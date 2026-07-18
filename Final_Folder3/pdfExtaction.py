import os
import re
import pandas as pd
import pdfplumber
from concurrent.futures import ThreadPoolExecutor, as_completed

# Settings
INPUT_FILE = "download_report1.xlsx"
OUTPUT_FILE = "ZOutput.xlsx"


def extract_from_pdf(pdf_path):
    text_content = ""
    full_path = os.path.abspath(pdf_path)
    try:
        with pdfplumber.open(full_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    text_content += text + "\n"
        return text_content
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
        return ""


CATALOG_LABELS = [r"Product\s*No\.?", r"Product\s*Number", r"Catalog\s*#", r"Catalog\s*Number",
                  r"Product\s*#", r"Item\s*Number", r"SKU", r"Part\s*No\.?",
                  r"Identification\s*Number", r"Product\s*Identification\s*Number",
                  r"SDS\s*Number", r"Material\s*Number", r"Formula\s*Number"]
TIGHT_CODE_PATTERN = r'\b[A-Z]{2,6}\d{3,7}[-.][A-Za-z0-9.]+\b'
BROAD_CODE_PATTERN = r'\b[A-Z]{0,6}\d{3,8}(?:[-.][A-Za-z0-9]+)?\b'


YEAR_PATTERN = re.compile(r'^(19|20)\d{2}$')
MAX_LABEL_DISTANCE = 150  # chars -- if nearest candidate is farther than this from the label, don't trust it


def extract_catalog_number(text):
    candidates = [(m.start(), m.group()) for m in re.finditer(TIGHT_CODE_PATTERN, text)]
    if not candidates:
        candidates = [(m.start(), m.group()) for m in re.finditer(BROAD_CODE_PATTERN, text)]

    # A catalog/identification number is essentially never a bare calendar
    # year (e.g. "2022") -- these are almost always issue-date fragments
    # that happen to sit near an "Identification Number" style label.
    candidates = [c for c in candidates if not YEAR_PATTERN.match(c[1])]

    if not candidates:
        return "Unknown"

    label_pattern = "|".join(CATALOG_LABELS)
    label_match = re.search(label_pattern, text, re.IGNORECASE)
    if label_match:
        label_pos = label_match.start()
        candidates.sort(key=lambda c: abs(c[0] - label_pos))
        nearest_pos, nearest_value = candidates[0]

        # If even the closest candidate is far from the label, the real
        # number is probably missing/differently formatted -- don't guess.
        if abs(nearest_pos - label_pos) <= MAX_LABEL_DISTANCE:
            return nearest_value
        return "Unknown"

    # No explicit catalog/identification label found in the document --
    # guessing from an unlabeled number is unreliable (it can grab a date,
    # zip code, phone number fragment, CFR citation, etc., as happened
    # with Dow SDS files where the "closest" unlabeled match was the
    # issue-date year). Report Unknown rather than a false positive.
    return "Unknown"


SIGNAL_WORD_LABEL_PATTERN = r"Signal\s*Word\s*[:\-]?\s*(Danger|Warning)"


def extract_signal_word(text):
    """
    Look for the actual 'Signal Word:' field (GHS classification, usually
    in Section 2 of an SDS) instead of just scanning the whole document
    for the words 'Danger'/'Warning', which appear many times elsewhere
    (hazard statements, precautions, etc.) and give false results.
    """
    label_match = re.search(SIGNAL_WORD_LABEL_PATTERN, text, re.IGNORECASE)
    if label_match:
        return label_match.group(1).capitalize()

    # Fallback: no explicit "Signal Word:" label found -- scan the whole
    # text as a last resort (less reliable, but better than nothing).
    upper_text = text.upper()
    if "DANGER" in upper_text:
        return "Danger"
    if "WARNING" in upper_text:
        return "Warning"
    return "None"


def smart_parse(text, row_catalog):
    # 1. CAS Pattern
    cas_match = re.search(r'\d{2,7}-\d{2}-\d', text)
    cas_no = cas_match.group(0) if cas_match else "N/A"

    # 2. Catalog Number (use Excel value if present, otherwise pull from PDF text)
    cat_num = str(row_catalog)
    if cat_num.lower() in ['nan', 'none', 'unknown_cat', ''] or not cat_num.strip():
        cat_num = extract_catalog_number(text)

    # 3. Signal Word (label-based, falls back to whole-document scan)
    signal = extract_signal_word(text)

    return cas_no, signal, cat_num


def process_single_record(row):
    # NOTE: column names match download_report.xlsx as produced by the
    # SDS downloader script (Product, Manufacturer, CAS, Link,
    # PDF URL, PDF File Name, Download Path, Status, Error Message).
    pdf_path = str(row.get("Download Path", "")).strip()

    if not pdf_path or pdf_path.lower() == "nan":
        print(f"⚠️ Skipped (no path recorded): {row.get('Product', 'Unknown')}")
        return None

    if not os.path.exists(pdf_path):
        print(f"⚠️ Skipped (file not found on disk): {pdf_path}")
        return None

    text = extract_from_pdf(pdf_path)
    if not text:
        print(f"⚠️ Skipped (no extractable text): {pdf_path}")
        return None

    # CAS, Signal word, and Catalog number
    cas, sig, cat = smart_parse(text, row.get("Catalog Number", ""))

    return {
        "Chemical_Name": str(row.get("Product", "Unknown")).strip(),
        "Manufacturer": str(row.get("Manufacturer", "Unknown")).strip(),
        "CAS_No": cas,
        "Signal_Word": sig,
        "Catalog_Number": cat,
        "PDF URL" :str(row.get("PDF URL", "")),
        "PDF File Name" :str(row.get("PDF File Name", "")),
        "Source_Link": str(row.get("Link", "")),
    }


def main():
    if not os.path.exists(INPUT_FILE):
        print(f"File {INPUT_FILE} not found!")
        return

    df = pd.read_excel(INPUT_FILE)

    if "Status" not in df.columns:
        print(f"'Status' column not found in {INPUT_FILE}. Columns present: {list(df.columns)}")
        return

    success_records = df[df["Status"] == "Success"]

    if success_records.empty:
        print("No successful download records found.")
        return

    print(f"Processing {len(success_records)} records...")
    results = []

    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = {executor.submit(process_single_record, row): row for _, row in success_records.iterrows()}
        for future in as_completed(futures):
            res = future.result()
            if res:
                results.append(res)
                print(f"✅ Processed: {res['Chemical_Name']}")

    if results:
        pd.DataFrame(results).to_excel(OUTPUT_FILE, index=False)
        print(f"\n🎉 Data saved to '{OUTPUT_FILE}'. {len(results)} of {len(success_records)} records extracted.")
    else:
        print("\n❌ No records were successfully processed — Output.xlsx was not created. "
              "Check the ⚠️ messages above (usually a missing/incorrect PDF path).")


if __name__ == "__main__":
    main()
