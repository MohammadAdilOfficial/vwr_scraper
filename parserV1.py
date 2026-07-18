import os
import re
import json
import requests
import pandas as pd
import pdfplumber
from concurrent.futures import ThreadPoolExecutor, as_completed

# Settings
INPUT_FILE = "download_report_Part_2_2.xlsx"
OUTPUT_FILE = "ZOutputTest.xlsx"

# Set to a small number (e.g. 20) to test on a subset first.
# Set to None to process ALL records (use this for the real 74k run).
TEST_LIMIT = 20

# ---- Ollama (local LLM fallback) settings ----
OLLAMA_URL = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "qwen2.5:1.5b"  # halka model -- 3.8GB RAM wali VM ke liye fit
USE_LLM_FALLBACK = True        # False kar do agar Ollama available nahi


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

# Generic fallback pattern: catches labels not in CATALOG_LABELS above,
# e.g. "Art. No.", "Ref.", "Cat. No.", "Product Code", "Item #", etc.
GENERIC_LABEL_PATTERN = re.compile(
    r'([A-Za-z][A-Za-z\.\s]{2,30}?(?:No\.?|Number|Code|ID|#|SKU|Cat\.?|Ref\.?|Art\.?))\s*[:\-]\s*'
    r'([A-Za-z0-9\-\./]{3,20})',
    re.IGNORECASE
)

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

    if candidates:
        label_pattern = "|".join(CATALOG_LABELS)
        label_match = re.search(label_pattern, text, re.IGNORECASE)
        if label_match:
            label_pos = label_match.start()
            candidates.sort(key=lambda c: abs(c[0] - label_pos))
            nearest_pos, nearest_value = candidates[0]

            if abs(nearest_pos - label_pos) <= MAX_LABEL_DISTANCE:
                return nearest_value

    # Fallback 1: generic label pattern (catches labels not in CATALOG_LABELS,
    # e.g. "Art. No.", "Ref.", "Cat. No.", "Product Code")
    for label, value in GENERIC_LABEL_PATTERN.findall(text):
        if not YEAR_PATTERN.match(value):
            return value

    # No explicit catalog/identification label found in the document --
    # guessing from an unlabeled number is unreliable (it can grab a date,
    # zip code, phone number fragment, CFR citation, etc.). Report Unknown
    # rather than a false positive -- LLM fallback handles this case later.
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


def extract_catalog_llm(text):
    """
    Local Ollama call -- used ONLY as a last-resort fallback when regex-based
    extraction returns 'Unknown'. No per-request billing; runs on your own
    GCP VM (CPU or GPU) via `ollama serve`.
    """
    try:
        resp = requests.post(
            OLLAMA_URL,
            json={
                "model": OLLAMA_MODEL,
                "prompt": (
                    "Extract only the catalog/product/item number from this "
                    "Safety Data Sheet (SDS) text. Respond with ONLY valid JSON, "
                    "no extra words, no markdown fences, in this exact format: "
                    '{"catalog_number": "..."}. '
                    "If no catalog/product/item number is found, use \"Unknown\".\n\n"
                    f"Text:\n{text[:2000]}"
                ),
                "stream": False,
            },
            timeout=60,
        )
        resp.raise_for_status()
        raw = resp.json()["response"].strip()

        # Strip possible markdown code fences the model might add anyway
        raw = re.sub(r'^```(json)?|```$', '', raw, flags=re.MULTILINE).strip()

        parsed = json.loads(raw)
        value = str(parsed.get("catalog_number", "Unknown")).strip()
        return value if value else "Unknown"
    except Exception as e:
        print(f"⚠️ Ollama fallback failed: {e}")
        return "Unknown"


def smart_parse(text, row_catalog):
    # 1. CAS Pattern
    cas_match = re.search(r'\d{2,7}-\d{2}-\d', text)
    cas_no = cas_match.group(0) if cas_match else "N/A"

    # 2. Catalog Number (use Excel value if present, otherwise pull from PDF text)
    cat_num = str(row_catalog)
    if cat_num.lower() in ['nan', 'none', 'unknown_cat', ''] or not cat_num.strip():
        cat_num = extract_catalog_number(text)

    # 3. LLM fallback -- only fires when regex-based methods gave up
    if cat_num == "Unknown" and USE_LLM_FALLBACK:
        cat_num = extract_catalog_llm(text)

    # 4. Signal Word (label-based, falls back to whole-document scan)
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
        "PDF URL": str(row.get("PDF URL", "")),
        "PDF File Name": str(row.get("PDF File Name", "")),
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

    if TEST_LIMIT:
        success_records = success_records.head(TEST_LIMIT)
        print(f"⚙️ TEST_LIMIT active -- only processing first {TEST_LIMIT} records.")

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
