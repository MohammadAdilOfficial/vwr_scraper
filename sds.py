import os
import time
import hashlib
import requests
import random
import threading
import pandas as pd
from urllib.parse import urljoin
from concurrent.futures import ThreadPoolExecutor, as_completed
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from dotenv import load_dotenv

load_dotenv()

# ==========================
# Configuration
# ==========================
CSV_FILE = "part_2.csv"
SAVE_FOLDER = "Final_Folder_Part_2_2"
REPORT_FILE = "download_report_Part_2_2.xlsx"
HEADLESS = True          # Multithreaded run mein True hi rakhein (headless=False multiple windows kholega)
NAV_TIMEOUT_MS = 60000
DROPDOWN_WAIT_MS = 1000
REQUEST_TIMEOUT_S = 60
DELAY_BETWEEN_PRODUCTS_S = 1
MAX_WORKERS = 5          # <-- Kitne products parallel process hon (thread count). Apni proxy list / CPU / RAM ke hisaab se adjust karein.

# --- PROXY CONFIGURATION (Loaded from .env) ---
PROXY_USER = os.getenv("PROXY_USER")
PROXY_PASS = os.getenv("PROXY_PASS")

raw_ips = os.getenv("PROXY_IPS", "")
PROXY_IPS = [ip.strip() for ip in raw_ips.split(",") if ip.strip()]

os.makedirs(SAVE_FOLDER, exist_ok=True)

# Status constants
STATUS_SUCCESS = "Success"
STATUS_NO_LINK = "No SDS Link"
STATUS_TIMEOUT = "Timeout"
STATUS_FAILED = "Failed"

# ==========================
# Thread-safety helpers
# ==========================
report_lock = threading.Lock()
counter_lock = threading.Lock()
print_lock = threading.Lock()

report_rows = []
counters = {
    STATUS_SUCCESS: 0,
    STATUS_NO_LINK: 0,
    STATUS_TIMEOUT: 0,
    STATUS_FAILED: 0,
}

# Har thread ka apna random.Random() instance -> shared random module par
# race condition / predictable-repeat pattern se bachne ke liye
thread_local = threading.local()


def get_thread_rng() -> random.Random:
    if not hasattr(thread_local, "rng"):
        # seed thread id + time se, taake har thread ki proxy pick alag ho
        thread_local.rng = random.Random(threading.get_ident() ^ int(time.time() * 1000))
    return thread_local.rng


def safe_print(*args, **kwargs):
    with print_lock:
        print(*args, **kwargs)


def add_report_row(row: dict) -> None:
    with report_lock:
        report_rows.append(row)


def increment_counter(status: str) -> None:
    with counter_lock:
        counters[status] += 1


# ==========================
# Helpers (same logic as before)
# ==========================
def sanitize_filename(name: str) -> str:
    illegal_chars = r'\/:*?"<>|'
    cleaned = "".join(c if c not in illegal_chars else "_" for c in str(name).strip())
    cleaned = " ".join(cleaned.split())
    return cleaned.rstrip(". ")


def build_filename(product: str, link: str) -> str:
    link_hash = hashlib.md5(link.encode("utf-8")).hexdigest()[:8]
    safe_product = sanitize_filename(product) or "Unknown_Product"
    return f"{safe_product}_{link_hash}.pdf"


def get_random_proxy() -> dict:
    """Har thread apni rng se proxy chunta hai, taake alag-alag IP milne ka chance zyada rahe."""
    if not PROXY_IPS:
        return {"requests": None, "playwright": None, "raw_ip": None}

    rng = get_thread_rng()
    selected_proxy = rng.choice(PROXY_IPS)

    requests_url = f"http://{PROXY_USER}:{PROXY_PASS}@{selected_proxy}"
    requests_proxies = {
        "http": requests_url,
        "https": requests_url
    }

    pw_server = f"http://{selected_proxy}"
    playwright_proxy = {
        "server": pw_server,
        "username": PROXY_USER,
        "password": PROXY_PASS
    }

    return {
        "requests": requests_proxies,
        "playwright": playwright_proxy,
        "raw_ip": selected_proxy
    }


def find_direct_sds_link(page) -> str | None:
    direct_selectors = [
        "span#sds_links a[href]",
        "a[href*='msds']",
        "a[href*='sds']",
        "a[href$='.pdf']",
    ]

    for selector in direct_selectors:
        locator = page.locator(selector)
        count = locator.count()

        for i in range(count):
            item = locator.nth(i)
            href = item.get_attribute("href")
            text = (item.inner_text() or "").strip().lower()

            if not href:
                continue

            looks_sds = (
                "sds" in href.lower()
                or "msds" in href.lower()
                or href.lower().endswith(".pdf")
                or "sds" in text
                or "safety data sheet" in text
            )

            if looks_sds:
                return href

    generic = page.locator("a.btn.btn-lg.btn-primary[href]")
    for i in range(generic.count()):
        item = generic.nth(i)
        text = (item.inner_text() or "").strip().lower()
        href = item.get_attribute("href")
        if href and "sds" in text:
            return href

    return None


def find_dropdown_sds_link(page) -> str | None:
    dropdown = page.locator("button.dropdown-toggle")

    if dropdown.count() == 0:
        return None

    try:
        dropdown.first.click()
        page.wait_for_timeout(DROPDOWN_WAIT_MS)
    except Exception:
        return None

    menu = page.locator("ul.dropdown-menu a")
    menu_count = menu.count()

    if menu_count == 0:
        return None

    for i in range(menu_count):
        item = menu.nth(i)
        href = item.get_attribute("href")
        text = (item.inner_text() or "").strip()

        if href and "US" in text and "English" in text:
            return href

    for i in range(menu_count):
        href = menu.nth(i).get_attribute("href")
        if href:
            return href

    return None


def is_valid_pdf_bytes(data: bytes) -> bool:
    return data[:5] == b"%PDF-"


def download_pdf_via_requests(session: requests.Session, pdf_url: str, referer: str, user_agent: str, proxy_dict: dict) -> bytes:
    headers = {
        "User-Agent": user_agent,
        "Referer": referer,
        "Accept": "application/pdf,*/*",
    }
    response = session.get(pdf_url, headers=headers, proxies=proxy_dict, timeout=REQUEST_TIMEOUT_S)
    response.raise_for_status()
    return response.content


def download_pdf_via_browser(context, pdf_url: str, referer: str) -> bytes:
    dl_page = context.new_page()
    try:
        try:
            with dl_page.expect_download(timeout=NAV_TIMEOUT_MS) as download_info:
                dl_page.goto(pdf_url, referer=referer, timeout=NAV_TIMEOUT_MS)
            download = download_info.value
            tmp_path = download.path()
            with open(tmp_path, "rb") as f:
                return f.read()
        except PlaywrightTimeoutError:
            response = dl_page.goto(pdf_url, referer=referer, timeout=NAV_TIMEOUT_MS)
            if response is not None:
                return response.body()
            raise
    finally:
        dl_page.close()


def download_pdf(context, session: requests.Session, pdf_url: str, referer: str,
                 user_agent: str, filepath: str, proxy_dict: dict) -> None:
    data = b""

    try:
        data = download_pdf_via_requests(session, pdf_url, referer, user_agent, proxy_dict)
    except Exception:
        data = b""

    if not is_valid_pdf_bytes(data):
        data = download_pdf_via_browser(context, pdf_url, referer)

    if not is_valid_pdf_bytes(data):
        raise ValueError("Downloaded content is not a valid PDF (missing %PDF- header)")

    with open(filepath, "wb") as f:
        f.write(data)


# ==========================
# Per-product worker (runs inside its own thread)
# ==========================
def process_product(row: dict) -> None:
    link = str(row.get("Link", "")).strip()
    name = str(row.get("Product", "")).strip()
    manufacturer = str(row.get("Manufacturer", "")).strip()
    cas = str(row.get("CAS", "")).strip()

    if not link or link.lower() == "nan":
        return

    safe_print(f"\n[Thread-{threading.get_ident()}] Processing: {name}")

    current_proxy = get_random_proxy()

    pdf_url = ""
    pdf_filename = ""
    download_path = ""
    status = ""
    error_message = ""

    # Har thread apna khud ka Playwright instance start karta hai
    # (sync Playwright ek hi thread ke andar chalna chahiye)
    with sync_playwright() as p:
        if current_proxy["playwright"]:
            safe_print(f"[Thread-{threading.get_ident()}] Using Proxy: {current_proxy['raw_ip']}")
            browser = p.chromium.launch(headless=HEADLESS, proxy=current_proxy["playwright"])
        else:
            browser = p.chromium.launch(headless=HEADLESS)

        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            page.goto(link, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
            page.wait_for_load_state("networkidle")

            found_url = find_dropdown_sds_link(page)

            if found_url:
                safe_print(f"[Thread-{threading.get_ident()}] ✓ Dropdown SDS Found")
            else:
                safe_print(f"[Thread-{threading.get_ident()}] Trying direct SDS link...")
                found_url = find_direct_sds_link(page)
                if found_url:
                    safe_print(f"[Thread-{threading.get_ident()}] ✓ Direct SDS Found")

            if not found_url:
                safe_print(f"[Thread-{threading.get_ident()}] ❌ No SDS Link Found")
                status = STATUS_NO_LINK
                increment_counter(STATUS_NO_LINK)
            else:
                pdf_url = urljoin(link, found_url)
                safe_print(f"[Thread-{threading.get_ident()}] PDF: {pdf_url}")

                session = requests.Session()
                for cookie in context.cookies():
                    session.cookies.set(cookie["name"], cookie["value"])

                user_agent = page.evaluate("navigator.userAgent")

                pdf_filename = build_filename(name, link)
                download_path = os.path.join(SAVE_FOLDER, pdf_filename)

                download_pdf(
                    context=context,
                    session=session,
                    pdf_url=pdf_url,
                    referer=link,
                    user_agent=user_agent,
                    filepath=download_path,
                    proxy_dict=current_proxy["requests"]
                )

                safe_print(f"[Thread-{threading.get_ident()}] ✓ Downloaded -> {download_path}")

                status = STATUS_SUCCESS
                increment_counter(STATUS_SUCCESS)

        except PlaywrightTimeoutError as e:
            safe_print(f"[Thread-{threading.get_ident()}] ❌ Timeout")
            status = STATUS_TIMEOUT
            error_message = str(e)
            increment_counter(STATUS_TIMEOUT)

        except Exception as e:
            safe_print(f"[Thread-{threading.get_ident()}] ❌ Error: {e}")
            status = STATUS_FAILED
            error_message = str(e)
            increment_counter(STATUS_FAILED)

        finally:
            try:
                page.close()
                context.close()
                browser.close()
            except Exception:
                pass

    add_report_row({
        "Product": name,
        "Manufacturer": manufacturer,
        "CAS": cas,
        "Link": link,
        "PDF URL": pdf_url,
        "PDF File Name": pdf_filename,
        "Download Path": download_path,
        "Status": status,
        "Error Message": error_message,
    })

    time.sleep(DELAY_BETWEEN_PRODUCTS_S)


# ==========================
# Main
# ==========================
def download_with_playwright(csv_file: str) -> None:
    df = pd.read_csv(csv_file)
    rows = [row for _, row in df.iterrows() if str(row.get("Link", "")).strip().lower() != "nan"]
    total_products = len(rows)

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [executor.submit(process_product, row) for _, row in df.iterrows()]
        for future in as_completed(futures):
            # Agar kisi worker ne exception raise ki (jo already andar handle ho jati hai),
            # phir bhi yahan future.result() call karke silent crash na hone dein
            try:
                future.result()
            except Exception as e:
                safe_print("Unhandled worker error:", e)

    # ==========================
    # Excel Report
    # ==========================
    report_df = pd.DataFrame(report_rows, columns=[
        "Product",
        "Manufacturer",
        "CAS",
        "Link",
        "PDF URL",
        "PDF File Name",
        "Download Path",
        "Status",
        "Error Message",
    ])
    report_df.to_excel(REPORT_FILE, index=False)

    print("\n" + "=" * 40)
    print("SUMMARY")
    print("=" * 40)
    print(f"Total Products : {total_products}")
    print(f"Downloaded     : {counters[STATUS_SUCCESS]}")
    print(f"Failed         : {counters[STATUS_FAILED]}")
    print(f"Timeout        : {counters[STATUS_TIMEOUT]}")
    print(f"No SDS Link    : {counters[STATUS_NO_LINK]}")
    print("=" * 40)
    print(f"Report saved to: {REPORT_FILE}")


if __name__ == "__main__":
    download_with_playwright(CSV_FILE)
    print("\nFinished.")
