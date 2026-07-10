# from playwright.sync_api import sync_playwright
# import pandas as pd
# import time

# def run():
#     # 1. Excel file read karein
#     df = pd.read_excel('chemical_list.xlsx')
    
#     # Results store karne ke liye list
#     results = []
    
#     with sync_playwright() as p:
#         browser = p.chromium.launch(headless=False)
#         page = browser.new_page()
#         page.goto("https://chemicalsafety.com/sds-search/", timeout=60000, wait_until="networkidle")
        
#         # 2. Excel ke har naam ke liye loop chalayein
#         for chemical in df['ChemicalName']:
#             print(f"Searching for: {chemical}")
            
#             # Input field clear aur fill karein
#             input_box = page.locator("input[placeholder*='product name']")
#             input_box.fill("") # Pehle clear karein
#             input_box.fill(str(chemical))
            
#             # Click search
#             page.click("#cs_btnSearch")
            
#             # Thora wait karein takay results load ho jayein
#             page.wait_for_timeout(2000) 
            
#             try:
#                 # Pehli row ka link nikalna
#                 first_link = page.locator("#cs_divResults table tbody tr").first.locator("a").get_attribute("href")
#                 results.append({"ChemicalName": chemical, "Link": first_link})
#                 print(f"Found: {first_link}")
#             except:
#                 results.append({"ChemicalName": chemical, "Link": "Not Found"})
#                 print("Link not found for this chemical.")
                
#         # 3. Nayi Excel file save karein
#         output_df = pd.DataFrame(results)
#         output_df.to_excel('extracted_links.xlsx', index=False)
#         print("Done! 'extracted_links.xlsx' file create ho gayi hai.")
        
#         browser.close()

# if __name__ == "__main__":
#     run()

from playwright.sync_api import sync_playwright, TimeoutError
import pandas as pd
import csv
import time

INPUT_FILE = "chemical_list.xlsx"
OUTPUT_FILE = "sds_results.csv"


def run():

    df = pd.read_excel(INPUT_FILE)

    all_results = []

    with sync_playwright() as p:

        browser = p.chromium.launch(headless=True)

        page = browser.new_page()

        page.goto(
            "https://chemicalsafety.com/sds-search/",
            wait_until="networkidle",
            timeout=60000
        )

        for index, row in df.iterrows():

            chemical = str(row["Chemical"]).strip()

            if chemical == "" or chemical.lower() == "nan":
                continue

            print(f"\nSearching : {chemical}")

            try:

                # Search box clear
                page.locator("input[placeholder*='product name']").fill("")

                # Chemical enter
                page.fill(
                    "input[placeholder*='product name']",
                    chemical
                )

                # Search button
                page.click("#cs_btnSearch")

                # Wait for results
                page.wait_for_selector(
                    "#cs_divResults table tbody tr",
                    timeout=20000
                )

                rows = page.locator(
                    "#cs_divResults table tbody tr"
                )

                total = rows.count()

                print(f"Found {total} rows")

                if total == 0:

                    all_results.append({
                        "Search Keyword": chemical,
                        "Product": "",
                        "Manufacturer": "",
                        "CAS": "",
                        "Link": "",
                        "Status": "No Result"
                    })

                    continue

                for i in range(total):

                    cols = rows.nth(i).locator("td")

                    if cols.count() < 6:
                        continue

                    try:
                        href = cols.nth(5).locator("a").get_attribute("href")
                    except:
                        href = ""

                    all_results.append({

                        "Search Keyword": chemical,

                        "Product": cols.nth(0).inner_text().strip(),

                        "Manufacturer": cols.nth(1).inner_text().strip(),

                        "CAS": cols.nth(2).inner_text().strip(),

                        "Link": href,

                        "Status": "Success"

                    })

                time.sleep(1)

            except TimeoutError:

                print("Timeout")

                all_results.append({

                    "Search Keyword": chemical,

                    "Product": "",

                    "Manufacturer": "",

                    "CAS": "",

                    "Link": "",

                    "Status": "Timeout"

                })

            except Exception as e:

                print(e)

                all_results.append({

                    "Search Keyword": chemical,

                    "Product": "",

                    "Manufacturer": "",

                    "CAS": "",

                    "Link": "",

                    "Status": "Error"

                })

        browser.close()

    with open(
        OUTPUT_FILE,
        "w",
        newline="",
        encoding="utf-8-sig"
    ) as f:

        writer = csv.DictWriter(
            f,
            fieldnames=[
                "Search Keyword",
                "Product",
                "Manufacturer",
                "CAS",
                "Link",
                "Status"
            ]
        )

        writer.writeheader()
        writer.writerows(all_results)

    print(f"\nSaved : {OUTPUT_FILE}")


if __name__ == "__main__":
    run()