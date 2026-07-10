from playwright.sync_api import sync_playwright
import pandas as pd

def scrape_vwr():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto("https://www.vwr.com/us/en/search/Chloroform")
        
        page.wait_for_load_state("networkidle")
        
        accordions = page.locator(".accordion-heading").all()
        all_data = []

        for item in accordions:
            try:
                item.scroll_into_view_if_needed()
                item.click()
                
                # Table ke rows ka wait karein
                page.wait_for_selector("li.item.item-container", timeout=5000)
                
                rows = page.locator("li.item.item-container")
                
                for row in rows.all():
                    if row.is_visible():
                        # Yahan .first() ka istemal kiya hai taaki strict mode error na aaye
                        # Hum specific class use kar rahe hain taaki duplicate na mile
                        cat_no_locator = row.locator("div.catalog-attribute.desktop-only[title='Catalog#']")
                        
                        if cat_no_locator.count() > 0:
                            cat_no = cat_no_locator.first.inner_text().strip()
                            
                            # SDS Link extract karna
                            sds_element = row.get_by_role("link", name="SDS")
                            sds_link = None
                            if sds_element.count() > 0:
                                href = sds_element.first.get_attribute("href")
                                if href:
                                    sds_link = "https://www.vwr.com" + href if href.startswith('/') else href

                            all_data.append({
                                'Catalog_Number': cat_no,
                                'SDS_Link': sds_link
                            })
                            print(f"Captured: {cat_no}")
                            
            except Exception as e:
                print(f"Skipping row/item due to error: {e}")
                continue
        
        df = pd.DataFrame(all_data)
        df.to_excel("vwr_data_clean.xlsx", index=False)
        print("Done! Data saved to vwr_data_clean.xlsx")
        
        browser.close()

if __name__ == "__main__":
    scrape_vwr()
