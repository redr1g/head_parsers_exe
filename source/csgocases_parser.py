import os
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from openpyxl import Workbook
import time

EXCEL_FILE = "Problematic Withdrawals.xlsx"
ITEM_COL = "steam_market_hash_name"
PRICE_COL = "csgocases_price"

# def initialize_driver(debug_port="127.0.0.1:9222", driver_path="D:/chromedriver-win64/chromedriver.exe"):
def initialize_driver():
    """Initialize Chrome WebDriver with existing browser session"""
    options = Options()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def search_skin(driver, search_input, skin_name):
    """Search for a skin and return the item blocks"""
    try:
        # Clear existing search text
        # search_input.send_keys(Keys.CONTROL + "a")
        # search_input.send_keys(Keys.DELETE)
        ActionChains(driver) \
        .click(search_input) \
        .key_down(Keys.CONTROL) \
        .send_keys("a") \
        .key_up(Keys.CONTROL) \
        .send_keys(Keys.DELETE) \
        .perform()

        time.sleep(1)
        
        # Input the skin name
        search_input.send_keys(skin_name)
        time.sleep(2)  # Wait for search results
        
        # Get all item blocks
        item_blocks = driver.find_elements(By.CLASS_NAME, "item-content")
        
        # ADDED: Clear search if too many results are found and retry
        if len(item_blocks) > 2 or not item_blocks:
            print(f"Too many / zero results for {skin_name}, refining search...")
            search_input.send_keys(Keys.CONTROL + "a")
            search_input.send_keys(Keys.ARROW_RIGHT)
            search_input.send_keys(Keys.BACKSPACE)
            time.sleep(2)  # Wait for refined search results
            # Get updated item blocks after refining
            item_blocks = driver.find_elements(By.CLASS_NAME, "item-content")
            print(f"After refining: found {len(item_blocks)} results")

        return item_blocks
    except Exception as e:
        print(f"Error searching for {skin_name}: {str(e)}")
        return None

def get_skin_price(item_blocks, skin_name):
    """Extract price for the specified skin, handling StatTrak designation"""
    # Determine if we're looking for a StatTrak version
    if not item_blocks:
        print(f"Error: No results found for skin: {skin_name}")
        return None

    is_stattrak = "StatTrak™" in skin_name
    is_souvenir = "Souvenir" in skin_name
    is_knife = "★" in skin_name

    for item in item_blocks:
        try:
            html = item.get_attribute('innerHTML')
            soup = BeautifulSoup(html, 'html.parser')
            
            # Find the image with the alt attribute containing the skin name
            img = soup.find('img')
            if not img or not img.get('alt'):
                continue
                
            alt_text = img.get('alt')
            
            if not is_souvenir and "Souvenir" in alt_text:
                print(f"Skipping Souvenir item for non-Souvenir search: {alt_text}")
                continue
            
            # YANDERE DEV MOMENT UAHDUASHDADAHDHADAHDASHGDHAAHAHJHHA
            if is_knife:
                if (is_stattrak and "★ StatTrak™" in alt_text) or (not is_stattrak and "★ StatTrak™" not in alt_text):
                    if skin_name.replace("★ StatTrak™ ", "") in alt_text:
                        # Extract price
                        price_span = soup.find('span', class_='resell-price-span')
                        if price_span:
                            price_text = price_span.text.strip()
                            price_float = float(price_text.replace('$', '').replace('€', '').replace(',', ''))
                            return price_float
                        
            # Check if the alt text matches our search criteria (StatTrak or not)
            if (is_stattrak and "StatTrak™" in alt_text) or (not is_stattrak and "StatTrak™" not in alt_text):
                if skin_name.replace("StatTrak™ ", "") in alt_text:
                    # Extract price
                    price_span = soup.find('span', class_='resell-price-span')
                    if price_span:
                        # return price_span.text.strip()
                        price_text = price_span.text.strip()
                        price_float = float(price_text.replace('$', '').replace('€', '').replace(',', ''))
                        return price_float
        except Exception as e:
            print(f"Error processing item: {str(e)}")
    
    return None

# =========================
# EXCEL WORKFLOW
# =========================
def choose_sheets(xls: pd.ExcelFile) -> list[str]:
    print("\nFound sheets:")
    for i, s in enumerate(xls.sheet_names, start=1):
        print(f"{i} - {s}")

    choice = input("\nEnter sheet number or 0 for ALL: ").strip()

    if choice == "0":
        return xls.sheet_names

    return [xls.sheet_names[int(choice) - 1]]


def process_sheets(sheet_names):
    driver = initialize_driver()

    search_input = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Search']"))
    )

    xls = pd.ExcelFile(EXCEL_FILE)

    writer = pd.ExcelWriter(
        EXCEL_FILE,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace",
    )

    try:
        total = len(sheet_names)
        idx = 0

        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)

            if sheet in sheet_names and ITEM_COL in df.columns:
                idx += 1
                print(f"\n=== Processing sheet {idx}/{total}: {sheet} ===")

                prices = []
                items = df[ITEM_COL].astype(str).tolist()

                for i, item in enumerate(items, start=1):
                    item_blocks = search_skin(driver, search_input, item)
                    price = get_skin_price(item_blocks, item)

                    # повторна спроба, як у твоєму main
                    if price is None and item_blocks:
                        price = get_skin_price(item_blocks, item)

                    if price is None:
                        prices.append("-")
                        print(f"[{i}/{len(items)}] ❌ {item}")
                    else:
                        prices.append(price)
                        print(f"[{i}/{len(items)}] ✅ {item} → {price}")

                    time.sleep(2)

                df[PRICE_COL] = prices

            # зберігаємо ЛИСТ ЗА ЛИСТОМ
            df.to_excel(writer, sheet_name=sheet, index=False)

    except KeyboardInterrupt:
        print("\n⚠ Interrupted by user (Ctrl+C)")
        print("💾 Saving Excel...")

    finally:
        writer.close()
        driver.quit()
        print("💾 Excel saved, Chrome detached")

def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ File '{EXCEL_FILE}' not found")
        return

    xls = pd.ExcelFile(EXCEL_FILE)
    sheets = choose_sheets(xls)

    process_sheets(sheets)
    print("\n✅ DONE")


if __name__ == "__main__":
    main()