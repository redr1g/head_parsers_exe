import os
import time
from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

EXCEL_FILE = "Problematic Withdrawals.xlsx"
ITEM_COL = "steam_market_hash_name"
PRICE_COL = "ggdrop_price"

def get_price(driver, name_input, query):
    name_input.click()
    name_input.send_keys(Keys.CONTROL + "a")
    name_input.send_keys(Keys.BACKSPACE) 
    # вводимо назву скіна
    name_input.send_keys(query)
    name_input.send_keys(Keys.ENTER)
    time.sleep(2)

    if not "|" in query:
            return None

    is_stattrak = "stattrak" in query.lower()

    if is_stattrak:
        try:
            # непосрєдствєнно прайс
            price_el = driver.find_element(By.CLASS_NAME, "item_price__aCda4")
            return price_el.text.strip()
        except:
            return None
    else:
        # Проходимо по всіх айтемах, шукаємо перший без "StatTrak"
        try:
            # грід зі скінами
            items = driver.find_elements(By.CLASS_NAME, "items_items__x8V9i")
            html = items[0].get_attribute('outerHTML')
            soup = BeautifulSoup(html, 'html.parser')
            prices = soup.find_all('div', class_='item_price__aCda4')
            if len(prices) == 1:
                price_el = driver.find_element(By.CLASS_NAME, "item_price__aCda4")
                return price_el.text.strip()
            if len(prices) >= 2:
                second_price = prices[1].text
                return second_price.strip()
            else:
                return None
        except:
            return None

def choose_sheets(xls: pd.ExcelFile) -> list[str]:
    print("\nFound sheets:")
    for i, name in enumerate(xls.sheet_names, start=1):
        print(f"{i} - {name}")

    choice = input("\nEnter sheet number or 0 for ALL: ").strip()
    if choice == "0":
        return xls.sheet_names

    return [xls.sheet_names[int(choice) - 1]]

def process_sheets(sheet_names: list[str]):
    options = Options()
    options.add_argument("--headless=new")

    driver = webdriver.Chrome(options=options)

    driver.get("https://ggdrop.com/items")
    driver.implicitly_wait(10)
    time.sleep(2)

    name_input = driver.find_element(By.CSS_SELECTOR, 'input[placeholder="Name"]')

    xls = pd.ExcelFile(EXCEL_FILE)
    writer = pd.ExcelWriter(
        EXCEL_FILE,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    )

    try:
        total = len(sheet_names)
        sheet_idx = 0

        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)

            if sheet in sheet_names and ITEM_COL in df.columns:
                sheet_idx += 1
                print(f"\n=== Processing sheet {sheet_idx}/{total}: {sheet} ===")

                prices = []
                items = df[ITEM_COL].astype(str).tolist()

                for i, item in enumerate(items, start=1):
                    raw_price = get_price(driver, name_input, item)

                    if raw_price is None:
                        prices.append("-")
                        print(f"[{i}/{len(items)}] No price found for {item}")
                        continue

                    try:
                        price = float(raw_price[:-1].replace(" ", ""))
                        prices.append(price)
                        print(f"[{i}/{len(items)}] {item}: {price}")
                    except:
                        prices.append("-")
                        print(f"[{i}/{len(items)}] No price found for {item}")

                df[PRICE_COL] = prices

            df.to_excel(writer, sheet_name=sheet, index=False)

    except KeyboardInterrupt:
        print("\n⚠ Interrupted by user. Processing stopped.")
        driver.quit()
        writer.close()
        return

    finally:
        driver.quit()

    writer.close()


# =========================
# ENTRYPOINT
# =========================

def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ File '{EXCEL_FILE}' not found")
        return

    xls = pd.ExcelFile(EXCEL_FILE)
    sheets = choose_sheets(xls)
    process_sheets(sheets)

    print("\n✅ Done.")


if __name__ == "__main__":
    main()