import os
import time
import re
import json
import pandas as pd
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager


# =========================
# CONFIG
# =========================
DEBUGGER_ADDR = "127.0.0.1:9222"

EXCEL_FILE = "Problematic Withdrawals.xlsx"
ITEM_COL = "steam_market_hash_name"
PRICE_COL = "casedrop_price"

SEARCH_INPUT_XPATH = "//input[@placeholder='Enter item name']"


# =========================
# CHROME (DEBUGGER)
# =========================
def get_debugger_driver() -> webdriver.Chrome:
    options = Options()
    options.add_experimental_option("debuggerAddress", DEBUGGER_ADDR)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver


def get_search_input(driver):
    return WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, SEARCH_INPUT_XPATH))
    )


# =========================
# PRICE PARSING
# =========================
def extract_price_number(text: str) -> float:
    cleaned = re.sub(r"[^\d.,]", "", text)
    cleaned = cleaned.replace(",", ".")
    return float(cleaned)


def get_skin_price(driver, search_input, skin_name: str):
    try:
        WebDriverWait(driver, 15).until(lambda d: search_input.is_enabled())

        search_input.click()
        search_input.send_keys(Keys.CONTROL + "a")
        search_input.send_keys(Keys.BACKSPACE)
        search_input.send_keys(skin_name)
        search_input.send_keys(Keys.ENTER)

        time.sleep(0.2)

        # NO ITEMS
        try:
            no_items = driver.find_element(By.CSS_SELECTOR, ".shop_items_list .itemEmpty")
            if "NO ITEMS" in no_items.text.upper():
                return None
        except:
            pass

        if "|" not in skin_name:
            return None

        items = driver.find_elements(By.CSS_SELECTOR, ".shop_items_list .item_container")
        if not items:
            return None

        is_stattrak = "stattrak" in skin_name.lower()

        for item in items:
            soup = BeautifulSoup(item.get_attribute("innerHTML"), "html.parser")

            has_track = soup.find("div", class_="info_track") is not None

            if not is_stattrak and has_track:
                continue
            if is_stattrak and not has_track:
                continue

            price_el = soup.find("div", class_="info_price")
            if price_el:
                return price_el.text.strip()

        return None

    except Exception as e:
        print(f"❌ Error for {skin_name}: {e}")
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
    driver = get_debugger_driver()
    search_input = get_search_input(driver)

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
                    raw = get_skin_price(driver, search_input, item)

                    if raw is None:
                        prices.append("-")
                        print(f"[{i}/{len(items)}] ❌ {item}")
                    else:
                        try:
                            price = extract_price_number(raw)
                            prices.append(price)
                            print(f"[{i}/{len(items)}] ✅ {item} → {price}")
                        except Exception:
                            prices.append("-")
                            print(f"[{i}/{len(items)}] ❌ {item}")

                df[PRICE_COL] = prices

            df.to_excel(writer, sheet_name=sheet, index=False)

    except KeyboardInterrupt:
        print("\n⚠ Interrupted by user (Ctrl+C)")

    finally:
        writer.close()
        driver.quit()
        print("💾 Excel saved, Chrome detached")


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
    print("\n✅ DONE")


if __name__ == "__main__":
    main()
