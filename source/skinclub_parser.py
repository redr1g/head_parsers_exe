import os
import re
import time
# import sys
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook

EXCEL_FILE = "Problematic Withdrawals.xlsx"
ITEM_COL = "steam_market_hash_name"
PRICE_COL = "skinclub_price"

def format_skinclub_url(skin_name):
    skin_name_clean = skin_name.replace("★", "").replace("StatTrak™", "").strip()

    # 🟣 Якщо це стікер
    if skin_name_clean.startswith("Sticker"):
        parts = skin_name_clean.split(" | ")
        base = parts[0].strip().lower().replace(" ", "-")  # "sticker"

        extra_tokens = []
        for p in parts[1:]:
            # Якщо є щось у дужках — додаємо і текст, і вміст дужок
            match = re.search(r'\(([^)]+)\)', p)
            if match:
                inside = match.group(1).strip().lower().replace(" ", "-")
                p_clean = re.sub(r'\([^)]+\)', '', p).strip().lower().replace(" ", "-")
                if p_clean:
                    extra_tokens.append(p_clean)
                extra_tokens.append(inside)
            else:
                if p.strip():
                    extra_tokens.append(p.strip().lower().replace(" ", "-"))

        # Об’єднуємо всі частини: heroic + holo + 2020-rmr → heroic-holo-2020-rmr
        name_slug = "-".join(extra_tokens)
        return f"https://wiki.skin.club/en/items/{base}-{name_slug}"

    if " | " not in skin_name and not skin_name.strip().startswith("Sticker"):
        print(f"⏭ Skipped (unsupported): {skin_name}")
        return None
    
    parts = skin_name_clean.split(" | ")
    weapon = parts[0].strip().lower().replace(" ", "-")
    name = parts[1].split(" (")[0].strip().lower().replace(" ", "-")

    skin_quality_match = re.search(r'\(([^)]+)\)', skin_name)
    skin_quality = skin_quality_match.group(1).strip().lower().replace(" ", "-") if skin_quality_match else "factory-new"
    is_stattrak = "StatTrak" in skin_name
    
    url = f"https://wiki.skin.club/en/items/{weapon}-{name}-{skin_quality}"
    if is_stattrak:
        url = f"https://wiki.skin.club/en/items/stattrak-{weapon}-{name}-{skin_quality}"
    
    return url

def get_skinclub_price(driver, skin_name):
    url = format_skinclub_url(skin_name)
    print(f"Trying: {skin_name} → {url}")

    try:
        driver.set_page_load_timeout(30)
        driver.get(url)
    except TimeoutException:
        print(f"⚠ Timeout loading page for {skin_name} (skipped after 30s)")
        return None
    except Exception as e:
        print(f"✖ Error opening {skin_name}: {e}")
        return None

    try:
        # 🟢 Якщо це стікер — шукаємо інший елемент
        if skin_name.strip().startswith("Sticker"):
            price_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((
                    By.CSS_SELECTOR,
                    "div.flex.items-center span.block.text-brand-300"
                ))
            )

            # 🕐 Чекаємо, поки з'явиться ціна з цифрами
            for _ in range(10):
                price_text = price_element.text.strip()
                if re.search(r"\$\s*\d", price_text):
                    break
                time.sleep(0.3)

            price_text = price_text.replace("$", "").strip()
            price_text = price_text.replace(" ", "").replace(",", "")
            if not price_text:
                raise ValueError("Price text is empty or not loaded yet")

            print(f"→ Found Sticker price: ${price_text}")
            return float(price_text)
        
        # Чекаємо головний контейнер
        main_container = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "bg-brand-700.rounded-lg"))
        )

        py1_container = main_container.find_element(By.CLASS_NAME, "py-1")
        rows = py1_container.find_elements(By.XPATH, ".//div[contains(@class, 'flex') and contains(@class, 'cursor-pointer')]")

        # Витягуємо якість зі скіна
        quality_from_name = ""
        if '(' in skin_name and ')' in skin_name:
            quality_from_name = skin_name.split('(')[-1].replace(')', '').strip().lower()

        for row in rows:
            try:
                quality_span = row.find_element(By.CLASS_NAME, "truncate.flex-1")
                quality_text = quality_span.text.strip().lower()

                # Перевірка на StatTrak
                if "stattrak" in skin_name.lower():
                    price_span = row.find_element(By.CSS_SELECTOR, ".truncate.text-rarity-stattrak.shrink-0")
                else:
                    price_span = row.find_element(By.CSS_SELECTOR, ".truncate.text-primary-green-900.shrink-0")

                price_text = price_span.text.strip().replace('$', '').replace(',', '')

                if quality_text == quality_from_name:
                    print(f"→ Found: ${price_text} for {skin_name}")
                    return float(price_text)

            except Exception:
                continue  # пропустити цей рядок, якщо елемент не знайдено

    except Exception as e:
        print(f"✖ Error fetching {skin_name}: {e}")

    print(f"⚠ No price found for {skin_name}")
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
    options.add_argument("--disable-blink-features=AutomationControlled")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

    xls = pd.ExcelFile(EXCEL_FILE)
    writer = pd.ExcelWriter(
        EXCEL_FILE,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    )

    try:
        total_sheets = len(sheet_names)

        for sheet_idx, sheet in enumerate(sheet_names, start=1):
            print(f"\n=== Sheet {sheet_idx}/{total_sheets}: {sheet} ===")

            df = pd.read_excel(xls, sheet_name=sheet)

            if ITEM_COL not in df.columns:
                print("Skipping - no steam_market_hash_name")
                continue

            items = df[ITEM_COL].astype(str).tolist()
            total_items = len(items)

            prices = []
            success = 0

            for i, item in enumerate(items, start=1):
                price = get_skinclub_price(driver, item)
                prices.append(price)

                if price is not None:
                    success += 1

                print(
                    f"[{sheet_idx}/{total_sheets}] "
                    f"[{i}/{total_items}] "
                    f"{item} → {price}"
                )

            df[PRICE_COL] = prices
            df.to_excel(writer, sheet_name=sheet, index=False)

            print(f"✔ Sheet done: {success}/{total_items}")

    finally:
        writer.close()
        driver.quit()

def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ File '{EXCEL_FILE}' not found")
        return

    xls = pd.ExcelFile(EXCEL_FILE)
    sheets = choose_sheets(xls)
    process_sheets(sheets)

    print("\n✅ Done. Skin.club prices added.")

if __name__ == "__main__":
    main()
