import re
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

# =========================
# CONFIG
# =========================

EXCEL_FILE = "Problematic Withdrawals.xlsx"
ITEM_COL = "steam_market_hash_name"
PRICE_COL = "skinclub_price"

# =========================
# URL FORMAT (UNCHANGED)
# =========================

def format_skinclub_url(skin_name):
    skin_name_clean = skin_name.replace("★", "").replace("StatTrak™", "").strip()

    if skin_name_clean.startswith("Sticker"):
        parts = skin_name_clean.split(" | ")
        base = parts[0].strip().lower().replace(" ", "-")

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

        name_slug = "-".join(extra_tokens)
        return f"https://wiki.skin.club/en/items/{base}-{name_slug}"

    parts = skin_name_clean.split(" | ")
    weapon = parts[0].strip().lower().replace(" ", "-")
    name = parts[1].split(" (")[0].strip().lower().replace(" ", "-")

    skin_quality_match = re.search(r'\(([^)]+)\)', skin_name)
    skin_quality = (
        skin_quality_match.group(1).strip().lower().replace(" ", "-")
        if skin_quality_match else "factory-new"
    )

    is_stattrak = "StatTrak" in skin_name

    url = f"https://wiki.skin.club/en/items/{weapon}-{name}-{skin_quality}"
    if is_stattrak:
        url = f"https://wiki.skin.club/en/items/stattrak-{weapon}-{name}-{skin_quality}"

    return url

# =========================
# SCRAPER (MINIMAL CHANGES)
# =========================

def get_skinclub_price(driver, skin_name):
    url = format_skinclub_url(skin_name)

    try:
        driver.set_page_load_timeout(30)
        driver.get(url)

        if "/items/" not in driver.current_url:
            return None

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
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div.flex.items-center span.block.text-brand-300")
                )
            )

            for _ in range(10):
                txt = price_element.text.strip()
                if re.search(r"\$\s*\d", txt):
                    break
                time.sleep(0.3)

            return float(txt.replace("$", "").replace(",", "").strip())

        main_container = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "bg-brand-700.rounded-lg"))
        )

        py1_container = main_container.find_element(By.CLASS_NAME, "py-1")
        rows = py1_container.find_elements(
            By.XPATH,
            ".//div[contains(@class,'flex') and contains(@class,'cursor-pointer')]"
        )

        quality_from_name = ""
        if "(" in skin_name and ")" in skin_name:
            quality_from_name = skin_name.split("(")[-1].replace(")", "").strip().lower()

        for row in rows:
            try:
                quality_text = row.find_element(
                    By.CLASS_NAME, "truncate.flex-1"
                ).text.strip().lower()

                if "stattrak" in skin_name.lower():
                    price_span = row.find_element(
                        By.CSS_SELECTOR, ".truncate.text-rarity-stattrak.shrink-0"
                    )
                else:
                    price_span = row.find_element(
                        By.CSS_SELECTOR, ".truncate.text-primary-green-900.shrink-0"
                    )

                if quality_text == quality_from_name:
                    return float(price_span.text.replace("$", "").replace(",", ""))

            except Exception:
                continue

    except WebDriverException:
        return None

    return None

# =========================
# EXCEL WORKFLOW (AS IN VARIANT 2)
# =========================

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
                    price = get_skinclub_price(driver, item)
                    
                    if price is not None:
                        print(f"[{i}/{len(items)}] {item} → {price}")
                        prices.append(price)
                    else:
                        print(f"[{i}/{len(items)}] ⚠ No price found for {item}")
                        prices.append("-")

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
