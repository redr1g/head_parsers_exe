import re
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd

EXCEL_FILE = "Problematic Withdrawals.xlsx"
ITEM_COL = "steam_market_hash_name"
PRICE_COL = "farmskins_price"
TMP_FILE = EXCEL_FILE.replace(".xlsx", "_temp.xlsx")

# =========================
# URL + PRICE HELPERS
# =========================

def format_skin_url(skin_name):
    skin_name = skin_name.replace("★ ", "").replace("StatTrak™", "").strip()

     # 🟣 Якщо це стікер
    if skin_name.startswith("Sticker"):
        parts = skin_name.split(" | ")
        base = parts[0].strip().lower().replace(" ", "-")

        extra_tokens = []
        for p in parts[1:]:
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
        return f"{base}-{name_slug}"

    if " | " not in skin_name and not skin_name.strip().startswith("Sticker"):
        print(f"⏭ Skipped (unsupported): {skin_name}")
        return None
    
    parts = skin_name.split(" | ")
    weapon = parts[0].strip().replace(" ", "-")
    name = parts[1].split(" (")[0].strip().replace(" ", "-")
    url = f"{weapon}-{name}".lower()
    return url


def extract_price_number(price: str | None):
    if not price or not isinstance(price, str):
        return None

    cleaned = (
        price.replace("\xa0", "")
        .replace(" ", "")
        .replace(",", ".")
        .replace("$", "")
    )

    try:
        return float(cleaned)
    except ValueError:
        return None


# =========================
# SCRAPER
# =========================

def get_skin_price(driver, skin_name):
    url = f"https://farmskins.com/items/{format_skin_url(skin_name)}"
    driver.get(url)

    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "item-statistics__row"))
        )
    except Exception:
        return None

    soup = BeautifulSoup(driver.page_source, "html.parser")

    # 🟣 Якщо це стікер
    if skin_name.strip().startswith("Sticker"):
        # Знаходимо всі спани з цією комбінацією класів
        price_spans = soup.select("div.item-statistics__row.item-statistics__padding.item-statistics__table span.item-statistics__col.item-statistics__span")

        if len(price_spans) >= 2:
            price = price_spans[1].get_text(strip=True)
            # print(f"→ Found Sticker price: {price}")
            return price

        print(f"⚠ Sticker price not found for {skin_name}")
        return None

    quality_match = re.search(r'\(([^)]+)\)', skin_name)
    quality = quality_match.group(1).strip() if quality_match else "Factory New"
    is_stattrak = "StatTrak" in skin_name

    rows = soup.select("div.item-statistics__row.item-statistics__padding.item-statistics__table")
    # time.sleep(5) # Uncomment for debugging
    for row in rows:
        cols = row.find_all("span", class_="item-statistics__col")
        if len(cols) >= 3:
            exterior = cols[0].get_text(strip=True)
            if quality.lower() in exterior.lower():
                price = cols[2 if is_stattrak else 1].get_text(strip=True)
                return price

    return None


# =========================
# EXCEL WORKFLOW
# =========================

def choose_sheets(xls: pd.ExcelFile) -> list[str]:
    print("\nFound sheets:")
    for i, name in enumerate(xls.sheet_names, start=1):
        print(f"{i} - {name}")

    choice = input("\nEnter sheet number or 0 for ALL: ").strip()

    if choice == "0":
        return xls.sheet_names

    idx = int(choice) - 1
    return [xls.sheet_names[idx]]

def safe_replace(src, dst, retries=5, delay=1.0):
    for i in range(retries):
        try:
            os.replace(src, dst)
            return
        except PermissionError:
            print(f"⚠ File locked, retry {i+1}/{retries}...")
            time.sleep(delay)
    raise PermissionError(f"Could not replace {dst} — file is locked")

def process_sheets(sheet_names: list[str]):
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-blink-features=AutomationControlled")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

    xls = pd.ExcelFile(EXCEL_FILE)
    writer = None
    wrote_anything = False

    try:
        writer = pd.ExcelWriter(
            TMP_FILE,
            engine="openpyxl",
            mode="w",
        )

        total_sheets = len(sheet_names)
        selected_index = 0

        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)

            if sheet in sheet_names and ITEM_COL in df.columns:
                selected_index += 1
                print(f"\n=== Processing sheet: {sheet} ===")

                items = df[ITEM_COL].astype(str).tolist()
                prices = []
                success = 0

                for i, item in enumerate(items, start=1):
                    raw_price = get_skin_price(driver, item)
                    price_number = extract_price_number(raw_price)
                    prices.append(price_number if price_number is not None else "-")

                    if price_number is not None:
                        success += 1

                    print(
                        f"[{selected_index}/{total_sheets}] "
                        f"[{i}/{len(items)}] "
                        f"{item} → {price_number}"
                    )

                df[PRICE_COL] = prices
                wrote_anything = True
                
                print(f"✔ Sheet done: {success}/{len(items)} prices found")

            df.to_excel(writer, sheet_name=sheet, index=False)

    except KeyboardInterrupt:
        print("\n⚠ Interrupted by user. Excel file was NOT modified.")
        driver.quit()
        return

    finally:
        driver.quit()
    if wrote_anything:
        writer.close()
        del writer
        time.sleep(0.5)
        safe_replace(TMP_FILE, EXCEL_FILE)

# =========================
# ENTRYPOINT
# =========================

def main():
    if os.path.exists(TMP_FILE):
        try:
            os.remove(TMP_FILE)
        except PermissionError:
            pass
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ File '{EXCEL_FILE}' not found")
        return

    xls = pd.ExcelFile(EXCEL_FILE)
    sheets = choose_sheets(xls)
    process_sheets(sheets)

    print("\n✅ Done.")


if __name__ == "__main__":
    main()
