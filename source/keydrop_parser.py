import os
import re
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup


# =========================
# CONFIG
# =========================

EXCEL_FILE = "Problematic Withdrawals.xlsx"
ITEM_COL = "steam_market_hash_name"
PRICE_COL = "keydrop_price"

def format_skin_url(skin_name):
    skin_name = skin_name.replace("★ ", "").replace("StatTrak™", "").strip()

    if skin_name.startswith("Sticker"):
        skin_name = re.sub(r"\s*\((Gold|Holo|Foil|Glitter)\)", "", skin_name, flags=re.IGNORECASE)
        parts = skin_name.split(" | ")
        base = parts[0].strip().replace(" ", "-")

        extra_tokens = []
        for p in parts[1:]:
            match = re.search(r"\(([^)]+)\)", p)
            if match:
                inside = match.group(1).strip().replace(" ", "-")
                p_clean = re.sub(r"\([^)]+\)", "", p).strip().replace(" ", "-")
                if p_clean:
                    extra_tokens.append(p_clean)
                extra_tokens.append(inside)
            else:
                if p.strip():
                    extra_tokens.append(p.strip().replace(" ", "-"))

        return f"{base}-{'-'.join(extra_tokens)}"

    is_stattrak = "StatTrak" in skin_name
    parts = skin_name.split(" | ")
    weapon = parts[0].strip().replace(" ", "-")
    name = parts[1].split(" (")[0].strip().replace(" ", "-")

    url = f"{weapon}-{name}"
    if is_stattrak:
        url = f"StatTrak-{url}"
    return url

def extract_price_number(price_str):
    return float(
        price_str.replace(" ", "")
        .replace(",", ".")
        .replace("\xa0", "")
        .replace("$", "")
        .strip()
    )

def get_skin_price(driver, skin_name):
    url = f"https://key-drop.com/ru/skins/product/{format_skin_url(skin_name)}"

    try:
        driver.set_page_load_timeout(20)
        driver.get(url)
    except TimeoutException:
        return None
    except Exception:
        return None

    if "|" not in skin_name:
        return None

    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "tr")))
    except Exception:
        return None

    soup = BeautifulSoup(driver.page_source, "html.parser")

    if skin_name.strip().startswith("Sticker"):
        # шукаємо таблицю з класом, де є елемент <td class="text-[#8BBCDD]">
        try:
            header_elem = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR,
                    "h2.mx-auto.flex.items-center.whitespace-nowrap.text-center.text-xl.font-semibold.uppercase.leading-tight.text-white.lg\\:px-6"
                ))
            )
            site_title = header_elem.text.strip().strip().upper()

            clean_expected = skin_name.strip().upper()

            if clean_expected != site_title:
                return None
        except Exception:
            return None
        
        table = soup.select_one("table.grid")
        if table:
            # у кожному рядку знаходимо всі <td>
            rows = table.select("tr")
            for row in rows:
                tds = row.find_all("td")
                if len(tds) >= 2:
                    price_text = tds[1].get_text(strip=True)
                    if "$" in price_text:
                        # print(f"→ Found Sticker price: {price_text}")
                        return price_text

        return None
    
    quality = skin_name.split("(")[1].replace(")", "").strip()
    is_stattrak = "StatTrak" in skin_name
    is_knife = "★" in skin_name

    rows = soup.find_all("tr")
    # time.sleep(5) # Uncomment for debugging
    for row in rows:
        cols = row.find_all("td")
        if len(cols) >= 2:
            label = cols[0].get_text(strip=True)
            if quality in label:
                if is_knife:
                    # Ніж: окрема сторінка для ST, беремо завжди другу колонку
                    return cols[1].get_text(strip=True)
                else:
                    # Звичайна зброя: беремо ST або норм ціну залежно від назви
                    if is_stattrak and len(cols) >= 3:
                        return cols[2].get_text(strip=True)  # ST
                    else:
                        return cols[1].get_text(strip=True)  # звичайна

    return None


# =========================
# EXCEL WORKFLOW (AS BEFORE)
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
    options = Options()
    options.add_argument("--headless")
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
                    raw = get_skin_price(driver, item)
                    if raw is None:
                        prices.append("-")
                        print(f"[{i}/{len(items)}] No price found for {item}")
                    else:
                        try:
                            price = extract_price_number(raw)
                            prices.append(price)
                            print(f"[{i}/{len(items)}] {item} → {price}")
                        except Exception:
                            prices.append("-")
                            print(f"[{i}/{len(items)}] No price found for {item}")

                df[PRICE_COL] = prices

            df.to_excel(writer, sheet_name=sheet, index=False)

    except KeyboardInterrupt:
        print("\n⚠ Interrupted by user. Processing stopped.")
        writer.close()
        driver.quit()
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
