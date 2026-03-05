"""
Hornbach Brikety Tracker – CI scraper + Google Sheets
=====================================================
Beží v GitHub Actions každý piatok o 16:00.
Scrapne drevené brikety, zapíše do Google Sheets.

Požiadavky: pip install playwright gspread google-auth openpyxl
            playwright install chromium
"""

import asyncio
import json
import os
import re
import sys
import unicodedata
from datetime import datetime

try:
    from playwright.async_api import async_playwright
except ImportError:
    print("pip install playwright && playwright install chromium")
    sys.exit(1)

try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    print("pip install gspread google-auth")
    sys.exit(1)

# ── Konfigurácia ──────────────────────────────────────────────────────────────
CATEGORY_URL = "https://www.hornbach.sk/c/krby-radiatory-a-klimatizacie/palivove-drevo-pelety-a-brikety/S15929/"
KEYWORDS = ["briketa", "brikety", "brikiet", "briket"]
EXCLUDE_KEYWORDS = ["hnedouholn", "uholn", "hnedouhol"]

CANONICAL_STORES = [
    "HORNBACH Nitra",
    "HORNBACH Bratislava - Ružinov",
    "HORNBACH Bratislava - Devínska Nová Ves",
    "HORNBACH Košice",
    "HORNBACH Prešov",
]

STORE_SHORT = {
    "HORNBACH Nitra": "Nitra",
    "HORNBACH Bratislava - Ružinov": "BA Ružinov",
    "HORNBACH Bratislava - Devínska Nová Ves": "BA Devínska",
    "HORNBACH Košice": "Košice",
    "HORNBACH Prešov": "Prešov",
}


def matches_keywords(text: str) -> bool:
    t = text.lower()
    t_plain = unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode("ascii")
    if not any(k in t or k in t_plain for k in KEYWORDS):
        return False
    if any(k in t or k in t_plain for k in EXCLUDE_KEYWORDS):
        return False
    return True


def canonicalize_store(raw_name: str):
    raw_lower = raw_name.lower().strip()
    if "etky predajne" in raw_lower or "vsetky predajne" in raw_lower:
        return None
    if "hornbach" not in raw_lower:
        return None
    for canonical in CANONICAL_STORES:
        canon_lower = canonical.lower()
        if raw_lower.startswith(canon_lower):
            return canonical
        city_part = canon_lower.replace("hornbach ", "").strip()
        city_keywords = [w.strip() for w in city_part.split("-")]
        last_keyword = city_keywords[-1].strip()
        if last_keyword and last_keyword in raw_lower:
            return canonical
    base = raw_name.split(",")[0].strip().rstrip(".")
    return base


def store_sort_key(s):
    for i, canonical in enumerate(CANONICAL_STORES):
        if s == canonical:
            return (i, s)
    return (99, s)


# ── Scraper ───────────────────────────────────────────────────────────────────
async def scrape():
    products = []
    print("Nacitavam kategoriu brikety...")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 900},
            locale="sk-SK",
        )
        page = await context.new_page()

        await page.goto(CATEGORY_URL, wait_until="networkidle", timeout=60000)

        # Cookie banner
        for sel in [
            "#onetrust-accept-btn-handler",
            "button:has-text('Prijat')", "button:has-text('Suhlasim')",
            "button:has-text('Akceptovat')", "button:has-text('Prijať')",
            "button:has-text('Súhlasím')",
        ]:
            try:
                btn = page.locator(sel).first
                if await btn.is_visible(timeout=2000):
                    await btn.click()
                    await page.wait_for_timeout(800)
                    break
            except Exception:
                pass

        # Scroll
        for _ in range(8):
            await page.evaluate("window.scrollBy(0, 800)")
            await page.wait_for_timeout(600)
        await page.evaluate("window.scrollTo(0, 0)")
        await page.wait_for_timeout(500)

        seen_urls = set()
        product_links = []

        all_links = await page.query_selector_all("a[href*='/p/']")
        print(f"  Linkov na stranke: {len(all_links)}")

        for link in all_links:
            try:
                href = await link.get_attribute("href")
                if not href or '/p/' not in href:
                    continue
                full_url = href if href.startswith("http") else f"https://www.hornbach.sk{href}"
                if full_url in seen_urls:
                    continue

                text = ""
                try:
                    text = (await link.inner_text()).strip()
                except Exception:
                    pass

                if not text or len(text) < 5:
                    try:
                        text = await page.evaluate("""el => {
                            const tile = el.closest('article') || el.closest('li') ||
                                         el.closest('[class*="Tile"]') || el.closest('[class*="product"]') ||
                                         el.parentElement;
                            return tile ? tile.innerText : '';
                        }""", link)
                        text = (text or "").strip()
                    except Exception:
                        pass

                if not text:
                    continue

                if matches_keywords(text):
                    seen_urls.add(full_url)
                    name = next(
                        (l.strip() for l in text.splitlines() if len(l.strip()) > 8),
                        text[:80]
                    )
                    product_links.append({"url": full_url, "name": name[:100]})
                    print(f"  + {name[:65]}")
            except Exception:
                pass

        print(f"\n{len(product_links)} produktov na spracovanie")

        if not product_links:
            print("Ziadne produkty.")
            await browser.close()
            return products

        # Detail produktu
        for idx, prod in enumerate(product_links):
            print(f"\n[{idx+1}/{len(product_links)}] {prod['name']}")

            prod_data = {
                "name": prod["name"],
                "ean": "",
                "artikel_nr": "",
                "price": "",
                "url": prod["url"],
                "stores": {}
            }

            m = re.search(r'/(\d{5,})/?', prod["url"])
            if m:
                prod_data["artikel_nr"] = m.group(1)

            pp = None
            try:
                pp = await context.new_page()
                await pp.goto(prod["url"], wait_until="domcontentloaded", timeout=45000)
                await pp.wait_for_timeout(2000)

                # Cena
                try:
                    for price_sel in [
                        "[data-testid='article-price'] [class*='value']",
                        "[class*='article-price'] [class*='value']",
                        "[class*='ArticlePrice'] [class*='value']",
                        "[class*='price-value']",
                        "[class*='Price'] strong",
                        "[class*='price'] strong",
                        "strong[class*='price']",
                    ]:
                        el = pp.locator(price_sel).first
                        if await el.is_visible(timeout=800):
                            prod_data["price"] = (await el.inner_text()).strip()
                            break
                except Exception:
                    pass

                if not prod_data["price"]:
                    try:
                        body_text = await pp.inner_text("body")
                        pm = re.search(r'(\d+,\d{2})\s*€', body_text)
                        if pm:
                            prod_data["price"] = f"{pm.group(1)} €"
                    except Exception:
                        pass

                # EAN
                for expand_text in [
                    "VIAC INFORMACII O VYROBKU", "VIAC INFORMÁCIÍ O VÝROBKU",
                    "Viac informacii", "Viac informácií",
                ]:
                    try:
                        btn = pp.get_by_text(expand_text, exact=False).first
                        if await btn.is_visible(timeout=800):
                            await btn.click()
                            await pp.wait_for_timeout(700)
                            break
                    except Exception:
                        pass

                try:
                    rows = await pp.query_selector_all("tr")
                    for row in rows:
                        row_text = await row.inner_text()
                        if "EAN" in row_text:
                            cells = await row.query_selector_all("td")
                            if len(cells) >= 2:
                                prod_data["ean"] = (await cells[-1].inner_text()).strip()
                            if prod_data["ean"]:
                                break
                except Exception:
                    pass

                if not prod_data["ean"]:
                    try:
                        body = await pp.inner_text("body")
                        em2 = re.search(r'EAN\s*\n?\s*([0-9]{8,}(?:[,\s]+[0-9]{8,})*)', body)
                        if em2:
                            prod_data["ean"] = em2.group(1).strip()
                    except Exception:
                        pass

                print(f"  EAN: {prod_data['ean'] or '-'}  Cena: {prod_data['price'] or '-'}")

                # Dostupnosť
                clicked = False
                for avail_text in [
                    "SKONTROLOVAT DOSTUPNOST", "SKONTROLOVAŤ DOSTUPNOSŤ",
                    "Skontrolovat dostupnost", "Skontrolovať dostupnosť",
                    "dostupnost v predajni", "dostupnosť v predajni",
                ]:
                    try:
                        btn = pp.get_by_text(avail_text, exact=False).first
                        if await btn.is_visible(timeout=1500):
                            await btn.click()
                            await pp.wait_for_timeout(3000)
                            clicked = True
                            break
                    except Exception:
                        pass

                if not clicked:
                    for btn_sel in [
                        "[class*='StoreAvailability'] button",
                        "[class*='store-availability'] button",
                        "button[class*='store']",
                    ]:
                        try:
                            btn = pp.locator(btn_sel).first
                            if await btn.is_visible(timeout=800):
                                await btn.click()
                                await pp.wait_for_timeout(3000)
                                clicked = True
                                break
                        except Exception:
                            pass

                # Modal
                modal_text = ""
                for modal_sel in ["[role='dialog']", "[class*='Modal']", "[class*='modal']", "[class*='Overlay']"]:
                    try:
                        modal = pp.locator(modal_sel).first
                        if await modal.is_visible(timeout=2000):
                            t = await modal.inner_text()
                            if "HORNBACH" in t:
                                modal_text = t
                                break
                    except Exception:
                        pass

                if not modal_text or "HORNBACH" not in modal_text:
                    try:
                        modal_text = await pp.inner_text("body")
                    except Exception:
                        pass

                # Parse stores
                if modal_text and "HORNBACH" in modal_text:
                    lines = [l.strip() for l in modal_text.splitlines() if l.strip()]
                    found_stores = {}
                    i = 0
                    while i < len(lines):
                        line = lines[i]
                        if "HORNBACH" in line:
                            canonical = canonicalize_store(line)
                            if canonical is not None:
                                stock_val = None
                                for j in range(i + 1, min(i + 15, len(lines))):
                                    jline = lines[j]
                                    if "HORNBACH" in jline and j > i:
                                        break
                                    sm = re.search(r'(\d[\d\s]*)\s*(balen[ií]e?|ks\b|kus)', jline, re.IGNORECASE)
                                    if sm:
                                        stock_val = int(sm.group(1).replace(" ", ""))
                                        break
                                    if re.match(r'^\d+$', jline):
                                        nearby = " ".join(lines[max(0, j-2):min(len(lines), j+3)]).lower()
                                        if any(w in nearby for w in ["dostupn", "skladom", "dispoz", "baleni", "balení", "objedn"]):
                                            stock_val = int(jline)
                                            break
                                    if re.search(r'nie je k dispoz|nedostupn|momentálne nie|momentalne nie|vypredané|vypredane|0 balení|0 balenie', jline, re.IGNORECASE):
                                        stock_val = 0
                                        break
                                    sm2 = re.search(r'dostupn[éeý]\s*>?\s*(\d+)', jline, re.IGNORECASE)
                                    if sm2:
                                        stock_val = int(sm2.group(1))
                                        break

                                if canonical in found_stores:
                                    if found_stores[canonical] == "?" and stock_val is not None:
                                        found_stores[canonical] = stock_val
                                else:
                                    found_stores[canonical] = stock_val if stock_val is not None else "?"

                                print(f"  {STORE_SHORT.get(canonical, canonical)}: {found_stores[canonical]}")
                        i += 1
                    prod_data["stores"] = found_stores

            except Exception as e:
                print(f"  Chyba: {e}")
            finally:
                if pp:
                    try:
                        await pp.close()
                    except Exception:
                        pass

            products.append(prod_data)
            await asyncio.sleep(0.5)

        await browser.close()

    print(f"\nHotovo – {len(products)} produktov")
    return products


# ── Google Sheets ─────────────────────────────────────────────────────────────
def get_sheets_client():
    """Vytvorí gspread klienta z env premennej alebo súboru."""
    creds_json = os.environ.get("GOOGLE_SHEETS_CREDS")
    if creds_json:
        creds_dict = json.loads(creds_json)
    else:
        # Lokálny vývoj — hľadaj súbor
        creds_path = os.path.join(os.path.dirname(__file__), "service_account.json")
        if os.path.exists(creds_path):
            with open(creds_path) as f:
                creds_dict = json.load(f)
        else:
            print("CHYBA: Nenájdený GOOGLE_SHEETS_CREDS env alebo service_account.json")
            sys.exit(1)

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


def write_to_sheets(products: list):
    """Zapíše aktuálny stav do Google Sheets.

    Sheet 'Data' — append nového stĺpca s dátumom:
    Riadok 1: hlavičky (Produkt, EAN, Artikl, Cena, [dátumy...])
    Riadok 2+: pre každú kombináciu produkt+predajňa

    Sheet 'Posledný beh' — prehľadná tabuľka aktuálneho stavu
    """
    sheet_id = os.environ.get("GOOGLE_SHEET_ID")
    if not sheet_id:
        print("CHYBA: GOOGLE_SHEET_ID env nie je nastavený")
        sys.exit(1)

    gc = get_sheets_client()
    spreadsheet = gc.open_by_key(sheet_id)
    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")
    date_col_header = datetime.now().strftime("%d.%m.%Y %H:%M")

    # Zozbieraj unikátne predajne
    all_stores = sorted(
        set(s for p in products for s in p["stores"].keys()),
        key=store_sort_key
    )

    # ── Sheet: Posledný beh ───────────────────────────────────────────────
    try:
        ws_current = spreadsheet.worksheet("Posledný beh")
        ws_current.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws_current = spreadsheet.add_worksheet("Posledný beh", rows=50, cols=15)

    headers = ["Produkt", "Artikl č.", "Cena"] + [STORE_SHORT.get(s, s) for s in all_stores]
    rows = [headers]
    for prod in products:
        row = [
            prod["name"],
            prod["artikel_nr"],
            prod["price"],
        ]
        for store in all_stores:
            val = prod["stores"].get(store, "–")
            row.append(val if val != "?" else "–")
        rows.append(row)

    ws_current.update(rows, value_input_option="USER_ENTERED")
    print(f"  'Posledný beh' zapísaný ({len(products)} produktov)")

    # ── Sheet: História ───────────────────────────────────────────────────
    # Formát: každý riadok = produkt+predajňa, každý stĺpec = dátum behu
    # To umožňuje jednoduché porovnanie a grafy

    try:
        ws_hist = spreadsheet.worksheet("História")
        existing = ws_hist.get_all_values()
    except gspread.exceptions.WorksheetNotFound:
        ws_hist = spreadsheet.add_worksheet("História", rows=100, cols=60)
        existing = []

    # Zostav row keys: "artikl_nr|store_canonical"
    row_keys = []
    for prod in products:
        for store in all_stores:
            row_keys.append({
                "key": f"{prod['artikel_nr']}|{store}",
                "name": prod["name"],
                "artikl": prod["artikel_nr"],
                "store": STORE_SHORT.get(store, store),
                "stock": prod["stores"].get(store, "–"),
            })

    if not existing or len(existing) < 2:
        # Prvý beh — vytvor celý sheet
        header_row = ["Produkt", "Artikl č.", "Predajňa", "Kľúč", date_col_header]
        data_rows = [header_row]
        for rk in row_keys:
            val = rk["stock"]
            data_rows.append([rk["name"], rk["artikl"], rk["store"], rk["key"],
                              val if val != "?" else "–"])
        ws_hist.update(data_rows, value_input_option="USER_ENTERED")
        print(f"  'História' inicializovaná ({len(row_keys)} riadkov)")
    else:
        # Existujúce dáta — pridaj nový stĺpec
        header_row = existing[0]

        # Skontroluj či dnes už bežal (rovnaký dátum)
        if date_col_header in header_row:
            col_idx = header_row.index(date_col_header)
            print(f"  'História' — dátum {date_col_header} už existuje, prepisujem stĺpec {col_idx+1}")
        else:
            col_idx = len(header_row)
            # Pridaj hlavičku nového stĺpca
            ws_hist.update_cell(1, col_idx + 1, date_col_header)
            print(f"  'História' — pridávam nový stĺpec {col_idx+1}: {date_col_header}")

        # Mapa existujúcich kľúčov na riadky
        key_col = 3  # stĺpec D (0-indexed: 3)
        existing_keys = {}
        for row_idx, row in enumerate(existing[1:], start=2):
            if len(row) > key_col:
                existing_keys[row[key_col]] = row_idx

        # Zapíš hodnoty
        cells_to_update = []
        new_rows = []
        for rk in row_keys:
            val = rk["stock"] if rk["stock"] != "?" else "–"
            if rk["key"] in existing_keys:
                row_num = existing_keys[rk["key"]]
                cells_to_update.append(gspread.Cell(row_num, col_idx + 1, val))
            else:
                # Nový produkt/predajňa — pridaj riadok
                new_row = [rk["name"], rk["artikl"], rk["store"], rk["key"]]
                # Doplň prázdne stĺpce pre predchádzajúce dátumy
                while len(new_row) < col_idx:
                    new_row.append("–")
                new_row.append(val)
                new_rows.append(new_row)

        if cells_to_update:
            ws_hist.update_cells(cells_to_update, value_input_option="USER_ENTERED")

        if new_rows:
            next_row = len(existing) + 1
            for nr in new_rows:
                ws_hist.update(f"A{next_row}", [nr], value_input_option="USER_ENTERED")
                next_row += 1

        print(f"  'História' aktualizovaná ({len(cells_to_update)} updated, {len(new_rows)} new)")

    # ── Sheet: Rozdiel ────────────────────────────────────────────────────
    # Porovná posledné 2 behy, zapíše rozdiel (záporné = predané)
    try:
        ws_hist_data = ws_hist.get_all_values()
    except Exception:
        ws_hist_data = []

    if ws_hist_data and len(ws_hist_data[0]) >= 6:  # Aspoň 2 dátumové stĺpce (4 fixné + 2)
        try:
            ws_diff = spreadsheet.worksheet("Rozdiel")
            ws_diff.clear()
        except gspread.exceptions.WorksheetNotFound:
            ws_diff = spreadsheet.add_worksheet("Rozdiel", rows=100, cols=15)

        header = ws_hist_data[0]
        last_col = len(header) - 1
        prev_col = last_col - 1

        # Len ak obe sú dátumové stĺpce (index >= 4)
        if prev_col >= 4:
            diff_header = ["Produkt", "Predajňa",
                           f"Stav {header[prev_col]}", f"Stav {header[last_col]}",
                           "Rozdiel", "Poznámka"]
            diff_rows = [diff_header]

            for row in ws_hist_data[1:]:
                name = row[0] if len(row) > 0 else ""
                store = row[2] if len(row) > 2 else ""
                prev_val = row[prev_col] if len(row) > prev_col else "–"
                curr_val = row[last_col] if len(row) > last_col else "–"

                try:
                    prev_n = int(str(prev_val).replace(" ", "").replace("–", ""))
                except ValueError:
                    prev_n = None
                try:
                    curr_n = int(str(curr_val).replace(" ", "").replace("–", ""))
                except ValueError:
                    curr_n = None

                if prev_n is not None and curr_n is not None:
                    diff = curr_n - prev_n
                    if diff < 0:
                        note = f"Predaných ~{abs(diff)} ks"
                    elif diff > 0:
                        note = f"Naskladnených {diff} ks"
                    else:
                        note = "Bez zmeny"
                else:
                    diff = "–"
                    note = "Nedostatok dát"

                diff_rows.append([name, store, prev_val, curr_val, diff, note])

            ws_diff.update(diff_rows, value_input_option="USER_ENTERED")
            print(f"  'Rozdiel' zapísaný ({len(diff_rows)-1} riadkov)")
        else:
            print("  'Rozdiel' — ešte nemám 2 behy na porovnanie")
    else:
        print("  'Rozdiel' — ešte nemám dosť dát")

    print(f"\nGoogle Sheets aktualizovaný: {now_str}")


# ── Main ──────────────────────────────────────────────────────────────────────
async def main():
    products = await scrape()
    if products:
        write_to_sheets(products)
    else:
        print("Žiadne produkty, sheets sa neaktualizujú.")
        sys.exit(1)


if __name__ == "__main__":
    asyncio.run(main())
