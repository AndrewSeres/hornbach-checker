"""
Hornbach Brikety Tracker – Multi-country CI scraper + Google Sheets
====================================================================
Beží v GitHub Actions každý piatok o 16:00.
Scrapuje SK + CZ + AT, zapisuje do Google Sheets.

Požiadavky: pip install playwright gspread google-auth
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


# ── Per-country config ─────────────────────────────────────────────────────────
COUNTRY_CONFIGS = [
    {
        "key": "SK",
        "category_url": "https://www.hornbach.sk/c/krby-radiatory-a-klimatizacie/palivove-drevo-pelety-a-brikety/S15929/",
        "keywords": ["briketa", "brikety", "brikiet", "briket", "pelet", "palivov", "drevo"],
        "exclude_keywords": ["hnedouholn", "uholn", "hnedouhol"],
        "locale": "sk-SK",
        "tab_prefix": "",          # empty = original tab names, backward-compatible
        "currency_re": r'(\d+[,\.]\d{2})\s*€',
    },
    {
        "key": "CZ",
        "category_url": "https://www.hornbach.cz/c/kamna-radiatory-a-klimatizace/brikety-pelety-a-palivove-drevo/S12498/",
        "keywords": ["briket", "pelet", "drevo"],  # drevo matches dřevo via NFKD
        "exclude_keywords": ["hnedouheln", "uheln"],   # catches hnědouhelné, uhelné via NFKD
        "locale": "cs-CZ",
        "tab_prefix": "CZ ",
        "currency_re": r'(\d+[,\.]\d{2})\s*(?:Kč|CZK|€)',
    },
    {
        "key": "AT",
        "category_url": "https://www.hornbach.at/c/heizen-klima-lueftung/holzbriketts-brennholz-heizpellets/S3028/",
        "keywords": ["brikett", "pellet", "brennholz"],
        "exclude_keywords": ["braunkohl"],             # catches Braunkohle, Braunkohlebrikett
        "locale": "de-AT",
        "tab_prefix": "AT ",
        "currency_re": r'(\d+[,\.]\d{2})\s*€',
    },
]

# Cookie button texts per language
COOKIE_SELECTORS = [
    "#onetrust-accept-btn-handler",
    "button:has-text('Prijať')", "button:has-text('Súhlasím')",
    "button:has-text('Prijat')", "button:has-text('Suhlasim')",
    "button:has-text('Akceptovat')",
    "button:has-text('Přijmout vše')", "button:has-text('Souhlasím')",
    "button:has-text('Alle akzeptieren')", "button:has-text('Akzeptieren')",
    "button:has-text('Alle Cookies akzeptieren')",
]

# Availability button texts per language
AVAIL_TEXTS = [
    # SK
    "SKONTROLOVAT DOSTUPNOST", "SKONTROLOVAŤ DOSTUPNOSŤ",
    "Skontrolovat dostupnost", "Skontrolovať dostupnosť",
    "dostupnost v predajni", "dostupnosť v predajni",
    # CZ
    "ZKONTROLOVAT DOSTUPNOST", "Zkontrolovat dostupnost",
    "dostupnost v prodejně", "DOSTUPNOST V PRODEJNĚ",
    # DE/AT
    "VERFÜGBARKEIT IN DER NÄHE PRÜFEN", "Verfügbarkeit in der Nähe prüfen",
    "VERFÜGBARKEIT PRÜFEN", "Verfügbarkeit prüfen",
    "IN DER NÄHE PRÜFEN", "In der Nähe prüfen",
    "Im Markt verfügbar", "Markt wechseln",
    "Verfügbarkeit in Märkten",
]

# EAN expand button texts
EAN_EXPAND_TEXTS = [
    "VIAC INFORMACII O VYROBKU", "VIAC INFORMÁCIÍ O VÝROBKU",
    "Viac informacii", "Viac informácií",
    "VÍCE INFORMACÍ O VÝROBKU", "Více informací",
    "MEHR INFORMATIONEN", "Mehr Informationen", "Produktdetails",
]

# Out-of-stock patterns (all languages)
OOS_RE = re.compile(
    r'nie je k dispoz|nedostupn|momentálne nie|momentalne nie|'
    r'vypredané|vypredane|0 balení|0 balenie|'
    r'není k dispozici|není skladem|momentálně není|'
    r'nicht verfügbar|ausverkauft|momentan nicht|nicht auf Lager|'
    r'0 Stück|0 St\.|0 ST\b',
    re.IGNORECASE
)

STOCK_RE = re.compile(
    r'(\d[\d\s]*)\s*(balen[ií]e?|ks\b|kus|Stück\b|Stk\b|St\.|ST\b)',
    re.IGNORECASE
)

STOCK_LABEL_RE = re.compile(
    r'dostupn[éeý]\s*>?\s*(\d+)|verfügbar[:\s]*(\d+)',
    re.IGNORECASE
)


# ── Helpers ────────────────────────────────────────────────────────────────────
def matches_keywords(text: str, keywords: list, exclude_keywords: list) -> bool:
    t = text.lower()
    t_plain = unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode("ascii")
    if not any(k in t or k in t_plain for k in keywords):
        return False
    if any(k in t or k in t_plain for k in exclude_keywords):
        return False
    return True


def canonicalize_store(raw_name: str) -> str | None:
    """Clean store name from modal line. Returns None for non-store lines."""
    raw_lower = raw_name.lower().strip()
    skip_fragments = [
        "etky predajne", "vsetky predajne",
        "vsechny prodejny", "alle filialen", "alle markte", "alle märkte",
        "vyhledat prodejnu", "markt suchen", "markt wechseln",
    ]
    if any(f in raw_lower for f in skip_fragments):
        return None
    if "hornbach" not in raw_lower:
        return None
    # Take up to first comma, strip trailing punctuation
    base = raw_name.split(",")[0].strip().rstrip(".")
    return re.sub(r'\s+', ' ', base).strip()


def store_short_name(canonical: str) -> str:
    """Strip HORNBACH prefix for display."""
    n = canonical.replace("HORNBACH", "").strip().lstrip("- ").strip()
    return re.sub(r'\s+', ' ', n).strip() or canonical


# ── Scraper ────────────────────────────────────────────────────────────────────
async def scrape_country(context, config: dict) -> list:
    products = []
    key = config["key"]
    keywords = config["keywords"]
    exclude_kw = config["exclude_keywords"]
    base_url = f"https://www.hornbach.{key.lower()}"

    print(f"\n{'='*60}")
    print(f"  Krajina: {key}  |  {config['category_url']}")
    print(f"{'='*60}")

    page = await context.new_page()
    await page.goto(config["category_url"], wait_until="networkidle", timeout=60000)

    # Dismiss cookie banner
    for sel in COOKIE_SELECTORS:
        try:
            btn = page.locator(sel).first
            if await btn.is_visible(timeout=2000):
                await btn.click()
                await page.wait_for_timeout(800)
                break
        except Exception:
            pass

    # Lazy-load products by scrolling
    for _ in range(8):
        await page.evaluate("window.scrollBy(0, 800)")
        await page.wait_for_timeout(600)
    await page.evaluate("window.scrollTo(0, 0)")
    await page.wait_for_timeout(500)

    all_links = await page.query_selector_all("a[href*='/p/']")
    print(f"  Linkov na stranke: {len(all_links)}")

    seen_urls: set[str] = set()
    product_links = []

    for link in all_links:
        try:
            href = await link.get_attribute("href")
            if not href or '/p/' not in href:
                continue
            full_url = href if href.startswith("http") else f"{base_url}{href}"
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
                                     el.closest('[class*="Tile"]') || el.closest('[class*="tile"]') ||
                                     el.closest('[class*="product"]') || el.parentElement;
                        return tile ? tile.innerText : '';
                    }""", link)
                    text = (text or "").strip()
                except Exception:
                    pass

            if not text:
                continue

            if matches_keywords(text, keywords, exclude_kw):
                seen_urls.add(full_url)
                name = next(
                    (l.strip() for l in text.splitlines() if len(l.strip()) > 8),
                    text[:80]
                )
                product_links.append({"url": full_url, "name": name[:100]})
                print(f"  + {name[:65]}")
        except Exception:
            pass

    await page.close()
    print(f"\n  {len(product_links)} produktov na spracovanie")

    if not product_links:
        print("  Ziadne produkty.")
        return products

    for idx, prod in enumerate(product_links):
        print(f"\n  [{idx+1}/{len(product_links)}] {prod['name']}")

        prod_data = {
            "name": prod["name"],
            "ean": "",
            "artikel_nr": "",
            "price": "",
            "url": prod["url"],
            "stores": {},
        }

        m = re.search(r'/(\d{5,})/?', prod["url"])
        if m:
            prod_data["artikel_nr"] = m.group(1)

        pp = None
        try:
            pp = await context.new_page()
            await pp.goto(prod["url"], wait_until="domcontentloaded", timeout=45000)
            await pp.wait_for_timeout(2000)

            # Price
            for price_sel in [
                "[data-testid='article-price'] [class*='value']",
                "[class*='article-price'] [class*='value']",
                "[class*='ArticlePrice'] [class*='value']",
                "[class*='price-value']",
                "[class*='Price'] strong",
                "[class*='price'] strong",
                "strong[class*='price']",
            ]:
                try:
                    el = pp.locator(price_sel).first
                    if await el.is_visible(timeout=800):
                        prod_data["price"] = (await el.inner_text()).strip()
                        break
                except Exception:
                    pass

            if not prod_data["price"]:
                try:
                    body_text = await pp.inner_text("body")
                    pm = re.search(config["currency_re"], body_text)
                    if pm:
                        prod_data["price"] = pm.group(0).strip()
                except Exception:
                    pass

            # Expand EAN section
            for expand_text in EAN_EXPAND_TEXTS:
                try:
                    btn = pp.get_by_text(expand_text, exact=False).first
                    if await btn.is_visible(timeout=800):
                        await btn.click()
                        await pp.wait_for_timeout(700)
                        break
                except Exception:
                    pass

            # EAN from table
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
                    em = re.search(r'EAN\s*\n?\s*([0-9]{8,}(?:[,\s]+[0-9]{8,})*)', body)
                    if em:
                        prod_data["ean"] = em.group(1).strip()
                except Exception:
                    pass

            print(f"  EAN: {prod_data['ean'] or '-'}  Cena: {prod_data['price'] or '-'}")

            # Click availability button
            clicked = False
            for avail_text in AVAIL_TEXTS:
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

            # Read modal text
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

            # Parse stores from modal
            if modal_text and "HORNBACH" in modal_text:
                lines = [l.strip() for l in modal_text.splitlines() if l.strip()]
                found_stores: dict[str, int | str] = {}
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
                                # "123 balení / Stück / ks"
                                sm = STOCK_RE.search(jline)
                                if sm:
                                    stock_val = int(sm.group(1).replace(" ", ""))
                                    break
                                # Bare number with contextual confirmation
                                if re.match(r'^\d+$', jline):
                                    nearby = " ".join(lines[max(0, j-2):min(len(lines), j+3)]).lower()
                                    if any(w in nearby for w in [
                                        "dostupn", "skladom", "dispoz", "baleni", "balení",
                                        "verfügb", "lager", "vorrätig", "k dispozici",
                                    ]):
                                        stock_val = int(jline)
                                        break
                                # Out-of-stock phrase
                                if OOS_RE.search(jline):
                                    stock_val = 0
                                    break
                                # "Dostupné > 50" / "verfügbar: 12"
                                sm2 = STOCK_LABEL_RE.search(jline)
                                if sm2:
                                    stock_val = int(sm2.group(1) or sm2.group(2))
                                    break

                            if canonical in found_stores:
                                if found_stores[canonical] == "?" and stock_val is not None:
                                    found_stores[canonical] = stock_val
                            else:
                                found_stores[canonical] = stock_val if stock_val is not None else "?"

                            print(f"  {store_short_name(canonical)}: {found_stores[canonical]}")
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

    print(f"\n  {key}: {len(products)} produktov hotovo")
    return products


# ── Google Sheets ──────────────────────────────────────────────────────────────
def get_sheets_client():
    creds_json = os.environ.get("GOOGLE_SHEETS_CREDS")
    if creds_json:
        creds_dict = json.loads(creds_json)
    else:
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


def write_to_sheets(products: list, spreadsheet, tab_prefix: str = ""):
    """Write one country's results to its three sheet tabs."""
    if not products:
        return

    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")
    date_col_header = datetime.now().strftime("%d.%m.%Y %H:%M")

    # Discover stores in order of first appearance (no hardcoded list needed)
    seen_stores: list[str] = []
    for p in products:
        for s in p["stores"]:
            if s not in seen_stores:
                seen_stores.append(s)
    all_stores = seen_stores

    tab_current = f"{tab_prefix}Posledný beh".strip()
    tab_hist    = f"{tab_prefix}História".strip()
    tab_diff    = f"{tab_prefix}Rozdiel".strip()

    # ── Posledný beh ──────────────────────────────────────────────────────────
    try:
        ws_current = spreadsheet.worksheet(tab_current)
        ws_current.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws_current = spreadsheet.add_worksheet(tab_current, rows=200, cols=30)

    headers = ["Produkt", "Artikl č.", "Cena"] + [store_short_name(s) for s in all_stores]
    rows = [headers]
    for prod in products:
        row = [prod["name"], prod["artikel_nr"], prod["price"]]
        for store in all_stores:
            val = prod["stores"].get(store, "–")
            row.append(val if val != "?" else "–")
        rows.append(row)

    ws_current.update(rows, value_input_option="USER_ENTERED")
    print(f"  '{tab_current}' → {len(products)} produktov")

    # ── História ───────────────────────────────────────────────────────────────
    try:
        ws_hist = spreadsheet.worksheet(tab_hist)
        existing = ws_hist.get_all_values()
        # Grow sheet if it's too small (fixes sheets created with old low defaults)
        needed_rows = max(1000, len(existing) + len(products) * 20)
        needed_cols = max(120, ws_hist.col_count)
        if ws_hist.row_count < needed_rows or ws_hist.col_count < needed_cols:
            ws_hist.resize(rows=needed_rows, cols=needed_cols)
    except gspread.exceptions.WorksheetNotFound:
        ws_hist = spreadsheet.add_worksheet(tab_hist, rows=1000, cols=120)
        existing = []

    row_keys = [
        {
            "key": f"{prod['artikel_nr']}|{store}",
            "name": prod["name"],
            "artikl": prod["artikel_nr"],
            "store": store_short_name(store),
            "stock": prod["stores"].get(store, "–"),
        }
        for prod in products
        for store in all_stores
    ]

    if not existing or len(existing) < 2:
        data_rows = [["Produkt", "Artikl č.", "Predajňa", "Kľúč", date_col_header]]
        for rk in row_keys:
            val = rk["stock"]
            data_rows.append([rk["name"], rk["artikl"], rk["store"], rk["key"],
                               val if val != "?" else "–"])
        ws_hist.update(data_rows, value_input_option="USER_ENTERED")
        print(f"  '{tab_hist}' → inicializovaná ({len(row_keys)} riadkov)")
    else:
        header_row = existing[0]
        if date_col_header in header_row:
            col_idx = header_row.index(date_col_header)
            print(f"  '{tab_hist}' → prepisujem stĺpec {col_idx+1} ({date_col_header})")
        else:
            col_idx = len(header_row)
            ws_hist.update_cell(1, col_idx + 1, date_col_header)
            print(f"  '{tab_hist}' → nový stĺpec {col_idx+1}: {date_col_header}")

        existing_keys = {
            row[3]: row_idx
            for row_idx, row in enumerate(existing[1:], start=2)
            if len(row) > 3
        }

        cells_to_update = []
        new_rows = []
        for rk in row_keys:
            val = rk["stock"] if rk["stock"] != "?" else "–"
            if rk["key"] in existing_keys:
                cells_to_update.append(gspread.Cell(existing_keys[rk["key"]], col_idx + 1, val))
            else:
                new_row = [rk["name"], rk["artikl"], rk["store"], rk["key"]]
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

        print(f"  '{tab_hist}' → {len(cells_to_update)} updated, {len(new_rows)} new")

    # ── Rozdiel ────────────────────────────────────────────────────────────────
    try:
        ws_hist_data = ws_hist.get_all_values()
    except Exception:
        ws_hist_data = []

    if ws_hist_data and len(ws_hist_data[0]) >= 6:
        try:
            ws_diff = spreadsheet.worksheet(tab_diff)
            ws_diff.clear()
        except gspread.exceptions.WorksheetNotFound:
            ws_diff = spreadsheet.add_worksheet(tab_diff, rows=1000, cols=15)

        header = ws_hist_data[0]
        last_col = len(header) - 1
        prev_col = last_col - 1

        if prev_col >= 4:
            diff_rows = [[
                "Produkt", "Predajňa",
                f"Stav {header[prev_col]}", f"Stav {header[last_col]}",
                "Rozdiel", "Poznámka",
            ]]

            for row in ws_hist_data[1:]:
                name      = row[0] if len(row) > 0 else ""
                store     = row[2] if len(row) > 2 else ""
                prev_val  = row[prev_col] if len(row) > prev_col else "–"
                curr_val  = row[last_col] if len(row) > last_col else "–"

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
                    note = (
                        f"Predaných ~{abs(diff)} ks" if diff < 0 else
                        f"Naskladnených {diff} ks" if diff > 0 else
                        "Bez zmeny"
                    )
                else:
                    diff, note = "–", "Nedostatok dát"

                diff_rows.append([name, store, prev_val, curr_val, diff, note])

            ws_diff.update(diff_rows, value_input_option="USER_ENTERED")
            print(f"  '{tab_diff}' → {len(diff_rows)-1} riadkov")
        else:
            print(f"  '{tab_diff}' → ešte nemám 2 behy na porovnanie")
    else:
        print(f"  '{tab_diff}' → ešte nemám dosť dát")

    print(f"  Google Sheets aktualizovaný: {now_str}")


# ── Main ───────────────────────────────────────────────────────────────────────
async def main():
    sheet_id = os.environ.get("GOOGLE_SHEET_ID")
    if not sheet_id:
        print("CHYBA: GOOGLE_SHEET_ID env nie je nastavený")
        sys.exit(1)

    gc = get_sheets_client()
    spreadsheet = gc.open_by_key(sheet_id)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        any_products = False

        for config in COUNTRY_CONFIGS:
            context = await browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
                viewport={"width": 1280, "height": 900},
                locale=config["locale"],
            )
            try:
                products = await scrape_country(context, config)
                if products:
                    write_to_sheets(products, spreadsheet, tab_prefix=config["tab_prefix"])
                    any_products = True
                else:
                    print(f"  {config['key']}: Žiadne produkty, sheets sa neaktualizujú.")
            finally:
                await context.close()

        await browser.close()

    if not any_products:
        sys.exit(1)


if __name__ == "__main__":
    asyncio.run(main())
