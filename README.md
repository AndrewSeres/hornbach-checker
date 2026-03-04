# Hornbach Brikety Tracker 🔥

Automatický scraper dostupnosti drevených brikiet v HORNBACH predajniach na Slovensku.

## Čo to robí

- Každý piatok o ~16:00 scrapne kategóriu brikiet na hornbach.sk
- Zapíše stav zásob do Google Sheets (5 predajní × N produktov)
- Porovná s predchádzajúcim behom a vypočíta rozdiel (predané/naskladnené kusy)

## Google Sheets štruktúra

| Sheet | Obsah |
|---|---|
| **Posledný beh** | Aktuálny stav — produkt × predajne |
| **História** | Všetky behy — každý stĺpec = dátum, riadky = produkt+predajňa |
| **Rozdiel** | Porovnanie posledných 2 behov s komentárom |

## Setup

### 1. Google Cloud Service Account

1. [Google Cloud Console](https://console.cloud.google.com/) → nový projekt
2. APIs & Services → Library → **Google Sheets API** → Enable
3. Credentials → Create Credentials → **Service Account** → vytvor
4. Klikni na account → Keys → Add Key → **JSON** → stiahne sa súbor
5. Vytvor Google Sheet, Share ho s `client_email` z JSON-u ako **Editor**
6. Skopíruj **Sheet ID** z URL (`/d/TOTO_JE_ID/edit`)

### 2. GitHub Secrets

V repo Settings → Secrets and variables → Actions:

- `GOOGLE_SHEETS_CREDS` = celý obsah JSON súboru
- `GOOGLE_SHEET_ID` = ID sheetu

### 3. Spustenie

- **Automaticky**: každý piatok ~16:00 CET
- **Manuálne**: Actions → Hornbach Brikety Scraper → Run workflow
- **Lokálne**: 
  ```bash
  export GOOGLE_SHEETS_CREDS=$(cat service_account.json)
  export GOOGLE_SHEET_ID=tvoje_sheet_id
  pip install -r requirements.txt
  playwright install chromium
  python scraper.py
  ```
