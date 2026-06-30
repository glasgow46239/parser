#!/usr/bin/env python3
"""
wiki_polls_scraper.py

Fetches the Wikipedia opinion polling table for the next UK general election,
cleans it into a structured format, and pushes it to a Google Sheet tab.

Usage:
    python wiki_polls_scraper.py [--tab-name "Polls tracker"] [--sheet-id ID]

Env vars required:
    GOOGLE_SERVICE_ACCOUNT_JSON
    GOOGLE_SHEET_ID
"""

import os, re, json, argparse, requests
from datetime import datetime
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials

WIKI_URL = ('https://en.wikipedia.org/wiki/'
            'Opinion_polling_for_the_next_United_Kingdom_general_election')

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Canonical output order, matching the BENMA party ordering
OUTPUT_COLS = ['Date', 'Pollster', 'Client', 'Area', 'Sample',
               'Con', 'Lab', 'LDem', 'Ref', 'Grn', 'SNP', 'PC',
               'Oth', 'Lead', 'Notes', 'scraped_at']

PARTY_MAP = {
    'con': 'Con', 'conservative': 'Con', 'tory': 'Con',
    'lab': 'Lab', 'labour': 'Lab',
    'ld': 'LDem', 'lib dem': 'LDem', 'lib dems': 'LDem',
    'liberal democrat': 'LDem', 'liberal democrats': 'LDem',
    'ref': 'Ref', 'reform': 'Ref', 'reform uk': 'Ref',
    'grn': 'Grn', 'green': 'Grn', 'greens': 'Grn',
    'snp': 'SNP', 'scottish national party': 'SNP',
    'pc': 'PC', 'plaid cymru': 'PC',
}

def _pct(raw):
    """'18%' → '18' (strip % and whitespace)."""
    if raw is None:
        return ''
    s = str(raw).strip().replace('%', '').strip()
    return s if s and s not in ('—', '-', '–', '−') else ''

def _parse_others(raw):
    """
    'Others' cells often look like '6% RB 3% YP 1% Other 2%' or just '4%'.
    Returns (canonical_pct_str, notes_str).
    The canonical value is the first bare percentage (total others).
    The notes string captures any named-party breakdown.
    """
    if not raw:
        return '', ''
    raw = raw.strip()
    # Simple case: just a percentage
    simple = re.match(r'^(\d+)%?$', raw)
    if simple:
        return simple.group(1), ''

    # Find the first percentage (total) and retain the rest as notes
    m = re.match(r'^(\d+)%?\s*(.*)', raw)
    if m:
        return m.group(1), m.group(2).strip()
    return '', raw

def _clean_lead(raw):
    """'Reform UK - 7' or 'Tie' → just the signed numeric or 'Tie'."""
    if not raw:
        return ''
    raw = str(raw).strip()
    if raw.lower() == 'tie':
        return 'Tie'
    # Wikipedia lead is usually just a number (positive = first party leads)
    m = re.search(r'-?\d+', raw)
    return m.group(0) if m else raw

def fetch_table(year=None):
    """
    Fetches all polling rows from the Wikipedia page.
    If year is given (e.g. 2026), only rows from that calendar year are returned.
    """
    resp = requests.get(WIKI_URL, headers={'User-Agent': 'BENMA-poll-scraper/1.0'}, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.content, 'html.parser')

    tables = soup.find_all('table', class_='wikitable')
    if not tables:
        raise RuntimeError("No wikitable found on the page -- layout may have changed")

    rows_out = []
    for table in tables:
        thead = table.find('thead') or table
        header_cells = thead.find_all(['th'])
        headers = [h.get_text(strip=True).lower() for h in header_cells]

        # Only process tables with the expected polling columns
        if not ('pollster' in headers and 'lab' in headers and 'con' in headers):
            continue

        col_idx = {h: i for i, h in enumerate(headers)}

        for tr in table.find_all('tr'):
            cells = tr.find_all(['td', 'th'])
            if len(cells) < 5:
                continue
            texts = [c.get_text(separator=' ', strip=True) for c in cells]

            date_raw = texts[col_idx.get('date(s) conducted', 0)]
            # Skip event rows (e.g. "26 Feb: Gorton and Denton by-election")
            if re.match(r'\d+\s+\w+:.*election', date_raw, re.IGNORECASE):
                continue
            if not re.search(r'\d', date_raw):
                continue

            def get(key, default=''):
                i = col_idx.get(key)
                return texts[i] if i is not None and i < len(texts) else default

            oth_raw = get('others')
            oth_pct, oth_notes = _parse_others(oth_raw)

            row = {
                'Date':     date_raw,
                'Pollster': get('pollster'),
                'Client':   get('client'),
                'Area':     get('area'),
                'Sample':   get('sample size').replace(',', ''),
                'Con':      _pct(get('con')),
                'Lab':      _pct(get('lab')),
                'LDem':     _pct(get('ld')),
                'Ref':      _pct(get('ref')),
                'Grn':      _pct(get('grn')),
                'SNP':      _pct(get('snp')),
                'PC':       _pct(get('pc')),
                'Oth':      oth_pct,
                'Lead':     _clean_lead(get('lead')),
                'Notes':    oth_notes,
                'scraped_at': datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ'),
            }
            rows_out.append(row)

    return rows_out


def push_to_sheet(rows, sheet_id, tab_name, creds_json):
    creds = Credentials.from_service_account_info(json.loads(creds_json), scopes=SCOPES)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(tab_name)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=tab_name, rows=2000, cols=len(OUTPUT_COLS))

    values = [OUTPUT_COLS]
    for row in rows:
        values.append([row.get(c, '') for c in OUTPUT_COLS])

    ws.update(values, value_input_option='USER_ENTERED')
    print(f"Wrote {len(rows)} polling rows to tab '{tab_name}'")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--tab-name', default=os.environ.get('WIKI_SHEET_TAB', 'Polls tracker'))
    ap.add_argument('--sheet-id', default=os.environ.get('GOOGLE_SHEET_ID'))
    ap.add_argument('--dry-run', action='store_true',
                    help='Print rows to stdout instead of writing to the sheet')
    args = ap.parse_args()

    print(f"Fetching {WIKI_URL} …")
    rows = fetch_table()
    print(f"Parsed {len(rows)} polling rows")

    if args.dry_run:
        import csv, sys
        w = csv.DictWriter(sys.stdout, fieldnames=OUTPUT_COLS)
        w.writeheader()
        w.writerows(rows)
        return

    creds_json = os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON')
    if not creds_json:
        raise SystemExit("GOOGLE_SERVICE_ACCOUNT_JSON not set")
    if not args.sheet_id:
        raise SystemExit("GOOGLE_SHEET_ID not set")

    push_to_sheet(rows, args.sheet_id, args.tab_name, creds_json)


if __name__ == '__main__':
    main()
