import os, json, argparse
import gspread
from google.oauth2.service_account import Credentials
from build_archive_rows import run, TARGET_HEADER

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def get_worksheet(sheet_id, tab_name, creds_json, clear=False):
    creds_dict = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(tab_name)
        if clear:
            # Delete and recreate to actually free the allocated cells,
            # not just clear content (which leaves rows pre-allocated)
            sh.del_worksheet(ws)
            ws = sh.add_worksheet(title=tab_name, rows=1000, cols=len(TARGET_HEADER) + 1)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=tab_name, rows=1000, cols=len(TARGET_HEADER) + 1)
    return ws

def push(xlsx_path, aliases_path, collapse_rules_path, party_order_path, question_labels_path,
         out_csv, sheet_id, tab_name, creds_json, tol=0.02, clear_first=False):
    summary = run(xlsx_path, aliases_path, collapse_rules_path, party_order_path,
                  question_labels_path, out_csv, tol=tol)
    rows_with_flags = summary['rows_with_flags']
    if not rows_with_flags:
        print("No rows to push -- nothing written to the sheet.")
        return summary

    ws = get_worksheet(sheet_id, tab_name, creds_json, clear=clear_first)
    existing = ws.get_all_values()
    values = []
    if not existing:
        values.append(TARGET_HEADER + ['flags'])
    for row, flag in rows_with_flags:
        values.append([str(v) for v in row] + [flag])

    ws.append_rows(values, value_input_option='USER_ENTERED')
    flagged_count = sum(1 for _, f in rows_with_flags if f)
    print(f"Pushed {len(rows_with_flags)} rows to sheet tab '{tab_name}' ({flagged_count} with a flag set)")
    return summary

if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsx')
    ap.add_argument('-o', '--out', default='archive_rows.csv')
    ap.add_argument('--aliases', default='crosstab_aliases.json')
    ap.add_argument('--collapse-rules', default='collapse_rules.json')
    ap.add_argument('--party-order', default='party_order.json')
    ap.add_argument('--question-labels', default='question_labels.json')
    ap.add_argument('--tol', type=float, default=0.02)
    ap.add_argument('--sheet-id', default=os.environ.get('GOOGLE_SHEET_ID'))
    ap.add_argument('--tab-name', default=os.environ.get('GOOGLE_SHEET_TAB', 'Review'))
    args = ap.parse_args()

    creds_json = os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON')
    if not creds_json:
        raise SystemExit("GOOGLE_SERVICE_ACCOUNT_JSON environment variable is not set")
    if not args.sheet_id:
        raise SystemExit("No sheet id provided (set GOOGLE_SHEET_ID or pass --sheet-id)")

    push(args.xlsx, args.aliases, args.collapse_rules, args.party_order,
         args.question_labels, args.out,
         args.sheet_id, args.tab_name, creds_json, tol=args.tol)
