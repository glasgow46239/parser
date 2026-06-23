import os, sys, glob, argparse, traceback
from push_to_sheet import push

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--raw-dir', default='raw_tables')
    ap.add_argument('--archive-dir', default='archive')
    ap.add_argument('--aliases', default='crosstab_aliases.json')
    ap.add_argument('--collapse-rules', default='collapse_rules.json')
    ap.add_argument('--party-order', default='party_order.json')
    ap.add_argument('--question-labels', default='question_labels.json')
    ap.add_argument('--sheet-id', default=os.environ.get('GOOGLE_SHEET_ID'))
    ap.add_argument('--tab-name', default=os.environ.get('GOOGLE_SHEET_TAB', 'Review'))
    ap.add_argument('--tol', type=float, default=0.02)
    args = ap.parse_args()

    creds_json = os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON')
    if not creds_json:
        raise SystemExit("GOOGLE_SERVICE_ACCOUNT_JSON environment variable is not set")
    if not args.sheet_id:
        raise SystemExit("No sheet id provided (set GOOGLE_SHEET_ID or pass --sheet-id)")

    os.makedirs(args.archive_dir, exist_ok=True)
    files = sorted(glob.glob(os.path.join(args.raw_dir, '*.xlsx')))
    print(f"Found {len(files)} xlsx file(s) in {args.raw_dir}/")

    succeeded, failed = [], []
    for f in files:
        base = os.path.splitext(os.path.basename(f))[0]
        out_csv = os.path.join(args.archive_dir, base + '.csv')
        print(f"\n=== {f} ===")
        try:
            push(f, args.aliases, args.collapse_rules, args.party_order,
                 args.question_labels, out_csv,
                 args.sheet_id, args.tab_name, creds_json, tol=args.tol)
            succeeded.append(f)
        except Exception as e:
            print(f"FAILED: {f}: {e}")
            traceback.print_exc()
            failed.append((f, str(e)))

    print(f"\n\n===== SUMMARY: {len(succeeded)}/{len(files)} succeeded, {len(failed)} failed =====")
    if failed:
        print("Failed files (everything else still got processed and pushed):")
        for f, err in failed:
            print(f"  {f}: {err}")
        sys.exit(1)  # mark the run as needing attention, but only after trying every file

if __name__ == '__main__':
    main()
