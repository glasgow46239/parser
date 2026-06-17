def build_merge_map(rules):
    m = {}
    for rule in rules.get('merge', []):
        for f in rule['from']:
            m[f.strip().lower()] = rule['into']
    return m

def collapse_records(records, columns, base_weighted, rules, tol=0.02):
    """Returns (collapsed_records, net_check_warnings).
    collapsed_records: list of {'label','raw':{idx:val},'pct':{idx:val}} dicts,
    using a dict keyed by column idx (sparser/clearer than fixed-width lists here).
    net_check_warnings: cases where a dropped NET row disagreed with our own
    recombined total by more than `tol` -- signals a merge rule is missing a
    label variant, not just sampling noise.
    """
    def norm(s):
        return (s or '').replace('\u2019', "'").replace('\u2018', "'").strip().lower()

    drop_patterns = [norm(p) for p in rules.get('drop_patterns', [])]
    merge_map = {norm(f): rule['into'] for rule in rules.get('merge', []) for f in rule['from']}
    col_idxs = [c['idx'] for c in columns]

    order = []
    merged_sum = {}  # into_label -> {idx: raw_sum}
    passthrough = {}
    net_rows = []  # stash NET rows we're about to drop, for the cross-check

    for rec in records:
        label = rec['label'] or ''
        ll = norm(label)
        if ll.startswith('net') or ll.startswith('net:'):
            net_rows.append(rec)
            continue
        if any(p in ll for p in drop_patterns):
            continue
        if ll in merge_map:
            target = merge_map[ll]
            if target not in merged_sum:
                merged_sum[target] = {idx: 0 for idx in col_idxs}
                order.append(('m', target))
            raw = rec['raw']
            if raw:
                for idx in col_idxs:
                    v = raw[idx] if idx < len(raw) else None
                    if isinstance(v, int):
                        merged_sum[target][idx] += v
            continue
        passthrough[label] = rec
        order.append(('p', label))

    collapsed = []
    seen = set()
    for kind, key in order:
        if (kind, key) in seen:
            continue
        seen.add((kind, key))
        if kind == 'm':
            raws = merged_sum[key]
            pct = {}
            for idx in col_idxs:
                b = base_weighted[idx] if base_weighted and idx < len(base_weighted) else None
                r = raws[idx]
                pct[idx] = round(r / b, 4) if isinstance(b, int) and b > 0 else None
            collapsed.append({'label': key, 'raw': raws, 'pct': pct})
        else:
            rec = passthrough[key]
            raw = {idx: (rec['raw'][idx] if rec['raw'] and idx < len(rec['raw']) else None) for idx in col_idxs}
            pct = {idx: (rec['pct'][idx] if rec['pct'] and idx < len(rec['pct']) else None) for idx in col_idxs}
            collapsed.append({'label': key, 'raw': raw, 'pct': pct})

    # sanity check: does our recombined Approve+Disapprove etc. match the pollster's
    # own NET row, before we throw that NET row away?
    warnings = []
    for net_rec in net_rows:
        net_label_clean = (net_rec['label'] or '')
        for prefix in ('net:', 'net'):
            if net_label_clean.lower().strip().startswith(prefix):
                net_label_clean = net_label_clean.strip()[len(prefix):]
                break
        net_label_clean = net_label_clean.strip().lower()
        for c in collapsed:
            if c['label'].lower() == net_label_clean:
                for idx in col_idxs:
                    stated = net_rec['pct'][idx] if net_rec['pct'] and idx < len(net_rec['pct']) else None
                    ours = c['pct'].get(idx)
                    if isinstance(stated, float) and isinstance(ours, float):
                        diff = abs(stated - ours)
                        if diff > tol:
                            warnings.append({
                                'net_label': net_rec['label'], 'our_label': c['label'],
                                'col_idx': idx, 'stated_net_pct': stated,
                                'our_recombined_pct': ours, 'diff': round(diff, 4)
                            })
    return collapsed, warnings
