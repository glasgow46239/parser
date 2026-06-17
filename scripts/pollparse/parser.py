from openpyxl import load_workbook
import json, sys, csv

def cell_kind(v):
    if v is None:
        return 'empty'
    if isinstance(v, bool):
        return 'other'
    if isinstance(v, int):
        return 'int'
    if isinstance(v, float):
        return 'float'
    if isinstance(v, str):
        s = v.strip()
        if s in ('', '-', 'N/A', 'n/a'):
            return 'empty'
        return 'other'
    return 'other'

def row_kind(values):
    kinds = [cell_kind(v) for v in values]
    n_int = kinds.count('int')
    n_float = kinds.count('float')
    if n_int == 0 and n_float == 0:
        return 'empty'
    return 'raw' if n_int >= n_float else 'pct'

def is_blank_row(row):
    return all(cell_kind(v) == 'empty' for v in row)

def label_of(row):
    if not row:
        return None
    v = row[0]
    if v is None:
        return None
    if isinstance(v, str) and v.strip() == '':
        return None
    return v

BASE_MARKERS = ('unweighted total', 'weighted total')

def is_base_row(row):
    lab = label_of(row)
    if not lab or not isinstance(lab, str):
        return False
    l = lab.lower()
    return l.startswith('base:') or l in BASE_MARKERS

def forward_fill(vals):
    out = []
    last = None
    for v in vals:
        if v is not None and not (isinstance(v, str) and v.strip() == ''):
            last = v
        out.append(last)
    return out

def parse_block(title, rows):
    """rows: list of tuples (row_idx, values-list) starting right after title row, ending at block end."""
    i = 0
    n = len(rows)
    # skip a 'Base: All Respondents' style descriptive line and blank lines before header rows
    while i < n and not is_base_row(rows[i][1]) and is_blank_row(rows[i][1]):
        i += 1
    header_rows = []
    base_rows = []
    # gather leading non-numeric "Base: xxx" descriptor line (Survation style, e.g. 'Base: All Respondents')
    while i < n:
        idx, vals = rows[i]
        lab = label_of(vals)
        if isinstance(lab, str) and lab.lower().startswith('base:') and row_kind(vals[1:]) == 'empty':
            i += 1
            continue
        if is_blank_row(vals):
            i += 1
            continue
        break
    # gather header rows: rows with empty col0 and some text in other cells, until a base row appears
    while i < n and not is_base_row(rows[i][1]):
        idx, vals = rows[i]
        if is_blank_row(vals):
            i += 1
            continue
        header_rows.append(vals)
        i += 1
    # gather base rows (Unweighted/Weighted total, possibly under 'Base:' labels)
    while i < n and is_base_row(rows[i][1]):
        base_rows.append(rows[i][1])
        i += 1
    # determine column count
    ncols = max((len(v) for v in header_rows + base_rows), default=0)
    # build column labels: support 1 or 2 header rows
    if len(header_rows) >= 2:
        group_row = forward_fill(header_rows[0])
        sub_row = header_rows[1]
    elif len(header_rows) == 1:
        group_row = [None] * ncols
        sub_row = header_rows[0]
    else:
        group_row = [None] * ncols
        sub_row = [None] * ncols
    columns = []
    for c in range(1, ncols):
        s_raw = sub_row[c] if c < len(sub_row) else None
        if s_raw is None:
            # no genuine subgroup label at this column -> padding/formatting artifact, skip
            continue
        g = group_row[c] if c < len(group_row) else None
        columns.append({'idx': c, 'group': g, 'subgroup': s_raw})
    # bases
    base_unweighted = None
    base_weighted = None
    for vals in base_rows:
        lab = str(label_of(vals)).lower()
        if 'unweighted' in lab:
            base_unweighted = vals
        elif 'weighted' in lab:
            base_weighted = vals
    # data rows: from i to end (caller already trimmed block to exclude trailing blank/footer)
    data_rows = [(idx, vals) for idx, vals in rows[i:]]
    # pair rows
    records = []
    j = 0
    m = len(data_rows)
    while j < m:
        idx, vals = data_rows[j]
        lab = label_of(vals)
        kind = row_kind(vals[1:])
        if kind == 'empty' or lab is None:
            j += 1
            continue
        raw_vals = pct_vals = None
        if j + 1 < m:
            idx2, vals2 = data_rows[j+1]
            lab2 = label_of(vals2)
            kind2 = row_kind(vals2[1:])
            if lab2 is None and kind2 != 'empty' and kind2 != kind:
                if kind == 'raw':
                    raw_vals, pct_vals = vals, vals2
                else:
                    pct_vals, raw_vals = vals, vals2
                j += 2
                records.append({'row': idx, 'label': lab, 'raw': raw_vals, 'pct': pct_vals})
                continue
        if kind == 'raw':
            raw_vals = vals
        else:
            pct_vals = vals
        records.append({'row': idx, 'label': lab, 'raw': raw_vals, 'pct': pct_vals})
        j += 1
    return {
        'title': title,
        'columns': columns,
        'base_unweighted': base_unweighted,
        'base_weighted': base_weighted,
        'records': records,
    }

def validate_block(block, tol=0.02):
    issues = []
    bw = block['base_weighted']
    if bw is None:
        return issues
    for rec in block['records']:
        if rec['raw'] is None or rec['pct'] is None:
            continue
        for col in block['columns']:
            c = col['idx']
            if c >= len(rec['raw']) or c >= len(rec['pct']) or c >= len(bw):
                continue
            raw = rec['raw'][c]
            pct = rec['pct'][c]
            base = bw[c]
            if cell_kind(raw) != 'int' or cell_kind(pct) != 'float' or cell_kind(base) != 'int':
                continue
            if base == 0:
                continue
            expected = raw / base
            diff = abs(expected - pct)
            if diff > tol:
                issues.append({
                    'row': rec['row'], 'label': rec['label'], 'col': col['subgroup'],
                    'raw': raw, 'base': base, 'expected_pct': round(expected,4),
                    'stated_pct': pct, 'diff': round(diff,4)
                })
    return issues
