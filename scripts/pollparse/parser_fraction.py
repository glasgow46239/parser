"""
parser_fraction.py

Handles pollster formats where values are stored as fractions (0.0–1.0) rather
than as integer raw counts paired with a separate percentage row.

Covers:
  - Find Out Now: one sheet per question, col A = label, data starts col B
  - More in Common: same layout but col A holds the question code on row 1;
    actual row labels are in col A from row 6 onwards once data begins
  - BMG: col A is blank, col B = label, 'Table X' markers in col B, data rows
    alternate raw_count_float / fraction / significance_code

The output shape matches parse_block() from parser.py:
  {
    'title': str,
    'columns': [{'idx': int, 'group': str|None, 'subgroup': str|None}],
    'base_weighted': list|None,   # always None for fraction files
    'base_unweighted': list|None, # None or extracted integers
    'records': [{'label': str, 'raw': None, 'pct': list}],
  }
"""

import re

# ── helpers ───────────────────────────────────────────────────────────────────

def _is_fraction(v):
    return isinstance(v, float) and 0.0 <= v <= 1.0

def _is_raw_count(v):
    """BMG stores raw counts as floats like 183.89 -- distinguishable from
    fractions because they are > 1 (unless a count genuinely rounds to 0 or 1,
    but those are handled fine as fractions anyway)."""
    return isinstance(v, float) and v > 1.0

def _row_has_fractions(row):
    nums = [v for v in row if isinstance(v, (int, float))]
    if len(nums) < 2:
        return False
    fractions = sum(1 for v in nums if _is_fraction(v))
    return fractions / len(nums) >= 0.6  # >60% of numeric cells are 0-1 fractions

def _row_is_sig_codes(row):
    """Significance code rows (BMG) look like 'V3,S3,T3...' strings."""
    non_empty = [v for v in row if v is not None and str(v).strip()]
    if not non_empty:
        return False
    strings = [v for v in non_empty if isinstance(v, str)]
    return len(strings) / len(non_empty) >= 0.5 and all(
        re.match(r'^[A-Z]\d', str(v).strip()[:2]) for v in strings
        if isinstance(v, str) and str(v).strip()
    )

def _clean_label(v):
    if v is None:
        return None
    s = str(v).strip()
    return s if s else None

def _forward_fill(vals):
    out = []
    last = None
    for v in vals:
        if v is not None and not (isinstance(v, str) and v.strip() == ''):
            last = v
        out.append(last)
    return out

# ── format detection ──────────────────────────────────────────────────────────

def is_fraction_format(rows):
    """
    Returns True if this block of rows uses the fraction (0-1) convention
    rather than integer-raw + fraction pairs.
    Check the first few data-looking rows.
    """
    fraction_rows = 0
    checked = 0
    for row in rows:
        nums = [v for v in row if isinstance(v, (int, float))]
        if len(nums) >= 2:
            checked += 1
            if _row_has_fractions(row):
                fraction_rows += 1
        if checked >= 6:
            break
    return checked > 0 and fraction_rows / checked >= 0.5


# ── Find Out Now / More in Common parser ─────────────────────────────────────

def _parse_fon_mic_block(title, rows):
    """
    Find Out Now and More in Common share almost the same layout:
      Row N-2: group headers (Gender, Age, Region, ...)   -- may be offset
      Row N-1: subgroup labels (All, Male, Female, ...)
      Row N+: data rows -- col A = label, col B onwards = fractions

    More in Common quirk: col A on row 1 holds the question code (same as
    sheet name); treat any col-A value that matches the title as a skip.
    """
    if not rows:
        return None

    # Find the header row: the one with the most non-None text values
    # that look like subgroup labels (not fractions)
    header_row_idx = None
    group_row_idx = None
    for i, row in enumerate(rows):
        text_vals = [v for v in row if isinstance(v, str) and v.strip()
                     and not _row_has_fractions(row)]
        num_vals = [v for v in row if isinstance(v, (int, float))]
        if text_vals and not num_vals and len(text_vals) >= 2:
            # candidate: a row with multiple text values, no numbers
            # pick the LAST such row before data begins as the subgroup row
            header_row_idx = i
        elif _row_has_fractions(row) and header_row_idx is not None:
            break

    if header_row_idx is None:
        return None

    # Look one row back for a group row
    if header_row_idx > 0:
        prev = rows[header_row_idx - 1]
        prev_text = [v for v in prev if isinstance(v, str) and v.strip()]
        if prev_text and not _row_has_fractions(prev):
            group_row_idx = header_row_idx - 1

    subgroup_row = rows[header_row_idx]
    group_row = rows[group_row_idx] if group_row_idx is not None else [None] * len(subgroup_row)
    group_row_ff = _forward_fill(list(group_row))

    # Build columns: skip col 0 (label column), start from col 1
    columns = []
    for col_idx in range(1, len(subgroup_row)):
        subgroup = subgroup_row[col_idx]
        if subgroup is None or (isinstance(subgroup, str) and not subgroup.strip()):
            continue
        group = group_row_ff[col_idx] if col_idx < len(group_row_ff) else None
        columns.append({
            'idx': col_idx,
            'group': _clean_label(group) if group != subgroup else None,
            'subgroup': _clean_label(subgroup),
        })

    if not columns:
        return None

    # Extract base (unweighted) from any 'Unweighted' / 'Base' / 'n=' row
    base_unweighted = None
    base_weighted = None

    records = []
    for row in rows[header_row_idx + 1:]:
        label = _clean_label(row[0]) if row else None
        if label is None:
            continue
        # Skip meta rows
        ll = label.lower()
        if any(k in ll for k in ('unweighted', 'base:', 'weighted total',
                                  'total base', 'n=')):
            # extract unweighted base if present
            if 'unweighted' in ll and base_unweighted is None:
                nums = [v for v in row[1:] if isinstance(v, int)]
                if nums:
                    base_unweighted = [None] + list(row[1:])
            continue
        if _row_is_sig_codes(row[1:]):
            continue
        if not _row_has_fractions(row):
            continue
        pct = [None] + list(row[1:])  # 1-indexed to match columns
        records.append({'label': label, 'raw': None, 'pct': pct})

    return {
        'title': title,
        'columns': columns,
        'base_weighted': base_weighted,
        'base_unweighted': base_unweighted,
        'records': records,
    }


# ── BMG parser ───────────────────────────────────────────────────────────────

def _parse_bmg_block(title, rows):
    """
    BMG format:
      Col A: always blank
      Col B: label (or ' ' for blank)
      Col C onwards: data
      Header structure:
        [blank, ' ', ' \n', 'Gender', ...]  <- group row
        [blank, ' ', 'Total', 'Male', ...]  <- subgroup row
        [blank, ' ', '(A)', '(B)', ...]     <- significance labels (skip)
        [blank, ' ']                        <- blank
        [blank, 'Unweighted row', n1, n2 ..]
        [blank, 'Base: Total', f1, f2, ...]
        [blank, 'PartyName', raw1, raw2, ...] <- raw count row (floats > 1)
        [blank, None, pct1, pct2, ...]        <- fraction row
        [blank, None, 'V3,S3,...', ...]       <- significance row (skip)
    """
    if not rows:
        return None

    header_row_idx = None
    group_row_idx = None

    for i, row in enumerate(rows):
        label_b = row[1] if len(row) > 1 else None
        label_s = (str(label_b).strip() if label_b is not None else '').lower()

        # Stop at the first sign of actual data so sig-code rows below don't
        # accidentally overwrite the real subgroup header_row_idx
        if 'unweighted' in label_s:
            break
        if any(_is_raw_count(v) or _is_fraction(v)
               for v in (row[2:] if len(row) > 2 else [])):
            break

        is_blank_b = (label_b is None
                      or (isinstance(label_b, str)
                          and label_b.strip() in ('', '\n', ' ', ' \n')))
        # Drop sig-code cells like (A),(B1),(AA) and dash characters
        _SIG = re.compile(r'^\([A-Z]+\d*\)$')
        text_from_c = [
            v for v in (row[2:] if len(row) > 2 else [])
            if isinstance(v, str) and v.strip()
            and not _SIG.match(v.strip())
            and v.strip() not in ('-', '\u2014')
        ]
        if is_blank_b and text_from_c and len(text_from_c) >= 2:
            header_row_idx = i


    if header_row_idx is None:
        return None

    # Look for group row one step before
    if header_row_idx > 0:
        prev = rows[header_row_idx - 1]
        prev_text = [v for v in (prev[2:] if len(prev) > 2 else [])
                     if isinstance(v, str) and v.strip()]
        if prev_text:
            group_row_idx = header_row_idx - 1

    subgroup_row = rows[header_row_idx]
    group_row = rows[group_row_idx] if group_row_idx is not None else [None] * len(subgroup_row)
    group_row_ff = _forward_fill(list(group_row))

    columns = []
    for col_idx in range(2, len(subgroup_row)):
        subgroup = subgroup_row[col_idx]
        if subgroup is None or (isinstance(subgroup, str) and not subgroup.strip()):
            continue
        # skip significance label columns like '(A)','(B)'
        if isinstance(subgroup, str) and re.match(r'^\([A-Z]+\d*\)$', subgroup.strip()):
            continue
        group = group_row_ff[col_idx] if col_idx < len(group_row_ff) else None
        columns.append({
            'idx': col_idx,
            'group': _clean_label(group) if group != subgroup else None,
            'subgroup': _clean_label(subgroup),
        })

    if not columns:
        return None

    base_unweighted = None
    base_weighted = None

    records = []
    data_rows = rows[header_row_idx + 1:]
    i = 0
    while i < len(data_rows):
        row = data_rows[i]
        label = _clean_label(row[1] if len(row) > 1 else None)

        if label is None or label.strip() in ('', ' ', '\n'):
            i += 1
            continue
        ll = label.lower()

        # Base rows
        if 'unweighted row' in ll:
            def _to_int(v):
                try: return int(round(float(v)))
                except (TypeError, ValueError): return None
            bases = [_to_int(v) for v in row[2:]]
            if any(b is not None for b in bases):
                base_unweighted = [None, None] + bases
            i += 1
            continue
        if 'base: total' in ll or 'base: all' in ll:
            i += 1
            continue
        if ll.startswith('comparison') or ll.startswith('independent') or ll.startswith('uppercase'):
            i += 1
            continue

        # Check if this is a data label row (next row will be the fractions)
        next_row = data_rows[i + 1] if i + 1 < len(data_rows) else None
        if next_row is not None and _row_has_fractions(next_row) and _clean_label(next_row[1]) is None:
            pct = [None, None] + list(next_row[2:])
            records.append({'label': label, 'raw': None, 'pct': pct})
            # skip the fraction row and any following significance row
            i += 2
            if i < len(data_rows) and _row_is_sig_codes(data_rows[i][2:]):
                i += 1
            continue

        # Fallback: the row itself might have fractions in col 2+
        if _row_has_fractions(row):
            pct = [None, None] + list(row[2:])
            records.append({'label': label, 'raw': None, 'pct': pct})
        i += 1

    return {
        'title': title,
        'columns': columns,
        'base_weighted': base_weighted,
        'base_unweighted': base_unweighted,
        'records': records,
    }


# ── public API ────────────────────────────────────────────────────────────────

def parse_fraction_block(title, rows, bmg_mode=False):
    """Main entry point -- routes to the right sub-parser."""
    if bmg_mode:
        return _parse_bmg_block(title, rows)
    return _parse_fon_mic_block(title, rows)
