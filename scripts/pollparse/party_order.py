def _match(label, aliases):
    ll = (label or '').strip().lower()
    for canon, variants in aliases.items():
        if ll in variants:
            return canon
    # fallback: substring containment for longer, unambiguous variants only
    for canon, variants in aliases.items():
        for v in variants:
            if len(v) >= 4 and v in ll:
                return canon
    return None

def _is_ni_party(label, ni_parties):
    ll = (label or '').strip().lower()
    return any(ll == p or (len(p) >= 4 and p in ll) for p in ni_parties)

def reorder_party_records(collapsed, base_weighted, col_idxs, config):
    """collapsed: list of {'label','raw':{idx:..},'pct':{idx:..}} from collapse_records.
    Returns (records, was_reordered, warnings)."""
    aliases = config['aliases']
    order = config['canonical_order']
    min_matches = config.get('min_matches_to_trigger', 3)
    ni_parties = config.get('ni_parties', [])

    matched = {}
    ni_matches = []
    unmatched = []
    for rec in collapsed:
        canon = _match(rec['label'], aliases)
        if canon and canon not in matched:
            matched[canon] = rec
        elif _is_ni_party(rec['label'], ni_parties):
            ni_matches.append(rec)  # NI party -- doesn't compete for XX, doesn't block reorder
        else:
            unmatched.append(rec)  # either no match, or a duplicate match of a party already seen

    if len(matched) < min_matches:
        return collapsed, False, []  # not a party-choice table; leave as-is

    if len(unmatched) > 1:
        # More than one leftover GB-party-like option than the single XX slot can hold.
        labels = [r['label'] for r in unmatched]
        return collapsed, False, [
            f"left UNREORDERED: {len(unmatched)} unmatched options ({', '.join(labels)}) "
            f"-- too many for the single XX slot without merging distinct parties"
        ]

    xx_rec = unmatched[0] if unmatched else None
    out = []
    for slot in order:
        if slot == 'XX':
            if xx_rec:
                out.append(xx_rec)
            continue
        if slot in matched:
            out.append(matched[slot])
    out.extend(ni_matches)  # tacked on at the end, original order, no fixed position
    warnings = []
    if ni_matches:
        warnings.append(
            f"{len(ni_matches)} Northern Ireland party option(s) appended after the GB order, "
            f"unordered ({', '.join(r['label'] for r in ni_matches)})"
        )
    return out, True, warnings
