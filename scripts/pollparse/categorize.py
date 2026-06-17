import json

def load_aliases(path):
    with open(path) as f:
        return json.load(f)

def categorize_columns(columns, aliases):
    """Mutates each column dict in place, adding a 'category' key.
    Unmatched columns get category=None and should be logged for review."""
    unmatched = []
    for col in columns:
        group = (col.get('group') or '').strip().lower()
        subgroup = (col.get('subgroup') or '').strip().lower()
        search_text = group if group else subgroup
        matched = None
        for category, patterns in aliases.items():
            for pat in patterns:
                if pat in search_text:
                    matched = category
                    break
            if matched:
                break
        col['category'] = matched
        if matched is None:
            unmatched.append((col['group'], col['subgroup']))
    return unmatched
