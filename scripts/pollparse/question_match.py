import json, re

# strips 'Q:V504.', 'Q1.', 'Q:LG2.', '[asked only to...]', leading/trailing whitespace
_NOISE = re.compile(
    r'^Q[:\s]?[A-Z0-9_]+[\.:]\s*'   # Q:V504. or Q1. at the start
    r'|^\s*\[.*?\]\s*',              # [asked only to those...] footnotes
    re.IGNORECASE | re.DOTALL
)

def _clean(text):
    t = (text or '').strip()
    t = _NOISE.sub('', t).strip()
    # collapse internal whitespace/newlines so multi-line question text matches cleanly
    t = re.sub(r'\s+', ' ', t)
    return t.lower()

def load_question_labels(path):
    with open(path) as f:
        return json.load(f)['labels']

def match_question(raw_text, labels):
    """Returns (canonical_label, matched_pattern) or ('UNMATCHED', '') if nothing fires."""
    cleaned = _clean(raw_text)
    for entry in labels:
        for pat in entry['patterns']:
            if re.search(pat, cleaned):
                return entry['canonical'], pat
    return 'UNMATCHED', ''
