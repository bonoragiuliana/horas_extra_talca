import re
import unicodedata

def normalize_text(s) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    s = re.sub(r"[^a-z0-9\s]+", " ", s)
    return " ".join(s.split())

def name_keys(raw: str):
    full = normalize_text(raw)
    toks = full.split()
    first2 = " ".join(toks[:2]) if len(toks) >= 2 else full
    last2 = " ".join(toks[-2:]) if len(toks) >= 2 else full
    return full, first2, last2