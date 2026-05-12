import re
import pandas as pd

def clean_id(val) -> str:
    if pd.isna(val): return ""
    if isinstance(val, int): return str(val)
    if isinstance(val, float) and float(val).is_integer(): return str(int(val))
    s = str(val).strip()
    if s.endswith(".0"): s = s[:-2]
    return s

def only_digits(s: str) -> str:
    return re.sub(r"\D", "", str(s or ""))

def id_key_from_any(val) -> str:
    digits = only_digits(clean_id(val))
    if not digits: return ""
    k = digits.lstrip("0")
    return k if k else "0"

def extract_id_parts(val):
    digits = only_digits(clean_id(val))
    key = id_key_from_any(val)
    cuil11 = digits if len(digits) == 11 else ""
    dni8 = ""
    if len(digits) == 11: dni8 = digits[2:10].zfill(8)
    elif len(digits) == 8: dni8 = digits
    elif len(digits) == 7: dni8 = digits.zfill(8)
    elif len(digits) > 8: dni8 = digits[-8:]
    return key, digits, cuil11, dni8

def guess_col(columns_norm, keywords):
    for i, c in enumerate(columns_norm):
        for k in keywords:
            if k in c: return i
    return None