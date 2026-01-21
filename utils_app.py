import math
import re
import unicodedata
from datetime import datetime, timedelta, date

import pandas as pd


# ============================================================
# HELPERS
# ============================================================
def normalize_text(s) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = "".join(
        c for c in unicodedata.normalize("NFKD", s)
        if not unicodedata.combining(c)
    )
    # saca comas, puntos, guiones, etc. (deja letras/numeros/espacios)
    s = re.sub(r"[^a-z0-9\s]+", " ", s)
    s = " ".join(s.split())
    return s


def name_keys(raw: str):
    """
    Devuelve 3 claves para matchear nombres aunque vengan con extras:
    - full: nombre normalizado completo
    - first2: primeras 2 palabras
    - last2: últimas 2 palabras
    """
    full = normalize_text(raw)
    toks = full.split()
    first2 = " ".join(toks[:2]) if len(toks) >= 2 else full
    last2 = " ".join(toks[-2:]) if len(toks) >= 2 else full
    return full, first2, last2


def clean_id(val) -> str:
    if pd.isna(val):
        return ""
    if isinstance(val, int):
        return str(val)
    if isinstance(val, float) and float(val).is_integer():
        return str(int(val))
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def only_digits(s: str) -> str:
    return re.sub(r"\D", "", str(s or ""))


def id_key_from_any(val) -> str:
    """
    Clave interna estable para matchear:
    - deja solo dígitos
    - quita ceros a la izquierda (00123 == 123)
    """
    digits = only_digits(clean_id(val))
    if not digits:
        return ""
    k = digits.lstrip("0")
    return k if k else "0"


def extract_id_parts(val):
    """
    Devuelve:
      - key: clave estable (sin ceros a la izquierda)
      - digits: dígitos completos
      - cuil11: si parece CUIL (11)
      - dni8: si se puede inferir DNI (8) (desde DNI o desde CUIL)
    """
    digits = only_digits(clean_id(val))
    key = id_key_from_any(val)

    cuil11 = digits if len(digits) == 11 else ""

    dni8 = ""
    if len(digits) == 11:
        dni8 = digits[2:10].zfill(8)
    elif len(digits) == 8:
        dni8 = digits
    elif len(digits) == 7:
        dni8 = digits.zfill(8)
    elif len(digits) > 8:
        dni8 = digits[-8:]

    return key, digits, cuil11, dni8


def parse_date(val):
    if pd.isna(val):
        return None

    # fecha excel serial
    if isinstance(val, (int, float)) and val > 30000:
        base = datetime(1899, 12, 30)
        return (base + timedelta(days=float(val))).date()

    if isinstance(val, (pd.Timestamp, datetime)):
        return val.date()

    s = str(val).strip()
    m = re.search(r'(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})', s)
    if m:
        s = m.group(1)

    s = s.replace(".", "/").replace("-", "/")
    dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.isna(dt):
        return None
    return dt.date()


def parse_time_only(val):
    if pd.isna(val):
        return None

    if isinstance(val, (pd.Timestamp, datetime)):
        return val.time()

    # hora excel como fracción de día
    if isinstance(val, (int, float)) and 0 <= float(val) < 1:
        total_seconds = int(round(float(val) * 24 * 3600))
        hh = total_seconds // 3600
        mm = (total_seconds % 3600) // 60
        ss = total_seconds % 60
        return datetime(2000, 1, 1, hh, mm, ss).time()

    s = str(val).strip()
    if not s:
        return None

    dt = pd.to_datetime(s, errors="coerce")
    if pd.isna(dt):
        return None
    return dt.time()


def round_dt_to_nearest_hour(dt: datetime) -> datetime:
    """
    Redondea al entero de hora más cercano:
    - 20:58 -> 21:00
    - 06:01 -> 06:00
    Regla: suma 30 min y luego trunca a la hora.
    """
    return (dt + timedelta(minutes=30)).replace(minute=0, second=0, microsecond=0)


def guess_col(columns_norm, keywords):
    for i, c in enumerate(columns_norm):
        for k in keywords:
            if k in c:
                return i
    return None


def excel_cell_to_date(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        dt = pd.to_datetime(str(v), errors="coerce", dayfirst=True)
        if pd.isna(dt):
            return None
        return dt.date()
    except Exception:
        return None


def round_hours(x: float, step: float = 0.25) -> float:
    """
    Redondeo de horas a "pasos" (por defecto 15 min = 0.25h).
    """
    try:
        x = float(x)
    except Exception:
        return 0.0
    if x <= 0:
        return 0.0
    return round(math.floor(x / step + 0.5) * step, 2)


def num_or_zero(v):
    try:
        if v is None or v == "":
            return 0.0
        return float(v)
    except Exception:
        return 0.0
