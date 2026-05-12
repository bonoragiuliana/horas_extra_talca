import math
import re
from datetime import datetime, timedelta, date
import pandas as pd

def parse_date(val):
    if pd.isna(val): return None
    if isinstance(val, (int, float)) and val > 30000:
        return (datetime(1899, 12, 30) + timedelta(days=float(val))).date()
    if isinstance(val, (pd.Timestamp, datetime)): return val.date()
    s = str(val).strip()
    m = re.search(r'(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})', s)
    if m: s = m.group(1)
    s = s.replace(".", "/").replace("-", "/")
    dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return dt.date() if not pd.isna(dt) else None

def parse_time_only(val):
    if pd.isna(val): return None
    if isinstance(val, (pd.Timestamp, datetime)): return val.time()
    if isinstance(val, (int, float)) and 0 <= float(val) < 1:
        total_seconds = int(round(float(val) * 24 * 3600))
        return datetime(2000, 1, 1, total_seconds // 3600, (total_seconds % 3600) // 60, total_seconds % 60).time()
    s = str(val).strip()
    if not s: return None
    dt = pd.to_datetime(s, errors="coerce")
    return dt.time() if not pd.isna(dt) else None

def round_dt_to_nearest_hour(dt: datetime) -> datetime:
    return (dt + timedelta(minutes=30)).replace(minute=0, second=0, microsecond=0)

def excel_cell_to_date(v):
    if v is None: return None
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date): return v
    try:
        dt = pd.to_datetime(str(v), errors="coerce", dayfirst=True)
        return dt.date() if not pd.isna(dt) else None
    except: return None

def round_hours(x: float, step: float = 0.25) -> float:
    try: x = float(x)
    except: return 0.0
    if x <= 0: return 0.0
    return round(math.floor(x / step + 0.5) * step, 2)

def num_or_zero(v):
    try: return float(v) if v not in (None, "") else 0.0
    except: return 0.0