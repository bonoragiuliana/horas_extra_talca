import os
import json
import unicodedata
import re
import traceback
import math
import calendar
from copy import copy
from datetime import datetime, timedelta, date

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

# =========================
# OPTIONAL: feriados auto
# =========================
try:
    import holidays as pyholidays
except Exception:
    pyholidays = None


# ============================================================
# CONFIG
# ============================================================
DEFAULT_CONFIG = {
    "company_name": "TALCA",
    "default_jornada_weekday": 8,
    "employees_excel_path": "",
    "template_excel_path": "",
    # feriados (se guardan como YYYY-MM-DD)
    "holidays": [],
    # estándar por día
    "saturday_standard_hours": 8,
    "sunday_standard_hours": 0,
    "holiday_standard_hours": 0,
    # auto feriados
    "auto_holidays_enabled": True,
    "auto_holidays_country": "AR",
    "auto_holidays_subdiv": "M",  # Mendoza
    "auto_holidays_observed": True
}

CONFIG_FILE = "config_horas_extra.json"
EMP_MASTER_DEFAULT_NAME = "datos empleados.xlsx"
TEMPLATE_DEFAULT_NAME = "formatosugerido.xlsx"

DOW_MAP = {0: "Mon", 1: "Tue", 2: "Wed", 3: "Thu", 4: "Fri", 5: "Sat", 6: "Sun"}


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
    s = " ".join(s.split())
    return s


def load_config() -> dict:
    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    # merge defaults
    for k, v in DEFAULT_CONFIG.items():
        if k not in cfg:
            cfg[k] = v
    return cfg


def save_config(cfg: dict) -> None:
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


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


def split_cuil_to_dni(id_value: str):
    digits = only_digits(id_value)
    if len(digits) == 11:
        return digits[2:10], digits
    if len(digits) >= 8:
        return digits[-8:], digits if len(digits) == 11 else ""
    return "", ""


def parse_date(val):
    if pd.isna(val):
        return None

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


def monday_of_week(d: date) -> date:
    return d - timedelta(days=d.weekday())


def week_key_sun_to_sat(d: date) -> date:
    # domingo pertenece a la semana que arranca el lunes siguiente
    if d.weekday() == 6:
        d = d + timedelta(days=1)
    return monday_of_week(d)


def hours_between(start_dt: datetime, end_dt: datetime) -> float:
    if end_dt < start_dt:
        end_dt = end_dt + timedelta(days=1)
    return (end_dt - start_dt).total_seconds() / 3600.0


def is_holiday(d: date, cfg: dict) -> bool:
    hol = set(cfg.get("holidays", []))
    return d.strftime("%Y-%m-%d") in hol


def guess_col(columns_norm, keywords):
    for i, c in enumerate(columns_norm):
        for k in keywords:
            if k in c:
                return i
    return None


# ============================================================
# AUTO HOLIDAYS (Argentina - Mendoza)
# ============================================================
def compute_auto_holidays_for_range(start: date, end: date, cfg: dict):
    """
    Devuelve feriados (nacionales + Mendoza) dentro del rango [start, end].
    Requiere paquete `holidays`.
    """
    if pyholidays is None:
        return []

    if start is None or end is None:
        return []

    years = list(range(start.year, end.year + 1))
    country = cfg.get("auto_holidays_country", "AR")
    subdiv = cfg.get("auto_holidays_subdiv", "M")
    observed = bool(cfg.get("auto_holidays_observed", True))

    try:
        hcal = pyholidays.country_holidays(country, subdiv=subdiv, years=years, observed=observed)
    except Exception:
        return []

    out = []
    cur = start
    while cur <= end:
        if cur in hcal:
            out.append(cur)
        cur += timedelta(days=1)

    return out


# ============================================================
# MASTER EMPLEADOS
# ============================================================
def load_employee_master(emp_path: str) -> pd.DataFrame:
    df = pd.read_excel(emp_path)
    df.columns = [str(c).strip() for c in df.columns]
    colmap = {normalize_text(c): c for c in df.columns}

    def col_like(keys):
        for k in keys:
            for norm, orig in colmap.items():
                if k in norm:
                    return orig
        return None

    c_empresa = col_like(["empresa"])
    c_sector = col_like(["sector"])
    c_id = col_like(["id"])
    c_name = col_like(["apellido", "nombre"])
    c_jornada = col_like(["jornada"])
    c_lav = col_like(["lav"])
    c_sab = col_like(["sab"])
    c_domfer = col_like(["domyfer", "dom", "fer"])

    need = [c_empresa, c_sector, c_id, c_name, c_jornada, c_lav, c_sab, c_domfer]
    if any(x is None for x in need):
        raise RuntimeError(
            "En datos empleados.xlsx faltan columnas esperadas.\n"
            "Necesito: EMPRESA, SECTOR, ID (CUIL/DNI), APELLIDO Y NOMBRE, JORNADA, $/hs LaV, $/hs Sab, $/hs DomYFer."
        )

    out = pd.DataFrame()
    out["raw_id"] = df[c_id].apply(clean_id)
    out["dni8"] = out["raw_id"].apply(lambda x: split_cuil_to_dni(x)[0])

    out["empresa_master"] = df[c_empresa].fillna("").astype(str).str.strip()
    out["sector_master"] = df[c_sector].fillna("").astype(str).str.strip().apply(normalize_text)
    out["nombre_master"] = df[c_name].fillna("").astype(str).str.strip()
    out["nombre_norm_rep"] = out["nombre_master"].apply(normalize_text)

    out["jornada_weekday"] = pd.to_numeric(df[c_jornada], errors="coerce").fillna(0).astype(float)
    out["rate_lav"] = pd.to_numeric(df[c_lav], errors="coerce").fillna(0).astype(float)
    out["rate_sab"] = pd.to_numeric(df[c_sab], errors="coerce").fillna(0).astype(float)
    out["rate_domfer"] = pd.to_numeric(df[c_domfer], errors="coerce").fillna(0).astype(float)

    out = out[out["dni8"] != ""].drop_duplicates(subset=["dni8"], keep="first")
    return out


# ============================================================
# LECTURA REPORTE VEOTIME
# ============================================================
def read_report_raw(path: str) -> pd.DataFrame:
    with open(path, "rb") as f:
        head = f.read(8192).lstrip()
    head_low = head.lower()

    # xlsx
    if head.startswith(b"PK"):
        return pd.read_excel(path, engine="openpyxl", header=None)

    # xls real (OLE)
    if head.startswith(b"\xD0\xCF\x11\xE0"):
        return pd.read_excel(path, engine="xlrd", header=None)

    # html disfrazado
    if b"<html" in head_low or b"<!doctype" in head_low or b"<table" in head_low:
        tables = pd.read_html(path)

        def score_table(t: pd.DataFrame) -> int:
            txt = " ".join([str(x) for x in list(t.columns)] + t.astype(str).values.ravel().tolist())
            txt = normalize_text(txt)
            score = 0
            for kw in ["fecha", "hora", "dni", "marc", "entrada", "salida"]:
                if kw in txt:
                    score += 1
            return score

        scored = [(score_table(t), t) for t in tables]
        scored.sort(key=lambda x: x[0], reverse=True)
        best = scored[0][1].copy()
        best.columns = list(range(best.shape[1]))
        return best.reset_index(drop=True)

    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        return pd.read_excel(path, engine="openpyxl", header=None)
    if ext == ".xls":
        return pd.read_excel(path, engine="xlrd", header=None)

    raise RuntimeError("Formato no soportado. Usá .xls o .xlsx.")


def find_header_row(df_raw, max_rows=80):
    for r in range(min(max_rows, len(df_raw))):
        row = df_raw.iloc[r].astype(str).tolist()
        text = " ".join(normalize_text(x) for x in row)

        has_fecha = "fecha" in text
        has_hora = "hora" in text or "time" in text
        has_dni = "dni" in text or "documento" in text or "legajo" in text
        has_marc = "marc" in text or "tipo" in text

        if has_fecha and has_hora and has_dni and has_marc:
            return r
    return None


def apply_header_row(df_raw, header_row):
    headers = df_raw.iloc[header_row].tolist()
    df = df_raw.iloc[header_row + 1:].copy()
    df.columns = [str(h).strip() for h in headers]
    df = df.dropna(axis=1, how="all")
    return df


def read_veotime_to_daily(path: str, emp_master: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    df_raw = read_report_raw(path)

    header_row = find_header_row(df_raw)
    if header_row is not None:
        df = apply_header_row(df_raw, header_row)
    else:
        df = df_raw.copy()
        df.columns = [str(c).strip() for c in df.iloc[0].tolist()]
        df = df.iloc[1:].copy()

    # debug
    df.head(200).to_csv("debug_veotime_head.csv", index=False, encoding="utf-8-sig")
    with open("debug_columnas.txt", "w", encoding="utf-8") as f:
        f.write("Columnas detectadas:\n")
        for c in df.columns:
            f.write(f"- {c}\n")

    cols_norm = [normalize_text(c) for c in df.columns]

    i_fecha = guess_col(cols_norm, ["fecha"])
    i_hora = guess_col(cols_norm, ["hora", "time"])
    i_marc = guess_col(cols_norm, ["marcaci", "marc", "tipo"])
    i_dni = guess_col(cols_norm, ["dni", "documento", "legajo"])
    i_nombre = guess_col(cols_norm, ["nombre", "empleado", "apellido", "colaborador"])

    missing = []
    if i_fecha is None: missing.append("Fecha")
    if i_hora is None: missing.append("Hora")
    if i_marc is None: missing.append("Marcación")
    if i_dni is None: missing.append("DNI")
    if i_nombre is None: missing.append("Nombre")

    if missing:
        raise RuntimeError(
            "No pude detectar columnas clave en el reporte VeoTime.\n"
            f"Faltan: {', '.join(missing)}\n\n"
            "Abrí debug_veotime_head.csv para ver encabezados."
        )

    events = pd.DataFrame()
    events["fecha"] = df.iloc[:, i_fecha].apply(parse_date)
    events["hora"] = df.iloc[:, i_hora].apply(parse_time_only)
    events["raw_id"] = df.iloc[:, i_dni].apply(clean_id)
    events["dni8"] = events["raw_id"].apply(lambda x: split_cuil_to_dni(x)[0])

    events["nombre_rep"] = df.iloc[:, i_nombre].astype(str).str.strip()
    events["nombre_norm_rep"] = events["nombre_rep"].apply(normalize_text)

    def norm_tipo(x):
        t = normalize_text(x)
        if "entrada" in t or "ingreso" in t or t.startswith("ent"):
            return "entrada"
        if "salida" in t or "egreso" in t or t.startswith("sal"):
            return "salida"
        return ""

    events["tipo"] = df.iloc[:, i_marc].apply(norm_tipo)

    events = events.dropna(subset=["fecha", "hora"])
    events = events[events["tipo"].isin(["entrada", "salida"])]
    events = events[events["dni8"] != ""]

    if events.empty:
        raise RuntimeError(
            "No quedaron eventos válidos (Entrada/Salida).\n"
            "Revisá debug_veotime_head.csv, especialmente la columna Marcación."
        )

    # =======================
    # FIX: evitar choque nombre_norm_rep
    # =======================
    merged = events.merge(emp_master, how="left", on="dni8", suffixes=("_rep", "_master"))

    # fallback por nombre para no matcheados
    no_match = merged["nombre_master"].isna()
    if no_match.any():
        emp_by_name = emp_master.drop_duplicates(subset=["nombre_norm_rep"], keep="first").copy()

        merged_no = merged.loc[no_match, ["nombre_norm_rep_rep"]].copy()
        merged_no = merged_no.merge(
            emp_by_name[["nombre_norm_rep", "empresa_master", "sector_master", "nombre_master",
                        "jornada_weekday", "rate_lav", "rate_sab", "rate_domfer"]],
            how="left",
            left_on="nombre_norm_rep_rep",
            right_on="nombre_norm_rep"
        )

        for col in ["empresa_master", "sector_master", "nombre_master",
                    "jornada_weekday", "rate_lav", "rate_sab", "rate_domfer"]:
            merged.loc[no_match, col] = merged_no[col].values

        merged_no.to_csv("debug_matcheados_por_nombre.csv", index=False, encoding="utf-8-sig")

    still_no = merged["nombre_master"].isna()
    if still_no.any():
        merged.loc[still_no, ["dni8", "nombre_rep"]].drop_duplicates().to_csv(
            "debug_no_matcheados.csv", index=False, encoding="utf-8-sig"
        )

    # empresa / sector / nombre finales
    merged["empresa_final"] = merged["empresa_master"].fillna("").astype(str).str.strip()
    merged["empresa_final"] = merged["empresa_final"].where(
        merged["empresa_final"] != "", cfg.get("company_name", "TALCA")
    )

    merged["sector_final"] = merged["sector_master"].fillna("").astype(str).str.strip()
    merged["sector_final"] = merged["sector_final"].where(merged["sector_final"] != "", "desconocido")
    merged["sector_final"] = merged["sector_final"].apply(normalize_text)

    merged["nombre_final"] = merged["nombre_master"].fillna("").astype(str).str.strip()
    merged["nombre_final"] = merged["nombre_final"].where(
        merged["nombre_final"] != "", merged["nombre_rep"].fillna("").astype(str)
    )

    merged["jornada_weekday"] = merged["jornada_weekday"].fillna(float(cfg.get("default_jornada_weekday", 8)))
    merged.loc[merged["jornada_weekday"] <= 0, "jornada_weekday"] = float(cfg.get("default_jornada_weekday", 8))

    merged["rate_lav"] = merged["rate_lav"].fillna(0.0)
    merged["rate_sab"] = merged["rate_sab"].fillna(0.0)
    merged["rate_domfer"] = merged["rate_domfer"].fillna(0.0)

    # construir diario (pair entrada->salida)
    daily_rows = []
    for (dni8, empresa, nombre, sector, fecha, jornada, rlav, rsab, rdom), sub in merged.groupby(
        ["dni8", "empresa_final", "nombre_final", "sector_final", "fecha",
         "jornada_weekday", "rate_lav", "rate_sab", "rate_domfer"],
        dropna=False
    ):
        sub = sub.sort_values("hora")
        last_entry = None
        total_hs = 0.0

        for _, r in sub.iterrows():
            if r["tipo"] == "entrada":
                last_entry = r["hora"]
            else:
                if last_entry is not None:
                    start_dt = datetime.combine(fecha, last_entry)
                    end_dt = datetime.combine(fecha, r["hora"])
                    total_hs += hours_between(start_dt, end_dt)
                    last_entry = None

        daily_rows.append({
            "dni": str(dni8),
            "empresa": str(empresa),
            "nombre": str(nombre),
            "sector": str(sector),
            "fecha": fecha,
            "horas_trab": float(total_hs),
            "jornada_weekday": float(jornada),
            "rate_lav": float(rlav),
            "rate_sab": float(rsab),
            "rate_domfer": float(rdom)
        })

    daily = pd.DataFrame(daily_rows)
    if daily.empty:
        raise RuntimeError("No pude construir horas diarias. Revisá debug_veotime_head.csv.")
    return daily


# ============================================================
# OVERTIME (HORAS REDONDAS)
# ============================================================
def compute_overtime_from_daily(daily: pd.DataFrame, cfg: dict) -> dict:
    d = daily.copy()
    d["semana_lunes"] = d["fecha"].apply(week_key_sun_to_sat)
    d["dow"] = d["fecha"].apply(lambda x: x.weekday())
    d["dow_key"] = d["dow"].map(DOW_MAP)

    def std_hours(row):
        # FERIADO: se calcula igual que un día normal (toma la jornada del empleado)
        if is_holiday(row["fecha"], cfg):
            hs = float(cfg.get("holiday_standard_hours", 0))
            # compatibilidad: si en tu config quedó 0, usamos la jornada (ej: 8)
            if hs <= 0:
                hs = float(row["jornada_weekday"])
            return hs

        if row["dow_key"] == "Sun":
            return float(cfg.get("sunday_standard_hours", 0))

        if row["dow_key"] == "Sat":
            return float(cfg.get("saturday_standard_hours", 8))

        return float(row["jornada_weekday"])

    d["std"] = d.apply(std_hours, axis=1)

    # redondeo hacia abajo: 8:55 a 13:02 => 4hs
    d["horas_trab_red"] = d["horas_trab"].apply(lambda x: math.floor(float(x)))
    d["std_red"] = d["std"].apply(lambda x: math.floor(float(x)))

    d["horas_extra"] = (d["horas_trab_red"] - d["std_red"]).clip(lower=0)

    def rate_for(row):
        if is_holiday(row["fecha"], cfg) or row["dow_key"] == "Sun":
            return float(row["rate_domfer"])
        if row["dow_key"] == "Sat":
            return float(row["rate_sab"])
        return float(row["rate_lav"])

    d["tarifa_he"] = d.apply(rate_for, axis=1)
    d["costo"] = d["horas_extra"] * d["tarifa_he"]

    weeks = {}
    for week_start, sub in d.groupby("semana_lunes"):
        weeks[week_start] = sub.copy()
    return weeks


# ============================================================
# TEMPLATE HELPERS
# ============================================================
def _safe_sheet_title(s: str) -> str:
    s = re.sub(r'[\\/*?:\[\]]', '-', str(s))
    return s[:31].strip() or "Semana"


def _col_for_date(d: date, cfg: dict) -> int:
    # En tu plantilla: Dom=5, Lun=6, ... Sab=11, Feriado=12
    if is_holiday(d, cfg):
        return 12
    if d.weekday() == 6:
        return 5
    return 6 + d.weekday()


def _clear_data_area(ws, start_row=7, max_col=20):
    for r in range(start_row, ws.max_row + 1):
        for c in range(1, 13):
            ws.cell(r, c).value = None
        ws.cell(r, 19).value = None
        ws.cell(r, 20).value = None


def _ensure_rows(ws, start_row, n_rows_needed, base_style_row=7, max_col=20):
    last_needed = start_row + n_rows_needed - 1
    if last_needed <= ws.max_row:
        return

    for r in range(ws.max_row + 1, last_needed + 1):
        ws.insert_rows(r)
        for c in range(1, max_col + 1):
            src = ws.cell(base_style_row, c)
            dst = ws.cell(r, c)
            dst._style = copy(src._style)
            dst.font = copy(src.font)
            dst.border = copy(src.border)
            dst.fill = copy(src.fill)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)
            dst.alignment = copy(src.alignment)


def _set_week_header(ws, week_start: date, cfg: dict):
    sun = week_start - timedelta(days=1)
    dates = [
        (5, sun),
        (6, week_start),
        (7, week_start + timedelta(days=1)),
        (8, week_start + timedelta(days=2)),
        (9, week_start + timedelta(days=3)),
        (10, week_start + timedelta(days=4)),
        (11, week_start + timedelta(days=5)),
    ]
    for col, d in dates:
        cell = ws.cell(4, col)
        cell.value = d
        cell.number_format = "dd/mm/yy"

    # Si hay 1+ feriados en la semana, mostramos el primero en el header
    hol = None
    for i in range(-1, 6):
        d = week_start + timedelta(days=i)
        if is_holiday(d, cfg):
            hol = d
            break
    ws.cell(4, 12).value = hol if hol else None
    ws.cell(4, 12).number_format = "dd/mm/yy"


# ============================================================
# OUTPUT: 1 HOJA POR SEMANA (TODOS LOS EMPLEADOS)
# ============================================================
def build_output_workbook_template(weeks: dict, out_path: str, template_path: str, cfg: dict):
    wb = load_workbook(template_path)
    tpl = wb.active
    created = []

    for week_start, df_week in sorted(weeks.items(), key=lambda x: x[0]):
        ws = wb.copy_worksheet(tpl)
        created.append(ws)

        ws.title = _safe_sheet_title(f"Semana {week_start.strftime('%d-%m')}")
        _set_week_header(ws, week_start, cfg)

        ws["M6"].value = ""
        ws["N6"].value = ""
        ws["O6"].value = ""

        _clear_data_area(ws, start_row=7, max_col=20)

        sub = df_week.copy()

        day_he = (
            sub.groupby(["empresa", "dni", "nombre", "sector", "fecha", "rate_lav", "rate_sab", "rate_domfer"], dropna=False)["horas_extra"]
            .sum()
            .reset_index()
        )

        employees = (
            day_he[["empresa", "dni", "nombre", "sector", "rate_lav", "rate_sab", "rate_domfer"]]
            .drop_duplicates()
            .sort_values(["empresa", "sector", "nombre"])
        )

        _ensure_rows(ws, start_row=7, n_rows_needed=len(employees), base_style_row=7, max_col=20)

        he_map = {}
        for _, r in day_he.iterrows():
            key = (str(r["empresa"]), str(r["dni"]), str(r["nombre"]), str(r["sector"]))
            he_map.setdefault(key, {})[r["fecha"]] = float(r["horas_extra"])

        for i, (_, emp) in enumerate(employees.iterrows()):
            row = 7 + i
            empresa = str(emp["empresa"])
            dni = str(emp["dni"])
            nombre = str(emp["nombre"])
            sector = str(emp["sector"])

            rate_lav = float(emp["rate_lav"])
            rate_sab = float(emp["rate_sab"])
            rate_domfer = float(emp["rate_domfer"])

            ws.cell(row, 1).value = empresa
            ws.cell(row, 2).value = sector
            ws.cell(row, 3).value = dni
            ws.cell(row, 4).value = nombre

            # inicializar E..L en 0
            for c in range(5, 13):
                ws.cell(row, c).value = 0
                ws.cell(row, c).number_format = "0"
                ws.cell(row, c).alignment = Alignment(horizontal="center", vertical="center")

            key = (empresa, dni, nombre, sector)
            emp_he_by_date = he_map.get(key, {})

            for dte, he in emp_he_by_date.items():
                col = _col_for_date(dte, cfg)
                ws.cell(row, col).value = int(he)
                ws.cell(row, col).number_format = "0"

            # costos por empleado
            cost_lav = 0.0
            cost_sab = 0.0
            cost_domfer = 0.0

            for dte, he in emp_he_by_date.items():
                he = float(he)
                if he <= 0:
                    continue
                if is_holiday(dte, cfg) or dte.weekday() == 6:
                    cost_domfer += he * rate_domfer
                elif dte.weekday() == 5:
                    cost_sab += he * rate_sab
                else:
                    cost_lav += he * rate_lav

            subtotal = cost_lav + cost_sab + cost_domfer

            ws.cell(row, 13).value = round(cost_lav, 0)
            ws.cell(row, 14).value = round(cost_sab, 0)
            ws.cell(row, 15).value = round(cost_domfer, 0)
            ws.cell(row, 16).value = round(subtotal, 0)
            ws.cell(row, 17).value = 0
            ws.cell(row, 18).value = round(subtotal, 0)

            for c in [13, 14, 15, 16, 17, 18]:
                ws.cell(row, c).number_format = "#,##0"
                ws.cell(row, c).alignment = Alignment(horizontal="center", vertical="center")

            ws.cell(row, 19).value = None
            ws.cell(row, 20).value = None

    wb.remove(tpl)

    if not created:
        raise RuntimeError("No se generaron hojas. Revisá si el reporte trae datos válidos.")

    wb.save(out_path)


# ============================================================
# UI PREMIUM + CALENDARIO INLINE
# ============================================================
def run_app():
    cfg = load_config()

    try:
        import ttkbootstrap as tb
    except Exception:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Falta ttkbootstrap", "Instalá: pip install ttkbootstrap")
        return

    THEME = "minty"
    root = tb.Window(themename=THEME)
    root.title("Liquidación de Horas Extra · TALCA")
    root.geometry("1120x700")
    root.minsize(1040, 660)

    FONT_TITLE = ("Segoe UI", 20, "bold")
    FONT_SUB = ("Segoe UI", 10)
    FONT_H2 = ("Segoe UI", 12, "bold")
    FONT_B = ("Segoe UI", 10, "bold")

    var_report = tk.StringVar(master=root, value="")
    var_emp = tk.StringVar(master=root, value=cfg.get("employees_excel_path", ""))
    var_tpl = tk.StringVar(master=root, value=cfg.get("template_excel_path", ""))

    var_use_holidays = tk.BooleanVar(master=root, value=False)
    var_auto_holidays = tk.BooleanVar(master=root, value=bool(cfg.get("auto_holidays_enabled", True)))

    status_var = tk.StringVar(master=root, value="Paso 1: elegí el reporte de VeoTime.")
    summary_range_var = tk.StringVar(master=root, value="—")
    summary_emps_var = tk.StringVar(master=root, value="—")
    summary_holidays_var = tk.StringVar(master=root, value="—")

    selected_holidays_manual = set()
    selected_holidays_auto = set()

    # IMPORTANT: Fix NameError btn_generate (se crea después)
    btn_generate = None

    def set_status(msg: str):
        status_var.set(msg)
        try:
            root.update_idletasks()
        except Exception:
            pass

    def month_name_es(m):
        names = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
        return names[m-1]

    def get_report_date_range_and_emps(report_path: str):
        try:
            df_raw = read_report_raw(report_path)
            header_row = find_header_row(df_raw)
            if header_row is not None:
                df = apply_header_row(df_raw, header_row)
            else:
                df = df_raw.copy()
                df.columns = [str(c).strip() for c in df.iloc[0].tolist()]
                df = df.iloc[1:].copy()

            cols_norm = [normalize_text(c) for c in df.columns]
            i_fecha = guess_col(cols_norm, ["fecha"])
            i_dni = guess_col(cols_norm, ["dni", "documento", "legajo"])

            if i_fecha is None:
                return None, None, None

            fechas = df.iloc[:, i_fecha].apply(parse_date).dropna()
            if fechas.empty:
                return None, None, None

            start = min(fechas)
            end = max(fechas)

            emps = None
            if i_dni is not None:
                tmp = df.iloc[:, i_dni].apply(clean_id)
                dni8 = tmp.apply(lambda x: split_cuil_to_dni(x)[0])
                dni8 = dni8[dni8 != ""]
                if not dni8.empty:
                    emps = int(dni8.nunique())

            return start, end, emps
        except Exception:
            return None, None, None

    def refresh_auto_holidays():
        selected_holidays_auto.clear()

        if not var_use_holidays.get():
            return
        if not var_auto_holidays.get():
            return

        rp = var_report.get().strip()
        if not rp or not os.path.exists(rp):
            return

        start, end, _ = get_report_date_range_and_emps(rp)
        if start is None or end is None:
            return

        try:
            auto_dates = compute_auto_holidays_for_range(start, end, cfg)
            for d in auto_dates:
                selected_holidays_auto.add(d)
        except Exception:
            pass

    # ===== Summary / state =====
    def update_generate_state():
        rp = var_report.get().strip()
        emp = var_emp.get().strip()
        tpl = var_tpl.get().strip()

        ok = bool(
            rp and os.path.exists(rp) and
            emp and os.path.exists(emp) and
            tpl and os.path.exists(tpl)
        )

        # Fix: si todavía no existe el botón, no hacemos nada
        if btn_generate is None:
            return

        btn_generate.configure(state=("normal" if ok else "disabled"))

    def update_summary():
        rp = var_report.get().strip()
        if rp and os.path.exists(rp):
            start, end, emps = get_report_date_range_and_emps(rp)
            summary_range_var.set(f"{start.strftime('%d/%m/%Y')} → {end.strftime('%d/%m/%Y')}" if start and end else "—")
            summary_emps_var.set(str(emps) if emps is not None else "—")
        else:
            summary_range_var.set("—")
            summary_emps_var.set("—")

        all_h = sorted(selected_holidays_auto | selected_holidays_manual) if var_use_holidays.get() else []
        summary_holidays_var.set(", ".join(d.strftime("%d/%m") for d in all_h) if all_h else "—")

        update_generate_state()

    # ====== default files local ======
    local_emp = os.path.join(os.getcwd(), EMP_MASTER_DEFAULT_NAME)
    if not var_emp.get().strip() and os.path.exists(local_emp):
        var_emp.set(local_emp)
        cfg["employees_excel_path"] = local_emp

    local_tpl = os.path.join(os.getcwd(), TEMPLATE_DEFAULT_NAME)
    if not var_tpl.get().strip() and os.path.exists(local_tpl):
        var_tpl.set(local_tpl)
        cfg["template_excel_path"] = local_tpl

    # ==========================================================
    # LAYOUT
    # ==========================================================
    header = tb.Frame(root, padding=(20, 18))
    header.pack(fill="x")

    left_h = tb.Frame(header)
    left_h.pack(side="left", fill="x", expand=True)

    tb.Label(left_h, text="Liquidación de Horas Extra", font=FONT_TITLE).pack(anchor="w")
    tb.Label(left_h, text="Generación automática desde VeoTime → Excel RRHH", font=FONT_SUB, foreground="#666").pack(anchor="w", pady=(4, 0))

    right_h = tb.Frame(header)
    right_h.pack(side="right")
    tb.Label(right_h, text="RRHH", bootstyle="info-inverse", padding=(10, 4)).pack(side="right", padx=(8, 0))
    tb.Label(right_h, text="TALCA", bootstyle="secondary-inverse", padding=(10, 4)).pack(side="right")

    body = tb.Frame(root, padding=(20, 0, 20, 12))
    body.pack(fill="both", expand=True)

    body.columnconfigure(0, weight=3)
    body.columnconfigure(1, weight=2)
    body.rowconfigure(0, weight=1)

    left = tb.Frame(body)
    left.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
    right = tb.Frame(body)
    right.grid(row=0, column=1, sticky="nsew", padx=(14, 0))

    nb = tb.Notebook(left, bootstyle="primary")
    nb.pack(fill="both", expand=True)

    tab_files = tb.Frame(nb, padding=16)
    tab_holidays = tb.Frame(nb, padding=16)
    tab_help = tb.Frame(nb, padding=16)

    nb.add(tab_files, text="  1 · Archivos  ")
    nb.add(tab_holidays, text="  2 · Feriados  ")
    nb.add(tab_help, text="  3 · Checklist  ")

    # Right: summary + status
    summary = tb.Labelframe(right, text="RESUMEN", padding=16, bootstyle="secondary")
    summary.pack(fill="x")

    def summary_row(parent, label, var):
        fr = tb.Frame(parent)
        fr.pack(fill="x", pady=8)
        tb.Label(fr, text=label, foreground="#666").pack(side="left")
        tb.Label(fr, textvariable=var, font=FONT_B).pack(side="right")

    summary_row(summary, "Rango detectado", summary_range_var)
    summary_row(summary, "Empleados (aprox.)", summary_emps_var)
    summary_row(summary, "Feriados (dd/mm)", summary_holidays_var)

    status_card = tb.Labelframe(right, text="ESTADO", padding=16, bootstyle="light")
    status_card.pack(fill="both", expand=True, pady=(14, 0))

    tb.Label(status_card, textvariable=status_var, wraplength=360, justify="left").pack(anchor="w")
    pb = tb.Progressbar(status_card, mode="indeterminate", bootstyle="success-striped")
    pb.pack(fill="x", pady=(14, 0))

    # ==========================================================
    # TAB 1: FILES
    # ==========================================================
    tb.Label(tab_files, text="Seleccioná los 3 archivos", font=FONT_H2).pack(anchor="w")

    def entry_file(parent, title, var, browse_cmd, hint=""):
        box = tb.Frame(parent)
        box.pack(fill="x", pady=10)

        top = tb.Frame(box)
        top.pack(fill="x")
        tb.Label(top, text=title, font=FONT_B).pack(side="left")
        if hint:
            tb.Label(top, text=hint, font=FONT_SUB, foreground="#666").pack(side="left", padx=(10, 0))

        row = tb.Frame(box)
        row.pack(fill="x", pady=(6, 0))
        e = tb.Entry(row, textvariable=var)
        e.pack(side="left", fill="x", expand=True)
        tb.Button(row, text="Buscar…", command=browse_cmd, bootstyle="secondary", width=12).pack(side="left", padx=(10, 0))
        return e

    def pick_report():
        p = filedialog.askopenfilename(
            title="Seleccionar reporte de VeoTime",
            filetypes=[("Excel", "*.xls *.xlsx")]
        )
        if p:
            var_report.set(p)
            set_status("Reporte seleccionado. Paso 2: revisá feriados (si aplica).")
            refresh_auto_holidays()
            update_summary()
            # refrescar calendario si estaba armado
            try:
                render_calendar()
                render_holiday_list()
            except Exception:
                pass

    def pick_emp():
        p = filedialog.askopenfilename(
            title="Seleccionar datos empleados.xlsx",
            filetypes=[("Excel", "*.xlsx *.xls")]
        )
        if p:
            var_emp.set(p)
            cfg["employees_excel_path"] = p
            save_config(cfg)
            update_summary()

    def pick_tpl():
        p = filedialog.askopenfilename(
            title="Seleccionar formatosugerido.xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if p:
            var_tpl.set(p)
            cfg["template_excel_path"] = p
            save_config(cfg)
            update_summary()

    entry_file(tab_files, "Reporte VeoTime (.xls/.xlsx)", var_report, pick_report, hint="Obligatorio")
    entry_file(tab_files, "Datos empleados.xlsx", var_emp, pick_emp, hint="Obligatorio")
    entry_file(tab_files, "Plantilla formatosugerido.xlsx", var_tpl, pick_tpl, hint="Obligatorio")

    tip_box = tb.Labelframe(tab_files, text="Tip", padding=12, bootstyle="info")
    tip_box.pack(fill="x", pady=(10, 0))
    tb.Label(
        tip_box,
        text="Si algo no matchea, se generan archivos debug (debug_veotime_head.csv / debug_no_matcheados.csv).",
        wraplength=760, justify="left"
    ).pack(anchor="w")

    # ==========================================================
    # TAB 2: HOLIDAYS (INLINE CALENDAR)
    # ==========================================================
    tb.Label(tab_holidays, text="Feriados de la semana", font=FONT_H2).pack(anchor="w")

    hol_top = tb.Frame(tab_holidays)
    hol_top.pack(fill="x", pady=(12, 10))

    cal_year = tk.IntVar(master=root, value=datetime.now().year)
    cal_month = tk.IntVar(master=root, value=datetime.now().month)

    def get_report_range():
        rp = var_report.get().strip()
        if rp and os.path.exists(rp):
            start, end, _ = get_report_date_range_and_emps(rp)
            return start, end
        return None, None

    def toggle_manual_date(dte: date):
        if not var_use_holidays.get():
            return
        if dte in selected_holidays_manual:
            selected_holidays_manual.remove(dte)
        else:
            selected_holidays_manual.add(dte)
        render_holiday_list()
        update_summary()
        render_calendar()

    def prev_month():
        y, m = cal_year.get(), cal_month.get()
        if m == 1:
            cal_year.set(y - 1)
            cal_month.set(12)
        else:
            cal_month.set(m - 1)
        render_calendar()

    def next_month():
        y, m = cal_year.get(), cal_month.get()
        if m == 12:
            cal_year.set(y + 1)
            cal_month.set(1)
        else:
            cal_month.set(m + 1)
        render_calendar()

    def goto_today():
        today = date.today()
        cal_year.set(today.year)
        cal_month.set(today.month)
        render_calendar()

    def on_toggle_holidays():
        sync_holiday_state()
        refresh_auto_holidays()
        render_holiday_list()
        update_summary()
        render_calendar()

    def on_toggle_auto():
        cfg["auto_holidays_enabled"] = bool(var_auto_holidays.get())
        save_config(cfg)
        refresh_auto_holidays()
        render_holiday_list()
        update_summary()
        render_calendar()

    chk_use = tb.Checkbutton(
        hol_top,
        text="Hubo feriados",
        variable=var_use_holidays,
        command=on_toggle_holidays,
        bootstyle="round-toggle"
    )
    chk_use.pack(side="left")

    chk_auto = tb.Checkbutton(
        hol_top,
        text="Detectar automáticamente (Argentina · Mendoza)",
        variable=var_auto_holidays,
        command=on_toggle_auto,
        bootstyle="round-toggle"
    )
    chk_auto.pack(side="left", padx=(14, 0))

    wrap = tb.Frame(tab_holidays)
    wrap.pack(fill="both", expand=True)

    wrap.columnconfigure(0, weight=3)
    wrap.columnconfigure(1, weight=2)
    wrap.rowconfigure(0, weight=1)

    cal_card = tb.Labelframe(wrap, text="CALENDARIO", padding=12, bootstyle="light")
    cal_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

    list_card = tb.Labelframe(wrap, text="FERIADOS APLICADOS", padding=12, bootstyle="secondary")
    list_card.grid(row=0, column=1, sticky="nsew", padx=(12, 0))

    cal_header = tb.Frame(cal_card)
    cal_header.pack(fill="x", pady=(0, 10))

    tb.Button(cal_header, text="◀", command=prev_month, bootstyle="secondary", width=4).pack(side="left")
    lbl_month = tb.Label(cal_header, text="", font=("Segoe UI", 11, "bold"))
    lbl_month.pack(side="left", padx=10)
    tb.Button(cal_header, text="▶", command=next_month, bootstyle="secondary", width=4).pack(side="left")
    tb.Button(cal_header, text="Hoy", command=goto_today, bootstyle="info", width=8).pack(side="right")

    legend = tb.Frame(cal_card)
    legend.pack(fill="x", pady=(0, 8))
    tb.Label(legend, text="MANUAL", bootstyle="primary-inverse", padding=(8, 3)).pack(side="left")
    tb.Label(legend, text="AUTO", bootstyle="info-inverse", padding=(8, 3)).pack(side="left", padx=(8, 0))
    tb.Label(legend, text="Click para agregar/quitar", foreground="#666").pack(side="left", padx=(10, 0))

    range_lbl = tb.Label(cal_card, text="", foreground="#666")
    range_lbl.pack(anchor="w", pady=(0, 8))

    grid = tb.Frame(cal_card)
    grid.pack(fill="both", expand=True)

    days_row = tb.Frame(grid)
    days_row.pack(fill="x")
    for dn in ["Lu", "Ma", "Mi", "Ju", "Vi", "Sa", "Do"]:
        tb.Label(days_row, text=dn, width=5, anchor="center", foreground="#666").pack(side="left", padx=2)

    cells = tb.Frame(grid)
    cells.pack(fill="both", expand=True, pady=(6, 0))

    def render_calendar():
        for w in cells.winfo_children():
            w.destroy()

        y, m = cal_year.get(), cal_month.get()
        lbl_month.config(text=f"{month_name_es(m)} {y}")

        start, end = get_report_range()
        if start and end:
            range_lbl.config(text=f"Rango del reporte: {start.strftime('%d/%m/%Y')} → {end.strftime('%d/%m/%Y')}")
        else:
            range_lbl.config(text="(Tip: al elegir el reporte, se limita el rango del calendario)")

        enabled = var_use_holidays.get()
        weeks = calendar.monthcalendar(y, m)

        for week in weeks:
            row = tb.Frame(cells)
            row.pack(fill="x", pady=2)

            for day in week:
                if day == 0:
                    tb.Label(row, text="", width=5).pack(side="left", padx=2)
                    continue

                dte = date(y, m, day)

                in_range = True
                if start and end:
                    in_range = (start <= dte <= end)

                is_manual = dte in selected_holidays_manual
                is_auto = dte in selected_holidays_auto

                if is_manual:
                    style = "primary"
                elif is_auto:
                    style = "info-outline"
                else:
                    style = "light"

                state = "normal"
                if (not enabled) or (not in_range):
                    state = "disabled"

                b = tb.Button(
                    row,
                    text=str(day),
                    width=5,
                    command=lambda dd=dte: toggle_manual_date(dd),
                    bootstyle=style
                )
                b.pack(side="left", padx=2)
                b.configure(state=state)

    def render_holiday_list():
        lb.delete(0, "end")
        all_dates = sorted(selected_holidays_auto | selected_holidays_manual)
        if not all_dates:
            lb.insert("end", "—")
            return
        for dte in all_dates:
            tags = []
            if dte in selected_holidays_auto:
                tags.append("AUTO")
            if dte in selected_holidays_manual:
                tags.append("MANUAL")
            suffix = f"   ·   {' + '.join(tags)}" if tags else ""
            lb.insert("end", dte.strftime("%d/%m/%Y") + suffix)

    def remove_selected_holiday():
        sel = lb.curselection()
        if not sel:
            return
        all_dates = sorted(selected_holidays_auto | selected_holidays_manual)
        if not all_dates:
            return
        dte = all_dates[sel[0]]

        if dte in selected_holidays_manual:
            selected_holidays_manual.remove(dte)
        elif dte in selected_holidays_auto:
            selected_holidays_auto.remove(dte)

        render_holiday_list()
        update_summary()
        render_calendar()

    def clear_all_holidays():
        selected_holidays_manual.clear()
        selected_holidays_auto.clear()
        render_holiday_list()
        update_summary()
        render_calendar()

    tb.Label(list_card, text="Seleccionados (auto + manual):", font=FONT_B).pack(anchor="w")
    lb = tk.Listbox(list_card, height=12)
    lb.pack(fill="both", expand=True, pady=(10, 10))

    actions = tb.Frame(list_card)
    actions.pack(fill="x")

    btn_remove_h = tb.Button(actions, text="Quitar seleccionado", command=remove_selected_holiday, bootstyle="secondary")
    btn_clear_h = tb.Button(actions, text="Limpiar todo", command=clear_all_holidays, bootstyle="secondary")
    btn_remove_h.pack(side="left")
    btn_clear_h.pack(side="right")

    def sync_holiday_state():
        enabled = var_use_holidays.get()
        st = "normal" if enabled else "disabled"

        try:
            chk_auto.configure(state=st)
        except Exception:
            pass

        try:
            btn_remove_h.configure(state=st)
            btn_clear_h.configure(state=st)
        except Exception:
            pass

        try:
            lb.configure(state=st)
        except Exception:
            pass

        if not enabled:
            selected_holidays_manual.clear()
            selected_holidays_auto.clear()

        render_holiday_list()
        update_summary()
        render_calendar()

    # ==========================================================
    # TAB 3: CHECKLIST
    # ==========================================================
    tb.Label(tab_help, text="Checklist rápido", font=FONT_H2).pack(anchor="w", pady=(0, 10))

    tips = [
        "• El reporte debe contener: Fecha / Hora / Marcación / DNI / Nombre.",
        "• Feriados se liquidan con tarifa DomYFer del empleado.",
        "• Horas extra se calculan en horas ENTERAS (sin minutos).",
        "• Si alguien no matchea: debug_no_matcheados.csv",
        "• Si matchea por nombre: debug_matcheados_por_nombre.csv",
    ]
    box = tb.Labelframe(tab_help, text="Notas", padding=14, bootstyle="light")
    box.pack(fill="both", expand=True)

    for t in tips:
        tb.Label(box, text=t, foreground="#444", wraplength=820, justify="left").pack(anchor="w", pady=6)

    # ==========================================================
    # FOOTER: GENERATE ALWAYS VISIBLE
    # ==========================================================
    footer = tb.Frame(root, padding=(20, 12))
    footer.pack(fill="x")

    tb.Label(footer, text="Cuando esté todo listo, generá el Excel con un clic.", foreground="#666").pack(side="left")

    def generate():
        rp = var_report.get().strip()
        emp = var_emp.get().strip()
        tpl = var_tpl.get().strip()

        default_out = os.path.join(
            os.path.dirname(rp) if rp else os.getcwd(),
            f"Liquidacion_Horas_Extra_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
        out_path = filedialog.asksaveasfilename(
            title="Guardar liquidación como",
            defaultextension=".xlsx",
            initialfile=os.path.basename(default_out),
            filetypes=[("Excel", "*.xlsx")]
        )
        if not out_path:
            return

        btn_generate.configure(state="disabled")
        pb.start(12)

        try:
            set_status("Procesando… leyendo empleados y reporte VeoTime…")

            if var_use_holidays.get():
                refresh_auto_holidays()
                all_h = sorted(selected_holidays_auto | selected_holidays_manual)
                cfg["holidays"] = [d.strftime("%Y-%m-%d") for d in all_h]
            else:
                cfg["holidays"] = []

            cfg["auto_holidays_enabled"] = bool(var_auto_holidays.get())
            save_config(cfg)

            emp_master = load_employee_master(emp)
            daily = read_veotime_to_daily(rp, emp_master, cfg)
            weeks = compute_overtime_from_daily(daily, cfg)
            build_output_workbook_template(weeks, out_path, tpl, cfg)

            set_status("Listo ✅ Excel generado correctamente.")
            messagebox.showinfo(
                "Listo ✅",
                f"Se generó el Excel:\n{out_path}\n\n"
                "Debug (si algo no matchea):\n"
                "• debug_veotime_head.csv\n"
                "• debug_no_matcheados.csv\n"
                "• debug_matcheados_por_nombre.csv"
            )

        except Exception as e:
            tbtxt = traceback.format_exc()
            print(tbtxt)
            with open("debug_error.txt", "w", encoding="utf-8") as f:
                f.write(tbtxt)
            set_status("Error ❌ Revisá debug_error.txt")
            messagebox.showerror("Error", f"{e}\n\nSe guardó el detalle en debug_error.txt")

        finally:
            pb.stop()
            update_generate_state()

    btn_generate = tb.Button(footer, text="Generar liquidación", command=generate, bootstyle="success", width=22)
    btn_generate.pack(side="right")

    # ==========================================================
    # Bindings + init
    # ==========================================================
    def on_any_change(*_):
        update_summary()

    var_report.trace_add("write", on_any_change)
    var_emp.trace_add("write", on_any_change)
    var_tpl.trace_add("write", on_any_change)

    # init
    refresh_auto_holidays()
    render_holiday_list()
    render_calendar()
    sync_holiday_state()
    update_summary()

    root.mainloop()


if __name__ == "__main__":
    run_app()
