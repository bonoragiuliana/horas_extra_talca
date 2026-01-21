import os
import re
from datetime import datetime, timedelta

import pandas as pd

from config_app import debug_write_df, debug_write_text
from utils_app import (
    normalize_text, guess_col, parse_date, parse_time_only,
    clean_id, extract_id_parts, name_keys, round_dt_to_nearest_hour
)


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
        has_id = ("dni" in text) or ("documento" in text) or ("legajo" in text) or ("id" in text)
        has_marc = "marc" in text or "tipo" in text

        if has_fecha and has_hora and has_id and has_marc:
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

    cols_norm = [normalize_text(c) for c in df.columns]

    i_fecha = guess_col(cols_norm, ["fecha"])
    i_hora = guess_col(cols_norm, ["hora", "time"])
    i_marc = guess_col(cols_norm, ["marcaci", "marc", "tipo"])
    i_id = guess_col(cols_norm, ["dni", "documento", "legajo", "id"])
    i_nombre = guess_col(cols_norm, ["nombre", "empleado", "apellido", "colaborador"])

    missing = []
    if i_fecha is None: missing.append("Fecha")
    if i_hora is None: missing.append("Hora")
    if i_marc is None: missing.append("Marcación")
    if i_id is None: missing.append("DNI/ID/Legajo")
    if i_nombre is None: missing.append("Nombre")

    if missing:
        debug_write_df(cfg, "debug_veotime_head.csv", df.head(200))
        debug_write_text(cfg, "debug_columnas.txt", "Columnas detectadas:\n" + "\n".join(f"- {c}" for c in df.columns))
        raise RuntimeError(
            "No pude detectar columnas clave en el reporte VeoTime.\n"
            f"Faltan: {', '.join(missing)}"
        )

    def norm_tipo(x):
        t = normalize_text(x)
        if "entrada" in t or "ingreso" in t or t.startswith("ent"):
            return "entrada"
        if "salida" in t or "egreso" in t or t.startswith("sal"):
            return "salida"
        return ""

    events = pd.DataFrame()
    events["fecha"] = df.iloc[:, i_fecha].apply(parse_date)
    events["hora"] = df.iloc[:, i_hora].apply(parse_time_only)
    events["tipo"] = df.iloc[:, i_marc].apply(norm_tipo)

    events["raw_id_rep"] = df.iloc[:, i_id].apply(clean_id).astype(str).str.strip()
    parts = events["raw_id_rep"].apply(lambda x: pd.Series(
        extract_id_parts(x),
        index=["id_key_rep", "id_digits_rep", "id_cuil11_rep", "id_dni8_rep"]
    ))
    events = pd.concat([events, parts], axis=1)

    events["nombre_rep"] = df.iloc[:, i_nombre].astype(str).str.strip()
    nk = events["nombre_rep"].apply(
        lambda s: pd.Series(name_keys(s), index=["nombre_norm_rep", "nombre_first2_rep", "nombre_last2_rep"])
    )
    events = pd.concat([events, nk], axis=1)

    # filtros
    events = events.dropna(subset=["fecha", "hora"])
    events = events[events["tipo"].isin(["entrada", "salida"])]
    events = events[events["id_key_rep"] != ""]

    if events.empty:
        debug_write_df(cfg, "debug_veotime_head.csv", df.head(200))
        raise RuntimeError("No quedaron eventos válidos (Entrada/Salida).")

    emp = emp_master.copy()

    # ---- mapeos por CLAVES ÚNICAS ----
    def unique_map(key_col):
        s = emp[[key_col, "midx"]].copy()
        s[key_col] = s[key_col].fillna("").astype(str).str.strip()
        s = s[s[key_col] != ""]
        if s.empty:
            return {}
        vc = s[key_col].value_counts()
        uniques = set(vc[vc == 1].index.astype(str))
        s = s[s[key_col].isin(uniques)]
        return dict(zip(s[key_col].astype(str), s["midx"]))

    map_digits = unique_map("id_digits")
    map_cuil = unique_map("id_cuil11")
    map_dni8 = unique_map("id_dni8")
    map_key = unique_map("id_key")

    map_name_full = unique_map("nombre_norm")
    map_name_f2 = unique_map("nombre_first2")
    map_name_l2 = unique_map("nombre_last2")

    # ---- matching por ID primero ----
    events["midx"] = events["id_digits_rep"].astype(str).map(map_digits)

    mask = events["midx"].isna()
    events.loc[mask, "midx"] = events.loc[mask, "id_cuil11_rep"].astype(str).map(map_cuil)

    mask = events["midx"].isna()
    events.loc[mask, "midx"] = events.loc[mask, "id_dni8_rep"].astype(str).map(map_dni8)

    mask = events["midx"].isna()
    events.loc[mask, "midx"] = events.loc[mask, "id_key_rep"].astype(str).map(map_key)

    # ---- si no matcheó por ID, matcheo por nombre (robusto) ----
    mask = events["midx"].isna()
    events.loc[mask, "midx"] = events.loc[mask, "nombre_norm_rep"].astype(str).map(map_name_full)

    mask = events["midx"].isna()
    events.loc[mask, "midx"] = events.loc[mask, "nombre_first2_rep"].astype(str).map(map_name_f2)

    mask = events["midx"].isna()
    events.loc[mask, "midx"] = events.loc[mask, "nombre_last2_rep"].astype(str).map(map_name_l2)

    merged = events.merge(emp, how="left", on="midx")

    still_no = merged["nombre_master"].isna()
    if still_no.any():
        debug_write_df(cfg, "debug_no_matcheados.csv",
                       merged.loc[still_no, ["raw_id_rep", "id_key_rep", "nombre_rep"]].drop_duplicates())

    # finales
    merged["empresa_final"] = merged["empresa_master"].fillna("").astype(str).str.strip()
    merged["empresa_final"] = merged["empresa_final"].where(
        merged["empresa_final"] != "",
        cfg.get("company_name", "TALCA")
    )

    merged["sector_final"] = merged["sector_master"].fillna("").astype(str).str.strip()
    merged["sector_final"] = merged["sector_final"].where(merged["sector_final"] != "", "desconocido")
    merged["sector_final"] = merged["sector_final"].apply(normalize_text)

    merged["nombre_final"] = merged["nombre_master"].fillna("").astype(str).str.strip()
    merged["nombre_final"] = merged["nombre_final"].where(
        merged["nombre_final"] != "",
        merged["nombre_rep"].fillna("").astype(str)
    )

    merged["jornada_weekday"] = merged["jornada_weekday"].fillna(float(cfg.get("default_jornada_weekday", 8)))
    merged.loc[merged["jornada_weekday"] <= 0, "jornada_weekday"] = float(cfg.get("default_jornada_weekday", 8))

    merged["rate_lav"] = merged["rate_lav"].fillna(0.0)
    merged["rate_sab"] = merged["rate_sab"].fillna(0.0)
    merged["rate_domfer"] = merged["rate_domfer"].fillna(0.0)

    # ID final: escribe SIEMPRE el ID del master si existe
    merged["id_display_final"] = merged["id_master_display"].fillna("").astype(str).str.strip()
    merged["id_display_final"] = merged["id_display_final"].where(
        merged["id_display_final"] != "",
        merged["raw_id_rep"].fillna("").astype(str).str.strip()
    )

    # clave interna estable para agrupar
    merged["id_key_final"] = merged["id_key"].fillna("").astype(str).str.strip()
    merged["id_key_final"] = merged["id_key_final"].where(
        merged["id_key_final"] != "",
        merged["id_key_rep"].fillna("").astype(str).str.strip()
    )

    # =========================================================
    # CONSTRUIR HORAS DIARIAS (TURNO NOCTURNO)
    # - Redondeo a hora más cercana (no minutos)
    # - Empareja Entrada -> Salida aunque sea al día siguiente
    # - Asigna TODAS las horas al día de la ENTRADA (como hace RRHH)
    # - Si cruza medianoche: suma +1 hora extra (bonus nocturno) al día de entrada
    # =========================================================

    merged["dt_raw"] = merged.apply(
        lambda r: datetime.combine(r["fecha"], r["hora"])
        if (r["fecha"] is not None and r["hora"] is not None)
        else None,
        axis=1
    )
    merged = merged.dropna(subset=["dt_raw"]).copy()

    merged["dt"] = merged["dt_raw"].apply(round_dt_to_nearest_hour)

    acc = {}

    def get_rec(static, dte: date):
        k = (static["dni"], dte)
        if k not in acc:
            acc[k] = {
                "dni": static["dni"],
                "dni_display": static["dni_display"],
                "empresa": static["empresa"],
                "nombre": static["nombre"],
                "sector": static["sector"],
                "fecha": dte,
                "horas_trab": 0,
                "night_bonus": 0,
                "jornada_weekday": static["jornada_weekday"],
                "rate_lav": static["rate_lav"],
                "rate_sab": static["rate_sab"],
                "rate_domfer": static["rate_domfer"],
            }
        return acc[k]

    emp_group_cols = [
        "id_key_final", "id_display_final",
        "empresa_final", "nombre_final", "sector_final",
        "jornada_weekday", "rate_lav", "rate_sab", "rate_domfer"
    ]

    for key, sub in merged.groupby(emp_group_cols, dropna=False):
        (idkey, iddisp, empresa, nombre, sector, jornada, rlav, rsab, rdom) = key

        static = {
            "dni": str(idkey).strip(),
            "dni_display": str(iddisp).strip() if str(iddisp).strip() else str(idkey).strip(),
            "empresa": str(empresa),
            "nombre": str(nombre),
            "sector": str(sector),
            "jornada_weekday": float(jornada),
            "rate_lav": float(rlav),
            "rate_sab": float(rsab),
            "rate_domfer": float(rdom),
        }

        if not static["dni"]:
            continue

        sub = sub.sort_values("dt")

        open_entry = None

        for _, r in sub.iterrows():
            if r["tipo"] == "entrada":
                open_entry = r["dt"]
                continue

            if r["tipo"] == "salida" and open_entry is not None:
                end_dt = r["dt"]

                while end_dt <= open_entry:
                    end_dt = end_dt + timedelta(days=1)

                hs = int((end_dt - open_entry).total_seconds() // 3600)
                if hs < 0:
                    hs = 0

                day_assigned = open_entry.date()
                rec = get_rec(static, day_assigned)
                rec["horas_trab"] += hs

                if open_entry.date() != end_dt.date():
                    rec["night_bonus"] += 1

                open_entry = None

    daily = pd.DataFrame(list(acc.values()))
    if daily.empty:
        raise RuntimeError("No pude construir horas diarias (no se formaron pares Entrada/Salida).")
    return daily
