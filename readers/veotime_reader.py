import os
import re
from datetime import datetime, timedelta, date
import pandas as pd
from config.settings import debug_write_df, debug_write_text
from utils.text_utils import normalize_text, name_keys
from utils.date_utils import parse_date, parse_time_only, round_dt_to_nearest_hour
from utils.id_utils import clean_id, extract_id_parts, guess_col

def read_report_raw(path: str) -> pd.DataFrame:
    with open(path, "rb") as f: head = f.read(8192).lstrip()
    head_low = head.lower()
    if head.startswith(b"PK"): return pd.read_excel(path, engine="openpyxl", header=None)
    if head.startswith(b"\xD0\xCF\x11\xE0"): return pd.read_excel(path, engine="xlrd", header=None)
    if b"<html" in head_low or b"<table" in head_low:
        dfs = pd.read_html(path, header=None)
        if not dfs: raise RuntimeError("HTML sin tablas.")
        def score_tab(t):
            txt = " ".join([str(x) for x in t.columns] + t.astype(str).values.ravel().tolist())
            txt = normalize_text(txt)
            return sum(1 for kw in ["fecha", "hora", "dni", "marc", "entrada", "salida"] if kw in txt)
        scored = [(score_tab(t), t) for t in dfs]
        scored.sort(key=lambda x: x[0], reverse=True)
        best = scored[0][1].copy()
        best.columns = list(range(best.shape[1]))
        return best.reset_index(drop=True)
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx": return pd.read_excel(path, engine="openpyxl", header=None)
    if ext == ".xls": return pd.read_excel(path, engine="xlrd", header=None)
    raise RuntimeError("Formato no soportado. Usá .xls o .xlsx.")

def find_header_row(df_raw, max_rows=80):
    for r in range(min(max_rows, len(df_raw))):
        row = df_raw.iloc[r].astype(str).tolist()
        text = " ".join(normalize_text(x) for x in row)
        if "fecha" in text and ("hora" in text or "time" in text) and \
           (("dni" in text) or ("documento" in text) or ("legajo" in text) or ("id" in text)) and \
           ("marc" in text or "tipo" in text):
            return r
    return None

def apply_header_row(df_raw, header_row):
    headers = df_raw.iloc[header_row].tolist()
    df = df_raw.iloc[header_row + 1:].copy()
    df.columns = [str(h).strip() for h in headers]
    return df.dropna(axis=1, how="all").reset_index(drop=True)

def read_veotime_to_daily(path: str, emp_master: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    df_raw = read_report_raw(path)
    header_row = find_header_row(df_raw)
    df = apply_header_row(df_raw, header_row) if header_row is not None else df_raw.copy()
    if header_row is None:
        df.columns = [str(c).strip() for c in df.iloc[0].tolist()]
        df = df.iloc[1:].reset_index(drop=True)

    cols_norm = [normalize_text(c) for c in df.columns]
    i_fecha = guess_col(cols_norm, ["fecha"])
    i_hora = guess_col(cols_norm, ["hora", "time"])
    i_marc = guess_col(cols_norm, ["marcaci", "marc", "tipo"])
    i_id = guess_col(cols_norm, ["dni", "documento", "legajo", "id"])
    i_nombre = guess_col(cols_norm, ["nombre", "empleado", "apellido", "colaborador"])

    missing = [n for n, i in [("Fecha",i_fecha),("Hora",i_hora),("Marcación",i_marc),("DNI/ID",i_id),("Nombre",i_nombre)] if i is None]
    if missing:
        debug_write_df(cfg, "debug_veotime_head.csv", df.head(200))
        debug_write_text(cfg, "debug_columnas.txt", "Columnas detectadas:\n" + "\n".join(f"- {c}" for c in df.columns))
        raise RuntimeError(f"No pude detectar columnas clave. Faltan: {', '.join(missing)}")

    def norm_tipo(x):
        t = normalize_text(x)
        if "entrada" in t or "ingreso" in t or t.startswith("ent"): return "entrada"
        if "salida" in t or "egreso" in t or t.startswith("sal"): return "salida"
        return ""

    events = pd.DataFrame()
    events["fecha"] = df.iloc[:, i_fecha].apply(parse_date)
    events["hora"] = df.iloc[:, i_hora].apply(parse_time_only)
    events["tipo"] = df.iloc[:, i_marc].apply(norm_tipo)
    events["raw_id_rep"] = df.iloc[:, i_id].apply(clean_id).astype(str).str.strip()
    parts = events["raw_id_rep"].apply(lambda x: pd.Series(extract_id_parts(x), index=["id_key_rep","id_digits_rep","id_cuil11_rep","id_dni8_rep"]))
    events = pd.concat([events, parts], axis=1)
    events["nombre_rep"] = df.iloc[:, i_nombre].astype(str).str.strip()
    nk = events["nombre_rep"].apply(lambda s: pd.Series(name_keys(s), index=["nombre_norm_rep","nombre_first2_rep","nombre_last2_rep"]))
    events = pd.concat([events, nk], axis=1)

    events = events.dropna(subset=["fecha", "hora"])
    events = events[events["tipo"].isin(["entrada", "salida"])]
    events = events[events["id_key_rep"] != ""]
    if events.empty:
        debug_write_df(cfg, "debug_veotime_head.csv", df.head(200))
        raise RuntimeError("No quedaron eventos válidos (Entrada/Salida).")

    emp = emp_master.copy()
    def unique_map(key_col):
        s = emp[[key_col, "midx"]].copy()
        s[key_col] = s[key_col].fillna("").astype(str).str.strip()
        s = s[s[key_col] != ""]
        if s.empty: return {}
        vc = s[key_col].value_counts()
        uniques = set(vc[vc == 1].index.astype(str))
        return dict(zip(s[s[key_col].isin(uniques)][key_col].astype(str), s[s[key_col].isin(uniques)]["midx"]))

    maps = {
        "id_digits": unique_map("id_digits"), "id_cuil11": unique_map("id_cuil11"),
        "id_dni8": unique_map("id_dni8"), "id_key": unique_map("id_key"),
        "nombre_norm": unique_map("nombre_norm"), "nombre_first2": unique_map("nombre_first2"),
        "nombre_last2": unique_map("nombre_last2")
    }
    rep_cols = ["id_digits_rep","id_cuil11_rep","id_dni8_rep","id_key_rep","nombre_norm_rep","nombre_first2_rep","nombre_last2_rep"]
    map_keys = ["id_digits","id_cuil11","id_dni8","id_key","nombre_norm","nombre_first2","nombre_last2"]
    for rc, mk in zip(rep_cols, map_keys):
        mask = events["midx"].isna() if "midx" in events.columns else pd.Series(True, index=events.index)
        if mask.any():
            events.loc[mask, "midx"] = events.loc[mask, rc].astype(str).map(maps[mk])

    merged = events.merge(emp, how="left", on="midx")
    still_no = merged["nombre_master"].isna()
    if still_no.any():
        debug_write_df(cfg, "debug_no_matcheados.csv", merged.loc[still_no, ["raw_id_rep","id_key_rep","nombre_rep"]].drop_duplicates())

    merged["empresa_final"] = merged["empresa_master"].fillna("").astype(str).str.strip()
    merged["empresa_final"] = merged["empresa_final"].where(merged["empresa_final"] != "", cfg.get("company_name", "TALCA"))
    merged["sector_final"] = merged["sector_master"].fillna("").astype(str).str.strip().apply(normalize_text).where(merged["sector_master"].notna(), "desconocido")
    merged["nombre_final"] = merged["nombre_master"].fillna("").astype(str).str.strip()
    merged["nombre_final"] = merged["nombre_final"].where(merged["nombre_final"] != "", merged["nombre_rep"].fillna("").astype(str))
    merged["jornada_weekday"] = merged["jornada_weekday"].fillna(float(cfg.get("default_jornada_weekday", 8)))

    merged["dt_raw"] = merged.apply(lambda r: datetime.combine(r["fecha"], r["hora"]) if r["fecha"] and r["hora"] else None, axis=1)
    merged = merged.dropna(subset=["dt_raw"]).copy()
    merged["dt"] = merged["dt_raw"].apply(round_dt_to_nearest_hour)

    acc = {}
    def get_rec(static, dte: date):
        k = (static["dni"], dte)
        if k not in acc:
            acc[k] = {"dni":static["dni"],"dni_display":static["dni_display"],"empresa":static["empresa"],
                      "nombre":static["nombre"],"sector":static["sector"],"fecha":dte,
                      "horas_trab":0,"night_bonus":0,"jornada_weekday":static["jornada_weekday"],
                      "rate_lav":static["rate_lav"],"rate_sab":static["rate_sab"],"rate_domfer":static["rate_domfer"]}
        return acc[k]

    for key, sub in merged.groupby(["id_key_final","id_display_final","empresa_final","nombre_final","sector_final","jornada_weekday","rate_lav","rate_sab","rate_domfer"], dropna=False):
        (idkey,iddisp,empresa,nombre,sector,jornada,rlav,rsab,rdom) = key
        static = {"dni":str(idkey).strip(),"dni_display":str(iddisp).strip() or str(idkey).strip(),
                  "empresa":str(empresa),"nombre":str(nombre),"sector":str(sector),
                  "jornada_weekday":float(jornada),"rate_lav":float(rlav),"rate_sab":float(rsab),"rate_domfer":float(rdom)}
        if not static["dni"]: continue
        sub = sub.sort_values("dt")
        open_entry = None
        for _, r in sub.iterrows():
            if r["tipo"] == "entrada":
                open_entry = r["dt"]
                continue
            if r["tipo"] == "salida" and open_entry is not None:
                end_dt = r["dt"]
                # Lógica reconstruida: si cruce medianoche, asigna al día de entrada
                entry_date = open_entry.date()
                hours = (end_dt - open_entry).total_seconds() / 3600
                rec = get_rec(static, entry_date)
                rec["horas_trab"] += int(hours)
                if end_dt.date() > entry_date:
                    rec["night_bonus"] += 1  # +1 hora extra bonus nocturno
                open_entry = None

    records = list(acc.values())
    return pd.DataFrame(records)