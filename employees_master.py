import pandas as pd

from utils_app import normalize_text, clean_id, extract_id_parts, name_keys


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
    c_id = col_like(["id"])  # en tu master se llama ID
    c_name = col_like(["apellido", "nombre"])
    c_jornada = col_like(["jornada"])
    c_lav = col_like(["lav"])
    c_sab = col_like(["sab"])
    c_domfer = col_like(["domyfer", "dom", "fer"])

    need = [c_empresa, c_sector, c_id, c_name, c_jornada, c_lav, c_sab, c_domfer]
    if any(x is None for x in need):
        raise RuntimeError(
            "En datos empleados.xlsx faltan columnas esperadas.\n"
            "Necesito: EMPRESA, SECTOR, ID, APELLIDO Y NOMBRE, JORNADA, $/hs LaV, $/hs Sab, $/hs DomYFer."
        )

    out = pd.DataFrame()
    out["id_master_display"] = df[c_id].apply(clean_id).astype(str).str.strip()

    parts = out["id_master_display"].apply(lambda x: pd.Series(
        extract_id_parts(x),
        index=["id_key", "id_digits", "id_cuil11", "id_dni8"]
    ))
    out = pd.concat([out, parts], axis=1)

    out["empresa_master"] = df[c_empresa].fillna("").astype(str).str.strip()
    out["sector_master"] = df[c_sector].fillna("").astype(str).str.strip().apply(normalize_text)

    out["nombre_master"] = df[c_name].fillna("").astype(str).str.strip()
    nk = out["nombre_master"].apply(
        lambda s: pd.Series(name_keys(s), index=["nombre_norm", "nombre_first2", "nombre_last2"])
    )
    out = pd.concat([out, nk], axis=1)

    out["jornada_weekday"] = pd.to_numeric(df[c_jornada], errors="coerce").fillna(0).astype(float)
    out["rate_lav"] = pd.to_numeric(df[c_lav], errors="coerce").fillna(0).astype(float)
    out["rate_sab"] = pd.to_numeric(df[c_sab], errors="coerce").fillna(0).astype(float)
    out["rate_domfer"] = pd.to_numeric(df[c_domfer], errors="coerce").fillna(0).astype(float)

    out = out[out["id_key"] != ""].copy()
    out = out.reset_index(drop=True)
    out["midx"] = out.index  # id interno estable

    return out
