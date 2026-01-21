import json
import os
from pathlib import Path

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
    # auto feriados
    "auto_holidays_enabled": True,
    "auto_holidays_country": "AR",
    "auto_holidays_subdiv": "M",  # Mendoza
    "auto_holidays_observed": True,
    # debug opcional
    "debug": False,
    # modo merge si un empleado ya estaba:
    # "sum" suma horas en la celda
    # "replace" pisa horas en la celda
    "merge_mode": "sum",

    # --- IMPORTANTE ---
    # Dejo estos por compatibilidad (ya no los usamos para calcular extra, porque ahora:
    # - Lun-Vie: extra = horas_trab - jornada
    # - Sab/Dom/Feriado: todo lo trabajado cuenta como extra
    "saturday_standard_hours": 0,
    "sunday_standard_hours": 0,
    "holiday_standard_hours": 0,
}

BASE_DIR = Path(__file__).resolve().parent
CONFIG_FILE = str(BASE_DIR / "config_horas_extra.json")

EMP_MASTER_DEFAULT_NAME = "datos empleados.xlsx"
TEMPLATE_DEFAULT_NAME = "formatosugerido.xlsx"

TEMPLATE_SHEET_NAME = "_TEMPLATE"  # quedará OCULTA en el Excel final


def load_config() -> dict:
    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    for k, v in DEFAULT_CONFIG.items():
        if k not in cfg:
            cfg[k] = v
    return cfg


def save_config(cfg: dict) -> None:
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def debug_enabled(cfg: dict) -> bool:
    return bool(cfg.get("debug", False))


def debug_write_text(cfg: dict, filename: str, content: str) -> None:
    if not debug_enabled(cfg):
        return
    with open(filename, "w", encoding="utf-8") as f:
        f.write(content)


def debug_write_df(cfg: dict, filename: str, df) -> None:
    if not debug_enabled(cfg):
        return
    df.to_csv(filename, index=False, encoding="utf-8-sig")
