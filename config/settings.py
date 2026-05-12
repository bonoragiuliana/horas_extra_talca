import json
import os
from pathlib import Path

DEFAULT_CONFIG = {
    "company_name": "TALCA",
    "default_jornada_weekday": 8,
    "employees_excel_path": "",
    "template_excel_path": "",
    "holidays": [],
    "auto_holidays_enabled": True,
    "auto_holidays_country": "AR",
    "auto_holidays_subdiv": "M",
    "auto_holidays_observed": True,
    "debug": False,
    "merge_mode": "sum",
    "saturday_standard_hours": 0,
    "sunday_standard_hours": 0,
    "holiday_standard_hours": 0,
}

PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_FILE = PROJECT_ROOT / "config" / "config_horas_extra.json"
EMP_MASTER_DEFAULT_NAME = "datos empleados.xlsx"
TEMPLATE_DEFAULT_NAME = "formatosugerido.xlsx"
TEMPLATE_SHEET_NAME = "_TEMPLATE"

def load_config() -> dict:
    if not CONFIG_FILE.exists():
        CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    for k, v in DEFAULT_CONFIG.items():
        if k not in cfg: cfg[k] = v
    return cfg

def save_config(cfg: dict) -> None:
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

def debug_write_text(cfg: dict, filename: str, content: str) -> None:
    if not cfg.get("debug", False): return
    with open(filename, "w", encoding="utf-8") as f: f.write(content)

def debug_write_df(cfg: dict, filename: str, df) -> None:
    if not cfg.get("debug", False): return
    df.to_csv(filename, index=False, encoding="utf-8-sig")