from datetime import date, timedelta
from typing import List

# try-except envuelve la importación de todo el módulo holidays.
# Si falla la instalación, el programa continúa, pero las funciones
# de autodetect feriados devolverán listas vacías.
try:
    # pyrefly: ignore [missing-import]
    import holidays as pyholidays
except Exception:
    pyholidays = None


def is_holiday(d: date, cfg: dict) -> bool:
    """Verifica si una fecha está en la lista de feriados del config."""
    if not isinstance(d, date):
        return False
    d_str = d.strftime("%Y-%m-%d")
    return d_str in set(cfg.get("holidays", []))


def compute_auto_holidays_for_range(start: date, end: date, cfg: dict) -> List[date]:
    """Calcula feriados automáticos (Argentina/Mendoza) en un rango."""
    if pyholidays is None:
        return []
    if not start or not end or start > end:
        return []

    years = list(range(start.year, end.year + 1))
    country = cfg.get("auto_holidays_country", "AR")
    subdiv = cfg.get("auto_holidays_subdiv", "M")
    observed = bool(cfg.get("auto_holidays_observed", True))

    try:
        hcal = pyholidays.country_holidays(
            country, subdiv=subdiv, years=years, observed=observed
        )
    except Exception:
        return []

    out = []
    cur = start
    while cur <= end:
        if cur in hcal:
            out.append(cur)
        cur += timedelta(days=1)
    return out