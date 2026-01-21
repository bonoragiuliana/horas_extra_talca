from datetime import date, timedelta

# =========================
# OPTIONAL: feriados auto
# =========================
try:
    import holidays as pyholidays
except Exception:
    pyholidays = None


def is_holiday(d: date, cfg: dict) -> bool:
    hol = set(cfg.get("holidays", []))
    return d.strftime("%Y-%m-%d") in hol


def compute_auto_holidays_for_range(start: date, end: date, cfg: dict):
    if pyholidays is None:
        return []
    if start is None or end is None:
        return []

    years = list(range(start.year, end.year + 1))
    country = cfg.get("auto_holidays_country", "AR")
    subdiv = cfg.get("auto_holidays_subdiv", "M")
    observed = bool(cfg.get("auto_holidays_observed", True))

    try:
        hcal = pyholidays.country_holidays(
            country,
            subdiv=subdiv,
            years=years,
            observed=observed
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
