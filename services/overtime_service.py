from datetime import timedelta, date
import pandas as pd
from services.holidays_service import is_holiday

DOW_MAP = {0: "Mon", 1: "Tue", 2: "Wed", 3: "Thu", 4: "Fri", 5: "Sat", 6: "Sun"}

def monday_of_week(d: date) -> date:
    return d - timedelta(days=d.weekday())

def week_key_sun_to_sat(d: date) -> date:
    # Domingo pertenece a la semana que arranca el lunes siguiente
    if d.weekday() == 6:
        d += timedelta(days=1)
    return monday_of_week(d)

def compute_overtime_from_daily(daily: pd.DataFrame, cfg: dict) -> dict[date, pd.DataFrame]:
    if daily.empty:
        return {}

    d = daily.copy()
    d["semana_lunes"] = d["fecha"].apply(week_key_sun_to_sat)
    d["dow"] = d["fecha"].apply(lambda x: x.weekday())
    d["dow_key"] = d["dow"].map(DOW_MAP)

    # Normalización de columnas numéricas
    d["horas_trab"] = pd.to_numeric(d["horas_trab"], errors="coerce").fillna(0).astype(int)
    d["jornada_weekday"] = pd.to_numeric(d["jornada_weekday"], errors="coerce").fillna(0).astype(int)
    d["night_bonus"] = pd.to_numeric(d.get("night_bonus", 0), errors="coerce").fillna(0).astype(int)

    # Cálculo: horas extra = (trabajadas - jornada) + bonus nocturno (mínimo 0)
    base_extra = (d["horas_trab"] - d["jornada_weekday"]).clip(lower=0).astype(int)
    d["horas_extra"] = (base_extra + d["night_bonus"]).astype(int)

    # Aplicación de tarifas según tipo de día
    def rate_for(row):
        if is_holiday(row["fecha"], cfg) or row["dow_key"] == "Sun":
            return float(row["rate_domfer"])
        if row["dow_key"] == "Sat":
            return float(row["rate_sab"])
        return float(row["rate_lav"])

    d["tarifa_he"] = d.apply(rate_for, axis=1)
    d["costo"] = d["horas_extra"] * d["tarifa_he"]

    # Agrupación por semana (Lunes a Domingo)
    weeks = {}
    for week_start, sub in d.groupby("semana_lunes"):
        weeks[week_start] = sub.copy()
    return weeks