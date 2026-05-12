from config.settings import save_config
from readers.employee_reader import load_employee_master
from readers.veotime_reader import read_veotime_to_daily
from services.overtime_service import compute_overtime_from_daily
from writers.excel_writer import update_or_build_output_workbook

def run_liquidacion(report_path: str, emp_path: str, tpl_path: str, out_path: str, holidays: list, cfg: dict) -> None:
    """
    Orquesta todo el flujo:
    1. Actualiza feriados en config
    2. Carga empleados
    3. Parsea VeoTime → eventos diarios
    4. Calcula horas extra por semana
    5. Escribe/actualiza Excel de salida
    """
    # Normalizar feriados a lista de strings YYYY-MM-DD
    cfg["holidays"] = [str(h).strip() for h in holidays if h]
    save_config(cfg)

    emp_master = load_employee_master(emp_path)
    daily = read_veotime_to_daily(report_path, emp_master, cfg)
    weeks = compute_overtime_from_daily(daily, cfg)
    update_or_build_output_workbook(weeks, out_path, tpl_path, cfg)