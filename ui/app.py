import os
import tkinter as tk
from tkinter import messagebox
from datetime import datetime, date
import ttkbootstrap as tb
from config.settings import load_config, save_config, EMP_MASTER_DEFAULT_NAME, TEMPLATE_DEFAULT_NAME

class UiState:
    def __init__(self, root):
        self.cfg = load_config()
        # Variables de ruta
        self.var_report = tk.StringVar(master=root, value="")
        self.var_emp = tk.StringVar(master=root, value=self.cfg.get("employees_excel_path", ""))
        self.var_tpl = tk.StringVar(master=root, value=self.cfg.get("template_excel_path", ""))
        self.var_oeste = tk.StringVar(master=root, value="")
        self.var_consultora = tk.StringVar(master=root, value="")
        self.var_tpl_print = tk.StringVar(master=root, value=self.cfg.get("print_template_excel_path", ""))
        
        # Variables de control
        self.var_use_holidays = tk.BooleanVar(master=root, value=False)
        self.var_auto_holidays = tk.BooleanVar(master=root, value=bool(self.cfg.get("auto_holidays_enabled", True)))
        self.cal_year = tk.IntVar(master=root, value=datetime.now().year)
        self.cal_month = tk.IntVar(master=root, value=datetime.now().month)
        
        # Estado y resumen
        self.status_var = tk.StringVar(master=root, value="Paso 1: elegí el reporte de VeoTime.")
        self.summary_range_var = tk.StringVar(master=root, value="—")
        self.summary_emps_var = tk.StringVar(master=root, value="—")
        self.summary_holidays_var = tk.StringVar(master=root, value="—")
        
        # Datos de feriados
        self.selected_holidays_manual = set()
        self.selected_holidays_auto = set()
        
        # Referencias UI para actualización
        self.btn_generate = None
        self.btn_print = None
        self.pb = None
        self.lbl_month = None
        self.range_lbl = None
        self.cal_card_grid = None
        self.holiday_lb = None

def run_app():
    try:
        import ttkbootstrap as tb
    except Exception:
        r = tk.Tk(); r.withdraw()
        messagebox.showerror("Falta ttkbootstrap", "Instalá: pip install ttkbootstrap")
        return

    THEME = "minty"
    root = tb.Window(themename=THEME)
    root.title("Liquidación de Horas Extra · TALCA")
    root.geometry("1120x720")
    root.minsize(1040, 680)

    st = UiState(root)

    from ui.tab_liquidacion import render_tab_liquidacion
    from ui.tab_feriados import render_tab_feriados
    from ui.tab_impresion import render_tab_impresion

    # Render layout base
    render_header(root, st)
    nb = tb.Notebook(root, bootstyle="primary")
    nb.pack(fill="both", expand=True, padx=20, pady=(10,0))
    
    render_tab_liquidacion(nb, st)
    render_tab_feriados(nb, st)
    render_tab_impresion(nb, st)
    
    render_sidebar(root, st)
    render_footer(root, st)

    # Bindings
    st.var_report.trace_add("write", lambda *_: update_summary(st))
    st.var_emp.trace_add("write", lambda *_: update_summary(st))
    st.var_tpl.trace_add("write", lambda *_: update_summary(st))
    
    # Init
    from ui.tab_feriados import refresh_auto_holidays, render_calendar, render_holiday_list, sync_holiday_state
    refresh_auto_holidays(st)
    render_holiday_list(st)
    render_calendar(st)
    sync_holiday_state(st)
    update_summary(st)
    
    root.mainloop()

def render_header(root, st):
    header = tb.Frame(root, padding=(20, 18))
    header.pack(fill="x")
    left_h = tb.Frame(header)
    left_h.pack(side="left", fill="x", expand=True)
    tb.Label(left_h, text="Liquidación de Horas Extra", font=("Segoe UI", 20, "bold")).pack(anchor="w")
    tb.Label(left_h, text="Generación automática desde VeoTime → Excel RRHH", font=("Segoe UI", 10), foreground="#666").pack(anchor="w", pady=(4,0))
    
    right_h = tb.Frame(header)
    right_h.pack(side="right")
    tb.Label(right_h, text="TALCA", bootstyle="secondary-inverse", padding=(10, 4)).pack(side="right")

def render_sidebar(root, st):
    sidebar = tb.Frame(root)
    sidebar.pack(side="right", fill="y", padx=(10,0))
    sidebar.configure(width=320)
    sidebar.pack_propagate(False)

    summary = tb.Labelframe(sidebar, text="RESUMEN", padding=16, bootstyle="secondary")
    summary.pack(fill="x")
    tb.Label(summary, text="Rango detectado", foreground="#666").pack(side="left")
    tb.Label(summary, textvariable=st.summary_range_var, font=("Segoe UI", 10, "bold")).pack(side="right")
    tb.Label(summary, text="Empleados (aprox.)", foreground="#666").pack(side="left", pady=4)
    tb.Label(summary, textvariable=st.summary_emps_var, font=("Segoe UI", 10, "bold")).pack(side="right", pady=4)
    tb.Label(summary, text="Feriados (dd/mm)", foreground="#666").pack(side="left", pady=4)
    tb.Label(summary, textvariable=st.summary_holidays_var, font=("Segoe UI", 10, "bold")).pack(side="right", pady=4)

    status_card = tb.Labelframe(sidebar, text="ESTADO", padding=16, bootstyle="light")
    status_card.pack(fill="both", expand=True, pady=(14,0))
    tb.Label(status_card, textvariable=st.status_var, wraplength=280, justify="left").pack(anchor="w")
    st.pb = tb.Progressbar(status_card, mode="indeterminate", bootstyle="success-striped")
    st.pb.pack(fill="x", pady=(14,0))

def render_footer(root, st):
    footer = tb.Frame(root, padding=(20, 12))
    footer.pack(fill="x")
    tb.Label(footer, text="Podés guardar un archivo nuevo o elegir uno existente para agregar empleados.", foreground="#666").pack(side="left")
    
    def on_generate_click():
        from ui.tab_liquidacion import generate as tab_generate
        tab_generate(st)

    st.btn_generate = tb.Button(footer, text="Generar / Actualizar", command=on_generate_click, bootstyle="success", width=22)
    st.btn_generate.pack(side="right")

def update_summary(st):
    from readers.veotime_reader import read_report_raw, find_header_row, apply_header_row
    from utils.text_utils import normalize_text
    from utils.date_utils import parse_date
    from utils.id_utils import clean_id, id_key_from_any

    rp = st.var_report.get().strip()
    if rp and os.path.exists(rp):
        try:
            df_raw = read_report_raw(rp)
            header_row = find_header_row(df_raw)
            if header_row is not None:
                df = apply_header_row(df_raw, header_row)
            else:
                df = df_raw.copy()
                df.columns = [str(c).strip() for c in df.iloc[0].tolist()]
                df = df.iloc[1:].copy()

            cols_norm = [normalize_text(c) for c in df.columns]
            def guess_col(kw):
                for i, c in enumerate(cols_norm):
                    if any(k in c for k in kw): return i
                return None

            i_fecha = guess_col(["fecha"])
            i_id = guess_col(["dni", "documento", "legajo", "id"])

            if i_fecha is not None:
                fechas = df.iloc[:, i_fecha].apply(parse_date).dropna()
                if not fechas.empty:
                    start, end = min(fechas), max(fechas)
                    st.summary_range_var.set(f"{start.strftime('%d/%m/%Y')} → {end.strftime('%d/%m/%Y')}")
                    if i_id is not None:
                        tmp = df.iloc[:, i_id].apply(clean_id)
                        keys = tmp.apply(id_key_from_any)
                        keys = keys[keys != ""]
                        st.summary_emps_var.set(str(int(keys.nunique())) if not keys.empty else "—")
                    else:
                        st.summary_emps_var.set("—")
                else:
                    st.summary_range_var.set("—"); st.summary_emps_var.set("—")
            else:
                st.summary_range_var.set("—"); st.summary_emps_var.set("—")
        except Exception:
            st.summary_range_var.set("—"); st.summary_emps_var.set("—")
    else:
        st.summary_range_var.set("—"); st.summary_emps_var.set("—")

    all_h = sorted(st.selected_holidays_auto | st.selected_holidays_manual) if st.var_use_holidays.get() else []
    st.summary_holidays_var.set(", ".join(d.strftime("%d/%m") for d in all_h) if all_h else "—")

    # Update buttons
    ok_gen = bool(st.var_report.get().strip() and os.path.exists(st.var_report.get().strip()) and
                  st.var_emp.get().strip() and os.path.exists(st.var_emp.get().strip()) and
                  st.var_tpl.get().strip() and os.path.exists(st.var_tpl.get().strip()))
    if st.btn_generate: st.btn_generate.configure(state="normal" if ok_gen else "disabled")

    ok_pr = bool(st.var_oeste.get().strip() and os.path.exists(st.var_oeste.get().strip()) and
                 st.var_consultora.get().strip() and os.path.exists(st.var_consultora.get().strip()) and
                 st.var_tpl_print.get().strip() and os.path.exists(st.var_tpl_print.get().strip()))
    if st.btn_print: st.btn_print.configure(state="normal" if ok_pr else "disabled")