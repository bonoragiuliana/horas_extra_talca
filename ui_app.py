import os
import calendar
import traceback
from datetime import datetime, date
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox

from config_app import load_config, save_config, EMP_MASTER_DEFAULT_NAME, TEMPLATE_DEFAULT_NAME
from holidays_auto import compute_auto_holidays_for_range
from employees_master import load_employee_master
from veotime_reader import read_report_raw, find_header_row, apply_header_row, read_veotime_to_daily
from overtime_calc import compute_overtime_from_daily
from excel_writer import update_or_build_output_workbook
from utils_app import normalize_text, guess_col, parse_date, clean_id, id_key_from_any


def run_app():
    cfg = load_config()

    try:
        import ttkbootstrap as tb
    except Exception:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Falta ttkbootstrap", "Instalá: pip install ttkbootstrap")
        return

    THEME = "minty"
    root = tb.Window(themename=THEME)
    root.title("Liquidación de Horas Extra · TALCA")
    root.geometry("1120x700")
    root.minsize(1040, 660)

    FONT_TITLE = ("Segoe UI", 20, "bold")
    FONT_SUB = ("Segoe UI", 10)
    FONT_H2 = ("Segoe UI", 12, "bold")
    FONT_B = ("Segoe UI", 10, "bold")

    var_report = tk.StringVar(master=root, value="")
    var_emp = tk.StringVar(master=root, value=cfg.get("employees_excel_path", ""))
    var_tpl = tk.StringVar(master=root, value=cfg.get("template_excel_path", ""))

    var_use_holidays = tk.BooleanVar(master=root, value=False)
    var_auto_holidays = tk.BooleanVar(master=root, value=bool(cfg.get("auto_holidays_enabled", True)))

    status_var = tk.StringVar(master=root, value="Paso 1: elegí el reporte de VeoTime.")
    summary_range_var = tk.StringVar(master=root, value="—")
    summary_emps_var = tk.StringVar(master=root, value="—")
    summary_holidays_var = tk.StringVar(master=root, value="—")

    selected_holidays_manual = set()
    selected_holidays_auto = set()

    btn_generate = None

    def set_status(msg: str):
        status_var.set(msg)
        try:
            root.update_idletasks()
        except Exception:
            pass

    def month_name_es(m):
        names = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
        return names[m-1]

    def get_report_date_range_and_emps(report_path: str):
        try:
            df_raw = read_report_raw(report_path)
            header_row = find_header_row(df_raw)
            if header_row is not None:
                df = apply_header_row(df_raw, header_row)
            else:
                df = df_raw.copy()
                df.columns = [str(c).strip() for c in df.iloc[0].tolist()]
                df = df.iloc[1:].copy()

            cols_norm = [normalize_text(c) for c in df.columns]
            i_fecha = guess_col(cols_norm, ["fecha"])
            i_id = guess_col(cols_norm, ["dni", "documento", "legajo", "id"])

            if i_fecha is None:
                return None, None, None

            fechas = df.iloc[:, i_fecha].apply(parse_date).dropna()
            if fechas.empty:
                return None, None, None

            start = min(fechas)
            end = max(fechas)

            emps = None
            if i_id is not None:
                tmp = df.iloc[:, i_id].apply(clean_id)
                keys = tmp.apply(id_key_from_any)
                keys = keys[keys != ""]
                if not keys.empty:
                    emps = int(keys.nunique())

            return start, end, emps
        except Exception:
            return None, None, None

    def refresh_auto_holidays():
        selected_holidays_auto.clear()

        if not var_use_holidays.get():
            return
        if not var_auto_holidays.get():
            return

        rp = var_report.get().strip()
        if not rp or not os.path.exists(rp):
            return

        start, end, _ = get_report_date_range_and_emps(rp)
        if start is None or end is None:
            return

        try:
            auto_dates = compute_auto_holidays_for_range(start, end, cfg)
            for d in auto_dates:
                selected_holidays_auto.add(d)
        except Exception:
            pass

    def update_generate_state():
        rp = var_report.get().strip()
        emp = var_emp.get().strip()
        tpl = var_tpl.get().strip()

        ok = bool(
            rp and os.path.exists(rp) and
            emp and os.path.exists(emp) and
            tpl and os.path.exists(tpl)
        )

        if btn_generate is None:
            return

        btn_generate.configure(state=("normal" if ok else "disabled"))

    def update_summary():
        rp = var_report.get().strip()
        if rp and os.path.exists(rp):
            start, end, emps = get_report_date_range_and_emps(rp)
            summary_range_var.set(f"{start.strftime('%d/%m/%Y')} → {end.strftime('%d/%m/%Y')}" if start and end else "—")
            summary_emps_var.set(str(emps) if emps is not None else "—")
        else:
            summary_range_var.set("—")
            summary_emps_var.set("—")

        all_h = sorted(selected_holidays_auto | selected_holidays_manual) if var_use_holidays.get() else []
        summary_holidays_var.set(", ".join(d.strftime("%d/%m") for d in all_h) if all_h else "—")

        update_generate_state()

    # defaults locales
    local_emp = os.path.join(os.getcwd(), EMP_MASTER_DEFAULT_NAME)
    if not var_emp.get().strip() and os.path.exists(local_emp):
        var_emp.set(local_emp)
        cfg["employees_excel_path"] = local_emp

    local_tpl = os.path.join(os.getcwd(), TEMPLATE_DEFAULT_NAME)
    if not var_tpl.get().strip() and os.path.exists(local_tpl):
        var_tpl.set(local_tpl)
        cfg["template_excel_path"] = local_tpl

    # ==========================================================
    # LAYOUT
    # ==========================================================
    header = tb.Frame(root, padding=(20, 18))
    header.pack(fill="x")

    left_h = tb.Frame(header)
    left_h.pack(side="left", fill="x", expand=True)

    tb.Label(left_h, text="Liquidación de Horas Extra", font=FONT_TITLE).pack(anchor="w")
    tb.Label(left_h, text="Generación automática desde VeoTime → Excel RRHH", font=FONT_SUB, foreground="#666").pack(anchor="w", pady=(4, 0))

    right_h = tb.Frame(header)
    right_h.pack(side="right")

    LOGO_PATH = Path(__file__).resolve().parent / "assets" / "talca_logo.png"
    logo_img = None
    try:
        from PIL import Image, ImageTk
        img = Image.open(LOGO_PATH)
        img = img.resize((160, 46), Image.LANCZOS)
        logo_img = ImageTk.PhotoImage(img)
    except Exception:
        try:
            logo_img = tk.PhotoImage(file=str(LOGO_PATH))
        except Exception:
            logo_img = None

    if logo_img:
        lbl_logo = tb.Label(right_h, image=logo_img)
        lbl_logo.pack(side="right")
        lbl_logo.image = logo_img
    else:
        tb.Label(right_h, text="TALCA", bootstyle="secondary-inverse", padding=(10, 4)).pack(side="right")

    body = tb.Frame(root, padding=(20, 0, 20, 12))
    body.pack(fill="both", expand=True)

    body.columnconfigure(0, weight=3)
    body.columnconfigure(1, weight=2)
    body.rowconfigure(0, weight=1)

    left = tb.Frame(body)
    left.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
    right = tb.Frame(body)
    right.grid(row=0, column=1, sticky="nsew", padx=(14, 0))

    nb = tb.Notebook(left, bootstyle="primary")
    nb.pack(fill="both", expand=True)

    tab_files = tb.Frame(nb, padding=16)
    tab_holidays = tb.Frame(nb, padding=16)

    nb.add(tab_files, text="  1 · Archivos  ")
    nb.add(tab_holidays, text="  2 · Feriados  ")

    summary = tb.Labelframe(right, text="RESUMEN", padding=16, bootstyle="secondary")
    summary.pack(fill="x")

    def summary_row(parent, label, var):
        fr = tb.Frame(parent)
        fr.pack(fill="x", pady=8)
        tb.Label(fr, text=label, foreground="#666").pack(side="left")
        tb.Label(fr, textvariable=var, font=FONT_B).pack(side="right")

    summary_row(summary, "Rango detectado", summary_range_var)
    summary_row(summary, "Empleados (aprox.)", summary_emps_var)
    summary_row(summary, "Feriados (dd/mm)", summary_holidays_var)

    status_card = tb.Labelframe(right, text="ESTADO", padding=16, bootstyle="light")
    status_card.pack(fill="both", expand=True, pady=(14, 0))

    tb.Label(status_card, textvariable=status_var, wraplength=360, justify="left").pack(anchor="w")
    pb = tb.Progressbar(status_card, mode="indeterminate", bootstyle="success-striped")
    pb.pack(fill="x", pady=(14, 0))

    # ==========================================================
    # TAB 1: FILES
    # ==========================================================
    tb.Label(tab_files, text="Seleccioná los 3 archivos", font=FONT_H2).pack(anchor="w")

    def entry_file(parent, title, var, browse_cmd, hint=""):
        box = tb.Frame(parent)
        box.pack(fill="x", pady=10)

        top = tb.Frame(box)
        top.pack(fill="x")
        tb.Label(top, text=title, font=FONT_B).pack(side="left")
        if hint:
            tb.Label(top, text=hint, font=FONT_SUB, foreground="#666").pack(side="left", padx=(10, 0))

        row = tb.Frame(box)
        row.pack(fill="x", pady=(6, 0))
        e = tb.Entry(row, textvariable=var)
        e.pack(side="left", fill="x", expand=True)
        tb.Button(row, text="Buscar…", command=browse_cmd, bootstyle="secondary", width=12).pack(side="left", padx=(10, 0))
        return e

    def pick_report():
        p = filedialog.askopenfilename(
            title="Seleccionar reporte de VeoTime",
            filetypes=[("Excel", "*.xls *.xlsx")]
        )
        if p:
            var_report.set(p)
            set_status("Reporte seleccionado. Paso 2: revisá feriados (si aplica).")
            refresh_auto_holidays()
            update_summary()
            try:
                render_calendar()
                render_holiday_list()
            except Exception:
                pass

    def pick_emp():
        p = filedialog.askopenfilename(
            title="Seleccionar datos empleados.xlsx",
            filetypes=[("Excel", "*.xlsx *.xls")]
        )
        if p:
            var_emp.set(p)
            cfg["employees_excel_path"] = p
            save_config(cfg)
            update_summary()

    def pick_tpl():
        p = filedialog.askopenfilename(
            title="Seleccionar formatosugerido.xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if p:
            var_tpl.set(p)
            cfg["template_excel_path"] = p
            save_config(cfg)
            update_summary()

    entry_file(tab_files, "Reporte VeoTime (.xls/.xlsx)", var_report, pick_report, hint="Obligatorio")
    entry_file(tab_files, "Datos empleados.xlsx", var_emp, pick_emp, hint="Obligatorio")
    entry_file(tab_files, "Plantilla formatosugerido.xlsx", var_tpl, pick_tpl, hint="Obligatorio")

    # ==========================================================
    # TAB 2: HOLIDAYS (INLINE CALENDAR)
    # ==========================================================
    tb.Label(tab_holidays, text="Feriados de la semana", font=FONT_H2).pack(anchor="w")

    hol_top = tb.Frame(tab_holidays)
    hol_top.pack(fill="x", pady=(12, 10))

    cal_year = tk.IntVar(master=root, value=datetime.now().year)
    cal_month = tk.IntVar(master=root, value=datetime.now().month)

    def get_report_range():
        rp = var_report.get().strip()
        if rp and os.path.exists(rp):
            start, end, _ = get_report_date_range_and_emps(rp)
            return start, end
        return None, None

    def toggle_manual_date(dte: date):
        if not var_use_holidays.get():
            return
        if dte in selected_holidays_manual:
            selected_holidays_manual.remove(dte)
        else:
            selected_holidays_manual.add(dte)
        render_holiday_list()
        update_summary()
        render_calendar()

    def prev_month():
        y, m = cal_year.get(), cal_month.get()
        if m == 1:
            cal_year.set(y - 1)
            cal_month.set(12)
        else:
            cal_month.set(m - 1)
        render_calendar()

    def next_month():
        y, m = cal_year.get(), cal_month.get()
        if m == 12:
            cal_year.set(y + 1)
            cal_month.set(1)
        else:
            cal_month.set(m + 1)
        render_calendar()

    def goto_today():
        today = date.today()
        cal_year.set(today.year)
        cal_month.set(today.month)
        render_calendar()

    def on_toggle_holidays():
        sync_holiday_state()
        refresh_auto_holidays()
        render_holiday_list()
        update_summary()
        render_calendar()

    def on_toggle_auto():
        cfg["auto_holidays_enabled"] = bool(var_auto_holidays.get())
        save_config(cfg)
        refresh_auto_holidays()
        render_holiday_list()
        update_summary()
        render_calendar()

    chk_use = tb.Checkbutton(
        hol_top,
        text="Hubo feriados",
        variable=var_use_holidays,
        command=on_toggle_holidays,
        bootstyle="round-toggle"
    )
    chk_use.pack(side="left")

    chk_auto = tb.Checkbutton(
        hol_top,
        text="Detectar automáticamente (Argentina · Mendoza)",
        variable=var_auto_holidays,
        command=on_toggle_auto,
        bootstyle="round-toggle"
    )
    chk_auto.pack(side="left", padx=(14, 0))

    wrap = tb.Frame(tab_holidays)
    wrap.pack(fill="both", expand=True)

    wrap.columnconfigure(0, weight=3)
    wrap.columnconfigure(1, weight=2)
    wrap.rowconfigure(0, weight=1)

    cal_card = tb.Labelframe(wrap, text="CALENDARIO", padding=12, bootstyle="light")
    cal_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

    list_card = tb.Labelframe(wrap, text="FERIADOS APLICADOS", padding=12, bootstyle="secondary")
    list_card.grid(row=0, column=1, sticky="nsew", padx=(12, 0))

    cal_header = tb.Frame(cal_card)
    cal_header.pack(fill="x", pady=(0, 10))

    tb.Button(cal_header, text="◀", command=prev_month, bootstyle="secondary", width=4).pack(side="left")
    lbl_month = tb.Label(cal_header, text="", font=("Segoe UI", 11, "bold"))
    lbl_month.pack(side="left", padx=10)
    tb.Button(cal_header, text="▶", command=next_month, bootstyle="secondary", width=4).pack(side="left")
    tb.Button(cal_header, text="Hoy", command=goto_today, bootstyle="info", width=8).pack(side="right")

    legend = tb.Frame(cal_card)
    legend.pack(fill="x", pady=(0, 8))
    tb.Label(legend, text="MANUAL", bootstyle="primary-inverse", padding=(8, 3)).pack(side="left")
    tb.Label(legend, text="AUTO", bootstyle="info-inverse", padding=(8, 3)).pack(side="left", padx=(8, 0))
    tb.Label(legend, text="Click para agregar/quitar", foreground="#666").pack(side="left", padx=(10, 0))

    range_lbl = tb.Label(cal_card, text="", foreground="#666")
    range_lbl.pack(anchor="w", pady=(0, 8))

    grid = tb.Frame(cal_card)
    grid.pack(fill="both", expand=True)

    days_row = tb.Frame(grid)
    days_row.pack(fill="x")
    for dn in ["Lu", "Ma", "Mi", "Ju", "Vi", "Sa", "Do"]:
        tb.Label(days_row, text=dn, width=5, anchor="center", foreground="#666").pack(side="left", padx=2)

    cells = tb.Frame(grid)
    cells.pack(fill="both", expand=True, pady=(6, 0))

    def render_calendar():
        for w in cells.winfo_children():
            w.destroy()

        y, m = cal_year.get(), cal_month.get()
        lbl_month.config(text=f"{month_name_es(m)} {y}")

        start, end = get_report_range()
        if start and end:
            range_lbl.config(text=f"Rango del reporte: {start.strftime('%d/%m/%Y')} → {end.strftime('%d/%m/%Y')}")
        else:
            range_lbl.config(text="(Tip: al elegir el reporte, se limita el rango del calendario)")

        enabled = var_use_holidays.get()
        weeks_ = calendar.monthcalendar(y, m)

        for week in weeks_:
            row = tb.Frame(cells)
            row.pack(fill="x", pady=2)

            for day in week:
                if day == 0:
                    tb.Label(row, text="", width=5).pack(side="left", padx=2)
                    continue

                dte = date(y, m, day)

                in_range = True
                if start and end:
                    in_range = (start <= dte <= end)

                is_manual = dte in selected_holidays_manual
                is_auto = dte in selected_holidays_auto

                if is_manual:
                    style = "primary"
                elif is_auto:
                    style = "info-outline"
                else:
                    style = "light"

                state = "normal"
                if (not enabled) or (not in_range):
                    state = "disabled"

                b = tb.Button(
                    row,
                    text=str(day),
                    width=5,
                    command=lambda dd=dte: toggle_manual_date(dd),
                    bootstyle=style
                )
                b.pack(side="left", padx=2)
                b.configure(state=state)

    def render_holiday_list():
        lb.delete(0, "end")
        all_dates = sorted(selected_holidays_auto | selected_holidays_manual)
        if not all_dates:
            lb.insert("end", "—")
            return
        for dte in all_dates:
            tags = []
            if dte in selected_holidays_auto:
                tags.append("AUTO")
            if dte in selected_holidays_manual:
                tags.append("MANUAL")
            suffix = f"   ·   {' + '.join(tags)}" if tags else ""
            lb.insert("end", dte.strftime("%d/%m/%Y") + suffix)

    def remove_selected_holiday():
        sel = lb.curselection()
        if not sel:
            return
        all_dates = sorted(selected_holidays_auto | selected_holidays_manual)
        if not all_dates:
            return
        dte = all_dates[sel[0]]

        if dte in selected_holidays_manual:
            selected_holidays_manual.remove(dte)
        elif dte in selected_holidays_auto:
            selected_holidays_auto.remove(dte)

        render_holiday_list()
        update_summary()
        render_calendar()

    def clear_all_holidays():
        selected_holidays_manual.clear()
        selected_holidays_auto.clear()
        render_holiday_list()
        update_summary()
        render_calendar()

    tb.Label(list_card, text="Seleccionados (auto + manual):", font=FONT_B).pack(anchor="w")
    lb = tk.Listbox(list_card, height=12)
    lb.pack(fill="both", expand=True, pady=(10, 10))

    actions = tb.Frame(list_card)
    actions.pack(fill="x")

    btn_remove_h = tb.Button(actions, text="Quitar seleccionado", command=remove_selected_holiday, bootstyle="secondary")
    btn_clear_h = tb.Button(actions, text="Limpiar todo", command=clear_all_holidays, bootstyle="secondary")
    btn_remove_h.pack(side="left")
    btn_clear_h.pack(side="right")

    def sync_holiday_state():
        enabled = var_use_holidays.get()
        st = "normal" if enabled else "disabled"

        try:
            chk_auto.configure(state=st)
        except Exception:
            pass

        try:
            btn_remove_h.configure(state=st)
            btn_clear_h.configure(state=st)
        except Exception:
            pass

        try:
            lb.configure(state=st)
        except Exception:
            pass

        if not enabled:
            selected_holidays_manual.clear()
            selected_holidays_auto.clear()

        render_holiday_list()
        update_summary()
        render_calendar()

    # ==========================================================
    # FOOTER: GENERATE
    # ==========================================================
    footer = tb.Frame(root, padding=(20, 12))
    footer.pack(fill="x")

    tb.Label(footer, text="Podés guardar un archivo nuevo o elegir uno existente para agregar empleados.", foreground="#666").pack(side="left")

    def generate():
        rp = var_report.get().strip()
        emp = var_emp.get().strip()
        tpl = var_tpl.get().strip()

        default_out = os.path.join(
            os.path.dirname(rp) if rp else os.getcwd(),
            "Liquidacion_Horas_Extra.xlsx"
        )

        out_path = filedialog.asksaveasfilename(
            title="Guardar / Actualizar liquidación",
            defaultextension=".xlsx",
            initialfile=os.path.basename(default_out),
            filetypes=[("Excel", "*.xlsx")]
        )
        if not out_path:
            return

        existed = os.path.exists(out_path)

        btn_generate.configure(state="disabled")
        pb.start(12)

        try:
            set_status("Procesando… leyendo empleados y reporte VeoTime…")

            if var_use_holidays.get():
                refresh_auto_holidays()
                all_h = sorted(selected_holidays_auto | selected_holidays_manual)
                cfg["holidays"] = [d.strftime("%Y-%m-%d") for d in all_h]
            else:
                cfg["holidays"] = []

            cfg["auto_holidays_enabled"] = bool(var_auto_holidays.get())
            save_config(cfg)

            emp_master = load_employee_master(emp)
            daily = read_veotime_to_daily(rp, emp_master, cfg)
            weeks = compute_overtime_from_daily(daily, cfg)

            set_status("Escribiendo Excel… (si ya existe, agrega empleados abajo)")
            update_or_build_output_workbook(weeks, out_path, tpl, cfg)

            set_status("Listo ✅ Excel generado/actualizado correctamente.")
            messagebox.showinfo(
                "Listo ✅",
                ("Se ACTUALIZÓ el Excel (se agregaron empleados debajo de los existentes):\n"
                 if existed else "Se generó el Excel:\n") + f"{out_path}"
            )

        except PermissionError:
            set_status("Error ❌ El archivo está abierto.")
            messagebox.showerror(
                "Archivo en uso",
                "Cerrá el Excel (está abierto) y volvé a intentar.\n"
                "Windows no deja guardar si el archivo está en uso."
            )

        except Exception as e:
            tbtxt = traceback.format_exc()
            with open("debug_error.txt", "w", encoding="utf-8") as f:
                f.write(tbtxt)
            set_status("Error ❌ Revisá debug_error.txt")
            messagebox.showerror("Error", f"{e}\n\nSe guardó el detalle en debug_error.txt")

        finally:
            pb.stop()
            update_generate_state()

    btn_generate = tb.Button(footer, text="Generar / Actualizar", command=generate, bootstyle="success", width=22)
    btn_generate.pack(side="right")

    # ==========================================================
    # Bindings + init
    # ==========================================================
    def on_any_change(*_):
        update_summary()

    var_report.trace_add("write", on_any_change)
    var_emp.trace_add("write", on_any_change)
    var_tpl.trace_add("write", on_any_change)

    refresh_auto_holidays()
    render_holiday_list()
    render_calendar()
    sync_holiday_state()
    update_summary()

    root.mainloop()
