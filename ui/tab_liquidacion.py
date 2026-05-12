import os, traceback, tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as tb
from services.liquidacion_service import run_liquidacion
from config.settings import save_config

def render_tab_liquidacion(nb, st):
    tab = tb.Frame(nb, padding=16)
    nb.add(tab, text="  1 · Archivos   ")
    tb.Label(tab, text="Seleccioná los 3 archivos", font=("Segoe UI", 12, "bold")).pack(anchor="w")

    def entry_file(parent, title, var, browse_cmd, hint=""):
        box = tb.Frame(parent); box.pack(fill="x", pady=10)
        top = tb.Frame(box); top.pack(fill="x")
        tb.Label(top, text=title, font=("Segoe UI", 10, "bold")).pack(side="left")
        if hint: tb.Label(top, text=hint, font=("Segoe UI", 9), foreground="#666").pack(side="left", padx=(10,0))
        row = tb.Frame(box); row.pack(fill="x", pady=(6,0))
        tb.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True)
        tb.Button(row, text="Buscar…", command=browse_cmd, bootstyle="secondary", width=12).pack(side="left", padx=(10,0))

    def pick_report():
        p = filedialog.askopenfilename(title="Reporte VeoTime", filetypes=[("Excel", "*.xls *.xlsx")])
        if p:
            st.var_report.set(p)
            st.status_var.set("Reporte seleccionado. Paso 2: revisá feriados.")
            from ui.tab_feriados import refresh_auto_holidays
            refresh_auto_holidays(st)
            from ui.app import update_summary
            update_summary(st)

    def pick_emp():
        p = filedialog.askopenfilename(title="Datos empleados.xlsx", filetypes=[("Excel", "*.xlsx *.xls")])
        if p:
            st.var_emp.set(p)
            st.cfg["employees_excel_path"] = p
            save_config(st.cfg)
            from ui.app import update_summary
            update_summary(st)

    def pick_tpl():
        p = filedialog.askopenfilename(title="Plantilla formatosugerido.xlsx", filetypes=[("Excel", "*.xlsx")])
        if p:
            st.var_tpl.set(p)
            st.cfg["template_excel_path"] = p
            save_config(st.cfg)
            from ui.app import update_summary
            update_summary(st)

    entry_file(tab, "Reporte VeoTime (.xls/.xlsx)", st.var_report, pick_report, hint="Obligatorio")
    entry_file(tab, "Datos empleados.xlsx", st.var_emp, pick_emp, hint="Obligatorio")
    entry_file(tab, "Plantilla formatosugerido.xlsx", st.var_tpl, pick_tpl, hint="Obligatorio")

def generate(st):
    rp = st.var_report.get().strip()
    emp = st.var_emp.get().strip()
    tpl = st.var_tpl.get().strip()

    if not all([rp, emp, tpl, os.path.exists(rp), os.path.exists(emp), os.path.exists(tpl)]):
        return messagebox.showerror("Faltan archivos", "Verificá los 3 archivos.")

    default_out = os.path.join(os.path.dirname(rp), "Liquidacion_Horas_Extra.xlsx")
    out_path = filedialog.asksaveasfilename(title="Guardar / Actualizar liquidación", defaultextension=".xlsx", initialfile=os.path.basename(default_out), filetypes=[("Excel", "*.xlsx")])
    if not out_path: return
    existed = os.path.exists(out_path)

    st.btn_generate.configure(state="disabled")
    st.pb.start(12)
    try:
        st.status_var.set("Procesando… leyendo empleados y reporte VeoTime…")
        from ui.tab_feriados import sync_holiday_state
        sync_holiday_state(st)
        holidays_list = [d.strftime("%Y-%m-%d") for d in (st.selected_holidays_auto | st.selected_holidays_manual)] if st.var_use_holidays.get() else []
        run_liquidacion(rp, emp, tpl, out_path, holidays_list, st.cfg)
        st.status_var.set("Listo ✅ Excel generado/actualizado correctamente.")
        messagebox.showinfo("Listo ✅", ("Se ACTUALIZÓ el Excel (se agregaron empleados debajo de los existentes):\n" if existed else "Se generó el Excel:\n") + f"{out_path}")
    except PermissionError:
        st.status_var.set("Error ❌ El archivo está abierto.")
        messagebox.showerror("Archivo en uso", "Cerrá el Excel (está abierto) y volvé a intentar.\nWindows no deja guardar si el archivo está en uso.")
    except Exception as e:
        with open("debug_error.txt", "w", encoding="utf-8") as f: f.write(traceback.format_exc())
        st.status_var.set("Error ❌ Revisá debug_error.txt")
        messagebox.showerror("Error", f"{e}\n\nSe guardó el detalle en debug_error.txt")
    finally:
        st.pb.stop()
        from ui.app import update_summary
        update_summary(st)