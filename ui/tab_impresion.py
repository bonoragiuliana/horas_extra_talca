import os, traceback, tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as tb
from services.rrhh_print_service import build_rrhh_print_workbook

def render_tab_impresion(nb, st):
    tab = tb.Frame(nb, padding=16)
    nb.add(tab, text="  3 · Impresión   ")
    tb.Label(tab, text="Generar planilla de impresión (RRHH)", font=("Segoe UI", 12, "bold")).pack(anchor="w")
    tb.Label(tab, text="Unifica OESTE + CONSULTORA en un solo Excel usando la plantilla de impresión.", foreground="#666").pack(anchor="w", pady=(6,0))

    def entry_file(parent, title, var, browse_cmd, hint=""):
        box = tb.Frame(parent); box.pack(fill="x", pady=10)
        top = tb.Frame(box); top.pack(fill="x")
        tb.Label(top, text=title, font=("Segoe UI", 10, "bold")).pack(side="left")
        if hint: tb.Label(top, text=hint, font=("Segoe UI", 9), foreground="#666").pack(side="left", padx=(10,0))
        row = tb.Frame(box); row.pack(fill="x", pady=(6,0))
        tb.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True)
        tb.Button(row, text="Buscar…", command=browse_cmd, bootstyle="secondary", width=12).pack(side="left", padx=(10,0))

    def pick_oeste():
        p = filedialog.askopenfilename(title="Planilla OESTE (generada)", filetypes=[("Excel", "*.xlsx *.xls")])
        if p: st.var_oeste.set(p)
        from ui.app import update_summary
        update_summary(st)

    def pick_consultora():
        p = filedialog.askopenfilename(title="Planilla CONSULTORA (generada)", filetypes=[("Excel", "*.xlsx *.xls")])
        if p: st.var_consultora.set(p)
        from ui.app import update_summary
        update_summary(st)

    def pick_tpl_print():
        p = filedialog.askopenfilename(title="Plantilla impresión (.xlsx)", filetypes=[("Excel", "*.xlsx")])
        if p: st.var_tpl_print.set(p)
        from config.settings import save_config
        st.cfg["print_template_excel_path"] = p
        save_config(st.cfg)
        from ui.app import update_summary
        update_summary(st)

    entry_file(tab, "Planilla OESTE (generada)", st.var_oeste, pick_oeste, hint="Obligatorio")
    entry_file(tab, "Planilla CONSULTORA (generada)", st.var_consultora, pick_consultora, hint="Obligatorio")
    entry_file(tab, "Plantilla impresión (.xlsx)", st.var_tpl_print, pick_tpl_print, hint="Obligatorio")

    print_actions = tb.Frame(tab); print_actions.pack(fill="x", pady=(16,0))
    
    def on_print_click():
        generate_rrhh_print(st)

    st.btn_print = tb.Button(print_actions, text="Generar impresión", command=on_print_click, bootstyle="info", width=22)
    st.btn_print.pack(side="right")
    from ui.app import update_summary
    update_summary(st)

def generate_rrhh_print(st):
    path_oeste = st.var_oeste.get().strip()
    path_consultora = st.var_consultora.get().strip()
    tpl_print = st.var_tpl_print.get().strip()

    if not (path_oeste and os.path.exists(path_oeste)):
        return messagebox.showerror("Falta archivo", "Seleccioná la planilla de OESTE.")
    if not (path_consultora and os.path.exists(path_consultora)):
        return messagebox.showerror("Falta archivo", "Seleccioná la planilla de CONSULTORA.")
    if not (tpl_print and os.path.exists(tpl_print)):
        return messagebox.showerror("Falta plantilla", "Seleccioná la plantilla de impresión.")

    out_path = filedialog.asksaveasfilename(title="Guardar planilla de impresión", defaultextension=".xlsx", initialfile="Impresion_RRHH.xlsx", filetypes=[("Excel", "*.xlsx")])
    if not out_path: return

    st.btn_print.configure(state="disabled")
    st.pb.start(12)
    try:
        st.status_var.set("Generando impresión RRHH… (Oeste + Consultora → una sola planilla)")
        build_rrhh_print_workbook(path_oeste=path_oeste, path_consultora=path_consultora, template_print_path=tpl_print, out_path=out_path)
        st.status_var.set("Listo ✅ Impresión RRHH generada.")
        messagebox.showinfo("Listo ✅", f"Se generó el archivo de impresión:\n{out_path}")
    except PermissionError:
        st.status_var.set("Error ❌ El archivo está abierto.")
        messagebox.showerror("Archivo en uso", "Cerrá el Excel (está abierto) y volvé a intentar.\nWindows no deja guardar si el archivo está en uso.")
    except Exception as e:
        with open("debug_error_rrhh_print.txt", "w", encoding="utf-8") as f: f.write(traceback.format_exc())
        st.status_var.set("Error ❌ Revisá debug_error_rrhh_print.txt")
        messagebox.showerror("Error", f"{e}\n\nSe guardó el detalle en debug_error_rrhh_print.txt")
    finally:
        st.pb.stop()
        from ui.app import update_summary
        update_summary(st)