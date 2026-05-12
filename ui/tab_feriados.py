import calendar, os, tkinter as tk
from datetime import date, datetime
import ttkbootstrap as tb
from config.settings import save_config
from services.holidays_service import compute_auto_holidays_for_range

def render_tab_feriados(nb, st):
    tab = tb.Frame(nb, padding=16)
    nb.add(tab, text="  2 · Feriados   ")
    tb.Label(tab, text="Feriados de la semana", font=("Segoe UI", 12, "bold")).pack(anchor="w")

    hol_top = tb.Frame(tab); hol_top.pack(fill="x", pady=(12, 10))

    st.chk_use = tb.Checkbutton(hol_top, text="Hubo feriados", variable=st.var_use_holidays, bootstyle="round-toggle")
    st.chk_use.pack(side="left")
    st.chk_auto = tb.Checkbutton(hol_top, text="Detectar automáticamente (Argentina · Mendoza)", variable=st.var_auto_holidays, bootstyle="round-toggle")
    st.chk_auto.pack(side="left", padx=(14, 0))

    wrap = tb.Frame(tab); wrap.pack(fill="both", expand=True)
    wrap.columnconfigure(0, weight=3); wrap.columnconfigure(1, weight=2); wrap.rowconfigure(0, weight=1)

    cal_card = tb.Labelframe(wrap, text="CALENDARIO", padding=12, bootstyle="light")
    cal_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

    list_card = tb.Labelframe(wrap, text="FERIADOS APLICADOS", padding=12, bootstyle="secondary")
    list_card.grid(row=0, column=1, sticky="nsew", padx=(12, 0))

    # Calendar Header
    cal_header = tb.Frame(cal_card); cal_header.pack(fill="x", pady=(0, 10))
    tb.Button(cal_header, text="◀", command=lambda: prev_month(st), bootstyle="secondary", width=4).pack(side="left")
    st.lbl_month = tb.Label(cal_header, text="", font=("Segoe UI", 11, "bold"))
    st.lbl_month.pack(side="left", padx=10)
    tb.Button(cal_header, text="▶", command=lambda: next_month(st), bootstyle="secondary", width=4).pack(side="left")
    tb.Button(cal_header, text="Hoy", command=lambda: goto_today(st), bootstyle="info", width=8).pack(side="right")

    legend = tb.Frame(cal_card); legend.pack(fill="x", pady=(0, 8))
    tb.Label(legend, text="MANUAL", bootstyle="primary-inverse", padding=(8, 3)).pack(side="left")
    tb.Label(legend, text="AUTO", bootstyle="info-inverse", padding=(8, 3)).pack(side="left", padx=(8, 0))
    tb.Label(legend, text="Click para agregar/quitar", foreground="#666").pack(side="left", padx=(10, 0))

    st.range_lbl = tb.Label(cal_card, text="", foreground="#666")
    st.range_lbl.pack(anchor="w", pady=(0, 8))

    days_row = tb.Frame(cal_card); days_row.pack(fill="x")
    for dn in ["Lu", "Ma", "Mi", "Ju", "Vi", "Sa", "Do"]:
        tb.Label(days_row, text=dn, width=5, anchor="center", foreground="#666").pack(side="left", padx=2)

    st.cal_grid = tb.Frame(cal_card); st.cal_grid.pack(fill="both", expand=True, pady=(6, 0))

    # Listbox & Actions
    tb.Label(list_card, text="Seleccionados (auto + manual):", font=("Segoe UI", 10, "bold")).pack(anchor="w")
    st.holiday_lb = tk.Listbox(list_card, height=12); st.holiday_lb.pack(fill="both", expand=True, pady=(10, 10))

    actions = tb.Frame(list_card); actions.pack(fill="x")
    tb.Button(actions, text="Quitar seleccionado", command=lambda: remove_selected_holiday(st), bootstyle="secondary").pack(side="left")
    tb.Button(actions, text="Limpiar todo", command=lambda: clear_all_holidays(st), bootstyle="secondary").pack(side="right")

def month_name_es(m):
    return ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"][m-1]

def get_report_range(st):
    rp = st.var_report.get().strip()
    if rp and os.path.exists(rp):
        from readers.veotime_reader import read_report_raw, find_header_row, apply_header_row
        from utils.text_utils import normalize_text
        from utils.date_utils import parse_date
        try:
            df_raw = read_report_raw(rp)
            hr = find_header_row(df_raw)
            df = apply_header_row(df_raw, hr) if hr else df_raw.copy()
            if hr is None:
                df.columns = [str(c).strip() for c in df.iloc[0].tolist()]; df = df.iloc[1:].copy()
            cols_norm = [normalize_text(c) for c in df.columns]
            i_fecha = None
            for i, c in enumerate(cols_norm):
                if "fecha" in c: i_fecha = i; break
            if i_fecha is not None:
                fechas = df.iloc[:, i_fecha].apply(parse_date).dropna()
                if not fechas.empty: return min(fechas), max(fechas)
        except: pass
    return None, None

def refresh_auto_holidays(st):
    st.selected_holidays_auto.clear()
    if not st.var_use_holidays.get() or not st.var_auto_holidays.get(): return
    start, end = get_report_range(st)
    if start and end:
        try:
            auto_dates = compute_auto_holidays_for_range(start, end, st.cfg)
            for d in auto_dates: st.selected_holidays_auto.add(d)
        except: pass

def sync_holiday_state(st):
    enabled = st.var_use_holidays.get()
    state = "normal" if enabled else "disabled"
    try:
        st.chk_auto.configure(state=state)
        st.holiday_lb.configure(state=state)
    except: pass
    if not enabled:
        st.selected_holidays_manual.clear()
        st.selected_holidays_auto.clear()
    render_holiday_list(st)
    from ui.app import update_summary
    update_summary(st)
    render_calendar(st)

def toggle_manual_date(st, dte: date):
    if not st.var_use_holidays.get(): return
    if dte in st.selected_holidays_manual: st.selected_holidays_manual.remove(dte)
    else: st.selected_holidays_manual.add(dte)
    render_holiday_list(st)
    render_calendar(st)
    from ui.app import update_summary
    update_summary(st)

def prev_month(st):
    y, m = st.cal_year.get(), st.cal_month.get()
    st.cal_year.set(y-1 if m==1 else y)
    st.cal_month.set(12 if m==1 else m-1)
    render_calendar(st)

def next_month(st):
    y, m = st.cal_year.get(), st.cal_month.get()
    st.cal_year.set(y+1 if m==12 else y)
    st.cal_month.set(1 if m==12 else m+1)
    render_calendar(st)

def goto_today(st):
    st.cal_year.set(date.today().year)
    st.cal_month.set(date.today().month)
    render_calendar(st)

def render_calendar(st):
    for w in st.cal_grid.winfo_children(): w.destroy()
    y, m = st.cal_year.get(), st.cal_month.get()
    st.lbl_month.config(text=f"{month_name_es(m)} {y}")

    start, end = get_report_range(st)
    if start and end: st.range_lbl.config(text=f"Rango del reporte: {start.strftime('%d/%m/%Y')} → {end.strftime('%d/%m/%Y')}")
    else: st.range_lbl.config(text="(Tip: al elegir el reporte, se limita el rango del calendario)")

    enabled = st.var_use_holidays.get()
    for week in calendar.monthcalendar(y, m):
        row = tb.Frame(st.cal_grid); row.pack(fill="x", pady=2)
        for day in week:
            if day == 0:
                tb.Label(row, text="", width=5).pack(side="left", padx=2)
                continue
            dte = date(y, m, day)
            in_range = True if not start or not end else (start <= dte <= end)
            is_manual = dte in st.selected_holidays_manual
            is_auto = dte in st.selected_holidays_auto

            style = "primary" if is_manual else ("info-outline" if is_auto else "light")
            state = "normal" if (enabled and in_range) else "disabled"

            b = tb.Button(row, text=str(day), width=5, bootstyle=style, state=state)
            b.pack(side="left", padx=2)
            b.configure(command=lambda dd=dte: toggle_manual_date(st, dd))

def render_holiday_list(st):
    st.holiday_lb.delete(0, "end")
    all_dates = sorted(st.selected_holidays_auto | st.selected_holidays_manual)
    if not all_dates:
        st.holiday_lb.insert("end", "—"); return
    for dte in all_dates:
        tags = []
        if dte in st.selected_holidays_auto: tags.append("AUTO")
        if dte in st.selected_holidays_manual: tags.append("MANUAL")
        suffix = f"   ·   {' + '.join(tags)}" if tags else ""
        st.holiday_lb.insert("end", f"{dte.strftime('%d/%m/%Y')}{suffix}")

def remove_selected_holiday(st):
    sel = st.holiday_lb.curselection()
    if not sel: return
    all_dates = sorted(st.selected_holidays_auto | st.selected_holidays_manual)
    if not all_dates: return
    dte = all_dates[sel[0]]
    if dte in st.selected_holidays_manual: st.selected_holidays_manual.remove(dte)
    elif dte in st.selected_holidays_auto: st.selected_holidays_auto.remove(dte)
    render_holiday_list(st)
    render_calendar(st)
    from ui.app import update_summary
    update_summary(st)

def clear_all_holidays(st):
    st.selected_holidays_manual.clear()
    st.selected_holidays_auto.clear()
    render_holiday_list(st)
    render_calendar(st)
    from ui.app import update_summary
    update_summary(st)