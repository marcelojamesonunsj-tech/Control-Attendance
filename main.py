from __future__ import annotations

import io
import pandas as pd
import streamlit as st


REQUIRED_COLS = ["Nombre", "Marc.", "Estado", "NvoEstado"]


# =========================
# UI
# =========================
def inject_css() -> None:
    st.markdown(
        """
        <style>
        .block-container { max-width: 1320px; padding-top: 1rem; padding-bottom: 1.4rem; }
        header, footer {visibility: hidden;}
        div[data-testid="stToolbar"] {visibility: hidden; height: 0px;}

        .kpi {
            border: 1px solid rgba(255,255,255,0.10);
            border-radius: 18px;
            padding: 14px 14px;
            background: rgba(255,255,255,0.03);
        }
        .kpi .label {opacity:.78; font-size:.92rem;}
        .kpi .value {font-size:1.65rem; font-weight:900; line-height:1.1;}
        .kpi .sub {opacity:.70; font-size:.86rem; margin-top:.18rem;}

        .pill {display:inline-block; padding:6px 10px; border-radius:999px;
               border:1px solid rgba(255,255,255,.12); background:rgba(255,255,255,.03);
               font-size:.85rem; opacity:.92;}

        .hr {height:1px; background: rgba(255,255,255,0.08); margin: 0.9rem 0 1.0rem 0;}
        </style>
        """,
        unsafe_allow_html=True,
    )


# =========================
# Helpers
# =========================
def minutes_to_hhmm(mins: int) -> str:
    mins = int(mins) if mins is not None else 0
    h = mins // 60
    m = mins % 60
    return f"{h:02d}:{m:02d}"


def delta_short(mins: int) -> str:
    mins = int(mins)
    sign = "+" if mins > 0 else "-" if mins < 0 else ""
    mins_abs = abs(mins)
    h, m = mins_abs // 60, mins_abs % 60
    if h == 0:
        return f"{sign}{m:02d}m" if sign else "0m"
    return f"{sign}{h}h {m:02d}m" if sign else f"{h}h {m:02d}m"


def read_excel_auto(file) -> pd.DataFrame:
    try:
        return pd.read_excel(file)
    except Exception:
        file.seek(0)
        try:
            return pd.read_excel(file, engine="openpyxl")
        except Exception:
            file.seek(0)
            return pd.read_excel(file, engine="xlrd")


def validate_format(df: pd.DataFrame) -> None:
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Formato incorrecto del reloj. Faltan columnas: {', '.join(missing)}")


def parse_and_clean(df: pd.DataFrame) -> pd.DataFrame:
    """
    Tu Excel fijo:
    - Nombre = DNI
    - Marc.  = Apellido, Nombre
    - Estado = dd/mm/yyyy HH:MM
    - NvoEstado = puede ser cualquier cosa (no confiamos)
    """
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    df["DNI"] = df["Nombre"].astype(str).str.replace(r"\D", "", regex=True).str.strip()
    df["Empleado"] = df["Marc."].astype(str).str.strip()

    df["FechaHora"] = pd.to_datetime(df["Estado"], errors="coerce", dayfirst=True)
    df = df.dropna(subset=["FechaHora"])

    df["Fecha"] = df["FechaHora"].dt.date
    df = df.sort_values(["Empleado", "DNI", "FechaHora"]).reset_index(drop=True)
    return df


def init_profiles(raw: pd.DataFrame) -> pd.DataFrame:
    base = raw[["DNI", "Empleado"]].drop_duplicates().sort_values(["Empleado", "DNI"]).reset_index(drop=True)
    if "profiles" not in st.session_state:
        p = base.copy()
        p["Tipo"] = "NO Docente"
        st.session_state["profiles"] = p
        return p

    p = st.session_state["profiles"].copy()
    merged = base.merge(p, on=["DNI", "Empleado"], how="left")
    merged["Tipo"] = merged["Tipo"].fillna("NO Docente")
    merged = merged[["DNI", "Empleado", "Tipo"]]
    st.session_state["profiles"] = merged
    return merged


def apply_profiles(raw: pd.DataFrame, profiles: pd.DataFrame) -> pd.DataFrame:
    m = raw.merge(profiles, on=["DNI", "Empleado"], how="left")
    m["Tipo"] = m["Tipo"].fillna("NO Docente")
    return m


def pair_alternating(times: list[pd.Timestamp]) -> tuple[int, int]:
    """
    Empareja alternando por tiempo: (t0,t1), (t2,t3)...
    Devuelve (minutos_por_tramos, pares)
    """
    times = [t for t in times if pd.notna(t)]
    times.sort()
    total = 0
    pairs = 0
    for i in range(0, len(times) - 1, 2):
        a, b = times[i], times[i + 1]
        if b >= a:
            total += int((b - a).total_seconds() // 60)
            pairs += 1
    return total, pairs


def calc_daily(raw: pd.DataFrame, expected_nodoc: int) -> pd.DataFrame:
    """
    - Docente: worked = tramos estimados (alternancia)
    - NO Docente: worked = (última - primera) "corrido"
    """
    rows = []
    for (dni, emp, tipo, day), g in raw.groupby(["DNI", "Empleado", "Tipo", "Fecha"], dropna=False):
        g = g.sort_values("FechaHora")
        times = g["FechaHora"].tolist()
        marc = int(g.shape[0])

        first = times[0] if times else pd.NaT
        last = times[-1] if times else pd.NaT

        worked_pairs, pairs = pair_alternating(times)

        span = 0
        if pd.notna(first) and pd.notna(last) and last >= first:
            span = int((last - first).total_seconds() // 60)

        incompleto = (marc < 2) or (pairs == 0 and tipo == "Docente") or (marc < 2 and tipo == "NO Docente")
        cortes = (pairs >= 2)

        if tipo == "Docente":
            worked = worked_pairs
            expected = 0
            saldo = 0
            cumple = ""
        else:
            worked = span if (pd.notna(first) and pd.notna(last) and marc >= 2) else 0
            expected = expected_nodoc if marc >= 1 else 0
            saldo = worked - expected if expected else 0
            if expected and not incompleto:
                cumple = "OK" if saldo >= 0 else "FALTA"
            elif expected and incompleto:
                cumple = "INCOMPLETO"
            else:
                cumple = ""

        rows.append(
            {
                "DNI": dni,
                "Empleado": emp,
                "Tipo": tipo,
                "Fecha": pd.to_datetime(day),
                "Primera": first,
                "Ultima": last,
                "Horas": minutes_to_hhmm(worked),
                "Minutos": int(worked),
                "Esperado_min": int(expected),
                "Esperado": minutes_to_hhmm(expected),
                "Saldo_min": int(saldo),
                "Saldo": delta_short(saldo),
                "Cumple": cumple,
                "Marcaciones": marc,
                "Pares_estimados": int(pairs),
                "Cortes": "SI" if cortes else "",
                "Incompleto": "SI" if incompleto else "",
            }
        )

    d = pd.DataFrame(rows)
    if d.empty:
        return d
    return d.sort_values(["Tipo", "Empleado", "DNI", "Fecha"]).reset_index(drop=True)


# =========================
# Corrección automática NO Docente
# =========================
def correct_missing_punches_for_employee(
    raw_emp: pd.DataFrame,
    expected_nodoc: int,
) -> tuple[pd.DataFrame, int]:
    """
    Aplica corrección SOLO para NO Docente:
    - Si un día tiene 1 marcación: crea la segunda para completar expected_nodoc
      * Si solo hay una marca (t): asume que esa es la entrada y agrega salida = t + expected
      * Si preferís lo inverso, se puede ajustar, pero tu ejemplo es entrada sin salida.
    - Si un día tiene 0 (no debería pasar), no toca.
    Devuelve (raw_emp_corregido, cantidad_correcciones)
    """
    if raw_emp.empty:
        return raw_emp, 0

    corrected = raw_emp.copy()
    corrected["Fecha"] = corrected["FechaHora"].dt.date
    nfix = 0

    fixes = []
    for day, g in corrected.groupby("Fecha"):
        times = sorted(g["FechaHora"].tolist())
        if len(times) == 1:
            t = times[0]
            # regla: si hay una sola marca, la tomamos como entrada (como tu ejemplo)
            fix_out = t + pd.to_timedelta(expected_nodoc, unit="m")
            fixes.append({"FechaHora": fix_out, "Fecha": day, "__FIX__": True})
            nfix += 1

    if fixes:
        fx = pd.DataFrame(fixes)
        # clonamos la estructura
        base = corrected.iloc[0:1].copy()
        base = base.iloc[0:0]  # vacío con columnas

        # construir filas artificiales
        fx_rows = []
        for _, r in fx.iterrows():
            row = corrected.iloc[0].copy()
            row["FechaHora"] = r["FechaHora"]
            row["Fecha"] = r["Fecha"]
            # marcadores opcionales
            row["NvoEstado"] = "AUTO_FIX"  # no importa para cálculo
            fx_rows.append(row)

        add = pd.DataFrame(fx_rows)
        corrected = pd.concat([corrected, add], ignore_index=True).sort_values("FechaHora").reset_index(drop=True)

    # recomputar Fecha
    corrected["Fecha"] = corrected["FechaHora"].dt.date
    return corrected, nfix


# =========================
# KPIs / Summary
# =========================
def summarize(daily: pd.DataFrame) -> pd.DataFrame:
    if daily.empty:
        return pd.DataFrame()

    s = (
        daily.groupby(["Empleado", "DNI", "Tipo"], as_index=False)
        .agg(
            Dias=("Fecha", "nunique"),
            Total_min=("Minutos", "sum"),
            Prom_min=("Minutos", "mean"),
            Incompletos=("Incompleto", lambda x: int((x == "SI").sum())),
            Cortes=("Cortes", lambda x: int((x == "SI").sum())),
            Marcaciones=("Marcaciones", "sum"),
            Esperado_min=("Esperado_min", "sum"),
            Saldo_min=("Saldo_min", "sum"),
        )
        .sort_values(["Tipo", "Empleado"])
        .reset_index(drop=True)
    )

    s["Total"] = s["Total_min"].round().astype(int).apply(minutes_to_hhmm)
    s["Prom/día"] = s["Prom_min"].round().astype(int).apply(minutes_to_hhmm)
    s["Saldo"] = s["Saldo_min"].apply(delta_short)

    def pct_row(r):
        if r["Tipo"] == "NO Docente" and r["Esperado_min"] > 0:
            return f"{(r['Total_min']/r['Esperado_min']*100):.0f}%"
        return ""

    s["Cumplimiento"] = s.apply(pct_row, axis=1)
    return s


def employee_detail_table(daily_emp: pd.DataFrame) -> pd.DataFrame:
    d = daily_emp.copy()
    d["Fecha"] = pd.to_datetime(d["Fecha"]).dt.date
    d["Primera"] = pd.to_datetime(d["Primera"], errors="coerce").dt.strftime("%H:%M")
    d["Ultima"] = pd.to_datetime(d["Ultima"], errors="coerce").dt.strftime("%H:%M")

    cols = [
        "Fecha", "Primera", "Ultima",
        "Horas", "Esperado", "Saldo",
        "Marcaciones", "Pares_estimados", "Cortes", "Incompleto", "Cumple"
    ]
    return d[cols].sort_values("Fecha").reset_index(drop=True)


def export_employee_excel(
    employee_name: str,
    dni: str,
    tipo: str,
    reduced: bool,
    expected: int,
    kpi_block: dict,
    daily_emp: pd.DataFrame,
    raw_emp: pd.DataFrame,
) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        cfg = pd.DataFrame([{
            "Empleado": employee_name,
            "DNI": dni,
            "Tipo": tipo,
            "Horario_reducido": "SI" if reduced else "NO",
            "Esperado_NO_Docente": minutes_to_hhmm(expected) if tipo == "NO Docente" else "",
        }])
        cfg.to_excel(writer, sheet_name="Empleado", index=False)

        kdf = pd.DataFrame([kpi_block])
        kdf.to_excel(writer, sheet_name="KPIs", index=False)

        daily_out = daily_emp.drop(
            columns=["Esperado_min", "Saldo_min"], errors="ignore"
        ).copy()
        daily_out.to_excel(writer, sheet_name="Detalle_Diario", index=False)

        raw_out = raw_emp.copy()
        raw_out["Fecha"] = raw_out["FechaHora"].dt.date
        raw_out["Hora"] = raw_out["FechaHora"].dt.strftime("%H:%M")
        raw_out = raw_out.sort_values("FechaHora")[["Fecha", "Hora"]]
        raw_out.to_excel(writer, sheet_name="Marcaciones", index=False)

        # widths
        for sheet in writer.book.sheetnames:
            ws = writer.book[sheet]
            for col in ws.columns:
                letter = col[0].column_letter
                max_len = 0
                for cell in col:
                    v = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, len(v))
                ws.column_dimensions[letter].width = min(max(10, max_len + 2), 48)

    out.seek(0)
    return out.getvalue()


# =========================
# App
# =========================
def main() -> None:
    st.set_page_config(page_title="Asistencia RRHH", page_icon="🕒", layout="centered")
    inject_css()

    # Header
    a, b, c = st.columns([1.2, 0.85, 1.0])
    with a:
        st.markdown("## 🕒 Asistencia RRHH")
    with b:
        reduced = st.toggle("Activar horario reducido", value=False)
    with c:
        st.markdown(
            f"""<div class="pill">NO Docente esperado: {"06:00" if reduced else "07:00"} · cálculo por tiempo</div>""",
            unsafe_allow_html=True,
        )

    file = st.file_uploader("", type=["xlsx", "xlsm", "xls"], label_visibility="collapsed")
    if not file:
        return

    df0 = read_excel_auto(file)
    validate_format(df0)
    raw0 = parse_and_clean(df0)

    profiles = init_profiles(raw0)

    # Tabs
    tabs = st.tabs(["General", "Empleado", "Perfiles"])

    expected = 360 if reduced else 420

    # =======================
    # GENERAL
    # =======================
    with tabs[0]:
        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        raw = apply_profiles(raw0, st.session_state["profiles"])
        daily = calc_daily(raw, expected)
        summary = summarize(daily)

        # KPIs globales
        total_min = int(daily["Minutos"].sum()) if not daily.empty else 0
        empleados = int(summary.shape[0]) if not summary.empty else 0
        dias = int(daily["Fecha"].nunique()) if not daily.empty else 0
        incompletos = int((daily["Incompleto"] == "SI").sum()) if not daily.empty else 0
        cortes = int((daily["Cortes"] == "SI").sum()) if not daily.empty else 0
        prom_dia = int(round(daily["Minutos"].mean())) if not daily.empty else 0

        nod = daily[daily["Tipo"] == "NO Docente"].copy()
        exp_sum = int(nod["Esperado_min"].sum()) if not nod.empty else 0
        nod_sum = int(nod["Minutos"].sum()) if not nod.empty else 0
        nod_pct = f"{(nod_sum/exp_sum*100):.0f}%" if exp_sum > 0 else ""

        doc = daily[daily["Tipo"] == "Docente"].copy()
        doc_total = int(doc["Minutos"].sum()) if not doc.empty else 0

        k1, k2, k3, k4 = st.columns(4)
        k1.markdown(f"""<div class="kpi"><div class="label">Empleados</div><div class="value">{empleados}</div><div class="sub">Días: {dias}</div></div>""", unsafe_allow_html=True)
        k2.markdown(f"""<div class="kpi"><div class="label">Total</div><div class="value">{minutes_to_hhmm(total_min)}</div><div class="sub">Prom/día: {minutes_to_hhmm(prom_dia)}</div></div>""", unsafe_allow_html=True)
        k3.markdown(f"""<div class="kpi"><div class="label">Incompletos</div><div class="value">{incompletos}</div><div class="sub">días con marcas insuficientes</div></div>""", unsafe_allow_html=True)
        k4.markdown(f"""<div class="kpi"><div class="label">Cortes</div><div class="value">{cortes}</div><div class="sub">varios tramos estimados</div></div>""", unsafe_allow_html=True)

        k5, k6, k7 = st.columns(3)
        k5.markdown(f"""<div class="kpi"><div class="label">NO Docente</div><div class="value">{minutes_to_hhmm(nod_sum)}</div><div class="sub">Cumplimiento: {nod_pct}</div></div>""", unsafe_allow_html=True)
        k6.markdown(f"""<div class="kpi"><div class="label">Docente</div><div class="value">{minutes_to_hhmm(doc_total)}</div><div class="sub">por tramos estimados</div></div>""", unsafe_allow_html=True)
        k7.markdown(f"""<div class="kpi"><div class="label">Marcaciones</div><div class="value">{int(raw.shape[0])}</div><div class="sub">filas del Excel</div></div>""", unsafe_allow_html=True)

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        # Gráfico: horas por día (global)
        if not daily.empty:
            by_day = daily.groupby("Fecha", as_index=False).agg(Minutos=("Minutos", "sum"))
            by_day["Horas"] = by_day["Minutos"] / 60.0
            by_day = by_day.sort_values("Fecha")
            st.markdown("### Horas totales por día")
            st.line_chart(by_day.set_index("Fecha")[["Horas"]], height=220)

        # Rankings
        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
        colL, colR = st.columns(2)

        with colL:
            st.markdown("### Top 10 · más horas")
            top_hours = (
                daily.groupby(["Empleado", "DNI", "Tipo"], as_index=False)
                .agg(Total_min=("Minutos", "sum"))
                .sort_values("Total_min", ascending=False)
                .head(10)
            )
            top_hours["Total"] = top_hours["Total_min"].apply(minutes_to_hhmm)
            st.dataframe(top_hours[["Empleado", "DNI", "Tipo", "Total"]], use_container_width=True, height=320)

        with colR:
            st.markdown("### Top 10 · más incompletos")
            top_inc = (
                daily.groupby(["Empleado", "DNI", "Tipo"], as_index=False)
                .agg(Incompletos=("Incompleto", lambda x: int((x == "SI").sum())))
                .sort_values("Incompletos", ascending=False)
                .head(10)
            )
            st.dataframe(top_inc, use_container_width=True, height=320)

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
        st.markdown("### Resumen")
        st.dataframe(summary, use_container_width=True, height=420)

        # guardar para Empleado
        st.session_state["__raw__"] = raw
        st.session_state["__daily__"] = daily
        st.session_state["__summary__"] = summary

    # =======================
    # EMPLEADO
    # =======================
    with tabs[1]:
        raw = st.session_state.get("__raw__", None)
        daily = st.session_state.get("__daily__", None)
        summary = st.session_state.get("__summary__", None)

        if raw is None or daily is None or summary is None or summary.empty:
            st.info("Cargá un Excel para ver esta sección.")
            return

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        disp = summary.copy()
        disp["Display"] = disp["Empleado"] + " · " + disp["DNI"].astype(str) + " · " + disp["Tipo"]
        selected = st.selectbox("", options=disp["Display"].tolist(), label_visibility="collapsed")
        r = disp[disp["Display"] == selected].iloc[0]
        emp, dni, tipo = r["Empleado"], r["DNI"], r["Tipo"]

        raw_emp = raw[(raw["Empleado"] == emp) & (raw["DNI"] == dni) & (raw["Tipo"] == tipo)].copy()
        raw_emp = raw_emp.sort_values("FechaHora")

        # toggle de corrección SOLO si NO Docente
        fix = False
        fixes_applied = 0

        left, right = st.columns([1.0, 1.0])
        with left:
            if tipo == "NO Docente":
                fix = st.button("Corregir falta de marcación", use_container_width=True)
            else:
                st.markdown("""<div class="pill">Docente: no aplica corrección automática</div>""", unsafe_allow_html=True)

        # aplicar corrección al empleado (solo si apretó botón)
        if tipo == "NO Docente" and fix:
            corrected_raw_emp, fixes_applied = correct_missing_punches_for_employee(raw_emp, expected)
            # reemplazar dentro del raw global para recalcular daily del empleado
            raw_corrected = raw.copy()
            mask = (raw_corrected["Empleado"] == emp) & (raw_corrected["DNI"] == dni) & (raw_corrected["Tipo"] == tipo)
            raw_corrected = raw_corrected[~mask]
            raw_corrected = pd.concat([raw_corrected, corrected_raw_emp], ignore_index=True)
            raw_corrected = raw_corrected.sort_values(["Empleado", "DNI", "FechaHora"]).reset_index(drop=True)

            # recalcular daily completo, pero mostramos el empleado
            daily_corrected = calc_daily(raw_corrected, expected)
            daily_emp = daily_corrected[(daily_corrected["Empleado"] == emp) & (daily_corrected["DNI"] == dni) & (daily_corrected["Tipo"] == tipo)].copy()
            raw_emp = corrected_raw_emp
        else:
            daily_emp = daily[(daily["Empleado"] == emp) & (daily["DNI"] == dni) & (daily["Tipo"] == tipo)].copy()

        # KPIs del empleado
        total_min = int(daily_emp["Minutos"].sum()) if not daily_emp.empty else 0
        dias = int(daily_emp["Fecha"].nunique()) if not daily_emp.empty else 0
        prom = int(round(daily_emp["Minutos"].mean())) if dias else 0
        inc = int((daily_emp["Incompleto"] == "SI").sum()) if not daily_emp.empty else 0
        cuts = int((daily_emp["Cortes"] == "SI").sum()) if not daily_emp.empty else 0
        marc = int(daily_emp["Marcaciones"].sum()) if not daily_emp.empty else 0

        exp_sum = int(daily_emp["Esperado_min"].sum()) if not daily_emp.empty else 0
        saldo_sum = int(daily_emp["Saldo_min"].sum()) if not daily_emp.empty else 0

        # EXTRAS DEL MES (NO Docente): solo positivos acumulados
        extras_min = 0
        faltas_min = 0
        if tipo == "NO Docente" and not daily_emp.empty:
            extras_min = int(daily_emp.loc[daily_emp["Saldo_min"] > 0, "Saldo_min"].sum())
            faltas_min = int((-daily_emp.loc[daily_emp["Saldo_min"] < 0, "Saldo_min"].sum()))

        pct = f"{(total_min/exp_sum*100):.0f}%" if exp_sum > 0 else ""

        k1, k2, k3, k4 = st.columns(4)
        k1.markdown(f"""<div class="kpi"><div class="label">{emp}</div><div class="value">{minutes_to_hhmm(total_min)}</div><div class="sub">{tipo} · DNI {dni}</div></div>""", unsafe_allow_html=True)
        k2.markdown(f"""<div class="kpi"><div class="label">Días</div><div class="value">{dias}</div><div class="sub">Prom/día: {minutes_to_hhmm(prom)}</div></div>""", unsafe_allow_html=True)
        k3.markdown(f"""<div class="kpi"><div class="label">Marcaciones</div><div class="value">{marc}</div><div class="sub">Cortes: {cuts} · Incompletos: {inc}</div></div>""", unsafe_allow_html=True)

        if tipo == "NO Docente":
            k4.markdown(
                f"""<div class="kpi"><div class="label">Extras del mes</div><div class="value">{minutes_to_hhmm(extras_min)}</div>
                <div class="sub">Faltas: {minutes_to_hhmm(faltas_min)} · Cumplimiento: {pct}</div></div>""",
                unsafe_allow_html=True
            )
        else:
            k4.markdown(f"""<div class="kpi"><div class="label">Total tramos</div><div class="value">{minutes_to_hhmm(total_min)}</div><div class="sub">Cálculo por alternancia</div></div>""", unsafe_allow_html=True)

        if fixes_applied:
            st.markdown(f"""<div class="pill">Correcciones aplicadas: {fixes_applied}</div>""", unsafe_allow_html=True)

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        # Gráfico del empleado: horas por día
        if not daily_emp.empty:
            chart_emp = daily_emp.copy()
            chart_emp = chart_emp.sort_values("Fecha")
            chart_emp["Horas_float"] = chart_emp["Minutos"] / 60.0
            st.markdown("### Horas por día (empleado)")
            st.bar_chart(chart_emp.set_index("Fecha")[["Horas_float"]], height=230)

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        # Tabla día a día
        st.dataframe(employee_detail_table(daily_emp), use_container_width=True, height=520)

        # Export SOLO empleado (incluye KPIs + diario + marcaciones)
        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        kpi_block = {
            "Empleado": emp,
            "DNI": dni,
            "Tipo": tipo,
            "Horario_reducido": "SI" if reduced else "NO",
            "Dias": dias,
            "Total": minutes_to_hhmm(total_min),
            "Prom_dia": minutes_to_hhmm(prom),
            "Incompletos": inc,
            "Cortes": cuts,
            "Cumplimiento_NO_Docente": pct if tipo == "NO Docente" else "",
            "Extras_mes_NO_Docente": minutes_to_hhmm(extras_min) if tipo == "NO Docente" else "",
            "Faltas_mes_NO_Docente": minutes_to_hhmm(faltas_min) if tipo == "NO Docente" else "",
            "Saldo_acumulado_NO_Docente": delta_short(saldo_sum) if tipo == "NO Docente" else "",
            "Correcciones_aplicadas": fixes_applied,
        }

        excel_emp = export_employee_excel(
            employee_name=emp,
            dni=str(dni),
            tipo=tipo,
            reduced=reduced,
            expected=expected,
            kpi_block=kpi_block,
            daily_emp=daily_emp,
            raw_emp=raw_emp,
        )
        st.download_button(
            "Exportar empleado (KPIs + diario + marcaciones)",
            data=excel_emp,
            file_name=f"empleado_{emp.replace(' ', '_')}_{dni}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # =======================
    # PERFILES
    # =======================
    with tabs[2]:
        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
        edited = st.data_editor(
            st.session_state["profiles"],
            use_container_width=True,
            height=520,
            hide_index=True,
            column_config={
                "Tipo": st.column_config.SelectboxColumn("Tipo", options=["NO Docente", "Docente"], required=True)
            },
        )
        st.session_state["profiles"] = edited


if __name__ == "__main__":
    main()
