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
        .block-container { max-width: 1280px; padding-top: 1rem; padding-bottom: 1.5rem; }
        header, footer {visibility: hidden;}
        div[data-testid="stToolbar"] {visibility: hidden; height: 0px;}
        .kpi {
            border: 1px solid rgba(255,255,255,0.10);
            border-radius: 18px;
            padding: 14px 14px;
            background: rgba(255,255,255,0.03);
        }
        .kpi .label {opacity:.75; font-size:.92rem;}
        .kpi .value {font-size:1.65rem; font-weight:900; line-height:1.1;}
        .kpi .sub {opacity:.70; font-size:.86rem; margin-top:.18rem;}
        .hr {height:1px; background: rgba(255,255,255,0.08); margin: 0.9rem 0 1.0rem 0;}
        .pill {display:inline-block; padding:6px 10px; border-radius:999px;
               border:1px solid rgba(255,255,255,.12); background:rgba(255,255,255,.03);
               font-size:.85rem; opacity:.9;}
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
    mins = abs(mins)
    h, m = mins // 60, mins % 60
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
        raise ValueError(f"Formato incorrecto. Faltan columnas: {', '.join(missing)}")


def parse_and_clean(df: pd.DataFrame) -> pd.DataFrame:
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


# =========================
# Core: cálculo SIN confiar en IN/OUT
# =========================
def pair_alternating(times: list[pd.Timestamp]) -> tuple[int, int]:
    """
    Empareja alternando: (t0,t1), (t2,t3) ...
    Devuelve (minutos_trabajados_por_pares, cantidad_de_pares)
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

        # flags
        incompleto = (marc < 2) or (pairs == 0)
        cortes = (pairs >= 2)  # hubo al menos 2 tramos estimados

        if tipo == "Docente":
            worked = worked_pairs
            expected = 0
            saldo = 0
            cumple = ""
        else:
            # NO Docente: usamos "corrido" (última - primera)
            # (porque tu regla es de corrido)
            worked = span if not incompleto else 0
            expected = expected_nodoc if (marc >= 1) else 0
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
                "Esperado": minutes_to_hhmm(expected),
                "Esperado_min": int(expected),
                "Saldo_min": int(saldo),
                "Saldo": delta_short(saldo),
                "Cumple": cumple,
                "Marcaciones": marc,
                "Pares_estimados": int(pairs),
                "Cortes": "SI" if cortes else "",
                "Incompleto": "SI" if incompleto else "",
                # auditoría:
                "Tramos_min_estimados": int(worked_pairs),
                "Corrido_min": int(span),
            }
        )
    d = pd.DataFrame(rows)
    if d.empty:
        return d
    return d.sort_values(["Tipo", "Empleado", "DNI", "Fecha"]).reset_index(drop=True)


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

    s = s.drop(columns=["Total_min", "Prom_min"])
    return s


def employee_detail_table(daily_emp: pd.DataFrame) -> pd.DataFrame:
    d = daily_emp.copy()
    d["Fecha"] = pd.to_datetime(d["Fecha"]).dt.date
    d["Primera"] = pd.to_datetime(d["Primera"], errors="coerce").dt.strftime("%H:%M")
    d["Ultima"] = pd.to_datetime(d["Ultima"], errors="coerce").dt.strftime("%H:%M")

    # orden limpio
    cols = [
        "Fecha", "Primera", "Ultima",
        "Horas", "Esperado", "Saldo",
        "Marcaciones", "Pares_estimados", "Cortes", "Incompleto", "Cumple"
    ]
    return d[cols].sort_values("Fecha").reset_index(drop=True)


def export_excel(summary: pd.DataFrame, daily: pd.DataFrame, raw: pd.DataFrame, profiles: pd.DataFrame, reduced: bool) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        cfg = pd.DataFrame([{
            "Horario_reducido": "SI" if reduced else "NO",
            "NO_Docente_esperado": "06:00" if reduced else "07:00",
            "Docente": "Libre (por tramos estimados)",
            "Metodo_tramos": "Alternancia por tiempo (no depende de IN/OUT)",
        }])
        cfg.to_excel(writer, sheet_name="Config", index=False)

        profiles.to_excel(writer, sheet_name="Perfiles", index=False)
        summary.to_excel(writer, sheet_name="Resumen", index=False)

        daily_out = daily.drop(columns=["Minutos", "Esperado_min", "Saldo_min", "Tramos_min_estimados", "Corrido_min"], errors="ignore")
        daily_out.to_excel(writer, sheet_name="Detalle_Diario", index=False)

        raw_out = raw[["DNI", "Empleado", "Tipo", "FechaHora"]].copy()
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
# Widgets de rankings
# =========================
def top_table(summary: pd.DataFrame, col: str, n=10, asc=False) -> pd.DataFrame:
    if summary.empty:
        return pd.DataFrame()
    t = summary.copy()
    if col not in t.columns:
        return pd.DataFrame()
    # ordenar con num si existe _min
    return t.sort_values(col, ascending=asc).head(n).reset_index(drop=True)


# =========================
# App
# =========================
def main() -> None:
    st.set_page_config(page_title="Asistencia RRHH", page_icon="🕒", layout="centered")
    inject_css()

    # Header
    h1, h2, h3 = st.columns([1.25, 0.85, 1.05])
    with h1:
        st.markdown("## 🕒 Asistencia RRHH")
    with h2:
        reduced = st.toggle("Activar horario reducido", value=False)
    with h3:
        st.markdown(f"""<div class="pill">NO Docente esperado: {"06:00" if reduced else "07:00"} · Tramos por tiempo</div>""",
                    unsafe_allow_html=True)

    file = st.file_uploader("", type=["xlsx", "xlsm", "xls"], label_visibility="collapsed")
    if not file:
        return

    df0 = read_excel_auto(file)
    validate_format(df0)
    raw0 = parse_and_clean(df0)

    profiles = init_profiles(raw0)

    tabs = st.tabs(["General", "Empleado", "Perfiles", "Exportar"])

    expected = 360 if reduced else 420

    # =======================
    # Tab General
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

        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f"""<div class="kpi"><div class="label">Empleados</div><div class="value">{empleados}</div><div class="sub">Días: {dias}</div></div>""", unsafe_allow_html=True)
        c2.markdown(f"""<div class="kpi"><div class="label">Total</div><div class="value">{minutes_to_hhmm(total_min)}</div><div class="sub">Prom/día: {minutes_to_hhmm(prom_dia)}</div></div>""", unsafe_allow_html=True)
        c3.markdown(f"""<div class="kpi"><div class="label">Incompletos</div><div class="value">{incompletos}</div><div class="sub"><2 marcas o sin pares</div></div>""", unsafe_allow_html=True)
        c4.markdown(f"""<div class="kpi"><div class="label">Cortes</div><div class="value">{cortes}</div><div class="sub">Múltiples tramos estimados</div></div>""", unsafe_allow_html=True)

        c5, c6, c7 = st.columns(3)
        c5.markdown(f"""<div class="kpi"><div class="label">NO Docente</div><div class="value">{minutes_to_hhmm(int(nod_sum))}</div><div class="sub">Cumplimiento: {nod_pct}</div></div>""", unsafe_allow_html=True)
        c6.markdown(f"""<div class="kpi"><div class="label">Docente</div><div class="value">{minutes_to_hhmm(int(doc_total))}</div><div class="sub">por tramos estimados</div></div>""", unsafe_allow_html=True)
        c7.markdown(f"""<div class="kpi"><div class="label">Marcaciones</div><div class="value">{int(raw.shape[0])}</div><div class="sub">filas del Excel válidas</div></div>""", unsafe_allow_html=True)

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        # Rankings
        # Creamos columnas numéricas para ranking: ya tenemos Saldo_min, Incompletos, Cortes, Marcaciones, Dias
        # Armamos mini tablas limpias
        s = summary.copy()

        # Para ranking de saldo, usamos Saldo_min (que está en summary)
        # Si no existe porque algo raro, lo omitimos.
        grid1, grid2 = st.columns(2)

        with grid1:
            st.markdown("### Rankings")
            # Top total horas (todos)
            t1 = s.copy()
            # reconstruimos un total min auxiliar si no está; pero summary ya lo perdió
            # Para ranking total, lo hacemos desde daily:
            total_emp = (
                daily.groupby(["Empleado", "DNI", "Tipo"], as_index=False)
                .agg(Total_min=("Minutos", "sum"))
                .sort_values("Total_min", ascending=False)
                .head(10)
            )
            total_emp["Total"] = total_emp["Total_min"].apply(minutes_to_hhmm)
            st.dataframe(total_emp[["Empleado", "DNI", "Tipo", "Total"]], use_container_width=True, height=320)

        with grid2:
            # Top saldos negativos (NO Docente)
            nod_daily = daily[daily["Tipo"] == "NO Docente"].copy()
            if not nod_daily.empty:
                saldo_emp = (
                    nod_daily.groupby(["Empleado", "DNI"], as_index=False)
                    .agg(Saldo_min=("Saldo_min", "sum"))
                    .sort_values("Saldo_min", ascending=True)
                    .head(10)
                )
                saldo_emp["Saldo"] = saldo_emp["Saldo_min"].apply(delta_short)
                st.markdown("### NO Docente · más falta")
                st.dataframe(saldo_emp[["Empleado", "DNI", "Saldo"]], use_container_width=True, height=320)
            else:
                st.markdown("### NO Docente · más falta")
                st.dataframe(pd.DataFrame(columns=["Empleado", "DNI", "Saldo"]), use_container_width=True, height=320)

        grid3, grid4 = st.columns(2)
        with grid3:
            st.markdown("### Incompletos (top)")
            inc_emp = (
                daily.groupby(["Empleado", "DNI", "Tipo"], as_index=False)
                .agg(Incompletos=("Incompleto", lambda x: int((x == "SI").sum())))
                .sort_values("Incompletos", ascending=False)
                .head(10)
            )
            st.dataframe(inc_emp, use_container_width=True, height=300)

        with grid4:
            st.markdown("### Cortes (top)")
            cut_emp = (
                daily.groupby(["Empleado", "DNI", "Tipo"], as_index=False)
                .agg(Cortes=("Cortes", lambda x: int((x == "SI").sum())))
                .sort_values("Cortes", ascending=False)
                .head(10)
            )
            st.dataframe(cut_emp, use_container_width=True, height=300)

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
        st.markdown("### Resumen completo")
        st.dataframe(summary, use_container_width=True, height=420)

        # Guardar para tabs siguientes
        st.session_state["__raw__"] = raw
        st.session_state["__daily__"] = daily
        st.session_state["__summary__"] = summary

    # =======================
    # Tab Empleado
    # =======================
    with tabs[1]:
        raw = st.session_state.get("__raw__", None)
        daily = st.session_state.get("__daily__", None)
        summary = st.session_state.get("__summary__", None)

        if raw is None or daily is None or summary is None or summary.empty:
            st.info("Cargá un Excel para ver esta sección.")
        else:
            st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

            disp = summary.copy()
            disp["Display"] = disp["Empleado"] + " · " + disp["DNI"].astype(str) + " · " + disp["Tipo"]
            selected = st.selectbox("", options=disp["Display"].tolist(), label_visibility="collapsed")
            r = disp[disp["Display"] == selected].iloc[0]
            emp, dni, tipo = r["Empleado"], r["DNI"], r["Tipo"]

            d_emp = daily[(daily["Empleado"] == emp) & (daily["DNI"] == dni) & (daily["Tipo"] == tipo)].copy()
            if d_emp.empty:
                st.warning("Sin datos para este empleado.")
            else:
                total_min = int(d_emp["Minutos"].sum())
                dias = int(d_emp["Fecha"].nunique())
                prom = int(round(d_emp["Minutos"].mean())) if dias else 0
                inc = int((d_emp["Incompleto"] == "SI").sum())
                cuts = int((d_emp["Cortes"] == "SI").sum())
                marc = int(d_emp["Marcaciones"].sum())
                pares = int(d_emp["Pares_estimados"].sum())

                # NO Docente extra KPIs
                exp_sum = int(d_emp["Esperado_min"].sum())
                saldo_sum = int(d_emp["Saldo_min"].sum())
                pct = f"{(total_min/exp_sum*100):.0f}%" if exp_sum > 0 else ""
                faltas = int(((d_emp["Saldo_min"] < 0) & (d_emp["Esperado_min"] > 0) & (d_emp["Incompleto"] != "SI")).sum())
                extras = int(((d_emp["Saldo_min"] > 0) & (d_emp["Esperado_min"] > 0) & (d_emp["Incompleto"] != "SI")).sum())

                k1, k2, k3, k4 = st.columns(4)
                k1.markdown(f"""<div class="kpi"><div class="label">{emp}</div><div class="value">{minutes_to_hhmm(total_min)}</div><div class="sub">{tipo}</div></div>""",
                            unsafe_allow_html=True)
                k2.markdown(f"""<div class="kpi"><div class="label">Días</div><div class="value">{dias}</div><div class="sub">Prom/día: {minutes_to_hhmm(prom)}</div></div>""",
                            unsafe_allow_html=True)
                k3.markdown(f"""<div class="kpi"><div class="label">Marcaciones</div><div class="value">{marc}</div><div class="sub">Pares estimados: {pares}</div></div>""",
                            unsafe_allow_html=True)
                k4.markdown(f"""<div class="kpi"><div class="label">Alertas</div><div class="value">{inc}</div><div class="sub">Cortes: {cuts}</div></div>""",
                            unsafe_allow_html=True)

                if tipo == "NO Docente" and pct:
                    a, b, c = st.columns(3)
                    a.markdown(f"""<div class="kpi"><div class="label">Cumplimiento</div><div class="value">{pct}</div><div class="sub">vs esperado</div></div>""",
                               unsafe_allow_html=True)
                    b.markdown(f"""<div class="kpi"><div class="label">Saldo</div><div class="value">{delta_short(saldo_sum)}</div><div class="sub">acumulado</div></div>""",
                               unsafe_allow_html=True)
                    c.markdown(f"""<div class="kpi"><div class="label">Días</div><div class="value">{extras}/{faltas}</div><div class="sub">extra / falta</div></div>""",
                               unsafe_allow_html=True)

                st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

                st.dataframe(employee_detail_table(d_emp), use_container_width=True, height=500)

                with st.expander("Marcaciones del empleado"):
                    raw_emp = raw[(raw["Empleado"] == emp) & (raw["DNI"] == dni) & (raw["Tipo"] == tipo)].copy()
                    raw_emp = raw_emp.sort_values("FechaHora")
                    raw_emp["Fecha"] = raw_emp["FechaHora"].dt.date
                    raw_emp["Hora"] = raw_emp["FechaHora"].dt.strftime("%H:%M")
                    st.dataframe(raw_emp[["Fecha", "Hora"]], use_container_width=True, height=380)

    # =======================
    # Tab Perfiles
    # =======================
    with tabs[2]:
        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
        st.data_editor(
            st.session_state["profiles"],
            use_container_width=True,
            height=520,
            hide_index=True,
            column_config={
                "Tipo": st.column_config.SelectboxColumn("Tipo", options=["NO Docente", "Docente"], required=True)
            },
        )

    # =======================
    # Tab Exportar
    # =======================
    with tabs[3]:
        raw = st.session_state.get("__raw__", None)
        daily = st.session_state.get("__daily__", None)
        summary = st.session_state.get("__summary__", None)

        if raw is None or daily is None or summary is None:
            st.info("Cargá un Excel primero.")
        else:
            st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
            excel_bytes = export_excel(summary, daily, raw, st.session_state["profiles"], reduced)
            st.download_button(
                "Descargar Excel",
                data=excel_bytes,
                file_name="asistencia_rrhh.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


if __name__ == "__main__":
    main()
