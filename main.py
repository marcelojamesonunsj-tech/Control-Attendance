from __future__ import annotations

import io
import pandas as pd
import streamlit as st

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


REQUIRED_COLS = ["Nombre", "Marc.", "Estado", "NvoEstado"]


# =========================
# UI
# =========================
def inject_css() -> None:
    st.markdown(
        """
        <style>
        .block-container { max-width: 1380px; padding-top: 1rem; padding-bottom: 1.4rem; }
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


def kpi_card(label: str, value: str, sub: str = "") -> None:
    st.markdown(
        f"""<div class="kpi"><div class="label">{label}</div><div class="value">{value}</div><div class="sub">{sub}</div></div>""",
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
    - NvoEstado = no confiamos (a veces todo queda como entrada)
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

        incompleto = (marc < 2) if tipo == "NO Docente" else (pairs == 0)
        cortes = (pairs >= 2)

        if tipo == "Docente":
            worked = worked_pairs
            expected = 0
            saldo = 0
            cumple = ""
        else:
            worked = span if (marc >= 2 and pd.notna(first) and pd.notna(last)) else 0
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
def correct_missing_punches_for_employee(raw_emp: pd.DataFrame, expected_nodoc: int) -> tuple[pd.DataFrame, int]:
    """
    SOLO NO Docente:
    - Si un día tiene 1 marcación: agrega la segunda para completar expected (6/7h)
      Regla: la única marca se toma como entrada.
    """
    if raw_emp.empty:
        return raw_emp, 0

    corrected = raw_emp.copy()
    corrected["Fecha"] = corrected["FechaHora"].dt.date

    fixes = []
    nfix = 0
    for day, g in corrected.groupby("Fecha"):
        times = sorted(g["FechaHora"].tolist())
        if len(times) == 1:
            t = times[0]
            fix_out = t + pd.to_timedelta(expected_nodoc, unit="m")
            fixes.append({"FechaHora": fix_out, "Fecha": day})
            nfix += 1

    if fixes:
        fx_rows = []
        template = corrected.iloc[0].copy()
        for f in fixes:
            row = template.copy()
            row["FechaHora"] = f["FechaHora"]
            row["Fecha"] = f["Fecha"]
            row["NvoEstado"] = "AUTO_FIX"
            row["Estado"] = row["FechaHora"].strftime("%d/%m/%Y %H:%M")
            fx_rows.append(row)

        add = pd.DataFrame(fx_rows)
        corrected = pd.concat([corrected, add], ignore_index=True).sort_values("FechaHora").reset_index(drop=True)

    corrected["Fecha"] = corrected["FechaHora"].dt.date
    return corrected, nfix


def correct_missing_punches_all(raw: pd.DataFrame, expected_nodoc: int) -> tuple[pd.DataFrame, int]:
    """
    Corrección masiva para TODOS los NO Docentes:
    - Para cada empleado NO Docente, agrega marca faltante en días con 1 sola marca.
    """
    if raw.empty:
        return raw, 0

    docentes = raw[raw["Tipo"] == "Docente"].copy()
    nodoc = raw[raw["Tipo"] == "NO Docente"].copy()
    if nodoc.empty:
        return raw, 0

    fixed_parts = []
    total_fixes = 0

    for (dni, emp), g in nodoc.groupby(["DNI", "Empleado"]):
        corrected, nfix = correct_missing_punches_for_employee(g.copy(), expected_nodoc)
        fixed_parts.append(corrected)
        total_fixes += nfix

    nodoc_fixed = pd.concat(fixed_parts, ignore_index=True) if fixed_parts else nodoc
    out = pd.concat([docentes, nodoc_fixed], ignore_index=True)
    out = out.sort_values(["Empleado", "DNI", "FechaHora"]).reset_index(drop=True)
    return out, total_fixes


# =========================
# KPIs / Summary
# =========================
def summarize(daily: pd.DataFrame) -> pd.DataFrame:
    """
    Agrega:
    - Extras (min/HH:MM) para NO Docente = SUM(max(0, saldo_dia))
    - Faltas (min/HH:MM) para NO Docente = SUM(max(0, -saldo_dia))
    - Saldo neto = SUM(saldo_dia)
    """
    if daily.empty:
        return pd.DataFrame()

    def extras_pos(x: pd.Series) -> int:
        return int(x[x > 0].sum())

    def faltas_pos(x: pd.Series) -> int:
        return int((-x[x < 0]).sum())

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
            Extras_min=("Saldo_min", extras_pos),
            Faltas_min=("Saldo_min", faltas_pos),
            Dias_OK=("Cumple", lambda x: int((x == "OK").sum())),
            Dias_FALTA=("Cumple", lambda x: int((x == "FALTA").sum())),
            Dias_INCOMPL=("Cumple", lambda x: int((x == "INCOMPLETO").sum())),
        )
        .sort_values(["Tipo", "Empleado"])
        .reset_index(drop=True)
    )

    s["Total"] = s["Total_min"].round().astype(int).apply(minutes_to_hhmm)
    s["Prom/día"] = s["Prom_min"].round().astype(int).apply(minutes_to_hhmm)

    s["Extras"] = s["Extras_min"].apply(minutes_to_hhmm)
    s["Faltas"] = s["Faltas_min"].apply(minutes_to_hhmm)
    s["Saldo"] = s["Saldo_min"].apply(delta_short)

    def pct_row(r):
        if r["Tipo"] == "NO Docente" and r["Esperado_min"] > 0:
            return f"{(r['Total_min']/r['Esperado_min']*100):.0f}%"
        return ""

    s["Cumplimiento"] = s.apply(pct_row, axis=1)

    # orden de columnas “lindo”
    cols = [
        "Empleado", "DNI", "Tipo",
        "Dias",
        "Total", "Total_min",
        "Prom/día",
        "Esperado_min",
        "Extras", "Extras_min",
        "Faltas", "Faltas_min",
        "Saldo", "Saldo_min",
        "Cumplimiento",
        "Marcaciones",
        "Incompletos",
        "Cortes",
        "Dias_OK", "Dias_FALTA", "Dias_INCOMPL",
    ]
    cols = [c for c in cols if c in s.columns]
    return s[cols]


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

        pd.DataFrame([kpi_block]).to_excel(writer, sheet_name="KPIs", index=False)

        daily_out = daily_emp.drop(columns=["Esperado_min", "Saldo_min"], errors="ignore").copy()
        daily_out.to_excel(writer, sheet_name="Detalle_Diario", index=False)

        raw_out = raw_emp.copy()
        raw_out["Fecha"] = raw_out["FechaHora"].dt.date
        raw_out["Hora"] = raw_out["FechaHora"].dt.strftime("%H:%M")
        raw_out = raw_out.sort_values("FechaHora")[["Fecha", "Hora"]]
        raw_out.to_excel(writer, sheet_name="Marcaciones", index=False)

        for sheet in writer.book.sheetnames:
            ws = writer.book[sheet]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                letter = col[0].column_letter
                max_len = 0
                for cell in col:
                    v = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, len(v))
                ws.column_dimensions[letter].width = min(max(10, max_len + 2), 52)

    out.seek(0)
    return out.getvalue()


# =========================
# Export General Excel (bonito)
# =========================
def _apply_excel_style(ws, table_name: str) -> None:
    # header style
    header_fill = PatternFill("solid", fgColor="1F4E79")  # azul oscuro
    header_font = Font(bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center", horizontal="center")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # table style
    max_row = ws.max_row
    max_col = ws.max_column
    if max_row >= 2 and max_col >= 1:
        ref = f"A1:{get_column_letter(max_col)}{max_row}"
        tab = Table(displayName=table_name, ref=ref)
        style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        tab.tableStyleInfo = style
        ws.add_table(tab)

    # column widths
    for col in range(1, max_col + 1):
        letter = get_column_letter(col)
        max_len = 0
        for r in range(1, max_row + 1):
            v = ws.cell(row=r, column=col).value
            s = "" if v is None else str(v)
            max_len = max(max_len, len(s))
        ws.column_dimensions[letter].width = min(max(10, max_len + 2), 60)


def export_general_excel(
    reduced: bool,
    expected: int,
    kpis_general: dict,
    summary_all: pd.DataFrame,
    extras_only: pd.DataFrame,
    daily: pd.DataFrame,
    raw: pd.DataFrame,
) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    def add_df(sheet_name: str, df: pd.DataFrame):
        ws = wb.create_sheet(sheet_name)
        # write header
        for j, col in enumerate(df.columns, start=1):
            ws.cell(row=1, column=j, value=col)
        # write rows
        for i, row in enumerate(df.itertuples(index=False), start=2):
            for j, val in enumerate(row, start=1):
                ws.cell(row=i, column=j, value=val)
        _apply_excel_style(ws, table_name=sheet_name.replace("-", "_").replace(" ", "_"))

    # KPIs General
    df_kpis = pd.DataFrame([{
        "Horario_reducido": "SI" if reduced else "NO",
        "Esperado_NO_Docente": minutes_to_hhmm(expected),
        **kpis_general
    }])
    add_df("KPIs_General", df_kpis)

    # Resumen (todos)
    add_df("Resumen_Empleados", summary_all.copy())

    # Extras (solo NO Docente)
    add_df("Extras_Empleados", extras_only.copy())

    # Detalle Diario
    daily_out = daily.copy()
    daily_out["Fecha"] = pd.to_datetime(daily_out["Fecha"]).dt.strftime("%Y-%m-%d")
    daily_out["Primera"] = pd.to_datetime(daily_out["Primera"], errors="coerce").dt.strftime("%Y-%m-%d %H:%M")
    daily_out["Ultima"] = pd.to_datetime(daily_out["Ultima"], errors="coerce").dt.strftime("%Y-%m-%d %H:%M")
    add_df("Detalle_Diario", daily_out.copy())

    # Marcaciones
    raw_out = raw.copy()
    raw_out["Fecha"] = pd.to_datetime(raw_out["FechaHora"]).dt.strftime("%Y-%m-%d")
    raw_out["Hora"] = pd.to_datetime(raw_out["FechaHora"]).dt.strftime("%H:%M")
    raw_out = raw_out.sort_values(["Empleado", "DNI", "FechaHora"])[["Empleado", "DNI", "Tipo", "Fecha", "Hora"]]
    add_df("Marcaciones", raw_out.copy())

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


# =========================
# Estadísticas generales (helpers)
# =========================
def safe_pct(a: int, b: int) -> str:
    if b <= 0:
        return "0%"
    return f"{(a / b * 100):.0f}%"


def histogram_hours(series_minutes: pd.Series, bin_hours: list[tuple[float, float]]) -> pd.DataFrame:
    hours = series_minutes.fillna(0).astype(int) / 60.0
    rows = []
    for lo, hi in bin_hours:
        label = f"{lo:.0f}-{hi:.0f}h" if hi < 999 else f"{lo:.0f}h+"
        if hi >= 999:
            cnt = int((hours >= lo).sum())
        else:
            cnt = int(((hours >= lo) & (hours < hi)).sum())
        rows.append({"Rango": label, "Días": cnt})
    return pd.DataFrame(rows)


# =========================
# App
# =========================
def main() -> None:
    st.set_page_config(page_title="Asistencia RRHH", page_icon="🕒", layout="centered")
    inject_css()

    a, b, c = st.columns([1.2, 0.9, 1.0])
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
    _ = init_profiles(raw0)

    expected = 360 if reduced else 420
    tabs = st.tabs(["General", "Empleado", "Perfiles"])

    # =======================
    # GENERAL
    # =======================
    with tabs[0]:
        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        raw = apply_profiles(raw0, st.session_state["profiles"])

        # Botón masivo
        left, right = st.columns([1.1, 1.0])
        with left:
            fix_all = st.button("Corregir faltas de marcación (TODOS)", use_container_width=True)
        with right:
            st.markdown("""<div class="pill">Aplica SOLO a NO Docentes (días con 1 marcación)</div>""", unsafe_allow_html=True)

        fixes_total = 0
        if fix_all:
            raw, fixes_total = correct_missing_punches_all(raw, expected)
            st.markdown(f"""<div class="pill">Correcciones aplicadas (total): {fixes_total}</div>""", unsafe_allow_html=True)

        daily = calc_daily(raw, expected)
        summary = summarize(daily)

        # Global KPIs
        total_min = int(daily["Minutos"].sum()) if not daily.empty else 0
        empleados = int(summary.shape[0]) if not summary.empty else 0
        dias = int(daily["Fecha"].nunique()) if not daily.empty else 0
        prom_dia = int(round(daily["Minutos"].mean())) if not daily.empty else 0
        mediana_dia = int(round(daily["Minutos"].median())) if not daily.empty else 0
        p90_dia = int(round(daily["Minutos"].quantile(0.90))) if not daily.empty else 0

        total_marc = int(daily["Marcaciones"].sum()) if not daily.empty else 0
        prom_marc_dia = (total_marc / max(int(daily.shape[0]), 1)) if not daily.empty else 0

        incompletos = int((daily["Incompleto"] == "SI").sum()) if not daily.empty else 0
        cortes = int((daily["Cortes"] == "SI").sum()) if not daily.empty else 0
        total_registros_dia = int(daily.shape[0]) if not daily.empty else 0

        nod = daily[daily["Tipo"] == "NO Docente"].copy()
        nod_sum = int(nod["Minutos"].sum()) if not nod.empty else 0
        nod_exp_sum = int(nod["Esperado_min"].sum()) if not nod.empty else 0
        nod_pct = f"{(nod_sum / nod_exp_sum * 100):.0f}%" if nod_exp_sum > 0 else ""

        # IMPORTANTE: extras = solo lo que supera esperado
        nod_extras = int(nod.loc[nod["Saldo_min"] > 0, "Saldo_min"].sum()) if not nod.empty else 0
        nod_faltas = int((-nod.loc[nod["Saldo_min"] < 0, "Saldo_min"].sum())) if not nod.empty else 0
        nod_saldo = int(nod["Saldo_min"].sum()) if not nod.empty else 0

        doc = daily[daily["Tipo"] == "Docente"].copy()
        doc_sum = int(doc["Minutos"].sum()) if not doc.empty else 0

        # KPIs burbujas
        r1 = st.columns(4)
        with r1[0]: kpi_card("Empleados", f"{empleados}", f"Días: {dias}")
        with r1[1]: kpi_card("Total", minutes_to_hhmm(total_min), f"Prom/día: {minutes_to_hhmm(prom_dia)}")
        with r1[2]: kpi_card("Mediana / P90", f"{minutes_to_hhmm(mediana_dia)}", f"P90: {minutes_to_hhmm(p90_dia)}")
        with r1[3]: kpi_card("Marcaciones", f"{total_marc}", f"Prom por día: {prom_marc_dia:.2f}")

        r2 = st.columns(4)
        with r2[0]: kpi_card("Incompletos", f"{incompletos}", f"{safe_pct(incompletos, total_registros_dia)} de los días")
        with r2[1]: kpi_card("Cortes", f"{cortes}", f"{safe_pct(cortes, total_registros_dia)} de los días")
        with r2[2]: kpi_card("NO Docente", minutes_to_hhmm(nod_sum), f"Cumplimiento: {nod_pct}")
        with r2[3]: kpi_card("Docente", minutes_to_hhmm(doc_sum), "Tramos estimados por tiempo")

        r3 = st.columns(4)
        with r3[0]: kpi_card("Extras NO Docente", minutes_to_hhmm(nod_extras), "Solo lo que supera el esperado")
        with r3[1]: kpi_card("Faltas NO Docente", minutes_to_hhmm(nod_faltas), "Debajo del esperado (no resta extras)")
        with r3[2]: kpi_card("Saldo neto NO Docente", delta_short(nod_saldo), "Extras - faltas (informativo)")
        with r3[3]:
            ok = int((nod["Cumple"] == "OK").sum()) if not nod.empty else 0
            fa = int((nod["Cumple"] == "FALTA").sum()) if not nod.empty else 0
            ic = int((nod["Cumple"] == "INCOMPLETO").sum()) if not nod.empty else 0
            kpi_card("NO Docente días", f"{ok}/{fa}/{ic}", "OK / FALTA / INCOMP")

        # Export GENERAL
        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        # tablas útiles para export
        extras_only = summary[summary["Tipo"] == "NO Docente"].copy()
        # en extras_only dejamos columnas clave bien “RRHH”
        if not extras_only.empty:
            extras_only = extras_only[[
                "Empleado", "DNI", "Tipo",
                "Dias",
                "Extras", "Extras_min",
                "Faltas", "Faltas_min",
                "Saldo", "Saldo_min",
                "Cumplimiento",
                "Incompletos"
            ]].rename(columns={
                "Extras_min": "Extras (min)",
                "Faltas_min": "Faltas (min)",
            })

        kpis_general = {
            "Empleados": empleados,
            "Dias_total": dias,
            "Total_HHMM": minutes_to_hhmm(total_min),
            "Prom_dia_HHMM": minutes_to_hhmm(prom_dia),
            "Incompletos": incompletos,
            "Cortes": cortes,
            "NO_Docente_Total_HHMM": minutes_to_hhmm(nod_sum),
            "NO_Docente_Cumplimiento": nod_pct,
            "NO_Docente_Extras_HHMM": minutes_to_hhmm(nod_extras),
            "NO_Docente_Faltas_HHMM": minutes_to_hhmm(nod_faltas),
            "Correcciones_masivas": fixes_total,
        }

        general_xlsx = export_general_excel(
            reduced=reduced,
            expected=expected,
            kpis_general=kpis_general,
            summary_all=summary.copy(),
            extras_only=extras_only.copy() if isinstance(extras_only, pd.DataFrame) else pd.DataFrame(),
            daily=daily.copy(),
            raw=raw.copy(),
        )

        st.download_button(
            "Exportar Resumen General (Excel)",
            data=general_xlsx,
            file_name="resumen_general_asistencia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        # Gráficos generales
        if not daily.empty:
            by_day = daily.groupby("Fecha", as_index=False).agg(Minutos=("Minutos", "sum"))
            by_day["Horas"] = by_day["Minutos"] / 60.0
            by_day = by_day.sort_values("Fecha")

            st.markdown("### Horas totales por día")
            st.line_chart(by_day.set_index("Fecha")[["Horas"]], height=220)

            # Por tipo
            by_day_tipo = daily.groupby(["Fecha", "Tipo"], as_index=False).agg(Minutos=("Minutos", "sum"))
            pivot = by_day_tipo.pivot(index="Fecha", columns="Tipo", values="Minutos").fillna(0) / 60.0
            st.markdown("### Horas por día (NO Docente vs Docente)")
            st.area_chart(pivot, height=220)

            # Histograma
            st.markdown("### Distribución (horas por día)")
            bins = [(0, 2), (2, 4), (4, 6), (6, 8), (8, 10), (10, 999)]
            hist = histogram_hours(daily["Minutos"], bins)
            hist = hist.set_index("Rango")
            st.bar_chart(hist[["Días"]], height=220)

            # Top 15 horas por empleado
            st.markdown("### Top 15 empleados por horas")
            top_hours = (
                daily.groupby(["Empleado", "DNI", "Tipo"], as_index=False)
                .agg(Total_min=("Minutos", "sum"))
                .sort_values("Total_min", ascending=False)
                .head(15)
            )
            top_hours["Total_horas"] = top_hours["Total_min"] / 60.0
            top_hours = top_hours.set_index("Empleado")
            st.bar_chart(top_hours[["Total_horas"]], height=260)

            # Incompletos por día
            st.markdown("### Incompletos por día")
            inc_day = (
                daily.assign(Incomp=(daily["Incompleto"] == "SI").astype(int))
                .groupby("Fecha", as_index=False)
                .agg(Incompletos=("Incomp", "sum"))
            )
            inc_day = inc_day.set_index("Fecha")
            st.bar_chart(inc_day[["Incompletos"]], height=220)

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        # Rankings / Tablas
        cL, cR = st.columns(2)

        with cL:
            st.markdown("### Top 10 · más incompletos")
            top_inc = (
                daily.groupby(["Empleado", "DNI", "Tipo"], as_index=False)
                .agg(Incompletos=("Incompleto", lambda x: int((x == "SI").sum())))
                .sort_values("Incompletos", ascending=False)
                .head(10)
            )
            st.dataframe(top_inc, use_container_width=True, height=320)

        with cR:
            st.markdown("### Top 10 · más cortes")
            top_cut = (
                daily.groupby(["Empleado", "DNI", "Tipo"], as_index=False)
                .agg(Cortes=("Cortes", lambda x: int((x == "SI").sum())))
                .sort_values("Cortes", ascending=False)
                .head(10)
            )
            st.dataframe(top_cut, use_container_width=True, height=320)

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
        st.markdown("### Resumen completo (incluye Extras)")
        st.dataframe(summary, use_container_width=True, height=440)

        # cache para Empleado
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

        raw_emp = raw[(raw["Empleado"] == emp) & (raw["DNI"] == dni) & (raw["Tipo"] == tipo)].copy().sort_values("FechaHora")

        # botón corrección (NO Docente)
        fix = False
        fixes_applied = 0
        if tipo == "NO Docente":
            fix = st.button("Corregir falta de marcación", use_container_width=True)

        if tipo == "NO Docente" and fix:
            corrected_raw_emp, fixes_applied = correct_missing_punches_for_employee(raw_emp, expected)

            raw_corrected = raw.copy()
            mask = (raw_corrected["Empleado"] == emp) & (raw_corrected["DNI"] == dni) & (raw_corrected["Tipo"] == tipo)
            raw_corrected = raw_corrected[~mask]
            raw_corrected = pd.concat([raw_corrected, corrected_raw_emp], ignore_index=True)
            raw_corrected = raw_corrected.sort_values(["Empleado", "DNI", "FechaHora"]).reset_index(drop=True)

            daily_corrected = calc_daily(raw_corrected, expected)
            daily_emp = daily_corrected[(daily_corrected["Empleado"] == emp) & (daily_corrected["DNI"] == dni) & (daily_corrected["Tipo"] == tipo)].copy()
            raw_emp = corrected_raw_emp
        else:
            daily_emp = daily[(daily["Empleado"] == emp) & (daily["DNI"] == dni) & (daily["Tipo"] == tipo)].copy()

        if daily_emp.empty:
            st.warning("Sin datos para este empleado.")
            return

        # KPIs empleado
        total_min = int(daily_emp["Minutos"].sum())
        dias = int(daily_emp["Fecha"].nunique())
        prom = int(round(daily_emp["Minutos"].mean())) if dias else 0

        inc = int((daily_emp["Incompleto"] == "SI").sum())
        cuts = int((daily_emp["Cortes"] == "SI").sum())
        marc_total = int(daily_emp["Marcaciones"].sum())
        marc_prom = marc_total / max(int(daily_emp.shape[0]), 1)

        pares_0 = int((daily_emp["Pares_estimados"] == 0).sum())

        max_day = int(daily_emp["Minutos"].max())
        min_day = int(daily_emp["Minutos"].min())

        exp_sum = int(daily_emp["Esperado_min"].sum())
        saldo_sum = int(daily_emp["Saldo_min"].sum())

        ok_days = int((daily_emp["Cumple"] == "OK").sum())
        falta_days = int((daily_emp["Cumple"] == "FALTA").sum())
        incom_days = int((daily_emp["Cumple"] == "INCOMPLETO").sum())

        extras_min = 0
        faltas_min = 0
        pct = ""
        if tipo == "NO Docente":
            extras_min = int(daily_emp.loc[daily_emp["Saldo_min"] > 0, "Saldo_min"].sum())
            faltas_min = int((-daily_emp.loc[daily_emp["Saldo_min"] < 0, "Saldo_min"].sum()))
            pct = f"{(total_min/exp_sum*100):.0f}%" if exp_sum > 0 else ""

        # KPIs en burbujas
        r1 = st.columns(4)
        with r1[0]: kpi_card(emp, minutes_to_hhmm(total_min), f"{tipo} · DNI {dni}")
        with r1[1]: kpi_card("Días", f"{dias}", f"Prom/día: {minutes_to_hhmm(prom)}")
        with r1[2]: kpi_card("Marcaciones", f"{marc_total}", f"Prom/día: {marc_prom:.2f} · Pares=0: {pares_0}")
        with r1[3]:
            if tipo == "NO Docente":
                kpi_card("Extras del mes", minutes_to_hhmm(extras_min), f"Faltas: {minutes_to_hhmm(faltas_min)} · Cumpl: {pct}")
            else:
                kpi_card("Total (Docente)", minutes_to_hhmm(total_min), "Por tramos estimados")

        r2 = st.columns(4)
        with r2[0]: kpi_card("Incompletos", f"{inc}", f"Cortes: {cuts}")
        with r2[1]: kpi_card("Máx / Mín día", minutes_to_hhmm(max_day), f"Mín: {minutes_to_hhmm(min_day)}")
        with r2[2]:
            if tipo == "NO Docente":
                kpi_card("Saldo neto", delta_short(saldo_sum), "Acumulado del mes")
            else:
                kpi_card("Cortes", f"{cuts}", "Días con varios tramos")
        with r2[3]:
            if tipo == "NO Docente":
                kpi_card("Días OK/FALTA/INC", f"{ok_days}/{falta_days}/{incom_days}", "Cumplimiento diario")
            else:
                kpi_card("Alertas", f"{inc}", "Días sin pares estimados")

        if fixes_applied:
            st.markdown(f"""<div class="pill">Correcciones aplicadas: {fixes_applied}</div>""", unsafe_allow_html=True)

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        # Gráficos empleado
        ch = daily_emp.sort_values("Fecha").copy()
        ch["Horas_float"] = ch["Minutos"] / 60.0

        st.markdown("### Horas por día (empleado)")
        st.bar_chart(ch.set_index("Fecha")[["Horas_float"]], height=240)

        if tipo == "NO Docente":
            st.markdown("### Saldo por día (NO Docente)")
            saldo_df = ch[["Fecha", "Saldo_min"]].copy()
            saldo_df["Saldo_horas"] = saldo_df["Saldo_min"] / 60.0
            st.bar_chart(saldo_df.set_index("Fecha")[["Saldo_horas"]], height=220)

        st.markdown("### Marcaciones por día")
        marc_df = ch[["Fecha", "Marcaciones"]].set_index("Fecha")
        st.bar_chart(marc_df[["Marcaciones"]], height=220)

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        # Tabla día a día
        st.dataframe(employee_detail_table(daily_emp), use_container_width=True, height=520)

        # Export empleado
        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        kpi_block = {
            "Empleado": emp,
            "DNI": dni,
            "Tipo": tipo,
            "Horario_reducido": "SI" if reduced else "NO",
            "Dias": dias,
            "Total": minutes_to_hhmm(total_min),
            "Prom_dia": minutes_to_hhmm(prom),
            "Marcaciones_total": marc_total,
            "Marcaciones_prom_dia": round(marc_prom, 2),
            "Incompletos": inc,
            "Cortes": cuts,
            "Max_dia": minutes_to_hhmm(max_day),
            "Min_dia": minutes_to_hhmm(min_day),
            "Cumplimiento_NO_Docente": pct if tipo == "NO Docente" else "",
            "Extras_mes_NO_Docente": minutes_to_hhmm(extras_min) if tipo == "NO Docente" else "",
            "Faltas_mes_NO_Docente": minutes_to_hhmm(faltas_min) if tipo == "NO Docente" else "",
            "Saldo_neto_NO_Docente": delta_short(saldo_sum) if tipo == "NO Docente" else "",
            "Dias_OK_FALTA_INCOMP": f"{ok_days}/{falta_days}/{incom_days}" if tipo == "NO Docente" else "",
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
            height=560,
            hide_index=True,
            column_config={
                "Tipo": st.column_config.SelectboxColumn("Tipo", options=["NO Docente", "Docente"], required=True)
            },
        )
        st.session_state["profiles"] = edited


if __name__ == "__main__":
    main()
