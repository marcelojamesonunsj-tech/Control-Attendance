from __future__ import annotations

import io
import hashlib
from dataclasses import dataclass
from typing import Dict, Tuple, List, Optional, Set

import pandas as pd
import streamlit as st

# Excel export styling
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo


# =========================
# Config
# =========================
APP_TITLE = "Asistencia RRHH"
EXPECTED_NO_REDUCED_MIN = 7 * 60
EXPECTED_REDUCED_MIN = 6 * 60

# Para evitar "correcciones" dobles cuando se recalcula
# clave: (dni, fecha_iso)
CorrectionKey = Tuple[str, str]  # ("12345678", "2026-01-15")

# Para evitar duplicados al importar
# clave: (dni, fecha_hora_iso)
MarkKey = Tuple[str, str]  # ("12345678", "2026-01-15T08:03:00")


@dataclass
class ParseResult:
    marks: pd.DataFrame  # columnas: dni, empleado, fecha_hora, fecha, tipo, is_synthetic
    dropped_rows: int
    invalid_dates: int
    empty_dni: int


# =========================
# UI / CSS
# =========================
def inject_css():
    st.markdown(
        """
        <style>
          :root {
            --bg: #0e1117;
            --panel: rgba(255,255,255,0.06);
            --panel2: rgba(255,255,255,0.08);
            --text: rgba(255,255,255,0.92);
            --muted: rgba(255,255,255,0.65);
            --accent: #2f7cf6;
            --good: #19c37d;
            --bad: #ff4d4f;
            --warn: #f5a623;
          }

          .stApp { background: var(--bg); color: var(--text); }
          h1, h2, h3, h4 { color: var(--text); }

          /* Bubble KPI cards */
          .kpi-wrap { display: flex; gap: 10px; flex-wrap: wrap; margin: 6px 0 12px 0; }
          .kpi {
            background: var(--panel);
            border: 1px solid rgba(255,255,255,0.10);
            border-radius: 999px;
            padding: 10px 14px;
            min-width: 170px;
            display: flex;
            flex-direction: column;
            justify-content: center;
          }
          .kpi .label { font-size: 12px; color: var(--muted); line-height: 1.2; }
          .kpi .value { font-size: 18px; font-weight: 700; line-height: 1.25; }
          .kpi .delta { font-size: 12px; color: var(--muted); margin-top: 2px; }

          /* Cleaner tabs + separators */
          .soft-panel {
            background: var(--panel);
            border: 1px solid rgba(255,255,255,0.10);
            border-radius: 14px;
            padding: 12px 12px 10px 12px;
          }

          /* Buttons slightly nicer */
          .stButton button {
            border-radius: 10px !important;
            border: 1px solid rgba(255,255,255,0.16) !important;
            background: rgba(255,255,255,0.08) !important;
            color: var(--text) !important;
          }
          .stButton button:hover { background: rgba(255,255,255,0.12) !important; }

          /* Hide dataframe toolbar if any (best effort) */
          div[data-testid="stToolbar"] { visibility: hidden; height: 0px; }

          /* Make selectboxes/pickers fit dark */
          .stSelectbox, .stMultiSelect, .stDateInput, .stTextInput {
            color: var(--text) !important;
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


def bubble_kpis(items: List[Tuple[str, str, str]]):
    """
    items: list of (label, value, delta_text)
    """
    parts = ['<div class="kpi-wrap">']
    for label, value, delta in items:
        parts.append(
            f"""
            <div class="kpi">
              <div class="label">{label}</div>
              <div class="value">{value}</div>
              <div class="delta">{delta}</div>
            </div>
            """
        )
    parts.append("</div>")
    st.markdown("".join(parts), unsafe_allow_html=True)


# =========================
# Helpers
# =========================
def fmt_hhmm_from_minutes(total_min: int) -> str:
    if total_min < 0:
        sign = "-"
        total_min = abs(total_min)
    else:
        sign = ""
    h = total_min // 60
    m = total_min % 60
    return f"{sign}{h:02d}:{m:02d}"


def safe_int_minutes_from_timedelta(td: pd.Timedelta) -> int:
    if pd.isna(td):
        return 0
    return int(round(td.total_seconds() / 60.0))


def clean_dni(value) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    digits = "".join(ch for ch in s if ch.isdigit())
    return digits


def stable_file_hash(file_bytes: bytes) -> str:
    return hashlib.sha256(file_bytes).hexdigest()


def read_excel_any(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """
    Soporta XLSX, XLSM, XLS (97-2003).
    No pide mapeo: espera columnas exactas: Nombre, Marc., Estado, NvoEstado.
    """
    name_lower = (filename or "").lower().strip()

    # Intento 1: pandas con engine por extensión
    if name_lower.endswith(".xls"):
        # Requiere xlrd (solo xls)
        return pd.read_excel(io.BytesIO(file_bytes), engine="xlrd")
    else:
        # xlsx/xlsm
        return pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")


def normalize_input(df: pd.DataFrame) -> pd.DataFrame:
    """
    Asegura nombres de columnas y deja solo las esperadas.
    """
    # Limpieza mínima por si hay espacios
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    required = ["Nombre", "Marc.", "Estado"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"El Excel no tiene las columnas requeridas: {missing}. Deben ser exactamente: Nombre, Marc., Estado, NvoEstado.")

    # Dejamos las relevantes (NvoEstado puede existir pero no se usa)
    keep = ["Nombre", "Marc.", "Estado"]
    if "NvoEstado" in df.columns:
        keep.append("NvoEstado")
    df = df[keep]

    return df


def parse_marks_from_excel(file_bytes: bytes, filename: str, profiles: Dict[str, str]) -> ParseResult:
    raw = read_excel_any(file_bytes, filename)
    raw = normalize_input(raw)

    # DNI limpio
    raw["dni"] = raw["Nombre"].apply(clean_dni)
    empty_dni = int((raw["dni"] == "").sum())
    raw = raw[raw["dni"] != ""].copy()

    # Empleado (Apellido, Nombre)
    raw["empleado"] = raw["Marc."].astype(str).str.strip()

    # FechaHora desde Estado (dd/mm/yyyy HH:MM) con dayfirst
    invalid_dates_before = len(raw)
    raw["fecha_hora"] = pd.to_datetime(raw["Estado"], errors="coerce", dayfirst=True)
    invalid_dates = int(raw["fecha_hora"].isna().sum())
    raw = raw.dropna(subset=["fecha_hora"]).copy()
    dropped_rows = (invalid_dates_before - len(raw))

    raw["fecha"] = raw["fecha_hora"].dt.date.astype(str)  # "YYYY-MM-DD"

    # Tipo desde perfiles (default NO Docente)
    raw["tipo"] = raw["dni"].map(profiles).fillna("NO Docente")
    raw["is_synthetic"] = False

    # Orden
    raw = raw.sort_values(["dni", "fecha_hora"]).reset_index(drop=True)

    marks = raw[["dni", "empleado", "fecha_hora", "fecha", "tipo", "is_synthetic"]].copy()
    return ParseResult(marks=marks, dropped_rows=dropped_rows, invalid_dates=invalid_dates, empty_dni=empty_dni)


def upsert_profiles_from_marks(marks: pd.DataFrame, profiles: Dict[str, str]) -> Dict[str, str]:
    """
    Agrega DNIs nuevos al diccionario de perfiles (default NO Docente).
    """
    out = dict(profiles)
    for dni in marks["dni"].unique().tolist():
        out.setdefault(dni, "NO Docente")
    return out


def merge_marks_keep_unique(existing: pd.DataFrame, new_marks: pd.DataFrame) -> pd.DataFrame:
    """
    Evita duplicados si se carga el mismo Excel 2 veces:
    Unicidad por (dni, fecha_hora).
    """
    if existing is None or existing.empty:
        merged = new_marks.copy()
    else:
        merged = pd.concat([existing, new_marks], ignore_index=True)

    merged["__k"] = merged["dni"].astype(str) + "||" + merged["fecha_hora"].astype("datetime64[ns]").dt.strftime("%Y-%m-%dT%H:%M:%S")
    merged = merged.drop_duplicates(subset=["__k"], keep="first").drop(columns=["__k"])
    merged = merged.sort_values(["dni", "fecha_hora"]).reset_index(drop=True)
    return merged


def expected_minutes(tipo: str, reduced: bool, apply_to_docente: bool) -> Optional[int]:
    if tipo == "NO Docente":
        return EXPECTED_REDUCED_MIN if reduced else EXPECTED_NO_REDUCED_MIN
    if tipo == "Docente":
        if apply_to_docente:
            return EXPECTED_REDUCED_MIN if reduced else EXPECTED_NO_REDUCED_MIN
        return None
    return None


def compute_daily_aggregates(
    marks: pd.DataFrame,
    reduced: bool,
    apply_expected_to_docente: bool,
    corrections_applied: Set[CorrectionKey],
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Devuelve:
      - daily: por dni+fecha con métricas (trabajado_min, esperado_min, extras_min, saldo_min, marcaciones, incompleto, cortes)
      - marks_out: marks con potenciales marcas sintéticas ya incluidas (si correcciones_applied contiene la key)
    """
    if marks is None or marks.empty:
        daily_cols = [
            "dni", "empleado", "tipo", "fecha",
            "primera", "ultima",
            "marcaciones",
            "trabajado_min", "esperado_min", "extras_min", "saldo_min",
            "incompleto", "cortes", "faltas_corregidas"
        ]
        return pd.DataFrame(columns=daily_cols), marks.copy()

    marks_work = marks.copy()
    marks_work["fecha_hora"] = pd.to_datetime(marks_work["fecha_hora"], errors="coerce")
    marks_work = marks_work.dropna(subset=["fecha_hora"]).copy()
    marks_work["fecha"] = marks_work["fecha_hora"].dt.date.astype(str)
    marks_work = marks_work.sort_values(["dni", "fecha_hora"]).reset_index(drop=True)

    synthetic_rows = []

    # Corrección automática (NO Docente): si el día tiene 1 marcación, agregar otra a +6h/+7h
    # SOLO si la key está en corrections_applied
    grouped = marks_work.groupby(["dni", "fecha"], sort=False)
    for (dni, fecha), g in grouped:
        tipo = (g["tipo"].iloc[0] if len(g) else "NO Docente)
        if tipo != "NO Docente":
            continue
        if len(g) != 1:
            continue

        key: CorrectionKey = (dni, fecha)
        if key not in corrections_applied:
            continue

        exp = expected_minutes("NO Docente", reduced, apply_expected_to_docente) or (EXPECTED_REDUCED_MIN if reduced else EXPECTED_NO_REDUCED_MIN)
        t0 = pd.to_datetime(g["fecha_hora"].iloc[0])
        t1 = t0 + pd.Timedelta(minutes=int(exp))

        synthetic_rows.append(
            {
                "dni": dni,
                "empleado": g["empleado"].iloc[0],
                "fecha_hora": t1,
                "fecha": fecha,
                "tipo": "NO Docente",
                "is_synthetic": True,
            }
        )

    if synthetic_rows:
        synth_df = pd.DataFrame(synthetic_rows)
        marks_work = pd.concat([marks_work, synth_df], ignore_index=True)
        marks_work = marks_work.sort_values(["dni", "fecha_hora"]).reset_index(drop=True)

    # Daily aggregates
    daily_rows = []
    grouped2 = marks_work.groupby(["dni", "fecha"], sort=False)
    for (dni, fecha), g in grouped2:
        g = g.sort_values("fecha_hora")
        empleado = g["empleado"].iloc[0]
        tipo = g["tipo"].iloc[0]
        n = int(len(g))
        primera = pd.to_datetime(g["fecha_hora"].iloc[0])
        ultima = pd.to_datetime(g["fecha_hora"].iloc[-1])

        faltas_corregidas = 0
        key = (dni, fecha)
        if tipo == "NO Docente" and (key in corrections_applied) and (n == 2) and (g["is_synthetic"].any()):
            faltas_corregidas = 1

        # trabajado
        trabajado_min = 0
        incompleto = False
        cortes = 0

        if n < 2:
            incompleto = True
            trabajado_min = 0
            cortes = 0
        else:
            if tipo == "NO Docente":
                # span
                trabajado_min = safe_int_minutes_from_timedelta(ultima - primera)
                incompleto = False
                # cortes: "interrupciones" inferidas por múltiples marcas
                # si hay más de 2 marcas, consideramos cortes = (n - 2)
                cortes = max(0, n - 2)
            else:
                # Docente: tramos alternados (t0-t1, t2-t3, ...)
                ts = g["fecha_hora"].tolist()
                pairs = (n // 2)
                total = 0
                for i in range(pairs):
                    t_in = pd.to_datetime(ts[2 * i])
                    t_out = pd.to_datetime(ts[2 * i + 1])
                    if t_out >= t_in:
                        total += safe_int_minutes_from_timedelta(t_out - t_in)
                trabajado_min = int(total)
                incompleto = (n % 2 == 1)  # queda impar => último tramo no cuenta
                # cortes: cantidad de tramos - 1
                cortes = max(0, pairs - 1)

        exp = expected_minutes(tipo, reduced, apply_expected_to_docente)
        esperado_min = int(exp) if exp is not None else None

        extras_min = 0
        saldo_min = 0
        if esperado_min is not None:
            extras_min = max(0, trabajado_min - esperado_min)
            saldo_min = trabajado_min - esperado_min

        daily_rows.append(
            {
                "dni": dni,
                "empleado": empleado,
                "tipo": tipo,
                "fecha": fecha,
                "primera": primera,
                "ultima": ultima,
                "marcaciones": n,
                "trabajado_min": trabajado_min,
                "esperado_min": esperado_min if esperado_min is not None else None,
                "extras_min": extras_min,
                "saldo_min": saldo_min if esperado_min is not None else None,
                "incompleto": bool(incompleto),
                "cortes": int(cortes),
                "faltas_corregidas": int(faltas_corregidas),
            }
        )

    daily = pd.DataFrame(daily_rows)
    if not daily.empty:
        daily = daily.sort_values(["fecha", "empleado"]).reset_index(drop=True)

        # columnas formato
        daily["primera"] = pd.to_datetime(daily["primera"]).dt.strftime("%H:%M")
        daily["ultima"] = pd.to_datetime(daily["ultima"]).dt.strftime("%H:%M")
        daily["horas"] = daily["trabajado_min"].apply(lambda x: fmt_hhmm_from_minutes(int(x)))
        daily["esperado"] = daily["esperado_min"].apply(lambda x: "" if pd.isna(x) else fmt_hhmm_from_minutes(int(x)))
        daily["extras"] = daily["extras_min"].apply(lambda x: fmt_hhmm_from_minutes(int(x)))
        daily["saldo"] = daily["saldo_min"].apply(lambda x: "" if pd.isna(x) else fmt_hhmm_from_minutes(int(x)))

    return daily, marks_work


def compute_employee_summary(daily: pd.DataFrame) -> pd.DataFrame:
    """
    Resumen por empleado (dni).
    """
    if daily is None or daily.empty:
        cols = ["empleado", "dni", "tipo", "dias", "marcaciones", "incompletos", "cortes", "total_horas", "total_extras", "extras_min"]
        return pd.DataFrame(columns=cols)

    agg = daily.groupby(["dni", "empleado", "tipo"], as_index=False).agg(
        dias=("fecha", "nunique"),
        marcaciones=("marcaciones", "sum"),
        incompletos=("incompleto", "sum"),
        cortes=("cortes", "sum"),
        trabajado_min=("trabajado_min", "sum"),
        extras_min=("extras_min", "sum"),
        faltas_corregidas=("faltas_corregidas", "sum"),
    )

    agg["total_horas"] = agg["trabajado_min"].apply(lambda x: fmt_hhmm_from_minutes(int(x)))
    agg["total_extras"] = agg["extras_min"].apply(lambda x: fmt_hhmm_from_minutes(int(x)))
    agg = agg.rename(columns={"dni": "DNI", "empleado": "Empleado", "tipo": "Tipo"})
    agg = agg[["Empleado", "DNI", "Tipo", "dias", "marcaciones", "incompletos", "cortes", "faltas_corregidas", "total_horas", "total_extras", "extras_min"]]
    return agg.sort_values(["extras_min", "total_horas"], ascending=[False, False]).reset_index(drop=True)


# =========================
# Excel export (bonito)
# =========================
def _apply_sheet_style(ws):
    ws.sheet_view.showGridLines = False


def _header_style():
    return Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E79"), Alignment(vertical="center")


def _thin_border():
    side = Side(style="thin", color="2B2B2B")
    return Border(left=side, right=side, top=side, bottom=side)


def write_table(ws, df: pd.DataFrame, table_name: str, freeze_pane: str = "A2", dni_cols: Optional[List[str]] = None):
    if df is None:
        df = pd.DataFrame()

    dni_cols = dni_cols or []

    # Escribo el DF
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
        ws.append(row)

    # Estilos
    header_font, header_fill, header_align = _header_style()
    border = _thin_border()

    max_row = ws.max_row
    max_col = ws.max_column

    # Freeze
    ws.freeze_panes = freeze_pane

    # Column widths
    for col in range(1, max_col + 1):
        letter = get_column_letter(col)
        # medir
        max_len = 0
        for cell in ws[letter]:
            if cell.value is None:
                continue
            max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[letter].width = min(max(10, max_len + 2), 45)

    # Header
    for c in range(1, max_col + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align

    # Body
    for r in range(2, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=False)

    # DNI como texto
    if dni_cols:
        col_map = {ws.cell(row=1, column=c).value: c for c in range(1, max_col + 1)}
        for name in dni_cols:
            if name in col_map:
                c_idx = col_map[name]
                for r in range(2, max_row + 1):
                    ws.cell(row=r, column=c_idx).number_format = "@"

    # Excel Table (autofilter + estilo)
    if max_row >= 1 and max_col >= 1:
        ref = f"A1:{get_column_letter(max_col)}{max_row}"
        table = Table(displayName=table_name, ref=ref)
        style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        table.tableStyleInfo = style
        ws.add_table(table)

    _apply_sheet_style(ws)


def workbook_to_bytes(wb: Workbook) -> bytes:
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def export_general_xlsx(kpis_general: pd.DataFrame, resumen_empleados: pd.DataFrame, solo_extras: pd.DataFrame,
                        detalle_diario: pd.DataFrame, marcaciones: pd.DataFrame) -> bytes:
    wb = Workbook()
    # Remove default
    wb.remove(wb.active)

    ws1 = wb.create_sheet("KPIs_General")
    write_table(ws1, kpis_general, "KPIsGeneral", freeze_pane="A2")

    ws2 = wb.create_sheet("Resumen_Empleados")
    write_table(ws2, resumen_empleados, "ResumenEmpleados", freeze_pane="A2", dni_cols=["DNI"])

    ws3 = wb.create_sheet("Solo_Extras")
    write_table(ws3, solo_extras, "SoloExtras", freeze_pane="A2", dni_cols=["DNI"])

    ws4 = wb.create_sheet("Detalle_Diario")
    write_table(ws4, detalle_diario, "DetalleDiario", freeze_pane="A2", dni_cols=["dni", "DNI"])

    ws5 = wb.create_sheet("Marcaciones")
    write_table(ws5, marcaciones, "Marcaciones", freeze_pane="A2", dni_cols=["dni", "DNI"])

    return workbook_to_bytes(wb)


def export_solo_extras_xlsx(solo_extras: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Solo_Extras"
    write_table(ws, solo_extras, "SoloExtras", freeze_pane="A2", dni_cols=["DNI"])
    return workbook_to_bytes(wb)


def export_empleado_xlsx(kpis_emp: pd.DataFrame, detalle: pd.DataFrame, marcaciones: pd.DataFrame) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    ws1 = wb.create_sheet("KPIs")
    write_table(ws1, kpis_emp, "KPIsEmpleado", freeze_pane="A2", dni_cols=["DNI"])

    ws2 = wb.create_sheet("Detalle_Diario")
    write_table(ws2, detalle, "DetalleEmpleado", freeze_pane="A2", dni_cols=["dni", "DNI"])

    ws3 = wb.create_sheet("Marcaciones")
    write_table(ws3, marcaciones, "MarcacionesEmpleado", freeze_pane="A2", dni_cols=["dni", "DNI"])

    return workbook_to_bytes(wb)


# =========================
# App State
# =========================
def ensure_state():
    if "profiles" not in st.session_state:
        st.session_state["profiles"] = {}  # dni -> "Docente"/"NO Docente"
    if "marks" not in st.session_state:
        st.session_state["marks"] = pd.DataFrame(columns=["dni", "empleado", "fecha_hora", "fecha", "tipo", "is_synthetic"])
    if "file_hashes" not in st.session_state:
        st.session_state["file_hashes"] = set()
    if "corrections" not in st.session_state:
        st.session_state["corrections"] = set()  # Set[CorrectionKey]
    if "last_parse_stats" not in st.session_state:
        st.session_state["last_parse_stats"] = None


# =========================
# UI blocks
# =========================
def kpis_general_block(marks: pd.DataFrame, daily: pd.DataFrame, resumen: pd.DataFrame):
    empleados = int(daily["dni"].nunique()) if not daily.empty else 0
    dias = int(daily["fecha"].nunique()) if not daily.empty else 0
    marcaciones = int(marks.shape[0]) if marks is not None else 0
    incompletos = int(daily["incompleto"].sum()) if not daily.empty else 0
    cortes = int(daily["cortes"].sum()) if not daily.empty else 0
    total_min = int(daily["trabajado_min"].sum()) if not daily.empty else 0
    total_extras_min = int(daily["extras_min"].sum()) if (not daily.empty and "extras_min" in daily.columns) else 0
    faltas_corr = int(daily["faltas_corregidas"].sum()) if not daily.empty else 0

    prom_diario_min = int(round(total_min / max(1, dias))) if dias else 0

    bubble_kpis(
        [
            ("Empleados", f"{empleados}", "únicos"),
            ("Días", f"{dias}", "con marcaciones"),
            ("Marcaciones", f"{marcaciones}", "total importadas"),
            ("Incompletos", f"{incompletos}", "días con 1 marca (o impar)"),
            ("Cortes", f"{cortes}", "interrupciones"),
            ("Total horas", f"{fmt_hhmm_from_minutes(total_min)}", "acumulado"),
            ("Promedio diario", f"{fmt_hhmm_from_minutes(prom_diario_min)}", "sobre días"),
            ("Total extras", f"{fmt_hhmm_from_minutes(total_extras_min)}", f"{total_extras_min} min"),
            ("Correcciones", f"{faltas_corr}", "faltas corregidas"),
        ]
    )


def make_kpis_general_df(marks: pd.DataFrame, daily: pd.DataFrame) -> pd.DataFrame:
    empleados = int(daily["dni"].nunique()) if not daily.empty else 0
    dias = int(daily["fecha"].nunique()) if not daily.empty else 0
    marcaciones = int(marks.shape[0]) if marks is not None else 0
    incompletos = int(daily["incompleto"].sum()) if not daily.empty else 0
    cortes = int(daily["cortes"].sum()) if not daily.empty else 0
    total_min = int(daily["trabajado_min"].sum()) if not daily.empty else 0
    total_extras_min = int(daily["extras_min"].sum()) if (not daily.empty and "extras_min" in daily.columns) else 0
    faltas_corr = int(daily["faltas_corregidas"].sum()) if not daily.empty else 0
    prom_diario_min = int(round(total_min / max(1, dias))) if dias else 0

    rows = [
        ("Empleados únicos", empleados),
        ("Días con marcaciones", dias),
        ("Marcaciones importadas", marcaciones),
        ("Días incompletos", incompletos),
        ("Cortes", cortes),
        ("Total horas (HH:MM)", fmt_hhmm_from_minutes(total_min)),
        ("Total horas (min)", total_min),
        ("Promedio diario (HH:MM)", fmt_hhmm_from_minutes(prom_diario_min)),
        ("Total extras (HH:MM)", fmt_hhmm_from_minutes(total_extras_min)),
        ("Total extras (min)", total_extras_min),
        ("Faltas corregidas", faltas_corr),
    ]
    return pd.DataFrame(rows, columns=["KPI", "Valor"])


def charts_general(daily: pd.DataFrame, resumen: pd.DataFrame):
    if daily.empty:
        return

    # Horas por día
    by_day = daily.groupby("fecha", as_index=False).agg(
        trabajado_min=("trabajado_min", "sum"),
        incompletos=("incompleto", "sum"),
        extras_min=("extras_min", "sum"),
    )
    by_day["horas"] = by_day["trabajado_min"].apply(lambda x: x / 60.0)
    by_day["extras_h"] = by_day["extras_min"].apply(lambda x: x / 60.0)

    c1, c2 = st.columns(2)
    with c1:
        st.caption("Horas por día")
        st.line_chart(by_day.set_index("fecha")[["horas"]])
    with c2:
        st.caption("Incompletos por día")
        st.bar_chart(by_day.set_index("fecha")[["incompletos"]])

    c3, c4 = st.columns(2)
    with c3:
        st.caption("Distribución (horas por día)")
        # Hist simple
        hist = by_day["horas"].round(2).value_counts().sort_index()
        st.bar_chart(hist)
    with c4:
        st.caption("Top empleados por extras")
        if resumen.empty:
            st.info("Sin datos.")
        else:
            top = resumen.sort_values("extras_min", ascending=False).head(10)[["Empleado", "extras_min"]].copy()
            top = top.set_index("Empleado")
            st.bar_chart(top)


def build_solo_extras_table(resumen: pd.DataFrame) -> pd.DataFrame:
    if resumen is None or resumen.empty:
        return pd.DataFrame(columns=["Empleado", "DNI", "Tipo", "Horas_extras", "Extras_min"])

    out = resumen.copy()
    out["Horas_extras"] = out["total_extras"]
    out = out.rename(columns={"extras_min": "Extras_min"})
    out = out[["Empleado", "DNI", "Tipo", "Horas_extras", "Extras_min"]]
    out = out.sort_values("Extras_min", ascending=False).reset_index(drop=True)
    return out


def build_detalle_diario_export(daily: pd.DataFrame) -> pd.DataFrame:
    if daily is None or daily.empty:
        return pd.DataFrame()

    cols = ["dni", "empleado", "tipo", "fecha", "primera", "ultima", "horas", "esperado", "extras", "saldo", "marcaciones", "incompleto", "cortes", "faltas_corregidas", "trabajado_min", "extras_min"]
    cols = [c for c in cols if c in daily.columns]
    d = daily[cols].copy()
    d = d.rename(columns={"dni": "DNI", "empleado": "Empleado", "tipo": "Tipo"})
    return d


def build_marcaciones_export(marks: pd.DataFrame) -> pd.DataFrame:
    if marks is None or marks.empty:
        return pd.DataFrame()

    m = marks.copy()
    m["fecha_hora"] = pd.to_datetime(m["fecha_hora"]).dt.strftime("%Y-%m-%d %H:%M")
    m = m.rename(columns={"dni": "DNI", "empleado": "Empleado", "tipo": "Tipo", "fecha_hora": "FechaHora", "is_synthetic": "Es_Sintetica"})
    return m[["DNI", "Empleado", "Tipo", "FechaHora", "fecha", "Es_Sintetica"]]


def employee_kpis(daily_emp: pd.DataFrame, marks_emp: pd.DataFrame, reduced: bool, apply_to_docente: bool) -> Tuple[List[Tuple[str, str, str]], pd.DataFrame]:
    if daily_emp.empty:
        items = [("Total horas", "00:00", ""), ("Días", "0", ""), ("Marcaciones", "0", ""), ("Extras", "00:00", "0 min")]
        return items, pd.DataFrame([("Sin datos", "")], columns=["KPI", "Valor"])

    tipo = daily_emp["tipo"].iloc[0]
    dni = daily_emp["dni"].iloc[0]
    empleado = daily_emp["empleado"].iloc[0]

    dias = int(daily_emp["fecha"].nunique())
    marc = int(marks_emp.shape[0]) if marks_emp is not None else int(daily_emp["marcaciones"].sum())
    incompletos = int(daily_emp["incompleto"].sum())
    cortes = int(daily_emp["cortes"].sum())
    total_min = int(daily_emp["trabajado_min"].sum())
    extras_min = int(daily_emp["extras_min"].sum()) if "extras_min" in daily_emp.columns else 0
    faltas_corr = int(daily_emp["faltas_corregidas"].sum()) if "faltas_corregidas" in daily_emp.columns else 0
    prom_min = int(round(total_min / max(1, dias))) if dias else 0

    exp = expected_minutes(tipo, reduced, apply_to_docente)
    cumplimiento = ""
    if exp is not None and dias:
        esperado_total = int(exp) * dias
        saldo_total = total_min - esperado_total
        cumplimiento = f"Saldo {fmt_hhmm_from_minutes(saldo_total)}"
    else:
        cumplimiento = "Sin esperado"

    items = [
        ("Empleado", empleado, f"DNI {dni}"),
        ("Tipo", tipo, "perfil"),
        ("Total horas", fmt_hhmm_from_minutes(total_min), "acumulado"),
        ("Prom/día", fmt_hhmm_from_minutes(prom_min), "promedio"),
        ("Días", f"{dias}", "con marcación"),
        ("Marcaciones", f"{marc}", "total"),
        ("Incompletos", f"{incompletos}", "días"),
        ("Cortes", f"{cortes}", "total"),
        ("Extras", fmt_hhmm_from_minutes(extras_min), f"{extras_min} min"),
        ("Correcciones", f"{faltas_corr}", "faltas corregidas"),
        ("Cumplimiento", cumplimiento, ""),
    ]

    kpi_df = pd.DataFrame(
        [
            ("Empleado", empleado),
            ("DNI", dni),
            ("Tipo", tipo),
            ("Días", dias),
            ("Marcaciones", marc),
            ("Incompletos", incompletos),
            ("Cortes", cortes),
            ("Total horas (HH:MM)", fmt_hhmm_from_minutes(total_min)),
            ("Total horas (min)", total_min),
            ("Extras (HH:MM)", fmt_hhmm_from_minutes(extras_min)),
            ("Extras (min)", extras_min),
            ("Faltas corregidas", faltas_corr),
            ("Cumplimiento", cumplimiento),
        ],
        columns=["KPI", "Valor"],
    )
    return items, kpi_df


# =========================
# Main App
# =========================
def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    inject_css()
    ensure_state()

    # Header minimal
    colA, colB, colC = st.columns([2.2, 1.2, 1.2])
    with colA:
        st.markdown(f"## {APP_TITLE}")
    with colB:
        reduced = st.toggle("Activar horario reducido", value=False)
    with colC:
        apply_to_docente = st.toggle("Aplicar 6/7 a docentes", value=False)

    # Upload
    with st.container():
        st.markdown('<div class="soft-panel">', unsafe_allow_html=True)
        up_col1, up_col2, up_col3 = st.columns([2.4, 1, 1])
        with up_col1:
            file = st.file_uploader(
                "Subir Excel del reloj (XLSX/XLSM/XLS)",
                type=["xlsx", "xlsm", "xls"],
                accept_multiple_files=False,
                label_visibility="visible",
            )
        with up_col2:
            if st.button("Limpiar todo", use_container_width=True):
                st.session_state["marks"] = pd.DataFrame(columns=["dni", "empleado", "fecha_hora", "fecha", "tipo", "is_synthetic"])
                st.session_state["file_hashes"] = set()
                st.session_state["corrections"] = set()
                st.session_state["last_parse_stats"] = None
                st.toast("Estado limpiado.")
        with up_col3:
            st.caption("Sin login · Simple para Ivana")

        st.markdown("</div>", unsafe_allow_html=True)

    # Parse / merge
    if file is not None:
        file_bytes = file.getvalue()
        h = stable_file_hash(file_bytes)

        # Evitar re-import doble por hash exacto
        if h not in st.session_state["file_hashes"]:
            try:
                pr = parse_marks_from_excel(file_bytes, file.name, st.session_state["profiles"])
                # Upsert perfiles
                st.session_state["profiles"] = upsert_profiles_from_marks(pr.marks, st.session_state["profiles"])
                # Actualizar tipo en marks según perfiles (por si se editó antes)
                pr.marks["tipo"] = pr.marks["dni"].map(st.session_state["profiles"]).fillna("NO Docente")
                st.session_state["marks"] = merge_marks_keep_unique(st.session_state["marks"], pr.marks)
                st.session_state["file_hashes"].add(h)
                st.session_state["last_parse_stats"] = pr
                st.toast("Excel importado.")
            except Exception as e:
                st.error(f"No pude leer el Excel: {e}")

    # Re-sync tipo de marks con perfiles actuales
    if not st.session_state["marks"].empty:
        st.session_state["marks"]["tipo"] = st.session_state["marks"]["dni"].map(st.session_state["profiles"]).fillna("NO Docente")

    # Compute daily + outputs
    daily, marks_with_synth = compute_daily_aggregates(
        st.session_state["marks"],
        reduced=reduced,
        apply_expected_to_docente=apply_to_docente,
        corrections_applied=st.session_state["corrections"],
    )
    resumen = compute_employee_summary(daily)
    solo_extras = build_solo_extras_table(resumen)

    # Tabs
    tab_general, tab_empleado, tab_perfiles = st.tabs(["General", "Empleado", "Perfiles"])

    # ========== General ==========
    with tab_general:
        if st.session_state["last_parse_stats"] is not None:
            pr: ParseResult = st.session_state["last_parse_stats"]
            # ultra corto y útil
            st.caption(
                f"Importación: marcaciones válidas {len(pr.marks)} · DNI vacíos {pr.empty_dni} · fechas inválidas {pr.invalid_dates}"
            )

        kpis_general_block(marks_with_synth, daily, resumen)

        # Botón corrección global (NO Docente con 1 marca)
        c1, c2, c3 = st.columns([1.3, 1.3, 3.4])
        with c1:
            if st.button("Corregir faltas de marcación (TODOS)", use_container_width=True, help="Solo NO Docente: días con 1 marcación => agrega otra a +6h/+7h"):
                # Aplicar correcciones a todas las keys posibles
                if not st.session_state["marks"].empty:
                    m = st.session_state["marks"]
                    tmp = m.copy()
                    tmp["fecha"] = pd.to_datetime(tmp["fecha_hora"]).dt.date.astype(str)
                    # solo NO Docente y días con 1 marca
                    g = tmp[tmp["tipo"] == "NO Docente"].groupby(["dni", "fecha"]).size().reset_index(name="n")
                    one = g[g["n"] == 1]
                    before = len(st.session_state["corrections"])
                    for _, r in one.iterrows():
                        st.session_state["corrections"].add((str(r["dni"]), str(r["fecha"])))
                    after = len(st.session_state["corrections"])
                    st.toast(f"Correcciones marcadas: +{after - before}")
        with c2:
            st.caption(f"Correcciones activas: {len(st.session_state['corrections'])}")
        with c3:
            st.caption("Las correcciones se aplican al recalcular. No se usa Entrada/Salida: solo horas.")

        # Charts
        charts_general(daily, resumen)

        # SOLO EXTRAS
        st.markdown("### SOLO EXTRAS")
        st.table(solo_extras)

        # Export SOLO extras
        ex1, ex2, ex3 = st.columns([1.2, 1.2, 2.6])
        with ex1:
            if not solo_extras.empty:
                bytes_x = export_solo_extras_xlsx(solo_extras)
                st.download_button(
                    "Exportar SOLO extras (Excel)",
                    data=bytes_x,
                    file_name="extras_empleados.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            else:
                st.button("Exportar SOLO extras (Excel)", disabled=True, use_container_width=True)
        with ex2:
            # Export general
            if not daily.empty or not marks_with_synth.empty:
                kpis_df = make_kpis_general_df(marks_with_synth, daily)
                detalle_export = build_detalle_diario_export(daily)
                marc_export = build_marcaciones_export(marks_with_synth)
                bytes_g = export_general_xlsx(
                    kpis_general=kpis_df,
                    resumen_empleados=resumen,
                    solo_extras=solo_extras,
                    detalle_diario=detalle_export,
                    marcaciones=marc_export,
                )
                st.download_button(
                    "Exportar resumen general (Excel)",
                    data=bytes_g,
                    file_name="resumen_general_asistencia.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            else:
                st.button("Exportar resumen general (Excel)", disabled=True, use_container_width=True)

        # tabla de empleados (resumen)
        st.markdown("### Resumen por empleado")
        if resumen.empty:
            st.info("Todavía no hay datos.")
        else:
            # st.table para que sea sin toolbar
            show = resumen.copy()
            show = show.drop(columns=["extras_min"], errors="ignore")
            st.table(show)

    # ========== Empleado ==========
    with tab_empleado:
        if st.session_state["marks"].empty:
            st.info("Subí un Excel para empezar.")
        else:
            # Selector
            options = []
            dni_to_label = {}
            for _, r in resumen.iterrows():
                label = f"{r['Empleado']} · {r['DNI']} · {r['Tipo']}"
                options.append(label)
                dni_to_label[label] = r["DNI"]

            if not options:
                st.info("No hay empleados para mostrar.")
            else:
                selected = st.selectbox("Empleado", options, index=0, label_visibility="collapsed")
                dni_sel = str(dni_to_label[selected])

                daily_emp = daily[daily["dni"].astype(str) == dni_sel].copy()
                marks_emp = marks_with_synth[marks_with_synth["dni"].astype(str) == dni_sel].copy()

                items, kpi_df = employee_kpis(daily_emp, marks_emp, reduced=reduced, apply_to_docente=apply_to_docente)
                bubble_kpis(items)

                # Botón corrección por empleado
                b1, b2, b3 = st.columns([1.2, 1.2, 2.6])
                with b1:
                    if st.button("Corregir falta de marcación", use_container_width=True, help="Solo NO Docente: días con 1 marcación => agrega otra a +6h/+7h"):
                        # marcar todas las keys de este empleado que tengan 1 marca en el día
                        if not st.session_state["marks"].empty:
                            m = st.session_state["marks"]
                            tmp = m[m["dni"].astype(str) == dni_sel].copy()
                            tmp["fecha"] = pd.to_datetime(tmp["fecha_hora"]).dt.date.astype(str)
                            tmp = tmp[tmp["tipo"] == "NO Docente"]
                            g = tmp.groupby(["dni", "fecha"]).size().reset_index(name="n")
                            one = g[g["n"] == 1]
                            before = len(st.session_state["corrections"])
                            for _, rr in one.iterrows():
                                st.session_state["corrections"].add((str(rr["dni"]), str(rr["fecha"])))
                            after = len(st.session_state["corrections"])
                            st.toast(f"Correcciones marcadas: +{after - before}")
                with b2:
                    # Export empleado
                    if not daily_emp.empty or not marks_emp.empty:
                        detalle_emp = build_detalle_diario_export(daily_emp)
                        marc_emp = build_marcaciones_export(marks_emp)
                        bytes_e = export_empleado_xlsx(kpis_emp=kpi_df, detalle=detalle_emp, marcaciones=marc_emp)
                        st.download_button(
                            "Exportar empleado (Excel)",
                            data=bytes_e,
                            file_name=f"empleado_{dni_sel}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                    else:
                        st.button("Exportar empleado (Excel)", disabled=True, use_container_width=True)

                # Tabla día a día
                st.markdown("### Detalle diario")
                if daily_emp.empty:
                    st.info("Sin detalle diario.")
                else:
                    # Tabla limpia
                    view_cols = [
                        "fecha", "primera", "ultima", "horas", "esperado", "extras", "saldo",
                        "marcaciones", "incompleto", "cortes", "faltas_corregidas",
                        "trabajado_min", "extras_min"
                    ]
                    view_cols = [c for c in view_cols if c in daily_emp.columns]
                    st.table(daily_emp[view_cols].rename(columns={"fecha": "Fecha"}))

                    # Gráficos (para NO Docente, extras por día; para Docente solo horas)
                    st.markdown("### Gráficos")
                    by_day = daily_emp.groupby("fecha", as_index=False).agg(
                        trabajado_min=("trabajado_min", "sum"),
                        extras_min=("extras_min", "sum"),
                    )
                    by_day["horas"] = by_day["trabajado_min"].apply(lambda x: x / 60.0)
                    by_day["extras_h"] = by_day["extras_min"].apply(lambda x: x / 60.0)

                    g1, g2 = st.columns(2)
                    with g1:
                        st.caption("Horas por día")
                        st.line_chart(by_day.set_index("fecha")[["horas"]])
                    with g2:
                        if (daily_emp["tipo"].iloc[0] == "NO Docente") or apply_to_docente:
                            st.caption("Extras por día")
                            st.bar_chart(by_day.set_index("fecha")[["extras_h"]])
                        else:
                            st.caption("Extras por día")
                            st.info("Extras no aplican a Docente (configurable arriba).")

                # Marcaciones
                st.markdown("### Marcaciones")
                if marks_emp.empty:
                    st.info("Sin marcaciones.")
                else:
                    mm = marks_emp.copy()
                    mm["FechaHora"] = pd.to_datetime(mm["fecha_hora"]).dt.strftime("%Y-%m-%d %H:%M")
                    mm = mm[["FechaHora", "fecha", "is_synthetic"]].rename(columns={"fecha": "Fecha", "is_synthetic": "Sintética"})
                    st.table(mm)

    # ========== Perfiles ==========
    with tab_perfiles:
        st.markdown("### Perfiles")
        # Tabla editable: DNI, Empleado (último conocido), Tipo
        if st.session_state["marks"].empty and not st.session_state["profiles"]:
            st.info("Subí un Excel para generar la lista de perfiles.")
        else:
            # Armar base
            last_name = {}
            if not st.session_state["marks"].empty:
                tmp = st.session_state["marks"].sort_values("fecha_hora").groupby("dni").tail(1)
                for _, r in tmp.iterrows():
                    last_name[str(r["dni"])] = str(r["empleado"])

            rows = []
            for dni, tipo in sorted(st.session_state["profiles"].items(), key=lambda x: x[0]):
                rows.append({"DNI": dni, "Empleado": last_name.get(dni, ""), "Tipo": tipo})

            prof_df = pd.DataFrame(rows)
            edited = st.data_editor(
                prof_df,
                use_container_width=True,
                num_rows="fixed",
                column_config={
                    "DNI": st.column_config.TextColumn(disabled=True),
                    "Empleado": st.column_config.TextColumn(disabled=True),
                    "Tipo": st.column_config.SelectboxColumn(options=["NO Docente", "Docente"]),
                },
                hide_index=True,
            )

            # Guardar cambios
            new_profiles = dict(st.session_state["profiles"])
            for _, r in edited.iterrows():
                dni = str(r["DNI"])
                tipo = str(r["Tipo"])
                if tipo not in ("NO Docente", "Docente"):
                    tipo = "NO Docente"
                new_profiles[dni] = tipo
            st.session_state["profiles"] = new_profiles

            # Re-sync tipo en marks
            if not st.session_state["marks"].empty:
                st.session_state["marks"]["tipo"] = st.session_state["marks"]["dni"].map(st.session_state["profiles"]).fillna("NO Docente")

            st.caption("Por defecto: NO Docente. Cambios se guardan en session_state.")


if __name__ == "__main__":
    main()
