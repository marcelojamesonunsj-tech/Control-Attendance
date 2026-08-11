from __future__ import annotations

import io
import re
import json
import html
import unicodedata
import calendar
from collections import defaultdict
from datetime import date

import pandas as pd
import streamlit as st

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.pagebreak import Break


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================
REQUIRED_COLS = ["Estado"]
OPTIONAL_COLS = ["Nombre", "Marc.", "NvoEstado"]

# Toda marcación entre 00:00 y 03:59 se considera salida del día laboral anterior.
# Esto evita que una salida a la 01:30 o 02:00 quede como un día nuevo.
WORKDAY_CUTOFF_HOUR = 4

# Dos marcas consecutivas dentro de este margen se consideran el mismo evento biométrico.
DUPLICATE_PUNCH_WINDOW_MINUTES = 2

MONTH_NAMES = {
    1: "ENERO",
    2: "FEBRERO",
    3: "MARZO",
    4: "ABRIL",
    5: "MAYO",
    6: "JUNIO",
    7: "JULIO",
    8: "AGOSTO",
    9: "SEPTIEMBRE",
    10: "OCTUBRE",
    11: "NOVIEMBRE",
    12: "DICIEMBRE",
}


# ============================================================
# UI
# ============================================================
def inject_css() -> None:
    st.markdown(
        """
        <style>
        :root{
            --unsj-blue-1:#002B63;
            --unsj-blue-2:#0A58CA;
            --unsj-blue-3:#79B8FF;
            --glass-1: rgba(255,255,255,.10);
            --glass-2: rgba(255,255,255,.05);
            --glass-br: rgba(255,255,255,.14);
            --txt: #F7FBFF;
            --txt-soft: rgba(247,251,255,.78);
            --shadow: 0 14px 40px rgba(0,0,0,.22);
        }

        html, body, [class*="css"] {
            color: var(--txt);
        }

        .stApp {
            background:
                radial-gradient(circle at 14% 18%, rgba(121,184,255,.24), transparent 24%),
                radial-gradient(circle at 85% 15%, rgba(10,88,202,.22), transparent 20%),
                radial-gradient(circle at 78% 80%, rgba(0,43,99,.24), transparent 25%),
                linear-gradient(135deg, #02101f 0%, #062349 38%, #0B3971 70%, #124E93 100%);
        }

        .block-container {
            max-width: 1220px;
            padding-top: 1rem;
            padding-bottom: 1.4rem;
        }

        header, footer {visibility: hidden;}
        div[data-testid="stToolbar"] {visibility: hidden; height: 0px;}

        .hero-wrap{
            background: linear-gradient(180deg, rgba(255,255,255,.10), rgba(255,255,255,.05));
            border: 1px solid rgba(255,255,255,.14);
            border-radius: 28px;
            padding: 22px 26px;
            box-shadow: var(--shadow);
            backdrop-filter: blur(18px) saturate(160%);
            -webkit-backdrop-filter: blur(18px) saturate(160%);
            text-align:center;
            margin-bottom: 14px;
        }

        .hero-title{
            font-size: 2.15rem;
            font-weight: 900;
            letter-spacing: .8px;
            color: #FFFFFF;
            line-height: 1.02;
            text-transform: uppercase;
        }

        .hero-sub{
            margin-top: 8px;
            font-size: .88rem;
            color: rgba(255,255,255,.76);
            letter-spacing: .7px;
            text-transform: uppercase;
        }

        .toolbar-wrap{
            background: linear-gradient(180deg, rgba(255,255,255,.08), rgba(255,255,255,.04));
            border: 1px solid rgba(255,255,255,.12);
            border-radius: 24px;
            padding: 14px 16px;
            box-shadow: 0 10px 30px rgba(0,0,0,.18);
            backdrop-filter: blur(16px) saturate(160%);
            -webkit-backdrop-filter: blur(16px) saturate(160%);
            margin-bottom: 14px;
        }

        .kpi {
            border: 1px solid rgba(255,255,255,0.14);
            border-radius: 22px;
            padding: 16px 16px;
            min-height: 132px;
            height: 100%;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            background: linear-gradient(180deg, rgba(255,255,255,.105), rgba(255,255,255,.045));
            box-shadow: 0 12px 30px rgba(0,0,0,.18);
            backdrop-filter: blur(16px) saturate(160%);
            -webkit-backdrop-filter: blur(16px) saturate(160%);
            overflow: hidden;
            position: relative;
        }

        .kpi::before {
            content: "";
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 3px;
            background: rgba(121,184,255,.72);
        }

        .kpi.positive::before { background: rgba(111,232,181,.90); }
        .kpi.negative::before { background: rgba(255,128,128,.90); }
        .kpi.warning::before  { background: rgba(255,207,112,.92); }
        .kpi.neutral::before  { background: rgba(121,184,255,.75); }

        .kpi.positive {
            background: linear-gradient(180deg, rgba(73,170,128,.15), rgba(255,255,255,.045));
        }

        .kpi.negative {
            background: linear-gradient(180deg, rgba(190,75,75,.14), rgba(255,255,255,.045));
        }

        .kpi.warning {
            background: linear-gradient(180deg, rgba(190,145,65,.14), rgba(255,255,255,.045));
        }

        .kpi .label {
            opacity: .86;
            font-size: .82rem;
            font-weight: 800;
            color: rgba(255,255,255,.82);
            text-transform: uppercase;
            letter-spacing: .55px;
            min-height: 30px;
        }

        .kpi .value {
            font-size: clamp(1.35rem, 2vw, 1.82rem);
            font-weight: 950;
            line-height: 1.04;
            margin-top: 9px;
            color:#FFFFFF;
            text-transform: uppercase;
            word-break: break-word;
        }

        .kpi .sub {
            opacity: .74;
            font-size: .76rem;
            margin-top: .55rem;
            color: rgba(255,255,255,.74);
            text-transform: uppercase;
            letter-spacing: .32px;
            line-height: 1.25;
            min-height: 30px;
        }

        .balance-hero {
            border: 1px solid rgba(255,255,255,.16);
            border-radius: 24px;
            padding: 18px 20px;
            margin: 10px 0 16px 0;
            background:
                radial-gradient(circle at 12% 20%, rgba(121,184,255,.18), transparent 34%),
                linear-gradient(180deg, rgba(255,255,255,.11), rgba(255,255,255,.045));
            box-shadow: 0 14px 34px rgba(0,0,0,.18);
            backdrop-filter: blur(18px) saturate(165%);
            -webkit-backdrop-filter: blur(18px) saturate(165%);
        }

        .balance-hero .eyebrow {
            font-size: .74rem;
            font-weight: 900;
            color: rgba(255,255,255,.68);
            letter-spacing: .7px;
            text-transform: uppercase;
        }

        .balance-hero .main {
            margin-top: 5px;
            font-size: 1.75rem;
            font-weight: 950;
            color: #fff;
            letter-spacing: .2px;
            text-transform: uppercase;
        }

        .balance-hero .desc {
            margin-top: 7px;
            font-size: .84rem;
            line-height: 1.42;
            color: rgba(255,255,255,.76);
            text-transform: uppercase;
            letter-spacing: .3px;
        }

        .pill {
            display:inline-block;
            padding:8px 12px;
            border-radius:999px;
            border:1px solid rgba(255,255,255,.14);
            background:linear-gradient(180deg, rgba(255,255,255,.10), rgba(255,255,255,.05));
            font-size:.80rem;
            color: rgba(255,255,255,.92);
            box-shadow: 0 8px 22px rgba(0,0,0,.14);
            backdrop-filter: blur(14px);
            -webkit-backdrop-filter: blur(14px);
            text-transform: uppercase;
            letter-spacing: .35px;
        }

        .section-title {
            font-size: 1.08rem;
            font-weight: 900;
            letter-spacing: .55px;
            text-transform: uppercase;
            margin: .25rem 0 .55rem 0;
            color: rgba(255,255,255,.95);
        }

        .hr {
            height:1px;
            background: linear-gradient(90deg, transparent, rgba(255,255,255,.22), transparent);
            margin: 1rem 0 1rem 0;
        }

        .calendar-wrap{
            background: linear-gradient(180deg, rgba(255,255,255,.07), rgba(255,255,255,.04));
            border: 1px solid rgba(255,255,255,.12);
            border-radius: 20px;
            padding: 14px;
            margin-top: 12px;
        }

        .calendar-weekday{
            text-align:center;
            font-size:.78rem;
            font-weight:800;
            color: rgba(255,255,255,.78);
            letter-spacing:.45px;
            margin-bottom:6px;
        }

        .calendar-note{
            font-size:.80rem;
            opacity:.82;
            letter-spacing:.25px;
            text-transform:uppercase;
        }

        div[data-testid="stFileUploader"] > section,
        div[data-testid="stDataFrame"],
        div[data-testid="stTable"],
        div[data-testid="stDataEditor"],
        div[data-testid="stMetric"],
        div[data-testid="stVerticalBlockBorderWrapper"],
        div[data-testid="stAlert"],
        div[data-testid="stExpander"] {
            border-radius: 22px !important;
        }

        div[data-testid="stFileUploader"] > section {
            background: linear-gradient(180deg, rgba(255,255,255,.10), rgba(255,255,255,.05)) !important;
            border: 1px solid rgba(255,255,255,.14) !important;
            box-shadow: 0 10px 28px rgba(0,0,0,.15) !important;
            backdrop-filter: blur(16px) saturate(160%);
            -webkit-backdrop-filter: blur(16px) saturate(160%);
        }

        .stButton > button,
        .stDownloadButton > button {
            width: 100%;
            border-radius: 16px !important;
            border: 1px solid rgba(255,255,255,.16) !important;
            background: linear-gradient(180deg, rgba(121,184,255,.24), rgba(10,88,202,.16)) !important;
            color: white !important;
            font-weight: 900 !important;
            min-height: 46px !important;
            box-shadow: 0 10px 24px rgba(0,0,0,.18) !important;
            backdrop-filter: blur(14px);
            -webkit-backdrop-filter: blur(14px);
            text-transform: uppercase !important;
            letter-spacing: .4px !important;
        }

        .stButton > button:hover,
        .stDownloadButton > button:hover {
            border-color: rgba(255,255,255,.24) !important;
            transform: translateY(-1px);
            box-shadow: 0 14px 28px rgba(0,0,0,.22) !important;
        }

        div[data-baseweb="select"] > div,
        .stTextInput > div > div > input,
        .stTextArea textarea,
        input[type="number"] {
            background: rgba(255,255,255,.08) !important;
            border: 1px solid rgba(255,255,255,.15) !important;
            border-radius: 14px !important;
            color: white !important;
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            text-transform: uppercase;
        }

        div[data-baseweb="tab-list"] {
            gap: 10px;
            background: transparent !important;
        }

        button[data-baseweb="tab"] {
            border-radius: 16px !important;
            padding: 10px 18px !important;
            background: rgba(255,255,255,.06) !important;
            border: 1px solid rgba(255,255,255,.12) !important;
            color: rgba(255,255,255,.88) !important;
            backdrop-filter: blur(14px);
            -webkit-backdrop-filter: blur(14px);
            text-transform: uppercase;
            letter-spacing: .35px;
        }

        button[data-baseweb="tab"][aria-selected="true"] {
            background: linear-gradient(180deg, rgba(121,184,255,.24), rgba(10,88,202,.15)) !important;
            border-color: rgba(255,255,255,.22) !important;
            color: #fff !important;
            box-shadow: 0 10px 24px rgba(0,0,0,.16);
        }

        [data-testid="stDataFrame"] > div,
        [data-testid="stDataEditor"] > div {
            background: linear-gradient(180deg, rgba(255,255,255,.10), rgba(255,255,255,.05)) !important;
            border: 1px solid rgba(255,255,255,.14) !important;
            box-shadow: 0 10px 28px rgba(0,0,0,.15) !important;
            backdrop-filter: blur(14px);
            -webkit-backdrop-filter: blur(14px);
        }

        [data-testid="stMarkdownContainer"] p,
        [data-testid="stMarkdownContainer"] li,
        .stCaption {
            color: rgba(255,255,255,.88);
        }


        .explain {
            border: 1px solid rgba(255,255,255,.13);
            border-radius: 18px;
            padding: 12px 14px;
            background: linear-gradient(180deg, rgba(255,255,255,.075), rgba(255,255,255,.035));
            box-shadow: 0 8px 22px rgba(0,0,0,.12);
            backdrop-filter: blur(14px);
            -webkit-backdrop-filter: blur(14px);
            margin: 8px 0 10px 0;
        }

        .explain .title {
            font-size: .84rem;
            font-weight: 900;
            letter-spacing: .45px;
            text-transform: uppercase;
            color: rgba(255,255,255,.93);
            margin-bottom: 4px;
        }

        .explain .body {
            font-size: .82rem;
            line-height: 1.35;
            color: rgba(255,255,255,.74);
            text-transform: uppercase;
            letter-spacing: .25px;
        }

        .mini-stat {
            border: 1px solid rgba(255,255,255,.11);
            border-radius: 18px;
            padding: 12px 12px;
            background: linear-gradient(180deg, rgba(255,255,255,.08), rgba(255,255,255,.035));
            box-shadow: 0 8px 20px rgba(0,0,0,.12);
            min-height: 92px;
        }

        .mini-stat .label {
            font-size:.76rem;
            font-weight:800;
            color:rgba(255,255,255,.72);
            text-transform:uppercase;
            letter-spacing:.35px;
        }

        .mini-stat .value {
            font-size:1.28rem;
            font-weight:900;
            color:#fff;
            margin-top:5px;
        }

        .mini-stat .help {
            font-size:.74rem;
            color:rgba(255,255,255,.66);
            margin-top:4px;
            line-height:1.25;
        }

        label, .st-emotion-cache-16txtl3, .st-emotion-cache-pkbazv {
            text-transform: uppercase !important;
            letter-spacing: .35px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def hero_header() -> None:
    st.markdown(
        """
        <div class="hero-wrap">
            <div class="hero-title">CONTROL DE ASISTENCIA APP</div>
            <div class="hero-sub">DESARROLLADA POR MARCELO JAMESON</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def kpi_card(label: str, value: str, sub: str = "", tone: str = "neutral") -> None:
    tone = tone if tone in {"neutral", "positive", "negative", "warning"} else "neutral"
    st.markdown(
        f"""
        <div class="kpi {tone}">
            <div class="label">{html.escape(str(label))}</div>
            <div class="value">{html.escape(str(value))}</div>
            <div class="sub">{html.escape(str(sub))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def explain_box(title: str, body: str) -> None:
    st.markdown(
        f"""
        <div class="explain">
            <div class="title">{html.escape(str(title))}</div>
            <div class="body">{html.escape(str(body))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def balance_hero(balance_minutes: int) -> None:
    if balance_minutes > 0:
        main_text = f"{minutes_to_hhmm(balance_minutes)} A FAVOR"
        desc = (
            "EL EMPLEADO YA COMPENSÓ SUS HORAS FALTANTES DEL PERÍODO Y CONSERVA ESTE EXCEDENTE "
            "COMO SALDO POSITIVO."
        )
    elif balance_minutes < 0:
        main_text = f"{minutes_to_hhmm(abs(balance_minutes))} POR COMPENSAR"
        desc = (
            "LAS HORAS EXTRA GENERADAS TODAVÍA NO ALCANZAN PARA CUBRIR LAS SALIDAS TEMPRANAS "
            "DEL PERÍODO. ESTE ES EL TIEMPO PENDIENTE."
        )
    else:
        main_text = "BALANCE EN CERO"
        desc = (
            "LAS HORAS EXTRA Y LAS HORAS FALTANTES DEL PERÍODO SE COMPENSAN EXACTAMENTE."
        )

    st.markdown(
        f"""
        <div class="balance-hero">
            <div class="eyebrow">BALANCE REAL DEL PERÍODO</div>
            <div class="main">{html.escape(main_text)}</div>
            <div class="desc">{html.escape(desc)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def mini_stat(label: str, value: str, help_text: str = "") -> None:
    st.markdown(
        f"""
        <div class="mini-stat">
            <div class="label">{html.escape(str(label))}</div>
            <div class="value">{html.escape(str(value))}</div>
            <div class="help">{html.escape(str(help_text))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_title(title: str) -> None:
    st.markdown(f'<div class="section-title">{html.escape(str(title))}</div>', unsafe_allow_html=True)


# ============================================================
# COPY
# ============================================================
def copy_table_button(df: pd.DataFrame, label: str, key: str) -> None:
    if df is None:
        df = pd.DataFrame()

    tsv = df.to_csv(sep="\t", index=False)
    payload = {"tsv": tsv}
    j = json.dumps(payload)

    st.components.v1.html(
        f"""
        <div style="display:flex; gap:10px; align-items:center; margin: 8px 0 10px 0;">
          <button id="btn_{key}" style="
              padding:12px 14px; border-radius:16px; border:1px solid rgba(255,255,255,.16);
              background:linear-gradient(180deg, rgba(121,184,255,.24), rgba(10,88,202,.16));
              color:white; font-weight:900; cursor:pointer; width:100%;
              box-shadow:0 10px 24px rgba(0,0,0,.18);
              backdrop-filter: blur(14px);
              text-transform:uppercase; letter-spacing:.35px;
          ">{html.escape(label)}</button>
          <span id="ok_{key}" style="opacity:.0; font-weight:800; color:white;">COPIADO ✅</span>
        </div>

        <script>
        const data_{key} = {j};
        const btn_{key} = document.getElementById("btn_{key}");
        const ok_{key} = document.getElementById("ok_{key}");

        btn_{key}.addEventListener("click", async () => {{
          try {{
            await navigator.clipboard.writeText(data_{key}.tsv);
            ok_{key}.style.opacity = "1";
            setTimeout(() => ok_{key}.style.opacity = "0", 1400);
          }} catch (e) {{
            alert("No se pudo copiar automáticamente. Probá con Chrome/Edge o habilitá permisos de portapapeles.");
          }}
        }});
        </script>
        """,
        height=66,
    )


# ============================================================
# HELPERS
# ============================================================
def minutes_to_hhmm(mins: int | float | None) -> str:
    mins = int(round(mins)) if mins is not None else 0
    h = mins // 60
    m = mins % 60
    return f"{h:02d}:{m:02d}"


def delta_short(mins: int | float) -> str:
    mins = int(round(mins))
    sign = "+" if mins > 0 else "-" if mins < 0 else ""
    mins_abs = abs(mins)
    h, m = mins_abs // 60, mins_abs % 60
    if h == 0:
        return f"{sign}{m:02d}M" if sign else "0M"
    return f"{sign}{h}H {m:02d}M" if sign else f"{h}H {m:02d}M"


def normalize_text_key(value: str) -> str:
    value = str(value or "").strip().upper()
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = re.sub(r"\s+", " ", value)
    return value


def display_dni(value: str) -> str:
    v = str(value or "").strip()
    return v if v else "SIN DNI"


def get_work_date(dt: pd.Timestamp, night_adjustment: bool = True) -> date:
    """
    Fecha laboral:
    - Si AJUSTE MADRUGADA está activo, 00:00 a 03:59 se considera día anterior.
    - Si está apagado, se usa la fecha calendario real.
    """
    if pd.isna(dt):
        return date.today()
    if night_adjustment and int(dt.hour) < WORKDAY_CUTOFF_HOUR:
        return (dt - pd.Timedelta(days=1)).date()
    return dt.date()


def read_legacy_biff2(file) -> pd.DataFrame:
    """
    Lector de respaldo para archivos .XLS antiguos exportados por algunos relojes.
    El archivo de ejemplo del reloj usa registros LABEL de BIFF2.
    """
    file.seek(0)
    data = file.read()
    if not isinstance(data, (bytes, bytearray)):
        raise ValueError("NO SE PUDO LEER EL ARCHIVO XLS ANTIGUO.")

    cells: dict[tuple[int, int], str] = {}
    pos = 0

    while pos + 4 <= len(data):
        record_id = int.from_bytes(data[pos:pos + 2], "little")
        length = int.from_bytes(data[pos + 2:pos + 4], "little")
        payload = data[pos + 4:pos + 4 + length]

        if pos + 4 + length > len(data):
            break

        # BIFF2 LABEL: row(2), col(2), attributes(3), string_length(1), text.
        if record_id == 0x0004 and len(payload) >= 8:
            row = int.from_bytes(payload[0:2], "little")
            col = int.from_bytes(payload[2:4], "little")
            string_length = int(payload[7])
            raw_text = payload[8:8 + string_length]
            text = raw_text.decode("latin1", errors="replace")
            cells[(row, col)] = text

        pos += 4 + length

    if not cells:
        raise ValueError("EL ARCHIVO XLS NO CONTIENE DATOS LEGIBLES.")

    max_row = max(r for r, _ in cells.keys())
    max_col = max(c for _, c in cells.keys())

    headers = [str(cells.get((0, c), "")).strip() for c in range(max_col + 1)]
    if not any(headers):
        raise ValueError("NO SE ENCONTRARON ENCABEZADOS EN EL XLS.")

    rows = []
    for row_idx in range(1, max_row + 1):
        row = {
            headers[col_idx] or f"Columna_{col_idx + 1}": cells.get((row_idx, col_idx), "")
            for col_idx in range(max_col + 1)
        }
        if any(str(v).strip() for v in row.values()):
            rows.append(row)

    return pd.DataFrame(rows)


def read_excel_auto(file) -> pd.DataFrame:
    errors = []

    try:
        file.seek(0)
        return pd.read_excel(file)
    except Exception as exc:
        errors.append(str(exc))

    try:
        file.seek(0)
        return pd.read_excel(file, engine="openpyxl")
    except Exception as exc:
        errors.append(str(exc))

    try:
        file.seek(0)
        return pd.read_excel(file, engine="xlrd")
    except Exception as exc:
        errors.append(str(exc))

    try:
        return read_legacy_biff2(file)
    except Exception as exc:
        errors.append(str(exc))

    raise ValueError(
        "NO SE PUDO LEER EL ARCHIVO. VERIFICÁ QUE SEA UN EXCEL DEL RELOJ VÁLIDO. "
        + " | ".join(errors[-2:])
    )


def validate_format(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    missing_required = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing_required:
        raise ValueError(
            f"FORMATO INCORRECTO DEL RELOJ. FALTA LA COLUMNA OBLIGATORIA: {', '.join(missing_required)}"
        )

    for col in OPTIONAL_COLS:
        if col not in df.columns:
            df[col] = ""

    return df


def parse_and_clean(df: pd.DataFrame, night_adjustment: bool = True) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # El reloj puede traer una columna llamada "Tipo".
    # La preservamos como dato original para que no choque con el perfil
    # DOCENTE / NO DOCENTE administrado por la aplicación.
    if "Tipo" in df.columns:
        df = df.rename(columns={"Tipo": "Tipo_reloj"})

    for col in ["Nombre", "Marc.", "Estado", "NvoEstado"]:
        if col not in df.columns:
            df[col] = ""

    df["__rowid__"] = range(1, len(df) + 1)

    df["DNI"] = (
        df["Nombre"]
        .fillna("")
        .astype(str)
        .str.replace(r"\D", "", regex=True)
        .str.strip()
    )

    df["Empleado"] = (
        df["Marc."]
        .fillna("")
        .astype(str)
        .replace({"nan": "", "None": ""})
        .str.strip()
    )

    df["FechaHora"] = pd.to_datetime(df["Estado"], errors="coerce", dayfirst=True)
    df = df.dropna(subset=["FechaHora"]).copy()

    def resolve_employee_name(row) -> str:
        emp = str(row["Empleado"] or "").strip()
        dni = str(row["DNI"] or "").strip()
        rowid = int(row["__rowid__"])

        if emp:
            return emp
        if dni:
            return f"SIN NOMBRE · DNI {dni}"
        return f"SIN IDENTIFICAR · REG {rowid}"

    df["Empleado"] = df.apply(resolve_employee_name, axis=1)

    def resolve_employee_key(row) -> str:
        dni = str(row["DNI"] or "").strip()
        emp = str(row["Empleado"] or "").strip()
        rowid = int(row["__rowid__"])

        if dni:
            return f"DNI::{dni}"
        if emp and not emp.startswith("SIN IDENTIFICAR · REG "):
            return f"NOMBRE::{normalize_text_key(emp)}"
        return f"REG::{rowid}"

    df["EmployeeKey"] = df.apply(resolve_employee_key, axis=1)

    # Fecha real calendario y fecha laboral corregida.
    df["FechaReal"] = df["FechaHora"].dt.date
    df["Fecha"] = df["FechaHora"].apply(lambda x: get_work_date(x, night_adjustment))
    df["Ajuste_madrugada"] = df.apply(
        lambda r: "SI" if r["Fecha"] != r["FechaReal"] else "",
        axis=1,
    )

    df = df.sort_values(["Empleado", "DNI", "FechaHora"]).reset_index(drop=True)
    df["DNI"] = df["DNI"].astype(str)
    return df


def init_profiles(raw: pd.DataFrame) -> pd.DataFrame:
    base = (
        raw[["EmployeeKey", "DNI", "Empleado"]]
        .drop_duplicates()
        .sort_values(["Empleado", "DNI"])
        .reset_index(drop=True)
    )

    if "profiles" not in st.session_state:
        p = base.copy()
        p["Tipo"] = "NO Docente"
        st.session_state["profiles"] = p
        return p

    p = st.session_state["profiles"].copy()
    merged = base.merge(p[["EmployeeKey", "Tipo"]], on="EmployeeKey", how="left")
    merged["Tipo"] = merged["Tipo"].fillna("NO Docente")
    merged = merged[["EmployeeKey", "DNI", "Empleado", "Tipo"]]
    st.session_state["profiles"] = merged
    return merged


def apply_profiles(raw: pd.DataFrame, profiles: pd.DataFrame) -> pd.DataFrame:
    base = raw.drop(columns=["Tipo"], errors="ignore").copy()
    m = base.merge(profiles[["EmployeeKey", "Tipo"]], on="EmployeeKey", how="left")
    m["Tipo"] = m["Tipo"].fillna("NO Docente")
    return m


# ============================================================
# FERIADOS
# ============================================================
def parse_holidays(text: str) -> set[date]:
    holidays: set[date] = set()
    if not text:
        return holidays

    for line in str(text).replace(",", "\n").splitlines():
        raw = line.strip()
        if not raw:
            continue
        try:
            holidays.add(pd.to_datetime(raw, dayfirst=True).date())
        except Exception:
            pass
    return holidays


def holidays_to_text(holidays: set[date]) -> str:
    if not holidays:
        return ""
    return "\n".join(sorted([d.strftime("%d/%m/%Y") for d in holidays]))


def init_holidays_state() -> None:
    if "holidays_set" not in st.session_state:
        st.session_state["holidays_set"] = set()
    if "holidays_text_input" not in st.session_state:
        st.session_state["holidays_text_input"] = ""
    if "holiday_calendar_month" not in st.session_state:
        today = date.today()
        st.session_state["holiday_calendar_month"] = today.month
    if "holiday_calendar_year" not in st.session_state:
        st.session_state["holiday_calendar_year"] = date.today().year


def sync_holidays_text_input_from_set() -> None:
    st.session_state["holidays_text_input"] = holidays_to_text(st.session_state["holidays_set"])


def apply_text_holidays_from_value(text_value: str) -> None:
    st.session_state["holidays_set"] = parse_holidays(text_value or "")
    sync_holidays_text_input_from_set()


def clear_all_holidays() -> None:
    st.session_state["holidays_set"] = set()
    st.session_state["holidays_text_input"] = ""


def toggle_holiday(day_value: date) -> None:
    holidays = set(st.session_state["holidays_set"])
    if day_value in holidays:
        holidays.remove(day_value)
    else:
        holidays.add(day_value)
    st.session_state["holidays_set"] = holidays
    sync_holidays_text_input_from_set()


def render_holiday_calendar() -> None:
    month = int(st.session_state["holiday_calendar_month"])
    year = int(st.session_state["holiday_calendar_year"])
    holidays = st.session_state["holidays_set"]

    st.markdown('<div class="calendar-wrap">', unsafe_allow_html=True)

    nav1, nav2, nav3, nav4 = st.columns([1.2, 1, 1, 1.4])
    with nav1:
        selected_month_name = st.selectbox(
            "MES",
            options=list(MONTH_NAMES.values()),
            index=month - 1,
            key="holiday_month_name_selector",
            label_visibility="collapsed",
        )
        month = {v: k for k, v in MONTH_NAMES.items()}[selected_month_name]
        st.session_state["holiday_calendar_month"] = month

    with nav2:
        selected_year = st.number_input(
            "AÑO",
            min_value=2000,
            max_value=2100,
            value=year,
            step=1,
            key="holiday_year_number",
            label_visibility="collapsed",
        )
        year = int(selected_year)
        st.session_state["holiday_calendar_year"] = year

    with nav3:
        if st.button("HOY", key="holiday_go_today"):
            today = date.today()
            st.session_state["holiday_calendar_month"] = today.month
            st.session_state["holiday_calendar_year"] = today.year
            st.rerun()

    with nav4:
        st.markdown(
            f"""<div class="pill">{MONTH_NAMES[month]} {year}</div>""",
            unsafe_allow_html=True,
        )

    week_cols = st.columns(7)
    weekday_names = ["LUN", "MAR", "MIÉ", "JUE", "VIE", "SÁB", "DOM"]
    for col, wd in zip(week_cols, weekday_names):
        with col:
            st.markdown(f'<div class="calendar-weekday">{wd}</div>', unsafe_allow_html=True)

    cal = calendar.Calendar(firstweekday=0)
    weeks = cal.monthdayscalendar(year, month)

    for week_idx, week in enumerate(weeks):
        cols = st.columns(7)
        for day_idx, day_num in enumerate(week):
            with cols[day_idx]:
                if day_num == 0:
                    st.markdown("&nbsp;", unsafe_allow_html=True)
                    continue

                current_date = date(year, month, day_num)
                is_holiday = current_date in holidays
                label = f"🟦 {day_num}" if is_holiday else f"{day_num}"
                key = f"holiday_btn_{year}_{month}_{day_num}_{week_idx}_{day_idx}"

                if st.button(label, key=key, use_container_width=True):
                    toggle_holiday(current_date)
                    st.rerun()

    st.markdown(
        '<div class="calendar-note">TOCÁ UN DÍA PARA AGREGARLO O QUITARLO COMO FERIADO.</div>',
        unsafe_allow_html=True,
    )
    st.markdown('</div>', unsafe_allow_html=True)


def is_weekend(day_value) -> bool:
    return pd.to_datetime(day_value).weekday() >= 5


def is_holiday(day_value, holidays: set[date]) -> bool:
    return pd.to_datetime(day_value).date() in holidays


def is_special_day(day_value, holidays: set[date]) -> bool:
    ts = pd.to_datetime(day_value)
    return ts.weekday() >= 5 or ts.date() in holidays


def day_type_label(day_value, holidays: set[date]) -> str:
    weekend = is_weekend(day_value)
    holiday = is_holiday(day_value, holidays)
    if weekend and holiday:
        return "FERIADO/FIN DE SEMANA"
    if holiday:
        return "FERIADO"
    if weekend:
        return "FIN DE SEMANA"
    return "HÁBIL"


# ============================================================
# CÁLCULOS DE MARCACIONES
# ============================================================
def pair_alternating(times: list[pd.Timestamp]) -> tuple[int, int]:
    """
    Suma pares alternados entrada/salida ya normalizados.
    Ejemplo: 07:00-14:00 + 15:00-21:00.
    """
    times = [pd.to_datetime(t) for t in times if pd.notna(t)]
    times.sort()
    total = 0
    pairs = 0
    for i in range(0, len(times) - 1, 2):
        a, b = times[i], times[i + 1]
        if b >= a:
            total += int((b - a).total_seconds() // 60)
            pairs += 1
    return total, pairs


def format_marks(times: list[pd.Timestamp]) -> str:
    times = [pd.to_datetime(t) for t in times if pd.notna(t)]
    times.sort()
    if not times:
        return ""
    return " · ".join(t.strftime("%H:%M") for t in times)


def normalize_duplicate_punches(
    times: list[pd.Timestamp],
    window_minutes: int = DUPLICATE_PUNCH_WINDOW_MINUTES,
) -> tuple[list[pd.Timestamp], int, list[list[pd.Timestamp]]]:
    """
    Agrupa ráfagas de marcas muy cercanas generadas por el biométrico.

    Regla:
    - una ráfaga <= window_minutes representa un único evento;
    - para una ENTRADA se conserva la primera marca de la ráfaga;
    - para una SALIDA se conserva la última marca de la ráfaga.

    Así:
      07:06, 07:06, 14:00 -> 07:06, 14:00
      07:00, 14:00, 15:00, 21:00 -> se mantienen las cuatro.
    """
    clean = [pd.to_datetime(t) for t in times if pd.notna(t)]
    clean.sort()
    if not clean:
        return [], 0, []

    bursts: list[list[pd.Timestamp]] = [[clean[0]]]
    max_gap = pd.Timedelta(minutes=window_minutes)

    for current in clean[1:]:
        if current - bursts[-1][-1] <= max_gap:
            bursts[-1].append(current)
        else:
            bursts.append([current])

    effective: list[pd.Timestamp] = []
    for event_index, burst in enumerate(bursts):
        # eventos pares = entrada -> la primera;
        # eventos impares = salida -> la última.
        chosen = burst[0] if event_index % 2 == 0 else burst[-1]
        effective.append(chosen)

    ignored = sum(max(0, len(burst) - 1) for burst in bursts)
    return effective, ignored, bursts


def duplicate_bursts_detail(bursts: list[list[pd.Timestamp]]) -> str:
    details = []
    for burst in bursts:
        if len(burst) <= 1:
            continue
        first = pd.to_datetime(burst[0]).strftime("%H:%M")
        last = pd.to_datetime(burst[-1]).strftime("%H:%M")
        if first == last:
            details.append(f"{first} ×{len(burst)}")
        else:
            details.append(f"{first}-{last} ×{len(burst)}")
    return " | ".join(details)


def build_pair_details(times: list[pd.Timestamp]) -> str:
    times = [pd.to_datetime(t) for t in times if pd.notna(t)]
    times.sort()
    details = []
    for i in range(0, len(times) - 1, 2):
        a, b = times[i], times[i + 1]
        if b >= a:
            mins = int((b - a).total_seconds() // 60)
            details.append(
                f"{a.strftime('%H:%M')}-{b.strftime('%H:%M')} ({minutes_to_hhmm(mins)})"
            )
    return " | ".join(details)


def split_balance(worked: int, expected: int, saldo: int) -> tuple[int, int, int]:
    normal = min(max(worked, 0), max(expected, 0)) if expected > 0 else 0
    extra = max(saldo, 0)
    faltante = max(-saldo, 0)
    return int(normal), int(extra), int(faltante)


def split_interval_by_day(start: pd.Timestamp, end: pd.Timestamp) -> list[tuple[pd.Timestamp, int]]:
    chunks = []
    current = start

    while current < end:
        next_midnight = (current.normalize() + pd.Timedelta(days=1))
        chunk_end = min(next_midnight, end)
        minutes = int((chunk_end - current).total_seconds() // 60)
        if minutes > 0:
            chunks.append((current.normalize(), minutes))
        current = chunk_end

    return chunks


def expected_and_saldo(worked: int, expected_nodoc: int, special: bool, has_mark: bool) -> tuple[int, int, str]:
    if special:
        expected = 0
        saldo = worked
        if worked > 0:
            cumple = "EXTRA"
        else:
            cumple = ""
        return expected, saldo, cumple

    expected = expected_nodoc if has_mark else 0
    saldo = worked - expected if expected else 0

    if expected:
        cumple = "OK" if saldo >= 0 else "FALTA"
    else:
        cumple = ""

    return expected, saldo, cumple


# ============================================================
# CÁLCULO NORMAL
# ============================================================
def calc_daily_standard(raw: pd.DataFrame, expected_nodoc: int, holidays: set[date]) -> pd.DataFrame:
    rows = []

    for (ekey, dni, emp, tipo, day), g in raw.groupby(
        ["EmployeeKey", "DNI", "Empleado", "Tipo", "Fecha"], dropna=False
    ):
        g = g.sort_values("FechaHora")

        raw_times = g["FechaHora"].tolist()
        raw_mark_count = int(g.shape[0])

        effective_times, duplicates_ignored, bursts = normalize_duplicate_punches(raw_times)
        effective_mark_count = len(effective_times)

        first = effective_times[0] if effective_times else pd.NaT
        last = effective_times[-1] if effective_times else pd.NaT

        worked_pairs, pairs = pair_alternating(effective_times)

        fecha_ts = pd.to_datetime(day)
        special = is_special_day(fecha_ts.date(), holidays)
        tipo_dia = day_type_label(fecha_ts.date(), holidays)

        # Después de limpiar dobles marcas, debe quedar una cantidad par.
        incompleto = (effective_mark_count % 2 != 0) or (effective_mark_count < 2)
        cortes = pairs >= 2

        if tipo == "Docente":
            worked = worked_pairs
            expected = 0
            saldo = 0
            cumple = "INCOMPLETO" if incompleto and effective_mark_count > 0 else ""
        else:
            # Suma todos los pares reales del día.
            # 07-14 + 15-21 = 13 horas trabajadas, no 14 horas corridas.
            worked = worked_pairs
            expected, saldo, cumple = expected_and_saldo(
                worked=worked,
                expected_nodoc=expected_nodoc,
                special=special,
                has_mark=effective_mark_count >= 1,
            )

            if incompleto and effective_mark_count > 0:
                cumple = "INCOMPLETO"

        normal_min, extra_min, faltante_min = split_balance(worked, expected, saldo)

        rows.append(
            {
                "EmployeeKey": ekey,
                "DNI": str(dni),
                "Empleado": emp,
                "Tipo": tipo,
                "Fecha": fecha_ts,
                "Tipo_dia": tipo_dia,
                "Es_fin_de_semana": "SI" if is_weekend(fecha_ts.date()) else "",
                "Es_feriado": "SI" if is_holiday(fecha_ts.date(), holidays) else "",
                "Primera": first,
                "Ultima": last,
                "Horas": minutes_to_hhmm(worked),
                "Minutos": int(worked),
                "Normal": minutes_to_hhmm(normal_min),
                "Normal_min": int(normal_min),
                "Extra_dia": minutes_to_hhmm(extra_min),
                "Extra_dia_min": int(extra_min),
                "Faltante_dia": minutes_to_hhmm(faltante_min),
                "Faltante_dia_min": int(faltante_min),
                # Auditoría: mostramos TODAS las marcas originales y las que realmente se usaron.
                "Marcaciones_detalle": format_marks(raw_times),
                "Marcaciones_efectivas_detalle": format_marks(effective_times),
                "Tramos_detalle": build_pair_details(effective_times),
                "Duplicadas_ignoradas": int(duplicates_ignored),
                "Duplicadas_detalle": duplicate_bursts_detail(bursts),
                "Esperado_min": int(expected),
                "Esperado": minutes_to_hhmm(expected),
                "Saldo_min": int(saldo),
                "Saldo": delta_short(saldo),
                "Cumple": cumple,
                "Marcaciones": raw_mark_count,
                "Marcaciones_efectivas": effective_mark_count,
                "Pares_estimados": int(pairs),
                "Cortes": "SI" if cortes else "",
                "Incompleto": "SI" if incompleto and effective_mark_count > 0 else "",
                "Ajuste_madrugada": "SI" if (g["Ajuste_madrugada"] == "SI").any() else "",
            }
        )

    d = pd.DataFrame(rows)
    if d.empty:
        return d
    return d.sort_values(["Tipo", "Empleado", "DNI", "Fecha"]).reset_index(drop=True)


def calc_daily_drivers(raw: pd.DataFrame, expected_nodoc: int, holidays: set[date]) -> pd.DataFrame:
    """
    Modo chofer:
    - limpia dobles lecturas biométricas;
    - arma pares globales por empleado, sin resetear al cambiar el día;
    - si marca hoy 14:00 y mañana 19:00, es un viaje continuo;
    - dentro de cada viaje, solo las primeras 7h/6h normales cuentan como normales;
    - sábados, domingos y feriados son extra siempre.
    """
    rows = []

    for (ekey, dni, emp, tipo), g_emp in raw.groupby(
        ["EmployeeKey", "DNI", "Empleado", "Tipo"], dropna=False
    ):
        g_emp = g_emp.sort_values("FechaHora").copy()

        if tipo == "Docente":
            g_doc = calc_daily_standard(g_emp.copy(), expected_nodoc, holidays)
            if not g_doc.empty:
                rows.extend(g_doc.to_dict("records"))
            continue

        raw_times = [pd.to_datetime(t) for t in g_emp["FechaHora"].tolist() if pd.notna(t)]
        raw_times.sort()
        effective_times, duplicates_ignored_total, bursts = normalize_duplicate_punches(raw_times)

        def workday_for_timestamp(ts: pd.Timestamp):
            matches = g_emp[g_emp["FechaHora"] == ts]
            if not matches.empty:
                return matches.iloc[0]["Fecha"]
            return pd.to_datetime(ts).date()

        effective_marks_by_day = defaultdict(list)
        for ts in effective_times:
            effective_marks_by_day[workday_for_timestamp(ts)].append(ts)

        duplicates_by_day = defaultdict(int)
        duplicate_detail_by_day = defaultdict(list)
        for burst in bursts:
            if len(burst) <= 1:
                continue
            # el evento efectivo de la ráfaga determina el día al cual pertenece.
            burst_index = bursts.index(burst)
            chosen = burst[0] if burst_index % 2 == 0 else burst[-1]
            day = workday_for_timestamp(chosen)
            duplicates_by_day[day] += len(burst) - 1
            detail = duplicate_bursts_detail([burst])
            if detail:
                duplicate_detail_by_day[day].append(detail)

        raw_day_stats = {}
        for day, g_day in g_emp.groupby("Fecha"):
            g_day = g_day.sort_values("FechaHora")
            raw_day_stats[day] = {
                "Marcaciones": int(g_day.shape[0]),
                "Primera": g_day["FechaHora"].iloc[0] if not g_day.empty else pd.NaT,
                "Ultima": g_day["FechaHora"].iloc[-1] if not g_day.empty else pd.NaT,
                "Ajuste_madrugada": "SI" if (g_day["Ajuste_madrugada"] == "SI").any() else "",
            }

        day_worked = defaultdict(int)
        day_expected = defaultdict(int)
        day_pairs = defaultdict(int)
        day_trip_details = defaultdict(list)

        for i in range(0, len(effective_times) - 1, 2):
            start = effective_times[i]
            end = effective_times[i + 1]
            if pd.isna(start) or pd.isna(end) or end < start:
                continue

            remaining_normal = expected_nodoc
            chunks = split_interval_by_day(start, end)

            for day_ts, minutes in chunks:
                day_date = day_ts.date()
                special = is_special_day(day_date, holidays)

                if special:
                    normal_chunk = 0
                else:
                    normal_chunk = min(minutes, remaining_normal)
                    remaining_normal -= normal_chunk

                day_worked[day_date] += minutes
                day_expected[day_date] += normal_chunk
                day_pairs[day_date] += 1

                day_trip_details[day_date].append(
                    f"{start.strftime('%d/%m %H:%M')} → {end.strftime('%d/%m %H:%M')}"
                )

        unmatched_day = None
        if len(effective_times) % 2 != 0:
            unmatched_day = workday_for_timestamp(effective_times[-1])

        all_days = set(raw_day_stats.keys()) | set(day_worked.keys()) | set(effective_marks_by_day.keys())

        for day in sorted(all_days):
            fecha_ts = pd.to_datetime(day)
            tipo_dia = day_type_label(day, holidays)

            raw_marc = raw_day_stats.get(day, {}).get("Marcaciones", 0)
            raw_marks_today = g_emp[g_emp["Fecha"] == day]["FechaHora"].tolist()
            effective_today = effective_marks_by_day.get(day, [])

            first = effective_today[0] if effective_today else raw_day_stats.get(day, {}).get("Primera", pd.NaT)
            last = effective_today[-1] if effective_today else raw_day_stats.get(day, {}).get("Ultima", pd.NaT)
            ajuste_madrugada = raw_day_stats.get(day, {}).get("Ajuste_madrugada", "")

            worked = int(day_worked.get(day, 0))
            expected = int(day_expected.get(day, 0))
            saldo = worked - expected
            pairs = int(day_pairs.get(day, 0))

            incompleto = unmatched_day == day
            cortes = pairs >= 2

            if incompleto:
                cumple = "INCOMPLETO"
            elif worked > 0 and expected == 0 and saldo > 0:
                cumple = "EXTRA"
            elif expected > 0:
                cumple = "OK" if saldo >= 0 else "FALTA"
            else:
                cumple = ""

            normal_min, extra_min, faltante_min = split_balance(worked, expected, saldo)

            rows.append(
                {
                    "EmployeeKey": ekey,
                    "DNI": str(dni),
                    "Empleado": emp,
                    "Tipo": tipo,
                    "Fecha": fecha_ts,
                    "Tipo_dia": tipo_dia,
                    "Es_fin_de_semana": "SI" if is_weekend(day) else "",
                    "Es_feriado": "SI" if is_holiday(day, holidays) else "",
                    "Primera": first,
                    "Ultima": last,
                    "Horas": minutes_to_hhmm(worked),
                    "Minutos": int(worked),
                    "Normal": minutes_to_hhmm(normal_min),
                    "Normal_min": int(normal_min),
                    "Extra_dia": minutes_to_hhmm(extra_min),
                    "Extra_dia_min": int(extra_min),
                    "Faltante_dia": minutes_to_hhmm(faltante_min),
                    "Faltante_dia_min": int(faltante_min),
                    "Marcaciones_detalle": format_marks(raw_marks_today),
                    "Marcaciones_efectivas_detalle": format_marks(effective_today),
                    "Tramos_detalle": " | ".join(day_trip_details.get(day, [])),
                    "Duplicadas_ignoradas": int(duplicates_by_day.get(day, 0)),
                    "Duplicadas_detalle": " | ".join(duplicate_detail_by_day.get(day, [])),
                    "Esperado_min": int(expected),
                    "Esperado": minutes_to_hhmm(expected),
                    "Saldo_min": int(saldo),
                    "Saldo": delta_short(saldo),
                    "Cumple": cumple,
                    "Marcaciones": int(raw_marc),
                    "Marcaciones_efectivas": int(len(effective_today)),
                    "Pares_estimados": int(pairs),
                    "Cortes": "SI" if cortes else "",
                    "Incompleto": "SI" if incompleto else "",
                    "Ajuste_madrugada": ajuste_madrugada,
                }
            )

    d = pd.DataFrame(rows)
    if d.empty:
        return d
    return d.sort_values(["Tipo", "Empleado", "DNI", "Fecha"]).reset_index(drop=True)


def calc_daily(raw: pd.DataFrame, expected_nodoc: int, holidays: set[date], driver_mode: bool) -> pd.DataFrame:
    if driver_mode:
        return calc_daily_drivers(raw, expected_nodoc, holidays)
    return calc_daily_standard(raw, expected_nodoc, holidays)


# ============================================================
# CORRECCIÓN AUTOMÁTICA NO DOCENTE
# ============================================================
def correct_missing_punches_for_employee(
    raw_emp: pd.DataFrame,
    expected_nodoc: int,
    night_adjustment: bool = True,
) -> tuple[pd.DataFrame, int]:
    """
    Corrección deliberadamente conservadora:
    solo agrega una salida automática si, después de ignorar dobles lecturas,
    queda UN único evento efectivo en el día.

    Nunca corrige automáticamente días con 3, 5, etc. eventos efectivos.
    """
    if raw_emp.empty:
        return raw_emp, 0

    corrected = raw_emp.copy()
    corrected["Fecha"] = corrected["FechaHora"].apply(
        lambda x: get_work_date(x, night_adjustment)
    )

    fixes = []
    nfix = 0

    for day, g in corrected.groupby("Fecha"):
        raw_times = sorted(g["FechaHora"].tolist())
        effective_times, _, _ = normalize_duplicate_punches(raw_times)

        if len(effective_times) == 1:
            t = effective_times[0]
            fix_out = t + pd.to_timedelta(expected_nodoc, unit="m")
            fixes.append(
                {
                    "FechaHora": fix_out,
                    "Fecha": get_work_date(fix_out, night_adjustment),
                }
            )
            nfix += 1

    if fixes:
        fx_rows = []
        template = corrected.iloc[0].copy()

        for f in fixes:
            row = template.copy()
            row["FechaHora"] = f["FechaHora"]
            row["FechaReal"] = row["FechaHora"].date()
            row["Fecha"] = f["Fecha"]
            row["Ajuste_madrugada"] = "SI" if row["Fecha"] != row["FechaReal"] else ""
            row["NvoEstado"] = "AUTO_FIX"
            row["Estado"] = row["FechaHora"].strftime("%d/%m/%Y %H:%M")
            fx_rows.append(row)

        corrected = (
            pd.concat([corrected, pd.DataFrame(fx_rows)], ignore_index=True)
            .sort_values("FechaHora")
            .reset_index(drop=True)
        )

    corrected["FechaReal"] = corrected["FechaHora"].dt.date
    corrected["Fecha"] = corrected["FechaHora"].apply(
        lambda x: get_work_date(x, night_adjustment)
    )
    corrected["Ajuste_madrugada"] = corrected.apply(
        lambda r: "SI" if r["Fecha"] != r["FechaReal"] else "",
        axis=1,
    )
    return corrected, nfix


def correct_missing_punches_all(raw: pd.DataFrame, expected_nodoc: int, night_adjustment: bool = True) -> tuple[pd.DataFrame, int]:
    if raw.empty:
        return raw, 0

    docentes = raw[raw["Tipo"] == "Docente"].copy()
    nodoc = raw[raw["Tipo"] == "NO Docente"].copy()
    if nodoc.empty:
        return raw, 0

    fixed_parts = []
    total_fixes = 0

    for ekey, g in nodoc.groupby(["EmployeeKey"]):
        corrected, nfix = correct_missing_punches_for_employee(g.copy(), expected_nodoc, night_adjustment)
        fixed_parts.append(corrected)
        total_fixes += nfix

    nodoc_fixed = pd.concat(fixed_parts, ignore_index=True) if fixed_parts else nodoc
    out = pd.concat([docentes, nodoc_fixed], ignore_index=True)
    out = out.sort_values(["Empleado", "DNI", "FechaHora"]).reset_index(drop=True)
    return out, total_fixes


# ============================================================
# RESÚMENES / TABLAS
# ============================================================
def summarize(daily: pd.DataFrame) -> pd.DataFrame:
    if daily.empty:
        return pd.DataFrame()

    def extras_pos(x: pd.Series) -> int:
        return int(x[x > 0].sum())

    def faltas_pos(x: pd.Series) -> int:
        return int((-x[x < 0]).sum())

    s = (
        daily.groupby(["EmployeeKey", "Empleado", "DNI", "Tipo"], as_index=False)
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
            Normal_min=("Normal_min", "sum"),
            Extra_dia_min=("Extra_dia_min", "sum"),
            Faltante_dia_min=("Faltante_dia_min", "sum"),
            Dias_OK=("Cumple", lambda x: int((x == "OK").sum())),
            Dias_FALTA=("Cumple", lambda x: int((x == "FALTA").sum())),
            Dias_INCOMPL=("Cumple", lambda x: int((x == "INCOMPLETO").sum())),
            Dias_EXTRA=("Cumple", lambda x: int((x == "EXTRA").sum())),
            Dias_Feriado=("Es_feriado", lambda x: int((x == "SI").sum()) if "Es_feriado" in daily.columns else 0),
            Dias_Findes=("Es_fin_de_semana", lambda x: int((x == "SI").sum()) if "Es_fin_de_semana" in daily.columns else 0),
            Dias_Madrugada=("Ajuste_madrugada", lambda x: int((x == "SI").sum()) if "Ajuste_madrugada" in daily.columns else 0),
        )
        .sort_values(["Tipo", "Empleado", "DNI"])
        .reset_index(drop=True)
    )

    s["DNI"] = s["DNI"].astype(str)
    s["Total"] = s["Total_min"].round().astype(int).apply(minutes_to_hhmm)
    s["Prom/día"] = s["Prom_min"].round().astype(int).apply(minutes_to_hhmm)
    s["Extras"] = s["Extras_min"].apply(minutes_to_hhmm)
    s["Faltas"] = s["Faltas_min"].apply(minutes_to_hhmm)
    s["Normal"] = s["Normal_min"].apply(minutes_to_hhmm) if "Normal_min" in s.columns else "00:00"
    s["Extra_dia"] = s["Extra_dia_min"].apply(minutes_to_hhmm) if "Extra_dia_min" in s.columns else "00:00"
    s["Faltante_dia"] = s["Faltante_dia_min"].apply(minutes_to_hhmm) if "Faltante_dia_min" in s.columns else "00:00"
    s["Saldo"] = s["Saldo_min"].apply(delta_short)

    def pct_row(r):
        if r["Tipo"] == "NO Docente" and r["Esperado_min"] > 0:
            return f"{(r['Total_min'] / r['Esperado_min'] * 100):.0f}%"
        return ""

    s["Cumplimiento"] = s.apply(pct_row, axis=1)

    cols = [
        "EmployeeKey",
        "Empleado", "DNI", "Tipo",
        "Dias",
        "Total", "Total_min",
        "Prom/día",
        "Normal", "Normal_min",
        "Esperado_min",
        "Extras", "Extras_min",
        "Extra_dia", "Extra_dia_min",
        "Faltas", "Faltas_min",
        "Faltante_dia", "Faltante_dia_min",
        "Saldo", "Saldo_min",
        "Cumplimiento",
        "Marcaciones",
        "Incompletos",
        "Cortes",
        "Dias_OK", "Dias_FALTA", "Dias_INCOMPL", "Dias_EXTRA",
        "Dias_Feriado", "Dias_Findes", "Dias_Madrugada",
    ]
    cols = [c for c in cols if c in s.columns]
    return s[cols]


def employee_detail_table(daily_emp: pd.DataFrame) -> pd.DataFrame:
    d = daily_emp.copy()
    d["Fecha"] = pd.to_datetime(d["Fecha"]).dt.date
    d["Primera"] = pd.to_datetime(d["Primera"], errors="coerce").dt.strftime("%H:%M")
    d["Ultima"] = pd.to_datetime(d["Ultima"], errors="coerce").dt.strftime("%H:%M")

    cols = [
        "Fecha", "Tipo_dia", "Primera", "Ultima",
        "Marcaciones_detalle", "Tramos_detalle",
        "Horas", "Normal", "Extra_dia", "Faltante_dia", "Esperado", "Saldo",
        "Marcaciones", "Pares_estimados", "Cortes", "Incompleto",
        "Cumple", "Ajuste_madrugada"
    ]
    cols = [c for c in cols if c in d.columns]
    return d[cols].sort_values("Fecha").reset_index(drop=True)


def raw_employee_marks_table(raw_emp: pd.DataFrame) -> pd.DataFrame:
    if raw_emp.empty:
        return pd.DataFrame()
    out = raw_emp.copy()
    out["Fecha_laboral"] = pd.to_datetime(out["Fecha"]).dt.strftime("%Y-%m-%d")
    out["Fecha_real"] = pd.to_datetime(out["FechaReal"]).dt.strftime("%Y-%m-%d")
    out["Hora"] = pd.to_datetime(out["FechaHora"], errors="coerce").dt.strftime("%H:%M")
    cols = ["Empleado", "DNI", "Tipo", "Fecha_laboral", "Fecha_real", "Hora", "Ajuste_madrugada", "NvoEstado"]
    return out[[c for c in cols if c in out.columns]].reset_index(drop=True)


def pretty_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df

    out = df.copy()
    if "EmployeeKey" in out.columns:
        out = out.drop(columns=["EmployeeKey"])
    if "DNI" in out.columns:
        out["DNI"] = out["DNI"].apply(display_dni)
    return out


def build_extras_only(summary: pd.DataFrame) -> pd.DataFrame:
    if summary.empty:
        return pd.DataFrame(columns=["Empleado", "DNI", "Tipo", "Horas_extras", "Extras_min"])
    extras_only = summary[summary["Tipo"] == "NO Docente"].copy()
    if extras_only.empty:
        return pd.DataFrame(columns=["Empleado", "DNI", "Tipo", "Horas_extras", "Extras_min"])
    extras_only = (
        extras_only[["Empleado", "DNI", "Tipo", "Extras", "Extras_min"]]
        .rename(columns={"Extras": "Horas_extras"})
        .sort_values("Extras_min", ascending=False)
        .reset_index(drop=True)
    )
    extras_only["DNI"] = extras_only["DNI"].apply(display_dni)
    return extras_only


def build_rankings(summary: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    if summary.empty:
        return pd.DataFrame(), pd.DataFrame()

    ranking_hours = summary.copy().sort_values("Total_min", ascending=False).reset_index(drop=True)
    ranking_hours = pretty_summary(ranking_hours).head(30)

    ranking_extras = summary.copy().sort_values("Extras_min", ascending=False).reset_index(drop=True)
    ranking_extras = pretty_summary(ranking_extras).head(30)

    return ranking_hours, ranking_extras


def build_inconsistencies(daily: pd.DataFrame) -> pd.DataFrame:
    if daily.empty:
        return pd.DataFrame()

    mask = (
        (daily.get("Incompleto", "") == "SI") |
        (daily.get("Cortes", "") == "SI") |
        (daily.get("Cumple", "") == "FALTA")
    )
    out = daily[mask].copy()
    if out.empty:
        return pd.DataFrame(columns=[
            "Fecha", "Empleado", "DNI", "Tipo", "Tipo_dia", "Horas", "Esperado",
            "Saldo", "Marcaciones", "Pares_estimados", "Cortes", "Incompleto", "Cumple"
        ])
    out["Fecha"] = pd.to_datetime(out["Fecha"]).dt.date
    cols = [
        "Fecha", "Empleado", "DNI", "Tipo", "Tipo_dia", "Marcaciones_detalle", "Tramos_detalle",
        "Horas", "Normal", "Extra_dia", "Faltante_dia", "Esperado",
        "Saldo", "Marcaciones", "Pares_estimados", "Cortes", "Incompleto", "Cumple",
        "Ajuste_madrugada"
    ]
    out = out[[c for c in cols if c in out.columns]]
    out["DNI"] = out["DNI"].apply(display_dni)
    return out.sort_values(["Fecha", "Empleado"]).reset_index(drop=True)


def calculate_general_indicators(daily: pd.DataFrame, summary: pd.DataFrame) -> dict:
    if daily is None or daily.empty:
        return {
            "dias_empleado": 0, "dias_ok": 0, "dias_falta": 0, "dias_extra": 0,
            "dias_incompletos": 0, "dias_cortes": 0, "dias_feriado": 0,
            "dias_findes": 0, "dias_madrugada": 0, "empleados_con_extra": 0,
            "empleados_con_falta": 0, "prom_extra_por_empleado": 0,
            "prom_horas_por_dia_empleado": 0, "mayor_extra_min": 0, "mayor_total_min": 0,
        }

    nod = daily[daily["Tipo"] == "NO Docente"].copy() if "Tipo" in daily.columns else daily.copy()
    dias_empleado = int(daily.shape[0])
    dias_ok = int((daily.get("Cumple", "") == "OK").sum())
    dias_falta = int((daily.get("Cumple", "") == "FALTA").sum())
    dias_extra = int((daily.get("Cumple", "") == "EXTRA").sum())
    dias_incompletos = int((daily.get("Incompleto", "") == "SI").sum())
    dias_cortes = int((daily.get("Cortes", "") == "SI").sum())
    dias_feriado = int((daily.get("Es_feriado", "") == "SI").sum()) if "Es_feriado" in daily.columns else 0
    dias_findes = int((daily.get("Es_fin_de_semana", "") == "SI").sum()) if "Es_fin_de_semana" in daily.columns else 0
    dias_madrugada = int((daily.get("Ajuste_madrugada", "") == "SI").sum()) if "Ajuste_madrugada" in daily.columns else 0

    empleados_con_extra = 0
    empleados_con_falta = 0
    prom_extra_por_empleado = 0
    mayor_extra_min = 0
    mayor_total_min = 0
    if summary is not None and not summary.empty:
        empleados_con_extra = int((summary.get("Extras_min", 0) > 0).sum())
        empleados_con_falta = int((summary.get("Faltas_min", 0) > 0).sum())
        prom_extra_por_empleado = int(round(summary.get("Extras_min", pd.Series(dtype=int)).mean())) if "Extras_min" in summary.columns else 0
        mayor_extra_min = int(summary.get("Extras_min", pd.Series([0])).max()) if "Extras_min" in summary.columns else 0
        mayor_total_min = int(summary.get("Total_min", pd.Series([0])).max()) if "Total_min" in summary.columns else 0

    prom_horas_por_dia_empleado = int(round(daily.get("Minutos", pd.Series(dtype=int)).mean())) if "Minutos" in daily.columns else 0

    return {
        "dias_empleado": dias_empleado,
        "dias_ok": dias_ok,
        "dias_falta": dias_falta,
        "dias_extra": dias_extra,
        "dias_incompletos": dias_incompletos,
        "dias_cortes": dias_cortes,
        "dias_feriado": dias_feriado,
        "dias_findes": dias_findes,
        "dias_madrugada": dias_madrugada,
        "empleados_con_extra": empleados_con_extra,
        "empleados_con_falta": empleados_con_falta,
        "prom_extra_por_empleado": prom_extra_por_empleado,
        "prom_horas_por_dia_empleado": prom_horas_por_dia_empleado,
        "mayor_extra_min": mayor_extra_min,
        "mayor_total_min": mayor_total_min,
    }


def build_daily_explained_table(daily: pd.DataFrame) -> pd.DataFrame:
    if daily is None or daily.empty:
        return pd.DataFrame()

    out = daily.copy()
    out["Fecha"] = pd.to_datetime(out["Fecha"]).dt.date
    out["DNI"] = out["DNI"].apply(display_dni)

    def explanation(r):
        tipo_dia = str(r.get("Tipo_dia", ""))
        horas = str(r.get("Horas", "00:00"))
        normal = str(r.get("Normal", "00:00"))
        extra = str(r.get("Extra_dia", "00:00"))
        faltante = str(r.get("Faltante_dia", "00:00"))
        marcas = str(r.get("Marcaciones_detalle", ""))
        tramos = str(r.get("Tramos_detalle", ""))

        if r.get("Incompleto", "") == "SI":
            return f"REVISAR: CANTIDAD IMPAR O FALTA DE MARCACIÓN. MARCAS: {marcas}"
        if tipo_dia in ["FERIADO", "FIN DE SEMANA", "FERIADO/FIN DE SEMANA"]:
            return f"DÍA ESPECIAL: TODO LO TRABAJADO ({horas}) CUENTA COMO EXTRA. MARCAS: {marcas}"
        if extra != "00:00":
            return f"TRABAJÓ {horas}. PRIMERO SE CUBRE LA JORNADA NORMAL ({normal}) Y EL RESTO ES EXTRA ({extra}). TRAMOS: {tramos}"
        if faltante != "00:00":
            return f"TRABAJÓ {horas}. LE FALTÓ {faltante} PARA COMPLETAR LA JORNADA. TRAMOS: {tramos}"
        return f"TRABAJÓ {horas}. CUMPLIÓ LA JORNADA ESPERADA. TRAMOS: {tramos}"

    out["Explicacion_calculo"] = out.apply(explanation, axis=1)

    cols = [
        "Fecha", "Empleado", "DNI", "Tipo", "Tipo_dia",
        "Marcaciones_detalle", "Tramos_detalle",
        "Horas", "Normal", "Extra_dia", "Faltante_dia", "Esperado", "Saldo",
        "Marcaciones", "Pares_estimados", "Cortes", "Incompleto", "Cumple",
        "Ajuste_madrugada", "Explicacion_calculo"
    ]
    return out[[c for c in cols if c in out.columns]].sort_values(["Fecha", "Empleado"]).reset_index(drop=True)


def build_employee_explained_stats(daily_emp: pd.DataFrame) -> dict:
    if daily_emp is None or daily_emp.empty:
        return {}
    total = int(daily_emp["Minutos"].sum())
    normal = int(daily_emp.get("Normal_min", pd.Series([0]*len(daily_emp))).sum()) if "Normal_min" in daily_emp.columns else 0
    extra = int(daily_emp.get("Extra_dia_min", pd.Series([0]*len(daily_emp))).sum()) if "Extra_dia_min" in daily_emp.columns else int(daily_emp.loc[daily_emp["Saldo_min"] > 0, "Saldo_min"].sum())
    faltante = int(daily_emp.get("Faltante_dia_min", pd.Series([0]*len(daily_emp))).sum()) if "Faltante_dia_min" in daily_emp.columns else int((-daily_emp.loc[daily_emp["Saldo_min"] < 0, "Saldo_min"].sum()))
    dias = int(daily_emp["Fecha"].nunique())
    marcas = int(daily_emp["Marcaciones"].sum())
    cortes = int((daily_emp.get("Cortes", "") == "SI").sum())
    incompletos = int((daily_emp.get("Incompleto", "") == "SI").sum())
    feriados = int((daily_emp.get("Es_feriado", "") == "SI").sum()) if "Es_feriado" in daily_emp.columns else 0
    findes = int((daily_emp.get("Es_fin_de_semana", "") == "SI").sum()) if "Es_fin_de_semana" in daily_emp.columns else 0
    madrugada = int((daily_emp.get("Ajuste_madrugada", "") == "SI").sum()) if "Ajuste_madrugada" in daily_emp.columns else 0
    prom = int(round(daily_emp["Minutos"].mean())) if dias else 0

    return {
        "total": total, "normal": normal, "extra": extra, "faltante": faltante,
        "dias": dias, "marcas": marcas, "cortes": cortes, "incompletos": incompletos,
        "feriados": feriados, "findes": findes, "madrugada": madrugada, "prom": prom
    }



def average_clock_time(values, overnight_as_next_day: bool = False) -> str:
    timestamps = pd.to_datetime(pd.Series(list(values)), errors="coerce").dropna()
    if timestamps.empty:
        return "—"

    minutes = []
    for ts in timestamps:
        minute = int(ts.hour) * 60 + int(ts.minute)
        if overnight_as_next_day and minute < WORKDAY_CUTOFF_HOUR * 60:
            minute += 24 * 60
        minutes.append(minute)

    average = int(round(sum(minutes) / len(minutes)))
    next_day = average >= 24 * 60
    average = average % (24 * 60)
    text = f"{average // 60:02d}:{average % 60:02d}"
    return f"{text} +1D" if next_day else text


def period_dates(raw: pd.DataFrame) -> tuple[date | None, date | None]:
    if raw is None or raw.empty or "Fecha" not in raw.columns:
        return None, None
    dates = pd.to_datetime(raw["Fecha"], errors="coerce").dropna()
    if dates.empty:
        return None, None
    return dates.min().date(), dates.max().date()


def expected_workdays(start: date | None, end: date | None, holidays: set[date]) -> list[date]:
    if start is None or end is None or end < start:
        return []

    result = []
    for ts in pd.date_range(start=start, end=end, freq="D"):
        current = ts.date()
        if ts.weekday() < 5 and current not in holidays:
            result.append(current)
    return result


def missing_workdays_by_employee(raw: pd.DataFrame, holidays: set[date]) -> dict[str, list[date]]:
    start, end = period_dates(raw)
    expected_dates = set(expected_workdays(start, end, holidays))
    result: dict[str, list[date]] = {}

    if raw is None or raw.empty:
        return result

    for employee_key, group in raw.groupby("EmployeeKey"):
        marked_dates = set(group["Fecha"].dropna().tolist())
        result[str(employee_key)] = sorted(expected_dates - marked_dates)

    return result


def compact_employee_summary(
    raw: pd.DataFrame,
    daily: pd.DataFrame,
    summary: pd.DataFrame,
    holidays: set[date],
) -> pd.DataFrame:
    if summary is None or summary.empty:
        return pd.DataFrame()

    missing_map = missing_workdays_by_employee(raw, holidays)
    rows = []

    for _, item in summary.iterrows():
        employee_key = str(item["EmployeeKey"])
        daily_emp = daily[daily["EmployeeKey"].astype(str) == employee_key].copy()

        valid_days = daily_emp[daily_emp["Incompleto"] != "SI"].copy()

        arrival_average = (
            average_clock_time(valid_days["Primera"])
            if not valid_days.empty
            else "—"
        )
        departure_average = (
            average_clock_time(valid_days["Ultima"], overnight_as_next_day=True)
            if not valid_days.empty
            else "—"
        )

        gross_extra = (
            int(valid_days["Extra_dia_min"].sum())
            if not valid_days.empty and "Extra_dia_min" in valid_days.columns
            else 0
        )
        missing_minutes = (
            int(valid_days["Faltante_dia_min"].sum())
            if not valid_days.empty and "Faltante_dia_min" in valid_days.columns
            else 0
        )

        net_balance = gross_extra - missing_minutes
        extra_favor = max(net_balance, 0)
        debt_pending = max(-net_balance, 0)

        normal_minutes = (
            int(daily_emp["Normal_min"].sum())
            if "Normal_min" in daily_emp.columns
            else int(item.get("Normal_min", 0))
        )
        total_minutes = (
            int(daily_emp["Minutos"].sum())
            if not daily_emp.empty
            else int(item.get("Total_min", 0))
        )
        expected_minutes = (
            int(valid_days["Esperado_min"].sum())
            if not valid_days.empty and "Esperado_min" in valid_days.columns
            else 0
        )

        incomplete_days = int((daily_emp["Incompleto"] == "SI").sum()) if not daily_emp.empty else 0
        days_with_extra = int((valid_days["Extra_dia_min"] > 0).sum()) if "Extra_dia_min" in valid_days.columns else 0
        days_with_missing_hours = int((valid_days["Faltante_dia_min"] > 0).sum()) if "Faltante_dia_min" in valid_days.columns else 0
        worked_days = int((daily_emp["Minutos"] > 0).sum()) if not daily_emp.empty else 0
        missing_days = len(missing_map.get(employee_key, []))
        duplicates_ignored = (
            int(daily_emp["Duplicadas_ignoradas"].sum())
            if "Duplicadas_ignoradas" in daily_emp.columns
            else 0
        )

        if incomplete_days > 0:
            status = "REVISAR MARCAS"
        elif net_balance < 0:
            status = "DEBE COMPENSAR"
        elif missing_days > 0:
            status = "REVISAR AUSENCIAS"
        elif net_balance > 0:
            status = "SALDO A FAVOR"
        else:
            status = "BALANCE EN CERO"

        rows.append(
            {
                "Empleado": item["Empleado"],
                "DNI": display_dni(item["DNI"]),
                "Tipo": item["Tipo"],

                # El dato principal para RRHH luego de compensar faltantes.
                "Horas_extra": minutes_to_hhmm(extra_favor),
                "Extras_min": extra_favor,

                # Transparencia del cálculo.
                "Extra_bruta": minutes_to_hhmm(gross_extra),
                "Extras_brutas_min": gross_extra,
                "Horas_faltantes": minutes_to_hhmm(missing_minutes),
                "Faltantes_min": missing_minutes,
                "Saldo_periodo": delta_short(net_balance),
                "Saldo_neto_min": net_balance,
                "Deuda_pendiente": minutes_to_hhmm(debt_pending),
                "Deuda_pendiente_min": debt_pending,

                "Días_con_extra": days_with_extra,
                "Días_con_horas_faltantes": days_with_missing_hours,
                "Días_trabajados": worked_days,
                "Días_sin_marcación": missing_days,
                "Días_marca_incompleta": incomplete_days,
                "Entrada_promedio": arrival_average,
                "Salida_promedio": departure_average,
                "Total_trabajado": minutes_to_hhmm(total_minutes),
                "Esperado_periodo": minutes_to_hhmm(expected_minutes),
                "Horas_normales": minutes_to_hhmm(normal_minutes),
                "Marcaciones": int(daily_emp["Marcaciones"].sum()) if not daily_emp.empty else 0,
                "Dobles_marcas_ignoradas": duplicates_ignored,
                "Estado": status,
                "EmployeeKey": employee_key,
            }
        )

    result = pd.DataFrame(rows)
    return result.sort_values(
        ["Saldo_neto_min", "Empleado"],
        ascending=[False, True],
    ).reset_index(drop=True)


def compact_daily_table(daily: pd.DataFrame) -> pd.DataFrame:
    if daily is None or daily.empty:
        return pd.DataFrame()

    out = daily.copy()
    out["Fecha"] = pd.to_datetime(out["Fecha"]).dt.date
    out["DNI"] = out["DNI"].apply(display_dni)
    out["Primera"] = pd.to_datetime(out["Primera"], errors="coerce").dt.strftime("%H:%M")
    out["Ultima"] = pd.to_datetime(out["Ultima"], errors="coerce").dt.strftime("%H:%M")

    def daily_status(row) -> str:
        if row.get("Incompleto", "") == "SI":
            base = "REVISAR MARCACIÓN"
        elif int(row.get("Extra_dia_min", 0)) > 0:
            base = "CON HORAS EXTRA"
        elif int(row.get("Faltante_dia_min", 0)) > 0:
            base = "HORAS FALTANTES"
        else:
            base = "OK"

        if int(row.get("Duplicadas_ignoradas", 0)) > 0:
            base += " · DOBLE MARCA IGNORADA"
        return base

    out["Estado_día"] = out.apply(daily_status, axis=1)

    columns = [
        "Fecha",
        "Empleado",
        "DNI",
        "Tipo_dia",
        "Primera",
        "Ultima",
        "Marcaciones_detalle",
        "Marcaciones_efectivas_detalle",
        "Tramos_detalle",
        "Horas",
        "Normal",
        "Extra_dia",
        "Faltante_dia",
        "Saldo",
        "Marcaciones",
        "Marcaciones_efectivas",
        "Duplicadas_ignoradas",
        "Duplicadas_detalle",
        "Estado_día",
        "Ajuste_madrugada",
    ]
    return out[[c for c in columns if c in out.columns]].sort_values(
        ["Fecha", "Empleado"]
    ).reset_index(drop=True)


def compact_review_table(daily: pd.DataFrame) -> pd.DataFrame:
    if daily is None or daily.empty:
        return pd.DataFrame()

    mask = (daily["Incompleto"] == "SI")
    if "Faltante_dia_min" in daily.columns:
        mask = mask | (daily["Faltante_dia_min"] > 0)

    review = compact_daily_table(daily[mask].copy())
    return review


def compact_employee_metrics(
    raw: pd.DataFrame,
    daily_emp: pd.DataFrame,
    employee_key: str,
    holidays: set[date],
) -> dict:
    missing_map = missing_workdays_by_employee(raw, holidays)

    base = {
        "extra": 0,
        "gross_extra": 0,
        "total": 0,
        "total_valid": 0,
        "normal": 0,
        "missing_minutes": 0,
        "net_balance": 0,
        "extra_favor": 0,
        "debt_pending": 0,
        "expected": 0,
        "worked_days": 0,
        "missing_days": len(missing_map.get(employee_key, [])),
        "incomplete_days": 0,
        "days_with_extra": 0,
        "days_with_missing": 0,
        "balanced_days": 0,
        "duplicates": 0,
        "largest_extra": 0,
        "largest_missing": 0,
        "coverage_pct": 0.0,
        "arrival": "—",
        "departure": "—",
    }

    if daily_emp is None or daily_emp.empty:
        return base

    # Días incompletos quedan FUERA del balance económico hasta ser revisados/corregidos.
    valid_days = daily_emp[daily_emp["Incompleto"] != "SI"].copy()

    gross_extra = (
        int(valid_days["Extra_dia_min"].sum())
        if not valid_days.empty and "Extra_dia_min" in valid_days.columns
        else 0
    )
    missing_minutes = (
        int(valid_days["Faltante_dia_min"].sum())
        if not valid_days.empty and "Faltante_dia_min" in valid_days.columns
        else 0
    )

    net_balance = gross_extra - missing_minutes
    extra_favor = max(net_balance, 0)
    debt_pending = max(-net_balance, 0)

    expected_minutes = (
        int(valid_days["Esperado_min"].sum())
        if not valid_days.empty and "Esperado_min" in valid_days.columns
        else 0
    )
    total_valid = (
        int(valid_days["Minutos"].sum())
        if not valid_days.empty and "Minutos" in valid_days.columns
        else 0
    )

    # El porcentaje puede superar 100% cuando existen extras.
    coverage_pct = (
        (total_valid / expected_minutes * 100)
        if expected_minutes > 0
        else 0.0
    )

    return {
        # compatibilidad con código anterior: "extra" pasa a representar extra A FAVOR.
        "extra": extra_favor,
        "gross_extra": gross_extra,
        "total": int(daily_emp["Minutos"].sum()),
        "total_valid": total_valid,
        "normal": int(daily_emp["Normal_min"].sum()) if "Normal_min" in daily_emp.columns else 0,
        "missing_minutes": missing_minutes,
        "net_balance": net_balance,
        "extra_favor": extra_favor,
        "debt_pending": debt_pending,
        "expected": expected_minutes,
        "worked_days": int((daily_emp["Minutos"] > 0).sum()),
        "missing_days": len(missing_map.get(employee_key, [])),
        "incomplete_days": int((daily_emp["Incompleto"] == "SI").sum()),
        "days_with_extra": int((valid_days["Extra_dia_min"] > 0).sum()) if "Extra_dia_min" in valid_days.columns else 0,
        "days_with_missing": int((valid_days["Faltante_dia_min"] > 0).sum()) if "Faltante_dia_min" in valid_days.columns else 0,
        "balanced_days": int((valid_days["Saldo_min"] == 0).sum()) if "Saldo_min" in valid_days.columns else 0,
        "duplicates": int(daily_emp["Duplicadas_ignoradas"].sum()) if "Duplicadas_ignoradas" in daily_emp.columns else 0,
        "largest_extra": int(valid_days["Extra_dia_min"].max()) if not valid_days.empty and "Extra_dia_min" in valid_days.columns else 0,
        "largest_missing": int(valid_days["Faltante_dia_min"].max()) if not valid_days.empty and "Faltante_dia_min" in valid_days.columns else 0,
        "coverage_pct": coverage_pct,
        "arrival": average_clock_time(valid_days["Primera"]) if not valid_days.empty else "—",
        "departure": average_clock_time(valid_days["Ultima"], overnight_as_next_day=True) if not valid_days.empty else "—",
    }


def build_employee_period_table(
    raw: pd.DataFrame,
    daily_emp: pd.DataFrame,
    employee_key: str,
    holidays: set[date],
) -> pd.DataFrame:
    """
    Planilla continua del período.

    Además del saldo diario, agrega SALDO ACUMULADO:
    cada saldo positivo compensa saldos negativos anteriores y viceversa.
    Los días sin marcación o incompletos quedan para revisión y NO alteran
    automáticamente el balance.
    """
    start, end = period_dates(raw)
    actual = compact_daily_table(daily_emp)

    if start is None or end is None:
        return actual

    expected_dates = set(expected_workdays(start, end, holidays))
    actual_dates = set(actual["Fecha"].tolist()) if not actual.empty else set()
    all_dates = sorted(expected_dates | actual_dates)

    if daily_emp is not None and not daily_emp.empty:
        emp = str(daily_emp.iloc[0]["Empleado"])
        dni = display_dni(daily_emp.iloc[0]["DNI"])
    else:
        raw_emp = raw[raw["EmployeeKey"].astype(str) == str(employee_key)]
        emp = str(raw_emp.iloc[0]["Empleado"]) if not raw_emp.empty else ""
        dni = display_dni(raw_emp.iloc[0]["DNI"]) if not raw_emp.empty else ""

    actual_map = {}
    if not actual.empty:
        for _, row in actual.iterrows():
            actual_map[row["Fecha"]] = row.to_dict()

    # Mapa numérico para poder acumular el saldo de días confirmados.
    saldo_map = {}
    incomplete_map = {}
    if daily_emp is not None and not daily_emp.empty:
        for _, row in daily_emp.iterrows():
            day = pd.to_datetime(row["Fecha"]).date()
            saldo_map[day] = int(row.get("Saldo_min", 0))
            incomplete_map[day] = str(row.get("Incompleto", "")) == "SI"

    rows = []
    running_balance = 0

    for day in all_dates:
        if day in actual_map:
            row = dict(actual_map[day])

            if not incomplete_map.get(day, False):
                running_balance += int(saldo_map.get(day, 0))
                row["Saldo_acumulado"] = delta_short(running_balance)
            else:
                row["Saldo_acumulado"] = f"{delta_short(running_balance)} · REVISAR"

            rows.append(row)
            continue

        rows.append(
            {
                "Fecha": day,
                "Empleado": emp,
                "DNI": dni,
                "Tipo_dia": day_type_label(day, holidays),
                "Primera": "",
                "Ultima": "",
                "Marcaciones_detalle": "",
                "Marcaciones_efectivas_detalle": "",
                "Tramos_detalle": "",
                "Horas": "00:00",
                "Normal": "00:00",
                "Extra_dia": "00:00",
                "Faltante_dia": "—",
                "Saldo": "REVISAR",
                "Saldo_acumulado": f"{delta_short(running_balance)} · SIN CAMBIO",
                "Marcaciones": 0,
                "Marcaciones_efectivas": 0,
                "Duplicadas_ignoradas": 0,
                "Duplicadas_detalle": "",
                "Estado_día": "SIN MARCACIÓN · REVISAR LICENCIA/AUSENCIA",
                "Ajuste_madrugada": "",
            }
        )

    result = pd.DataFrame(rows)

    preferred_order = [
        "Fecha", "Empleado", "DNI", "Tipo_dia",
        "Primera", "Ultima",
        "Marcaciones_detalle", "Marcaciones_efectivas_detalle", "Tramos_detalle",
        "Horas", "Normal", "Extra_dia", "Faltante_dia",
        "Saldo", "Saldo_acumulado",
        "Marcaciones", "Marcaciones_efectivas",
        "Duplicadas_ignoradas", "Duplicadas_detalle",
        "Estado_día", "Ajuste_madrugada",
    ]
    return result[[c for c in preferred_order if c in result.columns]]


def export_printable_employee_workbook(
    raw: pd.DataFrame,
    daily: pd.DataFrame,
    compact_summary: pd.DataFrame,
    employee_keys: list[str],
    holidays: set[date],
    period_text: str,
) -> bytes:
    """
    Genera una única hoja preparada para imprimir.
    Cada empleado comienza en una página nueva y usa exactamente la misma tabla.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Planillas"

    blue = "0A58CA"
    dark_blue = "002B63"
    light_blue = "DCEBFA"
    light_gray = "EEF2F6"
    white = "FFFFFF"
    thin_gray = Side(style="thin", color="B7C3D0")
    border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)

    headers = [
        "FECHA",
        "TIPO DÍA",
        "PRIMERA",
        "ÚLTIMA",
        "TODAS LAS MARCACIONES",
        "MARCAS USADAS",
        "TRAMOS CALCULADOS",
        "TRABAJADO",
        "NORMAL",
        "EXTRA",
        "FALTANTE",
        "SALDO DÍA",
        "SALDO ACUM.",
        "ESTADO",
        "DOBLES IGNORADAS",
    ]

    current_row = 1
    first_employee = True

    for employee_key in employee_keys:
        summary_match = compact_summary[
            compact_summary["EmployeeKey"].astype(str) == str(employee_key)
        ]
        daily_emp = daily[
            daily["EmployeeKey"].astype(str) == str(employee_key)
        ].copy()

        if summary_match.empty:
            continue

        s = summary_match.iloc[0]
        table = build_employee_period_table(
            raw=raw,
            daily_emp=daily_emp,
            employee_key=str(employee_key),
            holidays=holidays,
        )

        if not first_employee:
            ws.row_breaks.append(Break(id=current_row - 1))
        first_employee = False

        start_row = current_row

        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(headers))
        title_cell = ws.cell(current_row, 1, "CONTROL DE ASISTENCIA · PLANILLA DEL EMPLEADO")
        title_cell.font = Font(bold=True, color=white, size=16)
        title_cell.fill = PatternFill("solid", fgColor=dark_blue)
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[current_row].height = 26
        current_row += 1

        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(headers))
        info = (
            f"EMPLEADO: {s['Empleado']}   ·   DNI: {s['DNI']}   ·   TIPO: {s['Tipo']}   ·   PERÍODO: {period_text}"
        )
        c = ws.cell(current_row, 1, info)
        c.font = Font(bold=True, size=10)
        c.fill = PatternFill("solid", fgColor=light_blue)
        c.alignment = Alignment(horizontal="left")
        current_row += 1

        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(headers))
        totals = (
            f"EXTRA A FAVOR: {s['Horas_extra']}   ·   EXTRA GENERADA: {s.get('Extra_bruta', '00:00')}   ·   "
            f"A COMPENSAR: {s.get('Horas_faltantes', '00:00')}   ·   SALDO: {s.get('Saldo_periodo', '0M')}   ·   "
            f"TOTAL TRABAJADO: {s['Total_trabajado']}   ·   DÍAS SIN MARCACIÓN: {s['Días_sin_marcación']}   ·   "
            f"MARCAS INCOMPLETAS: {s['Días_marca_incompleta']}"
        )
        c = ws.cell(current_row, 1, totals)
        c.font = Font(bold=True, size=9)
        c.alignment = Alignment(horizontal="left", wrap_text=True)
        current_row += 1

        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(headers))
        legend = (
            "SALDO DÍA = TRABAJADO MENOS ESPERADO. SALDO ACUMULADO = SUMA CORRIDA DE LOS SALDOS DEL PERÍODO; "
            "LOS EXCEDENTES POSTERIORES PUEDEN COMPENSAR FALTANTES ANTERIORES. "
            f"DOBLE MARCA = DOS LECTURAS A {DUPLICATE_PUNCH_WINDOW_MINUTES} MINUTOS O MENOS; "
            "SE MUESTRAN TODAS, PERO PARA EL CÁLCULO SE USA UN SOLO EVENTO."
        )
        c = ws.cell(current_row, 1, legend)
        c.font = Font(italic=True, size=8, color="44546A")
        c.alignment = Alignment(horizontal="left", wrap_text=True)
        current_row += 2

        header_row = current_row
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(current_row, col_idx, header)
            cell.font = Font(bold=True, color=white, size=8)
            cell.fill = PatternFill("solid", fgColor=blue)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        ws.row_dimensions[current_row].height = 32
        current_row += 1

        for _, row in table.iterrows():
            values = [
                row.get("Fecha", ""),
                row.get("Tipo_dia", ""),
                row.get("Primera", ""),
                row.get("Ultima", ""),
                row.get("Marcaciones_detalle", ""),
                row.get("Marcaciones_efectivas_detalle", ""),
                row.get("Tramos_detalle", ""),
                row.get("Horas", ""),
                row.get("Normal", ""),
                row.get("Extra_dia", ""),
                row.get("Faltante_dia", ""),
                row.get("Saldo", ""),
                row.get("Saldo_acumulado", ""),
                row.get("Estado_día", ""),
                row.get("Duplicadas_ignoradas", 0),
            ]

            for col_idx, value in enumerate(values, start=1):
                if isinstance(value, date):
                    value = value.strftime("%d/%m/%Y")
                cell = ws.cell(current_row, col_idx, value)
                cell.border = border
                cell.alignment = Alignment(
                    horizontal="center" if col_idx not in [5, 6, 7, 14] else "left",
                    vertical="top",
                    wrap_text=True,
                )
                cell.font = Font(size=8)

                if "SIN MARCACIÓN" in str(row.get("Estado_día", "")):
                    cell.fill = PatternFill("solid", fgColor="FFF2CC")
                elif int(row.get("Duplicadas_ignoradas", 0) or 0) > 0:
                    cell.fill = PatternFill("solid", fgColor="E2F0D9")
                elif str(row.get("Estado_día", "")).startswith("REVISAR"):
                    cell.fill = PatternFill("solid", fgColor="FCE4D6")
                elif current_row % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor=light_gray)

            current_row += 1

        current_row += 2

    widths = {
        "A": 11, "B": 17, "C": 9, "D": 9,
        "E": 25, "F": 23, "G": 32,
        "H": 11, "I": 10, "J": 10, "K": 10,
        "L": 11, "M": 13, "N": 28, "O": 12,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = 0.2
    ws.page_margins.right = 0.2
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35
    ws.oddFooter.center.text = "Página &P de &N"
    ws.oddFooter.right.text = "Control de Asistencia"

    if current_row > 1:
        ws.print_area = f"A1:O{current_row - 1}"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def export_compact_excel(
    compact_summary: pd.DataFrame,
    daily_table: pd.DataFrame,
    review_table: pd.DataFrame,
    raw: pd.DataFrame,
    period_text: str,
    total_extra: int,
) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    def add_df(sheet_name: str, df: pd.DataFrame, text_cols: set[str] | None = None):
        text_cols = text_cols or set()
        ws = wb.create_sheet(sheet_name)

        if df is None or df.empty:
            df = pd.DataFrame(columns=["Sin_datos"])

        for col_idx, column in enumerate(df.columns, start=1):
            ws.cell(row=1, column=col_idx, value=str(column))

        for row_idx, row in enumerate(df.itertuples(index=False), start=2):
            for col_idx, value in enumerate(row, start=1):
                column = df.columns[col_idx - 1]
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                if column in text_cols and value is not None:
                    cell.value = str(value)
                    cell.number_format = "@"

        _apply_excel_style(ws, table_name=sheet_name)

    control = pd.DataFrame(
        [{
            "Periodo": period_text,
            "Extra_a_favor_total": minutes_to_hhmm(total_extra),
            "Criterio_ausencias": "DÍAS HÁBILES DEL PERÍODO SIN MARCACIÓN; REVISAR LICENCIAS Y JUSTIFICACIONES.",
        }]
    )
    add_df("Control", control)
    add_df("Resumen_Empleados", compact_summary.drop(columns=["EmployeeKey", "Extras_min"], errors="ignore"), {"DNI"})
    add_df("Detalle_Diario", daily_table, {"DNI"})
    add_df("Revisar", review_table, {"DNI"})

    raw_out = raw.copy()
    if not raw_out.empty:
        raw_out["DNI"] = raw_out["DNI"].apply(display_dni)
        raw_out["Fecha_laboral"] = pd.to_datetime(raw_out["Fecha"]).dt.strftime("%Y-%m-%d")
        raw_out["Fecha_real"] = pd.to_datetime(raw_out["FechaReal"]).dt.strftime("%Y-%m-%d")
        raw_out["Hora"] = pd.to_datetime(raw_out["FechaHora"]).dt.strftime("%H:%M")
        raw_out = raw_out[
            [
                "Empleado",
                "DNI",
                "Tipo",
                "Fecha_laboral",
                "Fecha_real",
                "Hora",
                "NvoEstado",
                "Ajuste_madrugada",
            ]
        ]
    add_df("Marcaciones", raw_out, {"DNI"})

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ============================================================
# EXPORT EXCEL
# ============================================================
def _safe_table_name(name: str) -> str:
    base = re.sub(r"[^A-Za-z0-9_]", "_", name)
    if not base:
        base = "Tabla"
    if not re.match(r"^[A-Za-z_]", base):
        base = f"_{base}"
    return base[:255]


def _apply_excel_style(ws, table_name: str) -> None:
    header_fill = PatternFill("solid", fgColor="0A58CA")
    header_font = Font(bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center", horizontal="center")

    ws.freeze_panes = "A2"

    max_row = ws.max_row
    max_col = ws.max_column
    if max_row >= 2 and max_col >= 1:
        ref = f"A1:{get_column_letter(max_col)}{max_row}"
        ws.auto_filter.ref = ref
        tab = Table(displayName=_safe_table_name(table_name), ref=ref)
        style = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        tab.tableStyleInfo = style
        try:
            ws.add_table(tab)
        except Exception:
            pass

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
    driver_mode: bool,
    holidays_text: str,
    expected: int,
    kpis_general: dict,
    summary_all: pd.DataFrame,
    extras_only: pd.DataFrame,
    ranking_hours: pd.DataFrame,
    ranking_extras: pd.DataFrame,
    inconsistencies: pd.DataFrame,
    daily: pd.DataFrame,
    raw: pd.DataFrame,
) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    def add_df(sheet_name: str, df: pd.DataFrame, text_cols: set[str] | None = None):
        text_cols = text_cols or set()
        ws = wb.create_sheet(sheet_name)

        if df is None or df.empty:
            df = pd.DataFrame(columns=["Sin_datos"])

        for j, col in enumerate(df.columns, start=1):
            ws.cell(row=1, column=j, value=str(col))

        for i, row in enumerate(df.itertuples(index=False), start=2):
            for j, val in enumerate(row, start=1):
                col_name = df.columns[j - 1]
                cell = ws.cell(row=i, column=j, value=val)
                if col_name in text_cols and val is not None:
                    cell.value = str(val)
                    cell.number_format = "@"

        _apply_excel_style(ws, table_name=sheet_name)

    df_kpis = pd.DataFrame([{
        "Horario_reducido": "SI" if reduced else "NO",
        "Control_choferes": "SI" if driver_mode else "NO",
        "Esperado_NO_Docente": minutes_to_hhmm(expected),
        "Corte_madrugada": f"00:00 a {WORKDAY_CUTOFF_HOUR - 1:02d}:59 cuenta como día laboral anterior",
        "Feriados_cargados": holidays_text.strip(),
        **kpis_general
    }])
    add_df("KPIs_General", df_kpis)

    summary_out = summary_all.copy()
    if "DNI" in summary_out.columns:
        summary_out["DNI"] = summary_out["DNI"].astype(str)
    add_df("Resumen_Empleados", summary_out, text_cols={"DNI", "EmployeeKey"})

    add_df("Solo_Extras", extras_only.copy(), text_cols={"DNI"})
    add_df("Ranking_Horas", ranking_hours.copy(), text_cols={"DNI"})
    add_df("Ranking_Extras", ranking_extras.copy(), text_cols={"DNI"})
    add_df("Inconsistencias", inconsistencies.copy(), text_cols={"DNI"})

    daily_out = daily.copy()
    if not daily_out.empty:
        daily_out["DNI"] = daily_out["DNI"].astype(str)
        daily_out["Fecha"] = pd.to_datetime(daily_out["Fecha"]).dt.strftime("%Y-%m-%d")
        daily_out["Primera"] = pd.to_datetime(daily_out["Primera"], errors="coerce").dt.strftime("%Y-%m-%d %H:%M")
        daily_out["Ultima"] = pd.to_datetime(daily_out["Ultima"], errors="coerce").dt.strftime("%Y-%m-%d %H:%M")
    add_df("Detalle_Diario", daily_out, text_cols={"DNI", "EmployeeKey"})

    raw_out = raw.copy()
    if not raw_out.empty:
        raw_out["DNI"] = raw_out["DNI"].astype(str)
        raw_out["Fecha_laboral"] = pd.to_datetime(raw_out["Fecha"]).dt.strftime("%Y-%m-%d")
        raw_out["Fecha_real"] = pd.to_datetime(raw_out["FechaReal"]).dt.strftime("%Y-%m-%d")
        raw_out["Hora"] = pd.to_datetime(raw_out["FechaHora"]).dt.strftime("%H:%M")
        raw_out = raw_out.sort_values(["Empleado", "DNI", "FechaHora"])[
            ["EmployeeKey", "Empleado", "DNI", "Tipo", "Fecha_laboral", "Fecha_real", "Hora", "Ajuste_madrugada"]
        ]
    add_df("Marcaciones", raw_out, text_cols={"DNI", "EmployeeKey"})

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


# ============================================================
# APP
# ============================================================
def main() -> None:
    st.set_page_config(page_title="CONTROL DE ASISTENCIA APP", page_icon="🫧", layout="wide")
    inject_css()
    hero_header()
    init_holidays_state()

    st.markdown('<div class="toolbar-wrap">', unsafe_allow_html=True)
    col1, col2, col3, col4 = st.columns([1, 1, 1, 1])
    with col1:
        reduced = st.toggle("HORARIO REDUCIDO", value=False)
    with col2:
        driver_mode = st.toggle("CONTROL DE CHOFERES", value=False)
    with col3:
        night_adjustment = st.toggle("AJUSTE MADRUGADA", value=True)
    with col4:
        st.markdown(
            f'<div class="pill">JORNADA {"06:00" if reduced else "07:00"}</div>',
            unsafe_allow_html=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)

    with st.expander("FERIADOS", expanded=False):
        holidays_text_value = st.text_area(
            "CARGAR FERIADOS (UNO POR LÍNEA O SEPARADOS POR COMA · DD/MM/AAAA O AAAA-MM-DD)",
            value=st.session_state.get("holidays_text_input", ""),
            height=100,
        )

        holiday_cols = st.columns(3)
        with holiday_cols[0]:
            if st.button("APLICAR FERIADOS", use_container_width=True):
                apply_text_holidays_from_value(holidays_text_value)
                st.rerun()
        with holiday_cols[1]:
            if st.button("LIMPIAR FERIADOS", use_container_width=True):
                clear_all_holidays()
                st.rerun()
        with holiday_cols[2]:
            st.markdown(
                f'<div class="pill">CARGADOS: {len(st.session_state["holidays_set"])}</div>',
                unsafe_allow_html=True,
            )

        render_holiday_calendar()

    holidays = set(st.session_state["holidays_set"])

    file = st.file_uploader("", type=["xlsx", "xlsm", "xls"], label_visibility="collapsed")
    if not file:
        return

    try:
        df0 = read_excel_auto(file)
        df0 = validate_format(df0)
        raw0 = parse_and_clean(df0, night_adjustment=night_adjustment)
    except Exception as exc:
        st.error(str(exc))
        return

    _ = init_profiles(raw0)
    expected = 360 if reduced else 420

    tabs = st.tabs(["GENERAL", "EMPLEADO", "PLANILLA / IMPRIMIR", "PERFILES"])

    with tabs[0]:
        raw = apply_profiles(raw0, st.session_state["profiles"])

        action_cols = st.columns([1.15, 1])
        with action_cols[0]:
            fix_all = False
            if not driver_mode:
                fix_all = st.button(
                    "CORREGIR DÍAS CON UNA SOLA MARCACIÓN",
                    use_container_width=True,
                )
        with action_cols[1]:
            st.markdown(
                '<div class="pill">CORRECCIÓN SEGURA: SOLO ACTÚA SI, DESPUÉS DE LIMPIAR DOBLES MARCAS, QUEDA UNA ÚNICA MARCACIÓN EFECTIVA</div>',
                unsafe_allow_html=True,
            )

        fixes_total = 0
        if fix_all:
            raw, fixes_total = correct_missing_punches_all(
                raw,
                expected,
                night_adjustment,
            )

        daily = calc_daily(raw, expected, holidays, driver_mode)
        summary = summarize(daily)

        compact_summary = compact_employee_summary(
            raw,
            daily,
            summary,
            holidays,
        )
        daily_table = compact_daily_table(daily)
        review_table = compact_review_table(daily)

        start_date, end_date = period_dates(raw)
        period_text = (
            f"{start_date.strftime('%d/%m/%Y')} AL {end_date.strftime('%d/%m/%Y')}"
            if start_date and end_date
            else "SIN PERÍODO"
        )

        total_extra = int(compact_summary["Extras_min"].sum()) if not compact_summary.empty else 0
        total_gross_extra = int(compact_summary["Extras_brutas_min"].sum()) if not compact_summary.empty and "Extras_brutas_min" in compact_summary.columns else total_extra
        total_missing_hours = int(compact_summary["Faltantes_min"].sum()) if not compact_summary.empty and "Faltantes_min" in compact_summary.columns else 0
        total_worked = int(daily["Minutos"].sum()) if not daily.empty else 0
        worked_days = int((daily["Minutos"] > 0).sum()) if not daily.empty else 0
        missing_days = int(compact_summary["Días_sin_marcación"].sum()) if not compact_summary.empty else 0
        incomplete_days = int((daily["Incompleto"] == "SI").sum()) if not daily.empty else 0
        duplicates_ignored = (
            int(daily["Duplicadas_ignoradas"].sum())
            if not daily.empty and "Duplicadas_ignoradas" in daily.columns
            else 0
        )

        complete_daily = daily[daily["Incompleto"] != "SI"].copy() if not daily.empty else pd.DataFrame()
        average_arrival = average_clock_time(complete_daily["Primera"]) if not complete_daily.empty else "—"
        average_departure = average_clock_time(
            complete_daily["Ultima"],
            overnight_as_next_day=True,
        ) if not complete_daily.empty else "—"

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        kpi_row_1 = st.columns(4)
        with kpi_row_1[0]:
            kpi_card("EXTRA A FAVOR", minutes_to_hhmm(total_extra), "EXTRA NETA DESPUÉS DE COMPENSAR FALTANTES", tone="positive")
        with kpi_row_1[1]:
            kpi_card("EXTRA GENERADA", minutes_to_hhmm(total_gross_extra), "SUMA DE TODOS LOS EXCEDENTES DIARIOS", tone="neutral")
        with kpi_row_1[2]:
            kpi_card("HORAS A COMPENSAR", minutes_to_hhmm(total_missing_hours), "SUMA DE SALIDAS TEMPRANAS / JORNADAS CORTAS", tone="negative" if total_missing_hours > 0 else "neutral")
        with kpi_row_1[3]:
            kpi_card("PERÍODO", period_text, f"{len(compact_summary)} EMPLEADOS PROCESADOS", tone="neutral")

        kpi_row_2 = st.columns(4)
        with kpi_row_2[0]:
            kpi_card("DÍAS SIN MARCACIÓN", str(missing_days), "LUNES A VIERNES; REVISAR LICENCIAS")
        with kpi_row_2[1]:
            kpi_card("MARCAS INCOMPLETAS", str(incomplete_days), "DÍAS CON CANTIDAD IMPAR DE MARCAS")
        with kpi_row_2[2]:
            kpi_card("ENTRADA PROMEDIO", average_arrival, "PROMEDIO DE PRIMERA MARCA")
        with kpi_row_2[3]:
            kpi_card("SALIDA PROMEDIO", average_departure, "PROMEDIO DE ÚLTIMA MARCA")

        explain_box(
            "CÓMO INTERPRETAR LOS DATOS",
            "EXTRA GENERADA = SUMA DE TODOS LOS EXCEDENTES DIARIOS. HORAS A COMPENSAR = SUMA DE LOS DÍAS DONDE TRABAJÓ MENOS DE LO ESPERADO. "
            "SALDO DEL PERÍODO = EXTRA GENERADA MENOS HORAS A COMPENSAR. EXTRA A FAVOR ES SOLO EL SALDO POSITIVO FINAL. "
            "LOS DÍAS SIN MARCACIÓN Y LOS DÍAS INCOMPLETOS QUEDAN PARA REVISIÓN Y NO SE DESCUENTAN AUTOMÁTICAMENTE."
        )
        st.markdown(
            f'<div class="pill">DOBLES MARCAS IGNORADAS: {duplicates_ignored} · DOS LECTURAS A 2 MINUTOS O MENOS CUENTAN COMO UN SOLO EVENTO</div>',
            unsafe_allow_html=True,
        )

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        section_title("RESUMEN POR EMPLEADO")
        summary_display = compact_summary.drop(
            columns=["EmployeeKey", "Extras_min"],
            errors="ignore",
        )
        copy_table_button(
            summary_display,
            "COPIAR RESUMEN POR EMPLEADO",
            key="copy_compact_summary",
        )
        st.dataframe(
            summary_display,
            use_container_width=True,
            height=390,
            hide_index=True,
            column_config={
                "Horas_extra": st.column_config.TextColumn(
                    "EXTRA A FAVOR",
                    help="SALDO POSITIVO FINAL DESPUÉS DE RESTAR LAS HORAS FALTANTES DEL PERÍODO.",
                ),
                "Extra_bruta": st.column_config.TextColumn(
                    "EXTRA GENERADA",
                    help="SUMA BRUTA DE TODOS LOS EXCEDENTES DIARIOS ANTES DE COMPENSAR FALTANTES.",
                ),
                "Horas_faltantes": st.column_config.TextColumn(
                    "HORAS A COMPENSAR",
                    help="SUMA DE LOS MINUTOS QUE FALTARON PARA COMPLETAR LA JORNADA EN DÍAS VÁLIDOS.",
                ),
                "Saldo_periodo": st.column_config.TextColumn(
                    "SALDO DEL PERÍODO",
                    help="EXTRA GENERADA MENOS HORAS A COMPENSAR. POSITIVO = A FAVOR. NEGATIVO = DEUDA.",
                ),
                "Deuda_pendiente": st.column_config.TextColumn(
                    "DEUDA PENDIENTE",
                    help="SI EL SALDO ES NEGATIVO, MUESTRA CUÁNTO DEBE COMPENSAR TODAVÍA.",
                ),
                "Días_sin_marcación": st.column_config.NumberColumn(
                    "DÍAS SIN MARCACIÓN",
                    help="DÍAS HÁBILES DEL PERÍODO SIN NINGUNA MARCA. REVISAR LICENCIAS O JUSTIFICACIONES.",
                ),
                "Entrada_promedio": st.column_config.TextColumn(
                    "ENTRADA PROMEDIO",
                    help="PROMEDIO DE LA PRIMERA MARCACIÓN DE LOS DÍAS COMPLETOS.",
                ),
                "Salida_promedio": st.column_config.TextColumn(
                    "SALIDA PROMEDIO",
                    help="PROMEDIO DE LA ÚLTIMA MARCACIÓN DE LOS DÍAS COMPLETOS.",
                ),
            },
        )

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        section_title("DETALLE DÍA A DÍA")
        copy_table_button(
            daily_table,
            "COPIAR DETALLE DÍA A DÍA",
            key="copy_compact_daily",
        )
        st.dataframe(
            daily_table,
            use_container_width=True,
            height=470,
            hide_index=True,
            column_config={
                "Marcaciones_detalle": st.column_config.TextColumn(
                    "TODAS LAS MARCACIONES",
                    help="TODAS LAS HORAS REGISTRADAS ESE DÍA, EN ORDEN.",
                    width="large",
                ),
                "Tramos_detalle": st.column_config.TextColumn(
                    "TRAMOS CALCULADOS",
                    help="PARES ENTRADA/SALIDA QUE REALMENTE SE USARON PARA SUMAR LAS HORAS.",
                    width="large",
                ),
                "Marcaciones_efectivas_detalle": st.column_config.TextColumn(
                    "MARCAS USADAS",
                    help="MARCACIONES DESPUÉS DE IGNORAR DOBLES LECTURAS MUY CERCANAS.",
                    width="large",
                ),
                "Extra_dia": st.column_config.TextColumn(
                    "EXTRA DEL DÍA",
                    help="LO QUE SUPERA LA JORNADA DEL DÍA. EN FERIADOS Y FINES DE SEMANA, TODO LO TRABAJADO ES EXTRA.",
                ),
                "Saldo": st.column_config.TextColumn(
                    "SALDO",
                    help="DIFERENCIA DEL DÍA. POSITIVO = EXTRA. NEGATIVO = TIEMPO FALTANTE.",
                ),
                "Duplicadas_ignoradas": st.column_config.NumberColumn(
                    "DOBLES IGNORADAS",
                    help="LECTURAS BIOMÉTRICAS MUY CERCANAS QUE SE MOSTRARON, PERO NO SE TOMARON COMO UNA ENTRADA/SALIDA NUEVA.",
                ),
            },
        )

        if not review_table.empty:
            st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
            with st.expander(
                f"MARCACIONES A REVISAR · {len(review_table)} CASOS",
                expanded=False,
            ):
                st.dataframe(
                    review_table,
                    use_container_width=True,
                    height=300,
                    hide_index=True,
                )

        compact_xlsx = export_compact_excel(
            compact_summary=compact_summary,
            daily_table=daily_table,
            review_table=review_table,
            raw=raw,
            period_text=period_text,
            total_extra=total_extra,
        )

        st.download_button(
            "EXPORTAR CONTROL DE ASISTENCIA (EXCEL)",
            data=compact_xlsx,
            file_name="control_asistencia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        if fixes_total:
            st.success(f"SE APLICARON {fixes_total} CORRECCIONES AUTOMÁTICAS.")

        st.session_state["__raw__"] = raw
        st.session_state["__daily__"] = daily
        st.session_state["__summary_compact__"] = compact_summary
        st.session_state["__expected__"] = expected
        st.session_state["__driver_mode__"] = driver_mode
        st.session_state["__holidays__"] = holidays
        st.session_state["__night_adjustment__"] = night_adjustment
        st.session_state["__period_text__"] = period_text

    with tabs[1]:
        raw = st.session_state.get("__raw__")
        daily = st.session_state.get("__daily__")
        compact_summary = st.session_state.get("__summary_compact__")
        expected = st.session_state.get("__expected__", expected)
        driver_mode = st.session_state.get("__driver_mode__", driver_mode)
        holidays = st.session_state.get("__holidays__", holidays)
        night_adjustment = st.session_state.get(
            "__night_adjustment__",
            night_adjustment,
        )

        if (
            raw is None
            or daily is None
            or compact_summary is None
            or compact_summary.empty
        ):
            st.info("CARGÁ UN EXCEL PARA VER ESTA SECCIÓN.")
            return

        options = compact_summary.copy()
        options["Display"] = (
            options["Empleado"].astype(str)
            + " · "
            + options["DNI"].astype(str)
            + " · "
            + options["Tipo"].astype(str)
        )
        selected = st.selectbox(
            "",
            options=options["Display"].tolist(),
            label_visibility="collapsed",
        )
        selected_row = options[options["Display"] == selected].iloc[0]
        employee_key = str(selected_row["EmployeeKey"])

        raw_emp = raw[
            raw["EmployeeKey"].astype(str) == employee_key
        ].copy().sort_values("FechaHora")
        daily_emp = daily[
            daily["EmployeeKey"].astype(str) == employee_key
        ].copy()

        if not driver_mode:
            if st.button(
                "CORREGIR DÍAS CON UNA SOLA MARCACIÓN DE ESTE EMPLEADO",
                use_container_width=True,
            ):
                corrected_emp, fixes = correct_missing_punches_for_employee(
                    raw_emp,
                    expected,
                    night_adjustment,
                )
                if fixes:
                    raw_corrected = raw[
                        raw["EmployeeKey"].astype(str) != employee_key
                    ].copy()
                    raw_corrected = pd.concat(
                        [raw_corrected, corrected_emp],
                        ignore_index=True,
                    ).sort_values(
                        ["Empleado", "DNI", "FechaHora"]
                    ).reset_index(drop=True)

                    daily_corrected = calc_daily(
                        raw_corrected,
                        expected,
                        holidays,
                        driver_mode,
                    )
                    raw_emp = corrected_emp
                    daily_emp = daily_corrected[
                        daily_corrected["EmployeeKey"].astype(str) == employee_key
                    ].copy()
                    st.success(f"SE APLICARON {fixes} CORRECCIONES.")
                else:
                    st.info("NO HAY DÍAS CON UNA ÚNICA MARCACIÓN.")

        metrics = compact_employee_metrics(
            raw,
            daily_emp,
            employee_key,
            holidays,
        )

        balance_hero(metrics["net_balance"])

        section_title("BALANCE DEL PERÍODO")
        employee_row_1 = st.columns(4)
        with employee_row_1[0]:
            balance_tone = "positive" if metrics["net_balance"] > 0 else "negative" if metrics["net_balance"] < 0 else "neutral"
            kpi_card(
                "SALDO DEL PERÍODO",
                delta_short(metrics["net_balance"]),
                "EXTRA GENERADA − HORAS A COMPENSAR",
                tone=balance_tone,
            )
        with employee_row_1[1]:
            kpi_card(
                "EXTRA A FAVOR",
                minutes_to_hhmm(metrics["extra_favor"]),
                "LO QUE QUEDA POSITIVO DESPUÉS DE COMPENSAR",
                tone="positive" if metrics["extra_favor"] > 0 else "neutral",
            )
        with employee_row_1[2]:
            kpi_card(
                "EXTRA GENERADA",
                minutes_to_hhmm(metrics["gross_extra"]),
                "SUMA BRUTA DE LOS EXCEDENTES DE CADA DÍA",
                tone="neutral",
            )
        with employee_row_1[3]:
            kpi_card(
                "HORAS A COMPENSAR",
                minutes_to_hhmm(metrics["missing_minutes"]),
                "TIEMPO QUE FALTÓ PARA COMPLETAR JORNADAS",
                tone="negative" if metrics["missing_minutes"] > 0 else "neutral",
            )

        section_title("JORNADA Y CUMPLIMIENTO")
        employee_row_2 = st.columns(4)
        with employee_row_2[0]:
            kpi_card(
                "TOTAL TRABAJADO",
                minutes_to_hhmm(metrics["total"]),
                "TODOS LOS TRAMOS CALCULADOS DEL PERÍODO",
                tone="neutral",
            )
        with employee_row_2[1]:
            kpi_card(
                "HORAS ESPERADAS",
                minutes_to_hhmm(metrics["expected"]),
                "SUMA DE LA JORNADA ESPERADA EN DÍAS VÁLIDOS",
                tone="neutral",
            )
        with employee_row_2[2]:
            kpi_card(
                "HORAS NORMALES",
                minutes_to_hhmm(metrics["normal"]),
                "TIEMPO IMPUTADO DENTRO DE LA JORNADA",
                tone="neutral",
            )
        with employee_row_2[3]:
            kpi_card(
                "CUMPLIMIENTO",
                f"{metrics['coverage_pct']:.1f}%",
                "TRABAJADO VÁLIDO / HORAS ESPERADAS",
                tone="positive" if metrics["coverage_pct"] >= 100 else "warning",
            )

        section_title("COMPORTAMIENTO DEL PERÍODO")
        employee_row_3 = st.columns(4)
        with employee_row_3[0]:
            kpi_card(
                "DÍAS CON EXTRA",
                str(metrics["days_with_extra"]),
                f"MAYOR EXTRA EN UN DÍA: {minutes_to_hhmm(metrics['largest_extra'])}",
                tone="positive" if metrics["days_with_extra"] > 0 else "neutral",
            )
        with employee_row_3[1]:
            kpi_card(
                "DÍAS CON FALTANTE",
                str(metrics["days_with_missing"]),
                f"MAYOR FALTANTE EN UN DÍA: {minutes_to_hhmm(metrics['largest_missing'])}",
                tone="negative" if metrics["days_with_missing"] > 0 else "neutral",
            )
        with employee_row_3[2]:
            kpi_card(
                "ENTRADA PROMEDIO",
                metrics["arrival"],
                "PROMEDIO DE PRIMERA MARCA EN DÍAS VÁLIDOS",
                tone="neutral",
            )
        with employee_row_3[3]:
            kpi_card(
                "SALIDA PROMEDIO",
                metrics["departure"],
                "PROMEDIO DE ÚLTIMA MARCA EN DÍAS VÁLIDOS",
                tone="neutral",
            )

        section_title("CONTROL Y REVISIÓN")
        employee_row_4 = st.columns(4)
        with employee_row_4[0]:
            kpi_card(
                "DÍAS TRABAJADOS",
                str(metrics["worked_days"]),
                "DÍAS CON HORAS CALCULADAS",
                tone="neutral",
            )
        with employee_row_4[1]:
            kpi_card(
                "DÍAS SIN MARCACIÓN",
                str(metrics["missing_days"]),
                "NO DESCUENTAN AUTOMÁTICAMENTE; REVISAR LICENCIA",
                tone="warning" if metrics["missing_days"] > 0 else "neutral",
            )
        with employee_row_4[2]:
            kpi_card(
                "MARCAS INCOMPLETAS",
                str(metrics["incomplete_days"]),
                "QUEDAN FUERA DEL BALANCE HASTA SER REVISADAS",
                tone="warning" if metrics["incomplete_days"] > 0 else "neutral",
            )
        with employee_row_4[3]:
            kpi_card(
                "DOBLES IGNORADAS",
                str(metrics["duplicates"]),
                "LECTURAS BIOMÉTRICAS REPETIDAS QUE NO ALTERAN EL CÁLCULO",
                tone="neutral",
            )

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        explain_box(
            "CÓMO FUNCIONA EL BALANCE",
            "CADA DÍA HÁBIL TIENE UNA JORNADA ESPERADA DE 7 HORAS, O 6 SI ESTÁ ACTIVO HORARIO REDUCIDO. "
            "SI TRABAJA MÁS, ESE EXCEDENTE SE SUMA COMO EXTRA GENERADA. SI TRABAJA MENOS, ESA DIFERENCIA SE SUMA COMO HORAS A COMPENSAR. "
            "AL FINAL DEL PERÍODO: EXTRA GENERADA − HORAS A COMPENSAR = SALDO. "
            "UN SALDO POSITIVO ES EXTRA A FAVOR; UN SALDO NEGATIVO ES TIEMPO QUE TODAVÍA DEBE COMPENSAR. "
            "LOS DÍAS SIN MARCACIÓN Y LOS DÍAS INCOMPLETOS NO MODIFICAN EL BALANCE HASTA QUE RRHH LOS REVISE."
        )

        employee_daily = build_employee_period_table(
            raw=raw,
            daily_emp=daily_emp,
            employee_key=employee_key,
            holidays=holidays,
        )
        section_title("DÍA A DÍA DEL EMPLEADO")
        copy_table_button(
            employee_daily,
            "COPIAR DÍA A DÍA DEL EMPLEADO",
            key="copy_employee_daily",
        )
        st.dataframe(
            employee_daily,
            use_container_width=True,
            height=520,
            hide_index=True,
            column_config={
                "Marcaciones_detalle": st.column_config.TextColumn(
                    "TODAS LAS MARCACIONES",
                    help="TODAS LAS LECTURAS DEL RELOJ, INCLUSO LAS DOBLES.",
                    width="large",
                ),
                "Marcaciones_efectivas_detalle": st.column_config.TextColumn(
                    "MARCAS USADAS",
                    help="LECTURAS QUE REALMENTE SE USARON PARA ARMAR LOS PARES ENTRADA/SALIDA.",
                    width="large",
                ),
                "Saldo": st.column_config.TextColumn(
                    "SALDO DEL DÍA",
                    help="DIFERENCIA DEL DÍA. POSITIVO = EXTRA; NEGATIVO = TIEMPO A COMPENSAR.",
                ),
                "Saldo_acumulado": st.column_config.TextColumn(
                    "SALDO ACUMULADO",
                    help="BALANCE CORRIDO DEL PERÍODO. UN DÍA CON EXTRA PUEDE COMPENSAR UN FALTANTE DE OTRO DÍA.",
                ),
                "Duplicadas_ignoradas": st.column_config.NumberColumn(
                    "DOBLES IGNORADAS",
                    help="MARCAS MUY CERCANAS QUE EL SISTEMA INTERPRETÓ COMO UNA SOLA LECTURA BIOMÉTRICA.",
                ),
            },
        )

        with st.expander("MARCACIONES ORIGINALES DEL RELOJ", expanded=False):
            raw_table = raw_employee_marks_table(raw_emp)
            st.dataframe(
                raw_table,
                use_container_width=True,
                height=350,
                hide_index=True,
            )

    with tabs[2]:
        raw = st.session_state.get("__raw__")
        daily = st.session_state.get("__daily__")
        compact_summary = st.session_state.get("__summary_compact__")
        holidays = st.session_state.get("__holidays__", holidays)
        period_text = st.session_state.get("__period_text__", "PERÍODO CARGADO")

        if (
            raw is None
            or daily is None
            or compact_summary is None
            or compact_summary.empty
        ):
            st.info("CARGÁ UN EXCEL PARA ARMAR LA PLANILLA.")
        else:
            explain_box(
                "PLANILLA LISTA PARA IMPRIMIR",
                "PODÉS ELEGIR UNO O VARIOS EMPLEADOS. EL ARCHIVO GENERA LA MISMA TABLA COMPLETA PARA CADA PERSONA, "
                "CON TODAS LAS MARCACIONES, MARCAS USADAS, TRAMOS, HORAS TRABAJADAS, NORMALES, EXTRA, FALTANTES Y SALDO. "
                "CADA EMPLEADO COMIENZA EN UNA PÁGINA NUEVA."
            )

            print_options = compact_summary.copy()
            print_options["Display"] = (
                print_options["Empleado"].astype(str)
                + " · "
                + print_options["DNI"].astype(str)
            )
            display_to_key = dict(
                zip(
                    print_options["Display"],
                    print_options["EmployeeKey"].astype(str),
                )
            )

            select_all = st.toggle(
                "INCLUIR TODOS LOS EMPLEADOS",
                value=False,
                key="print_all_employees",
            )

            if select_all:
                selected_displays = print_options["Display"].tolist()
                st.caption(f"SE INCLUIRÁN {len(selected_displays)} EMPLEADOS.")
            else:
                selected_displays = st.multiselect(
                    "EMPLEADOS A INCLUIR",
                    options=print_options["Display"].tolist(),
                    default=[],
                )

            selected_keys = [
                display_to_key[item]
                for item in selected_displays
                if item in display_to_key
            ]

            if selected_keys:
                preview_daily = daily[
                    daily["EmployeeKey"].astype(str).isin(selected_keys)
                ].copy()
                preview = compact_daily_table(preview_daily)

                section_title("VISTA PREVIA")
                st.dataframe(
                    preview,
                    use_container_width=True,
                    height=420,
                    hide_index=True,
                )

                printable_xlsx = export_printable_employee_workbook(
                    raw=raw,
                    daily=daily,
                    compact_summary=compact_summary,
                    employee_keys=selected_keys,
                    holidays=holidays,
                    period_text=period_text,
                )

                st.download_button(
                    "DESCARGAR PLANILLA LISTA PARA IMPRIMIR",
                    data=printable_xlsx,
                    file_name="planilla_empleados_imprimible.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            else:
                st.info("SELECCIONÁ AL MENOS UN EMPLEADO O ACTIVÁ «INCLUIR TODOS LOS EMPLEADOS».")

    with tabs[3]:
        profiles_view = st.session_state["profiles"].copy()
        profiles_view["DNI"] = profiles_view["DNI"].apply(display_dni)

        edited = st.data_editor(
            profiles_view,
            use_container_width=True,
            height=600,
            hide_index=True,
            disabled=["EmployeeKey", "DNI", "Empleado"],
            column_config={
                "EmployeeKey": st.column_config.TextColumn(
                    "ID INTERNO",
                    width="medium",
                ),
                "DNI": st.column_config.TextColumn("DNI", width="small"),
                "Empleado": st.column_config.TextColumn(
                    "EMPLEADO",
                    width="large",
                ),
                "Tipo": st.column_config.SelectboxColumn(
                    "TIPO",
                    options=["NO Docente", "Docente"],
                    required=True,
                ),
            },
        )

        real_profiles = st.session_state["profiles"].copy()
        real_profiles["Tipo"] = edited["Tipo"]
        st.session_state["profiles"] = real_profiles


if __name__ == "__main__":
    main()
