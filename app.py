
import os
import base64
import html
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import plotly.express as px
import plotly.graph_objects as go


# ============================================================
# CONFIGURACIÓN
# ============================================================
st.set_page_config(
    page_title="Indicadores de Productividad y Cobranza",
    page_icon="📊",
    layout="wide"
)

# Usa Excel, no CSV, porque ahora la base trae 2 hojas: Cartera y Cobranza.
RUTA_DEFAULT = "Base.xlsx"

# Nombre de la imagen de fondo. Debe estar en la misma carpeta que este script.
NOMBRE_IMAGEN_FONDO = "ChatGPT Image 19 may 2026, 11_58_09 a.m."

# Oculta la tarjeta superior de carga de archivo y el expander de control de datos.
# El tablero seguirá usando RUTA_DEFAULT como archivo base.
MOSTRAR_SECCION_ARCHIVO = False
MOSTRAR_CONTROL_DATOS = False

NIVELES_ESTRUCTURA = [
    "Unidad de Negocio",
    "Marca",
    "Region",
    "País",
    "Subdireccion",
    "Zona",
    "Sucursal",
    "Ruta",
]

# Filtros del menú lateral: SOLO estos 3.
FILTROS_LATERALES = [
    "Unidad de Negocio",
    "Marca",
    "País",
]

INDICADORES_BASE = [
    "Clientes Totales",
    "Clientes al corriente",
    "Faltas",
    "Nunca Abonados",
    "Cartera Total",
    "Saldo Cartera",
    "Saldo en atraso",
    "Saldo PP",
]

POSIBLES_COLUMNAS_COBRANZA_CARTERA = [
    "Cobranza Total",
    "Cobranza",
    "Recuperacion",
    "Recuperación",
    "Pago Total",
    "Pagos",
    "Pago Cobranza",
]

# Columnas reales detectadas en la hoja Cobranza de tu Base.xlsx:
# Año, Semana, Pais, Cuota Total Cobranza, Recuperación semana, % de Cumplimiento,
# Mejor semana, Peor semana.
COLUMNAS_COBRANZA_CUOTA = [
    "Cuota Total Cobranza",
    "Cuota total sin atraso",
    "Cuota Total Sin Atraso",
    "Cuota Total",
    "Cuota total",
    "Cuota",
    "Cuota total cobranza",
]

COLUMNAS_COBRANZA_PAGO = [
    "Recuperación semana",
    "Recuperacion semana",
    "Pago total sin atraso",
    "Pago Total Sin Atraso",
    "Pago Total",
    "Pago total",
    "Pago",
    "Cobranza Total",
    "Cobranza",
]

COLUMNAS_COBRANZA_CUMPLIMIENTO = [
    "% de Cumplimiento",
    "% Cumplimiento",
    "Cumplimiento",
    "% cumplimiento",
    "Porcentaje Cumplimiento",
]

COLUMNAS_COBRANZA_MEJOR = [
    "Mejor semana",
    "Mejor Semana",
]

COLUMNAS_COBRANZA_PEOR = [
    "Peor semana",
    "Peor Semana",
]

# Tipo de cambio a pesos mexicanos. La base original se conserva en moneda local;
# al elegir MXN en la barra superior, solo se convierten columnas monetarias.
TIPO_CAMBIO_MXN = {
    "COLOMBIA": 0.0047,
    "CO": 0.0047,
    "GUATEMALA": 2.44,
    "GT": 2.44,
    "PERU": 5.38,
    "PERÚ": 5.38,
    "PE": 5.38,
    "MX": 1,
    "MEXICO": 1,
    "MÉXICO": 1,
    "EL SALVADOR": 19.21,
    "SALVADOR": 19.21,
    "S": 19.21,
    "HONDURAS": 0.73,
    "HO": 0.73,
    "NICARAGUA": 0.5145,
    "NIC": 0.5145,
}

COLUMNAS_NO_MONETARIAS_EXACTAS = {
    "Clientes Totales",
    "Clientes al corriente",
    "Faltas",
    "Nunca Abonados",
    "Coordinadoras",
}

TERMINOS_NO_MONETARIOS = [
    "cliente",
    "clientes",
    "coord",
    "coordinadora",
    "coordinadoras",
    "faltas",
    "nunca abon",
    "cumplimiento",
    "semana",
    "año",
    "ano",
    "tipo",
    "pais",
    "país",
    "marca",
    "region",
    "ruta",
    "zona",
    "sucursal",
    "subdireccion",
    "unidad",
]

TERMINOS_MONETARIOS = [
    "cartera",
    "saldo",
    "cuota",
    "pago",
    "pagos",
    "cobranza",
    "recuperacion",
    "recuperación",
    "monto",
    "importe",
    "entregado",
    "pp",
    "mejor semana",
    "peor semana",
]


# ============================================================
# FONDO Y ESTILO
# ============================================================
def ruta_carpeta_script() -> Path:
    try:
        return Path(__file__).resolve().parent
    except Exception:
        return Path.cwd()


def buscar_imagen_fondo(nombre_imagen: str) -> Path | None:
    carpeta = ruta_carpeta_script()
    ruta_directa = carpeta / nombre_imagen

    if ruta_directa.exists():
        return ruta_directa

    for ext in [".png", ".jpg", ".jpeg", ".webp"]:
        ruta = carpeta / f"{nombre_imagen}{ext}"
        if ruta.exists():
            return ruta

    return None


@st.cache_data(show_spinner=False)
def imagen_a_base64(ruta_imagen: str) -> str:
    with open(ruta_imagen, "rb") as f:
        return base64.b64encode(f.read()).decode()


def imagen_logo_html(nombre_archivo: str, clase_css: str = "unidad-logo") -> str:
    """
    Convierte un logo local a HTML base64 para mostrarlo dentro de las tarjetas
    de la pantalla inicial. El archivo debe estar en la misma carpeta del script.
    """
    ruta = ruta_carpeta_script() / nombre_archivo

    if not ruta.exists():
        return '<div class="unidad-logo-placeholder">🏢</div>'

    logo_base64 = imagen_a_base64(str(ruta))

    extension = ruta.suffix.lower().replace(".", "")
    if extension in ["jpg", "jpeg"]:
        mime = "jpeg"
    elif extension == "png":
        mime = "png"
    elif extension == "webp":
        mime = "webp"
    else:
        mime = "png"

    return f'<img class="{clase_css}" src="data:image/{mime};base64,{logo_base64}" />'


def aplicar_fondo_pagina(nombre_imagen: str):
    ruta_imagen = buscar_imagen_fondo(nombre_imagen)

    if ruta_imagen is None:
        st.warning(
            "No encontré la imagen de fondo en la misma carpeta del script. "
            f"Revisa que exista el archivo: {nombre_imagen}.png, .jpg, .jpeg o .webp"
        )
        return

    fondo_base64 = imagen_a_base64(str(ruta_imagen))

    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image:
                linear-gradient(rgba(255,255,255,0.82), rgba(255,255,255,0.90)),
                url("data:image/png;base64,{fondo_base64}");
            background-size: cover;
            background-position: top center;
            background-repeat: no-repeat;
            background-attachment: fixed;
        }}

        header[data-testid="stHeader"] {{
            background: rgba(255,255,255,0);
        }}

        .main .block-container {{
            background: rgba(255,255,255,0.88);
            border-radius: 24px;
            padding: 1.4rem 2rem 2rem 2rem;
            margin-top: 1rem;
            margin-bottom: 2rem;
            box-shadow: 0 10px 35px rgba(15, 23, 42, 0.10);
            backdrop-filter: blur(2px);
        }}

        section[data-testid="stSidebar"] > div:first-child {{
            background: rgba(255,255,255,0.94);
            backdrop-filter: blur(3px);
        }}

        div[data-testid="stDataFrame"],
        div[data-testid="stTable"],
        div[data-testid="stPlotlyChart"] {{
            background: rgba(255,255,255,0.96);
            border-radius: 16px;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )


aplicar_fondo_pagina(NOMBRE_IMAGEN_FONDO)

st.markdown(
    """
    <style>
    .main .block-container {
        padding-top: 1.3rem;
        padding-bottom: 2rem;
        background: rgba(255,255,255,0.78) !important;
        border-radius: 28px !important;
        padding: 1.6rem 2rem 2.2rem 2rem !important;
        box-shadow: 0 14px 40px rgba(15, 23, 42, 0.12) !important;
        border: 1px solid rgba(226,232,240,0.85) !important;
    }

    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #eeeeee;
        padding: 14px 16px;
        border-radius: 16px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.04);
    }

    .titulo {
        font-size: 32px;
        font-weight: 800;
        margin-bottom: 0px;
    }

    .subtitulo {
        color: #666;
        font-size: 15px;
        margin-top: 0px;
        margin-bottom: 24px;
    }

    .kpi-card {
        background: rgba(255,255,255,0.98);
        border: 1px solid rgba(226,232,240,0.95);
        border-top: 5px solid #082567;
        padding: 18px 20px;
        border-radius: 16px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.10);
        min-height: 120px;
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }

    .kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 12px 30px rgba(15, 23, 42, 0.14);
    }

    .kpi-label {
        font-size: 13px;
        color: #082567;
        margin-bottom: 8px;
        font-weight: 700;
    }

    .kpi-value {
        font-size: 32px;
        font-weight: 700;
        color: #061a40;
        line-height: 1.1;
        margin-bottom: 12px;
    }

    .kpi-delta-positive {
        display: inline-block;
        font-size: 18px;
        font-weight: 800;
        color: #15803d;
        background: #dcfce7;
        padding: 5px 10px;
        border-radius: 999px;
    }

    .kpi-delta-negative {
        display: inline-block;
        font-size: 18px;
        font-weight: 800;
        color: #b91c1c;
        background: #fee2e2;
        padding: 5px 10px;
        border-radius: 999px;
    }

    .kpi-delta-neutral {
        display: inline-block;
        font-size: 18px;
        font-weight: 800;
        color: #475569;
        background: #f1f5f9;
        padding: 5px 10px;
        border-radius: 999px;
    }

    div[data-testid="stPlotlyChart"],
    div[data-testid="stDataFrame"],
    div[data-testid="stTable"] {
        background: rgba(255,255,255,0.96) !important;
        border: 1px solid rgba(226,232,240,0.95) !important;
        border-radius: 20px !important;
        padding: 14px !important;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.09) !important;
        overflow: hidden !important;
    }

    h2, h3 {
        color: #082567 !important;
        font-weight: 800 !important;
        letter-spacing: -0.2px;
    }

    h2::after, h3::after {
        content: "";
        display: block;
        width: 64px;
        height: 4px;
        background: #d9c322;
        border-radius: 999px;
        margin-top: 8px;
        margin-bottom: 6px;
    }

    div.stButton > button,
    div[data-testid="stDownloadButton"] > button {
        background: #082567 !important;
        color: white !important;
        border: 1px solid #082567 !important;
        border-radius: 12px !important;
        font-weight: 700 !important;
        box-shadow: 0 4px 12px rgba(8,37,103,0.18) !important;
    }

    div.stButton > button:hover,
    div[data-testid="stDownloadButton"] > button:hover {
        background: #d9c322 !important;
        color: #082567 !important;
        border: 1px solid #d9c322 !important;
    }

    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div {
        border-radius: 12px !important;
        background: rgba(255,255,255,0.98) !important;
    }

    div[data-testid="stAlert"] {
        background: rgba(239,246,255,0.98) !important;
        border-left: 6px solid #d9c322 !important;
        border-radius: 16px !important;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.08) !important;
        font-size: 19px !important;
        line-height: 1.55 !important;
    }

    .comentario-amplio {
        width: 80%;
        max-width: 80%;
        margin: 14px auto 22px auto;
        padding: 18px 26px;
        background: rgba(219,234,254,0.98);
        border-left: 8px solid #d9c322;
        border-radius: 18px;
        box-shadow: 0 8px 22px rgba(15,23,42,0.12);
        color: #082567;
        font-size: 19px;
        line-height: 1.55;
        font-weight: 600;
        box-sizing: border-box;
        text-align: left;
        position: static !important;
        clear: both;
        overflow-wrap: break-word;
        word-break: normal;
    }

    .comentario-amplio-titulo {
        display: inline-block;
        background: #082567;
        color: white;
        padding: 5px 12px;
        border-radius: 10px;
        margin-bottom: 10px;
        font-size: 15px;
        font-weight: 800;
    }

    .comentario-amplio-texto {
        display: block;
        white-space: normal;
    }

    .comentario-amplio strong {
        font-weight: 900;
    }

    @media (max-width: 900px) {
        .comentario-amplio {
            width: 96%;
            max-width: 96%;
            font-size: 17px;
            padding: 16px 18px;
        }
    }

    div[data-testid="stExpander"] {
        background: rgba(255,255,255,0.94) !important;
        border: 1px solid rgba(226,232,240,0.95) !important;
        border-radius: 16px !important;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.07) !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)



st.markdown(
    """
    <style>
    .top-filter-card {
        background: rgba(255,255,255,0.97);
        border: 1px solid rgba(226,232,240,0.95);
        border-radius: 18px;
        padding: 14px 16px 8px 16px;
        margin: 10px 0 18px 0;
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.08);
    }
    .top-filter-title {
        color: #082567;
        font-size: 18px;
        font-weight: 800;
        margin-bottom: 8px;
    }

    .landing-wrap {
        min-height: 24vh;
        display: flex;
        flex-direction: column;
        justify-content: flex-start;
        align-items: center;
        padding: 4vh 10px 12px 10px;
    }

    .landing-title {
        color: #082567;
        font-size: 42px;
        font-weight: 900;
        text-align: center;
        margin-bottom: 8px;
        letter-spacing: -0.5px;
    }

    .landing-subtitle {
        color: #334155;
        font-size: 18px;
        font-weight: 600;
        text-align: center;
        margin-bottom: 26px;
    }

    .unidad-card {
        background: rgba(255,255,255,0.96);
        border: 1px solid rgba(226,232,240,0.95);
        border-radius: 26px;
        padding: 28px 18px 20px 18px;
        min-height: 190px;
        box-shadow: 0 14px 34px rgba(15, 23, 42, 0.12);
        text-align: center;
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }

    .unidad-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 18px 42px rgba(15, 23, 42, 0.16);
    }

    .unidad-logo {
        width: 240px;
        height: 100px;
        object-fit: contain;
        display: block;
        margin: 0 auto 12px auto;
    }

    .unidad-logo-placeholder {
        width: 240px;
        height: 100px;
        display: flex;
        align-items: center;
        justify-content: center;
        margin: 0 auto 12px auto;
        font-size: 52px;
    }

    .unidad-card {
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }

    .unidad-name {
        color: #082567;
        font-size: 24px;
        font-weight: 900;
        margin-bottom: 6px;
    }

    .unidad-help {
        color: #64748b;
        font-size: 14px;
        font-weight: 600;
        margin-bottom: 14px;
    }

    .unidad-seleccionada-pill {
        display: inline-block;
        background: #082567;
        color: white;
        font-size: 14px;
        font-weight: 800;
        padding: 8px 14px;
        border-radius: 999px;
        margin-bottom: 6px;
    }

    .top-filter-card div[data-testid="stHorizontalBlock"] {
        align-items: end;
    }

    /* Barra superior 80% más grande */
    .top-filter-card label,
    .top-filter-card .stRadio label,
    .top-filter-card .stSelectbox label,
    .top-filter-card .stTextInput label,
    .top-filter-card .stFileUploader label {
        font-size: 18px !important;
        font-weight: 900 !important;
        color: #082567 !important;
    }

    .top-filter-card div[data-testid="stMarkdownContainer"] p,
    .top-filter-card div[data-baseweb="select"] span,
    .top-filter-card div[data-baseweb="radio"] label,
    .top-filter-card input {
        font-size: 18px !important;
        font-weight: 800 !important;
    }

    .top-filter-card div.stButton > button {
        min-height: 54px !important;
        font-size: 17px !important;
        font-weight: 900 !important;
        line-height: 1.15 !important;
        padding: 10px 12px !important;
    }

    .top-filter-title {
        font-size: 32px !important;
        line-height: 1.15 !important;
    }

    .unidad-seleccionada-pill {
        font-size: 21px !important;
        padding: 11px 18px !important;
    }

    /* Ventana emergente del resumen: ocupa cerca del 80% de pantalla */
    div[data-testid="stDialog"] div[role="dialog"] {
        width: 80vw !important;
        max-width: 80vw !important;
        height: 80vh !important;
        max-height: 80vh !important;
        overflow-y: auto !important;
        border-radius: 26px !important;
        padding: 18px 22px !important;
    }

    div[data-testid="stDialog"] h2 {
        color: #082567 !important;
        font-size: 30px !important;
        font-weight: 900 !important;
    }

    .modal-resumen-card {
        background: rgba(239,246,255,0.98);
        border-left: 8px solid #d9c322;
        border-radius: 20px;
        padding: 18px 22px;
        color: #082567;
        font-size: 22px;
        line-height: 1.45;
        font-weight: 700;
        box-shadow: 0 8px 22px rgba(15,23,42,0.12);
        margin: 8px 0 18px 0;
    }

    .modal-resumen-meta {
        background: #082567;
        color: white;
        border-radius: 999px;
        display: inline-block;
        padding: 8px 16px;
        font-size: 18px;
        font-weight: 900;
        margin: 0 8px 10px 0;
    }


    .top-bottom-opciones-card {
        background: rgba(255,255,255,0.97);
        border: 1px solid rgba(226,232,240,0.95);
        border-radius: 18px;
        padding: 12px 16px;
        margin: 10px 0 10px 0;
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.08);
    }

    .top-bottom-opciones-title {
        color: #082567;
        font-size: 20px;
        font-weight: 900;
    }
    </style>
    """,
    unsafe_allow_html=True
)



# ============================================================
# FIX VISUAL SEGURO PARA NAVEGADORES / TRADUCCIÓN / MODO OSCURO
# No cambia el peso/tamaño original de las letras; solo fuerza visibilidad.
# ============================================================
st.markdown(
    """
    <div class="notranslate" translate="no"></div>
    <style>
    html, body, .stApp, [data-testid="stAppViewContainer"] {
        opacity: 1 !important;
        filter: none !important;
    }

    label,
    [data-testid="stWidgetLabel"],
    [data-testid="stWidgetLabel"] p,
    [data-testid="stWidgetLabel"] span,
    .stSelectbox label,
    .stRadio label,
    .stNumberInput label,
    .stTextInput label,
    .stMultiSelect label {
        color: #082567 !important;
        -webkit-text-fill-color: #082567 !important;
        opacity: 1 !important;
        visibility: visible !important;
    }

    div[data-baseweb="select"],
    div[data-baseweb="select"] *,
    div[data-baseweb="popover"],
    div[data-baseweb="popover"] *,
    div[role="listbox"],
    div[role="listbox"] * {
        color: #061a40 !important;
        -webkit-text-fill-color: #061a40 !important;
        opacity: 1 !important;
        visibility: visible !important;
    }

    div[data-baseweb="select"] > div {
        background-color: #ffffff !important;
    }

    div[data-testid="stRadio"],
    div[data-testid="stRadio"] *,
    div[role="radiogroup"],
    div[role="radiogroup"] * {
        color: #061a40 !important;
        -webkit-text-fill-color: #061a40 !important;
        opacity: 1 !important;
        visibility: visible !important;
    }

    div.stButton > button,
    div.stButton > button *,
    div[data-testid="stDownloadButton"] > button,
    div[data-testid="stDownloadButton"] > button * {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        opacity: 1 !important;
        visibility: visible !important;
    }

    .kpi-card,
    .kpi-card *,
    .kpi-label,
    .kpi-value {
        opacity: 1 !important;
        visibility: visible !important;
    }

    input,
    textarea,
    select {
        color: #061a40 !important;
        -webkit-text-fill-color: #061a40 !important;
        background-color: #ffffff !important;
        opacity: 1 !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ============================================================
# FUNCIONES GENERALES
# ============================================================
def normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    reemplazos = {
        "Zona ": "Zona",
        "Clientes Nunca Abonados": "Nunca Abonados",
        "clientes nunca abonados": "Nunca Abonados",
        "Pais": "País",
        "pais": "País",
        "País ": "País",
        "Region ": "Region",
        "Subdirección": "Subdireccion",
        "subdireccion": "Subdireccion",
        "Semana": "Semana del año",
        "Semana ": "Semana del año",
        "Año ": "Año",
        "Recuperacion semana": "Recuperación semana",
        "% Cumplimiento": "% de Cumplimiento",
    }

    df = df.rename(columns={c: reemplazos.get(c, c) for c in df.columns})
    return df


def detectar_columna(posibles: list[str], columnas) -> str | None:
    columnas_set = set(columnas)

    for c in posibles:
        if c in columnas_set:
            return c

    # búsqueda flexible por minúsculas
    mapa = {str(c).strip().lower(): c for c in columnas}
    for c in posibles:
        if c.lower() in mapa:
            return mapa[c.lower()]

    return None


def normalizar_texto_tc(valor) -> str:
    if pd.isna(valor):
        return ""

    texto = str(valor).strip().upper()
    reemplazos = {
        "Á": "A",
        "É": "E",
        "Í": "I",
        "Ó": "O",
        "Ú": "U",
        "Ü": "U",
        "Ñ": "N",
    }
    for origen, destino in reemplazos.items():
        texto = texto.replace(origen, destino)

    return texto


def obtener_columna_pais(df_base: pd.DataFrame) -> str | None:
    for col in ["País", "Pais", "PAIS", "ID País", "ID Pais", "ID PAIS"]:
        if col in df_base.columns:
            return col
    return None


def es_columna_monetaria(nombre_columna: str) -> bool:
    nombre = str(nombre_columna).strip()
    nombre_norm = normalizar_texto_tc(nombre).lower()

    if nombre in COLUMNAS_NO_MONETARIAS_EXACTAS:
        return False

    # Porcentajes y cumplimientos nunca se convierten.
    if "cumplimiento" in nombre_norm or nombre_norm.startswith("%"):
        return False

    # Primero detecta términos monetarios para no descartar columnas como
    # "Recuperación semana", "Mejor semana" o "Peor semana" por traer la palabra semana.
    for termino in TERMINOS_MONETARIOS:
        if normalizar_texto_tc(termino).lower() in nombre_norm:
            return True

    for termino in TERMINOS_NO_MONETARIOS:
        if termino in nombre_norm:
            return False

    return False


def aplicar_tipo_cambio_mxn(df_base: pd.DataFrame, modo_moneda: str) -> pd.DataFrame:
    """
    Convierte columnas monetarias a pesos mexicanos cuando el usuario elige MXN.
    Las variables no monetarias como clientes, faltas y coordinadoras no se modifican.
    """
    df_tmp = df_base.copy()

    if modo_moneda != "Pesos mexicanos":
        return df_tmp

    col_pais = obtener_columna_pais(df_tmp)
    if col_pais is None:
        return df_tmp

    tc_map = {normalizar_texto_tc(k): v for k, v in TIPO_CAMBIO_MXN.items()}
    factor_tc = (
        df_tmp[col_pais]
        .apply(normalizar_texto_tc)
        .map(tc_map)
        .fillna(1)
        .astype(float)
    )

    columnas_convertir = [
        c for c in df_tmp.columns
        if c != col_pais
        and pd.api.types.is_numeric_dtype(df_tmp[c])
        and es_columna_monetaria(c)
    ]

    for col in columnas_convertir:
        df_tmp[col] = pd.to_numeric(df_tmp[col], errors="coerce").fillna(0) * factor_tc

    return df_tmp


def etiqueta_moneda(modo_moneda: str) -> str:
    return "MXN" if modo_moneda == "Pesos mexicanos" else "Moneda local"


def limpiar_datos(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = normalizar_columnas(df)

    for c in df.select_dtypes(include=["object"]).columns:
        df[c] = df[c].astype(str).str.strip()
        df[c] = df[c].replace({"nan": np.nan, "None": np.nan, "": np.nan})

    if "Semana del año" in df.columns:
        df["Semana del año"] = pd.to_numeric(df["Semana del año"], errors="coerce").astype("Int64")

    if "Año" in df.columns:
        df["Año"] = pd.to_numeric(df["Año"], errors="coerce").astype("Int64")

    columnas_posibles_numericas = list(set(
        INDICADORES_BASE
        + POSIBLES_COLUMNAS_COBRANZA_CARTERA
        + COLUMNAS_COBRANZA_CUOTA
        + COLUMNAS_COBRANZA_PAGO
        + COLUMNAS_COBRANZA_CUMPLIMIENTO
        + COLUMNAS_COBRANZA_MEJOR
        + COLUMNAS_COBRANZA_PEOR
    ))

    for c in columnas_posibles_numericas:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    if "Tipo Coordinadora" in df.columns:
        df["Tipo Coordinadora"] = df["Tipo Coordinadora"].fillna("NA")
        df["Tipo Coordinadora"] = df["Tipo Coordinadora"].replace({
            "NA": "Secundaria",
            "N/A": "Secundaria",
            "SECUNDARIA": "Secundaria",
            "Secundaria": "Secundaria",
            "Coordinadora Activa": "En Desarrollo",
            "Activa": "En Desarrollo",
            "Coordinadora en Desarrollo": "En Desarrollo",
            "En Desarrollo": "En Desarrollo",
            "Coordinadora Improductiva": "Improductiva",
            "Improductiva": "Improductiva",
            "Coordinadora Productiva": "Productiva",
            "Productiva": "Productiva",
        })

    filas_antes = len(df)
    df = df.drop_duplicates().copy()

    df.attrs["duplicados_exactos_eliminados"] = filas_antes - len(df)
    df.attrs["filas_antes_limpieza"] = filas_antes
    df.attrs["filas_despues_limpieza"] = len(df)

    return df


@st.cache_data(show_spinner=False)
def cargar_archivo(ruta_local: str | None, archivo_subido):
    df_cartera = None
    df_cobranza = None

    if archivo_subido is not None:
        nombre = archivo_subido.name.lower()

        if nombre.endswith(".csv"):
            df_cartera = pd.read_csv(archivo_subido, encoding="utf-8-sig")

        elif nombre.endswith((".xlsx", ".xlsm", ".xlsb", ".xls")):
            excel = pd.ExcelFile(archivo_subido)

            hoja_cartera = "Cartera" if "Cartera" in excel.sheet_names else excel.sheet_names[0]
            df_cartera = pd.read_excel(archivo_subido, sheet_name=hoja_cartera)

            if "Cobranza" in excel.sheet_names:
                df_cobranza = pd.read_excel(archivo_subido, sheet_name="Cobranza")

        else:
            raise ValueError("Formato no soportado. Usa CSV o Excel.")

    else:
        if not ruta_local or not os.path.exists(ruta_local):
            raise FileNotFoundError(
                f"No encontré el archivo en la ruta:\n{ruta_local}\n\n"
                "Puedes subir el archivo desde el panel lateral o corregir RUTA_DEFAULT."
            )

        if ruta_local.lower().endswith(".csv"):
            df_cartera = pd.read_csv(ruta_local, encoding="utf-8-sig")

        elif ruta_local.lower().endswith((".xlsx", ".xlsm", ".xlsb", ".xls")):
            excel = pd.ExcelFile(ruta_local)

            hoja_cartera = "Cartera" if "Cartera" in excel.sheet_names else excel.sheet_names[0]
            df_cartera = pd.read_excel(ruta_local, sheet_name=hoja_cartera)

            if "Cobranza" in excel.sheet_names:
                df_cobranza = pd.read_excel(ruta_local, sheet_name="Cobranza")

        else:
            raise ValueError("Formato no soportado. Usa CSV o Excel.")

    df_cartera = limpiar_datos(df_cartera)

    if df_cobranza is not None:
        df_cobranza = limpiar_datos(df_cobranza)

    return df_cartera, df_cobranza


def formato_numero(valor):
    if pd.isna(valor):
        return ""

    try:
        valor = float(valor)
    except Exception:
        return valor

    return f"{valor:,.0f}"


def formato_variacion(valor):
    if pd.isna(valor):
        return ""

    try:
        valor = float(valor)
    except Exception:
        return valor

    signo = "+" if valor > 0 else ""
    return f"{signo}{valor:,.0f}"


def formato_millones(valor):
    if pd.isna(valor):
        return ""

    try:
        valor = float(valor)
    except Exception:
        return valor

    return f"{valor / 1_000_000:,.2f} mill."


def formato_pct(valor, decimales=2, signo=False):
    if pd.isna(valor) or np.isinf(valor):
        return ""

    try:
        valor = float(valor)
    except Exception:
        return valor

    if abs(valor) > 1.5:
        valor = valor / 100

    prefijo = "+" if signo and valor > 0 else ""
    return f"{prefijo}{valor:.{decimales}%}"



def formato_eje_compacto(valor):
    """
    Formato corto para el eje Y.
    Evita que números muy grandes ocupen demasiado espacio y compriman la gráfica.
    """
    if pd.isna(valor):
        return ""

    try:
        valor = float(valor)
    except Exception:
        return str(valor)

    abs_valor = abs(valor)

    if abs_valor >= 1_000_000_000:
        return f"{valor / 1_000_000_000:,.2f}B"
    if abs_valor >= 1_000_000:
        return f"{valor / 1_000_000:,.1f}M"
    if abs_valor >= 1_000:
        return f"{valor / 1_000:,.0f}K"

    return f"{valor:,.0f}"


def crear_grafica_evolucion_fija(
    evol: pd.DataFrame,
    indicador_grafica: str,
    modo_moneda: str,
    altura: int = 430
):
    """
    Gráfica de evolución semanal con escala fija, sin zoom, sin desplazamiento
    y con etiquetas separadas para que se lean bien los datos y las variaciones.
    """
    df_plot = evol.copy()
    df_plot = df_plot.dropna(subset=["Semana del año", indicador_grafica]).copy()

    if df_plot.empty:
        return go.Figure(), {}

    df_plot["Semana del año"] = pd.to_numeric(df_plot["Semana del año"], errors="coerce")
    df_plot[indicador_grafica] = pd.to_numeric(df_plot[indicador_grafica], errors="coerce").fillna(0)
    df_plot["Variación vs anterior"] = pd.to_numeric(
        df_plot["Variación vs anterior"],
        errors="coerce"
    )

    df_plot = df_plot.sort_values("Semana del año").reset_index(drop=True)
    df_plot["Semana texto"] = df_plot["Semana del año"].apply(lambda x: f"S{int(x)}")

    df_plot["Texto valor"] = df_plot[indicador_grafica].apply(lambda x: f"{x:,.0f}")
    df_plot["Texto variacion"] = df_plot["Variación vs anterior"].apply(
        lambda x: formato_variacion(x) if pd.notna(x) else ""
    )

    valores_y = df_plot[indicador_grafica].astype(float)
    y_min = float(valores_y.min())
    y_max = float(valores_y.max())
    rango = y_max - y_min

    if rango == 0:
        base = max(abs(y_max), 1)
        y_min_fijo = y_min - base * 0.10
        y_max_fijo = y_max + base * 0.20
    else:
        y_min_fijo = y_min - rango * 0.22
        y_max_fijo = y_max + rango * 0.45

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=df_plot["Semana texto"],
            y=df_plot[indicador_grafica],
            mode="lines+markers",
            line=dict(color="#082567", width=3),
            marker=dict(
                size=10,
                color="#d9c322",
                line=dict(color="#082567", width=2)
            ),
            customdata=np.stack(
                [
                    df_plot["Texto variacion"],
                    df_plot["Semana del año"],
                    df_plot["Texto valor"],
                ],
                axis=-1
            ),
            hovertemplate=(
                "<b>Semana:</b> %{customdata[1]:.0f}<br>"
                f"<b>{indicador_grafica}:</b> %{{customdata[2]}}<br>"
                "<b>Variación vs anterior:</b> %{customdata[0]}"
                "<extra></extra>"
            ),
            cliponaxis=False
        )
    )

    # Valor y variación se muestran como anotaciones separadas.
    # Ajusta yshift_variacion / yshift_valor si quieres más o menos distancia.
    yshift_valor = 16
    yshift_variacion = 46

    for _, fila in df_plot.iterrows():
        fig.add_annotation(
            x=fila["Semana texto"],
            y=fila[indicador_grafica],
            text=fila["Texto valor"],
            showarrow=False,
            yshift=yshift_valor,
            font=dict(color="#082567", size=11),
            bgcolor="rgba(255,255,255,0)",
            borderwidth=0,
            borderpad=0
        )

        if pd.notna(fila["Variación vs anterior"]):
            fig.add_annotation(
                x=fila["Semana texto"],
                y=fila[indicador_grafica],
                text=f"<b>{formato_variacion(fila['Variación vs anterior'])}</b>",
                showarrow=False,
                yshift=yshift_variacion,
                font=dict(color="#082567", size=12),
                bgcolor="rgba(255,255,255,0.86)",
                bordercolor="rgba(8,37,103,0.16)",
                borderwidth=1,
                borderpad=3
            )

    ticks_y = np.linspace(y_min_fijo, y_max_fijo, 5)

    titulo_eje_y = (
        f"Monto ({etiqueta_moneda(modo_moneda)})"
        if es_columna_monetaria(indicador_grafica)
        else "Valor"
    )

    fig.update_layout(
        height=altura,
        showlegend=False,
        hovermode="x unified",
        dragmode=False,
        margin=dict(t=90, b=54, l=72, r=42),
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,1)",
        font=dict(color="#082567", size=12),
        xaxis_title="Semana",
        yaxis_title=titulo_eje_y,
        uirevision="grafica_evolucion_fija",
        transition_duration=0,
        clickmode="none"
    )

    fig.update_xaxes(
        type="category",
        categoryorder="array",
        categoryarray=df_plot["Semana texto"].tolist(),
        fixedrange=True,
        showgrid=False,
        zeroline=False,
        tickfont=dict(size=11, color="#64748b"),
        title_font=dict(size=12, color="#64748b")
    )

    fig.update_yaxes(
        fixedrange=True,
        range=[y_min_fijo, y_max_fijo],
        tickmode="array",
        tickvals=ticks_y,
        ticktext=[formato_eje_compacto(v) for v in ticks_y],
        gridcolor="rgba(148,163,184,0.25)",
        zeroline=False,
        tickfont=dict(size=11, color="#64748b"),
        title_font=dict(size=12, color="#64748b")
    )

    config = {
        "displayModeBar": False,
        "scrollZoom": False,
        "doubleClick": False,
        "responsive": False,
        "staticPlot": False,
    }

    return fig, config

def tarjeta_kpi(label, valor, variacion=None):
    valor_fmt = formato_numero(valor)

    if variacion is None or pd.isna(variacion):
        delta_html = ""
    else:
        variacion = float(variacion)
        variacion_fmt = formato_variacion(variacion)

        if variacion > 0:
            clase = "kpi-delta-positive"
            flecha = "↑"
        elif variacion < 0:
            clase = "kpi-delta-negative"
            flecha = "↓"
        else:
            clase = "kpi-delta-neutral"
            flecha = "→"

        delta_html = f'<div class="{clase}">{flecha} {variacion_fmt}</div>'

    html = f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{valor_fmt}</div>
        {delta_html}
    </div>
    """

    st.markdown(html, unsafe_allow_html=True)


def mostrar_boton_comentario(clave: str, texto: str):
    """
    Muestra el comentario de forma automática y dentro del flujo normal de la página.
    Ya no usa botón ni session_state, por lo que el comentario se recalcula en cada cambio
    de filtro, país, marca, moneda o variable seleccionada.
    """
    if texto is None or str(texto).strip() == "":
        return

    comentario_seguro = html.escape(str(texto)).replace("\n", "<br>")
    st.markdown(
        '<div class="comentario-amplio">'
        '<div class="comentario-amplio-titulo">Comentario</div>'
        '<div class="comentario-amplio-texto">' + comentario_seguro + '</div>'
        '</div>',
        unsafe_allow_html=True
    )


def _fmt_comentario(valor):
    try:
        return f"{float(valor):,.0f}"
    except Exception:
        return str(valor)


def filtrar_por_diccionario(df_base: pd.DataFrame, filtros: dict, excluir_col: str | None = None):
    df_tmp = df_base.copy()

    for col, seleccion in filtros.items():
        if col == excluir_col:
            continue

        if col in df_tmp.columns and seleccion:
            df_tmp = df_tmp[df_tmp[col].astype(str).isin(seleccion)]

    return df_tmp


def aplicar_filtros_base(df_base: pd.DataFrame, semanas_sel: list[int], filtros: dict):
    df_tmp = df_base.copy()

    if "Semana del año" in df_tmp.columns and semanas_sel:
        df_tmp = df_tmp[df_tmp["Semana del año"].isin(semanas_sel)]

    for col, seleccion in filtros.items():
        if col in df_tmp.columns and seleccion:
            df_tmp = df_tmp[df_tmp[col].astype(str).isin(seleccion)]

    return df_tmp


def aplicar_filtros_cobranza_todas_las_semanas(df_base: pd.DataFrame, filtros: dict):
    """
    Filtra Cobranza por estructura, pero NO por el semana de análisis.
    Esto permite que las gráficas de Cobranza siempre muestren todo el histórico
    disponible con los filtros de Unidad de Negocio / Marca / País aplicados.
    """
    df_tmp = df_base.copy()

    for col, seleccion in filtros.items():
        if col in df_tmp.columns and seleccion:
            df_tmp = df_tmp[df_tmp[col].astype(str).isin(seleccion)]

    return df_tmp



def aplicar_filtros_cobranza_desde_cartera(
    df_cobranza_base: pd.DataFrame,
    df_cartera_base: pd.DataFrame,
    filtros: dict
) -> pd.DataFrame:
    """
    Filtra Cobranza usando la selección hecha en Cartera.
    Esto corrige el caso donde Cobranza solo trae País, pero el usuario eligió
    Unidad de Negocio o Marca en la barra superior. En ese caso primero se obtiene
    desde Cartera la lista de países/marcas que corresponden a la unidad elegida
    y luego se aplica esa lista a Cobranza.
    """
    df_tmp = df_cobranza_base.copy()

    if df_tmp.empty:
        return df_tmp

    df_ref = filtrar_por_diccionario(df_cartera_base, filtros)

    if df_ref.empty:
        return df_tmp.iloc[0:0].copy()

    columnas_puente = [
        "Unidad de Negocio",
        "Marca",
        "País",
        "Subdireccion",
        "Zona",
        "Sucursal",
        "Ruta",
    ]

    for col in columnas_puente:
        if col in df_tmp.columns and col in df_ref.columns:
            valores_validos = (
                df_ref[col]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
                .tolist()
            )

            if valores_validos:
                df_tmp = df_tmp[df_tmp[col].astype(str).str.strip().isin(valores_validos)]

    return df_tmp


# ============================================================
# FUNCIONES CARTERA
# ============================================================
def detectar_columna_cobranza_cartera(df: pd.DataFrame):
    return detectar_columna(POSIBLES_COLUMNAS_COBRANZA_CARTERA, df.columns)


def consolidar_grano_correcto(df_base: pd.DataFrame, indicadores: list[str]) -> pd.DataFrame:
    df_tmp = df_base.copy()

    columnas_grano = [
        c for c in [
            "Semana del año",
            "Unidad de Negocio",
            "Marca",
            "Region",
            "País",
            "Subdireccion",
            "Zona",
            "Sucursal",
            "Ruta",
            "coordinadora_id",
            "Tipo Coordinadora",
        ]
        if c in df_tmp.columns
    ]

    if "coordinadora_id" not in df_tmp.columns:
        columnas_grano = [
            c for c in [
                "Semana del año",
                "Unidad de Negocio",
                "Marca",
                "Region",
                "País",
                "Subdireccion",
                "Zona",
                "Sucursal",
                "Ruta",
                "Tipo Coordinadora",
            ]
            if c in df_tmp.columns
        ]

    columna_cobranza = detectar_columna_cobranza_cartera(df_tmp)

    agg_dict = {}

    for col in indicadores:
        if col in df_tmp.columns:
            agg_dict[col] = "max"

    if columna_cobranza and columna_cobranza in df_tmp.columns:
        agg_dict[columna_cobranza] = "sum"

    if not columnas_grano or not agg_dict:
        return df_tmp

    return (
        df_tmp
        .groupby(columnas_grano, dropna=False, as_index=False)
        .agg(agg_dict)
    )


def calcular_resumen_actual_vs_anterior(df_filtrado: pd.DataFrame, indicadores: list[str], semana_actual: int):
    semanas_previas = sorted([
        int(s) for s in df_filtrado["Semana del año"].dropna().unique()
        if int(s) < int(semana_actual)
    ])

    semana_anterior = semanas_previas[-1] if semanas_previas else None

    actual = (
        df_filtrado[df_filtrado["Semana del año"] == semana_actual][indicadores]
        .sum(numeric_only=True)
    )

    if semana_anterior is not None:
        anterior = (
            df_filtrado[df_filtrado["Semana del año"] == semana_anterior][indicadores]
            .sum(numeric_only=True)
        )
    else:
        anterior = pd.Series(0, index=indicadores)

    resumen = pd.DataFrame({
        "Indicador": indicadores,
        f"Dato sem {semana_actual}": [actual.get(i, 0) for i in indicadores],
        "Variación vs sem ant": [actual.get(i, 0) - anterior.get(i, 0) for i in indicadores],
        "% Var": [
            np.nan if anterior.get(i, 0) == 0
            else (actual.get(i, 0) - anterior.get(i, 0)) / anterior.get(i, 0)
            for i in indicadores
        ],
    })

    return resumen, semana_anterior


def aplicar_formato_tabla(df_tabla: pd.DataFrame):
    df_fmt = df_tabla.copy()

    for c in df_fmt.columns:
        if c == "Indicador" or df_fmt[c].dtype == "object":
            continue

        if "% Var" in c:
            df_fmt[c] = df_fmt[c].apply(lambda x: formato_pct(x, 1, True))

        elif "Variación" in c or c.startswith("Var "):
            df_fmt[c] = df_fmt[c].apply(formato_variacion)

        else:
            df_fmt[c] = df_fmt[c].apply(formato_numero)

    return df_fmt


def tabla_por_nivel(df_filtrado: pd.DataFrame, nivel: str, indicadores: list[str], semana_actual: int):
    semanas_previas = sorted([
        int(s) for s in df_filtrado["Semana del año"].dropna().unique()
        if int(s) < int(semana_actual)
    ])

    semana_anterior = semanas_previas[-1] if semanas_previas else None

    actual = (
        df_filtrado[df_filtrado["Semana del año"] == semana_actual]
        .groupby(nivel, dropna=False)[indicadores]
        .sum()
        .reset_index()
    )

    if semana_anterior is not None:
        anterior = (
            df_filtrado[df_filtrado["Semana del año"] == semana_anterior]
            .groupby(nivel, dropna=False)[indicadores]
            .sum()
            .reset_index()
        )
    else:
        anterior = actual[[nivel]].copy()
        for i in indicadores:
            anterior[i] = 0

    salida = actual.merge(
        anterior,
        on=nivel,
        how="left",
        suffixes=("", " sem ant")
    ).fillna(0)

    for i in indicadores:
        salida[f"Var {i}"] = salida[i] - salida[f"{i} sem ant"]

    columnas = [nivel]
    for i in indicadores:
        columnas += [i, f"Var {i}"]

    return salida[columnas].sort_values(by=indicadores[0], ascending=False)


def construir_top_bottom_por_variable(
    df_filtrado: pd.DataFrame,
    nivel_top_bottom: str,
    variables_top_bottom: list[str],
    semana_actual: int,
    tipo_ranking: str = "Top",
    cantidad: int = 10
) -> pd.DataFrame:

    if df_filtrado is None or df_filtrado.empty:
        return pd.DataFrame()

    if nivel_top_bottom not in df_filtrado.columns:
        return pd.DataFrame()

    variables_validas = [
        v for v in variables_top_bottom
        if v in df_filtrado.columns and pd.api.types.is_numeric_dtype(df_filtrado[v])
    ]

    if not variables_validas:
        return pd.DataFrame()

    df_semana = df_filtrado[df_filtrado["Semana del año"] == semana_actual].copy()

    if df_semana.empty:
        return pd.DataFrame()

    agrupado = (
        df_semana
        .groupby(nivel_top_bottom, dropna=False)[variables_validas]
        .sum(numeric_only=True)
        .reset_index()
    )

    tablas = []
    ascendente = True if tipo_ranking == "Bottom" else False

    for variable in variables_validas:
        tabla_variable = (
            agrupado[[nivel_top_bottom, variable]]
            .sort_values(by=variable, ascending=ascendente)
            .head(cantidad)
            .copy()
        )

        tabla_variable.insert(0, "Tipo", tipo_ranking)
        tabla_variable.insert(1, "Variable", variable)
        tabla_variable.insert(2, "Ranking", range(1, len(tabla_variable) + 1))
        tabla_variable = tabla_variable.rename(columns={
            nivel_top_bottom: "Estructura",
            variable: "Valor"
        })

        tablas.append(tabla_variable)

    if not tablas:
        return pd.DataFrame()

    return pd.concat(tablas, ignore_index=True)[["Tipo", "Variable", "Ranking", "Estructura", "Valor"]]


def aplicar_formato_top_bottom(df_top_bottom: pd.DataFrame) -> pd.DataFrame:
    df_fmt = df_top_bottom.copy()

    if "Valor" in df_fmt.columns:
        df_fmt["Valor"] = df_fmt["Valor"].apply(formato_numero)

    return df_fmt


def obtener_ultima_semana_cobranza(df_cobranza_base: pd.DataFrame) -> int | None:
    if df_cobranza_base is None or df_cobranza_base.empty:
        return None
    if "Semana del año" not in df_cobranza_base.columns:
        return None

    df_tmp = df_cobranza_base.copy()
    df_tmp["Semana del año"] = pd.to_numeric(df_tmp["Semana del año"], errors="coerce")
    df_tmp = df_tmp.dropna(subset=["Semana del año"])

    if df_tmp.empty:
        return None

    if "Año" in df_tmp.columns:
        df_tmp["Año"] = pd.to_numeric(df_tmp["Año"], errors="coerce")
        df_tmp = df_tmp.sort_values(["Año", "Semana del año"])
    else:
        df_tmp = df_tmp.sort_values("Semana del año")

    return int(df_tmp.iloc[-1]["Semana del año"])


def construir_top_bottom_cobranza(
    df_cobranza_base: pd.DataFrame,
    nivel_top_bottom: str,
    variable_top_bottom: str,
    col_cuota: str,
    col_pago: str,
    col_cump: str,
    col_mejor: str | None,
    col_peor: str | None,
    tipo_ranking: str = "Top",
    cantidad: int = 10,
    semana_objetivo: int | None = None
) -> pd.DataFrame:
    if df_cobranza_base is None or df_cobranza_base.empty:
        return pd.DataFrame()

    if nivel_top_bottom not in df_cobranza_base.columns:
        return pd.DataFrame()

    if "Semana del año" not in df_cobranza_base.columns:
        return pd.DataFrame()

    df_tmp = df_cobranza_base.copy()
    df_tmp["Semana del año"] = pd.to_numeric(df_tmp["Semana del año"], errors="coerce")
    df_tmp = df_tmp.dropna(subset=["Semana del año"])

    if df_tmp.empty:
        return pd.DataFrame()

    if semana_objetivo is None:
        semana_objetivo = obtener_ultima_semana_cobranza(df_tmp)

    if semana_objetivo is None:
        return pd.DataFrame()

    df_semana = df_tmp[df_tmp["Semana del año"] == int(semana_objetivo)].copy()

    if df_semana.empty:
        return pd.DataFrame()

    for col in [col_cuota, col_pago, col_mejor, col_peor]:
        if col and col in df_semana.columns:
            df_semana[col] = pd.to_numeric(df_semana[col], errors="coerce").fillna(0)

    agg = {}
    if col_cuota and col_cuota in df_semana.columns:
        agg[col_cuota] = "sum"
    if col_pago and col_pago in df_semana.columns:
        agg[col_pago] = "sum"
    if col_mejor and col_mejor in df_semana.columns:
        agg[col_mejor] = "max"
    if col_peor and col_peor in df_semana.columns:
        agg[col_peor] = "min"

    if not agg:
        return pd.DataFrame()

    agrupado = (
        df_semana
        .groupby(nivel_top_bottom, dropna=False)
        .agg(agg)
        .reset_index()
    )

    if col_cuota in agrupado.columns and col_pago in agrupado.columns:
        agrupado[col_cump] = np.where(
            agrupado[col_cuota] == 0,
            np.nan,
            agrupado[col_pago] / agrupado[col_cuota]
        )

    variables_validas = [c for c in [col_cuota, col_pago, col_cump, col_mejor, col_peor] if c and c in agrupado.columns]

    if variable_top_bottom not in variables_validas:
        return pd.DataFrame()

    ascendente = True if tipo_ranking == "Bottom" else False

    salida = (
        agrupado[[nivel_top_bottom, variable_top_bottom]]
        .sort_values(variable_top_bottom, ascending=ascendente)
        .head(int(cantidad))
        .copy()
    )

    salida.insert(0, "Tipo", tipo_ranking)
    salida.insert(1, "Variable", variable_top_bottom)
    salida.insert(2, "Ranking", range(1, len(salida) + 1))
    salida = salida.rename(columns={
        nivel_top_bottom: "Estructura",
        variable_top_bottom: "Valor"
    })

    return salida[["Tipo", "Variable", "Ranking", "Estructura", "Valor"]]


def aplicar_formato_top_bottom_cobranza(df_top_bottom: pd.DataFrame) -> pd.DataFrame:
    df_fmt = df_top_bottom.copy()

    if "Valor" in df_fmt.columns:
        def _fmt_valor(row):
            variable = str(row.get("Variable", ""))
            valor = row.get("Valor")
            if "cumplimiento" in normalizar_texto_tc(variable).lower() or variable.strip().startswith("%"):
                return formato_pct(valor, 2, False)
            return formato_numero(valor)

        df_fmt["Valor"] = df_fmt.apply(_fmt_valor, axis=1)

    return df_fmt


def generar_comentario_top_bottom_cobranza(tabla_top_bottom, tipo_top_bottom, nivel_top_bottom, semana_actual):
    if tabla_top_bottom is None or tabla_top_bottom.empty:
        return "No hay información suficiente para comentar el Top / Bottom de cobranza."

    primera = tabla_top_bottom.iloc[0]
    variable = primera["Variable"]
    valor = primera["Valor"]

    if "cumplimiento" in normalizar_texto_tc(variable).lower() or str(variable).strip().startswith("%"):
        valor_fmt = formato_pct(valor, 2, False)
    else:
        valor_fmt = formato_numero(valor)

    comentario = (
        f"En la semana {semana_actual}, el {tipo_top_bottom} de cobranza por {nivel_top_bottom} "
        f"ubica a {primera['Estructura']} como principal registro en {variable}, con {valor_fmt}."
    )

    if tipo_top_bottom == "Top":
        comentario += " Esta vista permite identificar las estructuras con mayor aportación o cumplimiento dentro de cobranza."
    else:
        comentario += " Esta vista permite detectar las estructuras con menor volumen o menor cumplimiento dentro de cobranza."

    return comentario


def crear_llave_coordinadora_marca(df_base: pd.DataFrame, columna_id: str = "coordinadora_id") -> pd.DataFrame:
    df_tmp = df_base.copy()

    if columna_id not in df_tmp.columns:
        return df_tmp

    columnas_llave = [columna_id]

    for c in ["País", "Marca"]:
        if c in df_tmp.columns:
            columnas_llave.append(c)

    df_tmp["_llave_coordinadora_marca"] = (
        df_tmp[columnas_llave]
        .astype(str)
        .fillna("")
        .agg("|".join, axis=1)
    )

    return df_tmp


def obtener_categoria_unica_por_semana(
    df_base: pd.DataFrame,
    semana: int,
    columna_id: str = "coordinadora_id",
    columna_categoria: str = "Tipo Coordinadora"
):
    df_semana = df_base[df_base["Semana del año"] == semana].copy()

    columnas_necesarias = [columna_id, columna_categoria]

    for c in ["coordinadora_id", "País", "Marca"]:
        if c in df_semana.columns and c not in columnas_necesarias:
            columnas_necesarias.append(c)

    df_semana = df_semana[columnas_necesarias].dropna(
        subset=[columna_id, columna_categoria]
    )

    if df_semana.empty:
        return pd.DataFrame(columns=columnas_necesarias)

    prioridad = {
        "Productiva": 4,
        "En Desarrollo": 3,
        "Improductiva": 2,
        "Secundaria": 1,
    }

    df_semana["_prioridad_categoria"] = (
        df_semana[columna_categoria]
        .map(prioridad)
        .fillna(0)
    )

    salida = (
        df_semana
        .sort_values([columna_id, "_prioridad_categoria"], ascending=[True, False])
        .drop_duplicates(subset=[columna_id], keep="first")
        .drop(columns=["_prioridad_categoria"])
        .reset_index(drop=True)
    )

    return salida


def matriz_desplazamiento_coordinadoras(
    df_filtrado: pd.DataFrame,
    semana_origen: int,
    semana_destino: int,
    columna_id: str = "coordinadora_id",
    columna_categoria: str = "Tipo Coordinadora"
):
    if columna_id not in df_filtrado.columns:
        return None, None

    if columna_categoria not in df_filtrado.columns:
        return None, None

    df_tmp = crear_llave_coordinadora_marca(
        df_base=df_filtrado,
        columna_id=columna_id
    )

    columna_llave = "_llave_coordinadora_marca"

    df_origen = obtener_categoria_unica_por_semana(
        df_base=df_tmp,
        semana=semana_origen,
        columna_id=columna_llave,
        columna_categoria=columna_categoria
    ).rename(columns={columna_categoria: "Semana anterior"})

    df_destino = obtener_categoria_unica_por_semana(
        df_base=df_tmp,
        semana=semana_destino,
        columna_id=columna_llave,
        columna_categoria=columna_categoria
    ).rename(columns={columna_categoria: "Semana actual"})

    movimientos = df_origen.merge(
        df_destino,
        on=columna_llave,
        how="outer",
        suffixes=(" origen", " destino")
    )

    movimientos["Semana anterior"] = movimientos["Semana anterior"].fillna("Nueva")
    movimientos["Semana actual"] = movimientos["Semana actual"].fillna("Baja")

    if movimientos.empty:
        return movimientos, pd.DataFrame()

    matriz = pd.crosstab(
        movimientos["Semana anterior"],
        movimientos["Semana actual"]
    )

    orden_filas = ["Productiva", "En Desarrollo", "Improductiva", "Secundaria", "Nueva"]
    orden_columnas = ["Productiva", "En Desarrollo", "Improductiva", "Secundaria", "Baja"]

    filas_ordenadas = [c for c in orden_filas if c in matriz.index]
    columnas_ordenadas = [c for c in orden_columnas if c in matriz.columns]

    otras_filas = [c for c in matriz.index if c not in filas_ordenadas]
    otras_columnas = [c for c in matriz.columns if c not in columnas_ordenadas]

    matriz = matriz.loc[
        filas_ordenadas + otras_filas,
        columnas_ordenadas + otras_columnas
    ]

    matriz["Total general"] = matriz.sum(axis=1)
    total_general = matriz.sum(axis=0).to_frame().T
    total_general.index = ["Total general"]
    matriz = pd.concat([matriz, total_general])

    matriz.index.name = "Semana Anterior"
    matriz.columns.name = "Tipo Coordinadora Semana Actual"

    return movimientos, matriz


def estilo_matriz_desplazamiento(df_matriz: pd.DataFrame):
    ranking = {
        "Secundaria": 0,
        "Improductiva": 1,
        "En Desarrollo": 2,
        "Productiva": 3,
    }

    def colorear(data):
        estilos = pd.DataFrame("", index=data.index, columns=data.columns)

        for fila in data.index:
            for col in data.columns:
                if fila == "Total general" or col == "Total general":
                    estilos.loc[fila, col] = (
                        "background-color: #082567; color: white; "
                        "font-weight: 800; text-align: center;"
                    )

                elif fila == "Nueva":
                    estilos.loc[fila, col] = (
                        "background-color: #dbeafe; color: #1d4ed8; "
                        "font-weight: 800; text-align: center;"
                    )

                elif col == "Baja":
                    estilos.loc[fila, col] = (
                        "background-color: #ffedd5; color: #c2410c; "
                        "font-weight: 800; text-align: center;"
                    )

                elif fila in ranking and col in ranking:
                    if ranking[col] > ranking[fila]:
                        estilos.loc[fila, col] = "color: #059669; font-weight: 800; text-align: center;"
                    elif ranking[col] < ranking[fila]:
                        estilos.loc[fila, col] = "color: #dc2626; font-weight: 800; text-align: center;"
                    else:
                        estilos.loc[fila, col] = (
                            "background-color: #f1f5f9; color: #111827; "
                            "font-weight: 700; text-align: center;"
                        )
                else:
                    estilos.loc[fila, col] = "text-align: center;"

        return estilos

    return df_matriz.style.apply(colorear, axis=None).format("{:,.0f}")


def calcular_resumen_movimientos(movimientos: pd.DataFrame) -> dict:
    if movimientos is None or movimientos.empty:
        return {
            "Movimientos totales": 0,
            "Pérdida de Productivas": 0,
            "Aumento de improductivas": 0,
            "Aumento de productivas": 0,
            "Aumento de desarrollo": 0,
        }

    mov = movimientos.copy()
    movimientos_reales = mov[mov["Semana anterior"] != mov["Semana actual"]].copy()

    return {
        "Movimientos totales": len(movimientos_reales),
        "Pérdida de Productivas": len(
            mov[
                (mov["Semana anterior"] == "Productiva") &
                (mov["Semana actual"] != "Productiva") &
                (mov["Semana actual"] != "Baja")
            ]
        ),
        "Aumento de improductivas": len(
            mov[
                (mov["Semana anterior"] != "Improductiva") &
                (mov["Semana actual"] == "Improductiva") &
                (mov["Semana anterior"] != "Nueva")
            ]
        ),
        "Aumento de productivas": len(
            mov[
                (mov["Semana anterior"] != "Productiva") &
                (mov["Semana actual"] == "Productiva") &
                (mov["Semana anterior"] != "Nueva")
            ]
        ),
        "Aumento de desarrollo": len(
            mov[
                (mov["Semana anterior"] != "En Desarrollo") &
                (mov["Semana actual"] == "En Desarrollo") &
                (mov["Semana anterior"] != "Nueva")
            ]
        ),
    }


def mostrar_cuadro_resumen_movimientos(movimientos: pd.DataFrame):
    resumen_mov = calcular_resumen_movimientos(movimientos)

    html_resumen = f"""
    <html>
    <head>
        <style>
            body {{
                margin: 0;
                padding: 0;
                font-family: Arial, sans-serif;
                background: transparent;
            }}
            .cuadro-movimientos {{
                border: 1px solid #6b7280;
                border-radius: 2px;
                overflow: hidden;
                background: #ffffff;
                width: 100%;
                box-sizing: border-box;
            }}
            .cuadro-header {{
                background: #082567;
                color: white;
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 5px 10px;
                font-size: 17px;
                font-weight: 500;
            }}
            .cuadro-body {{
                padding: 12px 10px 8px 10px;
                font-size: 20px;
                line-height: 1.25;
                color: #111827;
            }}
            .fila-mov {{
                display: flex;
                justify-content: space-between;
                gap: 12px;
                margin-bottom: 2px;
                white-space: nowrap;
            }}
            .valor-rojo {{
                color: red;
                min-width: 45px;
                text-align: right;
            }}
            .valor-verde {{
                color: #059669;
                min-width: 45px;
                text-align: right;
            }}
        </style>
    </head>
    <body>
        <div class="cuadro-movimientos">
            <div class="cuadro-header">
                <span>Movimientos totales</span>
                <span>{resumen_mov["Movimientos totales"]:,.0f}</span>
            </div>

            <div class="cuadro-body">
                <div class="fila-mov">
                    <span>Pérdida de Productivas</span>
                    <span class="valor-rojo">{resumen_mov["Pérdida de Productivas"]:,.0f}</span>
                </div>
                <div class="fila-mov">
                    <span>Aumento de improductivas</span>
                    <span class="valor-rojo">{resumen_mov["Aumento de improductivas"]:,.0f}</span>
                </div>
                <div class="fila-mov">
                    <span>Aumento de productivas</span>
                    <span class="valor-verde">{resumen_mov["Aumento de productivas"]:,.0f}</span>
                </div>
                <div class="fila-mov">
                    <span>Aumento de desarrollo</span>
                    <span class="valor-verde">{resumen_mov["Aumento de desarrollo"]:,.0f}</span>
                </div>
            </div>
        </div>
    </body>
    </html>
    """

    components.html(html_resumen, height=150, scrolling=False)


# ============================================================
# FUNCIONES COBRANZA
# ============================================================
def preparar_cobranza(df_cobranza: pd.DataFrame):
    df_tmp = df_cobranza.copy()
    df_tmp = normalizar_columnas(df_tmp)

    col_cuota = detectar_columna(COLUMNAS_COBRANZA_CUOTA, df_tmp.columns)
    col_pago = detectar_columna(COLUMNAS_COBRANZA_PAGO, df_tmp.columns)
    col_cump = detectar_columna(COLUMNAS_COBRANZA_CUMPLIMIENTO, df_tmp.columns)
    col_mejor = detectar_columna(COLUMNAS_COBRANZA_MEJOR, df_tmp.columns)
    col_peor = detectar_columna(COLUMNAS_COBRANZA_PEOR, df_tmp.columns)

    if "Semana del año" not in df_tmp.columns:
        return df_tmp, col_cuota, col_pago, col_cump, col_mejor, col_peor

    if col_cuota:
        df_tmp[col_cuota] = pd.to_numeric(df_tmp[col_cuota], errors="coerce").fillna(0)

    if col_pago:
        df_tmp[col_pago] = pd.to_numeric(df_tmp[col_pago], errors="coerce").fillna(0)

    if col_mejor:
        df_tmp[col_mejor] = pd.to_numeric(df_tmp[col_mejor], errors="coerce").fillna(0)

    if col_peor:
        df_tmp[col_peor] = pd.to_numeric(df_tmp[col_peor], errors="coerce").fillna(0)

    if col_cump:
        df_tmp[col_cump] = pd.to_numeric(df_tmp[col_cump], errors="coerce")
        if df_tmp[col_cump].dropna().abs().median() > 1.5:
            df_tmp[col_cump] = df_tmp[col_cump] / 100

    elif col_cuota and col_pago:
        df_tmp["% de Cumplimiento"] = np.where(
            df_tmp[col_cuota] == 0,
            np.nan,
            df_tmp[col_pago] / df_tmp[col_cuota]
        )
        col_cump = "% de Cumplimiento"

    return df_tmp, col_cuota, col_pago, col_cump, col_mejor, col_peor


def consolidar_cobranza(
    df_cobranza: pd.DataFrame,
    col_cuota: str,
    col_pago: str,
    col_cump: str,
    col_mejor: str | None,
    col_peor: str | None,
    nivel: str | None = None
):
    df_tmp = df_cobranza.copy()

    if "Semana del año" not in df_tmp.columns:
        return pd.DataFrame()

    grupo = []

    if "Año" in df_tmp.columns:
        grupo.append("Año")

    grupo.append("Semana del año")

    if nivel and nivel in df_tmp.columns:
        grupo.append(nivel)

def consolidar_cobranza(
    df_cobranza: pd.DataFrame,
    col_cuota: str,
    col_pago: str,
    col_cump: str,
    col_mejor: str | None,
    col_peor: str | None,
    nivel: str | None = None
):
    df_tmp = df_cobranza.copy()

    if "Semana del año" not in df_tmp.columns:
        return pd.DataFrame()

    grupo = []

    if "Año" in df_tmp.columns:
        grupo.append("Año")

    grupo.append("Semana del año")

    if nivel and nivel in df_tmp.columns:
        grupo.append(nivel)

    # Solo se suman cuota y pago.
    # Mejor semana y peor semana se recalculan después del agrupado,
    # para que respeten moneda local / pesos mexicanos y el filtro aplicado.
    agg = {
        col_cuota: "sum",
        col_pago: "sum",
    }

    salida = (
        df_tmp
        .groupby(grupo, dropna=False)
        .agg(agg)
        .reset_index()
    )

    # Recalcula cumplimiento sobre los datos ya agrupados
    salida[col_cump] = np.where(
        salida[col_cuota] == 0,
        np.nan,
        salida[col_pago] / salida[col_cuota]
    )

    # Recalcula mejor y peor semana sobre el pago ya convertido y agrupado.
    # Esto corrige que "Mejor semana" quede por debajo de los recuperados.
    salida["Mejor semana"] = salida[col_pago].max()
    salida["Peor semana"] = salida[col_pago].min()

    col_mejor = "Mejor semana"
    col_peor = "Peor semana"

    salida = salida.sort_values(
        ["Año", "Semana del año"] if "Año" in salida.columns else ["Semana del año"]
    ).reset_index(drop=True)

    salida["Orden"] = range(len(salida))

    salida["Etiqueta semana"] = salida.apply(
        lambda r: f"{int(r['Año'])} S{int(r['Semana del año'])}"
        if "Año" in salida.columns and pd.notna(r["Año"])
        else f"S{int(r['Semana del año'])}",
        axis=1
    )

    return salida

def formato_tabla_detalle_cobranza(df_tabla, col_cuota, col_pago, col_cump, col_mejor, col_peor):
    df_fmt = df_tabla.copy()

    for c in [col_cuota, col_pago, col_mejor, col_peor]:
        if c and c in df_fmt.columns:
            df_fmt[c] = df_fmt[c].apply(formato_numero)

    if col_cump and col_cump in df_fmt.columns:
        df_fmt[col_cump] = df_fmt[col_cump].apply(lambda x: formato_pct(x, 2, False))

    return df_fmt


def crear_tabla_ultimas_5_cobranza(evol, col_cuota, col_pago, col_cump):
    """
    Construye la tabla de las últimas 5 semanas de cobranza.
    Para cada variable muestra el dato de la semana y, a un lado,
    la variación contra la semana inmediatamente anterior.
    """
    df_tmp = evol.copy().sort_values(
        ["Año", "Semana del año"] if "Año" in evol.columns else ["Semana del año"]
    ).reset_index(drop=True)

    if df_tmp.empty:
        return pd.DataFrame()

    # Variaciones contra la semana anterior dentro del histórico completo,
    # no solo dentro de las últimas 5 semanas.
    df_tmp[f"Var {col_cuota}"] = df_tmp[col_cuota].diff()
    df_tmp[f"Var {col_pago}"] = df_tmp[col_pago].diff()
    df_tmp[f"Var {col_cump}"] = df_tmp[col_cump].diff()

    # Se deja solo la etiqueta visible de la semana.
    # No se muestran las columnas Año ni Semana del año en esta tabla.
    columnas = [
        "Etiqueta semana",
        col_cuota,
        f"Var {col_cuota}",
        col_pago,
        f"Var {col_pago}",
        col_cump,
        f"Var {col_cump}",
    ]

    columnas = [c for c in columnas if c in df_tmp.columns]

    salida = df_tmp.tail(5)[columnas].copy()

    salida = salida.rename(columns={
        "Etiqueta semana": "Semana",
        f"Var {col_cuota}": f"Var {col_cuota}",
        f"Var {col_pago}": f"Var {col_pago}",
        f"Var {col_cump}": f"Var {col_cump}",
    })

    return salida


def formato_ultimas_5_cobranza(tabla, col_cuota, col_pago, col_cump):
    df_fmt = tabla.copy()

    columnas_monto = [
        col_cuota,
        f"Var {col_cuota}",
        col_pago,
        f"Var {col_pago}",
    ]

    for c in columnas_monto:
        if c in df_fmt.columns:
            if str(c).startswith("Var " ):
                df_fmt[c] = df_fmt[c].apply(formato_variacion)
            else:
                df_fmt[c] = df_fmt[c].apply(formato_numero)

    if col_cump in df_fmt.columns:
        df_fmt[col_cump] = df_fmt[col_cump].apply(lambda x: formato_pct(x, 2, False))

    col_var_cump = f"Var {col_cump}"
    if col_var_cump in df_fmt.columns:
        df_fmt[col_var_cump] = df_fmt[col_var_cump].apply(lambda x: formato_pct(x, 2, True))

    return df_fmt


def estilo_ultimas_5_cobranza(tabla):
    def pintar(data):
        estilos = pd.DataFrame("", index=data.index, columns=data.columns)

        for idx in data.index:
            for col in data.columns:
                if col == "Semana":
                    estilos.loc[idx, col] = (
                        "background-color: #082567; color: white; "
                        "font-weight: 800; text-align: center;"
                    )
                elif str(col).startswith("Var "):
                    estilos.loc[idx, col] = (
                        "background-color: #f8fafc; color: #082567; "
                        "font-weight: 800; text-align: right;"
                    )
                else:
                    estilos.loc[idx, col] = "text-align: right;"

        return estilos

    return tabla.style.apply(pintar, axis=None)


def grafica_cumplimiento(evol, col_cump):
    df_tmp = evol.copy().sort_values(
        ["Año", "Semana del año"] if "Año" in evol.columns else ["Semana del año"]
    ).reset_index(drop=True)

    # Asegura que se grafique TODO el histórico que llegue a evol_cobranza.
    # Gris = % de Cumplimiento real.
    # Azul = complemento o diferencia necesaria para llegar al % de cumplimiento de la mejor semana.
    mejor_cump = df_tmp[col_cump].max()
    df_tmp["Dif % Mejor Semana"] = (mejor_cump - df_tmp[col_cump]).clip(lower=0)
    df_tmp["Texto Dif % Mejor Semana"] = df_tmp["Dif % Mejor Semana"].apply(
        lambda x: formato_pct(x, 2, False) if pd.notna(x) and x > 0 else ""
    )

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=df_tmp["Etiqueta semana"],
            y=df_tmp[col_cump],
            name="% Cumplimiento",
            marker_color="#d9d9d9",
            text=df_tmp[col_cump].apply(lambda x: formato_pct(x, 2, False)),
            textposition="inside",
            textangle=-90,
            textfont=dict(color="#082567", size=10),
            cliponaxis=False,
            hovertemplate=(
                "<b>Semana:</b> %{x}<br>"
                "<b>% Cumplimiento:</b> %{y:.2%}<br>"
                f"<b>Mejor semana:</b> {mejor_cump:.2%}"
                "<extra></extra>"
            )
        )
    )

    fig.add_trace(
        go.Bar(
            x=df_tmp["Etiqueta semana"],
            y=df_tmp["Dif % Mejor Semana"],
            name="Dif % Mejor Semana",
            marker_color="#1f77b4",
            text=df_tmp["Texto Dif % Mejor Semana"],
            textposition="inside",
            textangle=-90,
            textfont=dict(color="white", size=10),
            cliponaxis=False,
            hovertemplate=(
                "<b>Semana:</b> %{x}<br>"
                "<b>Complemento vs mejor:</b> %{y:.2%}<br>"
                f"<b>Mejor semana:</b> {mejor_cump:.2%}"
                "<extra></extra>"
            )
        )
    )

    y_max = max(1, mejor_cump * 1.08 if pd.notna(mejor_cump) else 1)

    fig.update_layout(
        height=420,
        barmode="stack",
        bargap=0.25,
        yaxis_tickformat=".0%",
        yaxis_range=[0, y_max],
        xaxis_title=None,
        yaxis_title="% Cumplimiento y Dif % Mejor Semana",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        margin=dict(t=35, b=85, l=45, r=20),
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,1)",
        font=dict(color="#082567", size=11),
        xaxis=dict(
            gridcolor="rgba(148,163,184,0.25)",
            tickangle=-45,
            type="category",
            categoryorder="array",
            categoryarray=df_tmp["Etiqueta semana"].tolist()
        ),
        yaxis=dict(gridcolor="rgba(148,163,184,0.25)")
    )

    return fig

def grafica_cuota_pago(evol, col_cuota, col_pago, col_mejor, col_peor):
    df_tmp = evol.copy().sort_values(
        ["Año", "Semana del año"] if "Año" in evol.columns else ["Semana del año"]
    ).reset_index(drop=True)

    if df_tmp.empty:
        return go.Figure()

    # Asegura numéricos
    for col in [col_cuota, col_pago, col_mejor, col_peor]:
        if col and col in df_tmp.columns:
            df_tmp[col] = pd.to_numeric(df_tmp[col], errors="coerce").fillna(0)

    # Valores de referencia
    mejor_valor = (
        df_tmp[col_mejor].max()
        if col_mejor and col_mejor in df_tmp.columns
        else df_tmp[col_pago].max()
    )

    peor_valor = (
        df_tmp[col_peor].min()
        if col_peor and col_peor in df_tmp.columns
        else df_tmp[col_pago].min()
    )

    # Si Peor Semana viene en 0, no debe forzar la escala al piso.
    valores_reales = pd.concat([
        df_tmp[col_cuota],
        df_tmp[col_pago]
    ], ignore_index=True)

    valores_reales = valores_reales[
        valores_reales.notna() & 
        np.isfinite(valores_reales) & 
        (valores_reales > 0)
    ]

    if valores_reales.empty:
        y_min = 0
        y_max = 1
    else:
        y_min_real = float(valores_reales.min())
        y_max_real = float(valores_reales.max())
        rango = y_max_real - y_min_real

        if rango == 0:
            margen_inf = y_min_real * 0.12
            margen_sup = y_max_real * 0.18
        else:
            margen_inf = rango * 0.25
            margen_sup = rango * 0.25

        y_min = max(0, y_min_real - margen_inf)
        y_max = y_max_real + margen_sup

    # Solo incluir la línea de mejor semana en escala si está cerca del rango real
    if mejor_valor > 0:
        y_max = max(y_max, mejor_valor * 1.05)

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=df_tmp["Etiqueta semana"],
            y=[mejor_valor] * len(df_tmp),
            mode="lines",
            name="Mejor Semana",
            line=dict(color="#ff7f0e", width=2, dash="dot"),
            hovertemplate="<b>Mejor Semana:</b> %{y:,.0f}<extra></extra>"
        )
    )

    fig.add_trace(
        go.Scatter(
            x=df_tmp["Etiqueta semana"],
            y=df_tmp[col_cuota],
            mode="lines+markers+text",
            name="Cuota total",
            line=dict(color="black", width=2.5),
            marker=dict(color="black", size=6),
            text=df_tmp[col_cuota].apply(formato_millones),
            textposition="top center",
            textfont=dict(color="#1d4ed8", size=10),
            cliponaxis=False,
            hovertemplate="<b>Semana:</b> %{x}<br><b>Cuota total:</b> %{y:,.0f}<extra></extra>"
        )
    )

    fig.add_trace(
        go.Scatter(
            x=df_tmp["Etiqueta semana"],
            y=df_tmp[col_pago],
            mode="lines+markers+text",
            name="Pago total",
            line=dict(color="#00b050", width=2.5),
            marker=dict(color="#00b050", size=6),
            text=df_tmp[col_pago].apply(formato_millones),
            textposition="bottom center",
            textfont=dict(color="#00b050", size=10),
            cliponaxis=False,
            hovertemplate="<b>Semana:</b> %{x}<br><b>Pago total:</b> %{y:,.0f}<extra></extra>"
        )
    )

    # La línea peor semana se muestra, pero si vale 0 no se deja que aplaste la escala.
    fig.add_trace(
        go.Scatter(
            x=df_tmp["Etiqueta semana"],
            y=[peor_valor] * len(df_tmp),
            mode="lines",
            name="Peor Semana",
            line=dict(color="red", width=2, dash="dot"),
            hovertemplate="<b>Peor Semana:</b> %{y:,.0f}<extra></extra>"
        )
    )

    fig.update_layout(
        height=500,
        xaxis_title=None,
        yaxis_title=None,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0
        ),
        margin=dict(t=45, b=85, l=65, r=25),
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,1)",
        font=dict(color="#082567", size=11),
        hovermode="x unified",
        dragmode=False,
        xaxis=dict(
            gridcolor="rgba(148,163,184,0.25)",
            tickangle=-45,
            type="category",
            categoryorder="array",
            categoryarray=df_tmp["Etiqueta semana"].tolist(),
            fixedrange=True
        ),
        yaxis=dict(
            gridcolor="rgba(148,163,184,0.25)",
            range=[y_min, y_max],
            fixedrange=True,
            tickformat=",.0f"
        )
    )

    return fig

# ============================================================
# COMENTARIOS
# ============================================================
def generar_comentario_resumen(resumen: pd.DataFrame, semana_actual: int, semana_anterior):
    if resumen is None or resumen.empty:
        return "No hay información suficiente para generar un comentario."

    comentario = (
        f"En la semana {semana_actual}, el resumen muestra el comportamiento general "
        f"de los indicadores seleccionados."
    )

    if semana_anterior is not None and "Variación vs sem ant" in resumen.columns:
        positivos = resumen[resumen["Variación vs sem ant"] > 0].copy()
        negativos = resumen[resumen["Variación vs sem ant"] < 0].copy()

        if not positivos.empty:
            mejor = positivos.sort_values("Variación vs sem ant", ascending=False).iloc[0]
            comentario += (
                f" El mayor avance se observa en {mejor['Indicador']}, "
                f"con una variación de {_fmt_comentario(mejor['Variación vs sem ant'])} "
                f"respecto a la semana {semana_anterior}."
            )

        if not negativos.empty:
            peor = negativos.sort_values("Variación vs sem ant", ascending=True).iloc[0]
            comentario += (
                f" La principal oportunidad se concentra en {peor['Indicador']}, "
                f"con una disminución de {_fmt_comentario(peor['Variación vs sem ant'])}."
            )

    return comentario



def _columna_pais_disponible(df_base: pd.DataFrame) -> str | None:
    for c in ["País", "Pais", "PAIS"]:
        if c in df_base.columns:
            return c
    return None


def _direccion_favorable_indicador(indicador: str) -> int:
    """
    Define cómo interpretar la variación:
    +1 = subir es bueno.
    -1 = bajar es bueno.
    """
    nombre = normalizar_texto_tc(indicador).lower()
    if any(t in nombre for t in ["falta", "nunca abon", "atraso"]):
        return -1
    return 1


def _formato_valor_resumen_ia(indicador: str, valor) -> str:
    nombre = normalizar_texto_tc(indicador).lower()
    if "cumplimiento" in nombre or str(indicador).strip().startswith("%"):
        return formato_pct(valor, 2, True)
    return formato_variacion(valor)


def _construir_base_comparativo_pais(
    df_base: pd.DataFrame,
    indicadores: list[str],
    semana_actual: int,
    semana_anterior
) -> pd.DataFrame:
    col_pais = _columna_pais_disponible(df_base)
    if col_pais is None or semana_anterior is None or df_base is None or df_base.empty:
        return pd.DataFrame()

    indicadores_validos = [
        c for c in indicadores
        if c in df_base.columns and pd.api.types.is_numeric_dtype(df_base[c])
    ]
    if not indicadores_validos:
        return pd.DataFrame()

    actual = (
        df_base[df_base["Semana del año"] == semana_actual]
        .groupby(col_pais, dropna=False)[indicadores_validos]
        .sum(numeric_only=True)
        .reset_index()
    )
    anterior = (
        df_base[df_base["Semana del año"] == semana_anterior]
        .groupby(col_pais, dropna=False)[indicadores_validos]
        .sum(numeric_only=True)
        .reset_index()
    )

    salida = actual.merge(anterior, on=col_pais, how="outer", suffixes=("", " sem ant")).fillna(0)

    for ind in indicadores_validos:
        salida[f"Var {ind}"] = salida[ind] - salida[f"{ind} sem ant"]
        salida[f"% Var {ind}"] = np.where(
            salida[f"{ind} sem ant"] == 0,
            np.nan,
            salida[f"Var {ind}"] / salida[f"{ind} sem ant"]
        )

    salida = salida.rename(columns={col_pais: "País"})
    return salida


def _frase_metricas_relevantes_pais(fila: pd.Series, indicadores: list[str], max_metricas: int = 3) -> str:
    hallazgos = []
    for ind in indicadores:
        col_var = f"Var {ind}"
        if col_var not in fila.index:
            continue
        var = fila.get(col_var, 0)
        if pd.isna(var) or float(var) == 0:
            continue
        direccion = _direccion_favorable_indicador(ind)
        impacto = float(var) * direccion
        hallazgos.append((abs(impacto), impacto, ind, float(var)))

    hallazgos = sorted(hallazgos, key=lambda x: x[0], reverse=True)[:max_metricas]
    partes = []
    for _, impacto, ind, var in hallazgos:
        if impacto >= 0:
            partes.append(f"{ind} mejoró {_formato_valor_resumen_ia(ind, var)}")
        else:
            partes.append(f"{ind} presionó {_formato_valor_resumen_ia(ind, var)}")

    return "; ".join(partes)


def generar_resumen_ia_paises(
    df_base: pd.DataFrame,
    resumen: pd.DataFrame,
    indicadores: list[str],
    semana_actual: int,
    semana_anterior,
    modo_moneda: str,
    filtros_aplicados: dict | None = None
) -> str:
    """
    Genera un texto ejecutivo tipo IA usando los datos visibles en el tablero.
    No depende de una API externa: interpreta avances, retrocesos, países destacados
    y áreas de oportunidad con base en las variaciones contra la semana anterior.
    """
    if resumen is None or resumen.empty:
        return "No hay información suficiente para generar el resumen ejecutivo de la semana."

    if semana_anterior is None:
        return (
            f"En la semana {semana_actual} se cuenta con información para describir el nivel actual, "
            "pero no existe una semana anterior disponible para construir una lectura comparativa. "
            "La recomendación es validar la carga histórica para identificar avances, retrocesos y áreas de oportunidad."
        )

    moneda_txt = etiqueta_moneda(modo_moneda)
    col_dato = f"Dato sem {semana_actual}"

    positivos = resumen[resumen["Variación vs sem ant"] > 0].copy() if "Variación vs sem ant" in resumen.columns else pd.DataFrame()
    negativos = resumen[resumen["Variación vs sem ant"] < 0].copy() if "Variación vs sem ant" in resumen.columns else pd.DataFrame()

    texto = []
    texto.append(
        f"En la semana {semana_actual}, comparada contra la semana {semana_anterior}, "
        f"el tablero muestra una lectura general en {moneda_txt}. "
    )

    if not positivos.empty:
        top_avances = positivos.sort_values("Variación vs sem ant", ascending=False).head(3)
        avances_txt = ", ".join([
            f"{r['Indicador']} ({formato_variacion(r['Variación vs sem ant'])})"
            for _, r in top_avances.iterrows()
        ])
        texto.append(
            f"Lo más favorable se observa en {avances_txt}, lo que indica un avance operativo y/o financiero frente a la semana previa. "
        )

    if not negativos.empty:
        top_oportunidades = negativos.sort_values("Variación vs sem ant", ascending=True).head(3)
        oportunidades_txt = ", ".join([
            f"{r['Indicador']} ({formato_variacion(r['Variación vs sem ant'])})"
            for _, r in top_oportunidades.iterrows()
        ])
        texto.append(
            f"Las principales alertas del consolidado se concentran en {oportunidades_txt}; estos indicadores deben revisarse para distinguir si el movimiento responde a estacionalidad, recuperación insuficiente o deterioro en la calidad de cartera. "
        )

    comparativo_pais = _construir_base_comparativo_pais(
        df_base=df_base,
        indicadores=indicadores,
        semana_actual=semana_actual,
        semana_anterior=semana_anterior
    )

    if comparativo_pais.empty:
        texto.append(
            "Con los filtros actuales no se puede separar la lectura por país; el análisis se limita al consolidado visible. "
        )
        return "\n\n".join(texto)

    indicadores_validos = [i for i in indicadores if f"Var {i}" in comparativo_pais.columns]
    if not indicadores_validos:
        return "\n\n".join(texto)

    # Score ejecutivo: considera favorable que suban clientes/cartera/saldo sano y que bajen faltas/atraso/nunca abonados.
    comparativo_pais["Score ejecutivo"] = 0.0
    for ind in indicadores_validos:
        direccion = _direccion_favorable_indicador(ind)
        col_pct = f"% Var {ind}"
        col_var = f"Var {ind}"
        base_score = comparativo_pais[col_pct].replace([np.inf, -np.inf], np.nan).fillna(0)
        # Limita impactos extremos para que un país pequeño no domine solo por porcentaje.
        base_score = base_score.clip(lower=-1, upper=1)
        comparativo_pais["Score ejecutivo"] += base_score * direccion
        comparativo_pais[f"Impacto favorable {ind}"] = comparativo_pais[col_var] * direccion

    paises_visibles = comparativo_pais["País"].dropna().astype(str).nunique()

    if paises_visibles > 1:
        mejores = comparativo_pais.sort_values("Score ejecutivo", ascending=False).head(3)
        oportunidades = comparativo_pais.sort_values("Score ejecutivo", ascending=True).head(3)

        mejores_txt = []
        for _, fila in mejores.iterrows():
            detalle = _frase_metricas_relevantes_pais(fila, indicadores_validos, 2)
            if detalle:
                mejores_txt.append(f"{fila['País']}: {detalle}")
            else:
                mejores_txt.append(str(fila["País"]))

        oportunidades_txt = []
        for _, fila in oportunidades.iterrows():
            detalle = _frase_metricas_relevantes_pais(fila, indicadores_validos, 2)
            if detalle:
                oportunidades_txt.append(f"{fila['País']}: {detalle}")
            else:
                oportunidades_txt.append(str(fila["País"]))

        texto.append(
            "Por país, los mejores comportamientos relativos se observan en "
            + " | ".join(mejores_txt)
            + ". "
        )
        texto.append(
            "Las áreas de oportunidad se concentran en "
            + " | ".join(oportunidades_txt)
            + ". La prioridad es revisar las estructuras con mayor presión en atraso, faltas o disminución de clientes al corriente."
        )
    else:
        fila = comparativo_pais.iloc[0]
        pais = str(fila["País"])
        detalle = _frase_metricas_relevantes_pais(fila, indicadores_validos, 5)
        texto.append(
            f"Para {pais}, la lectura principal es: {detalle}. "
            "Los puntos favorables deben sostenerse en la siguiente semana, mientras que las variables que presionan el resultado requieren seguimiento por estructura para ubicar rutas, sucursales o zonas específicas."
        )

    # Recomendación ejecutiva final con foco en indicadores de oportunidad.
    indicadores_oportunidad = []
    for ind in indicadores_validos:
        direccion = _direccion_favorable_indicador(ind)
        total_var = resumen.loc[resumen["Indicador"] == ind, "Variación vs sem ant"]
        if not total_var.empty and float(total_var.iloc[0]) * direccion < 0:
            indicadores_oportunidad.append(ind)

    if indicadores_oportunidad:
        texto.append(
            "Recomendación ejecutiva: priorizar acciones sobre "
            + ", ".join(indicadores_oportunidad[:4])
            + ", revisando primero los países o estructuras donde la variación fue desfavorable y validando si el comportamiento se explica por concentración de cartera, atrasos o menor recuperación semanal."
        )
    else:
        texto.append(
            "Recomendación ejecutiva: mantener el seguimiento semanal para confirmar que los avances se sostengan y evitar que el crecimiento de cartera venga acompañado de mayor atraso o faltas."
        )

    return "\n\n".join(texto)



def calcular_resumen_cobranza_para_modal(
    df_cobranza_base: pd.DataFrame | None,
    df_cartera_base: pd.DataFrame,
    filtros_aplicados: dict,
    semana_referencia: int | None = None,
):
    """
    Construye un resumen de Cobranza para la ventana emergente.
    Usa los mismos filtros superiores del tablero y compara la última semana
    disponible de Cobranza contra la semana anterior disponible.
    """
    if df_cobranza_base is None or df_cobranza_base.empty:
        return pd.DataFrame(), None, None

    try:
        df_cob_filtrada = aplicar_filtros_cobranza_desde_cartera(
            df_cobranza_base=df_cobranza_base,
            df_cartera_base=df_cartera_base,
            filtros=filtros_aplicados,
        )
    except Exception:
        df_cob_filtrada = df_cobranza_base.copy()

    if df_cob_filtrada is None or df_cob_filtrada.empty:
        return pd.DataFrame(), None, None

    df_cob, col_cuota, col_pago, col_cump, col_mejor, col_peor = preparar_cobranza(df_cob_filtrada)

    if "Semana del año" not in df_cob.columns or col_cuota is None or col_pago is None:
        return pd.DataFrame(), None, None

    df_cob = df_cob.copy()
    df_cob["Semana del año"] = pd.to_numeric(df_cob["Semana del año"], errors="coerce")
    df_cob = df_cob.dropna(subset=["Semana del año"])

    if df_cob.empty:
        return pd.DataFrame(), None, None

    semanas_disponibles = sorted(df_cob["Semana del año"].astype(int).unique().tolist())

    if semana_referencia is not None:
        semanas_hasta_ref = [s for s in semanas_disponibles if s <= int(semana_referencia)]
        semana_actual_cob = semanas_hasta_ref[-1] if semanas_hasta_ref else semanas_disponibles[-1]
    else:
        semana_actual_cob = semanas_disponibles[-1]

    semanas_previas = [s for s in semanas_disponibles if s < semana_actual_cob]
    semana_anterior_cob = semanas_previas[-1] if semanas_previas else None

    def _agregar_semana(semana):
        df_sem = df_cob[df_cob["Semana del año"].astype(int) == int(semana)].copy()
        cuota = pd.to_numeric(df_sem[col_cuota], errors="coerce").fillna(0).sum()
        pago = pd.to_numeric(df_sem[col_pago], errors="coerce").fillna(0).sum()
        cumplimiento = np.nan if cuota == 0 else pago / cuota
        datos = {
            col_cuota: cuota,
            col_pago: pago,
            "% de Cumplimiento": cumplimiento,
        }
        if col_mejor and col_mejor in df_sem.columns:
            datos[col_mejor] = pd.to_numeric(df_sem[col_mejor], errors="coerce").fillna(0).max()
        if col_peor and col_peor in df_sem.columns:
            datos[col_peor] = pd.to_numeric(df_sem[col_peor], errors="coerce").fillna(0).min()
        return datos

    actual = _agregar_semana(semana_actual_cob)
    anterior = _agregar_semana(semana_anterior_cob) if semana_anterior_cob is not None else {}

    indicadores_cobranza = [col_cuota, col_pago, "% de Cumplimiento"]
    if col_mejor and col_mejor in actual:
        indicadores_cobranza.append(col_mejor)
    if col_peor and col_peor in actual:
        indicadores_cobranza.append(col_peor)

    filas = []
    for indicador in indicadores_cobranza:
        val_actual = actual.get(indicador, np.nan)
        val_anterior = anterior.get(indicador, np.nan)
        variacion = np.nan if pd.isna(val_anterior) else val_actual - val_anterior

        if indicador == "% de Cumplimiento":
            pct_var = variacion
        else:
            pct_var = np.nan if pd.isna(val_anterior) or val_anterior == 0 else variacion / val_anterior

        filas.append({
            "Indicador": indicador,
            f"Dato sem {semana_actual_cob}": val_actual,
            "Variación vs sem ant": variacion,
            "% Var": pct_var,
        })

    return pd.DataFrame(filas), semana_actual_cob, semana_anterior_cob


def aplicar_formato_tabla_resumen_mixto(df_tabla: pd.DataFrame) -> pd.DataFrame:
    """
    Formato para tablas de resumen que mezclan montos y porcentajes.

    Importante:
    En versiones recientes de pandas, una columna numérica ya no permite
    asignar directamente textos como "236,843,298" o "82.55%".
    Por eso primero convertimos la tabla a object/string-safe antes de
    reemplazar valores numéricos por valores formateados.
    """
    if df_tabla is None or df_tabla.empty:
        return pd.DataFrame()

    df_fmt = df_tabla.copy().astype(object)

    for idx, fila in df_tabla.iterrows():
        indicador = str(fila.get("Indicador", ""))
        indicador_norm = normalizar_texto_tc(indicador).lower()
        es_pct = indicador.strip().startswith("%") or "cumplimiento" in indicador_norm

        for col in df_fmt.columns:
            if col == "Indicador":
                df_fmt.at[idx, col] = indicador
                continue

            valor = fila.get(col)

            try:
                if es_pct:
                    if "Variación" in str(col) or "% Var" in str(col):
                        df_fmt.at[idx, col] = formato_pct(valor, 2, True)
                    else:
                        df_fmt.at[idx, col] = formato_pct(valor, 2, False)
                elif "% Var" in str(col):
                    df_fmt.at[idx, col] = formato_pct(valor, 1, True)
                elif "Variación" in str(col) or str(col).startswith("Var "):
                    df_fmt.at[idx, col] = formato_variacion(valor)
                else:
                    df_fmt.at[idx, col] = formato_numero(valor)
            except Exception:
                df_fmt.at[idx, col] = "" if pd.isna(valor) else str(valor)

    return df_fmt

def abrir_modal_resumen_pais(
    resumen: pd.DataFrame,
    semana_actual: int,
    semana_anterior,
    comentario_resumen: str,
    modo_moneda: str,
    filtros_aplicados: dict,
    resumen_cobranza: pd.DataFrame | None = None,
    semana_actual_cobranza: int | None = None,
    semana_anterior_cobranza: int | None = None,
):
    """
    Abre un resumen ejecutivo en ventana emergente.
    La ventana se agranda por CSS para cubrir aproximadamente el 80% de la pantalla.
    """
    @st.dialog("Resumen semanal de todo el país", width="large")
    def _dialogo_resumen():
        filtros_visibles = []
        for col, val in filtros_aplicados.items():
            if val:
                valores = ", ".join([str(x) for x in val])
                filtros_visibles.append(f"{col}: {valores}")

        filtros_texto = " | ".join(filtros_visibles) if filtros_visibles else "Todos los países / todas las marcas"
        semana_anterior_txt = semana_anterior if semana_anterior is not None else "sin semana anterior"

        st.markdown(
            f"""
            <span class="modal-resumen-meta">Semana actual: {semana_actual}</span>
            <span class="modal-resumen-meta">Comparativo: {semana_anterior_txt}</span>
            <span class="modal-resumen-meta">Moneda: {etiqueta_moneda(modo_moneda)}</span>
            """,
            unsafe_allow_html=True
        )

        st.caption(f"Filtros aplicados: {filtros_texto}")

        st.markdown("### Lectura ejecutiva generada por IA")
        comentario_seguro = html.escape(str(comentario_resumen)).replace("\n", "<br><br>")
        st.markdown(
            f'<div class="modal-resumen-card">{comentario_seguro}</div>',
            unsafe_allow_html=True
        )

        if resumen is not None and not resumen.empty:
            metricas_modal = [
                c for c in [
                    "Clientes Totales",
                    "Clientes al corriente",
                    "Cartera Total",
                    "Saldo en atraso",
                ]
                if c in resumen["Indicador"].astype(str).tolist()
            ]

            if metricas_modal:
                cols_modal = st.columns(len(metricas_modal))
                for idx, indicador in enumerate(metricas_modal):
                    fila = resumen[resumen["Indicador"] == indicador].iloc[0]
                    with cols_modal[idx]:
                        tarjeta_kpi(
                            label=indicador,
                            valor=fila.get(f"Dato sem {semana_actual}", 0),
                            variacion=fila.get("Variación vs sem ant", np.nan)
                            if semana_anterior is not None else None
                        )

            st.markdown("### Detalle del resumen")
            st.dataframe(
                aplicar_formato_tabla(resumen),
                use_container_width=True,
                hide_index=True,
                height=380
            )

        if resumen_cobranza is not None and not resumen_cobranza.empty:
            st.markdown("### Resumen de cobranza")

            metricas_cob_modal = [
                c for c in ["Cuota Total Cobranza", "Recuperación semana", "% de Cumplimiento"]
                if c in resumen_cobranza["Indicador"].astype(str).tolist()
            ]

            if metricas_cob_modal and semana_actual_cobranza is not None:
                col_cob_modal = st.columns(len(metricas_cob_modal))
                for idx, indicador in enumerate(metricas_cob_modal):
                    fila_cob = resumen_cobranza[resumen_cobranza["Indicador"] == indicador].iloc[0]
                    valor_cob = fila_cob.get(f"Dato sem {semana_actual_cobranza}", np.nan)
                    var_cob = fila_cob.get("Variación vs sem ant", np.nan)

                    with col_cob_modal[idx]:
                        if indicador.strip().startswith("%") or "cumplimiento" in normalizar_texto_tc(indicador).lower():
                            valor_txt = formato_pct(valor_cob, 2, False)
                            var_txt = formato_pct(var_cob, 2, True) if pd.notna(var_cob) else ""
                            st.markdown(
                                f"""
                                <div class="kpi-card">
                                    <div class="kpi-label">{indicador}</div>
                                    <div class="kpi-value">{valor_txt}</div>
                                    <div class="kpi-delta-neutral">{var_txt}</div>
                                </div>
                                """,
                                unsafe_allow_html=True
                            )
                        else:
                            tarjeta_kpi(
                                label=indicador,
                                valor=valor_cob,
                                variacion=var_cob if semana_anterior_cobranza is not None else None
                            )

            st.dataframe(
                aplicar_formato_tabla_resumen_mixto(resumen_cobranza),
                use_container_width=True,
                hide_index=True,
                height=230
            )
        else:
            st.info("No se encontró información de Cobranza para incluirla en este resumen.")



    _dialogo_resumen()


def generar_comentario_evolucion(evol: pd.DataFrame, indicador: str):
    if evol is None or evol.empty:
        return "No hay información suficiente para comentar la evolución."

    evol_tmp = evol.dropna(subset=[indicador]).copy()

    if evol_tmp.empty:
        return "No hay valores disponibles para comentar la evolución."

    ultima = evol_tmp.iloc[-1]
    semana_ultima = int(ultima["Semana del año"])
    valor_ultimo = ultima[indicador]

    comentario = (
        f"La evolución semanal de {indicador} cierra en la semana {semana_ultima} "
        f"con {_fmt_comentario(valor_ultimo)}."
    )

    if len(evol_tmp) >= 2:
        anterior = evol_tmp.iloc[-2]
        variacion = valor_ultimo - anterior[indicador]

        if variacion > 0:
            comentario += f" Frente a la semana previa presenta un incremento de {_fmt_comentario(variacion)}."
        elif variacion < 0:
            comentario += f" Frente a la semana previa presenta una disminución de {_fmt_comentario(variacion)}."
        else:
            comentario += " Frente a la semana previa se mantiene sin variación."

    semana_max = int(evol_tmp.loc[evol_tmp[indicador].idxmax(), "Semana del año"])
    valor_max = evol_tmp[indicador].max()

    comentario += f" El punto más alto del periodo se observa en la semana {semana_max}, con {_fmt_comentario(valor_max)}."

    return comentario


def generar_comentario_pie(pie: pd.DataFrame):
    if pie is None or pie.empty:
        return "No hay información suficiente para comentar la distribución."

    total = pie["Coordinadoras"].sum()

    if total == 0:
        return "No hay coordinadoras disponibles para comentar la distribución."

    pie_tmp = pie.copy()
    pie_tmp["Participación"] = pie_tmp["Coordinadoras"] / total
    principal = pie_tmp.sort_values("Coordinadoras", ascending=False).iloc[0]

    comentario = (
        f"La distribución de coordinadoras está concentrada principalmente en "
        f"{principal['Tipo Coordinadora']}, con {_fmt_comentario(principal['Coordinadoras'])} "
        f"coordinadoras, equivalentes al {principal['Participación']:.1%} del total."
    )

    impro = pie_tmp[
        pie_tmp["Tipo Coordinadora"]
        .astype(str)
        .str.contains("Improductiva", case=False, na=False)
    ]

    if not impro.empty:
        part_impro = impro["Coordinadoras"].sum() / total
        comentario += f" La participación de coordinadoras improductivas es de {part_impro:.1%}."

    return comentario


def generar_comentario_matriz(matriz, movimientos, semana_origen, semana_destino):
    if matriz is None or matriz.empty or movimientos is None or movimientos.empty:
        return "No hay información suficiente para comentar la matriz de desplazamiento."

    resumen_mov = calcular_resumen_movimientos(movimientos)

    comentario = (
        f"Entre la semana {semana_origen} y la semana {semana_destino}, "
        f"se registran {_fmt_comentario(resumen_mov['Movimientos totales'])} movimientos totales de coordinadoras. "
        f"Destacan {_fmt_comentario(resumen_mov['Aumento de productivas'])} aumentos hacia Productiva "
        f"y {_fmt_comentario(resumen_mov['Aumento de desarrollo'])} aumentos hacia En Desarrollo."
    )

    if resumen_mov["Pérdida de Productivas"] > 0:
        comentario += f" Como foco de atención, se observan {_fmt_comentario(resumen_mov['Pérdida de Productivas'])} pérdidas de Productivas."

    if resumen_mov["Aumento de improductivas"] > 0:
        comentario += f" También se identifican {_fmt_comentario(resumen_mov['Aumento de improductivas'])} aumentos hacia Improductiva."

    return comentario


def generar_comentario_top_bottom(tabla_top_bottom, tipo_top_bottom, nivel_top_bottom, semana_actual):
    if tabla_top_bottom is None or tabla_top_bottom.empty:
        return "No hay información suficiente para comentar el Top / Bottom."

    primera = tabla_top_bottom.iloc[0]

    comentario = (
        f"En la semana {semana_actual}, el {tipo_top_bottom} por {nivel_top_bottom} "
        f"muestra como principal registro a {primera['Estructura']} en la variable "
        f"{primera['Variable']}, con {_fmt_comentario(primera['Valor'])}."
    )

    if tipo_top_bottom == "Top":
        comentario += " Esta vista permite identificar las estructuras con mayor aportación."
    else:
        comentario += " Esta vista permite detectar las estructuras con menor desempeño o menor volumen."

    return comentario


def generar_comentario_detalle(detalle, nivel, semana_actual):
    if detalle is None or detalle.empty:
        return "No hay información suficiente para comentar el detalle agrupado."

    columnas_metricas = [
        c for c in detalle.columns
        if c != nivel and not c.startswith("Var ")
    ]

    if not columnas_metricas:
        return "No hay indicadores numéricos suficientes para comentar el detalle."

    indicador_principal = columnas_metricas[0]
    top = detalle.sort_values(indicador_principal, ascending=False).iloc[0]

    comentario = (
        f"En el detalle por {nivel}, la estructura con mayor valor en {indicador_principal} "
        f"durante la semana {semana_actual} es {top[nivel]}, con {_fmt_comentario(top[indicador_principal])}."
    )

    col_var = f"Var {indicador_principal}"

    if col_var in detalle.columns:
        mejor_var = detalle.sort_values(col_var, ascending=False).iloc[0]
        peor_var = detalle.sort_values(col_var, ascending=True).iloc[0]

        comentario += (
            f" El mayor crecimiento lo presenta {mejor_var[nivel]}, "
            f"con {_fmt_comentario(mejor_var[col_var])}."
        )

        if peor_var[col_var] < 0:
            comentario += (
                f" La mayor disminución se observa en {peor_var[nivel]}, "
                f"con {_fmt_comentario(peor_var[col_var])}."
            )

    return comentario


def comentario_cobranza_cumplimiento(evol, col_cump):
    if evol is None or evol.empty:
        return "No hay información suficiente para comentar el cumplimiento."

    df_tmp = evol.dropna(subset=[col_cump]).copy()

    if df_tmp.empty:
        return "No hay datos válidos de cumplimiento."

    actual = df_tmp.iloc[-1]
    mejor = df_tmp.loc[df_tmp[col_cump].idxmax()]
    peor = df_tmp.loc[df_tmp[col_cump].idxmin()]

    comentario = (
        f"El cumplimiento cierra en {actual['Etiqueta semana']} con {actual[col_cump]:.2%}. "
        f"La mejor semana del periodo fue {mejor['Etiqueta semana']}, con {mejor[col_cump]:.2%}, "
        f"mientras que la menor lectura se presentó en {peor['Etiqueta semana']}, con {peor[col_cump]:.2%}."
    )

    if len(df_tmp) >= 2:
        ant = df_tmp.iloc[-2]
        var = actual[col_cump] - ant[col_cump]

        if var > 0:
            comentario += f" Frente a la semana previa, el cumplimiento mejora {var:.2%}."
        elif var < 0:
            comentario += f" Frente a la semana previa, el cumplimiento disminuye {abs(var):.2%}."
        else:
            comentario += " Frente a la semana previa, el cumplimiento se mantiene estable."

    return comentario


def comentario_cobranza_cuota_pago(evol, col_cuota, col_pago, col_cump):
    if evol is None or evol.empty:
        return "No hay información suficiente para comentar la cobranza."

    df_tmp = evol.copy().sort_values(["Año", "Semana del año"] if "Año" in evol.columns else ["Semana del año"])
    actual = df_tmp.iloc[-1]
    brecha = actual[col_cuota] - actual[col_pago]

    comentario = (
        f"En {actual['Etiqueta semana']}, la cuota total es de {_fmt_comentario(actual[col_cuota])} "
        f"y el pago total alcanza {_fmt_comentario(actual[col_pago])}, "
        f"equivalente a un cumplimiento de {actual[col_cump]:.2%}."
    )

    if brecha > 0:
        comentario += f" La brecha pendiente contra la cuota es de {_fmt_comentario(brecha)}."
    else:
        comentario += f" El pago supera la cuota por {_fmt_comentario(abs(brecha))}."

    if len(df_tmp) >= 2:
        ant = df_tmp.iloc[-2]
        var_pago = actual[col_pago] - ant[col_pago]

        if var_pago > 0:
            comentario += f" Contra la semana anterior, el pago total aumenta {_fmt_comentario(var_pago)}."
        elif var_pago < 0:
            comentario += f" Contra la semana anterior, el pago total disminuye {_fmt_comentario(var_pago)}."
        else:
            comentario += " Contra la semana anterior, el pago total no muestra variación."

    return comentario


def comentario_tabla_cobranza(tabla, col_cuota, col_pago, col_cump):
    if tabla is None or tabla.empty:
        return "No hay información suficiente para comentar la tabla de cobranza."

    df_tmp = tabla.copy()
    ultima = df_tmp.iloc[-1]
    semana = ultima.get("Semana", "la última semana")

    comentario = (
        f"La tabla muestra las últimas 5 semanas disponibles de cobranza, "
        f"incluyendo el dato semanal y la variación contra la semana anterior para cada variable. "
        f"En {semana}, la cuota total es de {_fmt_comentario(ultima[col_cuota])}, "
        f"la recuperación semana es de {_fmt_comentario(ultima[col_pago])} "
        f"y el cumplimiento alcanza {ultima[col_cump]:.2%}."
    )

    col_var_pago = f"Var {col_pago}"
    col_var_cump = f"Var {col_cump}"

    if col_var_pago in df_tmp.columns and pd.notna(ultima[col_var_pago]):
        var_pago = ultima[col_var_pago]
        if var_pago > 0:
            comentario += f" Frente a la semana previa, la recuperación aumenta {_fmt_comentario(var_pago)}."
        elif var_pago < 0:
            comentario += f" Frente a la semana previa, la recuperación disminuye {_fmt_comentario(var_pago)}."
        else:
            comentario += " Frente a la semana previa, la recuperación se mantiene sin variación."

    if col_var_cump in df_tmp.columns and pd.notna(ultima[col_var_cump]):
        var_cump = ultima[col_var_cump]
        if var_cump > 0:
            comentario += f" El cumplimiento mejora {var_cump:.2%}."
        elif var_cump < 0:
            comentario += f" El cumplimiento disminuye {abs(var_cump):.2%}."
        else:
            comentario += " El cumplimiento se mantiene estable."

    return comentario


# ============================================================
# CARGA DE DATOS
# ============================================================
# Encabezado superior eliminado para que no se duplique con el título central.

# La sección de archivo queda oculta para que el tablero entre directo al análisis.
# Si después necesitas volver a verla, cambia MOSTRAR_SECCION_ARCHIVO = True.
if MOSTRAR_SECCION_ARCHIVO:
    st.markdown('<div class="top-filter-card"><div class="top-filter-title">Archivo</div>', unsafe_allow_html=True)
    col_ruta_archivo, col_upload_archivo = st.columns([2.2, 1])

    with col_ruta_archivo:
        ruta_local = st.text_input("Ruta local del archivo", value=RUTA_DEFAULT)

    with col_upload_archivo:
        archivo_subido = st.file_uploader(
            "O sube aquí tu CSV/Excel",
            type=["csv", "xlsx", "xlsm", "xlsb", "xls"]
        )

    st.markdown('</div>', unsafe_allow_html=True)
else:
    ruta_local = RUTA_DEFAULT
    archivo_subido = None

try:
    df, df_cobranza = cargar_archivo(ruta_local, archivo_subido)
except Exception as e:
    st.error(str(e))
    st.stop()


# ============================================================
# VALIDACIONES CARTERA
# ============================================================
columnas_faltantes = [
    c for c in ["Semana del año", "Tipo Coordinadora"]
    if c not in df.columns
]

if columnas_faltantes:
    st.error(f"Faltan columnas obligatorias en Cartera: {columnas_faltantes}")
    st.stop()

niveles_disponibles = [c for c in NIVELES_ESTRUCTURA if c in df.columns]
indicadores_disponibles = [c for c in INDICADORES_BASE if c in df.columns]
columna_cobranza_cartera = detectar_columna_cobranza_cartera(df)

if not niveles_disponibles:
    st.error("No encontré columnas de estructura para agrupar.")
    st.stop()

if not indicadores_disponibles:
    st.error("No encontré columnas numéricas de indicadores.")
    st.stop()


# ============================================================
# PANTALLA INICIAL POR UNIDAD DE NEGOCIO
# ============================================================
if "Unidad de Negocio" in df.columns:
    unidades_negocio = sorted(df["Unidad de Negocio"].dropna().astype(str).unique())
else:
    unidades_negocio = []

if unidades_negocio:
    unidad_guardada = st.session_state.get("unidad_negocio_app", None)

    if unidad_guardada not in unidades_negocio:
        st.session_state["unidad_negocio_app"] = None
        unidad_guardada = None

    if unidad_guardada is None:
        st.markdown(
            """
            <div class="landing-wrap">
                <div class="landing-title">Indicadores de Productividad y Cobranza</div>
                <div class="landing-subtitle">Semanal</div>
            </div>
            """,
            unsafe_allow_html=True
        )

        logos_unidad = {
            "PRESICO": "Logo.jpg",
            "PRESICO LATAM": "Presico sin fondo LATAM.jpg",
        }

        cols_unidades = st.columns(min(len(unidades_negocio), 4))

        for idx, unidad in enumerate(unidades_negocio):
            unidad_texto = str(unidad).strip()
            unidad_key = normalizar_texto_tc(unidad_texto)
            nombre_logo = logos_unidad.get(unidad_key)

            if nombre_logo:
                logo_html = imagen_logo_html(nombre_logo)
            else:
                logo_html = '<div class="unidad-logo-placeholder">🏢</div>'

            with cols_unidades[idx % len(cols_unidades)]:
                st.markdown(
                    f"""
                    <div class="unidad-card">
                        {logo_html}
                        <div class="unidad-name">{html.escape(unidad_texto)}</div>
                        <div class="unidad-help">Entrar al tablero</div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                if st.button(f"Entrar a {unidad_texto}", key=f"btn_unidad_{idx}", use_container_width=True):
                    st.session_state["unidad_negocio_app"] = unidad_texto
                    st.rerun()

        st.stop()

    unidad_negocio_seleccionada = st.session_state.get("unidad_negocio_app", None)
else:
    st.session_state["unidad_negocio_app"] = None
    unidad_negocio_seleccionada = None


# ============================================================
# BARRA SUPERIOR
# ============================================================
st.markdown('<div class="top-filter-card"><div class="top-filter-title">Filtros</div>', unsafe_allow_html=True)

filtros = {}

if unidad_negocio_seleccionada is not None:
    filtros["Unidad de Negocio"] = [unidad_negocio_seleccionada]

base_para_filtros = filtrar_por_diccionario(df, filtros)

col_unidad_actual, col_modulo, col_moneda, col_marca, col_pais, col_resumen_pais, col_cambiar = st.columns([1.10, 0.90, 1.10, 0.95, 0.95, 1.10, 0.85])

with col_unidad_actual:
    if unidad_negocio_seleccionada is not None:
        st.markdown(
            f'<div class="unidad-seleccionada-pill">{html.escape(str(unidad_negocio_seleccionada))}</div>',
            unsafe_allow_html=True
        )
    else:
        st.markdown('<div class="unidad-seleccionada-pill">Todas las unidades</div>', unsafe_allow_html=True)

with col_modulo:
    modulo_seleccionado = st.selectbox(
        "Vista",
        options=["Cartera", "Cobranza"],
        index=0,
        key="modulo_superior"
    )

with col_moneda:
    modo_moneda = st.radio(
        "Moneda",
        options=["Moneda local", "Pesos mexicanos"],
        index=0,
        horizontal=True,
        key="modo_moneda_superior"
    )

# Marca y País quedan como únicas opciones de filtro dentro del tablero.
for col_filtro, col_streamlit in [("Marca", col_marca), ("País", col_pais)]:
    if col_filtro in df.columns:
        df_opciones = filtrar_por_diccionario(base_para_filtros, filtros, excluir_col=col_filtro)
        valores = sorted(df_opciones[col_filtro].dropna().astype(str).unique())
        opciones = ["Todos"] + valores
        with col_streamlit:
            seleccion = st.selectbox(
                col_filtro,
                options=opciones,
                index=0,
                key=f"filtro_superior_{col_filtro}"
            )
        if seleccion != "Todos":
            filtros[col_filtro] = [seleccion]
    else:
        with col_streamlit:
            st.caption(f"Sin columna {col_filtro}")

# IMPORTANTE:
# El resumen se abre solo en el clic de este botón.
# No se deja guardado como estado persistente, porque si el usuario lo cierra
# con la X de Streamlit, el estado no se limpia y el resumen se vuelve a abrir
# al presionar cualquier otro botón del tablero.
abrir_resumen_pais_click = False

with col_resumen_pais:
    abrir_resumen_pais_click = st.button(
        "Resumen semana país",
        key="btn_abrir_resumen_pais",
        use_container_width=True
    )


with col_cambiar:
    if unidad_negocio_seleccionada is not None:
        if st.button("Cambiar unidad", key="btn_cambiar_unidad", use_container_width=True):
            st.session_state.pop("unidad_negocio_app", None)
            st.rerun()

st.caption(
    "La evolución semanal muestra todo el histórico filtrado. La distribución por coordinadora usa siempre la última semana disponible. "
    "El selector de moneda solo convierte variables monetarias; clientes y coordinadoras permanecen igual."
)
st.markdown('</div>', unsafe_allow_html=True)

# Aplica conversión de moneda después de construir los filtros, para conservar las opciones originales.
df = aplicar_tipo_cambio_mxn(df, modo_moneda)
if df_cobranza is not None:
    df_cobranza = aplicar_tipo_cambio_mxn(df_cobranza, modo_moneda)

# Siempre se usan todos los indicadores disponibles; ya no hay selector múltiple.
indicadores_sel = indicadores_disponibles.copy()
semanas = sorted([int(s) for s in df["Semana del año"].dropna().unique()])
semanas_sel = semanas

# Nivel fijo para tablas de detalle; ya no se muestra en la barra superior.
niveles_detalle_preferidos = ["Marca", "País", "Subdireccion", "Zona", "Sucursal", "Ruta", "Unidad de Negocio"]
nivel = next((c for c in niveles_detalle_preferidos if c in niveles_disponibles), niveles_disponibles[0])

if not semanas:
    st.warning("No hay semanas disponibles en la base.")
    st.stop()

if not indicadores_sel:
    st.warning("No hay indicadores disponibles en la base.")
    st.stop()

# ============================================================
# FILTROS CARTERA
# ============================================================
# Se filtra por Unidad de Negocio / Marca / País, pero se conserva TODO el histórico de semanas.
df_filtrado_original = aplicar_filtros_base(
    df_base=df,
    semanas_sel=[],
    filtros=filtros
)

if df_filtrado_original.empty:
    st.warning("No hay datos de Cartera con los filtros seleccionados.")
    st.stop()

filas_antes_consolidar = len(df_filtrado_original)

df_filtrado = consolidar_grano_correcto(
    df_base=df_filtrado_original,
    indicadores=indicadores_sel
)

filas_despues_consolidar = len(df_filtrado)

if df_filtrado.empty:
    st.warning("No hay datos después de consolidar el grano correcto.")
    st.stop()

semanas_historial_filtrado = sorted([int(s) for s in df_filtrado["Semana del año"].dropna().unique()])
semana_actual = semanas_historial_filtrado[-1]
semana_ultima_historial = semanas_historial_filtrado[-1]

# Resumen general disponible para el botón de la barra superior.
# Usa los filtros generales activos y la última semana disponible del país / unidad seleccionada.
resumen_general_pais, semana_anterior_general_pais = calcular_resumen_actual_vs_anterior(
    df_filtrado=df_filtrado,
    indicadores=indicadores_sel,
    semana_actual=semana_actual
)
comentario_general_pais = generar_resumen_ia_paises(
    df_base=df_filtrado,
    resumen=resumen_general_pais,
    indicadores=indicadores_sel,
    semana_actual=semana_actual,
    semana_anterior=semana_anterior_general_pais,
    modo_moneda=modo_moneda,
    filtros_aplicados=filtros
)

if abrir_resumen_pais_click:
    # El resumen se abre únicamente cuando se oprime el botón correspondiente.
    # Esto evita que se vuelva a abrir al usar filtros, descargar archivos,
    # cambiar unidad después de haberlo cerrado con la X.
    st.session_state["modal_activo"] = None

    resumen_cobranza_modal, semana_actual_cob_modal, semana_anterior_cob_modal = calcular_resumen_cobranza_para_modal(
        df_cobranza_base=df_cobranza,
        df_cartera_base=df,
        filtros_aplicados=filtros,
        semana_referencia=semana_actual,
    )

    abrir_modal_resumen_pais(
        resumen=resumen_general_pais,
        semana_actual=semana_actual,
        semana_anterior=semana_anterior_general_pais,
        comentario_resumen=comentario_general_pais,
        modo_moneda=modo_moneda,
        filtros_aplicados=filtros,
        resumen_cobranza=resumen_cobranza_modal,
        semana_actual_cobranza=semana_actual_cob_modal,
        semana_anterior_cobranza=semana_anterior_cob_modal,
    )

# ============================================================
# CONTROL DE DATOS
# ============================================================
duplicados_eliminados = df.attrs.get("duplicados_exactos_eliminados", 0)

# Oculto por solicitud: no se muestra el bloque de control de datos cargados.
# Si necesitas revisarlo, cambia MOSTRAR_CONTROL_DATOS = True.
if MOSTRAR_CONTROL_DATOS:
    with st.expander("Control de datos cargados", expanded=False):
        st.write(f"Filas originales Cartera: {df.attrs.get('filas_antes_limpieza', len(df)):,}")
        st.write(f"Filas usadas después de quitar duplicados exactos: {df.attrs.get('filas_despues_limpieza', len(df)):,}")
        st.write(f"Duplicados exactos eliminados: {duplicados_eliminados:,}")
        st.write(f"Filas filtradas antes de consolidar (histórico completo): {filas_antes_consolidar:,}")
        st.write(f"Filas después de consolidar grano correcto: {filas_despues_consolidar:,}")

        if df_cobranza is not None:
            st.success(f"Hoja Cobranza detectada correctamente: {len(df_cobranza):,} filas.")
            st.write("Columnas Cobranza:", list(df_cobranza.columns))
        else:
            st.info("No se detectó hoja Cobranza. Recuerda que CSV no puede tener segunda hoja.")



# ============================================================
# VISTA SELECCIONADA
# ============================================================
if modulo_seleccionado == "Cartera":
    # ============================================================
    # RESUMEN CARTERA
    # ============================================================
    # Oculto por solicitud: el resumen ejecutivo general ya no se muestra
    # directamente en la página. Solo aparece al presionar el botón
    # "Resumen semana país" de la barra superior.

    # ============================================================
    # GRÁFICAS CARTERA
    # ============================================================
    comentario_evolucion = ""
    comentario_pie = ""
    col1, col2 = st.columns([1.15, 0.85])

    with col1:
        st.subheader("Evolución semanal")

        metricas_evolucion = [
            c for c in [
                "Clientes Totales",
                "Clientes al corriente",
                "Faltas",
                "Nunca Abonados",
                "Cartera Total",
                "Saldo Cartera",
                "Saldo en atraso",
                "Saldo PP"
            ]
            if c in df_filtrado.columns
        ]

        if columna_cobranza_cartera and columna_cobranza_cartera in df_filtrado.columns:
            metricas_evolucion.append(columna_cobranza_cartera)

        if metricas_evolucion:
            col_grafica, col_menu = st.columns([4, 1])

            with col_menu:
                indicador_grafica = st.selectbox(
                    "Indicador",
                    options=metricas_evolucion,
                    index=metricas_evolucion.index("Cartera Total")
                    if "Cartera Total" in metricas_evolucion
                    else 0,
                    key="indicador_evolucion"
                )

            evol = (
                df_filtrado
                .groupby("Semana del año", dropna=False)[indicador_grafica]
                .sum()
                .reset_index()
                .sort_values("Semana del año")
            )

            evol["Variación vs anterior"] = evol[indicador_grafica].diff()

            comentario_evolucion = generar_comentario_evolucion(
                evol=evol,
                indicador=indicador_grafica
            )

            with col_grafica:
                fig_linea, config_linea = crear_grafica_evolucion_fija(
                    evol=evol,
                    indicador_grafica=indicador_grafica,
                    modo_moneda=modo_moneda,
                    altura=430
                )

                st.plotly_chart(
                    fig_linea,
                    use_container_width=True,
                    config=config_linea
                )

        else:
            st.info("No hay indicadores disponibles para la gráfica de evolución semanal.")


    with col2:
        st.subheader(f"Distribución por tipo de coordinadora | Última semana: {semana_ultima_historial}")

        if "coordinadora_id" in df_filtrado_original.columns:
            df_pie_base = crear_llave_coordinadora_marca(
                df_base=df_filtrado_original,
                columna_id="coordinadora_id"
            )

            df_sem_actual_unico = obtener_categoria_unica_por_semana(
                df_base=df_pie_base,
                semana=semana_ultima_historial,
                columna_id="_llave_coordinadora_marca",
                columna_categoria="Tipo Coordinadora"
            )

            pie = (
                df_sem_actual_unico
                .groupby("Tipo Coordinadora", dropna=False)["_llave_coordinadora_marca"]
                .nunique()
                .reset_index(name="Coordinadoras")
            )

        else:
            df_sem_actual = df_filtrado_original[df_filtrado_original["Semana del año"] == semana_ultima_historial].copy()

            pie = (
                df_sem_actual
                .groupby("Tipo Coordinadora", dropna=False)
                .size()
                .reset_index(name="Coordinadoras")
            )

        pie = pie.sort_values("Coordinadoras", ascending=False).copy()
        pie["Etiqueta"] = pie.apply(
            lambda r: f"{r['Tipo Coordinadora']}<br>{float(r['Coordinadoras']):,.0f}<br>{(float(r['Coordinadoras']) / max(float(pie['Coordinadoras'].sum()), 1)):.1%}",
            axis=1
        )

        colores_tipo_coordinadora = {
            "Productiva": "#ffa0a4",
            "En Desarrollo": "#7ec0ee",
            "Improductiva": "#0b70c9",
            "Secundaria": "#ff2d2d",
        }

        fig_pie = go.Figure(
            data=[
                go.Pie(
                    labels=pie["Tipo Coordinadora"],
                    values=pie["Coordinadoras"],
                    hole=0.42,
                    sort=False,
                    direction="clockwise",
                    text=pie["Etiqueta"],
                    texttemplate="%{text}",
                    textposition="outside",
                    automargin=True,
                    marker=dict(
                        colors=[
                            colores_tipo_coordinadora.get(str(tipo), None)
                            for tipo in pie["Tipo Coordinadora"]
                        ],
                        line=dict(color="white", width=2)
                    ),
                    hovertemplate=(
                        "<b>%{label}</b><br>"
                        "Coordinadoras: %{value:,.0f}<br>"
                        "Participación: %{percent:.1%}"
                        "<extra></extra>"
                    ),
                    insidetextorientation="radial"
                )
            ]
        )

        fig_pie.update_traces(
            textfont=dict(size=13, color="#082567", family="Arial"),
            pull=[0.02] * len(pie)
        )

        fig_pie.update_layout(
            height=500,
            legend_title=None,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.18,
                xanchor="center",
                x=0.5,
                font=dict(size=12, color="#082567")
            ),
            paper_bgcolor="rgba(255,255,255,0)",
            plot_bgcolor="rgba(255,255,255,0)",
            font=dict(color="#082567", size=13),
            margin=dict(t=30, b=95, l=95, r=95),
            uniformtext_minsize=11,
            uniformtext_mode="show"
        )

        st.plotly_chart(
            fig_pie,
            use_container_width=True,
            config={"displayModeBar": False, "responsive": True}
        )
        comentario_pie = generar_comentario_pie(pie)


    mostrar_boton_comentario("grafica_evolucion", comentario_evolucion)
    mostrar_boton_comentario("pie_coordinadoras", comentario_pie)


    # ============================================================
    # MATRIZ DE DESPLAZAMIENTO DE COORDINADORAS
    # ============================================================
    st.subheader("Matriz de desplazamiento de coordinadoras por categoría")

    if "coordinadora_id" not in df_filtrado_original.columns:
        st.warning(
            "No se puede generar la matriz porque no existe la columna 'coordinadora_id'. "
            "Esta columna es necesaria para identificar a la misma coordinadora entre semanas."
        )

    else:
        semanas_disponibles_matriz = sorted([
            int(s) for s in df_filtrado_original["Semana del año"].dropna().unique()
        ])

        if len(semanas_disponibles_matriz) < 2:
            st.info("Selecciona al menos dos semanas para construir la matriz de desplazamiento.")

        else:
            col_origen, col_destino = st.columns(2)

            with col_origen:
                semana_origen = st.selectbox(
                    "Semana anterior / origen",
                    options=semanas_disponibles_matriz[:-1],
                    index=max(0, len(semanas_disponibles_matriz) - 2),
                    key="semana_origen_matriz"
                )

            semanas_destino_validas = [
                s for s in semanas_disponibles_matriz
                if s > semana_origen
            ]

            with col_destino:
                semana_destino = st.selectbox(
                    "Semana actual / destino",
                    options=semanas_destino_validas,
                    index=len(semanas_destino_validas) - 1,
                    key="semana_destino_matriz"
                )

            movimientos, matriz_movimientos = matriz_desplazamiento_coordinadoras(
                df_filtrado=df_filtrado_original,
                semana_origen=semana_origen,
                semana_destino=semana_destino
            )

            if matriz_movimientos is None:
                st.warning("No se pudo construir la matriz. Revisa que existan las columnas necesarias.")

            elif matriz_movimientos.empty:
                st.info(
                    f"No hay coordinadoras entre la semana {semana_origen} "
                    f"y la semana {semana_destino}."
                )

            else:
                llave_matriz = "_llave_coordinadora_marca"
                total_origen = movimientos[movimientos["Semana anterior"] != "Nueva"][llave_matriz].nunique()
                total_destino = movimientos[movimientos["Semana actual"] != "Baja"][llave_matriz].nunique()
                total_nuevas = movimientos[movimientos["Semana anterior"] == "Nueva"][llave_matriz].nunique()
                total_bajas = movimientos[movimientos["Semana actual"] == "Baja"][llave_matriz].nunique()

                st.caption(
                    f"Lectura: las filas muestran la categoría en la semana {semana_origen}; "
                    f"las columnas muestran la categoría en la semana {semana_destino}. "
                    "Los valores verdes son mejoras de categoría, los rojos son retrocesos, "
                    "la diagonal muestra permanencia, la fila Nueva muestra altas y la columna Baja muestra salidas."
                )

                col_m1, col_m2, col_m3, col_m4 = st.columns(4)

                with col_m1:
                    st.metric(f"Coordinadoras sem {semana_origen}", f"{total_origen:,.0f}")

                with col_m2:
                    st.metric(f"Coordinadoras sem {semana_destino}", f"{total_destino:,.0f}")

                with col_m3:
                    st.metric("Nuevas", f"{total_nuevas:,.0f}")

                with col_m4:
                    st.metric("Bajas", f"{total_bajas:,.0f}")

                col_matriz_mov, col_resumen_mov = st.columns([2.2, 1])

                with col_matriz_mov:
                    st.dataframe(
                        estilo_matriz_desplazamiento(matriz_movimientos),
                        use_container_width=True
                    )

                with col_resumen_mov:
                    mostrar_cuadro_resumen_movimientos(movimientos)

                comentario_matriz = generar_comentario_matriz(
                    matriz=matriz_movimientos,
                    movimientos=movimientos,
                    semana_origen=semana_origen,
                    semana_destino=semana_destino
                )

                mostrar_boton_comentario("matriz_movimientos", comentario_matriz)

                with st.expander("Ver detalle de coordinadoras desplazadas", expanded=False):
                    detalle_movimientos = movimientos.copy()
                    detalle_movimientos = detalle_movimientos.rename(columns={
                        "_llave_coordinadora_marca": "Llave Coordinadora Marca",
                        "coordinadora_id origen": "Coordinadora ID origen",
                        "coordinadora_id destino": "Coordinadora ID destino",
                        "País origen": "País origen",
                        "País destino": "País destino",
                        "Marca origen": "Marca origen",
                        "Marca destino": "Marca destino",
                    })

                    st.dataframe(
                        detalle_movimientos,
                        use_container_width=True,
                        hide_index=True
                    )


    # ============================================================
    # TOP / BOTTOM POR VARIABLE
    # ============================================================
    st.subheader("Top / Bottom por variable")
    comentario_top_bottom = ""

    niveles_top_bottom = [
        c for c in [
            "País",
            "Subdireccion",
            "Zona",
            "Sucursal",
            "Ruta",
        ]
        if c in df_filtrado.columns
    ]

    variables_top_bottom_disponibles = [
        c for c in indicadores_disponibles
        if c in df_filtrado.columns and pd.api.types.is_numeric_dtype(df_filtrado[c])
    ]

    if columna_cobranza_cartera and columna_cobranza_cartera in df_filtrado.columns:
        if columna_cobranza_cartera not in variables_top_bottom_disponibles:
            variables_top_bottom_disponibles.append(columna_cobranza_cartera)

    if not niveles_top_bottom:
        st.info("No hay niveles disponibles desde País hasta Ruta para construir el Top / Bottom.")

    elif not variables_top_bottom_disponibles:
        st.info("No hay variables numéricas disponibles para construir el Top / Bottom.")

    else:
        st.markdown(
            """
            <div class="top-bottom-opciones-card">
                <div class="top-bottom-opciones-title">Opciones</div>
            </div>
            """,
            unsafe_allow_html=True
        )

        col_tipo_tb, col_nivel_tb, col_variable_tb, col_cantidad_tb = st.columns([1.0, 1.35, 1.7, 1.25], gap="large")

        with col_tipo_tb:
            tipo_top_bottom = st.radio(
                "Vista",
                options=["Top", "Bottom"],
                horizontal=True,
                key="tipo_top_bottom"
            )

        with col_nivel_tb:
            nivel_top_bottom = st.selectbox(
                "Estructura",
                options=niveles_top_bottom,
                index=0,
                key="nivel_top_bottom"
            )

        with col_variable_tb:
            variable_top_bottom = st.selectbox(
                "Variable",
                options=variables_top_bottom_disponibles,
                index=0,
                key="variable_top_bottom"
            )

        with col_cantidad_tb:
            cantidad_top_bottom = st.number_input(
                "Cantidad de registros por variable",
                min_value=1,
                max_value=50,
                value=10,
                step=1,
                key="cantidad_top_bottom"
            )

        variables_top_bottom = [variable_top_bottom]

        if not variable_top_bottom:
            st.info("Selecciona una variable para mostrar el Top / Bottom.")
        else:
            tabla_top_bottom = construir_top_bottom_por_variable(
                df_filtrado=df_filtrado,
                nivel_top_bottom=nivel_top_bottom,
                variables_top_bottom=variables_top_bottom,
                semana_actual=semana_actual,
                tipo_ranking=tipo_top_bottom,
                cantidad=int(cantidad_top_bottom)
            )

            if tabla_top_bottom.empty:
                st.info("No hay datos para mostrar con la selección actual.")
            else:
                st.caption(
                    f"Semana {semana_actual} | {tipo_top_bottom} "
                    f"por {nivel_top_bottom} para la variable seleccionada."
                )

                comentario_top_bottom = generar_comentario_top_bottom(
                    tabla_top_bottom=tabla_top_bottom,
                    tipo_top_bottom=tipo_top_bottom,
                    nivel_top_bottom=nivel_top_bottom,
                    semana_actual=semana_actual
                )

                st.dataframe(
                    aplicar_formato_top_bottom(tabla_top_bottom),
                    use_container_width=True,
                    hide_index=True
                )

                csv_top_bottom = tabla_top_bottom.to_csv(index=False).encode("utf-8-sig")

                st.download_button(
                    label="Descargar Top / Bottom",
                    data=csv_top_bottom,
                    file_name=f"top_bottom_{nivel_top_bottom}_semana_{semana_actual}.csv",
                    mime="text/csv"
                )

    mostrar_boton_comentario("top_bottom", comentario_top_bottom)

    # ============================================================
    # TABLA POR NIVEL
    # ============================================================
    st.subheader(f"Detalle agrupado por {nivel}")

    detalle = tabla_por_nivel(
        df_filtrado=df_filtrado,
        nivel=nivel,
        indicadores=indicadores_sel,
        semana_actual=semana_actual
    )

    comentario_detalle = generar_comentario_detalle(
        detalle=detalle,
        nivel=nivel,
        semana_actual=semana_actual
    )
    mostrar_boton_comentario("detalle_agrupado", comentario_detalle)

    st.dataframe(
        aplicar_formato_tabla(detalle),
        use_container_width=True,
        hide_index=True
    )


    # ============================================================
    # DESCARGA
    # ============================================================
    csv = detalle.to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        label="Descargar detalle agrupado",
        data=csv,
        file_name=f"detalle_{nivel}_semana_{semana_actual}.csv",
        mime="text/csv"
    )

else:
    # ============================================================
    # SECCIÓN COBRANZA
    # ============================================================
    st.subheader("Cobranza")

    if df_cobranza is None:
        st.info(
            "No se encontró una hoja llamada 'Cobranza'. "
            "Usa un archivo Excel con hojas 'Cartera' y 'Cobranza'."
        )

    else:
        df_cobranza_preparada, col_cuota, col_pago, col_cump, col_mejor, col_peor = preparar_cobranza(df_cobranza)

        if "Semana del año" not in df_cobranza_preparada.columns:
            st.warning("La hoja Cobranza no contiene 'Semana' o 'Semana del año'.")

        elif col_cuota is None or col_pago is None:
            st.warning(
                "No pude detectar las columnas de cuota y pago en Cobranza. "
                "La base debe traer 'Cuota Total Cobranza' y 'Recuperación semana'."
            )
            st.write("Columnas encontradas en Cobranza:", list(df_cobranza_preparada.columns))

        else:
            # Aplica SOLO filtros de estructura en Cobranza.
            # No se filtra por la semana de análisis para que las gráficas
            # muestren todo el histórico disponible, no solo las últimas semanas seleccionadas.
            df_cobranza_filtrada = aplicar_filtros_cobranza_desde_cartera(
                df_cobranza_base=df_cobranza_preparada,
                df_cartera_base=df,
                filtros=filtros
            )

            if df_cobranza_filtrada.empty:
                st.info("No hay datos de Cobranza con los filtros seleccionados.")

            else:
                # Base de Cobranza filtrada por Unidad / Marca / País.
                # Se conserva antes de aplicar el selector interno de nivel para que
                # el Top / Bottom pueda comparar todas las estructuras disponibles.
                df_cobranza_top_bottom_base = df_cobranza_filtrada.copy()

                niveles_cobranza_disponibles = [
                    c for c in NIVELES_ESTRUCTURA
                    if c in df_cobranza_filtrada.columns
                ]

                # En tu hoja Cobranza actual solo viene País, por eso el nivel disponible será Total o País.
                col_cob_menu, col_cob_info = st.columns([1.25, 3])

                with col_cob_menu:
                    nivel_cobranza = st.selectbox(
                        "Nivel de estructura para Cobranza",
                        options=["Total"] + niveles_cobranza_disponibles,
                        index=0,
                        key="nivel_cobranza"
                    )

                nivel_cobranza_real = None if nivel_cobranza == "Total" else nivel_cobranza

                if nivel_cobranza_real is not None:
                    opciones_nivel_cob = sorted(
                        df_cobranza_filtrada[nivel_cobranza_real].dropna().astype(str).unique()
                    )

                    with col_cob_menu:
                        estructura_cobranza = st.selectbox(
                            f"Selecciona {nivel_cobranza_real}",
                            options=opciones_nivel_cob,
                            index=0,
                            key="estructura_cobranza"
                        )

                    df_cobranza_filtrada = df_cobranza_filtrada[
                        df_cobranza_filtrada[nivel_cobranza_real].astype(str) == estructura_cobranza
                    ].copy()

                evol_cobranza = consolidar_cobranza(
                    df_cobranza=df_cobranza_filtrada,
                    col_cuota=col_cuota,
                    col_pago=col_pago,
                    col_cump=col_cump,
                    col_mejor=col_mejor,
                    col_peor=col_peor,
                    nivel=None
                )

                if evol_cobranza.empty:
                    st.info("No hay datos suficientes para mostrar Cobranza con la selección actual.")

                else:
                    col_mejor_final = col_mejor if col_mejor and col_mejor in evol_cobranza.columns else "Mejor semana"
                    col_peor_final = col_peor if col_peor and col_peor in evol_cobranza.columns else "Peor semana"

                    with col_cob_info:
                        st.caption(
                            f"Columnas detectadas: cuota = '{col_cuota}', pago = '{col_pago}', "
                            f"cumplimiento = '{col_cump}', mejor = '{col_mejor_final}', peor = '{col_peor_final}'. Moneda: {etiqueta_moneda(modo_moneda)}."
                        )

                    # ------------------------------
                    # Tabla base de Cobranza
                    # ------------------------------
                    st.markdown("**Tabla semanal de cobranza**")
                    comentario_base = comentario_cobranza_cuota_pago(
                        evol=evol_cobranza,
                        col_cuota=col_cuota,
                        col_pago=col_pago,
                        col_cump=col_cump
                    )
                    mostrar_boton_comentario("cobranza_tabla_base", comentario_base)

                    columnas_tabla_cob = [
                        c for c in [
                            "Año",
                            "Semana del año",
                            col_cuota,
                            col_pago,
                            col_cump,
                            col_mejor_final,
                            col_peor_final,
                        ]
                        if c in evol_cobranza.columns
                    ]

                    tabla_detalle_cob = evol_cobranza[columnas_tabla_cob].copy()
                    tabla_detalle_cob = tabla_detalle_cob.rename(columns={"Semana del año": "Semana"})

                    st.dataframe(
                        formato_tabla_detalle_cobranza(
                            tabla_detalle_cob,
                            col_cuota=col_cuota,
                            col_pago=col_pago,
                            col_cump=col_cump,
                            col_mejor=col_mejor_final,
                            col_peor=col_peor_final
                        ),
                        use_container_width=True,
                        hide_index=True
                    )

                    # ------------------------------
                    # Gráfica cumplimiento
                    # ------------------------------
                    st.markdown("**% Cumplimiento semanal**")
                    mostrar_boton_comentario(
                        "cobranza_cumplimiento",
                        comentario_cobranza_cumplimiento(evol_cobranza, col_cump)
                    )

                    fig_cump = grafica_cumplimiento(evol_cobranza, col_cump)
                    st.plotly_chart(fig_cump, use_container_width=True)

                    # ------------------------------
                    # Gráfica cuota vs pago
                    # ------------------------------
                    st.markdown("**Cuota total vs Pago total**")
                    mostrar_boton_comentario(
                        "cobranza_cuota_pago",
                        comentario_cobranza_cuota_pago(evol_cobranza, col_cuota, col_pago, col_cump)
                    )

                    fig_cp = grafica_cuota_pago(
                        evol=evol_cobranza,
                        col_cuota=col_cuota,
                        col_pago=col_pago,
                        col_mejor=col_mejor_final,
                        col_peor=col_peor_final
                    )
                    st.plotly_chart(fig_cp, use_container_width=True)

                    # ------------------------------
                    # Tabla últimas 5 semanas
                    # ------------------------------
                    tabla_ultimas_5_cobranza = crear_tabla_ultimas_5_cobranza(
                        evol=evol_cobranza,
                        col_cuota=col_cuota,
                        col_pago=col_pago,
                        col_cump=col_cump
                    )

                    st.markdown("**Últimas 5 semanas de cobranza**")
                    mostrar_boton_comentario(
                        "cobranza_ultimas_5",
                        comentario_tabla_cobranza(tabla_ultimas_5_cobranza, col_cuota, col_pago, col_cump)
                    )

                    tabla_ultimas_5_fmt = formato_ultimas_5_cobranza(
                        tabla=tabla_ultimas_5_cobranza,
                        col_cuota=col_cuota,
                        col_pago=col_pago,
                        col_cump=col_cump
                    )

                    st.dataframe(
                        estilo_ultimas_5_cobranza(tabla_ultimas_5_fmt),
                        use_container_width=True,
                        hide_index=True
                    )


                    # ------------------------------
                    # Top / Bottom de Cobranza
                    # ------------------------------
                    st.markdown("**Top / Bottom de cobranza por variable**")

                    niveles_top_bottom_cobranza = [
                        c for c in [
                            "País",
                            "Subdireccion",
                            "Zona",
                            "Sucursal",
                            "Ruta",
                        ]
                        if c in df_cobranza_top_bottom_base.columns
                    ]

                    variables_top_bottom_cobranza = [
                        c for c in [
                            col_cuota,
                            col_pago,
                            col_cump,
                            col_mejor_final,
                            col_peor_final,
                        ]
                        if c and c in df_cobranza_top_bottom_base.columns or c == col_cump
                    ]

                    # Quita duplicados conservando el orden.
                    variables_top_bottom_cobranza = list(dict.fromkeys(variables_top_bottom_cobranza))

                    semana_top_bottom_cobranza = obtener_ultima_semana_cobranza(df_cobranza_top_bottom_base)

                    if not niveles_top_bottom_cobranza:
                        st.info("No hay niveles de estructura disponibles en la hoja Cobranza para construir el Top / Bottom.")
                    elif not variables_top_bottom_cobranza:
                        st.info("No hay variables de cobranza disponibles para construir el Top / Bottom.")
                    elif semana_top_bottom_cobranza is None:
                        st.info("No hay semanas válidas en Cobranza para construir el Top / Bottom.")
                    else:
                        col_tabla_top_bottom_cob, col_opciones_top_bottom_cob = st.columns([2.2, 1])

                        with col_opciones_top_bottom_cob:
                            tipo_top_bottom_cobranza = st.radio(
                                "Top / Bottom",
                                options=["Top", "Bottom"],
                                horizontal=True,
                                key="tipo_top_bottom_cobranza"
                            )

                            nivel_top_bottom_cobranza = st.selectbox(
                                "Estructura",
                                options=niveles_top_bottom_cobranza,
                                index=0,
                                key="nivel_top_bottom_cobranza"
                            )

                            variable_top_bottom_cobranza = st.selectbox(
                                "Variable de cobranza",
                                options=variables_top_bottom_cobranza,
                                index=0,
                                key="variable_top_bottom_cobranza"
                            )

                            cantidad_top_bottom_cobranza = st.number_input(
                                "Cantidad",
                                min_value=3,
                                max_value=30,
                                value=10,
                                step=1,
                                key="cantidad_top_bottom_cobranza"
                            )

                        tabla_top_bottom_cobranza = construir_top_bottom_cobranza(
                            df_cobranza_base=df_cobranza_top_bottom_base,
                            nivel_top_bottom=nivel_top_bottom_cobranza,
                            variable_top_bottom=variable_top_bottom_cobranza,
                            col_cuota=col_cuota,
                            col_pago=col_pago,
                            col_cump=col_cump,
                            col_mejor=col_mejor_final,
                            col_peor=col_peor_final,
                            tipo_ranking=tipo_top_bottom_cobranza,
                            cantidad=int(cantidad_top_bottom_cobranza),
                            semana_objetivo=semana_top_bottom_cobranza
                        )

                        with col_tabla_top_bottom_cob:
                            if tabla_top_bottom_cobranza.empty:
                                st.info("No hay datos suficientes para mostrar el Top / Bottom de Cobranza con la selección actual.")
                            else:
                                st.caption(
                                    f"Semana {semana_top_bottom_cobranza} | {tipo_top_bottom_cobranza} "
                                    f"por {nivel_top_bottom_cobranza} | Moneda: {etiqueta_moneda(modo_moneda)}"
                                )

                                st.dataframe(
                                    aplicar_formato_top_bottom_cobranza(tabla_top_bottom_cobranza),
                                    use_container_width=True,
                                    hide_index=True
                                )

                                csv_top_bottom_cobranza = tabla_top_bottom_cobranza.to_csv(index=False).encode("utf-8-sig")

                                st.download_button(
                                    label="Descargar Top / Bottom Cobranza",
                                    data=csv_top_bottom_cobranza,
                                    file_name=f"top_bottom_cobranza_{nivel_top_bottom_cobranza}_semana_{semana_top_bottom_cobranza}.csv",
                                    mime="text/csv"
                                )

                        comentario_top_bottom_cobranza = generar_comentario_top_bottom_cobranza(
                            tabla_top_bottom=tabla_top_bottom_cobranza,
                            tipo_top_bottom=tipo_top_bottom_cobranza,
                            nivel_top_bottom=nivel_top_bottom_cobranza,
                            semana_actual=semana_top_bottom_cobranza
                        )
                        mostrar_boton_comentario("top_bottom_cobranza", comentario_top_bottom_cobranza)
