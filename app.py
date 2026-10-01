
import os
import base64
import html
import re
from pathlib import Path
from io import BytesIO
from urllib.parse import quote

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import plotly.express as px
import plotly.graph_objects as go

try:
    from pptx import Presentation
    from pptx.chart.data import ChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    PPTX_DISPONIBLE = True
except Exception:
    PPTX_DISPONIBLE = False


# ============================================================
# CONFIGURACIÓN
# ============================================================
st.set_page_config(
    page_title="Indicadores de Productividad y Cobranza",
    page_icon="📊",
    layout="wide"
)

# Todas las rutas se resuelven desde la carpeta donde vive este script. Esto
# funciona igual al ejecutarlo localmente o al desplegarlo en Streamlit Cloud.
try:
    CARPETA_APP = Path(__file__).resolve().parent
except NameError:
    CARPETA_APP = Path.cwd()

# Cartera se consolida desde un Parquet o desde los Excel guardados en esta
# carpeta. En Streamlit acepta tanto Concentrados/ como la raíz junto a app.py.
# La variable de entorno permite cambiar la ubicación sin editar el código.
CARPETA_CONCENTRADOS_REPOSITORIO = CARPETA_APP / "Concentrados"
RUTA_CARTERA_CONCENTRADOS = Path(
    os.environ.get(
        "CARPETA_CARTERA_CONCENTRADOS",
        str(
            CARPETA_CONCENTRADOS_REPOSITORIO
            if CARPETA_CONCENTRADOS_REPOSITORIO.is_dir()
            else CARPETA_APP
        ),
    )
)
RUTA_DEFAULT = str(RUTA_CARTERA_CONCENTRADOS)

# Los Parquet semanales están en la carpeta Cobranza del repositorio.
RUTA_COBRANZA_PARQUET = CARPETA_APP / "Cobranza"

# Fórmulas equivalentes a las medidas del modelo de Cobranza.
COLUMNAS_PARQUET_CUOTA_TOTAL = [
    "Cuota_cobranza_del_dia",
    "Cuota_cobranza_1_sem",
    "Pago_Adelanto",
]

COLUMNAS_PARQUET_RECUPERACION = [
    "Pago_cobranza_del_dia",
    "Pago_Pendiente_sin_atraso",
    "Pago_Adelanto",
    "Recuperado_cobranza_1_sem",
]

# Las gráficas de Cobranza muestran únicamente los periodos más recientes.
# Las tablas y archivos descargables conservan todo el histórico disponible.
MAX_SEMANAS_VISUALES_COBRANZA = 20

# Nombre de la imagen de fondo. Debe estar en la misma carpeta que este script.
NOMBRE_IMAGEN_FONDO = "ChatGPT Image 19 may 2026, 11_58_09 a.m."

# Oculta la tarjeta superior de carga de archivo y el expander de control de datos.
# El tablero seguirá usando RUTA_DEFAULT como origen de Cartera.
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
    "IP",
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
    "IP",
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
    return CARPETA_APP


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

# ============================================================
# VISTA PARA TELÉFONO
# ============================================================
if st.session_state.get("vista_telefono", False):
    st.markdown(
        """
        <style>
        /* Simula un ancho tipo iPhone en toda la aplicación sin tocar la lógica. */
        [data-testid="stAppViewContainer"] > .main {
            max-width: 430px !important;
            margin-left: auto !important;
            margin-right: auto !important;
            border-left: 1px solid rgba(226,232,240,0.95) !important;
            border-right: 1px solid rgba(226,232,240,0.95) !important;
            box-shadow: 0 0 30px rgba(15,23,42,0.18) !important;
            overflow-x: hidden !important;
        }
        .main .block-container {
            max-width: 430px !important;
            padding-left: 0.85rem !important;
            padding-right: 0.85rem !important;
        }
        div[data-testid="stHorizontalBlock"] {
            flex-wrap: wrap !important;
            gap: 0.55rem !important;
        }
        div[data-testid="column"] {
            min-width: 100% !important;
            width: 100% !important;
            flex: 1 1 100% !important;
        }
        .landing-title {
            font-size: 30px !important;
            line-height: 1.1 !important;
        }
        .landing-subtitle {
            font-size: 15px !important;
        }
        .unidad-logo,
        .unidad-logo-placeholder {
            width: 190px !important;
            height: 78px !important;
        }
        .top-filter-title,
        h1, h2, h3 {
            font-size: 26px !important;
            line-height: 1.15 !important;
        }
        .comentario-amplio {
            width: 100% !important;
            max-width: 100% !important;
            font-size: 15px !important;
            padding: 14px 16px !important;
        }
        .kpi-value {
            font-size: 26px !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

st.markdown(
    """
    <style>
    .main .block-container {
        padding-top: 1.3rem;
        padding-bottom: 2rem;
        background: var(--app-bg-card, rgba(255,255,255,0.94)) !important;
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
        background: var(--app-panel-bg, #ffffff) !important;
        color: var(--app-text-main, #111827) !important;
        border: 1px solid var(--app-border, rgba(226,232,240,0.95));
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
# MODO CLARO / OSCURO AUTOMÁTICO
# Este bloque debe quedar después de los estilos principales.
# Corrige labels invisibles en usuarios con modo oscuro del navegador/Windows.
# ============================================================
st.markdown(
    """
    <style>
    :root {
        --app-bg-card: rgba(255,255,255,0.94);
        --app-panel-bg: rgba(255,255,255,0.97);
        --app-panel-solid: #ffffff;
        --app-text-main: #111827;
        --app-text-soft: #374151;
        --app-brand: #082567;
        --app-brand-soft: #dbeafe;
        --app-accent: #d9c322;
        --app-border: rgba(226,232,240,0.95);
        --app-input-bg: #ffffff;
        --app-chart-bg: rgba(255,255,255,0.96);
        --app-shadow: 0 8px 24px rgba(15, 23, 42, 0.09);
    }

    html,
    body,
    .stApp,
    [data-testid="stAppViewContainer"] {
        color-scheme: light dark !important;
    }

    .main .block-container {
        background: var(--app-bg-card) !important;
        color: var(--app-text-main) !important;
        border-color: var(--app-border) !important;
    }

    .top-filter-card,
    .top-bottom-opciones-card,
    div[data-testid="stPlotlyChart"],
    div[data-testid="stDataFrame"],
    div[data-testid="stTable"],
    div[data-testid="stMetric"],
    .kpi-card,
    .unidad-card,
    div[data-testid="stExpander"] {
        background: var(--app-panel-bg) !important;
        color: var(--app-text-main) !important;
        border-color: var(--app-border) !important;
        box-shadow: var(--app-shadow) !important;
    }

    .top-filter-title,
    h1, h2, h3, h4,
    .titulo,
    .kpi-label,
    .unidad-name,
    .top-bottom-opciones-title {
        color: var(--app-brand) !important;
        -webkit-text-fill-color: var(--app-brand) !important;
        opacity: 1 !important;
    }

    .subtitulo,
    .unidad-help,
    div[data-testid="stMarkdownContainer"] p {
        color: var(--app-text-soft) !important;
        -webkit-text-fill-color: var(--app-text-soft) !important;
        opacity: 1 !important;
    }

    /* Labels de filtros: Vista, Moneda, Marca, País */
    .top-filter-card label,
    .top-filter-card label p,
    .top-filter-card div[data-testid="stWidgetLabel"],
    .top-filter-card div[data-testid="stWidgetLabel"] p,
    .top-filter-card .stRadio label,
    .top-filter-card .stSelectbox label,
    .top-filter-card .stTextInput label,
    .top-filter-card .stFileUploader label {
        color: var(--app-brand) !important;
        -webkit-text-fill-color: var(--app-brand) !important;
        opacity: 1 !important;
        font-weight: 900 !important;
    }

    /* Texto de radios y contenido de filtros */
    .top-filter-card p,
    .top-filter-card span,
    .top-filter-card div[role="radiogroup"] label,
    .top-filter-card div[role="radiogroup"] p,
    .top-filter-card div[role="radiogroup"] span,
    .top-filter-card div[data-baseweb="radio"] label,
    .top-filter-card div[data-baseweb="radio"] p,
    .top-filter-card div[data-baseweb="radio"] span {
        color: var(--app-text-main) !important;
        -webkit-text-fill-color: var(--app-text-main) !important;
        opacity: 1 !important;
    }

    /* Selectbox e inputs */
    .top-filter-card div[data-baseweb="select"] > div,
    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div,
    input,
    textarea,
    select {
        background-color: var(--app-input-bg) !important;
        color: var(--app-text-main) !important;
        -webkit-text-fill-color: var(--app-text-main) !important;
        border-color: var(--app-border) !important;
        opacity: 1 !important;
    }

    .top-filter-card div[data-baseweb="select"] span,
    .top-filter-card div[data-baseweb="select"] div,
    .top-filter-card div[data-baseweb="select"] input,
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] div,
    div[data-baseweb="select"] input {
        color: var(--app-text-main) !important;
        -webkit-text-fill-color: var(--app-text-main) !important;
        opacity: 1 !important;
    }

    .top-filter-card div[data-baseweb="select"] svg,
    div[data-baseweb="select"] svg {
        fill: var(--app-text-main) !important;
        color: var(--app-text-main) !important;
        opacity: 1 !important;
    }

    div[role="listbox"],
    ul[role="listbox"] {
        background-color: var(--app-input-bg) !important;
        color: var(--app-text-main) !important;
    }

    div[role="option"],
    li[role="option"] {
        background-color: var(--app-input-bg) !important;
        color: var(--app-text-main) !important;
        -webkit-text-fill-color: var(--app-text-main) !important;
    }

    div[role="option"]:hover,
    li[role="option"]:hover {
        background-color: rgba(148,163,184,0.20) !important;
        color: var(--app-text-main) !important;
        -webkit-text-fill-color: var(--app-text-main) !important;
    }

    .kpi-value {
        color: var(--app-text-main) !important;
        -webkit-text-fill-color: var(--app-text-main) !important;
    }

    .unidad-seleccionada-pill,
    .top-filter-card div.stButton > button,
    div.stButton > button,
    div[data-testid="stDownloadButton"] > button {
        background-color: var(--app-brand) !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        border-color: var(--app-brand) !important;
        opacity: 1 !important;
    }

    .top-filter-card div.stButton > button p,
    .top-filter-card div.stButton > button span,
    div.stButton > button p,
    div.stButton > button span,
    div[data-testid="stDownloadButton"] > button p,
    div[data-testid="stDownloadButton"] > button span {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        opacity: 1 !important;
    }

    .comentario-amplio {
        background: var(--app-brand-soft) !important;
        color: var(--app-brand) !important;
        border-left-color: var(--app-accent) !important;
    }

    /* Modo oscuro automático según Windows/navegador */
    @media (prefers-color-scheme: dark) {
        :root {
            --app-bg-card: rgba(15,23,42,0.94);
            --app-panel-bg: rgba(15,23,42,0.96);
            --app-panel-solid: #0f172a;
            --app-text-main: #f8fafc;
            --app-text-soft: #e5e7eb;
            --app-brand: #dbeafe;
            --app-brand-soft: rgba(30,41,59,0.98);
            --app-accent: #d9c322;
            --app-border: rgba(148,163,184,0.35);
            --app-input-bg: #111827;
            --app-chart-bg: rgba(15,23,42,0.96);
            --app-shadow: 0 8px 24px rgba(0,0,0,0.28);
        }

        .stApp {
            background-image: linear-gradient(rgba(15,23,42,0.80), rgba(15,23,42,0.88)) !important;
            background-color: #020617 !important;
            background-attachment: scroll !important;
        }

        header[data-testid="stHeader"] {
            background: rgba(2,6,23,0) !important;
        }

        .main .block-container {
            box-shadow: 0 14px 40px rgba(0,0,0,0.35) !important;
        }

        .top-filter-title,
        h1, h2, h3, h4,
        .titulo,
        .kpi-label,
        .unidad-name,
        .top-bottom-opciones-title,
        .top-filter-card label,
        .top-filter-card label p,
        .top-filter-card div[data-testid="stWidgetLabel"],
        .top-filter-card div[data-testid="stWidgetLabel"] p,
        .top-filter-card .stRadio label,
        .top-filter-card .stSelectbox label {
            color: #dbeafe !important;
            -webkit-text-fill-color: #dbeafe !important;
            opacity: 1 !important;
        }

        .top-filter-card p,
        .top-filter-card span,
        .top-filter-card div[role="radiogroup"] label,
        .top-filter-card div[role="radiogroup"] p,
        .top-filter-card div[role="radiogroup"] span,
        .subtitulo,
        .unidad-help,
        div[data-testid="stMarkdownContainer"] p {
            color: #e5e7eb !important;
            -webkit-text-fill-color: #e5e7eb !important;
            opacity: 1 !important;
        }

        .top-filter-card div[data-baseweb="select"] > div,
        div[data-baseweb="select"] > div,
        div[data-baseweb="input"] > div,
        input,
        textarea,
        select {
            background-color: #111827 !important;
            color: #f8fafc !important;
            -webkit-text-fill-color: #f8fafc !important;
            border: 1.5px solid rgba(219,234,254,0.45) !important;
        }

        .top-filter-card div[data-baseweb="select"] span,
        .top-filter-card div[data-baseweb="select"] div,
        .top-filter-card div[data-baseweb="select"] input,
        div[data-baseweb="select"] span,
        div[data-baseweb="select"] div,
        div[data-baseweb="select"] input {
            color: #f8fafc !important;
            -webkit-text-fill-color: #f8fafc !important;
            opacity: 1 !important;
        }

        .top-filter-card div[data-baseweb="select"] svg,
        div[data-baseweb="select"] svg {
            fill: #f8fafc !important;
            color: #f8fafc !important;
        }

        div[role="listbox"],
        ul[role="listbox"],
        div[role="option"],
        li[role="option"] {
            background-color: #111827 !important;
            color: #f8fafc !important;
            -webkit-text-fill-color: #f8fafc !important;
        }

        div[role="option"]:hover,
        li[role="option"]:hover {
            background-color: #1e293b !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
        }

        .comentario-amplio {
            background: rgba(30,41,59,0.98) !important;
            color: #f8fafc !important;
            -webkit-text-fill-color: #f8fafc !important;
        }

        .comentario-amplio-titulo {
            background: #dbeafe !important;
            color: #082567 !important;
            -webkit-text-fill-color: #082567 !important;
        }

        .kpi-delta-positive {
            color: #bbf7d0 !important;
            background: rgba(22,101,52,0.45) !important;
        }

        .kpi-delta-negative {
            color: #fecaca !important;
            background: rgba(127,29,29,0.45) !important;
        }

        .kpi-delta-neutral {
            color: #e5e7eb !important;
            background: rgba(71,85,105,0.55) !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True
)


# Ajuste puntual: botones/píldoras del resumen visibles en modo claro y oscuro.
st.markdown(
    """
    <style>
    .modal-resumen-meta {
        background-color: #082567 !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        opacity: 1 !important;
    }

    .modal-resumen-meta * {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        opacity: 1 !important;
    }

    div.stButton > button,
    div.stButton > button *,
    div[data-testid="stDownloadButton"] > button,
    div[data-testid="stDownloadButton"] > button * {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
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
    if modo_moneda != "Pesos mexicanos":
        # En moneda local no hay nada que transformar. Devolver la misma base
        # evita duplicar en memoria cientos de miles de registros en cada rerun.
        return df_base

    df_tmp = df_base.copy()

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
    return "Pesos mexicanos (MXN)" if modo_moneda == "Pesos mexicanos" else "Moneda local"


def obtener_query_param(nombre: str, default: str = "") -> str:
    """Lee parámetros de URL de forma compatible con versiones recientes y anteriores de Streamlit."""
    try:
        valor = st.query_params.get(nombre, default)
        if isinstance(valor, list):
            return valor[0] if valor else default
        return valor if valor is not None else default
    except Exception:
        try:
            valores = st.experimental_get_query_params().get(nombre, [default])
            return valores[0] if valores else default
        except Exception:
            return default


def ir_a_pagina_resumen():
    """Abre el resumen en una vista completa del tablero, no en modal."""
    try:
        st.query_params["vista"] = "resumen_pais"
    except Exception:
        st.experimental_set_query_params(vista="resumen_pais")
    st.rerun()


def volver_al_tablero():
    """Regresa al tablero principal limpiando el parámetro de vista."""
    try:
        if "vista" in st.query_params:
            del st.query_params["vista"]
    except Exception:
        st.experimental_set_query_params()
    st.rerun()


def filtrar_niveles_top_bottom_visibles(df_base: pd.DataFrame, niveles: list[str]) -> list[str]:
    """
    Evita mostrar Top / Bottom por niveles que solo tienen un valor visible.
    Así, cuando Presico México queda como único país, no aparece el Top por País.
    """
    niveles_visibles = []
    if df_base is None or df_base.empty:
        return niveles_visibles

    for nivel in niveles:
        if nivel not in df_base.columns:
            continue

        n_unicos = (
            df_base[nivel]
            .dropna()
            .astype(str)
            .str.strip()
            .replace("", np.nan)
            .dropna()
            .nunique()
        )

        if n_unicos > 1:
            niveles_visibles.append(nivel)

    return niveles_visibles


def limpiar_datos(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = normalizar_columnas(df)

    columnas_texto = [
        c for c in df.columns
        if pd.api.types.is_object_dtype(df[c])
        or pd.api.types.is_string_dtype(df[c])
    ]
    for c in columnas_texto:
        df[c] = df[c].astype("string").str.strip()
        df[c] = df[c].replace({"nan": pd.NA, "None": pd.NA, "": pd.NA})

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

    # IP = Índice de Productividad. Se calcula como porcentaje base 0-100.
    if "Clientes al corriente" in df.columns and "Clientes Totales" in df.columns:
        clientes_corriente = pd.to_numeric(df["Clientes al corriente"], errors="coerce").fillna(0)
        clientes_totales = pd.to_numeric(df["Clientes Totales"], errors="coerce").fillna(0)
        df["IP"] = np.where(clientes_totales == 0, 0, (clientes_corriente / clientes_totales) * 100)

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


def _periodo_desde_nombre_parquet(ruta: Path) -> tuple[int, int] | None:
    """Extrae (año, semana) del nombre estándar de los consolidados."""
    coincidencia = re.search(
        r"Semana_(\d{1,2})_(\d{4})_",
        ruta.name,
        flags=re.IGNORECASE,
    )
    if coincidencia is None:
        return None

    semana = int(coincidencia.group(1))
    anio = int(coincidencia.group(2))
    if not 1 <= semana <= 53:
        return None
    return anio, semana


def firma_archivos_cobranza_parquet(ruta_carpeta: str) -> tuple:
    """
    Devuelve una firma de los archivos que invalida la caché cuando se agrega
    o actualiza un consolidado. Si hay más de un archivo para la misma
    semana/año, conserva el de modificación más reciente.
    """
    carpeta = Path(ruta_carpeta)
    if not carpeta.is_dir():
        raise FileNotFoundError(
            "No encontré la carpeta de Parquet de Cobranza:\n"
            f"{carpeta}\n\n"
            "Crea la carpeta Cobranza junto al script y sube allí los archivos .parquet."
        )

    archivos_por_periodo: dict[tuple[int, int], Path] = {}
    for ruta in carpeta.glob("*.parquet"):
        periodo = _periodo_desde_nombre_parquet(ruta)
        if periodo is None:
            continue

        ruta_actual = archivos_por_periodo.get(periodo)
        if ruta_actual is None or ruta.stat().st_mtime_ns > ruta_actual.stat().st_mtime_ns:
            archivos_por_periodo[periodo] = ruta

    if not archivos_por_periodo:
        raise FileNotFoundError(
            f"No encontré archivos .parquet semanales válidos en:\n{carpeta}"
        )

    firma = []
    for (anio, semana), ruta in sorted(archivos_por_periodo.items()):
        datos_archivo = ruta.stat()
        firma.append(
            (str(ruta), anio, semana, datos_archivo.st_size, datos_archivo.st_mtime_ns)
        )
    return tuple(firma)


@st.cache_data(show_spinner=False, ttl=300)
def cargar_cobranza_desde_parquet(ruta_carpeta: str, firma_archivos: tuple) -> pd.DataFrame:
    """
    Consolida los Parquet semanales de Cobranza en una sola tabla.

    El año y la semana se toman del nombre del archivo para que el histórico
    funcione incluso si esas columnas no existen dentro del Parquet.
    """
    partes = []
    errores = []

    for ruta_texto, anio, semana, _tamano, _fecha_modificacion in firma_archivos:
        ruta = Path(ruta_texto)
        try:
            parte = pd.read_parquet(ruta)
        except ImportError as exc:
            raise RuntimeError(
                "Falta el motor para leer archivos Parquet. Instala 'pyarrow' "
                "en el mismo entorno donde ejecutas Streamlit:\n"
                "pip install pyarrow"
            ) from exc
        except Exception as exc:
            errores.append(f"{ruta.name}: {exc}")
            continue

        if parte is None or parte.empty:
            continue

        renombres = {}
        for columna in parte.columns:
            nombre = str(columna).strip()
            clave = re.sub(
                r"[^a-z0-9]+",
                "_",
                normalizar_texto_tc(nombre).lower(),
            ).strip("_")

            if clave == "pais":
                renombres[columna] = "País"
            elif clave == "unidad_de_negocio":
                # En los consolidados de Cobranza este campo contiene marcas
                # (La Casita, Presico MX, etc.), no la unidad superior del tablero.
                renombres[columna] = "Marca"
            elif clave == "ruta":
                renombres[columna] = "Ruta"
            elif clave.startswith("semana_del_a"):
                renombres[columna] = "Semana del año"

        parte = parte.rename(columns=renombres)

        columnas_requeridas = list(dict.fromkeys(
            COLUMNAS_PARQUET_CUOTA_TOTAL + COLUMNAS_PARQUET_RECUPERACION
        ))
        faltantes = [c for c in columnas_requeridas if c not in parte.columns]
        if faltantes:
            errores.append(
                f"{ruta.name}: faltan columnas requeridas {faltantes}"
            )
            continue

        for columna in columnas_requeridas:
            parte[columna] = pd.to_numeric(parte[columna], errors="coerce").fillna(0)

        # Cuota Total Cobranza = cuota del día + cuota de una semana + adelanto.
        parte["Cuota Total Cobranza"] = parte[COLUMNAS_PARQUET_CUOTA_TOTAL].sum(axis=1)

        # Recuperación semana = pago del día + pendiente sin atraso + adelanto
        # + recuperado de una semana. Pago_Adelanto_No_Natural no interviene.
        parte["Recuperación semana"] = parte[COLUMNAS_PARQUET_RECUPERACION].sum(axis=1)
        parte["Año"] = int(anio)
        parte["Semana del año"] = int(semana)

        cuota = pd.to_numeric(parte["Cuota Total Cobranza"], errors="coerce").fillna(0)
        pago = pd.to_numeric(parte["Recuperación semana"], errors="coerce").fillna(0)
        parte["% de Cumplimiento"] = np.where(cuota == 0, np.nan, pago / cuota)
        partes.append(parte)

    if errores:
        detalle = "\n".join(f"- {mensaje}" for mensaje in errores[:10])
        adicionales = len(errores) - 10
        if adicionales > 0:
            detalle += f"\n- ... y {adicionales} archivo(s) más."
        raise RuntimeError(
            "No fue posible cargar completo el histórico Parquet de Cobranza:\n"
            f"{detalle}"
        )

    if not partes:
        raise ValueError(
            "Los archivos Parquet de Cobranza no contienen registros utilizables."
        )

    cobranza = pd.concat(partes, ignore_index=True, sort=False)
    cobranza = limpiar_datos(cobranza)
    cobranza.attrs["ruta_origen"] = ruta_carpeta
    cobranza.attrs["archivos_cargados"] = len(partes)
    return cobranza


COLUMNAS_ORIGEN_CARTERA_CONCENTRADOS = {
    "Semana",
    "FECHA",
    "Unidad de negocio",
    "Unidad de Negocio",
    "Ruta",
    "id_y_localidad",
    "Clientes Totales",
    "Clientes al corriente",
    "Faltas",
    "Cartera Total",
    "Cartera sin atrasos",
    "Nunca Abonada",
    "Coord Totales",
    "Coord prod",
    "Coord en desarrollo",
    "Coord impro",
    "País",
}


def firma_archivos_cartera_concentrados(ruta_carpeta: str) -> tuple:
    """Firma los Parquet/Excel de Cartera para invalidar la caché al cambiarlos."""
    carpeta = Path(ruta_carpeta)
    if not carpeta.is_dir():
        raise FileNotFoundError(
            "No encontré la carpeta de concentrados de Cartera:\n"
            f"{carpeta}"
        )

    archivos_parquet = sorted(
        ruta
        for ruta in carpeta.iterdir()
        if ruta.is_file()
        and not ruta.name.startswith("~$")
        and ruta.suffix.lower() == ".parquet"
    )
    archivos_excel = sorted(
        ruta
        for ruta in carpeta.iterdir()
        if ruta.is_file()
        and not ruta.name.startswith("~$")
        and ruta.suffix.lower() in {".xlsx", ".xlsm", ".xlsb", ".xls"}
    )
    # Si existe un Parquet consolidado, se usa exclusivamente ese formato para
    # evitar volver a leer los Excel grandes o duplicar los mismos registros.
    archivos = archivos_parquet if archivos_parquet else archivos_excel
    if not archivos:
        raise FileNotFoundError(
            f"No encontré archivos Parquet o Excel de Cartera en:\n{carpeta}"
        )

    return tuple(
        (
            str(ruta),
            ruta.stat().st_size,
            ruta.stat().st_mtime_ns,
        )
        for ruta in archivos
    )


def _detectar_hoja_cartera_concentrado(excel: pd.ExcelFile) -> tuple[str, list[str]]:
    """Localiza la hoja de detalle por sus encabezados, sin depender del nombre."""
    requeridas = {"Semana", "Clientes Totales", "Cartera Total", "Ruta"}

    for hoja in excel.sheet_names:
        try:
            encabezados = pd.read_excel(excel, sheet_name=hoja, nrows=0)
        except Exception:
            continue

        columnas = [str(columna).strip() for columna in encabezados.columns]
        if requeridas.issubset(set(columnas)):
            return hoja, columnas

    raise ValueError(
        "No encontré una hoja de detalle de Cartera. Se requieren al menos "
        "las columnas Semana, Clientes Totales, Cartera Total y Ruta."
    )


def _preparar_cartera_concentrado(parte: pd.DataFrame, archivo: str) -> pd.DataFrame:
    """Adapta el formato de los concentrados al esquema usado por el tablero."""
    parte = parte.copy()
    parte.columns = [str(columna).strip() for columna in parte.columns]
    parte = parte.rename(
        columns={
            "Semana": "Semana del año",
            "FECHA": "Fecha",
            # En estos libros, la variante con 'negocio' en minúscula es la marca.
            "Unidad de negocio": "Marca",
            "id_y_localidad": "coordinadora_id",
            "Cartera sin atrasos": "Saldo Cartera",
            "Nunca Abonada": "Nunca Abonados",
        }
    )

    if "Marca" in parte.columns:
        parte["Marca"] = (
            parte["Marca"]
            .astype("string")
            .str.strip()
            .str.upper()
            .replace({"": pd.NA, "-": pd.NA})
        )

    if "País" in parte.columns:
        parte["País"] = parte["País"].replace(
            {"Peru": "Perú", "PERU": "Perú", "Mexico": "México", "MEXICO": "México"}
        )

    if "Semana del año" in parte.columns:
        parte["Semana del año"] = pd.to_numeric(
            parte["Semana del año"]
            .astype("string")
            .str.extract(r"(\d{1,2})", expand=False),
            errors="coerce",
        ).astype("Int64")

    if "Fecha" in parte.columns:
        fechas = pd.to_datetime(parte["Fecha"], errors="coerce")
        parte["Año"] = fechas.dt.year.astype("Int64")

    columnas_numericas = [
        "Clientes Totales",
        "Clientes al corriente",
        "Faltas",
        "Cartera Total",
        "Saldo Cartera",
        "Nunca Abonados",
        "Coord Totales",
        "Coord prod",
        "Coord en desarrollo",
        "Coord impro",
    ]
    for columna in columnas_numericas:
        if columna in parte.columns:
            parte[columna] = pd.to_numeric(parte[columna], errors="coerce").fillna(0)

    if {"Cartera Total", "Saldo Cartera"}.issubset(parte.columns):
        parte["Saldo en atraso"] = (
            parte["Cartera Total"] - parte["Saldo Cartera"]
        ).clip(lower=0)

    # Cada fila del concentrado representa una coordinadora/localidad. Las
    # banderas indican su clasificación en la semana correspondiente.
    if "Tipo Coordinadora" not in parte.columns:
        productiva = parte.get("Coord prod", pd.Series(0, index=parte.index)).gt(0)
        desarrollo = parte.get("Coord en desarrollo", pd.Series(0, index=parte.index)).gt(0)
        improductiva = parte.get("Coord impro", pd.Series(0, index=parte.index)).gt(0)
        parte["Tipo Coordinadora"] = np.select(
            [productiva, desarrollo, improductiva],
            ["Productiva", "En Desarrollo", "Improductiva"],
            default="Secundaria",
        )

    parte = parte.drop(columns=["Fecha"], errors="ignore")
    return parte


def optimizar_memoria_cartera(df: pd.DataFrame) -> pd.DataFrame:
    """Reduce memoria sin cambiar valores ni columnas usadas por el tablero."""
    columnas_categoria = [
        "Unidad de Negocio",
        "Marca",
        "País",
        "Region",
        "Subdireccion",
        "Zona",
        "Sucursal",
        "Ruta",
        "coordinadora_id",
        "Tipo Coordinadora",
    ]
    for columna in columnas_categoria:
        if columna in df.columns:
            df[columna] = df[columna].astype("category")

    for columna in [
        "Semana del año",
        "Año",
        "Clientes Totales",
        "Clientes al corriente",
        "Faltas",
        "Nunca Abonados",
    ]:
        if columna in df.columns:
            df[columna] = pd.to_numeric(
                df[columna], errors="coerce", downcast="integer"
            )

    if "IP" in df.columns:
        df["IP"] = pd.to_numeric(df["IP"], errors="coerce", downcast="float")

    return df


@st.cache_resource(show_spinner=False)
def cargar_cartera_desde_concentrados(
    ruta_carpeta: str,
    firma_archivos: tuple,
) -> pd.DataFrame:
    """Combina automáticamente las hojas de detalle de todos los concentrados."""
    partes = []
    errores = []
    hojas_cargadas = []

    for ruta_texto, _tamano, _fecha_modificacion in firma_archivos:
        ruta = Path(ruta_texto)
        try:
            if ruta.suffix.lower() == ".parquet":
                parte = pd.read_parquet(ruta)
                hoja = "Parquet consolidado"
                if "Semana del año" not in parte.columns:
                    parte = _preparar_cartera_concentrado(parte, ruta.name)
            else:
                with pd.ExcelFile(ruta) as excel:
                    hoja, columnas = _detectar_hoja_cartera_concentrado(excel)
                    columnas_utiles = [
                        columna
                        for columna in columnas
                        if columna in COLUMNAS_ORIGEN_CARTERA_CONCENTRADOS
                    ]
                    parte = pd.read_excel(
                        excel,
                        sheet_name=hoja,
                        usecols=columnas_utiles,
                    )
                parte = _preparar_cartera_concentrado(parte, ruta.name)
        except Exception as exc:
            errores.append(f"{ruta.name}: {exc}")
            continue

        if parte is None or parte.empty:
            errores.append(f"{ruta.name}: la hoja {hoja} no contiene registros")
            continue

        partes.append(parte)
        hojas_cargadas.append(f"{ruta.name} [{hoja}]")

    if errores:
        detalle = "\n".join(f"- {mensaje}" for mensaje in errores)
        raise RuntimeError(
            "No fue posible cargar todos los concentrados de Cartera:\n"
            f"{detalle}"
        )

    if not partes:
        raise ValueError("Los concentrados no contienen registros de Cartera utilizables.")

    solo_parquet_preparado = all(
        Path(ruta_texto).suffix.lower() == ".parquet"
        for ruta_texto, _tamano, _fecha_modificacion in firma_archivos
    ) and all("Semana del año" in parte.columns for parte in partes)

    cartera = partes[0] if len(partes) == 1 else pd.concat(
        partes, ignore_index=True, sort=False
    )

    if solo_parquet_preparado:
        # El consolidado entregado ya está limpio y sin duplicados. Repetir
        # drop_duplicates sobre 837 mil filas multiplica el uso de memoria.
        filas_antes = len(cartera)
        cartera.attrs["duplicados_exactos_eliminados"] = 0
        cartera.attrs["filas_antes_limpieza"] = filas_antes
        cartera.attrs["filas_despues_limpieza"] = filas_antes
        cartera.attrs["preconsolidado_cartera"] = True
    else:
        cartera = limpiar_datos(cartera)
        cartera.attrs["preconsolidado_cartera"] = False

    cartera = optimizar_memoria_cartera(cartera)
    cartera.attrs["ruta_origen"] = ruta_carpeta
    cartera.attrs["archivos_cargados"] = len(partes)
    cartera.attrs["hojas_cargadas"] = hojas_cargadas
    return cartera


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



def dataframe_a_xlsx_bytes(df_exportar: pd.DataFrame, nombre_hoja: str = "Datos") -> bytes:
    """Convierte una tabla a archivo Excel en memoria para descarga."""
    salida = BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df_tmp = df_exportar.copy()
        # Excel no acepta zonas horarias en datetimes.
        for col in df_tmp.columns:
            if pd.api.types.is_datetime64_any_dtype(df_tmp[col]):
                try:
                    df_tmp[col] = df_tmp[col].dt.tz_localize(None)
                except Exception:
                    pass
        df_tmp.to_excel(writer, index=False, sheet_name=str(nombre_hoja)[:31] or "Datos")
        ws = writer.sheets[str(nombre_hoja)[:31] or "Datos"]
        for idx, col in enumerate(df_tmp.columns, start=1):
            muestra = df_tmp[col].head(200).astype("string").fillna("")
            max_len = max([len(str(col))] + [len(str(v)) for v in muestra.tolist()])
            ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = min(max(max_len + 2, 12), 45)
    return salida.getvalue()


def boton_descargar_xlsx(df_exportar: pd.DataFrame, etiqueta: str, nombre_archivo: str, key: str | None = None):
    if df_exportar is None or df_exportar.empty:
        return
    st.download_button(
        label=etiqueta,
        data=dataframe_a_xlsx_bytes(df_exportar, "Datos"),
        file_name=nombre_archivo if nombre_archivo.lower().endswith(".xlsx") else f"{nombre_archivo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=key,
    )

def mostrar_grafica_adaptable(fig: go.Figure, config: dict | None = None, key: str | None = None):
    """
    Muestra la gráfica Plotly normal.
    Se eliminó la lógica de imagen ampliable / kaleido para mantener todo estable.
    """
    config = config or {
        "displayModeBar": False,
        "scrollZoom": False,
        "doubleClick": False,
        "responsive": True,
    }
    st.plotly_chart(fig, use_container_width=True, config=config)


def mostrar_grafica_burbujas_amplia(fig: go.Figure, config: dict | None = None, key: str | None = None):
    """
    Muestra la gráfica de burbujas con un ancho fijo grande y scroll horizontal.
    Esto evita que en teléfono se comprima y se vea chica/amontonada.
    No usa kaleido ni imágenes: mantiene Plotly interactivo.
    """
    config = config or {
        "displayModeBar": False,
        "scrollZoom": False,
        "doubleClick": False,
        "responsive": False,
    }

    fig_amplia = go.Figure(fig)
    fig_amplia.update_layout(
        autosize=False,
        width=1350,
        height=860,
        margin=dict(t=95, b=90, l=105, r=80),
        font=dict(color="#111827", size=15),
        title=dict(
            text="Clientes Totales, Porcentaje de Faltas y Faltas por Sucursal",
            x=0.5,
            xanchor="center",
            font=dict(size=22, color="#111827")
        ),
    )
    fig_amplia.update_xaxes(title_font=dict(size=18), tickfont=dict(size=14))
    fig_amplia.update_yaxes(title_font=dict(size=18), tickfont=dict(size=14))

    html_plot = fig_amplia.to_html(
        full_html=False,
        include_plotlyjs="cdn",
        config=config,
        default_width="1350px",
        default_height="860px",
    )

    components.html(
        f"""
        <style>
        .chart-scroll-wrap {{
            width: 100%;
            overflow-x: auto;
            overflow-y: hidden;
            -webkit-overflow-scrolling: touch;
            background: rgba(255,255,255,0.98);
            border: 1px solid rgba(226,232,240,0.95);
            border-radius: 22px;
            padding: 12px;
            box-sizing: border-box;
            text-align: center;
        }}
        .chart-scroll-inner {{
            width: 1350px;
            min-width: 1350px;
            margin-left: auto;
            margin-right: auto;
            display: inline-block;
            text-align: left;
        }}
        .chart-scroll-hint {{
            font-family: Arial, sans-serif;
            color: #082567;
            font-size: 14px;
            font-weight: 800;
            margin: 0 auto 8px auto;
            text-align: center;
        }}
        </style>
        <div class="chart-scroll-wrap">
            <div class="chart-scroll-hint">Desliza horizontalmente para ver la gráfica completa.</div>
            <div class="chart-scroll-inner">{html_plot}</div>
        </div>
        """,
        height=915,
        scrolling=True,
    )



def mostrar_grafica_cobranza_amplia(
    fig: go.Figure,
    titulo: str = "Gráfica de cobranza",
    ancho: int = 1350,
    alto: int = 620,
    config: dict | None = None,
    key: str | None = None
):
    """
    Muestra las gráficas de Cobranza con ancho fijo y desplazamiento horizontal,
    igual que la gráfica de burbujas de Cartera. Esto evita que en teléfono
    se compriman las semanas, etiquetas y leyendas.
    """
    config = config or {
        "displayModeBar": False,
        "scrollZoom": False,
        "doubleClick": False,
        "responsive": False,
    }

    fig_amplia = go.Figure(fig)
    fig_amplia.update_layout(
        autosize=False,
        width=ancho,
        height=alto,
        margin=dict(t=85, b=95, l=90, r=fig.layout.margin.r if fig.layout.margin and fig.layout.margin.r else 80),
        font=dict(color="#082567", size=13),
    )
    fig_amplia.update_xaxes(tickfont=dict(size=12), title_font=dict(size=14))
    fig_amplia.update_yaxes(tickfont=dict(size=12), title_font=dict(size=14))

    html_plot = fig_amplia.to_html(
        full_html=False,
        include_plotlyjs="cdn",
        config=config,
        default_width=f"{ancho}px",
        default_height=f"{alto}px",
    )

    components.html(
        f"""
        <style>
        .chart-scroll-wrap-cobranza {{
            width: 100%;
            overflow-x: auto;
            overflow-y: hidden;
            -webkit-overflow-scrolling: touch;
            background: rgba(255,255,255,0.98);
            border: 1px solid rgba(226,232,240,0.95);
            border-radius: 22px;
            padding: 12px;
            box-sizing: border-box;
            text-align: center;
        }}
        .chart-scroll-inner-cobranza {{
            width: {ancho}px;
            min-width: {ancho}px;
            margin-left: auto;
            margin-right: auto;
            display: inline-block;
            text-align: left;
        }}
        .chart-scroll-hint-cobranza {{
            font-family: Arial, sans-serif;
            color: #082567;
            font-size: 14px;
            font-weight: 800;
            margin: 0 auto 8px auto;
            text-align: center;
        }}
        </style>
        <div class="chart-scroll-wrap-cobranza">
            <div class="chart-scroll-hint-cobranza">Desliza horizontalmente para ver la gráfica completa.</div>
            <div class="chart-scroll-inner-cobranza">{html_plot}</div>
        </div>
        """,
        height=alto + 70,
        scrolling=True,
    )
def crear_grafica_evolucion_fija(
    evol: pd.DataFrame,
    indicador_grafica: str,
    modo_moneda: str,
    altura: int = 500
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
    mostrar_valores_y_variaciones = len(df_plot) <= 8

    es_ip = normalizar_texto_tc(indicador_grafica).strip().lower() == "ip"

    if es_ip:
        df_plot["Texto valor"] = df_plot[indicador_grafica].apply(lambda x: f"{float(x):,.1f}%")
        df_plot["Texto variacion"] = df_plot["Variación vs anterior"].apply(
            lambda x: f"{float(x):+,.1f} pp" if pd.notna(x) else ""
        )
    else:
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
        y_max_fijo = y_max + base * (0.20 if mostrar_valores_y_variaciones else 0.10)
    else:
        y_min_fijo = y_min - rango * (0.22 if mostrar_valores_y_variaciones else 0.10)
        y_max_fijo = y_max + rango * (0.45 if mostrar_valores_y_variaciones else 0.10)

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=df_plot["Semana texto"],
            y=df_plot[indicador_grafica],
            mode="lines+markers" if mostrar_valores_y_variaciones else "lines",
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
                (
                    "<b>Semana:</b> %{customdata[1]:.0f}<br>"
                    f"<b>{indicador_grafica}:</b> %{{customdata[2]}}<br>"
                    "<b>Variación vs anterior:</b> %{customdata[0]}"
                    "<extra></extra>"
                )
                if mostrar_valores_y_variaciones
                else "<b>Semana:</b> %{customdata[1]:.0f}<extra></extra>"
            ),
            cliponaxis=False
        )
    )

    # Valor y variación se muestran como anotaciones separadas.
    # Ajusta yshift_variacion / yshift_valor si quieres más o menos distancia.
    yshift_valor = 16
    yshift_variacion = 46

    if mostrar_valores_y_variaciones:
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
                    text=f"<b>{fila['Texto variacion']}</b>",
                    showarrow=False,
                    yshift=yshift_variacion,
                    font=dict(color="#082567", size=12),
                    bgcolor="rgba(255,255,255,0.86)",
                    bordercolor="rgba(8,37,103,0.16)",
                    borderwidth=1,
                    borderpad=3
                )

    ticks_y = np.linspace(y_min_fijo, y_max_fijo, 5)

    if es_ip:
        titulo_eje_y = "IP (%)"
    else:
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
        ticktext=[f"{v:,.1f}%" for v in ticks_y] if es_ip else [formato_eje_compacto(v) for v in ticks_y],
        gridcolor="rgba(148,163,184,0.25)",
        zeroline=False,
        tickfont=dict(size=11, color="#64748b"),
        title_font=dict(size=12, color="#64748b")
    )

    config = {
        "displayModeBar": False,
        "scrollZoom": False,
        "doubleClick": False,
        "responsive": True,
        "staticPlot": False,
        "modeBarButtonsToRemove": [
            "zoom2d", "pan2d", "select2d", "lasso2d", "zoomIn2d", "zoomOut2d",
            "autoScale2d", "resetScale2d"
        ],
    }

    return fig, config

def tarjeta_kpi(label, valor, variacion=None):
    label_norm = normalizar_texto_tc(label).lower()
    es_porcentaje = str(label).strip().startswith("%") or "cumplimiento" in label_norm

    valor_fmt = formato_pct(valor, 2, False) if es_porcentaje else formato_numero(valor)

    if variacion is None or pd.isna(variacion):
        delta_html = ""
    else:
        variacion = float(variacion)
        variacion_fmt = formato_pct(variacion, 2, True) if es_porcentaje else formato_variacion(variacion)

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
    df_tmp = df_base

    for col, seleccion in filtros.items():
        if col == excluir_col:
            continue

        if col in df_tmp.columns and seleccion:
            df_tmp = df_tmp[df_tmp[col].astype(str).isin(seleccion)]

    return df_tmp


def aplicar_filtros_base(df_base: pd.DataFrame, semanas_sel: list[int], filtros: dict):
    df_tmp = df_base

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

    # País siempre funciona como puente entre la Unidad de Negocio elegida en
    # Cartera y los Parquet. Las demás dimensiones solo deben restringir
    # Cobranza cuando el usuario las seleccionó explícitamente. Así, Marca =
    # "Todos" conserva también marcas que existen únicamente en Cobranza,
    # como Presico MX y Presico CD.
    columnas_puente = ["País"]
    columnas_puente.extend([
        col for col in ["Marca", "Subdireccion", "Zona", "Sucursal", "Ruta"]
        if col in filtros and filtros.get(col)
    ])

    for col in columnas_puente:
        if col in df_tmp.columns and col in df_ref.columns:
            # Cartera y Cobranza pueden escribir el mismo catálogo de forma
            # distinta (por ejemplo, "México" frente a "Mexico"). Comparamos
            # una versión normalizada para ignorar acentos, mayúsculas y espacios.
            valores_validos = {
                normalizar_texto_tc(valor)
                for valor in df_ref[col].dropna().tolist()
                if normalizar_texto_tc(valor)
            }

            if valores_validos:
                valores_cobranza = df_tmp[col].apply(normalizar_texto_tc)
                df_tmp = df_tmp[valores_cobranza.isin(valores_validos)].copy()

    return df_tmp


def recalcular_ip_agregado(df_base: pd.DataFrame) -> pd.DataFrame:
    """Recalcula IP a nivel agregado para evitar sumar porcentajes."""
    if df_base.attrs.get("preconsolidado_cartera", False):
        return df_base

    df_tmp = df_base.copy()
    if "Clientes al corriente" in df_tmp.columns and "Clientes Totales" in df_tmp.columns:
        clientes_corriente = pd.to_numeric(df_tmp["Clientes al corriente"], errors="coerce").fillna(0)
        clientes_totales = pd.to_numeric(df_tmp["Clientes Totales"], errors="coerce").fillna(0)
        df_tmp["IP"] = np.where(clientes_totales == 0, 0, (clientes_corriente / clientes_totales) * 100)
    return df_tmp


# ============================================================
# FUNCIONES CARTERA
# ============================================================
def detectar_columna_cobranza_cartera(df: pd.DataFrame):
    return detectar_columna(POSIBLES_COLUMNAS_COBRANZA_CARTERA, df.columns)


def consolidar_grano_correcto(df_base: pd.DataFrame, indicadores: list[str]) -> pd.DataFrame:
    if df_base.attrs.get("preconsolidado_cartera", False):
        return df_base

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

    if "IP" in indicadores and "Clientes al corriente" in actual.index and "Clientes Totales" in actual.index:
        actual["IP"] = 0 if actual.get("Clientes Totales", 0) == 0 else (actual.get("Clientes al corriente", 0) / actual.get("Clientes Totales", 0)) * 100
        anterior["IP"] = 0 if anterior.get("Clientes Totales", 0) == 0 else (anterior.get("Clientes al corriente", 0) / anterior.get("Clientes Totales", 0)) * 100

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
    """
    Formatea tablas para visualización.
    Cambio importante:
    - IP se muestra como porcentaje cuando aparece como indicador en filas.
    - La variación de IP se muestra como puntos porcentuales, no como número entero.
    """
    if df_tabla is None or df_tabla.empty:
        return pd.DataFrame()

    df_fmt = df_tabla.copy().astype(object)

    for idx, fila in df_tabla.iterrows():
        indicador = str(fila.get("Indicador", ""))
        indicador_norm = normalizar_texto_tc(indicador).lower()
        es_ip = indicador_norm == "ip"
        es_pct = indicador.strip().startswith("%") or "cumplimiento" in indicador_norm

        for col in df_fmt.columns:
            if col == "Indicador":
                df_fmt.at[idx, col] = indicador
                continue

            valor = fila.get(col)
            nombre_col = str(col).strip()

            try:
                if es_ip:
                    if "Variación" in nombre_col or nombre_col.startswith("Var "):
                        df_fmt.at[idx, col] = "" if pd.isna(valor) else f"{float(valor):+,.1f} pp"
                    elif "% Var" in nombre_col:
                        df_fmt.at[idx, col] = formato_pct(valor, 1, True)
                    else:
                        df_fmt.at[idx, col] = "" if pd.isna(valor) else f"{float(valor):,.1f}%"

                elif es_pct:
                    if "Variación" in nombre_col or "% Var" in nombre_col or nombre_col.startswith("Var "):
                        df_fmt.at[idx, col] = formato_pct(valor, 2, True)
                    else:
                        df_fmt.at[idx, col] = formato_pct(valor, 2, False)

                elif "% Var" in nombre_col:
                    df_fmt.at[idx, col] = formato_pct(valor, 1, True)

                elif "Variación" in nombre_col or nombre_col.startswith("Var "):
                    df_fmt.at[idx, col] = formato_variacion(valor)

                else:
                    df_fmt.at[idx, col] = formato_numero(valor)
            except Exception:
                df_fmt.at[idx, col] = "" if pd.isna(valor) else str(valor)

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

    actual = recalcular_ip_agregado(actual)
    anterior = recalcular_ip_agregado(anterior)

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
    agrupado = recalcular_ip_agregado(agrupado)

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


def obtener_ultimo_periodo_cobranza(
    df_cobranza_base: pd.DataFrame,
) -> tuple[int | None, int] | None:
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
        df_tmp = df_tmp.dropna(subset=["Año"])
        if df_tmp.empty:
            return None
        df_tmp = df_tmp.sort_values(["Año", "Semana del año"])
        ultima = df_tmp.iloc[-1]
        return int(ultima["Año"]), int(ultima["Semana del año"])
    else:
        df_tmp = df_tmp.sort_values("Semana del año")
        return None, int(df_tmp.iloc[-1]["Semana del año"])


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
    semana_objetivo: int | None = None,
    anio_objetivo: int | None = None,
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
        periodo_objetivo = obtener_ultimo_periodo_cobranza(df_tmp)
        if periodo_objetivo is not None:
            anio_objetivo, semana_objetivo = periodo_objetivo

    if semana_objetivo is None:
        return pd.DataFrame()

    df_semana = df_tmp[df_tmp["Semana del año"] == int(semana_objetivo)].copy()
    if anio_objetivo is not None and "Año" in df_semana.columns:
        anios = pd.to_numeric(df_semana["Año"], errors="coerce")
        df_semana = df_semana[anios == int(anio_objetivo)].copy()

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
    """
    Crea una llave de movimiento SIN perder coordinadoras/registros.

    Antes la llave era coordinadora_id + País + Marca. Eso agrupaba como una sola
    coordinadora los casos donde el mismo ID aparecía en más de una Ruta, y por eso
    la matriz de desplazamiento no cuadraba con el total de la base.

    Ahora la llave incluye campos de estructura disponibles, especialmente Ruta.
    Además agrega un consecutivo por semana para no perder filas incluso si existen
    registros repetidos con la misma combinación de estructura.
    """
    df_tmp = df_base.copy()

    if columna_id not in df_tmp.columns:
        return df_tmp

    columnas_llave = [columna_id]

    # Ruta es clave para PRESICO porque hay coordinadoras con el mismo ID/País/Marca
    # en rutas distintas. Se incluyen más columnas sólo si existen en la base.
    for c in ["Unidad de Negocio", "País", "Marca", "Subdireccion", "Zona", "Sucursal", "Ruta"]:
        if c in df_tmp.columns and c not in columnas_llave:
            columnas_llave.append(c)

    df_tmp["_base_llave_coordinadora_marca"] = (
        df_tmp[columnas_llave]
        .astype("string")
        .fillna("")
        .apply(lambda s: s.str.strip())
        .agg("|".join, axis=1)
    )

    # Consecutivo por semana + llave base para conservar TODOS los registros.
    # Sin esto, si hubiera dos filas exactamente iguales, también se compactarían.
    if "Semana del año" in df_tmp.columns:
        df_tmp["_consecutivo_llave"] = (
            df_tmp
            .groupby(["Semana del año", "_base_llave_coordinadora_marca"], dropna=False)
            .cumcount()
            .astype(str)
        )
    else:
        df_tmp["_consecutivo_llave"] = (
            df_tmp
            .groupby(["_base_llave_coordinadora_marca"], dropna=False)
            .cumcount()
            .astype(str)
        )

    df_tmp["_llave_coordinadora_marca"] = (
        df_tmp["_base_llave_coordinadora_marca"]
        + "|OCURRENCIA_"
        + df_tmp["_consecutivo_llave"]
    )

    return df_tmp


def obtener_categoria_unica_por_semana(
    df_base: pd.DataFrame,
    semana: int,
    columna_id: str = "coordinadora_id",
    columna_categoria: str = "Tipo Coordinadora"
):
    """
    Devuelve una fila por registro de la semana, no una fila por coordinadora única.

    Antes se hacía drop_duplicates por llave, lo cual reducía el total de coordinadoras
    en la matriz. Para que el total cuadre con la base, aquí se conserva cada registro.
    """
    df_semana = df_base[df_base["Semana del año"] == semana].copy()

    columnas_necesarias = [columna_id, columna_categoria]

    for c in [
        "coordinadora_id",
        "Unidad de Negocio",
        "País",
        "Marca",
        "Subdireccion",
        "Zona",
        "Sucursal",
        "Ruta",
        "_base_llave_coordinadora_marca",
        "_consecutivo_llave",
    ]:
        if c in df_semana.columns and c not in columnas_necesarias:
            columnas_necesarias.append(c)

    df_semana = df_semana[columnas_necesarias].dropna(
        subset=[columna_id, columna_categoria]
    )

    if df_semana.empty:
        return pd.DataFrame(columns=columnas_necesarias)

    return df_semana.reset_index(drop=True)


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

    movimientos["Semana anterior"] = (
        movimientos["Semana anterior"].astype("string").fillna("Nueva")
    )
    movimientos["Semana actual"] = (
        movimientos["Semana actual"].astype("string").fillna("Baja")
    )

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


def matriz_desplazamiento_html(df_matriz: pd.DataFrame) -> str:
    """Renderiza la matriz completa como tabla responsiva, sin scroll interno."""
    ranking = {
        "Secundaria": 0,
        "Improductiva": 1,
        "En Desarrollo": 2,
        "Productiva": 3,
    }

    encabezados = [
        '<th class="matriz-col-origen">Semana anterior</th>',
        *[f"<th>{html.escape(str(col))}</th>" for col in df_matriz.columns],
    ]
    filas_html = []

    for fila in df_matriz.index:
        celdas = [f'<th scope="row">{html.escape(str(fila))}</th>']
        for columna in df_matriz.columns:
            clases = []
            if fila == "Total general" or columna == "Total general":
                clases.append("matriz-total")
            elif fila == "Nueva":
                clases.append("matriz-nueva")
            elif columna == "Baja":
                clases.append("matriz-baja")
            elif fila in ranking and columna in ranking:
                if ranking[columna] > ranking[fila]:
                    clases.append("matriz-mejora")
                elif ranking[columna] < ranking[fila]:
                    clases.append("matriz-retroceso")
                else:
                    clases.append("matriz-permanencia")

            valor = df_matriz.loc[fila, columna]
            valor_txt = "—" if pd.isna(valor) else f"{float(valor):,.0f}"
            celdas.append(
                f'<td class="{" ".join(clases)}">{html.escape(valor_txt)}</td>'
            )

        clase_fila = " class=\"matriz-fila-total\"" if fila == "Total general" else ""
        filas_html.append(f"<tr{clase_fila}>{''.join(celdas)}</tr>")

    return (
        '<div class="matriz-responsiva-sin-scroll">'
        '<table aria-label="Matriz de desplazamiento de coordinadoras">'
        f"<thead><tr>{''.join(encabezados)}</tr></thead>"
        f"<tbody>{''.join(filas_html)}</tbody>"
        "</table></div>"
    )


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


def tabla_improductivas_por_marca(movimientos: pd.DataFrame) -> pd.DataFrame:
    """
    Cuenta coordinadoras que pasaron a Improductiva por marca,
    sin importar su categoría anterior, excluyendo las que ya eran Improductivas.
    """
    if movimientos is None or movimientos.empty:
        return pd.DataFrame(columns=["Marca", "Coordinadoras que pasaron a Improductiva"])

    mov = movimientos.copy()
    mov = mov[
        (mov["Semana actual"].astype(str) == "Improductiva") &
        (mov["Semana anterior"].astype(str) != "Improductiva") &
        (mov["Semana anterior"].astype(str) != "Nueva")
    ].copy()

    if mov.empty:
        return pd.DataFrame(columns=["Marca", "Coordinadoras que pasaron a Improductiva"])

    if "Marca destino" in mov.columns:
        mov["Marca"] = mov["Marca destino"]
    elif "Marca origen" in mov.columns:
        mov["Marca"] = mov["Marca origen"]
    else:
        mov["Marca"] = "Sin marca"

    # Cuenta registros de movimiento, no llaves únicas, para no excluir coordinadoras.
    conteo = mov.groupby("Marca", dropna=False).size().reset_index(name="_llave_coordinadora_marca")

    conteo = conteo.rename(columns={"_llave_coordinadora_marca": "Coordinadoras que pasaron a Improductiva"})
    conteo["Marca"] = conteo["Marca"].astype("string").fillna("Sin marca")
    return conteo.sort_values("Coordinadoras que pasaron a Improductiva", ascending=False).reset_index(drop=True)


def mostrar_cuadro_resumen_movimientos(movimientos: pd.DataFrame):
    resumen_mov = calcular_resumen_movimientos(movimientos)

    st.markdown(
        f"""
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
        """,
        unsafe_allow_html=True
    )

    tabla_imp = tabla_improductivas_por_marca(movimientos)
    st.markdown(
        '<div class="cuadro-marca-title">Coordinadoras que pasaron a Improductiva por marca</div>',
        unsafe_allow_html=True
    )
    if tabla_imp.empty:
        st.markdown(
            '<div class="cuadro-marca-vacio">Sin movimientos hacia Improductiva.</div>',
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            '<div class="tabla-marca-responsiva">'
            + tabla_imp.to_html(
                index=False,
                border=0,
                classes="tabla-marca-compacta",
            )
            + '</div>',
            unsafe_allow_html=True,
        )
        boton_descargar_xlsx(
            tabla_imp,
            "Descargar improductivas por marca XLSX",
            "improductivas_por_marca.xlsx",
            key="descargar_improductivas_marca_xlsx"
        )


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


def limitar_ultimas_semanas_cobranza(
    evol: pd.DataFrame,
    limite: int = MAX_SEMANAS_VISUALES_COBRANZA,
) -> pd.DataFrame:
    """Ordena por año/semana y conserva los últimos periodos para los visuales."""
    if evol is None or evol.empty or "Semana del año" not in evol.columns:
        return pd.DataFrame() if evol is None else evol.copy()

    df_tmp = evol.copy()
    df_tmp["Semana del año"] = pd.to_numeric(
        df_tmp["Semana del año"], errors="coerce"
    )
    columnas_orden = ["Semana del año"]

    if "Año" in df_tmp.columns:
        df_tmp["Año"] = pd.to_numeric(df_tmp["Año"], errors="coerce")
        columnas_orden = ["Año", "Semana del año"]

    df_tmp = df_tmp.dropna(subset=columnas_orden)
    return (
        df_tmp
        .sort_values(columnas_orden)
        .tail(max(1, int(limite)))
        .reset_index(drop=True)
    )


def grafica_cumplimiento(evol, col_cump, modo_moneda=None):
    df_tmp = limitar_ultimas_semanas_cobranza(evol)

    if df_tmp.empty:
        return go.Figure()

    # Asegura que se grafique TODO el histórico que llegue a evol_cobranza.
    # Gris = % de Cumplimiento real.
    # Azul = complemento o diferencia necesaria para llegar al % de cumplimiento de la mejor semana.
    df_tmp[col_cump] = pd.to_numeric(df_tmp[col_cump], errors="coerce")
    mejor_cump = df_tmp[col_cump].max()
    df_tmp["Dif % Mejor Semana"] = (mejor_cump - df_tmp[col_cump]).clip(lower=0)
    df_tmp["Texto Dif % Mejor Semana"] = df_tmp["Dif % Mejor Semana"].apply(
        lambda x: formato_pct(x, 2, False) if pd.notna(x) and x > 0 else ""
    )

    # Detecta la mejor semana para resaltarla con contorno dorado.
    idx_mejor = df_tmp[col_cump].idxmax() if df_tmp[col_cump].notna().any() else None
    colores_borde = ["rgba(0,0,0,0)" for _ in range(len(df_tmp))]
    anchos_borde = [0 for _ in range(len(df_tmp))]

    if idx_mejor is not None and pd.notna(idx_mejor):
        posicion_mejor = int(idx_mejor)
        colores_borde[posicion_mejor] = "#d9c322"
        anchos_borde[posicion_mejor] = 4

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=df_tmp["Etiqueta semana"],
            y=df_tmp[col_cump],
            name="% Cumplimiento",
            marker=dict(
                color="#d9d9d9",
                line=dict(color=colores_borde, width=anchos_borde)
            ),
            text=df_tmp[col_cump].apply(lambda x: formato_pct(x, 2, False)),
            textposition="inside",
            textangle=0,
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
            marker=dict(
                color="#1f77b4",
                line=dict(color=colores_borde, width=anchos_borde)
            ),
            text=df_tmp["Texto Dif % Mejor Semana"],
            # La diferencia contra la mejor semana se muestra encima de la barra.
            # Así no se pierde cuando el segmento azul es muy pequeño.
            textposition="outside",
            textangle=0,
            textfont=dict(color="#082567", size=11),
            cliponaxis=False,
            hovertemplate=(
                "<b>Semana:</b> %{x}<br>"
                "<b>Complemento vs mejor:</b> %{y:.2%}<br>"
                f"<b>Mejor semana:</b> {mejor_cump:.2%}"
                "<extra></extra>"
            )
        )
    )

    if idx_mejor is not None and pd.notna(idx_mejor):
        fila_mejor = df_tmp.loc[idx_mejor]
        fig.add_annotation(
            x=fila_mejor["Etiqueta semana"],
            y=mejor_cump,
            text="<b>Mejor semana</b>",
            showarrow=True,
            arrowhead=2,
            arrowsize=1,
            arrowwidth=1.5,
            arrowcolor="#d9c322",
            yshift=28,
            font=dict(color="#082567", size=12),
            bgcolor="rgba(255,255,255,0.92)",
            bordercolor="#d9c322",
            borderwidth=1,
            borderpad=4
        )

    y_max = max(1, mejor_cump * 1.22 if pd.notna(mejor_cump) else 1)

    fig.update_layout(
        height=420,
        barmode="stack",
        bargap=0.25,
        yaxis_tickformat=".0%",
        yaxis_range=[0, y_max],
        xaxis_title=None,
        yaxis_title="% Cumplimiento y Dif % Mejor Semana",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        margin=dict(t=62, b=85, l=45, r=20),
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

def _siguiente_semana_etiqueta(fila_ultima):
    semana = int(fila_ultima["Semana del año"])
    anio = int(fila_ultima["Año"]) if "Año" in fila_ultima.index and pd.notna(fila_ultima["Año"]) else None

    if semana >= 52:
        semana_sig = 1
        anio_sig = anio + 1 if anio is not None else None
    else:
        semana_sig = semana + 1
        anio_sig = anio

    if anio_sig is not None:
        return f"{anio_sig} S{semana_sig} Pronóstico", semana_sig, anio_sig

    return f"S{semana_sig} Pronóstico", semana_sig, None


def _pronostico_siguiente_semana(serie):
    """
    Pronóstico simple y estable para tablero ejecutivo:
    último valor + promedio de las últimas variaciones disponibles.
    Si hay muy poco histórico, conserva el último dato.
    """
    s = pd.to_numeric(serie, errors="coerce").dropna().astype(float)

    if s.empty:
        return np.nan

    ultimo = float(s.iloc[-1])

    if len(s) < 3:
        return max(0, ultimo)

    difs = s.diff().dropna().tail(3)

    if difs.empty:
        return max(0, ultimo)

    pronostico = ultimo + float(difs.mean())
    return max(0, pronostico)


def _html_recuadro_pronostico(ultima, pronostico_cuota, pronostico_pago, col_cuota, col_pago, moneda_nombre):
    cumplimiento_ult = np.nan
    cumplimiento_proy = np.nan

    if float(ultima[col_cuota]) != 0:
        cumplimiento_ult = float(ultima[col_pago]) / float(ultima[col_cuota])

    if float(pronostico_cuota) != 0:
        cumplimiento_proy = float(pronostico_pago) / float(pronostico_cuota)

    return (
        "<b>Última semana</b><br>"
        f"Semana: {ultima['Etiqueta semana']}<br>"
        f"Cuota: {formato_numero(ultima[col_cuota])}<br>"
        f"Pago: {formato_numero(ultima[col_pago])}<br>"
        f"Cumplimiento: {formato_pct(cumplimiento_ult, 2, False)}<br><br>"
        "<b>Pronóstico próxima semana</b><br>"
        f"Cuota: {formato_numero(pronostico_cuota)}<br>"
        f"Pago: {formato_numero(pronostico_pago)}<br>"
        f"Cumplimiento: {formato_pct(cumplimiento_proy, 2, False)}<br><br>"
        f"<b>Moneda:</b> {moneda_nombre}"
    )


def grafica_cuota_pago(evol, col_cuota, col_pago, col_mejor, col_peor, modo_moneda="Moneda local"):
    df_tmp = limitar_ultimas_semanas_cobranza(evol)

    if df_tmp.empty:
        return go.Figure()

    # Asegura numéricos
    for col in [col_cuota, col_pago, col_mejor, col_peor]:
        if col and col in df_tmp.columns:
            df_tmp[col] = pd.to_numeric(df_tmp[col], errors="coerce").fillna(0)

    moneda_nombre = etiqueta_moneda(modo_moneda)

    # Pronóstico de la siguiente semana.
    ultima = df_tmp.iloc[-1].copy()
    etiqueta_pronostico, semana_sig, anio_sig = _siguiente_semana_etiqueta(ultima)
    pronostico_cuota = _pronostico_siguiente_semana(df_tmp[col_cuota])
    pronostico_pago = _pronostico_siguiente_semana(df_tmp[col_pago])

    fila_pronostico = ultima.copy()
    fila_pronostico["Etiqueta semana"] = etiqueta_pronostico
    fila_pronostico["Semana del año"] = semana_sig
    if "Año" in df_tmp.columns and anio_sig is not None:
        fila_pronostico["Año"] = anio_sig
    fila_pronostico[col_cuota] = pronostico_cuota
    fila_pronostico[col_pago] = pronostico_pago

    df_plot = pd.concat([df_tmp, pd.DataFrame([fila_pronostico])], ignore_index=True)
    df_plot["_posicion_x"] = range(len(df_plot))
    df_plot["_tipo_dato"] = "Real"
    df_plot.loc[df_plot.index[-1], "_tipo_dato"] = "Pronóstico"

    # Valores de referencia con datos reales, no con el pronóstico.
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

    valores_reales = pd.concat([
        df_plot[col_cuota],
        df_plot[col_pago]
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
            margen_sup = rango * 0.34

        y_min = max(0, y_min_real - margen_inf)
        y_max = y_max_real + margen_sup

    if mejor_valor > 0:
        y_max = max(y_max, mejor_valor * 1.05)

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=df_plot["_posicion_x"],
            y=[mejor_valor] * len(df_plot),
            mode="lines",
            name="Mejor Semana",
            line=dict(color="#ff7f0e", width=2, dash="dot"),
            hovertemplate=f"<b>Mejor Semana:</b> %{{y:,.0f}}<br><b>Moneda:</b> {moneda_nombre}<extra></extra>"
        )
    )

    # Tramo real de cuota
    fig.add_trace(
        go.Scatter(
            x=df_tmp.index,
            y=df_tmp[col_cuota],
            mode="lines+markers+text",
            name="Cuota total",
            line=dict(color="black", width=2.5),
            marker=dict(color="black", size=6),
            text=df_tmp[col_cuota].apply(formato_millones),
            textposition="top center",
            textfont=dict(color="#1d4ed8", size=10),
            cliponaxis=False,
            hovertemplate=(
                "<b>Semana:</b> %{customdata}<br>"
                "<b>Cuota total:</b> %{y:,.0f}<br>"
                f"<b>Moneda:</b> {moneda_nombre}"
                "<extra></extra>"
            ),
            customdata=df_tmp["Etiqueta semana"]
        )
    )

    # Tramo pronosticado de cuota: une última semana real con pronóstico.
    fig.add_trace(
        go.Scatter(
            x=[len(df_tmp) - 1, len(df_tmp)],
            y=[df_tmp.iloc[-1][col_cuota], pronostico_cuota],
            mode="lines+markers+text",
            name="Cuota pronóstico",
            line=dict(color="black", width=2.5, dash="dash"),
            marker=dict(color="black", size=[6, 9], symbol=["circle", "diamond"]),
            text=["", formato_millones(pronostico_cuota)],
            textposition="top center",
            textfont=dict(color="#1d4ed8", size=10),
            cliponaxis=False,
            hovertemplate=(
                "<b>Semana:</b> %{customdata}<br>"
                "<b>Cuota total:</b> %{y:,.0f}<br>"
                f"<b>Moneda:</b> {moneda_nombre}"
                "<extra></extra>"
            ),
            customdata=[df_tmp.iloc[-1]["Etiqueta semana"], etiqueta_pronostico]
        )
    )

    # Tramo real de pago
    fig.add_trace(
        go.Scatter(
            x=df_tmp.index,
            y=df_tmp[col_pago],
            mode="lines+markers+text",
            name="Pago total",
            line=dict(color="#00b050", width=2.5),
            marker=dict(color="#00b050", size=6),
            text=df_tmp[col_pago].apply(formato_millones),
            textposition="bottom center",
            textfont=dict(color="#00b050", size=10),
            cliponaxis=False,
            hovertemplate=(
                "<b>Semana:</b> %{customdata}<br>"
                "<b>Pago total:</b> %{y:,.0f}<br>"
                f"<b>Moneda:</b> {moneda_nombre}"
                "<extra></extra>"
            ),
            customdata=df_tmp["Etiqueta semana"]
        )
    )

    # Tramo pronosticado de pago: une última semana real con pronóstico.
    fig.add_trace(
        go.Scatter(
            x=[len(df_tmp) - 1, len(df_tmp)],
            y=[df_tmp.iloc[-1][col_pago], pronostico_pago],
            mode="lines+markers+text",
            name="Pago pronóstico",
            line=dict(color="#00b050", width=2.5, dash="dash"),
            marker=dict(color="#00b050", size=[6, 9], symbol=["circle", "diamond"]),
            text=["", formato_millones(pronostico_pago)],
            textposition="bottom center",
            textfont=dict(color="#00b050", size=10),
            cliponaxis=False,
            hovertemplate=(
                "<b>Semana:</b> %{customdata}<br>"
                "<b>Pago total:</b> %{y:,.0f}<br>"
                f"<b>Moneda:</b> {moneda_nombre}"
                "<extra></extra>"
            ),
            customdata=[df_tmp.iloc[-1]["Etiqueta semana"], etiqueta_pronostico]
        )
    )

    fig.add_trace(
        go.Scatter(
            x=df_plot["_posicion_x"],
            y=[peor_valor] * len(df_plot),
            mode="lines",
            name="Peor Semana",
            line=dict(color="red", width=2, dash="dot"),
            hovertemplate=f"<b>Peor Semana:</b> %{{y:,.0f}}<br><b>Moneda:</b> {moneda_nombre}<extra></extra>"
        )
    )

    # Línea vertical divisoria entre última semana real y pronóstico.
    posicion_division = len(df_tmp) - 0.5
    fig.add_shape(
        type="line",
        x0=posicion_division,
        x1=posicion_division,
        y0=y_min,
        y1=y_max,
        line=dict(color="#082567", width=2, dash="dash")
    )

    fig.add_annotation(
        x=posicion_division,
        y=y_max,
        text="<b>Pronóstico</b>",
        showarrow=False,
        yshift=14,
        font=dict(color="#082567", size=12),
        bgcolor="rgba(255,255,255,0.92)",
        bordercolor="#082567",
        borderwidth=1,
        borderpad=4
    )

    # Recuadro derecho con última semana y pronóstico.
    fig.add_annotation(
        xref="paper",
        yref="paper",
        x=1.01,
        y=0.96,
        xanchor="left",
        yanchor="top",
        align="left",
        text=_html_recuadro_pronostico(
            ultima=ultima,
            pronostico_cuota=pronostico_cuota,
            pronostico_pago=pronostico_pago,
            col_cuota=col_cuota,
            col_pago=col_pago,
            moneda_nombre=moneda_nombre
        ),
        showarrow=False,
        font=dict(color="#082567", size=11),
        bgcolor="rgba(255,255,255,0.96)",
        bordercolor="#d9c322",
        borderwidth=2,
        borderpad=10
    )

    fig.update_layout(
        height=540,
        xaxis_title=None,
        yaxis_title=f"Monto ({moneda_nombre})",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0
        ),
        margin=dict(t=55, b=85, l=65, r=285),
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,1)",
        font=dict(color="#082567", size=11),
        hovermode="x unified",
        dragmode=False,
        xaxis=dict(
            gridcolor="rgba(148,163,184,0.25)",
            tickangle=-45,
            tickmode="array",
            tickvals=df_plot["_posicion_x"].tolist(),
            ticktext=df_plot["Etiqueta semana"].tolist(),
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
    if nombre == "ip":
        try:
            return "" if pd.isna(valor) else f"{float(valor):+,.1f} pp"
        except Exception:
            return str(valor)
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
            f"{r['Indicador']} ({_formato_valor_resumen_ia(r['Indicador'], r['Variación vs sem ant'])})"
            for _, r in top_avances.iterrows()
        ])
        texto.append(
            f"Lo más favorable se observa en {avances_txt}, lo que indica un avance operativo y/o financiero frente a la semana previa. "
        )

    if not negativos.empty:
        top_oportunidades = negativos.sort_values("Variación vs sem ant", ascending=True).head(3)
        oportunidades_txt = ", ".join([
            f"{r['Indicador']} ({_formato_valor_resumen_ia(r['Indicador'], r['Variación vs sem ant'])})"
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

    df_cob_filtrada = aplicar_filtros_cobranza_desde_cartera(
        df_cobranza_base=df_cobranza_base,
        df_cartera_base=df_cartera_base,
        filtros=filtros_aplicados,
    )

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

    if "Año" in df_cob.columns:
        df_cob["Año"] = pd.to_numeric(df_cob["Año"], errors="coerce")
        df_cob = df_cob.dropna(subset=["Año"])
        if df_cob.empty:
            return pd.DataFrame(), None, None

        periodos_disponibles = sorted({
            (int(anio), int(semana))
            for anio, semana in zip(df_cob["Año"], df_cob["Semana del año"])
        })
        periodo_actual = periodos_disponibles[-1]
    else:
        semanas_disponibles = sorted(
            df_cob["Semana del año"].astype(int).unique().tolist()
        )
        if semana_referencia is not None:
            semanas_hasta_ref = [
                s for s in semanas_disponibles if s <= int(semana_referencia)
            ]
            semana_actual = semanas_hasta_ref[-1] if semanas_hasta_ref else semanas_disponibles[-1]
        else:
            semana_actual = semanas_disponibles[-1]
        periodos_disponibles = [(None, semana) for semana in semanas_disponibles]
        periodo_actual = (None, semana_actual)

    indice_actual = periodos_disponibles.index(periodo_actual)
    periodo_anterior = periodos_disponibles[indice_actual - 1] if indice_actual > 0 else None
    semana_actual_cob = periodo_actual[1]
    semana_anterior_cob = periodo_anterior[1] if periodo_anterior is not None else None

    def _agregar_periodo(periodo):
        anio, semana = periodo
        df_sem = df_cob[df_cob["Semana del año"].astype(int) == int(semana)].copy()
        if anio is not None and "Año" in df_sem.columns:
            df_sem = df_sem[df_sem["Año"].astype(int) == int(anio)].copy()
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

    actual = _agregar_periodo(periodo_actual)
    anterior = _agregar_periodo(periodo_anterior) if periodo_anterior is not None else {}

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


def altura_tabla_resumen(df_tabla: pd.DataFrame, alto_fila: int = 36, alto_header: int = 42, max_height: int = 340) -> int:
    """Altura compacta para tablas del resumen; evita espacios en blanco al final."""
    try:
        n = len(df_tabla) if df_tabla is not None else 0
    except Exception:
        n = 0
    return int(min(max_height, max(120, alto_header + (n + 1) * alto_fila)))


def generar_parrafo_resumen_cobranza(resumen_cobranza: pd.DataFrame | None, semana_actual_cobranza: int | None, semana_anterior_cobranza: int | None) -> str:
    """Construye una lectura ejecutiva corta con los datos visibles de Cobranza."""
    if resumen_cobranza is None or resumen_cobranza.empty or semana_actual_cobranza is None:
        return ""

    col_dato = f"Dato sem {semana_actual_cobranza}"
    if col_dato not in resumen_cobranza.columns:
        return ""

    def valor_indicador(nombre: str):
        fila = resumen_cobranza[resumen_cobranza["Indicador"].astype(str).str.lower() == nombre.lower()]
        if fila.empty:
            return None, None
        valor = fila.iloc[0].get(col_dato, np.nan)
        variacion = fila.iloc[0].get("Variación vs sem ant", np.nan)
        return valor, variacion

    cuota, var_cuota = valor_indicador("Cuota Total Cobranza")
    pago, var_pago = valor_indicador("Recuperación semana")
    cumplimiento, var_cump = valor_indicador("% de Cumplimiento")

    partes = []
    semana_ant_txt = semana_anterior_cobranza if semana_anterior_cobranza is not None else "la semana anterior"

    if cuota is not None and pago is not None:
        try:
            brecha = float(cuota) - float(pago)
            partes.append(
                f"En Cobranza, para la semana {semana_actual_cobranza}, la cuota total es de {formato_numero(cuota)} "
                f"y la recuperación semanal alcanza {formato_numero(pago)}, con una brecha pendiente de {formato_numero(brecha)}."
            )
        except Exception:
            pass

    if cumplimiento is not None and pd.notna(cumplimiento):
        partes.append(f"El cumplimiento de cobranza se ubica en {formato_pct(cumplimiento, 2, False)}.")

    movimientos = []
    if var_cuota is not None and pd.notna(var_cuota):
        movimientos.append(f"la cuota varía {formato_variacion(var_cuota)}")
    if var_pago is not None and pd.notna(var_pago):
        movimientos.append(f"la recuperación varía {formato_variacion(var_pago)}")
    if var_cump is not None and pd.notna(var_cump):
        movimientos.append(f"el cumplimiento cambia {formato_pct(var_cump, 2, True)}")

    if movimientos:
        partes.append(f"Contra la semana {semana_ant_txt}, " + ", ".join(movimientos) + ".")

    return " ".join(partes)


def generar_powerpoint_analisis(
    resumen: pd.DataFrame,
    df_cartera: pd.DataFrame,
    semana_actual: int,
    semana_anterior,
    comentario_resumen: str,
    unidad: str,
    modo_moneda: str,
    secciones: list[str],
    incluir_comentarios: bool = True,
) -> bytes | None:
    """Construye una presentación 16:9 editable con la información filtrada."""
    if not PPTX_DISPONIBLE:
        return None

    azul = RGBColor(8, 37, 103)
    amarillo = RGBColor(240, 207, 44)
    gris = RGBColor(100, 116, 139)
    gris_claro = RGBColor(226, 232, 240)
    verde = RGBColor(31, 157, 96)
    rojo = RGBColor(214, 69, 65)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    def agregar_texto(slide, texto, x, y, ancho, alto, tamano=18, negrita=False, color=azul, alineacion=PP_ALIGN.LEFT):
        caja = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(ancho), Inches(alto))
        marco = caja.text_frame
        marco.clear()
        marco.word_wrap = True
        parrafo = marco.paragraphs[0]
        parrafo.alignment = alineacion
        corrida = parrafo.add_run()
        corrida.text = str(texto)
        corrida.font.name = "Arial"
        corrida.font.size = Pt(tamano)
        corrida.font.bold = negrita
        corrida.font.color.rgb = color
        return caja

    def agregar_encabezado(slide, titulo, subtitulo=""):
        agregar_texto(slide, titulo, .55, .28, 8.9, .55, 26, True, azul)
        linea = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(.55), Inches(.91), Inches(1.05), Inches(.05))
        linea.fill.solid()
        linea.fill.fore_color.rgb = amarillo
        linea.line.fill.background()
        agregar_texto(slide, subtitulo, 9.1, .33, 3.65, .34, 10, False, gris, PP_ALIGN.RIGHT)

    def aplicar_estilo_tabla(tabla, tamano=10):
        for fila_idx, fila in enumerate(tabla.rows):
            for celda in fila.cells:
                celda.margin_left = Inches(.05)
                celda.margin_right = Inches(.05)
                celda.margin_top = Inches(.02)
                celda.margin_bottom = Inches(.02)
                celda.fill.solid()
                celda.fill.fore_color.rgb = azul if fila_idx == 0 else RGBColor(246, 249, 253)
                for parrafo in celda.text_frame.paragraphs:
                    for corrida in parrafo.runs:
                        corrida.font.name = "Arial"
                        corrida.font.size = Pt(tamano)
                        corrida.font.bold = fila_idx == 0
                        corrida.font.color.rgb = RGBColor(255, 255, 255) if fila_idx == 0 else azul

    unidad_texto = str(unidad or "Todas las unidades")
    comparativo_texto = f"Semana {semana_actual} vs. semana {semana_anterior}" if semana_anterior is not None else f"Semana {semana_actual}"

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    agregar_encabezado(slide, "Análisis semanal", f"{unidad_texto} · {comparativo_texto}")
    agregar_texto(slide, "Lectura ejecutiva", .65, 1.17, 4.4, .42, 18, True, azul)
    comentario_ppt = str(comentario_resumen or "Sin comentario disponible.")
    if not incluir_comentarios:
        comentario_ppt = "La presentación se generó sin comentarios ejecutivos."
    agregar_texto(slide, comentario_ppt, .65, 1.68, 5.05, 4.8, 14, False, gris)

    metricas_portada = []
    if resumen is not None and not resumen.empty:
        for indicador in ["Clientes Totales", "Clientes al corriente", "Cartera Total", "Saldo en atraso"]:
            fila = resumen[resumen["Indicador"].astype(str) == indicador]
            if not fila.empty:
                metricas_portada.append((
                    indicador,
                    fila.iloc[0].get(f"Dato sem {semana_actual}", 0),
                    fila.iloc[0].get("Variación vs sem ant", np.nan),
                ))

    for indice, (indicador, valor, variacion) in enumerate(metricas_portada[:4]):
        columna = indice % 2
        fila_tarjeta = indice // 2
        x = 6.08 + columna * 3.15
        y = 1.55 + fila_tarjeta * 2.25
        tarjeta = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(2.85), Inches(1.78))
        tarjeta.fill.solid()
        tarjeta.fill.fore_color.rgb = RGBColor(238, 245, 253)
        tarjeta.line.color.rgb = gris_claro
        agregar_texto(slide, indicador, x + .18, y + .18, 2.45, .3, 11, True, azul)
        agregar_texto(slide, formato_numero(valor), x + .18, y + .58, 2.45, .52, 23, True, azul)
        if pd.notna(variacion):
            color_var = verde if float(variacion) >= 0 else rojo
            agregar_texto(slide, formato_variacion(variacion), x + .18, y + 1.23, 2.45, .28, 11, True, color_var)

    if "KPIs" in secciones and resumen is not None and not resumen.empty:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        agregar_encabezado(slide, "KPIs", f"{unidad_texto} · {comparativo_texto}")
        filas_kpi = []
        for _, fila in resumen.head(10).iterrows():
            filas_kpi.append([
                str(fila.get("Indicador", "")),
                formato_numero(fila.get(f"Dato sem {semana_actual}", 0)),
                formato_variacion(fila.get("Variación vs sem ant", 0)),
            ])
        forma = slide.shapes.add_table(len(filas_kpi) + 1, 3, Inches(.75), Inches(1.3), Inches(11.8), Inches(5.55))
        tabla = forma.table
        for columna, encabezado in enumerate(["Indicador", f"Semana {semana_actual}", "Variación"]):
            tabla.cell(0, columna).text = encabezado
        for fila_idx, valores in enumerate(filas_kpi, start=1):
            for columna, valor in enumerate(valores):
                tabla.cell(fila_idx, columna).text = str(valor)
        aplicar_estilo_tabla(tabla, 11)

    if "Evolución semanal" in secciones and "Semana del año" in df_cartera.columns:
        indicador_evolucion = "Cartera Total" if "Cartera Total" in df_cartera.columns else None
        if indicador_evolucion is None:
            candidatos = [c for c in INDICADORES_BASE if c in df_cartera.columns]
            indicador_evolucion = candidatos[0] if candidatos else None
        if indicador_evolucion:
            evolucion = df_cartera.groupby("Semana del año", dropna=False)[indicador_evolucion].sum().reset_index().sort_values("Semana del año").tail(8)
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            agregar_encabezado(slide, "Evolución semanal", f"{indicador_evolucion} · {etiqueta_moneda(modo_moneda)}")
            datos = ChartData()
            datos.categories = [f"S{int(x)}" for x in evolucion["Semana del año"].tolist()]
            datos.add_series(indicador_evolucion, [float(x) for x in evolucion[indicador_evolucion].tolist()])
            grafica = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(.8), Inches(1.35), Inches(11.7), Inches(5.25), datos).chart
            grafica.has_legend = False
            grafica.value_axis.has_major_gridlines = True
            grafica.value_axis.major_gridlines.format.line.color.rgb = gris_claro

    if "Tipo de coordinadora" in secciones and {"Semana del año", "Tipo Coordinadora"}.issubset(df_cartera.columns):
        distribucion = df_cartera[df_cartera["Semana del año"] == semana_actual].groupby("Tipo Coordinadora", dropna=False).size().reset_index(name="Coordinadoras").sort_values("Coordinadoras", ascending=False)
        if not distribucion.empty:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            agregar_encabezado(slide, "Tipo de coordinadora", f"Semana {semana_actual}")
            datos = ChartData()
            datos.categories = distribucion["Tipo Coordinadora"].astype(str).tolist()
            datos.add_series("Coordinadoras", distribucion["Coordinadoras"].astype(float).tolist())
            grafica = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(.8), Inches(1.25), Inches(7.0), Inches(5.5), datos).chart
            grafica.has_legend = True
            grafica.legend.position = XL_LEGEND_POSITION.RIGHT
            agregar_texto(slide, formato_numero(distribucion["Coordinadoras"].sum()), 8.45, 2.05, 3.5, .75, 31, True, azul)
            agregar_texto(slide, "Coordinadoras registradas", 8.45, 2.82, 3.5, .45, 15, False, gris)

    if "Matriz de movimientos" in secciones and "Semana del año" in df_cartera.columns:
        semanas = sorted(int(x) for x in df_cartera["Semana del año"].dropna().unique())
        if len(semanas) >= 2:
            _, matriz = matriz_desplazamiento_coordinadoras(df_cartera, semanas[-2], semanas[-1])
            if matriz is not None and not matriz.empty:
                matriz_ppt = matriz.copy().iloc[:7, :7]
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                agregar_encabezado(slide, "Matriz de movimientos", f"Semana {semanas[-2]} vs. semana {semanas[-1]}")
                forma = slide.shapes.add_table(len(matriz_ppt.index) + 1, len(matriz_ppt.columns) + 1, Inches(.65), Inches(1.25), Inches(12.0), Inches(5.7))
                tabla = forma.table
                tabla.cell(0, 0).text = "Origen / destino"
                for columna, etiqueta in enumerate(matriz_ppt.columns, start=1):
                    tabla.cell(0, columna).text = str(etiqueta)
                for fila_idx, (indice, fila) in enumerate(matriz_ppt.iterrows(), start=1):
                    tabla.cell(fila_idx, 0).text = str(indice)
                    for columna, valor in enumerate(fila.tolist(), start=1):
                        tabla.cell(fila_idx, columna).text = formato_numero(valor)
                aplicar_estilo_tabla(tabla, 9)

    if "Top / Bottom" in secciones and "Semana del año" in df_cartera.columns:
        nivel_top = next((c for c in ["País", "Sucursal", "Zona", "Ruta"] if c in df_cartera.columns), None)
        indicador_top = "Cartera Total" if "Cartera Total" in df_cartera.columns else None
        if nivel_top and indicador_top:
            ranking = df_cartera[df_cartera["Semana del año"] == semana_actual].groupby(nivel_top, dropna=False)[indicador_top].sum().reset_index().sort_values(indicador_top, ascending=False)
            if not ranking.empty:
                seleccion_ranking = pd.concat([ranking.head(5), ranking.tail(5)]).drop_duplicates(subset=[nivel_top])
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                agregar_encabezado(slide, "Top / Bottom", f"{indicador_top} por {nivel_top}")
                datos = ChartData()
                datos.categories = seleccion_ranking[nivel_top].astype(str).tolist()
                datos.add_series(indicador_top, seleccion_ranking[indicador_top].astype(float).tolist())
                grafica = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(.9), Inches(1.25), Inches(11.6), Inches(5.65), datos).chart
                grafica.has_legend = False

    if "Conclusiones" in secciones:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        agregar_encabezado(slide, "Conclusiones", f"{unidad_texto} · semana {semana_actual}")
        agregar_texto(slide, str(comentario_resumen or "Sin conclusiones disponibles."), .9, 1.35, 11.3, 4.9, 18, False, azul)
        agregar_texto(slide, "Fuente: tablero de Productividad y Cobranza", .9, 6.65, 11.3, .3, 10, False, gris)

    salida = BytesIO()
    prs.save(salida)
    salida.seek(0)
    return salida.getvalue()


def abrir_modal_resumen_pais(
    resumen: pd.DataFrame,
    semana_actual: int,
    semana_anterior,
    comentario_resumen: str,
    modo_moneda: str,
    filtros_aplicados: dict,
    df_cartera: pd.DataFrame | None = None,
    unidad: str = "",
    resumen_cobranza: pd.DataFrame | None = None,
    semana_actual_cobranza: int | None = None,
    semana_anterior_cobranza: int | None = None,
):
    """Renderiza la tercera etapa: Análisis y generación de PowerPoint."""
    filtros_visibles = []
    for col, val in filtros_aplicados.items():
        if val:
            valores = ", ".join([str(x) for x in val])
            filtros_visibles.append(f"{col}: {valores}")

    filtros_texto = " · ".join(filtros_visibles) if filtros_visibles else "Todos los países y marcas"
    semana_anterior_txt = semana_anterior if semana_anterior is not None else "sin semana anterior"
    comentario_resumen_completo = str(comentario_resumen)
    parrafo_cobranza = generar_parrafo_resumen_cobranza(
        resumen_cobranza=resumen_cobranza,
        semana_actual_cobranza=semana_actual_cobranza,
        semana_anterior_cobranza=semana_anterior_cobranza,
    )
    if parrafo_cobranza:
        comentario_resumen_completo = comentario_resumen_completo + "\n\n" + parrafo_cobranza

    metricas_modal = []
    if resumen is not None and not resumen.empty:
        metricas_modal = [
            c for c in ["Clientes Totales", "Clientes al corriente", "Cartera Total", "Saldo en atraso"]
            if c in resumen["Indicador"].astype(str).tolist()
        ]

    tarjetas_html = []
    for indicador in metricas_modal[:4]:
        fila = resumen[resumen["Indicador"] == indicador].iloc[0]
        valor = fila.get(f"Dato sem {semana_actual}", 0)
        variacion = fila.get("Variación vs sem ant", np.nan) if semana_anterior is not None else np.nan
        clase_var = "positiva" if pd.notna(variacion) and float(variacion) >= 0 else "negativa"
        texto_var = formato_variacion(variacion) if pd.notna(variacion) else "Sin comparativo"
        tarjetas_html.append(
            '<div class="analisis-kpi">'
            f'<span>{html.escape(indicador)}</span>'
            f'<strong>{formato_numero(valor)}</strong>'
            f'<small class="{clase_var}">{texto_var}</small>'
            '</div>'
        )

    comentario_seguro = html.escape(comentario_resumen_completo).replace("\n", "<br><br>")
    unidad_segura = html.escape(str(unidad or "Todas las unidades"))

    st.markdown(
        """
        <style>
        html, body, .stApp, [data-testid="stAppViewContainer"],
        [data-testid="stAppViewContainer"] > .main {
            width:100% !important;
            max-width:100% !important;
            height:100vh !important;
            max-height:100vh !important;
            overflow:hidden !important;
        }
        [data-testid="stAppViewContainer"] .main .block-container {
            width:100% !important;
            max-width:100% !important;
            height:100vh !important;
            max-height:100vh !important;
            overflow:hidden !important;
            box-sizing:border-box !important;
            padding:.35rem .65rem .45rem !important;
        }
        .st-key-filtros_superiores,
        .st-key-ayuda_tablero,
        [data-testid="stElementContainer"]:has(.gestion-encabezado),
        div[data-testid="stMarkdownContainer"]:has(.gestion-encabezado) {
            display:none !important;
            height:0 !important;
            min-height:0 !important;
            margin:0 !important;
            padding:0 !important;
        }
        .st-key-analisis_pagina {
            width:100% !important;
            max-width:100% !important;
            height:calc(100vh - .8rem) !important;
            max-height:calc(100vh - .8rem) !important;
            overflow:hidden !important;
        }
        .st-key-analisis_pagina > div[data-testid="stVerticalBlock"] { gap:.45rem !important; }
        .analisis-encabezado {
            display:flex; align-items:center; justify-content:space-between; gap:20px;
            height:58px; padding:7px 16px; box-sizing:border-box;
            background:#ffffff; border:1px solid #dbe3ee; border-radius:10px;
            box-shadow:0 5px 16px rgba(15,23,42,.06);
        }
        .analisis-marca { color:#082567; font-size:20px; font-weight:950; white-space:nowrap; }
        .analisis-pasos { display:flex; align-items:center; justify-content:flex-end; gap:11px; flex:1; }
        .analisis-paso { display:flex; align-items:center; gap:7px; color:#52698d; font-size:12px; font-weight:800; white-space:nowrap; }
        .analisis-paso i { display:flex; align-items:center; justify-content:center; width:27px; height:27px; border-radius:50%; background:#d9e9fb; color:#082567; font-style:normal; }
        .analisis-paso.activo { color:#082567; }
        .analisis-paso.activo i { background:#e7c42d; color:#ffffff; }
        .analisis-linea { width:45px; height:2px; background:#a9c8ed; }
        .analisis-slide-preview, .st-key-analisis_configuracion {
            height:calc(100vh - 86px); max-height:calc(100vh - 86px);
            overflow:hidden; box-sizing:border-box; background:#ffffff;
            border:1px solid #dbe3ee; border-radius:11px;
            box-shadow:0 7px 20px rgba(15,23,42,.07);
        }
        .analisis-slide-preview { padding:clamp(16px,2vw,28px); }
        .analisis-titulo { color:#082567; font-size:clamp(26px,2.5vw,42px); line-height:1; font-weight:950; margin-bottom:8px; }
        .analisis-subtitulo { color:#64748b; font-size:12px; margin-bottom:15px; }
        .analisis-kpis { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:10px; margin:4px 0 14px; }
        .analisis-kpi { min-width:0; padding:11px 12px; border-radius:9px; background:#eef5fd; border:1px solid #dce8f6; }
        .analisis-kpi span { display:block; color:#587098; font-size:10px; font-weight:800; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
        .analisis-kpi strong { display:block; color:#082567; font-size:clamp(17px,1.6vw,26px); line-height:1.15; margin:5px 0 3px; }
        .analisis-kpi small { font-size:10px; font-weight:900; }
        .analisis-kpi small.positiva { color:#17864f; }
        .analisis-kpi small.negativa { color:#c83d3d; }
        .analisis-lectura-titulo { color:#082567; font-size:16px; font-weight:950; margin:5px 0 7px; }
        .analisis-lectura {
            color:#334b74; font-size:clamp(12px,1.05vw,16px); line-height:1.42;
            max-height:calc(100vh - 335px); overflow:hidden;
            display:-webkit-box; -webkit-line-clamp:12; -webkit-box-orient:vertical;
        }
        .st-key-analisis_configuracion { padding:15px 16px 12px; }
        .analisis-config-encabezado h3 { color:#082567; font-size:20px; margin:0 0 2px; }
        .analisis-config-encabezado p { color:#64748b; font-size:11px; margin:0 0 8px; }
        .st-key-analisis_configuracion > div[data-testid="stVerticalBlock"] { gap:.22rem !important; }
        .st-key-analisis_configuracion [data-testid="stCheckbox"] { margin:0 !important; }
        .st-key-analisis_configuracion [data-testid="stCheckbox"] label { padding:3px 7px !important; border-radius:6px; background:#eef5fd; }
        .st-key-analisis_configuracion [data-testid="stSelectbox"] { margin-top:2px !important; }
        .st-key-analisis_configuracion [data-testid="stDownloadButton"] button { min-height:42px !important; }
        @media (max-width:900px) {
            html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stAppViewContainer"] > .main { overflow-y:auto !important; height:auto !important; max-height:none !important; }
            [data-testid="stAppViewContainer"] .main .block-container, .st-key-analisis_pagina { height:auto !important; max-height:none !important; overflow:visible !important; }
            .analisis-pasos { overflow-x:auto; justify-content:flex-start; }
            .analisis-paso b { display:none; }
            .analisis-slide-preview, .st-key-analisis_configuracion { height:auto; max-height:none; }
            .analisis-kpis { grid-template-columns:repeat(2,minmax(0,1fr)); }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    pagina_analisis = st.container(key="analisis_pagina")
    with pagina_analisis:
        col_volver, col_cabecera = st.columns([1.05, 10.95], gap="small")
        with col_volver:
            if st.button("← Gestión", key="btn_volver_desde_resumen", use_container_width=True):
                volver_al_tablero()
        with col_cabecera:
            st.markdown(
                '<div class="analisis-encabezado">'
                '<div class="analisis-marca">Gestión de Portafolio</div>'
                '<div class="analisis-pasos">'
                '<div class="analisis-paso"><i>✓</i><b>1. Seleccionar unidad</b></div>'
                '<div class="analisis-linea"></div>'
                '<div class="analisis-paso"><i>✓</i><b>2. Gestión</b></div>'
                '<div class="analisis-linea"></div>'
                '<div class="analisis-paso activo"><i>3</i><b>3. Análisis</b></div>'
                '</div></div>',
                unsafe_allow_html=True,
            )

        col_vista, col_config = st.columns([3.55, 1.35], gap="medium")
        with col_vista:
            st.markdown(
                '<div class="analisis-slide-preview">'
                '<div class="analisis-titulo">Análisis semanal</div>'
                f'<div class="analisis-subtitulo">{unidad_segura} · Semana {semana_actual} · Comparativo {semana_anterior_txt} · {html.escape(etiqueta_moneda(modo_moneda))}<br>{html.escape(filtros_texto)}</div>'
                f'<div class="analisis-kpis">{"".join(tarjetas_html)}</div>'
                '<div class="analisis-lectura-titulo">Lectura ejecutiva</div>'
                f'<div class="analisis-lectura">{comentario_seguro}</div>'
                '</div>',
                unsafe_allow_html=True,
            )

        with col_config:
            configuracion = st.container(key="analisis_configuracion")
            with configuracion:
                st.markdown('<div class="analisis-config-encabezado"><h3>Generar PowerPoint</h3><p>Selecciona las secciones que incluirá el archivo.</p></div>', unsafe_allow_html=True)
                opciones_ppt = [
                    "KPIs",
                    "Evolución semanal",
                    "Tipo de coordinadora",
                    "Matriz de movimientos",
                    "Top / Bottom",
                    "Conclusiones",
                ]
                secciones_ppt = [
                    opcion for opcion in opciones_ppt
                    if st.checkbox(opcion, value=True, key=f"ppt_{normalizar_texto_tc(opcion).lower().replace(' ', '_').replace('/', '_')}")
                ]
                st.selectbox("Formato", ["16:9"], key="formato_powerpoint", disabled=True)
                incluir_comentarios = st.toggle("Incluir comentarios", value=True, key="ppt_incluir_comentarios")

                datos_ppt = generar_powerpoint_analisis(
                    resumen=resumen,
                    df_cartera=df_cartera if df_cartera is not None else pd.DataFrame(),
                    semana_actual=semana_actual,
                    semana_anterior=semana_anterior,
                    comentario_resumen=comentario_resumen_completo,
                    unidad=unidad,
                    modo_moneda=modo_moneda,
                    secciones=secciones_ppt,
                    incluir_comentarios=incluir_comentarios,
                )
                if datos_ppt is None:
                    st.error("Falta instalar python-pptx para generar la presentación.")
                else:
                    st.download_button(
                        "Generar PowerPoint",
                        data=datos_ppt,
                        file_name=f"analisis_{normalizar_texto_tc(unidad or 'unidad').lower().replace(' ', '_')}_semana_{semana_actual}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key="descargar_powerpoint_analisis",
                        use_container_width=True,
                    )


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
        ruta_local = st.text_input("Ruta local del archivo o carpeta", value=RUTA_DEFAULT)

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
    if archivo_subido is None and Path(ruta_local).is_dir():
        with st.spinner("Cargando concentrados de Cartera..."):
            firma_cartera = firma_archivos_cartera_concentrados(ruta_local)
            df = cargar_cartera_desde_concentrados(ruta_local, firma_cartera)
        _df_cobranza_excel = None
    else:
        # Se conserva la carga manual de un CSV/Excel para pruebas puntuales.
        df, _df_cobranza_excel = cargar_archivo(ruta_local, archivo_subido)
except Exception as e:
    st.error(str(e))
    st.stop()

# Cobranza es independiente. Si sus Parquet no están disponibles, Cartera
# sigue funcionando y la vista Cobranza muestra el motivo.
error_cobranza = None
try:
    firma_cobranza = firma_archivos_cobranza_parquet(str(RUTA_COBRANZA_PARQUET))
    df_cobranza = cargar_cobranza_desde_parquet(
        str(RUTA_COBRANZA_PARQUET),
        firma_cobranza,
    )
except Exception as e:
    df_cobranza = None
    error_cobranza = str(e)


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
if "IP" in df.columns and "IP" not in indicadores_disponibles:
    indicadores_disponibles.append("IP")
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

    # Las tarjetas de los mapas son enlaces a esta misma aplicación. El
    # parámetro permite que toda la tarjeta sea clicable y abre directamente
    # el tablero correspondiente, sin un botón intermedio.
    unidad_query = st.query_params.get("unidad")
    if unidad_query:
        unidad_desde_mapa = next(
            (
                unidad
                for unidad in unidades_negocio
                if normalizar_texto_tc(unidad) == normalizar_texto_tc(unidad_query)
            ),
            None,
        )
        if unidad_desde_mapa is not None:
            cambio_de_unidad = unidad_guardada != unidad_desde_mapa
            seccion_desde_enlace = str(st.query_params.get("seccion", "kpis"))
            if seccion_desde_enlace not in {
                "kpis",
                "evolucion",
                "coordinadoras",
                "movimientos",
                "top-bottom",
                "conclusiones",
                "cobranza-barras",
                "cobranza-lineas",
            }:
                seccion_desde_enlace = "kpis"
            st.session_state["unidad_negocio_app"] = unidad_desde_mapa
            st.session_state["gestion_seccion_activa"] = seccion_desde_enlace
            if cambio_de_unidad:
                st.session_state["modo_moneda_superior"] = "Moneda local"
                st.session_state.pop("filtro_superior_País", None)
                st.session_state.pop("filtro_superior_Marca", None)
            st.query_params.clear()
            st.rerun()

    if unidad_guardada is None:
        # La pantalla inicial funciona como el primer paso de la secuencia:
        # alcance -> gestión -> análisis.
        unidad_inicio = st.session_state.get("unidad_negocio_inicio")
        if unidad_inicio not in unidades_negocio:
            unidad_inicio = next(
                (
                    unidad
                    for unidad in unidades_negocio
                    if normalizar_texto_tc(unidad) == "PRESICO"
                ),
                unidades_negocio[0],
            )
            st.session_state["unidad_negocio_inicio"] = unidad_inicio

        st.markdown(
            """
            <div class="inicio-hero">
                <div class="inicio-kicker">REPORTE EJECUTIVO SEMANAL</div>
                <div class="inicio-title">Indicadores de Productividad y Cobranza</div>
                <div class="inicio-subtitle">Selecciona la unidad de negocio para abrir el tablero.</div>
            </div>
            <div class="inicio-stepper">
                <div class="inicio-step activo"><span>1</span><b>Seleccionar unidad</b></div>
                <div class="inicio-linea"></div>
                <div class="inicio-step"><span>2</span><b>Gestión</b></div>
                <div class="inicio-linea"></div>
                <div class="inicio-step"><span>3</span><b>Análisis</b></div>
            </div>
            <div class="inicio-section-title">1. Elige la unidad de negocio</div>
            """,
            unsafe_allow_html=True,
        )

        logos_unidad = {
            "PRESICO": "Logo.jpg",
            "PRESICO LATAM": "Presico sin fondo LATAM.jpg",
        }
        mapas_unidad = {
            "PRESICO": "mapa_mexico.png",
            "PRESICO LATAM": "mapa_latam.png",
        }

        cols_unidades = st.columns(min(len(unidades_negocio), 2), gap="large")
        for idx, unidad in enumerate(unidades_negocio):
            unidad_texto = str(unidad).strip()
            unidad_key = normalizar_texto_tc(unidad_texto)
            nombre_logo = logos_unidad.get(unidad_key)
            nombre_mapa = mapas_unidad.get(unidad_key)
            unidad_parametro = unidad_key.replace(" ", "%20")

            mapa_html = (
                imagen_logo_html(nombre_mapa, "mapa-silueta mapa-latam" if "LATAM" in unidad_key else "mapa-silueta mapa-mexico")
                if nombre_mapa
                else '<div class="mapa-placeholder">Mapa</div>'
            )
            logo_html = (
                imagen_logo_html(nombre_logo, "mapa-logo-overlay mapa-logo-latam" if "LATAM" in unidad_key else "mapa-logo-overlay mapa-logo-mexico")
                if nombre_logo
                else '<div class="mapa-logo-texto">PRÉSICO</div>'
            )
            alcance = "México" if unidad_key == "PRESICO" else "LATAM"
            detalle = "Operación nacional" if unidad_key == "PRESICO" else "Operación por país"

            with cols_unidades[idx % len(cols_unidades)]:
                st.markdown(
                    f"""
                    <a class="mapa-card-link" href="?unidad={unidad_parametro}" target="_self">
                        <div class="mapa-card">
                            <div class="mapa-visual">
                                {mapa_html}
                                <div class="mapa-logo-contenedor">{logo_html}</div>
                            </div>
                            <div class="mapa-card-title">{html.escape(alcance)}</div>
                            <div class="mapa-card-subtitle">{html.escape(detalle)}</div>
                            <div class="mapa-card-accion">Abrir tablero →</div>
                        </div>
                    </a>
                    """,
                    unsafe_allow_html=True,
                )

        ruta_calendario_presico = CARPETA_APP / "CALENDARIO PRESICO.png"
        @st.dialog("Calendario Présico 2026", width="large")
        def mostrar_calendario_presico():
            if ruta_calendario_presico.exists():
                calendario_base64 = imagen_a_base64(str(ruta_calendario_presico))
                st.markdown(
                    f"""
                    <div class="calendario-presico-viewport">
                        <img
                            class="calendario-presico-imagen"
                            src="data:image/png;base64,{calendario_base64}"
                            alt="Calendario Présico 2026"
                        />
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            else:
                st.info(
                    "Agrega CALENDARIO PRESICO.png en la misma carpeta de app.py para mostrar el calendario."
                )

        if st.button(
            "📅 Calendario Présico · Cierres mensuales",
            key="btn_calendario_presico",
            use_container_width=True,
        ):
            mostrar_calendario_presico()

        # Estilo autocontenido de la portada. Se inyecta antes de detener la
        # ejecución porque el tablero se abre únicamente al pulsar un mapa.
        ruta_landing_marca_agua = buscar_imagen_fondo(NOMBRE_IMAGEN_FONDO)
        landing_background = ""
        if ruta_landing_marca_agua is not None:
            landing_fondo_base64 = imagen_a_base64(str(ruta_landing_marca_agua))
            landing_background = f'''
                background-image:
                    linear-gradient(rgba(248,250,252,0.94), rgba(248,250,252,0.98)),
                    url("data:image/png;base64,{landing_fondo_base64}") !important;
            '''

        st.markdown(
            f"""
            <style>
            html, body, .stApp, [data-testid="stAppViewContainer"] {{
                color-scheme: light !important;
                background-color:#f8fafc !important;
            }}
            .stApp, [data-testid="stAppViewContainer"] {{
                {landing_background}
                background-size:cover !important;
                background-position:top center !important;
                background-repeat:no-repeat !important;
                background-attachment:fixed !important;
            }}
            .main .block-container,
            [data-testid="stMainBlockContainer"] {{
                max-width:1280px !important;
                padding-top:.75rem !important;
                padding-bottom:1rem !important;
                background:transparent !important;
                border:0 !important;
                box-shadow:none !important;
            }}
            [data-testid="stVerticalBlock"]
            > [data-testid="stElementContainer"]:has(style) {{
                display:none !important;
                height:0 !important;
                margin:0 !important;
                padding:0 !important;
            }}
            .inicio-hero {{ text-align:center; padding:2px 12px 8px; }}
            .inicio-kicker {{
                color:#d48a24; font-size:12px; font-weight:900;
                letter-spacing:2.2px; margin-bottom:5px;
            }}
            .inicio-title {{
                color:#082567; font-size:36px; line-height:1.06;
                font-weight:950; letter-spacing:-.8px;
            }}
            .inicio-subtitle {{
                color:#475569; font-size:15px; font-weight:600; margin-top:7px;
            }}
            .inicio-stepper {{
                display:flex; align-items:center; justify-content:center;
                max-width:760px; margin:10px auto 16px;
            }}
            .inicio-step {{
                min-width:145px; display:flex; flex-direction:column;
                align-items:center; gap:4px; color:#94a3b8; font-size:11px;
            }}
            .inicio-step span {{
                width:29px; height:29px; border-radius:50%; display:flex;
                align-items:center; justify-content:center; background:#e2e8f0;
                color:#64748b; font-weight:900;
            }}
            .inicio-step.activo {{ color:#082567; }}
            .inicio-step.activo span {{ background:#082567; color:#ffffff; }}
            .inicio-linea {{ height:3px; background:#dbe3ee; flex:1; margin-bottom:18px; }}
            .inicio-section-title {{
                color:#082567; font-size:18px; font-weight:900; margin:0 0 8px;
            }}
            .mapa-card-link {{ text-decoration:none !important; color:inherit !important; }}
            .mapa-card {{
                position:relative; background:rgba(255,255,255,.98);
                border:2px solid #e2e8f0; border-radius:22px;
                padding:8px 18px 13px; min-height:245px; text-align:center;
                box-shadow:0 10px 26px rgba(15,23,42,.08);
                transition:transform .16s ease, border-color .16s ease, box-shadow .16s ease;
                cursor:pointer;
            }}
            .mapa-card:hover {{
                transform:translateY(-4px); border-color:#d99932;
                box-shadow:0 16px 34px rgba(8,37,103,.16);
            }}
            .mapa-visual {{
                position:relative; height:160px; display:flex;
                align-items:center; justify-content:center; overflow:hidden;
            }}
            .mapa-silueta {{
                width:100%; height:100%; object-fit:contain; display:block;
                filter:drop-shadow(0 8px 9px rgba(15,23,42,.12));
            }}
            .mapa-latam {{ transform:scale(.91); }}
            .mapa-logo-contenedor {{
                position:absolute; inset:0; display:flex; align-items:center;
                justify-content:center; pointer-events:none;
            }}
            .mapa-logo-overlay {{
                width:142px; height:66px; object-fit:contain; display:block;
                background:rgba(255,255,255,.91); border-radius:13px;
                padding:6px 10px; box-shadow:0 5px 15px rgba(15,23,42,.14);
            }}
            .mapa-logo-latam {{ width:134px; height:72px; }}
            .mapa-logo-contenedor .unidad-logo-placeholder {{
                width:110px; height:64px; margin:0; border-radius:13px;
                background:rgba(255,255,255,.92); font-size:34px;
            }}
            .mapa-card-title {{ color:#082567; font-size:20px; font-weight:950; }}
            .mapa-card-subtitle {{ color:#64748b; font-size:13px; font-weight:650; margin-top:1px; }}
            .mapa-card-accion {{ color:#b56f1f; font-size:12px; font-weight:900; margin-top:5px; }}
            .st-key-btn_calendario_presico button {{
                justify-content:flex-start !important; text-align:left !important;
                background:rgba(255,255,255,.97) !important; color:#082567 !important;
                border:1px solid #dbe3ee !important; border-radius:14px !important;
                min-height:48px !important; padding:0 18px !important;
                font-size:17px !important; font-weight:900 !important;
                box-shadow:0 7px 20px rgba(15,23,42,.06) !important;
            }}
            .st-key-btn_calendario_presico button:hover {{
                border-color:#d99932 !important; color:#082567 !important;
            }}
            .calendario-presico-viewport {{
                width:100%; height:min(70vh, 720px); min-height:340px;
                display:flex; align-items:center; justify-content:center;
                overflow:hidden; background:#ffffff;
            }}
            .calendario-presico-imagen {{
                display:block; width:100%; height:100%; max-width:100%;
                max-height:100%; object-fit:contain;
            }}
            @media (max-width:800px) {{
                .inicio-title {{ font-size:27px; }}
                .inicio-step b {{ display:none; }}
                .inicio-step {{ min-width:40px; }}
                .inicio-linea {{ margin-bottom:0; }}
                .mapa-card {{ min-height:220px; }}
                .mapa-visual {{ height:140px; }}
                .calendario-presico-viewport {{ height:auto; min-height:0; }}
                .calendario-presico-imagen {{ height:auto; }}
            }}
            </style>
            """,
            unsafe_allow_html=True,
        )

        st.stop()

        df_inicio = df[
            df["Unidad de Negocio"].astype(str).str.strip() == str(unidad_inicio).strip()
        ].copy()
        unidad_inicio_norm = normalizar_texto_tc(unidad_inicio)
        es_inicio_latam = "LATAM" in unidad_inicio_norm

        opciones_moneda_inicio = (
            ["Moneda local", "Pesos mexicanos"]
            if es_inicio_latam
            else ["Moneda local"]
        )
        if st.session_state.get("inicio_moneda") not in opciones_moneda_inicio:
            st.session_state["inicio_moneda"] = "Moneda local"

        paises_inicio = []
        if "País" in df_inicio.columns:
            paises_inicio = sorted(
                df_inicio["País"].dropna().astype(str).str.strip().unique().tolist()
            )

        modo_moneda_inicio = st.session_state.get("inicio_moneda", "Moneda local")
        if es_inicio_latam and modo_moneda_inicio == "Moneda local":
            opciones_pais_inicio = paises_inicio or ["Sin datos"]
            pais_default_inicio = (
                "Guatemala"
                if "Guatemala" in opciones_pais_inicio
                else opciones_pais_inicio[0]
            )
        else:
            opciones_pais_inicio = ["Todos"] + paises_inicio
            pais_default_inicio = (
                paises_inicio[0]
                if not es_inicio_latam and len(paises_inicio) == 1
                else "Todos"
            )
        if st.session_state.get("inicio_pais") not in opciones_pais_inicio:
            st.session_state["inicio_pais"] = pais_default_inicio

        pais_inicio_estado = st.session_state.get("inicio_pais", pais_default_inicio)
        df_marcas_inicio = df_inicio
        if (
            pais_inicio_estado not in ["Todos", "Sin datos"]
            and "País" in df_marcas_inicio.columns
        ):
            df_marcas_inicio = df_marcas_inicio[
                df_marcas_inicio["País"].astype(str).str.strip()
                == str(pais_inicio_estado).strip()
            ]
        marcas_inicio = []
        if "Marca" in df_marcas_inicio.columns:
            marcas_inicio = sorted(
                df_marcas_inicio["Marca"].dropna().astype(str).str.strip().unique().tolist()
            )
        opciones_marca_inicio = ["Todos"] + marcas_inicio
        if st.session_state.get("inicio_marca") not in opciones_marca_inicio:
            st.session_state["inicio_marca"] = "Todos"

        semanas_inicio = sorted(
            int(semana)
            for semana in df_inicio["Semana del año"].dropna().unique()
        )
        rango_inicio_default = (
            semanas_inicio[-8] if len(semanas_inicio) >= 8 else semanas_inicio[0],
            semanas_inicio[-1],
        )
        anio_calendario_inicio = 2026
        fecha_minima_inicio = date.fromisocalendar(
            anio_calendario_inicio,
            semanas_inicio[0],
            1,
        )
        fecha_maxima_inicio = date.fromisocalendar(
            anio_calendario_inicio,
            semanas_inicio[-1],
            7,
        )
        fechas_inicio_default = (
            date.fromisocalendar(
                anio_calendario_inicio,
                rango_inicio_default[0],
                1,
            ),
            date.fromisocalendar(
                anio_calendario_inicio,
                rango_inicio_default[1],
                7,
            ),
        )
        rango_fechas_estado = st.session_state.get("inicio_rango_fechas")
        if not (
            isinstance(rango_fechas_estado, (list, tuple))
            and len(rango_fechas_estado) == 2
            and isinstance(rango_fechas_estado[0], date)
            and isinstance(rango_fechas_estado[1], date)
            and fecha_minima_inicio <= rango_fechas_estado[0] <= rango_fechas_estado[1] <= fecha_maxima_inicio
        ):
            rango_fechas_estado = fechas_inicio_default
            st.session_state["inicio_rango_fechas"] = rango_fechas_estado

        st.markdown(
            '<div class="inicio-section-title inicio-filtros-title">2. Configura la vista inicial</div>',
            unsafe_allow_html=True,
        )
        col_pais_inicio, col_marca_inicio, col_moneda_inicio = st.columns(3)

        with col_pais_inicio:
            pais_inicio = st.selectbox(
                "País",
                options=opciones_pais_inicio,
                key="inicio_pais",
            )
        with col_marca_inicio:
            marca_inicio = st.selectbox(
                "Marca",
                options=opciones_marca_inicio,
                key="inicio_marca",
            )
        with col_moneda_inicio:
            moneda_inicio = st.selectbox(
                "Moneda",
                options=opciones_moneda_inicio,
                key="inicio_moneda",
            )

        st.markdown(
            """
            <div class="calendario-inicio-titulo">📅 Rango de fechas · Calendario Présico</div>
            <div class="calendario-inicio-ayuda">
                Arrastra desde el primer día hasta el último día del periodo. El tablero convertirá
                automáticamente la selección al rango de semanas correspondiente.
            </div>
            """,
            unsafe_allow_html=True,
        )

        # Fechas corporativas recreadas a partir del calendario Présico 2026.
        eventos_calendario_inicio = []
        fechas_cierre = [
            "2026-01-31", "2026-02-28", "2026-03-28", "2026-05-02",
            "2026-05-30", "2026-06-27", "2026-08-01", "2026-08-29",
            "2026-09-26", "2026-10-31", "2026-11-28", "2026-12-26",
        ]
        fechas_festivas = [
            "2026-01-01", "2026-03-16", "2026-04-03", "2026-04-04",
            "2026-05-01", "2026-09-16", "2026-11-16", "2026-12-25",
        ]
        fechas_aniversario = {
            "2026-02-02": "Aniv. La Lupita",
            "2026-02-22": "Aniv. La Casita",
            "2026-04-16": "Aniv. La Güerita",
            "2026-10-23": "Aniv. Préstamos Unidos",
            "2026-10-28": "Aniv. Présico",
            "2026-12-05": "Aniv. La Moderna",
        }
        for fecha_evento in fechas_cierre:
            eventos_calendario_inicio.append(
                {
                    "title": "Cierre",
                    "start": fecha_evento,
                    "allDay": True,
                    "backgroundColor": "#34436f",
                    "borderColor": "#34436f",
                    "textColor": "#ffffff",
                }
            )
        for fecha_evento in fechas_festivas:
            eventos_calendario_inicio.append(
                {
                    "title": "Festivo",
                    "start": fecha_evento,
                    "allDay": True,
                    "backgroundColor": "#f5df28",
                    "borderColor": "#f5df28",
                    "textColor": "#5b4a00",
                }
            )
        for fecha_evento, titulo_evento in fechas_aniversario.items():
            eventos_calendario_inicio.append(
                {
                    "title": titulo_evento,
                    "start": fecha_evento,
                    "allDay": True,
                    "backgroundColor": "#f5822a",
                    "borderColor": "#f5822a",
                    "textColor": "#ffffff",
                }
            )
        eventos_calendario_inicio.append(
            {
                "title": "Liberación de listados",
                "start": "2026-03-30",
                "allDay": True,
                "backgroundColor": "#84c35a",
                "borderColor": "#84c35a",
                "textColor": "#173b15",
            }
        )

        # El rango vigente se dibuja como fondo azul para que permanezca visible
        # después del rerun que produce la selección del componente.
        eventos_calendario_inicio.insert(
            0,
            {
                "title": "Periodo seleccionado",
                "start": rango_fechas_estado[0].isoformat(),
                "end": (rango_fechas_estado[1] + timedelta(days=1)).isoformat(),
                "allDay": True,
                "display": "background",
                "backgroundColor": "#bcd7ff",
            },
        )

        if calendario_interactivo is not None:
            opciones_calendario_inicio = {
                "initialView": "multiMonthTwoMonths",
                "initialDate": fecha_minima_inicio.isoformat(),
                "locale": "es",
                "firstDay": 1,
                "selectable": True,
                "selectMirror": True,
                "unselectAuto": False,
                "weekNumbers": True,
                "weekText": "S",
                "showNonCurrentDates": False,
                "fixedWeekCount": False,
                "multiMonthMaxColumns": 2,
                "views": {
                    "multiMonthTwoMonths": {
                        "type": "multiMonth",
                        "duration": {"months": 2},
                    }
                },
                "headerToolbar": {
                    "left": "prev",
                    "center": "title",
                    "right": "next",
                },
                "validRange": {
                    "start": fecha_minima_inicio.isoformat(),
                    "end": (fecha_maxima_inicio + timedelta(days=1)).isoformat(),
                },
                "height": "auto",
            }
            estado_calendario_inicio = calendario_interactivo(
                events=eventos_calendario_inicio,
                options=opciones_calendario_inicio,
                custom_css="""
                    .fc { font-family: Arial, sans-serif; color: #19315d; }
                    .fc .fc-toolbar-title { color: #082567; font-size: 1.25rem; font-weight: 900; }
                    .fc .fc-button-primary { background: #082567; border-color: #082567; }
                    .fc .fc-multimonth { border: 0; }
                    .fc .fc-multimonth-month { border: 1px solid #dbe3ee; border-radius: 14px; padding: 10px; }
                    .fc .fc-multimonth-title { color: #b56f1f; font-weight: 900; text-transform: uppercase; }
                    .fc .fc-col-header-cell-cushion { color: #475569; font-weight: 800; }
                    .fc .fc-daygrid-day-number { color: #1e293b; font-weight: 700; }
                    .fc .fc-day-today { background: rgba(217,153,50,.12) !important; }
                    .fc .fc-highlight { background: rgba(8,37,103,.22) !important; }
                    .fc .fc-event { border-radius: 5px; font-size: .68rem; font-weight: 800; }
                    .fc .fc-daygrid-week-number { background: #d9902f; color: white; font-weight: 900; }
                """,
                callbacks=["select"],
                key="calendario_presico_inicio_v1",
            ) or {}

            if estado_calendario_inicio.get("callback") == "select":
                seleccion_calendario = estado_calendario_inicio.get("select", {})
                inicio_iso = str(seleccion_calendario.get("start", ""))[:10]
                fin_exclusivo_iso = str(seleccion_calendario.get("end", ""))[:10]
                try:
                    nueva_fecha_inicio = date.fromisoformat(inicio_iso)
                    nueva_fecha_fin = date.fromisoformat(fin_exclusivo_iso) - timedelta(days=1)
                except (TypeError, ValueError):
                    nueva_fecha_inicio = None
                    nueva_fecha_fin = None

                if (
                    nueva_fecha_inicio is not None
                    and fecha_minima_inicio <= nueva_fecha_inicio <= nueva_fecha_fin <= fecha_maxima_inicio
                    and (nueva_fecha_inicio, nueva_fecha_fin) != tuple(rango_fechas_estado)
                ):
                    st.session_state["inicio_rango_fechas"] = (
                        nueva_fecha_inicio,
                        nueva_fecha_fin,
                    )
                    st.rerun()
        else:
            # Respaldo nativo: mantiene la aplicación operativa incluso si la
            # dependencia visual todavía no se ha instalado en Streamlit Cloud.
            rango_fechas_nativo = st.date_input(
                "Selecciona la fecha inicial y final",
                value=tuple(rango_fechas_estado),
                min_value=fecha_minima_inicio,
                max_value=fecha_maxima_inicio,
                format="DD/MM/YYYY",
                key="calendario_presico_respaldo",
            )
            if isinstance(rango_fechas_nativo, (list, tuple)) and len(rango_fechas_nativo) == 2:
                st.session_state["inicio_rango_fechas"] = tuple(rango_fechas_nativo)
                rango_fechas_estado = tuple(rango_fechas_nativo)

        rango_fechas_estado = tuple(st.session_state["inicio_rango_fechas"])
        semana_inicio_calendario = int(rango_fechas_estado[0].isocalendar().week)
        semana_fin_calendario = int(rango_fechas_estado[1].isocalendar().week)
        rango_inicio = (
            min(semanas_inicio, key=lambda semana: abs(semana - semana_inicio_calendario)),
            min(semanas_inicio, key=lambda semana: abs(semana - semana_fin_calendario)),
        )
        if rango_inicio[0] > rango_inicio[1]:
            rango_inicio = (rango_inicio[1], rango_inicio[0])
        st.session_state["inicio_rango_semanas"] = rango_inicio

        meses_es = [
            "enero", "febrero", "marzo", "abril", "mayo", "junio",
            "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
        ]
        st.markdown(
            f"""
            <div class="calendario-inicio-resumen">
                <b>Periodo seleccionado:</b>
                {rango_fechas_estado[0].day} de {meses_es[rango_fechas_estado[0].month - 1]}
                al {rango_fechas_estado[1].day} de {meses_es[rango_fechas_estado[1].month - 1]} de {anio_calendario_inicio}
                <span>Semanas {rango_inicio[0]}–{rango_inicio[1]}</span>
            </div>
            <div class="calendario-leyenda">
                <span><i class="leyenda-cierre"></i>Cierre mensual</span>
                <span><i class="leyenda-festivo"></i>Día festivo</span>
                <span><i class="leyenda-aniversario"></i>Aniversario</span>
                <span><i class="leyenda-listado"></i>Liberación de listados</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        col_entrar_izq, col_entrar, col_entrar_der = st.columns([1.2, 1, 1.2])
        with col_entrar:
            if st.button(
                "Ver resumen semanal  →",
                key="btn_entrar_desde_inicio",
                type="primary",
                use_container_width=True,
            ):
                st.session_state["unidad_negocio_app"] = unidad_inicio
                st.session_state["filtro_superior_País"] = pais_inicio
                st.session_state["filtro_superior_Marca"] = marca_inicio
                st.session_state["modo_moneda_superior"] = moneda_inicio
                st.session_state["rango_semanas_evolucion_input_v3"] = tuple(rango_inicio)
                st.session_state["rango_semanas_evolucion_aplicado_v3"] = tuple(rango_inicio)
                st.rerun()

        # Modo claro forzado para que la portada conserve la identidad visual
        # aun cuando el navegador o Streamlit estén configurados en tema oscuro.
        ruta_landing_marca_agua = buscar_imagen_fondo(NOMBRE_IMAGEN_FONDO)
        landing_background = ""
        if ruta_landing_marca_agua is not None:
            landing_fondo_base64 = imagen_a_base64(str(ruta_landing_marca_agua))
            landing_background = f'''
                background-image:
                    linear-gradient(rgba(248,250,252,0.94), rgba(248,250,252,0.98)),
                    url("data:image/png;base64,{landing_fondo_base64}") !important;
            '''

        st.markdown(
            f"""
            <style>
            html, body, .stApp, [data-testid="stAppViewContainer"] {{
                color-scheme: light !important;
                background-color: #f8fafc !important;
            }}
            .stApp, [data-testid="stAppViewContainer"] {{
                {landing_background}
                background-size: cover !important;
                background-position: top center !important;
                background-repeat: no-repeat !important;
                background-attachment: fixed !important;
            }}
            .main .block-container,
            [data-testid="stMainBlockContainer"] {{
                max-width: 1280px !important;
                padding-top: .75rem !important;
                padding-bottom: 1rem !important;
                background: transparent !important;
                border: 0 !important;
                box-shadow: none !important;
            }}
            [data-testid="stVerticalBlock"]
            > [data-testid="stElementContainer"]:has(style) {{
                display:none !important;
                height:0 !important;
                margin:0 !important;
                padding:0 !important;
            }}
            .inicio-hero {{ text-align:center; padding: 8px 12px 12px; }}
            .inicio-kicker {{
                color:#d48a24; font-size:13px; font-weight:900;
                letter-spacing:2.2px; margin-bottom:8px;
            }}
            .inicio-title {{
                color:#082567; font-size:42px; line-height:1.08;
                font-weight:950; letter-spacing:-1px;
            }}
            .inicio-subtitle {{
                color:#475569; font-size:17px; font-weight:600; margin-top:10px;
            }}
            .inicio-stepper {{
                display:flex; align-items:center; justify-content:center;
                max-width:820px; margin:18px auto 28px;
            }}
            .inicio-step {{
                min-width:150px; display:flex; flex-direction:column;
                align-items:center; gap:6px; color:#94a3b8; font-size:12px;
            }}
            .inicio-step span {{
                width:34px; height:34px; border-radius:50%; display:flex;
                align-items:center; justify-content:center; background:#e2e8f0;
                color:#64748b; font-weight:900;
            }}
            .inicio-step.activo {{ color:#082567; }}
            .inicio-step.activo span {{ background:#082567; color:#ffffff; }}
            .inicio-linea {{ height:3px; background:#dbe3ee; flex:1; margin-bottom:23px; }}
            .inicio-section-title {{
                color:#082567; font-size:19px; font-weight:900; margin:4px 0 12px;
            }}
            .mapa-card {{
                position:relative; background:rgba(255,255,255,0.98);
                border:2px solid #e2e8f0; border-radius:24px;
                padding:14px 20px 18px; min-height:300px; text-align:center;
                box-shadow:0 12px 32px rgba(15,23,42,0.08);
                transition:transform .15s ease, border-color .15s ease, box-shadow .15s ease;
            }}
            .mapa-card.seleccionada {{
                border-color:#d99932;
                box-shadow:0 14px 36px rgba(8,37,103,0.15), 0 0 0 3px rgba(217,153,50,0.13);
            }}
            .mapa-check {{
                position:absolute; right:16px; top:14px; width:30px; height:30px;
                display:flex; align-items:center; justify-content:center;
                border-radius:50%; background:#d99932; color:#ffffff;
                font-weight:950; font-size:18px; z-index:5;
            }}
            .mapa-card:not(.seleccionada) .mapa-check {{ background:transparent; }}
            .mapa-visual {{
                position:relative; height:210px; display:flex;
                align-items:center; justify-content:center; overflow:hidden;
            }}
            .mapa-silueta {{
                width:100%; height:100%; object-fit:contain; display:block;
                filter:drop-shadow(0 10px 10px rgba(15,23,42,0.12));
            }}
            .mapa-latam {{ transform:scale(.90); }}
            .mapa-logo-contenedor {{
                position:absolute; inset:0; display:flex; align-items:center;
                justify-content:center; pointer-events:none;
            }}
            .mapa-logo-overlay {{
                width:155px; height:74px; object-fit:contain; display:block;
                background:rgba(255,255,255,.90); border-radius:14px;
                padding:7px 11px; box-shadow:0 6px 18px rgba(15,23,42,.14);
            }}
            .mapa-logo-latam {{ width:145px; height:82px; }}
            .mapa-logo-contenedor .unidad-logo-placeholder {{
                width:120px; height:72px; margin:0; border-radius:14px;
                background:rgba(255,255,255,.92); font-size:38px;
            }}
            .mapa-card-title {{ color:#082567; font-size:22px; font-weight:950; margin-top:2px; }}
            .mapa-card-subtitle {{ color:#64748b; font-size:14px; font-weight:650; margin-top:2px; }}
            .inicio-filtros-title {{ margin-top:26px; }}
            .calendario-inicio-titulo {{
                color:#082567; font-size:19px; font-weight:950; margin:24px 0 3px;
            }}
            .calendario-inicio-ayuda {{
                color:#64748b; font-size:14px; font-weight:600; margin-bottom:10px;
            }}
            .calendario-inicio-resumen {{
                display:flex; align-items:center; gap:8px; flex-wrap:wrap;
                color:#1e293b; background:#ffffff; border:1px solid #dbe3ee;
                border-radius:12px; padding:11px 14px; margin-top:8px;
                box-shadow:0 5px 14px rgba(15,23,42,.05);
            }}
            .calendario-inicio-resumen b {{ color:#082567; }}
            .calendario-inicio-resumen span {{
                margin-left:auto; color:#ffffff; background:#082567;
                border-radius:999px; padding:4px 10px; font-weight:850;
            }}
            .calendario-leyenda {{
                display:flex; flex-wrap:wrap; gap:18px; margin:10px 2px 22px;
                color:#475569; font-size:12px; font-weight:750;
            }}
            .calendario-leyenda span {{ display:flex; align-items:center; gap:6px; }}
            .calendario-leyenda i {{
                width:13px; height:13px; display:inline-block; border-radius:4px;
            }}
            .leyenda-cierre {{ background:#34436f; }}
            .leyenda-festivo {{ background:#f5df28; }}
            .leyenda-aniversario {{ background:#f5822a; }}
            .leyenda-listado {{ background:#84c35a; }}
            div.stButton > button {{ border-radius:12px !important; font-weight:850 !important; }}
            div.stButton > button[kind="primary"] {{
                background:#082567 !important; color:#ffffff !important;
                border-color:#082567 !important; min-height:48px;
            }}
            [data-testid="stExpander"] {{
                background:rgba(255,255,255,.96); border:1px solid #e2e8f0;
                border-radius:14px; margin-top:12px;
            }}
            @media (max-width: 800px) {{
                .inicio-title {{ font-size:30px; }}
                .inicio-step b {{ display:none; }}
                .inicio-step {{ min-width:45px; }}
                .inicio-linea {{ margin-bottom:0; }}
                .mapa-card {{ min-height:255px; }}
                .mapa-visual {{ height:170px; }}
            }}
            </style>
            """,
            unsafe_allow_html=True,
        )

        st.stop()

    unidad_negocio_seleccionada = st.session_state.get("unidad_negocio_app", None)
else:
    st.session_state["unidad_negocio_app"] = None
    unidad_negocio_seleccionada = None


# ============================================================
# BARRA SUPERIOR
# ============================================================
def es_unidad_latam(valor) -> bool:
    return "LATAM" in normalizar_texto_tc(valor)


st.markdown(
    f"""
    <style>
    .gestion-encabezado {{
        display:flex; align-items:center; justify-content:space-between; gap:24px;
        background:rgba(255,255,255,.97); border:1px solid #dbe3ee;
        border-radius:16px; padding:14px 20px; margin:0 0 14px;
        box-shadow:0 8px 24px rgba(15,23,42,.07);
    }}
    .gestion-marca {{ display:flex; align-items:center; gap:12px; min-width:220px; }}
    .gestion-icono {{ height:34px; display:flex; align-items:flex-end; gap:4px; }}
    .gestion-icono i {{ width:7px; display:block; border-radius:2px 2px 0 0; }}
    .gestion-icono i:nth-child(1) {{ height:13px; background:#082567; }}
    .gestion-icono i:nth-child(2) {{ height:23px; background:#d99932; }}
    .gestion-icono i:nth-child(3) {{ height:31px; background:#f2cf29; }}
    .gestion-titulo {{ color:#082567; font-size:24px; font-weight:950; line-height:1; }}
    .gestion-unidad {{ color:#64748b; font-size:12px; font-weight:750; margin-top:4px; }}
    .gestion-progreso {{ display:flex; align-items:center; flex:1; justify-content:flex-end; max-width:760px; }}
    .gestion-paso {{
        display:flex; align-items:center; gap:8px; white-space:nowrap;
        color:#94a3b8; font-size:13px;
    }}
    .gestion-paso span {{
        width:30px; height:30px; display:flex; align-items:center; justify-content:center;
        border-radius:50%; background:#e2e8f0; color:#64748b; font-weight:950;
    }}
    .gestion-paso.completado {{ color:#36548c; }}
    .gestion-paso.completado span {{ background:#d9e9fb; color:#082567; }}
    .gestion-paso.activo {{ color:#082567; }}
    .gestion-paso.activo span {{ background:#e7c42d; color:#ffffff; }}
    .gestion-trazo {{ height:2px; width:48px; background:#cbd5e1; margin:0 12px; }}
    .gestion-trazo.completado {{ background:#a9c8ed; }}
    .gestion-nav {{
        display:grid; grid-template-columns:repeat(7, minmax(125px,1fr)); gap:8px;
        margin:12px 0 18px; overflow-x:auto; padding-bottom:3px;
    }}
    .gestion-nav a {{
        display:flex; align-items:center; justify-content:center; gap:7px;
        min-height:46px; padding:7px 10px; border-radius:12px;
        background:#eef5fd; border:1px solid #d7e6f7; color:#082567 !important;
        text-decoration:none !important; font-size:12px; font-weight:850;
        text-align:center; transition:.15s ease;
    }}
    .gestion-nav a:hover {{ background:#082567; color:#ffffff !important; transform:translateY(-2px); }}
    .gestion-nav-numero {{
        width:22px; height:22px; border-radius:50%; background:#ffffff;
        display:inline-flex; align-items:center; justify-content:center;
        color:#082567; font-weight:950; flex:0 0 auto;
    }}
    .gestion-seccion-ancla {{ scroll-margin-top:20px; }}
    .gestion-seccion-cabecera {{
        display:flex; align-items:center; gap:10px; margin:22px 0 10px;
        padding:11px 14px; background:linear-gradient(90deg,#eef5fd 0%,rgba(255,255,255,.45) 100%);
        border-left:5px solid #d99932; border-radius:0 12px 12px 0;
        color:#082567; font-size:20px; font-weight:950;
    }}
    .gestion-seccion-cabecera span {{
        width:29px; height:29px; display:flex; align-items:center; justify-content:center;
        border-radius:50%; background:#082567; color:#ffffff; font-size:13px;
    }}
    @media (max-width:900px) {{
        .gestion-encabezado {{ align-items:flex-start; flex-direction:column; }}
        .gestion-progreso {{ width:100%; justify-content:flex-start; overflow-x:auto; }}
        .gestion-paso b {{ font-size:11px; }}
        .gestion-nav {{ grid-template-columns:repeat(7, 135px); }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

texto_indicaciones = (
    "El botón <b>ⓘ</b> contiene Moneda, Marca, País y el alcance de coordinadoras. "
    "Estos filtros ajustan los KPIs, tablas, movimientos, análisis de coordinadoras y las dos "
    "diapositivas de Cobranza. La tarjeta <b>Cobranza · Barras</b> muestra el cumplimiento semanal "
    "y <b>Cobranza · Líneas</b> compara cuota contra recuperación. "
)

if es_unidad_latam(unidad_negocio_seleccionada):
    texto_indicaciones += (
        "En <b>Presico LATAM</b>, si seleccionas <b>Pesos mexicanos</b>, puedes ver las variables "
        "consolidadas con <b>todos</b> los países y consultar el <b>Top / Bottom por país</b>; "
        "si seleccionas <b>Moneda local</b>, la visualización se hace <b>solo por país</b> "
        "para no mezclar monedas. "
    )

texto_indicaciones += (
    "El botón <b>Ir a Análisis</b> abre la tercera etapa, donde se muestra la lectura del último corte "
    "y se genera el PowerPoint; el botón <b>Cambiar unidad</b> regresa a la selección inicial."
)

st.markdown(
    """
    <style>
    .st-key-ayuda_tablero {
        position:relative !important;
        z-index:30 !important;
        width:clamp(430px,31vw,555px) !important;
        max-width:calc(100% - 690px) !important;
        margin:-73px 0 25px clamp(260px,23vw,430px) !important;
    }
    .st-key-ayuda_tablero [data-testid="stHorizontalBlock"] { align-items:center !important; }
    .gestion-progreso { max-width:600px !important; }
    .st-key-ayuda_tablero [data-testid="stPopover"] > button {
        width:38px !important; min-width:38px !important; height:38px !important;
        min-height:38px !important; padding:0 !important; border-radius:50% !important;
        font-size:19px !important; box-shadow:0 4px 12px rgba(8,37,103,.15) !important;
    }
    .st-key-ayuda_tablero div.stButton > button { min-height:42px !important; }
    @media (max-width:1100px) {
        .st-key-ayuda_tablero {
            width:100% !important; max-width:100% !important;
            margin:-4px 0 4px !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# AJUSTE ESTRUCTURAL MEDIDO EN STREAMLIT CLOUD
# - elimina los huecos que generan los bloques <style> sin altura;
# - elimina la fila reservada por cada diapositiva oculta;
# - reduce el padding nativo superior/inferior de Streamlit;
# - conserva los arcos del fondo sin la marca de agua central.
# ============================================================
if unidad_negocio_seleccionada is not None:
    seccion_cloud_solicitada = str(
        st.query_params.get("seccion")
        or st.session_state.get("gestion_seccion_activa", "kpis")
    )
    if seccion_cloud_solicitada not in {
        "kpis", "evolucion", "coordinadoras", "movimientos",
        "top-bottom", "conclusiones", "cobranza-barras", "cobranza-lineas",
    }:
        seccion_cloud_solicitada = "kpis"
    seccion_cloud_ajustada = seccion_cloud_solicitada.replace("-", "_")

    st.markdown(
        f"""
        <style>
        [data-testid="stMainBlockContainer"] {{
            width:100% !important;
            max-width:100% !important;
            padding:10px clamp(12px,1.5vw,24px) 8px !important;
            margin:0 !important;
            box-sizing:border-box !important;
        }}

        /* El bloque raíz de Gestión usa un solo ritmo vertical compacto. */
        [data-testid="stVerticalBlock"]:has(.st-key-gestion_header_barra) {{
            gap:5px !important;
        }}

        /* Cada st.markdown que solo inyecta CSS medía 0 px, pero Streamlit le
           agregaba 16 px de separación. No debe ocupar una fila visual. */
        [data-testid="stVerticalBlock"]:has(.st-key-gestion_header_barra)
        > [data-testid="stElementContainer"]:has(style) {{
            display:none !important;
            height:0 !important;
            min-height:0 !important;
            margin:0 !important;
            padding:0 !important;
        }}

        /* Los siete slides ocultos dejaban siete filas vacías. Se ocultan sus
           wrappers y se conserva únicamente la diapositiva seleccionada. */
        [data-testid="stVerticalBlock"]:has(.st-key-gestion_header_barra)
        > [data-testid="stLayoutWrapper"]:has([class*="st-key-gestion_slide_"]) {{
            display:none !important;
            height:0 !important;
            min-height:0 !important;
            margin:0 !important;
        }}

        [data-testid="stVerticalBlock"]:has(.st-key-gestion_header_barra)
        > [data-testid="stLayoutWrapper"]:has(.st-key-gestion_slide_{seccion_cloud_ajustada}) {{
            display:flex !important;
            height:auto !important;
            min-height:0 !important;
        }}

        .st-key-gestion_header_barra {{
            margin:0 !important;
        }}

        [class*="st-key-gestion_slide_"] {{
            margin:0 !important;
        }}

        .gestion-carrusel-real {{
            margin:0 !important;
        }}

        /* El popover tiene un div intermedio; se apunta al botón real. */
        .stApp .st-key-gestion_header_barra [data-testid="stPopoverButton"] {{
            width:36px !important;
            min-width:36px !important;
            max-width:36px !important;
            height:34px !important;
            min-height:34px !important;
            padding:0 !important;
            border-radius:9px !important;
            background:linear-gradient(180deg,#ffffff 0%,#edf4fc 100%) !important;
            color:#082567 !important;
            -webkit-text-fill-color:#082567 !important;
            border:1px solid #b9cee7 !important;
            box-shadow:0 2px 6px rgba(15,23,42,.07) !important;
        }}

        .stApp .st-key-gestion_header_barra [data-testid="stPopoverButton"] * {{
            color:#082567 !important;
            -webkit-text-fill-color:#082567 !important;
            opacity:1 !important;
        }}

        @media (max-width:900px) {{
            [data-testid="stMainBlockContainer"] {{
                padding:7px 8px 8px !important;
            }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

abrir_resumen_pais_click = False
barra_gestion = st.container(key="gestion_header_barra")
with barra_gestion:
    col_marca_barra, col_acciones_barra, col_pasos_barra = st.columns(
        [1, 1, 1],
        gap="small",
    )

    with col_marca_barra:
        st.markdown(
            f"""
            <div class="gestion-marca">
                <span class="gestion-icono"><i></i><i></i><i></i></span>
                <div>
                    <div class="gestion-titulo">Gestión</div>
                    <div class="gestion-unidad">{html.escape(str(unidad_negocio_seleccionada or 'Todas las unidades'))}</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col_acciones_barra:
        col_ayuda, col_resumen_pais, col_cambiar = st.columns(
            [.24, 1, 1],
            gap="small",
        )
        with col_ayuda:
            filtros_popover = st.popover("ⓘ", use_container_width=True)
        with col_resumen_pais:
            abrir_resumen_pais_click = st.button(
                "Ir a Análisis",
                key="btn_abrir_resumen_pais",
                use_container_width=True,
            )
        with col_cambiar:
            if unidad_negocio_seleccionada is not None:
                if st.button("Cambiar unidad", key="btn_cambiar_unidad", use_container_width=True):
                    st.session_state.pop("unidad_negocio_app", None)
                    st.query_params.clear()
                    st.rerun()

    with col_pasos_barra:
        st.markdown(
            """
            <div class="gestion-progreso">
                <div class="gestion-paso completado"><span>✓</span><b>1. Seleccionar unidad</b></div>
                <div class="gestion-trazo completado"></div>
                <div class="gestion-paso activo"><span>2</span><b>2. Gestión</b></div>
                <div class="gestion-trazo"></div>
                <div class="gestion-paso"><span>3</span><b>3. Análisis</b></div>
            </div>
            """,
            unsafe_allow_html=True,
        )


filtros = {}

if unidad_negocio_seleccionada is not None:
    filtros["Unidad de Negocio"] = [unidad_negocio_seleccionada]

base_para_filtros = filtrar_por_diccionario(df, filtros)

modulo_seleccionado = "Cartera"
filtros_popover.__enter__()
st.markdown("#### Filtros")
st.markdown(texto_indicaciones, unsafe_allow_html=True)

# Detecta el alcance de países antes de mostrar moneda.
# Regla solicitada:
# - Presico México: solo Moneda local.
# - LATAM en Moneda local: no permite "Todos" en País y entra Guatemala por default.
# - LATAM en Pesos mexicanos: sí permite "Todos".
paises_base_filtro = []
if "País" in base_para_filtros.columns:
    paises_base_filtro = sorted(base_para_filtros["País"].dropna().astype(str).str.strip().unique().tolist())

unidad_norm = normalizar_texto_tc(unidad_negocio_seleccionada) if unidad_negocio_seleccionada is not None else ""
paises_norm = [normalizar_texto_tc(x) for x in paises_base_filtro]
es_presico_mexico = (
    unidad_norm == "PRESICO"
    and len(paises_norm) == 1
    and paises_norm[0] in ["MEXICO", "MX"]
)
es_latam = "LATAM" in unidad_norm

opciones_moneda = ["Moneda local"] if es_presico_mexico else ["Moneda local", "Pesos mexicanos"]

# Si el usuario venía de Pesos mexicanos y ahora entra a Presico México,
# se fuerza automáticamente a Moneda local para evitar una opción inválida.
if st.session_state.get("modo_moneda_superior") not in opciones_moneda:
    st.session_state["modo_moneda_superior"] = "Moneda local"

modo_moneda = st.radio(
    "Moneda",
    options=opciones_moneda,
    index=opciones_moneda.index(st.session_state.get("modo_moneda_superior", "Moneda local")),
    horizontal=True,
    key="modo_moneda_superior"
)

# Marca y País quedan como únicas opciones de filtro dentro del tablero.
# En LATAM, si ya hay un país seleccionado, Marca se limita a las marcas disponibles de ese país.
pais_seleccionado_estado = st.session_state.get("filtro_superior_País")
for col_filtro in ["Marca", "País"]:
    if col_filtro in df.columns:
        df_opciones = filtrar_por_diccionario(base_para_filtros, filtros, excluir_col=col_filtro)

        if (
            col_filtro == "Marca"
            and es_latam
            and pais_seleccionado_estado
            and pais_seleccionado_estado not in ["Todos", "Sin datos"]
            and "País" in df_opciones.columns
        ):
            df_opciones = df_opciones[
                df_opciones["País"].astype(str).str.strip() == str(pais_seleccionado_estado).strip()
            ]

        valores = sorted(df_opciones[col_filtro].dropna().astype(str).str.strip().unique())

        if col_filtro == "País" and es_latam and modo_moneda == "Moneda local":
            # En LATAM + Moneda local no se permite consolidar "Todos" porque mezcla monedas.
            opciones = valores
            if not opciones:
                opciones = ["Sin datos"]
            default_pais = "Guatemala" if "Guatemala" in opciones else opciones[0]
            valor_actual = st.session_state.get(f"filtro_superior_{col_filtro}")
            if valor_actual not in opciones:
                st.session_state[f"filtro_superior_{col_filtro}"] = default_pais
            index_default = opciones.index(st.session_state.get(f"filtro_superior_{col_filtro}", default_pais))
        else:
            opciones = ["Todos"] + valores
            valor_actual = st.session_state.get(f"filtro_superior_{col_filtro}")
            if valor_actual not in opciones:
                st.session_state[f"filtro_superior_{col_filtro}"] = "Todos"
            index_default = opciones.index(st.session_state.get(f"filtro_superior_{col_filtro}", "Todos"))

        seleccion = st.selectbox(
            col_filtro,
            options=opciones,
            index=index_default,
            key=f"filtro_superior_{col_filtro}"
        )
        if seleccion != "Todos" and seleccion != "Sin datos":
            filtros[col_filtro] = [seleccion]
    else:
        st.caption(f"Sin columna {col_filtro}")

# Filtro adicional de coordinadoras DESACTIVADO.
# Antes permitía elegir "Sin secundarias" y eso reducía el total de coordinadoras.
# Ahora siempre se conservan TODAS las coordinadoras para KPIs, tablas, gráficas y análisis.
excluir_secundarias_cartera = False
if "Tipo Coordinadora" in df.columns:
    st.caption("Coordinadoras: Todas")
filtros_popover.__exit__(None, None, None)

secciones_gestion_validas = [
    "kpis",
    "evolucion",
    "coordinadoras",
    "movimientos",
    "top-bottom",
    "conclusiones",
    "cobranza-barras",
    "cobranza-lineas",
]

seccion_query = st.query_params.get("seccion")
if seccion_query:
    seccion_query = str(seccion_query)
    if seccion_query in secciones_gestion_validas:
        st.session_state["gestion_seccion_activa"] = seccion_query

seccion_gestion_activa = str(
    st.session_state.get("gestion_seccion_activa", "kpis")
)
if seccion_gestion_activa not in secciones_gestion_validas:
    seccion_gestion_activa = "kpis"
    st.session_state["gestion_seccion_activa"] = seccion_gestion_activa

if modulo_seleccionado == "Cartera":
    st.markdown(
        f"""
        <style>
        [class*="st-key-gestion_slide_"] {{ display:none !important; }}
        .st-key-gestion_slide_{seccion_gestion_activa.replace('-', '_')} {{
            display:block !important;
            position:relative;
            width:100% !important;
            max-width:100% !important;
            box-sizing:border-box !important;
            min-height:52vh;
            max-height:64vh;
            overflow-y:auto;
            overflow-x:hidden;
            background:rgba(255,255,255,.98);
            border:1px solid #dbe3ee;
            border-radius:17px;
            padding:24px 28px 28px;
            box-shadow:0 12px 34px rgba(15,23,42,.10);
        }}
        .st-key-gestion_slide_{seccion_gestion_activa.replace('-', '_')}::before {{
            content:""; position:absolute; left:-55px; top:-75px;
            width:230px; height:125px; border-radius:50%;
            border:18px solid #f0cf2c; border-right-color:transparent;
            border-bottom-color:transparent; transform:rotate(-10deg);
            pointer-events:none; opacity:.92;
        }}
        .st-key-gestion_slide_{seccion_gestion_activa.replace('-', '_')}::after {{
            content:""; position:absolute; right:-70px; bottom:-92px;
            width:260px; height:150px; border-radius:50%;
            border:22px solid #082567; border-left-color:transparent;
            border-top-color:transparent; transform:rotate(-8deg);
            pointer-events:none; opacity:.96;
        }}
        .st-key-gestion_slide_{seccion_gestion_activa.replace('-', '_')} [data-testid="stPlotlyChart"] {{
            width:100% !important; max-width:100% !important; overflow:hidden !important;
        }}
        .st-key-gestion_slide_{seccion_gestion_activa.replace('-', '_')} [data-testid="stPlotlyChart"] > div,
        .st-key-gestion_slide_{seccion_gestion_activa.replace('-', '_')} .js-plotly-plot,
        .st-key-gestion_slide_{seccion_gestion_activa.replace('-', '_')} .plot-container {{
            width:100% !important; max-width:100% !important;
        }}
        .st-key-gestion_slide_{seccion_gestion_activa.replace('-', '_')} [data-testid="stDataFrame"],
        .st-key-gestion_slide_{seccion_gestion_activa.replace('-', '_')} [data-testid="stTable"] {{
            width:100% !important; max-width:100% !important; overflow-x:auto !important;
        }}
        .st-key-gestion_slide_{seccion_gestion_activa.replace('-', '_')} img,
        .st-key-gestion_slide_{seccion_gestion_activa.replace('-', '_')} svg {{
            max-width:100% !important;
        }}
        .gestion-miniaturas {{
            display:grid; grid-template-columns:repeat(7,minmax(145px,1fr)); gap:10px;
            margin:12px 0 4px; overflow-x:auto; padding:2px 2px 8px;
        }}
        .gestion-miniatura {{
            position:relative; min-height:92px; border:2px solid #dbe3ee;
            border-radius:12px; background:linear-gradient(145deg,#ffffff 0%,#eef5fd 100%);
            padding:12px 10px 10px 42px; color:#082567 !important;
            text-decoration:none !important; box-shadow:0 5px 14px rgba(15,23,42,.07);
            overflow:hidden; transition:.15s ease;
        }}
        .gestion-miniatura:hover {{ transform:translateY(-3px); border-color:#d99932; }}
        .gestion-miniatura.activa {{ border-color:#0b70c9; box-shadow:0 0 0 2px rgba(11,112,201,.13); }}
        .gestion-miniatura-numero {{
            position:absolute; left:9px; bottom:9px; width:25px; height:25px;
            border-radius:50%; display:flex; align-items:center; justify-content:center;
            background:#082567; color:#ffffff; font-size:12px; font-weight:950;
        }}
        .gestion-miniatura-titulo {{ font-size:12px; font-weight:950; line-height:1.15; }}
        .gestion-miniatura-preview {{
            height:38px; margin-top:8px; border-radius:6px;
            background:
                linear-gradient(90deg,transparent 0 8%,#76b7ea 8% 16%,transparent 16% 24%,#f0cf2c 24% 37%,transparent 37% 45%,#082567 45% 58%,transparent 58% 66%,#f49aa0 66% 78%,transparent 78%);
            opacity:.70;
        }}
        .gestion-miniatura:nth-child(1) .gestion-miniatura-preview {{
            width:38px; border-radius:50%;
            background:conic-gradient(#0b70c9 0 40%,#f49aa0 40% 62%,#76b7ea 62% 84%,#ff2d2d 84% 100%);
            box-shadow:inset 0 0 0 10px #ffffff;
        }}
        .gestion-miniatura:nth-child(3) .gestion-miniatura-preview {{
            background:linear-gradient(165deg,transparent 0 42%,#0b70c9 43% 48%,transparent 49% 58%,#f0cf2c 59% 64%,transparent 65%);
        }}
        .gestion-miniatura:nth-child(5) .gestion-miniatura-preview {{
            background:
                repeating-linear-gradient(0deg,rgba(8,37,103,.15) 0 8px,transparent 8px 10px),
                repeating-linear-gradient(90deg,#d7e6f7 0 14px,#5b8fc9 14px 28px,#173b73 28px 42px);
        }}
        .gestion-miniatura:nth-child(6) .gestion-miniatura-preview {{
            background:
                linear-gradient(90deg,#0b70c9 0 75%,transparent 75%) 0 2px/100% 7px no-repeat,
                linear-gradient(90deg,#0b70c9 0 55%,transparent 55%) 0 15px/100% 7px no-repeat,
                linear-gradient(90deg,#f0cf2c 0 88%,transparent 88%) 0 28px/100% 7px no-repeat;
        }}
        .gestion-miniatura:nth-child(7) .gestion-miniatura-preview {{
            background:
                linear-gradient(90deg,#94a3b8 0 82%,transparent 82%) 0 3px/100% 5px no-repeat,
                linear-gradient(90deg,#94a3b8 0 66%,transparent 66%) 0 16px/100% 5px no-repeat,
                linear-gradient(90deg,#94a3b8 0 74%,transparent 74%) 0 29px/100% 5px no-repeat;
        }}
        .st-key-gestion_carrusel {{
            width:100%; overflow-x:auto; overflow-y:hidden; padding:3px 2px 9px;
            scrollbar-color:#9db5d4 transparent; scrollbar-width:thin;
        }}
        .st-key-gestion_carrusel [data-testid="stHorizontalBlock"] {{
            min-width:1080px; gap:10px !important;
        }}
        .st-key-gestion_carrusel button {{
            min-height:94px !important; height:94px !important;
            white-space:normal !important; line-height:1.2 !important;
            align-items:flex-start !important; justify-content:flex-start !important;
            text-align:left !important; padding:12px 11px !important;
            border:2px solid #dbe3ee !important; border-radius:12px !important;
            color:#082567 !important; font-size:12px !important; font-weight:900 !important;
            background:
                linear-gradient(90deg,#0b70c9 0 66%,transparent 66%) 12px 58px/70% 6px no-repeat,
                linear-gradient(90deg,#f0cf2c 0 48%,transparent 48%) 12px 72px/70% 6px no-repeat,
                linear-gradient(145deg,#ffffff 0%,#eef5fd 100%) !important;
            box-shadow:0 5px 14px rgba(15,23,42,.07) !important;
        }}
        .st-key-gestion_carrusel button:hover {{
            transform:translateY(-2px); border-color:#d99932 !important;
        }}
        .st-key-gestion_mini_{seccion_gestion_activa.replace('-', '_')} button {{
            border-color:#0b70c9 !important;
            box-shadow:0 0 0 2px rgba(11,112,201,.14),0 7px 18px rgba(15,23,42,.10) !important;
        }}
        @media (max-width:900px) {{
            .gestion-miniaturas {{ grid-template-columns:repeat(7,145px); }}
            .st-key-gestion_slide_{seccion_gestion_activa.replace('-', '_')} {{
                width:100% !important; max-width:100% !important;
                max-height:68vh; min-height:46vh; padding:18px 14px 22px;
                overflow-y:auto; overflow-x:hidden;
            }}
            .st-key-gestion_slide_{seccion_gestion_activa.replace('-', '_')} [data-testid="stHorizontalBlock"] {{
                flex-wrap:wrap !important;
            }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

# ============================================================
# MINIATURAS FIELES Y DIAPOSITIVA SIN ADORNOS EXTERIORES
# ============================================================
if globals().get("modulo_seleccionado") == "Cartera":
    miniatura_activa_limpia = str(
        st.session_state.get("gestion_seccion_activa", "kpis")
    ).replace("-", "_")

    st.markdown(
        f"""
        <style>
        /* Quita por completo las curvas decorativas amarilla y azul. */
        [class*="st-key-gestion_slide_"]::before,
        [class*="st-key-gestion_slide_"]::after {{
            content:none !important;
            display:none !important;
            border:0 !important;
            background:none !important;
        }}

        /* La diapositiva ocupa el ancho útil y crece con el contenido, sin scroll propio. */
        .st-key-gestion_slide_{miniatura_activa_limpia} {{
            width:100% !important;
            max-width:100% !important;
            min-width:0 !important;
            min-height:0 !important;
            height:auto !important;
            max-height:none !important;
            overflow:visible !important;
            margin:6px 0 10px !important;
            padding:clamp(14px,1.7vw,24px) clamp(14px,2vw,28px) 22px !important;
            box-sizing:border-box !important;
            border:1px solid #d8e1ed !important;
            border-radius:12px !important;
            background:#ffffff !important;
            box-shadow:0 8px 22px rgba(15,23,42,.08) !important;
        }}

        .st-key-gestion_slide_{miniatura_activa_limpia} .gestion-seccion-cabecera {{
            margin:0 0 9px !important;
        }}

        .st-key-gestion_slide_{miniatura_activa_limpia} [data-testid="stPlotlyChart"],
        .st-key-gestion_slide_{miniatura_activa_limpia} [data-testid="stDataFrame"],
        .st-key-gestion_slide_{miniatura_activa_limpia} [data-testid="stTable"],
        .st-key-gestion_slide_{miniatura_activa_limpia} img,
        .st-key-gestion_slide_{miniatura_activa_limpia} svg {{
            max-width:100% !important;
            box-sizing:border-box !important;
        }}

        /* Carrusel real: solo esta franja puede desplazarse horizontalmente. */
        .gestion-carrusel-real {{
            display:flex;
            flex-flow:row nowrap;
            align-items:stretch;
            gap:10px;
            width:100%;
            max-width:100%;
            overflow-x:auto;
            overflow-y:hidden;
            box-sizing:border-box;
            padding:4px 2px 11px;
            margin:0;
            overscroll-behavior-inline:contain;
            scrollbar-width:thin;
            scrollbar-color:#8ea7c6 transparent;
        }}

        .gestion-carrusel-real::-webkit-scrollbar {{ height:8px; }}
        .gestion-carrusel-real::-webkit-scrollbar-track {{ background:transparent; }}
        .gestion-carrusel-real::-webkit-scrollbar-thumb {{
            background:#8ea7c6;
            border-radius:999px;
        }}

        .gestion-miniatura-real {{
            position:relative;
            flex:1 0 178px;
            min-width:178px;
            height:124px;
            box-sizing:border-box;
            overflow:hidden;
            padding:10px 10px 9px;
            border:1px solid #d4deea;
            border-radius:10px;
            background:#ffffff;
            box-shadow:0 4px 11px rgba(15,23,42,.08);
            color:#082567 !important;
            text-decoration:none !important;
            transition:border-color .14s ease,box-shadow .14s ease,transform .14s ease;
        }}

        .gestion-miniatura-real:hover {{
            border-color:#8fb6e3;
            box-shadow:0 7px 16px rgba(15,23,42,.13);
            transform:translateY(-1px);
            color:#082567 !important;
            text-decoration:none !important;
        }}

        .gestion-miniatura-real.activa {{
            border:2px solid #0b70c9;
            padding:9px 9px 8px;
            box-shadow:0 0 0 2px rgba(11,112,201,.11),0 6px 15px rgba(15,23,42,.11);
        }}

        .gestion-miniatura-titulo-real {{
            position:relative;
            z-index:3;
            min-height:23px;
            overflow:hidden;
            color:#082567 !important;
            font-size:11px;
            font-weight:950;
            line-height:1.05;
            text-align:left;
            white-space:normal;
        }}

        .gestion-miniatura-numero-real {{
            position:absolute;
            left:8px;
            bottom:7px;
            z-index:5;
            display:flex;
            align-items:center;
            justify-content:center;
            width:23px;
            height:23px;
            border-radius:50%;
            background:#082567;
            color:#ffffff !important;
            font-size:11px;
            font-weight:950;
            box-shadow:0 2px 5px rgba(8,37,103,.22);
        }}

        /* 1. Resumen ejecutivo: dona y tarjetas de resultado. */
        .mini-resumen {{
            position:absolute; left:38px; right:9px; top:34px; bottom:9px;
            display:grid; grid-template-columns:55px 1fr; gap:9px; align-items:center;
        }}
        .mini-dona {{
            width:48px; height:48px; border-radius:50%;
            background:conic-gradient(#0b70c9 0 41%,#f49aa0 41% 63%,#76b7ea 63% 84%,#ff2d2d 84% 100%);
            box-shadow:inset 0 0 0 13px #ffffff;
        }}
        .mini-resumen-datos {{ display:flex; flex-direction:column; gap:4px; }}
        .mini-resumen-datos b {{ color:#082567; font-size:8px; line-height:1; }}
        .mini-resumen-datos strong {{ color:#18864b; font-size:8px; line-height:1; }}
        .mini-resumen-datos i {{ display:block; height:5px; border-radius:2px; background:#dbeafe; width:88%; }}
        .mini-resumen-datos i:last-child {{ width:66%; background:#dcfce7; }}

        /* 2. KPIs: tres tarjetas y evolución en columnas. */
        .mini-kpis {{ position:absolute; left:39px; right:9px; top:34px; bottom:8px; }}
        .mini-kpis > span {{
            display:inline-flex; width:29%; height:20px; margin-right:2%;
            border-radius:4px; background:#eaf3fd; align-items:center; gap:3px; padding:2px 3px;
            box-sizing:border-box;
        }}
        .mini-kpis > span b {{ color:#1d9a5b; font-size:7px; }}
        .mini-kpis > span i {{ height:4px; width:70%; background:#8eabd0; border-radius:2px; }}
        .mini-columnas {{ position:absolute; inset:26px 0 0; display:flex; align-items:flex-end; justify-content:space-around; border-bottom:1px solid #cbd5e1; }}
        .mini-columnas i {{ width:9%; background:#0b70c9; height:45%; }}
        .mini-columnas i:nth-child(2) {{ height:76%; background:#f0cf2c; }}
        .mini-columnas i:nth-child(3) {{ height:58%; background:#76b7ea; }}
        .mini-columnas i:nth-child(4) {{ height:90%; }}
        .mini-columnas i:nth-child(5) {{ height:64%; background:#f0cf2c; }}
        .mini-columnas i:nth-child(6) {{ height:82%; background:#76b7ea; }}

        /* 3. Evolución semanal: gráfica de líneas completa. */
        .mini-evolucion {{
            position:absolute; left:39px; right:9px; top:35px; bottom:10px;
            background:repeating-linear-gradient(0deg,transparent 0 15px,#e5eaf1 15px 16px);
            border-bottom:1px solid #b8c5d6;
        }}
        .mini-evolucion svg {{ width:100%; height:100%; overflow:visible; }}
        .mini-evolucion path {{ fill:none; stroke:#1976c9; stroke-width:3; }}
        .mini-evolucion path.linea-secundaria {{ stroke:#f0c928; stroke-width:2.5; }}

        /* 4. Tipo de coordinadora: dona y leyenda por categoría. */
        .mini-coordinadoras {{
            position:absolute; left:39px; right:9px; top:34px; bottom:9px;
            display:grid; grid-template-columns:58px 1fr; gap:8px; align-items:center;
        }}
        .mini-dona-coord {{ width:53px; height:53px; box-shadow:inset 0 0 0 15px #ffffff; }}
        .mini-leyenda {{ display:flex; flex-direction:column; gap:6px; }}
        .mini-leyenda i {{ display:block; height:5px; width:88%; border-radius:2px; background:#0b70c9; }}
        .mini-leyenda i:nth-child(2) {{ width:72%; background:#f49aa0; }}
        .mini-leyenda i:nth-child(3) {{ width:80%; background:#76b7ea; }}
        .mini-leyenda i:nth-child(4) {{ width:54%; background:#ff2d2d; }}

        /* 5. Matriz: celdas de intensidad como la tabla principal. */
        .mini-matriz {{
            position:absolute; left:42px; right:12px; top:37px; bottom:12px;
            display:grid; grid-template-columns:repeat(5,1fr); grid-template-rows:repeat(3,1fr);
            gap:2px; padding:3px; border-radius:5px; background:#eaf2fb;
        }}
        .mini-matriz i {{ background:#d9e8f8; border-radius:1px; }}
        .mini-matriz i:nth-child(2), .mini-matriz i:nth-child(6), .mini-matriz i:nth-child(13) {{ background:#91b4dc; }}
        .mini-matriz i:nth-child(3), .mini-matriz i:nth-child(8), .mini-matriz i:nth-child(12) {{ background:#4f7fb7; }}
        .mini-matriz i:nth-child(7), .mini-matriz i:nth-child(9) {{ background:#163c73; }}

        /* 6. Top / Bottom: dos rankings lado a lado. */
        .mini-top-bottom {{
            position:absolute; left:39px; right:9px; top:36px; bottom:11px;
            display:grid; grid-template-columns:1fr 1fr; gap:10px;
        }}
        .mini-top-bottom > div {{ display:flex; flex-direction:column; gap:6px; padding-top:3px; }}
        .mini-top-bottom i {{ display:block; height:6px; background:#0b70c9; }}
        .mini-top-bottom div:first-child i:nth-child(1) {{ width:92%; }}
        .mini-top-bottom div:first-child i:nth-child(2) {{ width:72%; }}
        .mini-top-bottom div:first-child i:nth-child(3) {{ width:55%; }}
        .mini-top-bottom div:first-child i:nth-child(4) {{ width:38%; }}
        .mini-top-bottom div:last-child i {{ background:#efc91f; }}
        .mini-top-bottom div:last-child i:nth-child(1) {{ width:88%; }}
        .mini-top-bottom div:last-child i:nth-child(2) {{ width:66%; }}
        .mini-top-bottom div:last-child i:nth-child(3) {{ width:51%; }}
        .mini-top-bottom div:last-child i:nth-child(4) {{ width:34%; }}

        /* 7. Conclusiones: tres hallazgos ejecutivos. */
        .mini-conclusiones {{
            position:absolute; left:40px; right:10px; top:37px; bottom:10px;
            display:flex; flex-direction:column; gap:7px;
        }}
        .mini-conclusiones span {{ display:grid; grid-template-columns:17px 1fr; align-items:center; gap:6px; }}
        .mini-conclusiones b {{
            display:flex; align-items:center; justify-content:center; width:16px; height:16px;
            border-radius:50%; background:#dff4eb; color:#177553; font-size:8px;
        }}
        .mini-conclusiones span:nth-child(2) b {{ background:#fff3cd; color:#9a6b00; }}
        .mini-conclusiones span:nth-child(3) b {{ background:#e6f0fb; color:#0b70c9; }}
        .mini-conclusiones i {{ display:block; height:5px; border-radius:2px; background:#94a3b8; width:92%; }}
        .mini-conclusiones span:nth-child(2) i {{ width:72%; }}
        .mini-conclusiones span:nth-child(3) i {{ width:82%; }}

        .mini-cobranza-barras {{
            position:absolute; left:40px; right:10px; top:38px; bottom:10px;
            display:flex; align-items:flex-end; justify-content:space-around;
            gap:4px; border-bottom:1px solid #b8c5d6;
        }}
        .mini-cobranza-barras i {{ display:block; width:8%; height:52%; background:#bfc7d3; }}
        .mini-cobranza-barras i:nth-child(2) {{ height:74%; }}
        .mini-cobranza-barras i:nth-child(3) {{ height:62%; }}
        .mini-cobranza-barras i:nth-child(4) {{ height:88%; background:#0b70c9; }}
        .mini-cobranza-barras i:nth-child(5) {{ height:69%; }}
        .mini-cobranza-barras i:nth-child(6) {{ height:81%; background:#0b70c9; }}
        .mini-cobranza-barras i:nth-child(7) {{ height:58%; }}
        .mini-cobranza-barras i:nth-child(8) {{ height:93%; background:#e7c42d; }}

        .mini-cobranza-lineas {{
            position:absolute; left:40px; right:10px; top:38px; bottom:11px;
            background:repeating-linear-gradient(0deg,transparent 0 14px,#e5eaf1 14px 15px);
            border-bottom:1px solid #b8c5d6;
        }}
        .mini-cobranza-lineas svg {{ width:100%; height:100%; overflow:visible; }}
        .mini-cobranza-lineas path {{ fill:none; stroke:#082567; stroke-width:3; }}
        .mini-cobranza-lineas path.linea-pago {{ stroke:#e7c42d; stroke-width:2.5; }}

        @media (max-width:900px) {{
            [data-testid="stAppViewContainer"] .main .block-container {{
                padding-left:10px !important;
                padding-right:10px !important;
            }}
            .gestion-miniatura-real {{ flex-basis:174px; min-width:174px; }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

# Aplica conversión de moneda después de construir los filtros, para conservar las opciones originales.
df = aplicar_tipo_cambio_mxn(df, modo_moneda)
if df_cobranza is not None:
    df_cobranza = aplicar_tipo_cambio_mxn(df_cobranza, modo_moneda)

# No se excluyen coordinadoras secundarias.
# Se deja el bloque desactivado para conservar el resto del script intacto.
if False and modulo_seleccionado == "Cartera" and excluir_secundarias_cartera and "Tipo Coordinadora" in df.columns:
    df = df[df["Tipo Coordinadora"].astype(str).str.strip().str.lower() != "secundaria"].copy()

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
df_filtrado = recalcular_ip_agregado(df_filtrado)

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
    ir_a_pagina_resumen()

if obtener_query_param("vista") == "resumen_pais":
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
        df_cartera=df_filtrado_original,
        unidad=unidad_negocio_seleccionada or "",
        resumen_cobranza=resumen_cobranza_modal,
        semana_actual_cobranza=semana_actual_cob_modal,
        semana_anterior_cobranza=semana_anterior_cob_modal,
    )
    st.stop()



# ============================================================
# GRÁFICA DE BURBUJAS: CLIENTES Y FALTAS POR SUCURSAL
# ============================================================
def preparar_datos_burbujas_faltas(df_base: pd.DataFrame, nivel_estructura: str = "Sucursal") -> tuple[pd.DataFrame, str | None]:
    """
    Construye la base de la burbuja por semana y sucursal.
    X = Clientes Totales.
    Y = Faltas / Clientes Totales.
    Tamaño = Faltas.
    No usa días de atraso ni escala de color.
    """
    if df_base is None or df_base.empty:
        return pd.DataFrame(), None

    nivel_estructura = "Sucursal"

    if nivel_estructura not in df_base.columns:
        return pd.DataFrame(), None

    columnas_necesarias = ["Semana del año", nivel_estructura, "Clientes Totales", "Faltas"]
    faltantes = [c for c in columnas_necesarias if c not in df_base.columns]
    if faltantes:
        return pd.DataFrame(), None

    df_tmp = df_base.copy()
    df_tmp["Semana del año"] = pd.to_numeric(df_tmp["Semana del año"], errors="coerce")
    df_tmp["Clientes Totales"] = pd.to_numeric(df_tmp["Clientes Totales"], errors="coerce").fillna(0)
    df_tmp["Faltas"] = pd.to_numeric(df_tmp["Faltas"], errors="coerce").fillna(0)
    df_tmp = df_tmp.dropna(subset=["Semana del año", nivel_estructura]).copy()

    if df_tmp.empty:
        return pd.DataFrame(), None

    columnas_group = ["Semana del año", nivel_estructura]
    if "Año" in df_tmp.columns:
        df_tmp["Año"] = pd.to_numeric(df_tmp["Año"], errors="coerce")
        columnas_group = ["Año", "Semana del año", nivel_estructura]

    agrupado = (
        df_tmp
        .groupby(columnas_group, dropna=False)
        .agg(
            **{
                "Clientes Totales": ("Clientes Totales", "sum"),
                "Faltas": ("Faltas", "sum"),
            }
        )
        .reset_index()
    )

    agrupado["Porcentaje de Faltas"] = np.where(
        agrupado["Clientes Totales"] == 0,
        0,
        agrupado["Faltas"] / agrupado["Clientes Totales"]
    )

    if "Año" in agrupado.columns:
        agrupado = agrupado.sort_values(["Año", "Semana del año", nivel_estructura])
        agrupado["Etiqueta semana"] = agrupado.apply(
            lambda r: f"{int(r['Año'])} - S{int(r['Semana del año'])}" if pd.notna(r["Año"]) else f"S{int(r['Semana del año'])}",
            axis=1
        )
    else:
        agrupado = agrupado.sort_values(["Semana del año", nivel_estructura])
        agrupado["Etiqueta semana"] = agrupado["Semana del año"].apply(lambda x: f"S{int(x)}")

    agrupado["Estructura"] = agrupado[nivel_estructura].astype(str)
    agrupado["Faltas texto"] = agrupado["Faltas"].apply(lambda x: f"{x:,.0f}")
    agrupado["Clientes texto"] = agrupado["Clientes Totales"].apply(lambda x: f"{x:,.0f}")
    agrupado["% Faltas texto"] = agrupado["Porcentaje de Faltas"].apply(lambda x: f"{x:.1%}")

    return agrupado, None


def generar_comentario_burbujas_faltas(df_burbujas: pd.DataFrame, nivel_estructura: str = "Sucursal", semana_visible: str | None = None) -> str:
    if df_burbujas is None or df_burbujas.empty:
        return "No hay información suficiente para comentar la matriz de clientes y faltas por sucursal."

    nivel_estructura = "Sucursal"
    df_tmp = df_burbujas.copy()
    if semana_visible is not None:
        df_sem = df_tmp[df_tmp["Etiqueta semana"] == semana_visible].copy()
        if df_sem.empty:
            df_sem = df_tmp.copy()
    else:
        semana_visible = df_tmp["Etiqueta semana"].iloc[-1]
        df_sem = df_tmp[df_tmp["Etiqueta semana"] == semana_visible].copy()

    if df_sem.empty:
        return "No hay registros válidos para comentar la gráfica de burbujas."

    med_clientes = df_sem["Clientes Totales"].median()
    med_pct = df_sem["Porcentaje de Faltas"].median()

    zona_critica = df_sem[
        (df_sem["Clientes Totales"] >= med_clientes) &
        (df_sem["Porcentaje de Faltas"] >= med_pct)
    ].copy()

    if not zona_critica.empty:
        foco = zona_critica.sort_values(["Faltas", "Porcentaje de Faltas"], ascending=False).iloc[0]
        lectura_foco = (
            f"El principal foco operativo está en {foco['Estructura']}, porque combina una base alta de clientes "
            f"({foco['Clientes Totales']:,.0f}) con un porcentaje elevado de faltas ({foco['Porcentaje de Faltas']:.1%}) "
            f"y {foco['Faltas']:,.0f} faltas."
        )
    else:
        foco = df_sem.sort_values(["Porcentaje de Faltas", "Faltas"], ascending=False).iloc[0]
        lectura_foco = (
            f"El mayor foco por proporción de faltas está en {foco['Estructura']}, con "
            f"{foco['Porcentaje de Faltas']:.1%} y {foco['Faltas']:,.0f} faltas."
        )

    mayor_clientes = df_sem.sort_values("Clientes Totales", ascending=False).iloc[0]
    menor_faltas_rel = df_sem.sort_values("Porcentaje de Faltas", ascending=True).iloc[0]

    return (
        f"En {semana_visible}, la lectura por sucursal separa cuatro cuadrantes: sucursales con más clientes y más faltas, "
        f"más clientes y menos faltas, menos clientes y más faltas, y menos clientes y menos faltas. {lectura_foco} "
        f"La sucursal con mayor volumen de clientes es {mayor_clientes['Estructura']} "
        f"({mayor_clientes['Clientes Totales']:,.0f} clientes). "
        f"Como referencia positiva, {menor_faltas_rel['Estructura']} registra el menor porcentaje de faltas "
        f"({menor_faltas_rel['Porcentaje de Faltas']:.1%})."
    )


def crear_grafica_burbujas_faltas(
    df_burbujas: pd.DataFrame,
    nivel_estructura: str = "Sucursal",
    modo_interactivo: bool = False,
    semana_estatica: str | None = None,
) -> tuple[go.Figure, dict]:
    if df_burbujas is None or df_burbujas.empty:
        return go.Figure(), {"displayModeBar": False, "responsive": True}

    nivel_estructura = "Sucursal"
    df_plot = df_burbujas.copy()
    df_plot["Tamaño burbuja"] = pd.to_numeric(df_plot["Faltas"], errors="coerce").fillna(0).clip(lower=1)

    # Para que la gráfica no se amontone, no se etiqueta cada burbuja.
    # Se muestran solo las sucursales más relevantes; el resto se consulta con hover.
    df_plot["Etiqueta visible"] = ""
    for _, idx_semana in df_plot.groupby("Etiqueta semana").groups.items():
        tmp = df_plot.loc[list(idx_semana)].copy()
        idx_etiquetas = set()
        # Solo se etiquetan los casos más relevantes para que no se amontone el texto.
        # El resto de sucursales se consulta al pasar el cursor sobre la burbuja.
        for col in ["Clientes Totales", "Faltas", "Porcentaje de Faltas"]:
            idx_etiquetas.update(tmp.nlargest(4, col).index.tolist())
        idx_etiquetas.update(tmp.nsmallest(2, "Porcentaje de Faltas").index.tolist())
        df_plot.loc[list(idx_etiquetas), "Etiqueta visible"] = df_plot.loc[list(idx_etiquetas), "Estructura"]

    max_x = max(float(df_plot["Clientes Totales"].max()), 1)
    max_y = max(float(df_plot["Porcentaje de Faltas"].max()), 0.01)

    # Evita que las etiquetas de burbujas pegadas al origen tapen el nombre
    # del cuadrante "Menos clientes y menos faltas". La burbuja sigue visible
    # y el nombre de la sucursal se conserva en el hover.
    mascara_etiquetas_origen = (
        (pd.to_numeric(df_plot["Clientes Totales"], errors="coerce").fillna(0) <= max_x * 0.16)
        & (pd.to_numeric(df_plot["Porcentaje de Faltas"], errors="coerce").fillna(0) <= max_y * 0.18)
    )
    df_plot.loc[mascara_etiquetas_origen, "Etiqueta visible"] = ""

    if modo_interactivo:
        fig = px.scatter(
            df_plot,
            x="Clientes Totales",
            y="Porcentaje de Faltas",
            size="Tamaño burbuja",
            animation_frame="Etiqueta semana",
            animation_group="Estructura",
            hover_name="Estructura",
            text="Etiqueta visible",
            size_max=42,
            range_x=[0, max_x * 1.12],
            range_y=[0, min(max_y * 1.18, 1.05)],
            labels={
                "Clientes Totales": "Clientes Totales",
                "Porcentaje de Faltas": "Porcentaje de Faltas",
                "Tamaño burbuja": "Faltas",
            },
            custom_data=["Clientes texto", "Faltas texto", "% Faltas texto"],
        )
    else:
        if semana_estatica is None:
            semana_estatica = df_plot["Etiqueta semana"].iloc[-1]
        df_sem = df_plot[df_plot["Etiqueta semana"] == semana_estatica].copy()
        if df_sem.empty:
            df_sem = df_plot[df_plot["Etiqueta semana"] == df_plot["Etiqueta semana"].iloc[-1]].copy()
            semana_estatica = df_sem["Etiqueta semana"].iloc[-1]

        fig = px.scatter(
            df_sem,
            x="Clientes Totales",
            y="Porcentaje de Faltas",
            size="Tamaño burbuja",
            hover_name="Estructura",
            text="Etiqueta visible",
            size_max=42,
            range_x=[0, max_x * 1.12],
            range_y=[0, min(max_y * 1.18, 1.05)],
            labels={
                "Clientes Totales": "Clientes Totales",
                "Porcentaje de Faltas": "Porcentaje de Faltas",
                "Tamaño burbuja": "Faltas",
            },
            custom_data=["Clientes texto", "Faltas texto", "% Faltas texto"],
        )

    med_x = float(df_plot["Clientes Totales"].median())
    med_y = float(df_plot["Porcentaje de Faltas"].median())

    fig.update_traces(
        textposition="top center",
        textfont=dict(size=11, color="#111827"),
        marker=dict(color="#5aa9f2", line=dict(color="white", width=2), opacity=0.62),
        hovertemplate=(
            "<b>%{hovertext}</b><br>"
            "Clientes Totales: %{customdata[0]}<br>"
            "Faltas: %{customdata[1]}<br>"
            "% Faltas: %{customdata[2]}"
            "<extra></extra>"
        ),
    )

    fig.add_vline(x=med_x, line_width=2.5, line_color="rgba(80,80,80,0.80)")
    fig.add_hline(y=med_y, line_width=2.5, line_color="rgba(80,80,80,0.80)")

    fig.add_annotation(
        x=max_x * 0.03, y=min(max_y * 1.10, 0.98), text="<b><i>Más faltas<br>y menos clientes</i></b>",
        showarrow=False, xanchor="left", yanchor="top", font=dict(size=14, color="#082567")
    )
    fig.add_annotation(
        x=max_x * 1.08, y=min(max_y * 1.10, 0.98), text="<b><i>Más clientes<br>y más faltas</i></b>",
        showarrow=False, xanchor="right", yanchor="top", font=dict(size=14, color="#082567")
    )
    fig.add_annotation(
        x=max_x * 0.03, y=0, text="<b><i>Menos clientes<br>y menos faltas</i></b>",
        showarrow=False, xanchor="left", yanchor="bottom", font=dict(size=14, color="#082567")
    )
    fig.add_annotation(
        x=max_x * 1.08, y=0, text="<b><i>Más clientes<br>y menos faltas</i></b>",
        showarrow=False, xanchor="right", yanchor="bottom", font=dict(size=14, color="#082567")
    )

    fig.update_layout(
        title=dict(text="Clientes Totales, Porcentaje de Faltas y Faltas por Sucursal", x=0.0, xanchor="left", font=dict(size=16, color="#111827")),
        height=760,
        margin=dict(t=90, b=80, l=95, r=70),
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,1)",
        font=dict(color="#111827", size=14),
        showlegend=False,
        dragmode=False,
    )

    fig.update_xaxes(
        title="Clientes Totales",
        showgrid=False,
        zeroline=False,
        fixedrange=True,
        tickformat=",.0f",
    )
    fig.update_yaxes(
        title="Porcentaje de Faltas",
        showgrid=False,
        zeroline=False,
        fixedrange=True,
        tickformat=".0%",
    )

    if modo_interactivo and fig.layout.updatemenus:
        fig.layout.updatemenus[0].buttons[0].args[1]["frame"] = {"duration": 850, "redraw": True}
        fig.layout.updatemenus[0].buttons[0].args[1]["transition"] = {"duration": 450}

    config = {
        "displayModeBar": False,
        "scrollZoom": False,
        "doubleClick": False,
        "responsive": True,
        "modeBarButtonsToRemove": [
            "zoom2d", "pan2d", "select2d", "lasso2d", "zoomIn2d", "zoomOut2d",
            "autoScale2d", "resetScale2d"
        ],
    }

    return fig, config

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
    slide_kpis = st.container(key="gestion_slide_kpis")
    slide_kpis.__enter__()
    st.markdown(
        '<div id="gestion-kpis" class="gestion-seccion-ancla"></div>'
        '<div class="gestion-seccion-cabecera"><span>1</span>KPIs de la última semana</div>',
        unsafe_allow_html=True,
    )
    kpis_gestion = [
        indicador
        for indicador in [
            "Clientes Totales",
            "Clientes al corriente",
            "Cartera Total",
            "Saldo en atraso",
        ]
        if indicador in resumen_general_pais["Indicador"].astype(str).tolist()
    ] if resumen_general_pais is not None and not resumen_general_pais.empty else []

    if kpis_gestion:
        columnas_kpis_gestion = st.columns(len(kpis_gestion))
        for indice_kpi, indicador_kpi in enumerate(kpis_gestion):
            fila_kpi = resumen_general_pais[
                resumen_general_pais["Indicador"] == indicador_kpi
            ].iloc[0]
            with columnas_kpis_gestion[indice_kpi]:
                tarjeta_kpi(
                    indicador_kpi,
                    fila_kpi.get(f"Dato sem {semana_actual}", 0),
                    fila_kpi.get("Variación vs sem ant", np.nan)
                    if semana_anterior_general_pais is not None
                    else None,
                )
    else:
        st.info("No hay indicadores disponibles para mostrar KPIs con los filtros actuales.")

    slide_kpis.__exit__(None, None, None)
    slide_evolucion = st.container(key="gestion_slide_evolucion")
    slide_evolucion.__enter__()

    # ============================================================
    # GRÁFICAS CARTERA
    # ============================================================
    comentario_evolucion = ""
    comentario_pie = ""
    st.markdown(
        '<div id="gestion-evolucion" class="gestion-seccion-ancla"></div>'
        '<div class="gestion-seccion-cabecera"><span>2</span>Evolución</div>',
        unsafe_allow_html=True,
    )
    col1 = st.container()
    col2 = st.container()

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
                "Saldo PP",
                "IP"
            ]
            if c in df_filtrado.columns
        ]

        if columna_cobranza_cartera and columna_cobranza_cartera in df_filtrado.columns:
            metricas_evolucion.append(columna_cobranza_cartera)

        if metricas_evolucion:
            col_grafica, col_menu = st.columns([4, 1])

            semanas_evolucion = sorted(
                int(semana)
                for semana in df_filtrado["Semana del año"].dropna().unique()
            )
            rango_semanas_default = (
                semanas_evolucion[-8] if len(semanas_evolucion) >= 8 else semanas_evolucion[0],
                semanas_evolucion[-1],
            )

            # Limpia las llaves usadas por versiones anteriores del control.
            st.session_state.pop("rango_semanas_evolucion", None)
            st.session_state.pop("semana_inicio_evolucion", None)
            st.session_state.pop("semana_fin_evolucion", None)
            st.session_state.pop("rango_semanas_evolucion_input", None)
            st.session_state.pop("rango_semanas_evolucion_aplicado", None)
            st.session_state.pop("rango_semanas_evolucion_input_v2", None)
            st.session_state.pop("rango_semanas_evolucion_aplicado_v2", None)

            def rango_evolucion_valido(valor) -> bool:
                return (
                    isinstance(valor, (list, tuple))
                    and len(valor) == 2
                    and valor[0] in semanas_evolucion
                    and valor[1] in semanas_evolucion
                    and valor[0] <= valor[1]
                )

            if not rango_evolucion_valido(
                st.session_state.get("rango_semanas_evolucion_input_v3")
            ):
                # El valor debe pasarse explícitamente al widget como tupla para
                # que Streamlit lo construya en modo rango (dos extremos).
                st.session_state.pop("rango_semanas_evolucion_input_v3", None)

            if not rango_evolucion_valido(
                st.session_state.get("rango_semanas_evolucion_aplicado_v3")
            ):
                st.session_state["rango_semanas_evolucion_aplicado_v3"] = rango_semanas_default

            with col_menu:
                # El formulario evita ejecutar nuevamente toda la aplicación
                # mientras se arrastran los extremos del rango.
                with st.form("form_rango_semanas_evolucion_v3", clear_on_submit=False):
                    rango_semanas_input = st.select_slider(
                        "Rango de semanas",
                        options=semanas_evolucion,
                        value=rango_semanas_default,
                        format_func=lambda semana: f"S{int(semana)}",
                        key="rango_semanas_evolucion_input_v3",
                    )
                    aplicar_rango_semanas = st.form_submit_button(
                        "Aplicar rango",
                        use_container_width=True,
                    )

                if aplicar_rango_semanas:
                    st.session_state["rango_semanas_evolucion_aplicado_v3"] = tuple(
                        rango_semanas_input
                    )

                (
                    semana_inicio_evolucion,
                    semana_fin_evolucion,
                ) = st.session_state["rango_semanas_evolucion_aplicado_v3"]

                indicador_grafica = st.selectbox(
                    "Indicador",
                    options=metricas_evolucion,
                    index=metricas_evolucion.index("Cartera Total")
                    if "Cartera Total" in metricas_evolucion
                    else 0,
                    key="indicador_evolucion"
                )

            df_evolucion = df_filtrado[
                df_filtrado["Semana del año"].between(
                    semana_inicio_evolucion,
                    semana_fin_evolucion,
                )
            ].copy()

            if indicador_grafica == "IP" and {"Clientes al corriente", "Clientes Totales"}.issubset(df_evolucion.columns):
                evol_base_ip = (
                    df_evolucion
                    .groupby("Semana del año", dropna=False)[["Clientes al corriente", "Clientes Totales"]]
                    .sum(numeric_only=True)
                    .reset_index()
                    .sort_values("Semana del año")
                )
                evol_base_ip["IP"] = np.where(
                    evol_base_ip["Clientes Totales"] == 0,
                    0,
                    (evol_base_ip["Clientes al corriente"] / evol_base_ip["Clientes Totales"]) * 100
                )
                evol = evol_base_ip[["Semana del año", "IP"]].copy()
            else:
                evol = (
                    df_evolucion
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

                mostrar_grafica_adaptable(
                    fig_linea,
                    config=config_linea,
                    key="grafica_evolucion_semanal"
                )

        else:
            st.info("No hay indicadores disponibles para la gráfica de evolución semanal.")


    mostrar_boton_comentario("grafica_evolucion", comentario_evolucion)
    slide_evolucion.__exit__(None, None, None)
    slide_coordinadoras = st.container(key="gestion_slide_coordinadoras")
    slide_coordinadoras.__enter__()
    st.markdown(
        '<div id="gestion-coordinadoras" class="gestion-seccion-ancla"></div>'
        '<div class="gestion-seccion-cabecera"><span>3</span>Tipo de coordinadora</div>',
        unsafe_allow_html=True,
    )
    with col2:
        st.subheader(f"Distribución por tipo de coordinadora | Última semana: {semana_ultima_historial}")

        # CORRECCIÓN:
        # Antes esta gráfica hacía una llave coordinadora_id + País + Marca y después
        # quitaba duplicados. En PRESICO México eso reducía el total porque existen
        # coordinadoras con el mismo ID/País/Marca en más de una Ruta.
        # Para no excluir ninguna coordinadora/registro de la base, aquí se cuenta
        # directamente cada fila de la última semana filtrada.
        df_sem_actual = df_filtrado_original[
            df_filtrado_original["Semana del año"] == semana_ultima_historial
        ].copy()

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
                    domain=dict(x=[0.10, 0.90], y=[0.08, 0.92]),
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
            textfont=dict(size=12, color="#082567", family="Arial"),
            pull=[0] * len(pie)
        )

        fig_pie.update_layout(
            height=520,
            showlegend=False,
            paper_bgcolor="rgba(255,255,255,0)",
            plot_bgcolor="rgba(255,255,255,0)",
            font=dict(color="#082567", size=13),
            margin=dict(t=45, b=65, l=125, r=125),
            uniformtext_minsize=11,
            uniformtext_mode="show"
        )

        fig_pie.update_layout(dragmode=False)
        mostrar_grafica_adaptable(
            fig_pie,
            config={"displayModeBar": False, "scrollZoom": False, "doubleClick": False, "responsive": True},
            key="grafica_pie_coordinadoras"
        )
        comentario_pie = generar_comentario_pie(pie)


    mostrar_boton_comentario("pie_coordinadoras", comentario_pie)


    # ============================================================
    # MATRIZ DE DESPLAZAMIENTO DE COORDINADORAS
    # ============================================================
    slide_coordinadoras.__exit__(None, None, None)
    slide_movimientos = st.container(key="gestion_slide_movimientos")
    slide_movimientos.__enter__()
    st.markdown(
        '<div id="gestion-movimientos" class="gestion-seccion-ancla"></div>'
        '<div class="gestion-seccion-cabecera"><span>4</span>Movimientos de coordinadoras</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="movimientos-subtitulo">Matriz de desplazamiento por categoría</div>',
        unsafe_allow_html=True,
    )

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

                col_m1, col_m2, col_m3, col_m4 = st.columns(4)

                with col_m1:
                    st.metric(f"Coordinadoras sem {semana_origen}", f"{total_origen:,.0f}")

                with col_m2:
                    st.metric(f"Coordinadoras sem {semana_destino}", f"{total_destino:,.0f}")

                with col_m3:
                    st.metric("Nuevas", f"{total_nuevas:,.0f}")

                with col_m4:
                    st.metric("Bajas", f"{total_bajas:,.0f}")

                col_matriz_mov, col_resumen_mov = st.columns([2.65, 1], gap="small")

                with col_matriz_mov:
                    st.markdown(
                        matriz_desplazamiento_html(matriz_movimientos),
                        unsafe_allow_html=True,
                    )
                    boton_descargar_xlsx(
                        matriz_movimientos.reset_index() if hasattr(matriz_movimientos, "reset_index") else matriz_movimientos,
                        "Descargar matriz de movimientos XLSX",
                        f"matriz_movimientos_semana_{semana_origen}_{semana_destino}.xlsx",
                        key="descargar_matriz_movimientos_xlsx"
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

                    # Muestra únicamente coordinadoras que sí cambiaron de categoría
                    # entre la semana anterior y la semana actual.
                    # Esto elimina los registros que permanecen igual, que antes saturaban la tabla.
                    detalle_movimientos = detalle_movimientos[
                        detalle_movimientos["Semana anterior"].astype(str)
                        != detalle_movimientos["Semana actual"].astype(str)
                    ].copy()

                    if detalle_movimientos.empty:
                        st.info("No hubo coordinadoras desplazadas entre las semanas seleccionadas.")
                    else:
                        # Se consolida la identificación en una sola columna, usando origen cuando existe
                        # y destino cuando se trata de coordinadoras nuevas.
                        if "coordinadora_id origen" in detalle_movimientos.columns:
                            detalle_movimientos["Coordinadora ID"] = detalle_movimientos["coordinadora_id origen"]
                        else:
                            detalle_movimientos["Coordinadora ID"] = np.nan

                        if "coordinadora_id destino" in detalle_movimientos.columns:
                            detalle_movimientos["Coordinadora ID"] = detalle_movimientos["Coordinadora ID"].fillna(
                                detalle_movimientos["coordinadora_id destino"]
                            )

                        if "País origen" in detalle_movimientos.columns:
                            detalle_movimientos["País"] = detalle_movimientos["País origen"]
                        else:
                            detalle_movimientos["País"] = np.nan

                        if "País destino" in detalle_movimientos.columns:
                            detalle_movimientos["País"] = detalle_movimientos["País"].fillna(
                                detalle_movimientos["País destino"]
                            )

                        if "Marca origen" in detalle_movimientos.columns:
                            detalle_movimientos["Marca"] = detalle_movimientos["Marca origen"]
                        else:
                            detalle_movimientos["Marca"] = np.nan

                        if "Marca destino" in detalle_movimientos.columns:
                            detalle_movimientos["Marca"] = detalle_movimientos["Marca"].fillna(
                                detalle_movimientos["Marca destino"]
                            )

                        columnas_detalle = [
                            "Coordinadora ID",
                            "País",
                            "Marca",
                            "Semana anterior",
                            "Semana actual",
                        ]
                        columnas_detalle = [
                            c for c in columnas_detalle
                            if c in detalle_movimientos.columns
                        ]

                        st.dataframe(
                            detalle_movimientos[columnas_detalle],
                            use_container_width=True,
                            hide_index=True
                        )
                        boton_descargar_xlsx(
                            detalle_movimientos[columnas_detalle],
                            "Descargar detalle de coordinadoras XLSX",
                            f"detalle_coordinadoras_desplazadas_{semana_origen}_{semana_destino}.xlsx",
                            key="descargar_detalle_coordinadoras_xlsx"
                        )


    # ============================================================
    # BURBUJAS: CLIENTES Y FALTAS POR SUCURSAL
    # Se mantiene disponible sin alargar la diapositiva principal.
    # ============================================================
    detalle_burbujas = st.expander(
        "Ver análisis complementario de clientes y faltas por sucursal",
        expanded=False,
    )
    detalle_burbujas.__enter__()
    st.subheader("Mapa de clientes y faltas por sucursal")

    nivel_burbujas = "Sucursal"

    if nivel_burbujas not in df_filtrado_original.columns:
        st.info("No existe la columna 'Sucursal' para construir la gráfica de burbujas.")
    elif not all(c in df_filtrado_original.columns for c in ["Clientes Totales", "Faltas", "Semana del año"]):
        st.info("Para esta gráfica se necesitan las columnas 'Clientes Totales', 'Faltas' y 'Semana del año'.")
    else:
        st.markdown(
            """
            <div class="top-bottom-opciones-card">
                <div class="top-bottom-opciones-title">Opciones de matriz de clientes y faltas por sucursal</div>
            </div>
            """,
            unsafe_allow_html=True
        )

        col_modo_burbuja, col_semana_burbuja = st.columns([1.35, 1.35], gap="large")

        datos_burbujas, _ = preparar_datos_burbujas_faltas(
            df_base=df_filtrado_original,
            nivel_estructura=nivel_burbujas
        )

        if datos_burbujas.empty:
            st.info("No hay datos suficientes para construir la gráfica de burbujas con los filtros actuales.")
        else:
            # Gráfica fija: se elimina el botón de movimiento/animación para evitar saturación
            # y errores en vista móvil. La semana se elige manualmente.
            modo_interactivo_burbujas = False

            semanas_burbujas = datos_burbujas["Etiqueta semana"].dropna().drop_duplicates().tolist()

            with col_semana_burbuja:
                semana_burbuja_estatica = st.selectbox(
                    "Semana",
                    options=semanas_burbujas,
                    index=len(semanas_burbujas) - 1,
                    key="semana_burbujas_faltas"
                )

            with col_modo_burbuja:
                st.caption("Gráfica fija por sucursal. Desliza horizontalmente para verla completa en teléfono.")

            fig_burbujas, config_burbujas = crear_grafica_burbujas_faltas(
                df_burbujas=datos_burbujas,
                nivel_estructura=nivel_burbujas,
                modo_interactivo=modo_interactivo_burbujas,
                semana_estatica=semana_burbuja_estatica,
            )

            mostrar_grafica_burbujas_amplia(
                fig_burbujas,
                config=config_burbujas,
                key="grafica_burbujas_faltas"
            )

            comentario_burbujas = generar_comentario_burbujas_faltas(
                df_burbujas=datos_burbujas,
                nivel_estructura=nivel_burbujas,
                semana_visible=None if modo_interactivo_burbujas else semana_burbuja_estatica
            )
            mostrar_boton_comentario("burbujas_clientes_faltas", comentario_burbujas)

            st.caption("La gráfica queda fija por Sucursal. El tamaño de la burbuja representa el número de faltas; no se usa días de atraso ni escala de color.")

    detalle_burbujas.__exit__(None, None, None)


    # ============================================================
    # TOP / BOTTOM POR VARIABLE
    # ============================================================
    slide_movimientos.__exit__(None, None, None)
    slide_top_bottom = st.container(key="gestion_slide_top_bottom")
    slide_top_bottom.__enter__()
    st.markdown(
        '<div id="gestion-top-bottom" class="gestion-seccion-ancla"></div>'
        '<div class="gestion-seccion-cabecera"><span>5</span>Top / Bottom</div>',
        unsafe_allow_html=True,
    )
    st.subheader("Top / Bottom por variable")
    comentario_top_bottom = ""

    niveles_top_bottom = filtrar_niveles_top_bottom_visibles(
        df_filtrado,
        ["País", "Subdireccion", "Zona", "Sucursal", "Ruta"]
    )

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

                boton_descargar_xlsx(
                    tabla_top_bottom,
                    "Descargar Top / Bottom XLSX",
                    f"top_bottom_{nivel_top_bottom}_semana_{semana_actual}.xlsx",
                    key="descargar_top_bottom_xlsx"
                )

    mostrar_boton_comentario("top_bottom", comentario_top_bottom)

    conclusiones_gestion = [
        texto
        for texto in [
            comentario_general_pais,
            comentario_evolucion,
            comentario_pie,
            comentario_top_bottom,
        ]
        if texto is not None and str(texto).strip()
    ]

    # ============================================================
    # TABLA POR NIVEL
    # ============================================================
    detalle_top_bottom = st.expander(
        f"Ver detalle agrupado por {nivel}",
        expanded=False,
    )
    detalle_top_bottom.__enter__()
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
    boton_descargar_xlsx(
        detalle,
        "Descargar detalle agrupado XLSX",
        f"detalle_{nivel}_semana_{semana_actual}.xlsx",
        key="descargar_detalle_agrupado_xlsx"
    )

    detalle_top_bottom.__exit__(None, None, None)

    slide_top_bottom.__exit__(None, None, None)

    slide_conclusiones = st.container(key="gestion_slide_conclusiones")
    slide_conclusiones.__enter__()
    st.markdown(
        '<div id="gestion-conclusiones" class="gestion-seccion-ancla"></div>'
        '<div class="gestion-seccion-cabecera"><span>6</span>Conclusiones</div>',
        unsafe_allow_html=True,
    )
    mostrar_boton_comentario(
        "gestion_conclusiones_dinamicas",
        " ".join(dict.fromkeys(conclusiones_gestion)),
    )
    slide_conclusiones.__exit__(None, None, None)

    # ============================================================
    # COBRANZA COMO DOS DIAPOSITIVAS DEL CARRUSEL
    # ============================================================
    evol_cobranza_tarjetas = pd.DataFrame()
    evol_cobranza_visual_tarjetas = pd.DataFrame()
    col_cuota_tarjeta = None
    col_pago_tarjeta = None
    col_cump_tarjeta = None
    col_mejor_tarjeta = None
    col_peor_tarjeta = None
    error_cobranza_tarjetas = ""

    if df_cobranza is None:
        error_cobranza_tarjetas = error_cobranza or "No se encontró información de Cobranza."
    else:
        (
            df_cobranza_preparada_tarjetas,
            col_cuota_tarjeta,
            col_pago_tarjeta,
            col_cump_tarjeta,
            col_mejor_tarjeta,
            col_peor_tarjeta,
        ) = preparar_cobranza(df_cobranza)

        if "Semana del año" not in df_cobranza_preparada_tarjetas.columns:
            error_cobranza_tarjetas = "La base de Cobranza no contiene Semana o Semana del año."
        elif col_cuota_tarjeta is None or col_pago_tarjeta is None:
            error_cobranza_tarjetas = "No se detectaron las columnas de cuota y recuperación en Cobranza."
        else:
            df_cobranza_filtrada_tarjetas = aplicar_filtros_cobranza_desde_cartera(
                df_cobranza_base=df_cobranza_preparada_tarjetas,
                df_cartera_base=df,
                filtros=filtros,
            )
            evol_cobranza_tarjetas = consolidar_cobranza(
                df_cobranza=df_cobranza_filtrada_tarjetas,
                col_cuota=col_cuota_tarjeta,
                col_pago=col_pago_tarjeta,
                col_cump=col_cump_tarjeta,
                col_mejor=col_mejor_tarjeta,
                col_peor=col_peor_tarjeta,
                nivel=None,
            )
            evol_cobranza_visual_tarjetas = limitar_ultimas_semanas_cobranza(
                evol_cobranza_tarjetas,
                MAX_SEMANAS_VISUALES_COBRANZA,
            )
            if evol_cobranza_tarjetas.empty:
                error_cobranza_tarjetas = "No hay datos de Cobranza con los filtros seleccionados."

    slide_cobranza_barras = st.container(key="gestion_slide_cobranza_barras")
    slide_cobranza_barras.__enter__()
    st.markdown(
        '<div id="gestion-cobranza-barras" class="gestion-seccion-ancla"></div>'
        '<div class="gestion-seccion-cabecera"><span>7</span>Cobranza · Cumplimiento</div>',
        unsafe_allow_html=True,
    )
    if error_cobranza_tarjetas:
        st.info(error_cobranza_tarjetas)
    else:
        tabla_cobranza_barras = evol_cobranza_visual_tarjetas.copy()
        if col_cump_tarjeta and col_cump_tarjeta in tabla_cobranza_barras.columns:
            mejor_cumplimiento_tarjeta = pd.to_numeric(
                tabla_cobranza_barras[col_cump_tarjeta], errors="coerce"
            ).max()
            tabla_cobranza_barras["Diferencia vs mejor semana"] = (
                mejor_cumplimiento_tarjeta
                - pd.to_numeric(tabla_cobranza_barras[col_cump_tarjeta], errors="coerce")
            ).clip(lower=0)
        columnas_barras = [
            c for c in [
                "Año",
                "Semana del año",
                "Etiqueta semana",
                col_cump_tarjeta,
                "Diferencia vs mejor semana",
            ]
            if c and c in tabla_cobranza_barras.columns
        ]
        tabla_cobranza_barras = tabla_cobranza_barras[columnas_barras].copy()

        col_grafica_barras, col_tabla_barras = st.columns([3.15, 1.05], gap="medium")
        with col_grafica_barras:
            fig_cobranza_barras = grafica_cumplimiento(
                evol_cobranza_visual_tarjetas,
                col_cump_tarjeta,
                modo_moneda,
            )
            fig_cobranza_barras.update_layout(
                height=430,
                margin=dict(l=35, r=20, t=48, b=45),
                dragmode=False,
            )
            fig_cobranza_barras.update_xaxes(fixedrange=True)
            fig_cobranza_barras.update_yaxes(fixedrange=True)
            st.plotly_chart(
                fig_cobranza_barras,
                width="stretch",
                config={"displayModeBar": False, "scrollZoom": False, "doubleClick": False, "responsive": True},
                key="grafica_cobranza_barras_tarjeta",
            )
        with col_tabla_barras:
            st.markdown("**Datos de la gráfica**")
            st.dataframe(tabla_cobranza_barras, width="stretch", hide_index=True, height=320)
            boton_descargar_xlsx(
                tabla_cobranza_barras,
                "Exportar Excel",
                f"cobranza_cumplimiento_semana_{semana_actual}.xlsx",
                key="exportar_excel_cobranza_barras",
            )
    slide_cobranza_barras.__exit__(None, None, None)

    slide_cobranza_lineas = st.container(key="gestion_slide_cobranza_lineas")
    slide_cobranza_lineas.__enter__()
    st.markdown(
        '<div id="gestion-cobranza-lineas" class="gestion-seccion-ancla"></div>'
        '<div class="gestion-seccion-cabecera"><span>8</span>Cobranza · Cuota vs recuperación</div>',
        unsafe_allow_html=True,
    )
    if error_cobranza_tarjetas:
        st.info(error_cobranza_tarjetas)
    else:
        col_mejor_final_tarjeta = (
            col_mejor_tarjeta
            if col_mejor_tarjeta and col_mejor_tarjeta in evol_cobranza_visual_tarjetas.columns
            else "Mejor semana"
        )
        col_peor_final_tarjeta = (
            col_peor_tarjeta
            if col_peor_tarjeta and col_peor_tarjeta in evol_cobranza_visual_tarjetas.columns
            else "Peor semana"
        )
        columnas_lineas = [
            c for c in [
                "Año",
                "Semana del año",
                "Etiqueta semana",
                col_cuota_tarjeta,
                col_pago_tarjeta,
                col_mejor_final_tarjeta,
                col_peor_final_tarjeta,
            ]
            if c and c in evol_cobranza_visual_tarjetas.columns
        ]
        tabla_cobranza_lineas = evol_cobranza_visual_tarjetas[columnas_lineas].copy()

        col_grafica_lineas, col_tabla_lineas = st.columns([3.15, 1.05], gap="medium")
        with col_grafica_lineas:
            fig_cobranza_lineas = grafica_cuota_pago(
                evol=evol_cobranza_visual_tarjetas,
                col_cuota=col_cuota_tarjeta,
                col_pago=col_pago_tarjeta,
                col_mejor=col_mejor_final_tarjeta,
                col_peor=col_peor_final_tarjeta,
                modo_moneda=modo_moneda,
            )
            fig_cobranza_lineas.update_layout(
                height=430,
                margin=dict(l=35, r=20, t=48, b=45),
                dragmode=False,
            )
            fig_cobranza_lineas.update_xaxes(fixedrange=True)
            fig_cobranza_lineas.update_yaxes(fixedrange=True)
            st.plotly_chart(
                fig_cobranza_lineas,
                width="stretch",
                config={"displayModeBar": False, "scrollZoom": False, "doubleClick": False, "responsive": True},
                key="grafica_cobranza_lineas_tarjeta",
            )
        with col_tabla_lineas:
            st.markdown("**Datos de la gráfica**")
            st.dataframe(tabla_cobranza_lineas, width="stretch", hide_index=True, height=320)
            boton_descargar_xlsx(
                tabla_cobranza_lineas,
                "Exportar Excel",
                f"cobranza_cuota_recuperacion_semana_{semana_actual}.xlsx",
                key="exportar_excel_cobranza_lineas",
            )
    slide_cobranza_lineas.__exit__(None, None, None)

    miniaturas_gestion = [
        (
            "kpis",
            "KPIs",
            '<div class="mini-kpis">'
            '<span><b>▲</b><i></i></span><span><b>●</b><i></i></span><span><b>▲</b><i></i></span>'
            '<div class="mini-columnas"><i></i><i></i><i></i><i></i><i></i><i></i></div>'
            '</div>',
        ),
        (
            "evolucion",
            "Evolución semanal",
            '<div class="mini-evolucion">'
            '<svg viewBox="0 0 180 58" preserveAspectRatio="none" aria-hidden="true">'
            '<path d="M2 50 L28 38 L52 43 L77 23 L103 31 L128 15 L154 21 L178 7" />'
            '<path class="linea-secundaria" d="M2 55 L30 49 L57 35 L86 42 L116 30 L146 32 L178 20" />'
            '</svg></div>',
        ),
        (
            "coordinadoras",
            "Tipo de coordinadora",
            '<div class="mini-coordinadoras">'
            '<div class="mini-dona mini-dona-coord"></div>'
            '<div class="mini-leyenda"><i></i><i></i><i></i><i></i></div>'
            '</div>',
        ),
        (
            "movimientos",
            "Matriz de movimientos",
            '<div class="mini-matriz">'
            '<i></i><i></i><i></i><i></i><i></i>'
            '<i></i><i></i><i></i><i></i><i></i>'
            '<i></i><i></i><i></i><i></i><i></i>'
            '</div>',
        ),
        (
            "top-bottom",
            "Top / Bottom",
            '<div class="mini-top-bottom">'
            '<div><i></i><i></i><i></i><i></i></div>'
            '<div><i></i><i></i><i></i><i></i></div>'
            '</div>',
        ),
        (
            "conclusiones",
            "Conclusiones",
            '<div class="mini-conclusiones">'
            '<span><b>✓</b><i></i></span><span><b>!</b><i></i></span><span><b>→</b><i></i></span>'
            '</div>',
        ),
        (
            "cobranza-barras",
            "Cobranza · Barras",
            '<div class="mini-cobranza-barras">'
            '<i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i>'
            '</div>',
        ),
        (
            "cobranza-lineas",
            "Cobranza · Líneas",
            '<div class="mini-cobranza-lineas">'
            '<svg viewBox="0 0 180 58" preserveAspectRatio="none" aria-hidden="true">'
            '<path d="M2 46 L26 38 L49 43 L73 26 L98 31 L122 17 L148 22 L178 9" />'
            '<path class="linea-pago" d="M2 54 L26 49 L49 38 L73 41 L98 35 L122 28 L148 31 L178 21" />'
            '</svg></div>',
        ),
    ]

    unidad_miniaturas = quote(str(unidad_negocio_seleccionada or ""), safe="")
    tarjetas_miniaturas = []
    for numero_miniatura, (slug_miniatura, titulo_miniatura, vista_miniatura) in enumerate(
        miniaturas_gestion,
        start=1,
    ):
        clase_activa = " activa" if slug_miniatura == seccion_gestion_activa else ""
        tarjetas_miniaturas.append(
            f'<a class="gestion-miniatura-real{clase_activa}" '
            f'href="?unidad={unidad_miniaturas}&amp;seccion={slug_miniatura}" target="_self">'
            f'<div class="gestion-miniatura-titulo-real">{html.escape(titulo_miniatura)}</div>'
            f'{vista_miniatura}'
            f'<span class="gestion-miniatura-numero-real">{numero_miniatura}</span>'
            '</a>'
        )

    st.markdown(
        '<div class="gestion-carrusel-real">' + "".join(tarjetas_miniaturas) + '</div>',
        unsafe_allow_html=True,
    )

else:
    # ============================================================
    # SECCIÓN COBRANZA
    # ============================================================
    st.subheader("Cobranza")

    if df_cobranza is None:
        st.info(
            "La información de Cobranza no está disponible, pero la vista "
            "Cartera continúa funcionando con los concentrados."
        )
        if error_cobranza:
            st.caption(error_cobranza)

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
                    evol_cobranza_visual = limitar_ultimas_semanas_cobranza(
                        evol_cobranza,
                        MAX_SEMANAS_VISUALES_COBRANZA,
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
                    boton_descargar_xlsx(
                        tabla_detalle_cob,
                        "Descargar detalle de cobranza XLSX",
                        "detalle_cobranza.xlsx",
                        key="descargar_detalle_cobranza_xlsx"
                    )

                    # ------------------------------
                    # Gráfica cumplimiento
                    # ------------------------------
                    st.markdown("**% Cumplimiento semanal**")
                    st.caption(
                        f"Visuales ordenados cronológicamente: últimas "
                        f"{len(evol_cobranza_visual)} semanas disponibles."
                    )
                    mostrar_boton_comentario(
                        "cobranza_cumplimiento",
                        comentario_cobranza_cumplimiento(evol_cobranza_visual, col_cump)
                    )

                    fig_cump = grafica_cumplimiento(evol_cobranza_visual, col_cump, modo_moneda)
                    fig_cump.update_layout(dragmode=False)
                    fig_cump.update_xaxes(fixedrange=True)
                    fig_cump.update_yaxes(fixedrange=True)
                    mostrar_grafica_cobranza_amplia(
                        fig_cump,
                        titulo="% Cumplimiento semanal",
                        ancho=1350,
                        alto=520,
                        config={"displayModeBar": False, "scrollZoom": False, "doubleClick": False, "responsive": False},
                        key="grafica_cumplimiento_cobranza"
                    )

                    # ------------------------------
                    # Gráfica cuota vs pago
                    # ------------------------------
                    st.markdown("**Cuota total vs Pago total**")
                    mostrar_boton_comentario(
                        "cobranza_cuota_pago",
                        comentario_cobranza_cuota_pago(evol_cobranza_visual, col_cuota, col_pago, col_cump)
                    )

                    fig_cp = grafica_cuota_pago(
                        evol=evol_cobranza_visual,
                        col_cuota=col_cuota,
                        col_pago=col_pago,
                        col_mejor=col_mejor_final,
                        col_peor=col_peor_final,
                        modo_moneda=modo_moneda
                    )
                    fig_cp.update_layout(dragmode=False)
                    fig_cp.update_xaxes(fixedrange=True)
                    fig_cp.update_yaxes(fixedrange=True)
                    mostrar_grafica_cobranza_amplia(
                        fig_cp,
                        titulo="Cuota total vs Pago total",
                        ancho=1450,
                        alto=650,
                        config={"displayModeBar": False, "scrollZoom": False, "doubleClick": False, "responsive": False},
                        key="grafica_cuota_pago_cobranza"
                    )

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
                    boton_descargar_xlsx(
                        tabla_ultimas_5_cobranza,
                        "Descargar últimas 5 semanas XLSX",
                        "ultimas_5_semanas_cobranza.xlsx",
                        key="descargar_ultimas_5_cobranza_xlsx"
                    )


                    # ------------------------------
                    # Top / Bottom de Cobranza
                    # ------------------------------
                    mostrar_top_bottom_cobranza = not es_presico_mexico

                    if mostrar_top_bottom_cobranza:
                        st.markdown("**Top / Bottom de cobranza por variable**")

                        niveles_top_bottom_cobranza = filtrar_niveles_top_bottom_visibles(
                            df_cobranza_top_bottom_base,
                            ["País", "Subdireccion", "Zona", "Sucursal", "Ruta"]
                        )

                        variables_top_bottom_cobranza = [
                            c for c in [
                                col_cuota,
                                col_pago,
                                col_cump,
                                col_mejor_final,
                                col_peor_final,
                            ]
                            if (c and c in df_cobranza_top_bottom_base.columns) or c == col_cump
                        ]

                        # Quita duplicados conservando el orden.
                        variables_top_bottom_cobranza = list(dict.fromkeys(variables_top_bottom_cobranza))

                        periodo_top_bottom_cobranza = obtener_ultimo_periodo_cobranza(
                            df_cobranza_top_bottom_base
                        )
                        if periodo_top_bottom_cobranza is None:
                            anio_top_bottom_cobranza = None
                            semana_top_bottom_cobranza = None
                        else:
                            anio_top_bottom_cobranza, semana_top_bottom_cobranza = periodo_top_bottom_cobranza

                        if not niveles_top_bottom_cobranza:
                            st.info("No hay niveles de estructura disponibles en la hoja Cobranza para construir el Top / Bottom.")
                        elif not variables_top_bottom_cobranza:
                            st.info("No hay variables de cobranza disponibles para construir el Top / Bottom.")
                        elif semana_top_bottom_cobranza is None:
                            st.info("No hay semanas válidas en Cobranza para construir el Top / Bottom.")
                        else:
                            col_tabla_top_bottom_cob, col_opciones_top_bottom_cob = st.columns([3.4, 1.1], gap="small")

                            # La estructura ya no se muestra como filtro lateral. Se define automáticamente:
                            # si existe País, se usa País; si no, se toma el primer nivel disponible.
                            nivel_top_bottom_cobranza = "País" if "País" in niveles_top_bottom_cobranza else niveles_top_bottom_cobranza[0]
                            cantidad_top_bottom_cobranza = 30

                            with col_opciones_top_bottom_cob:
                                tipo_top_bottom_cobranza = st.radio(
                                    "Top / Bottom",
                                    options=["Top", "Bottom"],
                                    horizontal=True,
                                    key="tipo_top_bottom_cobranza"
                                )

                                variable_top_bottom_cobranza = st.selectbox(
                                    "Variable de cobranza",
                                    options=variables_top_bottom_cobranza,
                                    index=0,
                                    key="variable_top_bottom_cobranza"
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
                                semana_objetivo=semana_top_bottom_cobranza,
                                anio_objetivo=anio_top_bottom_cobranza,
                            )

                            with col_tabla_top_bottom_cob:
                                if tabla_top_bottom_cobranza.empty:
                                    st.info("No hay datos suficientes para mostrar el Top / Bottom de Cobranza con la selección actual.")
                                else:
                                    periodo_caption = (
                                        f"{anio_top_bottom_cobranza} S{semana_top_bottom_cobranza}"
                                        if anio_top_bottom_cobranza is not None
                                        else f"Semana {semana_top_bottom_cobranza}"
                                    )
                                    st.caption(
                                        f"{periodo_caption} | {tipo_top_bottom_cobranza} "
                                        f"por {nivel_top_bottom_cobranza} | Moneda: {etiqueta_moneda(modo_moneda)}"
                                    )

                                    st.dataframe(
                                        aplicar_formato_top_bottom_cobranza(tabla_top_bottom_cobranza),
                                        use_container_width=True,
                                        hide_index=True
                                    )

                                    boton_descargar_xlsx(
                                        tabla_top_bottom_cobranza,
                                        "Descargar Top / Bottom Cobranza XLSX",
                                        f"top_bottom_cobranza_{nivel_top_bottom_cobranza}_"
                                        f"{anio_top_bottom_cobranza or 'sin_anio'}_"
                                        f"semana_{semana_top_bottom_cobranza}.xlsx",
                                        key="descargar_top_bottom_cobranza_xlsx"
                                    )

                            comentario_top_bottom_cobranza = generar_comentario_top_bottom_cobranza(
                                tabla_top_bottom=tabla_top_bottom_cobranza,
                                tipo_top_bottom=tipo_top_bottom_cobranza,
                                nivel_top_bottom=nivel_top_bottom_cobranza,
                                semana_actual=semana_top_bottom_cobranza
                            )
                            mostrar_boton_comentario("top_bottom_cobranza", comentario_top_bottom_cobranza)
# ============================================================
# AJUSTE FINAL DE CONTRASTE SOLO PARA MODO OSCURO
# ============================================================
st.markdown(
    """
    <style>
    /*
       Este bloque NO modifica el modo claro.
       Solo se activa cuando la computadora/navegador está en modo oscuro.
       Corrige textos que se pierden por bajo contraste en filtros, botones,
       gráficas Plotly, tablas y comentarios.
    */
    @media (prefers-color-scheme: dark) {

        /* Colores base del tablero oscuro */
        :root {
            --dark-bg-main: #0f172a;
            --dark-bg-card: #111827;
            --dark-bg-card-2: #1e293b;
            --dark-border: rgba(219,234,254,0.38);
            --dark-text: #f8fafc;
            --dark-text-soft: #e5e7eb;
            --dark-title: #ffffff;
            --dark-accent: #facc15;
            --dark-button-bg: #dbeafe;
            --dark-button-text: #082567;
        }

        html,
        body,
        .stApp,
        [data-testid="stAppViewContainer"] {
            color-scheme: dark !important;
        }

        .main .block-container {
            background: rgba(15,23,42,0.96) !important;
            color: var(--dark-text) !important;
            border: 1px solid var(--dark-border) !important;
        }

        /* Tarjetas principales */
        .top-filter-card,
        .kpi-card,
        .top-bottom-opciones-card,
        div[data-testid="stPlotlyChart"],
        div[data-testid="stDataFrame"],
        div[data-testid="stTable"],
        div[data-testid="stMetric"],
        div[data-testid="stExpander"] {
            background: rgba(17,24,39,0.96) !important;
            color: var(--dark-text) !important;
            border-color: var(--dark-border) !important;
            box-shadow: 0 8px 24px rgba(0,0,0,0.35) !important;
        }

        /* Títulos generales */
        .top-filter-title,
        .titulo,
        h1, h2, h3, h4,
        .landing-title,
        .unidad-name,
        .kpi-label {
            color: var(--dark-title) !important;
            -webkit-text-fill-color: var(--dark-title) !important;
            opacity: 1 !important;
            text-shadow: 0 1px 2px rgba(0,0,0,0.45) !important;
        }

        h2::after,
        h3::after {
            background: var(--dark-accent) !important;
            opacity: 1 !important;
        }

        /* Textos secundarios */
        .subtitulo,
        .landing-subtitle,
        .unidad-help,
        div[data-testid="stMarkdownContainer"] p,
        .top-filter-card p,
        .top-filter-card span,
        .top-filter-card small {
            color: var(--dark-text-soft) !important;
            -webkit-text-fill-color: var(--dark-text-soft) !important;
            opacity: 1 !important;
        }

        /* Labels de filtros: Vista, Moneda, Marca, País, Indicador, etc. */
        .top-filter-card label,
        .top-filter-card label p,
        .top-filter-card div[data-testid="stWidgetLabel"],
        .top-filter-card div[data-testid="stWidgetLabel"] p,
        div[data-testid="stWidgetLabel"],
        div[data-testid="stWidgetLabel"] p {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            opacity: 1 !important;
            font-weight: 900 !important;
        }

        /* Radio buttons: Moneda local / Pesos mexicanos */
        .top-filter-card div[role="radiogroup"] label,
        .top-filter-card div[role="radiogroup"] p,
        .top-filter-card div[role="radiogroup"] span,
        .top-filter-card div[data-baseweb="radio"] label,
        .top-filter-card div[data-baseweb="radio"] p,
        .top-filter-card div[data-baseweb="radio"] span {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            opacity: 1 !important;
            font-weight: 800 !important;
        }

        /* Selectbox y campos */
        .top-filter-card div[data-baseweb="select"] > div,
        div[data-baseweb="select"] > div,
        input,
        textarea,
        select {
            background-color: #0b1220 !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            border: 1.5px solid rgba(219,234,254,0.70) !important;
            opacity: 1 !important;
        }

        .top-filter-card div[data-baseweb="select"] span,
        .top-filter-card div[data-baseweb="select"] div,
        .top-filter-card div[data-baseweb="select"] input,
        div[data-baseweb="select"] span,
        div[data-baseweb="select"] div,
        div[data-baseweb="select"] input {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            opacity: 1 !important;
            font-weight: 800 !important;
        }

        .top-filter-card div[data-baseweb="select"] svg,
        div[data-baseweb="select"] svg {
            fill: #ffffff !important;
            color: #ffffff !important;
            opacity: 1 !important;
        }

        /* Menú desplegable abierto */
        div[role="listbox"],
        ul[role="listbox"] {
            background-color: #0b1220 !important;
            color: #ffffff !important;
            border: 1px solid rgba(219,234,254,0.50) !important;
        }

        div[role="option"],
        li[role="option"] {
            background-color: #0b1220 !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
        }

        div[role="option"]:hover,
        li[role="option"]:hover {
            background-color: #1e293b !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
        }

        /* Botones: mejor contraste que el azul claro con texto blanco */
        div.stButton > button,
        div[data-testid="stDownloadButton"] > button,
        .top-filter-card div.stButton > button {
            background: var(--dark-button-bg) !important;
            color: var(--dark-button-text) !important;
            -webkit-text-fill-color: var(--dark-button-text) !important;
            border: 1px solid var(--dark-button-bg) !important;
            opacity: 1 !important;
            font-weight: 900 !important;
        }

        div.stButton > button p,
        div.stButton > button span,
        div[data-testid="stDownloadButton"] > button p,
        div[data-testid="stDownloadButton"] > button span,
        .top-filter-card div.stButton > button p,
        .top-filter-card div.stButton > button span {
            color: var(--dark-button-text) !important;
            -webkit-text-fill-color: var(--dark-button-text) !important;
            opacity: 1 !important;
            font-weight: 900 !important;
        }

        div.stButton > button:hover,
        div[data-testid="stDownloadButton"] > button:hover {
            background: #facc15 !important;
            color: #082567 !important;
            -webkit-text-fill-color: #082567 !important;
            border-color: #facc15 !important;
        }

        /* Pastilla unidad seleccionada */
        .unidad-seleccionada-pill {
            background: var(--dark-button-bg) !important;
            color: var(--dark-button-text) !important;
            -webkit-text-fill-color: var(--dark-button-text) !important;
            border: 1px solid var(--dark-button-bg) !important;
            opacity: 1 !important;
            font-weight: 900 !important;
        }

        /* Comentarios */
        .comentario-amplio {
            background: rgba(30,41,59,0.98) !important;
            color: #ffffff !important;
            border-left-color: var(--dark-accent) !important;
            box-shadow: 0 10px 26px rgba(0,0,0,0.38) !important;
        }

        .comentario-amplio,
        .comentario-amplio-texto,
        .comentario-amplio strong {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            opacity: 1 !important;
        }

        .comentario-amplio-titulo {
            background: var(--dark-button-bg) !important;
            color: var(--dark-button-text) !important;
            -webkit-text-fill-color: var(--dark-button-text) !important;
            opacity: 1 !important;
        }

        /* Dataframes / tablas en modo oscuro */
        div[data-testid="stDataFrame"] *,
        div[data-testid="stTable"] * {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
        }

        /* ============================================================
           Plotly en modo oscuro
           Mejora contraste SOLO de textos que quedan sobre fondo oscuro.
           No se cambia el modo claro.
        ============================================================ */

        /* Ejes, títulos de ejes, leyendas y etiquetas externas de pie/dona. */
        div[data-testid="stPlotlyChart"] .xtick text,
        div[data-testid="stPlotlyChart"] .ytick text,
        div[data-testid="stPlotlyChart"] .g-xtitle text,
        div[data-testid="stPlotlyChart"] .g-ytitle text,
        div[data-testid="stPlotlyChart"] .legend text,
        div[data-testid="stPlotlyChart"] .legendtext,
        div[data-testid="stPlotlyChart"] .pielabel text,
        div[data-testid="stPlotlyChart"] .slicetext {
            fill: #ffffff !important;
            color: #ffffff !important;
            opacity: 1 !important;
            paint-order: stroke !important;
            stroke: rgba(4, 10, 25, 0.95) !important;
            stroke-width: 3px !important;
            stroke-linejoin: round !important;
            font-weight: 800 !important;
        }

        /* Números dentro de barras claras: se mantienen oscuros para que no se pierdan. */
        div[data-testid="stPlotlyChart"] .bartext {
            opacity: 1 !important;
            paint-order: stroke !important;
            stroke: rgba(255,255,255,0.95) !important;
            stroke-width: 1.8px !important;
            stroke-linejoin: round !important;
            font-weight: 800 !important;
        }

        /* Anotaciones de variación en cajas blancas: texto azul oscuro legible. */
        div[data-testid="stPlotlyChart"] .annotation text,
        div[data-testid="stPlotlyChart"] .annotation-text {
            fill: #082567 !important;
            color: #082567 !important;
            opacity: 1 !important;
            paint-order: stroke !important;
            stroke: rgba(255,255,255,0.98) !important;
            stroke-width: 2px !important;
            font-weight: 900 !important;
        }

        /* Trazos/textos de línea dentro de áreas blancas: conservar azul oscuro. */
        div[data-testid="stPlotlyChart"] .scatterlayer text {
            opacity: 1 !important;
            paint-order: stroke !important;
            stroke: rgba(255,255,255,0.98) !important;
            stroke-width: 2px !important;
            stroke-linejoin: round !important;
            font-weight: 800 !important;
        }

        /* Cards de entrada */
        .unidad-card {
            background: rgba(17,24,39,0.96) !important;
            border-color: var(--dark-border) !important;
            color: #ffffff !important;
        }

        .unidad-card * {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            opacity: 1 !important;
        }

        /* Métricas */
        .kpi-value {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
        }

        /* Evita que temas/extensiones vuelvan transparentes los textos */
        button,
        button *,
        label,
        label *,
        input,
        textarea {
            opacity: 1 !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ============================================================
# AJUSTE FINAL: FORZAR VISTA EN MODO CLARO
# Este bloque queda al final para sobreescribir cualquier regla de modo oscuro
# del navegador, Windows o extensiones.
# ============================================================
st.markdown(
    """
    <style>
    html, body, .stApp, [data-testid="stAppViewContainer"] {
        color-scheme: light !important;
        background-color: #ffffff !important;
    }

    @media (prefers-color-scheme: dark) {
        :root {
            --app-bg-card: rgba(255,255,255,0.94) !important;
            --app-panel-bg: rgba(255,255,255,0.97) !important;
            --app-panel-solid: #ffffff !important;
            --app-text-main: #111827 !important;
            --app-text-soft: #374151 !important;
            --app-brand: #082567 !important;
            --app-brand-soft: #dbeafe !important;
            --app-accent: #d9c322 !important;
            --app-border: rgba(226,232,240,0.95) !important;
            --app-input-bg: #ffffff !important;
            --app-chart-bg: rgba(255,255,255,0.96) !important;
            --app-shadow: 0 8px 24px rgba(15, 23, 42, 0.09) !important;
        }

        .stApp {
            background-color: #ffffff !important;
            background-image: linear-gradient(rgba(255,255,255,0.82), rgba(255,255,255,0.90)) !important;
        }

        .main .block-container,
        .top-filter-card,
        .kpi-card,
        .unidad-card,
        .top-bottom-opciones-card,
        div[data-testid="stPlotlyChart"],
        div[data-testid="stDataFrame"],
        div[data-testid="stTable"],
        div[data-testid="stMetric"],
        div[data-testid="stExpander"] {
            background: rgba(255,255,255,0.97) !important;
            color: #111827 !important;
            -webkit-text-fill-color: initial !important;
            border-color: rgba(226,232,240,0.95) !important;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.09) !important;
        }

        h1, h2, h3, h4,
        .titulo,
        .top-filter-title,
        .kpi-label,
        .unidad-name,
        .top-bottom-opciones-title,
        .top-filter-card label,
        .top-filter-card label p,
        div[data-testid="stWidgetLabel"],
        div[data-testid="stWidgetLabel"] p {
            color: #082567 !important;
            -webkit-text-fill-color: #082567 !important;
            text-shadow: none !important;
        }

        .subtitulo,
        .landing-subtitle,
        .unidad-help,
        div[data-testid="stMarkdownContainer"] p,
        .top-filter-card p,
        .top-filter-card span,
        .top-filter-card small,
        .kpi-value {
            color: #111827 !important;
            -webkit-text-fill-color: #111827 !important;
            text-shadow: none !important;
        }

        .top-filter-card div[data-baseweb="select"] > div,
        div[data-baseweb="select"] > div,
        div[data-baseweb="input"] > div,
        input, textarea, select {
            background-color: #ffffff !important;
            color: #111827 !important;
            -webkit-text-fill-color: #111827 !important;
            border: 1px solid rgba(226,232,240,0.95) !important;
        }

        .top-filter-card div[data-baseweb="select"] span,
        .top-filter-card div[data-baseweb="select"] div,
        .top-filter-card div[data-baseweb="select"] input,
        div[data-baseweb="select"] span,
        div[data-baseweb="select"] div,
        div[data-baseweb="select"] input {
            color: #111827 !important;
            -webkit-text-fill-color: #111827 !important;
        }

        div[role="listbox"], ul[role="listbox"],
        div[role="option"], li[role="option"] {
            background-color: #ffffff !important;
            color: #111827 !important;
            -webkit-text-fill-color: #111827 !important;
        }

        div.stButton > button,
        div[data-testid="stDownloadButton"] > button,
        .top-filter-card div.stButton > button,
        .unidad-seleccionada-pill {
            background: #082567 !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            border-color: #082567 !important;
        }

        .comentario-amplio,
        .modal-resumen-card,
        .resumen-pagina-card {
            background: #dbeafe !important;
            color: #082567 !important;
            -webkit-text-fill-color: #082567 !important;
            border-left-color: #d9c322 !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# AJUSTE FINAL: COMENTARIOS SIEMPRE LEGIBLES EN MODO CLARO FORZADO
# Este bloque queda al final para evitar que reglas previas de modo oscuro
# dejen el texto blanco sobre tarjetas azules claras.
# ============================================================
st.markdown(
    """
    <style>
    .comentario-amplio,
    .comentario-amplio *,
    .comentario-amplio-texto,
    .comentario-amplio-texto *,
    .modal-resumen-card,
    .modal-resumen-card *,
    .resumen-pagina-card,
    .resumen-pagina-card * {
        background-color: transparent;
        color: #082567 !important;
        -webkit-text-fill-color: #082567 !important;
        text-shadow: none !important;
        opacity: 1 !important;
    }

    .comentario-amplio,
    .modal-resumen-card,
    .resumen-pagina-card {
        background: #dbeafe !important;
        border-left: 8px solid #d9c322 !important;
        box-shadow: 0 8px 22px rgba(15,23,42,0.12) !important;
    }

    .comentario-amplio-texto {
        font-weight: 800 !important;
        line-height: 1.55 !important;
    }
    

    /* Botones visibles cuando Windows/navegador está en modo oscuro, pero el tablero fuerza claro */
    div.stButton > button,
    div[data-testid="stDownloadButton"] > button,
    .top-filter-card div.stButton > button,
    .unidad-seleccionada-pill {
        background: #082567 !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        border: 1px solid #082567 !important;
        opacity: 1 !important;
        font-weight: 900 !important;
        min-height: 42px !important;
    }

    div.stButton > button *,
    div[data-testid="stDownloadButton"] > button *,
    .top-filter-card div.stButton > button *,
    .unidad-seleccionada-pill * {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        opacity: 1 !important;
        font-weight: 900 !important;
    }

    div.stButton > button:hover,
    div[data-testid="stDownloadButton"] > button:hover {
        background: #d9c322 !important;
        color: #082567 !important;
        -webkit-text-fill-color: #082567 !important;
        border-color: #d9c322 !important;
    }

    div.stButton > button:hover *,
    div[data-testid="stDownloadButton"] > button:hover * {
        color: #082567 !important;
        -webkit-text-fill-color: #082567 !important;
    }

    /* Cuadro de movimientos usando la tipografía general de Streamlit */
    .cuadro-movimientos {
        border: 1px solid rgba(8,37,103,0.35);
        border-radius: 10px;
        overflow: hidden;
        background: #ffffff;
        width: 100%;
        box-sizing: border-box;
        font-family: inherit !important;
        color: #111827 !important;
        margin-bottom: 14px;
    }

    .cuadro-header {
        background: #082567;
        color: #ffffff !important;
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 7px 10px;
        font-size: 16px;
        font-weight: 800;
        font-family: inherit !important;
    }

    .cuadro-header * {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
    }

    .cuadro-body {
        padding: 12px 10px 8px 10px;
        font-size: 18px;
        line-height: 1.30;
        color: #111827 !important;
        font-family: inherit !important;
    }

    .fila-mov {
        display: flex;
        justify-content: space-between;
        gap: 12px;
        margin-bottom: 3px;
        white-space: nowrap;
        font-family: inherit !important;
    }

    .valor-rojo {
        color: #dc2626 !important;
        -webkit-text-fill-color: #dc2626 !important;
        min-width: 45px;
        text-align: right;
        font-weight: 900;
    }

    .valor-verde {
        color: #059669 !important;
        -webkit-text-fill-color: #059669 !important;
        min-width: 45px;
        text-align: right;
        font-weight: 900;
    }

    .cuadro-marca-title {
        margin-top: 12px;
        background: #082567;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        border-radius: 10px 10px 0 0;
        padding: 8px 10px;
        font-weight: 900;
        font-family: inherit !important;
    }

    .cuadro-marca-vacio {
        background: #ffffff;
        color: #111827 !important;
        -webkit-text-fill-color: #111827 !important;
        border: 1px solid rgba(8,37,103,0.20);
        border-top: 0;
        border-radius: 0 0 10px 10px;
        padding: 10px;
        font-weight: 700;
    }
</style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# AJUSTE PUNTUAL V4: BOTONES DEL RESUMEN Y TABLAS COMPACTAS
# No modifica lógica, solo visibilidad de botones y altura de tablas.
# ============================================================
st.markdown(
    """
    <style>
    /* Botones visibles también cuando la computadora está en modo claro */
    .stApp button[kind="primary"],
    .stApp button[kind="secondary"],
    .stApp div.stButton > button,
    .stApp div[data-testid="stDownloadButton"] > button,
    .stApp .top-filter-card div.stButton > button {
        background-color: #082567 !important;
        background: #082567 !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        border: 1px solid #082567 !important;
        border-radius: 12px !important;
        opacity: 1 !important;
        visibility: visible !important;
        box-shadow: 0 6px 16px rgba(8,37,103,0.22) !important;
    }

    .stApp button[kind="primary"] *,
    .stApp button[kind="secondary"] *,
    .stApp div.stButton > button *,
    .stApp div[data-testid="stDownloadButton"] > button *,
    .stApp .top-filter-card div.stButton > button * {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        opacity: 1 !important;
        visibility: visible !important;
        font-weight: 900 !important;
    }

    .stApp button[kind="primary"]:hover,
    .stApp button[kind="secondary"]:hover,
    .stApp div.stButton > button:hover,
    .stApp div[data-testid="stDownloadButton"] > button:hover {
        background-color: #d9c322 !important;
        background: #d9c322 !important;
        color: #082567 !important;
        -webkit-text-fill-color: #082567 !important;
        border-color: #d9c322 !important;
    }

    .stApp button[kind="primary"]:hover *,
    .stApp button[kind="secondary"]:hover *,
    .stApp div.stButton > button:hover *,
    .stApp div[data-testid="stDownloadButton"] > button:hover * {
        color: #082567 !important;
        -webkit-text-fill-color: #082567 !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ============================================================
# AJUSTE FINAL: FONDO BLANCO FORZADO + IMAGEN COMO MARCA DE AGUA
# TAMBIÉN CUANDO WINDOWS/NAVEGADOR ESTÁ EN MODO OSCURO
# ============================================================
ruta_imagen_marca_agua = buscar_imagen_fondo(NOMBRE_IMAGEN_FONDO)
if ruta_imagen_marca_agua is not None:
    fondo_marca_agua_base64 = imagen_a_base64(str(ruta_imagen_marca_agua))
    st.markdown(
        f"""
        <style>
        html,
        body,
        .stApp,
        [data-testid="stAppViewContainer"] {{
            color-scheme: light !important;
            background-color: #ffffff !important;
        }}

        .stApp,
        [data-testid="stAppViewContainer"] {{
            background-image:
                linear-gradient(rgba(255,255,255,0.82), rgba(255,255,255,0.91)),
                url("data:image/png;base64,{fondo_marca_agua_base64}") !important;
            background-size: cover !important;
            background-position: top center !important;
            background-repeat: no-repeat !important;
            background-attachment: fixed !important;
            background-color: #ffffff !important;
        }}

        @media (prefers-color-scheme: dark) {{
            html,
            body,
            .stApp,
            [data-testid="stAppViewContainer"] {{
                color-scheme: light !important;
                background-color: #ffffff !important;
            }}

            .stApp,
            [data-testid="stAppViewContainer"] {{
                background-image:
                    linear-gradient(rgba(255,255,255,0.82), rgba(255,255,255,0.91)),
                    url("data:image/png;base64,{fondo_marca_agua_base64}") !important;
                background-size: cover !important;
                background-position: top center !important;
                background-repeat: no-repeat !important;
                background-attachment: fixed !important;
                background-color: #ffffff !important;
            }}

            header[data-testid="stHeader"] {{
                background: rgba(255,255,255,0) !important;
            }}
        }}
        </style>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# MAQUETA FINAL DE GESTIÓN EN FORMATO PRESENTACIÓN
# Esta regla va al final para que el carrusel no herede el estilo azul
# global de los demás botones del tablero.
# ============================================================
if globals().get("modulo_seleccionado") == "Cartera":
    miniatura_activa = str(
        st.session_state.get("gestion_seccion_activa", "kpis")
    ).replace("-", "_")

    st.markdown(
        f"""
        <style>
        /* La página nunca genera desplazamiento lateral; solamente la tira inferior. */
        html, body, .stApp, [data-testid="stAppViewContainer"],
        [data-testid="stAppViewContainer"] > .main {{
            max-width:100% !important;
            overflow-x:hidden !important;
        }}

        [data-testid="stAppViewContainer"] .main .block-container {{
            width:100% !important;
            max-width:100% !important;
            padding-left:clamp(14px,2vw,30px) !important;
            padding-right:clamp(14px,2vw,30px) !important;
            padding-bottom:18px !important;
            box-sizing:border-box !important;
        }}

        /* Una sola diapositiva grande, ajustada al ancho y sin scroll interno. */
        [class*="st-key-gestion_slide_"] {{
            display:none !important;
        }}

        .st-key-gestion_slide_{miniatura_activa} {{
            display:block !important;
            position:relative !important;
            width:100% !important;
            max-width:100% !important;
            min-height:clamp(470px,58vh,650px) !important;
            height:auto !important;
            max-height:none !important;
            overflow:visible !important;
            box-sizing:border-box !important;
            padding:clamp(18px,2.1vw,30px) clamp(16px,2.4vw,34px) 28px !important;
            margin:8px 0 12px !important;
            border:1px solid #d9e2ee !important;
            border-radius:16px !important;
            background:rgba(255,255,255,.985) !important;
            box-shadow:0 12px 32px rgba(15,23,42,.11) !important;
        }}

        .st-key-gestion_slide_{miniatura_activa} > div,
        .st-key-gestion_slide_{miniatura_activa} > div > div {{
            position:relative;
            z-index:1;
        }}

        .st-key-gestion_slide_{miniatura_activa} > div[data-testid="stVerticalBlock"] {{
            gap:.65rem !important;
        }}

        .st-key-gestion_slide_{miniatura_activa} h1,
        .st-key-gestion_slide_{miniatura_activa} h2,
        .st-key-gestion_slide_{miniatura_activa} h3 {{
            margin-top:.2rem !important;
            margin-bottom:.35rem !important;
        }}

        .st-key-gestion_slide_{miniatura_activa} [data-testid="stPlotlyChart"],
        .st-key-gestion_slide_{miniatura_activa} [data-testid="stPlotlyChart"] > div,
        .st-key-gestion_slide_{miniatura_activa} .js-plotly-plot,
        .st-key-gestion_slide_{miniatura_activa} .plot-container {{
            width:100% !important;
            max-width:100% !important;
        }}

        .st-key-gestion_slide_{miniatura_activa} [data-testid="stDataFrame"],
        .st-key-gestion_slide_{miniatura_activa} [data-testid="stTable"] {{
            width:100% !important;
            max-width:100% !important;
        }}

        /* Franja de miniaturas: la única zona con desplazamiento horizontal. */
        .st-key-gestion_carrusel {{
            display:block !important;
            width:100% !important;
            max-width:100% !important;
            overflow-x:auto !important;
            overflow-y:hidden !important;
            overscroll-behavior-inline:contain;
            scrollbar-width:thin;
            scrollbar-color:#8fa8c8 transparent;
            padding:2px 1px 10px !important;
            margin:0 !important;
        }}

        .st-key-gestion_carrusel::-webkit-scrollbar {{ height:8px; }}
        .st-key-gestion_carrusel::-webkit-scrollbar-track {{ background:transparent; }}
        .st-key-gestion_carrusel::-webkit-scrollbar-thumb {{
            background:#8fa8c8; border-radius:999px;
        }}

        .st-key-gestion_carrusel > div,
        .st-key-gestion_carrusel > div > div {{
            width:max-content !important;
            min-width:100% !important;
            max-width:none !important;
        }}

        .st-key-gestion_carrusel [data-testid="stHorizontalBlock"] {{
            display:flex !important;
            flex-flow:row nowrap !important;
            align-items:stretch !important;
            width:max-content !important;
            min-width:100% !important;
            max-width:none !important;
            gap:10px !important;
        }}

        .st-key-gestion_carrusel [data-testid="column"],
        .st-key-gestion_carrusel [data-testid="stColumn"] {{
            flex:0 0 clamp(154px,13.2vw,190px) !important;
            width:clamp(154px,13.2vw,190px) !important;
            min-width:154px !important;
            max-width:190px !important;
        }}

        /* Miniaturas blancas con esquinas corporativas, no botones azules. */
        .stApp .st-key-gestion_carrusel div.stButton > button,
        .stApp .st-key-gestion_carrusel button,
        .stApp [class*="st-key-gestion_mini_"] div.stButton > button,
        .stApp [class*="st-key-gestion_mini_"] button {{
            position:relative !important;
            display:flex !important;
            align-items:flex-start !important;
            justify-content:flex-start !important;
            width:100% !important;
            min-width:0 !important;
            height:108px !important;
            min-height:108px !important;
            max-height:108px !important;
            padding:10px 10px 58px 11px !important;
            overflow:hidden !important;
            white-space:normal !important;
            text-align:left !important;
            line-height:1.1 !important;
            color:#082567 !important;
            -webkit-text-fill-color:#082567 !important;
            border:2px solid #d9e2ee !important;
            border-radius:11px !important;
            background:
                radial-gradient(circle at -8% -18%, transparent 0 45px, #f0cf2c 46px 57px, transparent 58px),
                radial-gradient(circle at 110% 125%, transparent 0 51px, #082567 52px 70px, transparent 71px),
                linear-gradient(145deg,#ffffff 0%,#f7faff 100%) !important;
            box-shadow:0 4px 12px rgba(15,23,42,.08) !important;
            opacity:1 !important;
            visibility:visible !important;
            transform:none !important;
        }}

        .stApp .st-key-gestion_carrusel div.stButton > button:hover,
        .stApp .st-key-gestion_carrusel button:hover,
        .stApp [class*="st-key-gestion_mini_"] div.stButton > button:hover,
        .stApp [class*="st-key-gestion_mini_"] button:hover {{
            color:#082567 !important;
            -webkit-text-fill-color:#082567 !important;
            border-color:#d5b91b !important;
            background:
                radial-gradient(circle at -8% -18%, transparent 0 45px, #f0cf2c 46px 57px, transparent 58px),
                radial-gradient(circle at 110% 125%, transparent 0 51px, #082567 52px 70px, transparent 71px),
                linear-gradient(145deg,#ffffff 0%,#fffdf1 100%) !important;
            box-shadow:0 7px 18px rgba(15,23,42,.13) !important;
            transform:translateY(-1px) !important;
        }}

        .stApp .st-key-gestion_carrusel button *,
        .stApp [class*="st-key-gestion_mini_"] button *,
        .stApp .st-key-gestion_carrusel button:hover *,
        .stApp [class*="st-key-gestion_mini_"] button:hover * {{
            color:#082567 !important;
            -webkit-text-fill-color:#082567 !important;
            font-size:11px !important;
            font-weight:900 !important;
            line-height:1.08 !important;
            opacity:1 !important;
            visibility:visible !important;
        }}

        .stApp [class*="st-key-gestion_mini_"] button::before {{
            position:absolute;
            left:8px;
            bottom:7px;
            z-index:4;
            display:flex;
            align-items:center;
            justify-content:center;
            width:24px;
            height:24px;
            border-radius:50%;
            background:#082567;
            color:#ffffff;
            -webkit-text-fill-color:#ffffff;
            font-size:12px;
            font-weight:950;
            box-shadow:0 2px 5px rgba(8,37,103,.22);
        }}

        .stApp [class*="st-key-gestion_mini_"] button::after {{
            content:"";
            position:absolute;
            left:42px;
            right:13px;
            bottom:13px;
            z-index:2;
            height:34px;
            border-radius:5px;
            opacity:.92;
        }}

        .st-key-gestion_mini_resumen button::before {{ content:"1"; }}
        .st-key-gestion_mini_kpis button::before {{ content:"2"; }}
        .st-key-gestion_mini_evolucion button::before {{ content:"3"; }}
        .st-key-gestion_mini_coordinadoras button::before {{ content:"4"; }}
        .st-key-gestion_mini_movimientos button::before {{ content:"5"; }}
        .st-key-gestion_mini_top_bottom button::before {{ content:"6"; }}
        .st-key-gestion_mini_conclusiones button::before {{ content:"7"; }}

        .st-key-gestion_mini_resumen button::after {{
            left:56px !important; right:auto !important; bottom:10px !important;
            width:40px; height:40px; border-radius:50% !important;
            background:conic-gradient(#0b70c9 0 41%,#f49aa0 41% 63%,#76b7ea 63% 84%,#ff2d2d 84% 100%);
            box-shadow:inset 0 0 0 10px #ffffff;
        }}

        .st-key-gestion_mini_kpis button::after,
        .st-key-gestion_mini_coordinadoras button::after {{
            background:
                linear-gradient(to top,#0b70c9 0 74%,transparent 74%) 2% 100%/11% 100% no-repeat,
                linear-gradient(to top,#f0cf2c 0 46%,transparent 46%) 22% 100%/11% 100% no-repeat,
                linear-gradient(to top,#76b7ea 0 88%,transparent 88%) 42% 100%/11% 100% no-repeat,
                linear-gradient(to top,#0b70c9 0 60%,transparent 60%) 62% 100%/11% 100% no-repeat,
                linear-gradient(to top,#f0cf2c 0 80%,transparent 80%) 82% 100%/11% 100% no-repeat;
            border-bottom:1px solid #b8c5d6;
        }}

        .st-key-gestion_mini_evolucion button::after {{
            background:
                linear-gradient(158deg,transparent 0 21%,#0b70c9 22% 25%,transparent 26% 43%,#0b70c9 44% 47%,transparent 48% 64%,#f0cf2c 65% 69%,transparent 70%),
                repeating-linear-gradient(0deg,transparent 0 10px,rgba(148,163,184,.22) 10px 11px);
        }}

        .st-key-gestion_mini_movimientos button::after {{
            background:
                repeating-linear-gradient(0deg,rgba(255,255,255,.38) 0 7px,transparent 7px 9px),
                repeating-linear-gradient(90deg,#d7e6f7 0 13px,#7ea9d8 13px 26px,#173b73 26px 39px);
        }}

        .st-key-gestion_mini_top_bottom button::after {{
            background:
                linear-gradient(90deg,#0b70c9 0 75%,transparent 75%) 0 1px/48% 6px no-repeat,
                linear-gradient(90deg,#0b70c9 0 57%,transparent 57%) 0 13px/48% 6px no-repeat,
                linear-gradient(90deg,#0b70c9 0 88%,transparent 88%) 0 25px/48% 6px no-repeat,
                linear-gradient(90deg,#f0cf2c 0 83%,transparent 83%) 100% 1px/45% 6px no-repeat,
                linear-gradient(90deg,#f0cf2c 0 61%,transparent 61%) 100% 13px/45% 6px no-repeat,
                linear-gradient(90deg,#f0cf2c 0 72%,transparent 72%) 100% 25px/45% 6px no-repeat;
        }}

        .st-key-gestion_mini_conclusiones button::after {{
            background:
                radial-gradient(circle,#1d7f64 0 4px,transparent 5px) 0 2px/10px 10px no-repeat,
                radial-gradient(circle,#0b70c9 0 4px,transparent 5px) 0 15px/10px 10px no-repeat,
                radial-gradient(circle,#d9b600 0 4px,transparent 5px) 0 28px/10px 10px no-repeat,
                linear-gradient(90deg,#94a3b8 0 82%,transparent 82%) 16px 4px/calc(100% - 16px) 4px no-repeat,
                linear-gradient(90deg,#94a3b8 0 64%,transparent 64%) 16px 17px/calc(100% - 16px) 4px no-repeat,
                linear-gradient(90deg,#94a3b8 0 74%,transparent 74%) 16px 30px/calc(100% - 16px) 4px no-repeat;
        }}

        .stApp .st-key-gestion_mini_{miniatura_activa} button {{
            border-color:#0b70c9 !important;
            box-shadow:0 0 0 2px rgba(11,112,201,.13),0 7px 18px rgba(15,23,42,.12) !important;
        }}

        @media (max-width:900px) {{
            .st-key-gestion_slide_{miniatura_activa} {{
                min-height:0 !important;
                max-height:none !important;
                overflow:visible !important;
                padding:17px 13px 22px !important;
            }}

            .st-key-gestion_slide_{miniatura_activa} [data-testid="stHorizontalBlock"] {{
                flex-wrap:wrap !important;
            }}

            .st-key-gestion_carrusel [data-testid="stHorizontalBlock"] {{
                flex-wrap:nowrap !important;
            }}

            .st-key-gestion_carrusel [data-testid="column"],
            .st-key-gestion_carrusel [data-testid="stColumn"] {{
                flex:0 0 158px !important;
                width:158px !important;
                min-width:158px !important;
                max-width:158px !important;
            }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

# Regla definitiva: la tarjeta seleccionada usa todo el ancho, no tiene
# adornos circulares y nunca crea un scroll interno.
if globals().get("modulo_seleccionado") == "Cartera":
    seccion_final_ajustada = str(
        st.session_state.get("gestion_seccion_activa", "kpis")
    ).replace("-", "_")
    st.markdown(
        f"""
        <style>
        [class*="st-key-gestion_slide_"]::before,
        [class*="st-key-gestion_slide_"]::after {{
            content:none !important;
            display:none !important;
            border:0 !important;
            background:none !important;
        }}
        .st-key-gestion_slide_{seccion_final_ajustada} {{
            width:100% !important;
            max-width:100% !important;
            min-width:0 !important;
            min-height:0 !important;
            height:auto !important;
            max-height:none !important;
            overflow:visible !important;
            margin:6px 0 10px !important;
            padding:clamp(14px,1.7vw,24px) clamp(14px,2vw,28px) 22px !important;
            border-radius:12px !important;
            background:#ffffff !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# ENCABEZADO COMPACTO: CONTROLES DENTRO DE LA MISMA TARJETA
# ============================================================
st.markdown(
    """
    <style>
    .gestion-encabezado {
        position:relative !important;
        min-height:78px !important;
        padding-top:11px !important;
        padding-bottom:11px !important;
        margin-bottom:0 !important;
        box-sizing:border-box !important;
    }

    /* Barra real en una sola fila: marca, acciones centradas y avance. */
    .st-key-gestion_header_barra {
        width:100% !important;
        max-width:100% !important;
        margin:0 0 3px !important;
        padding:8px 14px !important;
        box-sizing:border-box !important;
        background:rgba(255,255,255,.97) !important;
        border:1px solid #dbe3ee !important;
        border-radius:15px !important;
        box-shadow:0 7px 20px rgba(15,23,42,.065) !important;
    }

    .st-key-gestion_header_barra > [data-testid="stVerticalBlock"] {
        gap:0 !important;
    }

    .st-key-gestion_header_barra [data-testid="stHorizontalBlock"] {
        align-items:center !important;
    }

    [data-testid="stVerticalBlock"]:has(> .st-key-gestion_header_barra) {
        gap:.32rem !important;
    }

    .st-key-gestion_header_barra .gestion-marca {
        min-width:0 !important;
        gap:10px !important;
    }

    .st-key-gestion_header_barra .gestion-titulo {
        font-size:21px !important;
    }

    .st-key-gestion_header_barra .gestion-unidad {
        margin-top:3px !important;
        font-size:10.5px !important;
    }

    .st-key-gestion_header_barra .gestion-progreso {
        width:100% !important;
        max-width:none !important;
        justify-content:flex-end !important;
    }

    .st-key-gestion_header_barra .gestion-paso {
        gap:6px !important;
        font-size:10.5px !important;
    }

    .st-key-gestion_header_barra .gestion-paso span {
        width:26px !important;
        min-width:26px !important;
        height:26px !important;
        font-size:11px !important;
    }

    .st-key-gestion_header_barra .gestion-trazo {
        width:32px !important;
        margin:0 7px !important;
    }

    .stApp .st-key-gestion_header_barra button {
        min-height:34px !important;
        height:34px !important;
        padding:0 10px !important;
        border-radius:9px !important;
        font-size:10.5px !important;
        box-shadow:0 2px 6px rgba(15,23,42,.07) !important;
    }

    .stApp .st-key-gestion_header_barra button * {
        font-size:10.5px !important;
        line-height:1 !important;
    }

    .stApp .st-key-gestion_header_barra [data-testid="stPopover"] > button {
        width:36px !important;
        min-width:36px !important;
        max-width:36px !important;
        padding:0 !important;
        background:#eef5fd !important;
        color:#082567 !important;
        -webkit-text-fill-color:#082567 !important;
        border:1px solid #b9cee7 !important;
    }

    @media (min-width:1101px) {
        .st-key-ayuda_tablero {
            position:relative !important;
            z-index:40 !important;
            width:clamp(400px,25vw,450px) !important;
            max-width:450px !important;
            height:0 !important;
            min-height:0 !important;
            overflow:visible !important;
            margin:0 auto !important;
            padding:0 !important;
            transform:translateY(-59px) !important;
        }
        .st-key-ayuda_tablero > div,
        .st-key-ayuda_tablero > div > div {
            overflow:visible !important;
        }

        /* El contenedor de acciones no reserva otra fila debajo del encabezado. */
        [data-testid="stElementContainer"]:has(> .st-key-ayuda_tablero),
        [data-testid="stElementContainer"]:has(.st-key-ayuda_tablero) {
            min-height:0 !important;
            height:0 !important;
            margin:0 !important;
            padding:0 !important;
            overflow:visible !important;
        }
    }

    .st-key-ayuda_tablero [data-testid="stHorizontalBlock"] {
        gap:7px !important;
        align-items:center !important;
        justify-content:center !important;
    }

    /* Base común: botones ligeros, integrados con la tarjeta blanca. */
    .stApp .st-key-ayuda_tablero button {
        height:36px !important;
        min-height:36px !important;
        padding:0 13px !important;
        border-radius:9px !important;
        background:linear-gradient(180deg,#ffffff 0%,#f3f7fc 100%) !important;
        color:#082567 !important;
        -webkit-text-fill-color:#082567 !important;
        border:1px solid #cfdae8 !important;
        box-shadow:0 2px 6px rgba(15,23,42,.07) !important;
        font-size:11.5px !important;
        font-weight:850 !important;
        letter-spacing:0 !important;
    }
    .stApp .st-key-ayuda_tablero button * {
        color:#082567 !important;
        -webkit-text-fill-color:#082567 !important;
        font-size:11.5px !important;
        font-weight:850 !important;
    }

    /* Información: control corto con la misma geometría del grupo. */
    .stApp .st-key-ayuda_tablero [data-testid="stHorizontalBlock"] > div:nth-child(1) button,
    .stApp .st-key-ayuda_tablero [data-testid="stPopover"] > button {
        width:38px !important;
        min-width:38px !important;
        max-width:38px !important;
        padding:0 !important;
        background:#eef5fd !important;
        border-color:#b9cee7 !important;
        font-size:17px !important;
    }

    /* Acción principal: marfil con acento dorado, sin bloque azul pesado. */
    .stApp .st-key-ayuda_tablero [data-testid="stHorizontalBlock"] > div:nth-child(2) button,
    .stApp div[class*="st-key-btn_abrir_resumen_pais"] button {
        background:linear-gradient(180deg,#fffef8 0%,#fff8d9 100%) !important;
        border-color:#e2c44b !important;
        box-shadow:inset 3px 0 0 #e7c42d,0 2px 6px rgba(15,23,42,.07) !important;
    }

    /* Acción secundaria: azul muy claro, coherente con el paso completado. */
    .stApp .st-key-ayuda_tablero [data-testid="stHorizontalBlock"] > div:nth-child(3) button,
    .stApp div[class*="st-key-btn_cambiar_unidad"] button {
        background:linear-gradient(180deg,#ffffff 0%,#edf4fc 100%) !important;
        border-color:#b9cee7 !important;
    }

    .stApp .st-key-ayuda_tablero button:hover {
        transform:translateY(-1px) !important;
        border-color:#8eacd0 !important;
        box-shadow:0 4px 9px rgba(15,23,42,.10) !important;
    }

    /* Ritmo compacto y uniforme: encabezado -> indicador -> contenido. */
    [data-testid="stElementContainer"]:has(.gestion-encabezado) {
        margin-bottom:0 !important;
        padding-bottom:0 !important;
    }

    [class*="st-key-gestion_slide_"] {
        margin:0 0 4px !important;
        padding:9px 16px 12px !important;
        border-radius:12px !important;
    }

    [class*="st-key-gestion_slide_"] > [data-testid="stVerticalBlock"],
    [class*="st-key-gestion_slide_"] [data-testid="stVerticalBlock"] {
        gap:.38rem !important;
    }

    [class*="st-key-gestion_slide_"] .gestion-seccion-cabecera {
        min-height:36px !important;
        margin:0 0 3px !important;
        padding:6px 10px !important;
        gap:8px !important;
        border-left-width:4px !important;
        border-radius:0 9px 9px 0 !important;
        font-size:17px !important;
        line-height:1.1 !important;
    }

    [class*="st-key-gestion_slide_"] .gestion-seccion-cabecera span {
        width:25px !important;
        min-width:25px !important;
        height:25px !important;
        font-size:11px !important;
    }

    [class*="st-key-gestion_slide_"] h1,
    [class*="st-key-gestion_slide_"] h2,
    [class*="st-key-gestion_slide_"] h3 {
        margin-top:0 !important;
        margin-bottom:.2rem !important;
        padding-top:0 !important;
    }

    [data-testid="stElementContainer"]:has(.gestion-carrusel-real) {
        margin-top:0 !important;
        padding-top:0 !important;
    }

    .gestion-carrusel-real {
        gap:8px !important;
        margin:0 !important;
        padding:2px 1px 5px !important;
    }

    .gestion-miniatura-real {
        height:116px !important;
    }

    /* Los textos de estos dos botones prevalecen sobre el tema global azul. */
    .stApp div[class*="st-key-btn_abrir_resumen_pais"] button,
    .stApp div[class*="st-key-btn_abrir_resumen_pais"] button p,
    .stApp div[class*="st-key-btn_abrir_resumen_pais"] button span,
    .stApp div[class*="st-key-btn_abrir_resumen_pais"] button div,
    .stApp div[class*="st-key-btn_cambiar_unidad"] button,
    .stApp div[class*="st-key-btn_cambiar_unidad"] button p,
    .stApp div[class*="st-key-btn_cambiar_unidad"] button span,
    .stApp div[class*="st-key-btn_cambiar_unidad"] button div {
        color:#082567 !important;
        -webkit-text-fill-color:#082567 !important;
        opacity:1 !important;
        visibility:visible !important;
        text-shadow:none !important;
    }

    .stApp .st-key-ayuda_tablero [data-testid="stHorizontalBlock"] > div:nth-child(2) button,
    .stApp .st-key-ayuda_tablero [data-testid="stHorizontalBlock"] > div:nth-child(2) button *,
    .stApp .st-key-ayuda_tablero [data-testid="stHorizontalBlock"] > div:nth-child(3) button,
    .stApp .st-key-ayuda_tablero [data-testid="stHorizontalBlock"] > div:nth-child(3) button * {
        color:#082567 !important;
        -webkit-text-fill-color:#082567 !important;
        opacity:1 !important;
        visibility:visible !important;
        text-shadow:none !important;
    }

    /* Matriz estática y fluida: todas las columnas caben, sin scroll horizontal. */
    .matriz-responsiva-sin-scroll {
        width:100% !important;
        max-width:100% !important;
        overflow:visible !important;
        border:1px solid #dbe3ee;
        border-radius:10px;
        background:#ffffff;
        box-sizing:border-box;
    }

    .matriz-responsiva-sin-scroll table {
        width:100% !important;
        max-width:100% !important;
        table-layout:fixed !important;
        border-collapse:separate;
        border-spacing:0;
        color:#111827;
        font-size:clamp(9px,.68vw,12px);
        line-height:1.05;
    }

    .matriz-responsiva-sin-scroll th,
    .matriz-responsiva-sin-scroll td {
        height:32px;
        padding:5px 4px;
        border-right:1px solid #dbe3ee;
        border-bottom:1px solid #dbe3ee;
        text-align:right;
        vertical-align:middle;
        white-space:normal;
        overflow-wrap:anywhere;
        box-sizing:border-box;
    }

    .matriz-responsiva-sin-scroll thead th {
        height:34px;
        background:#f8fafc;
        color:#64748b;
        font-weight:750;
        text-align:left;
    }

    .matriz-responsiva-sin-scroll th:first-child {
        width:15%;
        text-align:left;
        padding-left:8px;
    }

    .matriz-responsiva-sin-scroll tr:last-child > * { border-bottom:0; }
    .matriz-responsiva-sin-scroll tr > *:last-child { border-right:0; }
    .matriz-responsiva-sin-scroll .matriz-fila-total > th {
        background:#082567;
        color:#ffffff !important;
        -webkit-text-fill-color:#ffffff !important;
        font-weight:900;
    }
    .matriz-responsiva-sin-scroll .matriz-total {
        background:#082567;
        color:#ffffff !important;
        -webkit-text-fill-color:#ffffff !important;
        font-weight:900;
    }
    .matriz-responsiva-sin-scroll .matriz-nueva {
        background:#dbeafe;
        color:#1d4ed8 !important;
        -webkit-text-fill-color:#1d4ed8 !important;
        font-weight:850;
    }
    .matriz-responsiva-sin-scroll .matriz-baja {
        background:#ffedd5;
        color:#c2410c !important;
        -webkit-text-fill-color:#c2410c !important;
        font-weight:850;
    }
    .matriz-responsiva-sin-scroll .matriz-mejora {
        color:#059669 !important;
        -webkit-text-fill-color:#059669 !important;
        font-weight:850;
    }
    .matriz-responsiva-sin-scroll .matriz-retroceso {
        color:#dc2626 !important;
        -webkit-text-fill-color:#dc2626 !important;
        font-weight:850;
    }
    .matriz-responsiva-sin-scroll .matriz-permanencia {
        background:#f1f5f9;
        color:#111827 !important;
        -webkit-text-fill-color:#111827 !important;
        font-weight:800;
    }

    /* La tabla lateral también crece con sus filas y no crea un viewport propio. */
    .tabla-marca-responsiva {
        width:100%;
        overflow:visible;
        border:1px solid #dbe3ee;
        border-top:0;
        border-radius:0 0 9px 9px;
        background:#ffffff;
    }
    .tabla-marca-compacta {
        width:100%;
        table-layout:fixed;
        border-collapse:collapse;
        font-size:11px;
        color:#334155;
    }
    .tabla-marca-compacta th,
    .tabla-marca-compacta td {
        padding:5px 7px;
        border-bottom:1px solid #e2e8f0;
        text-align:left;
        overflow-wrap:anywhere;
    }
    .tabla-marca-compacta th {
        background:#f8fafc;
        color:#64748b;
        font-weight:750;
    }
    .tabla-marca-compacta th:last-child,
    .tabla-marca-compacta td:last-child {
        width:62%;
        text-align:right;
    }
    .tabla-marca-compacta tr:last-child td { border-bottom:0; }

    /* Ajustes de densidad exclusivos para la diapositiva de movimientos. */
    .st-key-gestion_slide_movimientos {
        overflow:visible !important;
        padding:8px 14px 10px !important;
    }
    .st-key-gestion_slide_movimientos > [data-testid="stVerticalBlock"],
    .st-key-gestion_slide_movimientos [data-testid="stVerticalBlock"] {
        gap:.28rem !important;
    }
    .st-key-gestion_slide_movimientos h2,
    .st-key-gestion_slide_movimientos h3 {
        font-size:clamp(19px,1.45vw,26px) !important;
        line-height:1.05 !important;
        margin:0 !important;
        padding:0 !important;
    }
    .st-key-gestion_slide_movimientos .movimientos-subtitulo {
        margin:0 0 2px !important;
        color:#082567 !important;
        -webkit-text-fill-color:#082567 !important;
        font-size:clamp(16px,1.2vw,21px) !important;
        font-weight:900 !important;
        line-height:1.05 !important;
        letter-spacing:-.01em !important;
    }
    .st-key-gestion_slide_movimientos [data-testid="stWidgetLabel"] p,
    .st-key-gestion_slide_movimientos [data-testid="stCaptionContainer"] p {
        font-size:11px !important;
        line-height:1.15 !important;
        margin:0 !important;
    }
    .st-key-gestion_slide_movimientos div[data-baseweb="select"] > div {
        min-height:34px !important;
        height:34px !important;
    }
    .st-key-gestion_slide_movimientos [data-testid="stMetric"] {
        min-height:68px !important;
        padding:7px 12px !important;
        border-radius:10px !important;
    }
    .st-key-gestion_slide_movimientos [data-testid="stMetricLabel"] p {
        font-size:11px !important;
        line-height:1.05 !important;
    }
    .st-key-gestion_slide_movimientos [data-testid="stMetricValue"] {
        font-size:clamp(22px,1.7vw,30px) !important;
        line-height:1 !important;
    }
    .st-key-gestion_slide_movimientos .cuadro-movimientos {
        margin-bottom:5px !important;
        border-radius:9px !important;
    }
    .st-key-gestion_slide_movimientos .cuadro-header {
        padding:5px 8px !important;
        font-size:12px !important;
    }
    .st-key-gestion_slide_movimientos .cuadro-body {
        padding:7px 8px 5px !important;
        font-size:13px !important;
        line-height:1.13 !important;
    }
    .st-key-gestion_slide_movimientos .fila-mov {
        margin-bottom:2px !important;
        gap:6px !important;
    }
    .st-key-gestion_slide_movimientos .cuadro-marca-title {
        margin-top:4px !important;
        padding:6px 8px !important;
        min-height:30px !important;
        display:flex !important;
        align-items:center !important;
        box-sizing:border-box !important;
        border-radius:9px 9px 0 0 !important;
        font-size:11px !important;
        line-height:1.15 !important;
    }
    .st-key-gestion_slide_movimientos div[data-testid="stDownloadButton"] > button {
        min-height:32px !important;
        height:32px !important;
        padding:0 11px !important;
        border-radius:8px !important;
        font-size:10px !important;
        box-shadow:0 3px 8px rgba(8,37,103,.14) !important;
    }
    .st-key-gestion_slide_movimientos div[data-testid="stDownloadButton"] > button * {
        font-size:10px !important;
        line-height:1 !important;
    }
    .st-key-gestion_slide_movimientos .comentario-amplio {
        width:min(82%,1180px) !important;
        margin:3px auto 0 !important;
        padding:8px 14px !important;
        border-left-width:5px !important;
        border-radius:12px !important;
        box-shadow:0 4px 12px rgba(15,23,42,.08) !important;
    }
    .st-key-gestion_slide_movimientos .comentario-amplio-texto {
        font-size:clamp(10px,.8vw,12px) !important;
        line-height:1.35 !important;
        font-weight:750 !important;
    }

    @media (max-width:1450px) and (min-width:1101px) {
        .st-key-gestion_header_barra .gestion-paso {
            gap:4px !important;
            font-size:9px !important;
        }
        .st-key-gestion_header_barra .gestion-paso span {
            width:23px !important;
            min-width:23px !important;
            height:23px !important;
            font-size:9px !important;
        }
        .st-key-gestion_header_barra .gestion-trazo {
            width:20px !important;
            margin:0 4px !important;
        }
        .stApp .st-key-gestion_header_barra button,
        .stApp .st-key-gestion_header_barra button * {
            font-size:9.5px !important;
        }
    }

    @media (max-width:1300px) and (min-width:901px) {
        .matriz-responsiva-sin-scroll table { font-size:8.5px; }
        .matriz-responsiva-sin-scroll th,
        .matriz-responsiva-sin-scroll td { padding:4px 2px; }
        .st-key-ayuda_tablero { width:390px !important; }
    }

    @media (max-width:1100px) {
        .gestion-encabezado {
            min-height:0 !important;
            margin-bottom:6px !important;
        }
        .st-key-ayuda_tablero {
            width:100% !important;
            max-width:100% !important;
            height:auto !important;
            min-height:0 !important;
            margin:0 0 5px !important;
            transform:none !important;
        }
        [class*="st-key-gestion_slide_"] {
            padding:8px 10px 11px !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)
