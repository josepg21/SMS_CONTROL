"""App Streamlit: Seguimiento de prendas SMS.

Importar Excel, editar estados/observaciones, ver dashboard y exportar reportes.
Desplegar en red:
    streamlit run app.py --server.address 0.0.0.0 --server.port 8501
Otras máquinas acceden:  http://IP_DE_ESTA_MAQUINA:8501
"""
import os
import time
import datetime

import streamlit as st

import datos
import db
import reporte

st.set_page_config(page_title='Seguimiento SMS', layout='wide')


def inyectar_css():
    st.markdown("""
    <style>
    /* ===== Paleta neutra blanco/negro (diseño) =====
       Los GRÁFICOS e INDICADORES del dashboard usan sus propios
       colores libres (definidos en CHART_COLORS/plotly), no se tocan. */
    :root {
        --sms-bg-top: #0B0B0F;
        --sms-bg-bottom: #16161D;
        --sms-glass: rgba(255,255,255,0.05);
        --sms-glass-strong: rgba(255,255,255,0.09);
        --sms-glass-border: rgba(255,255,255,0.12);
        --sms-text: #F5F5F7;
        --sms-muted: rgba(255,255,255,0.55);
        --sms-border: rgba(255,255,255,0.10);
        --sms-black: #000000;
        --sms-white: #FFFFFF;
    }

    /* Fondo con gradiente suave y textura de "vidrio" */
    .stApp {
        background:
            radial-gradient(1200px 600px at 10% -10%, rgba(255,255,255,0.06), transparent 60%),
            radial-gradient(1000px 500px at 110% 10%, rgba(255,255,255,0.05), transparent 55%),
            linear-gradient(160deg, var(--sms-bg-top), var(--sms-bg-bottom));
    }

    /* Tipografía: Inter para texto visible */
    .stApp, h1, h2, h3, h4, p, .stMarkdown, .stCaption,
    [data-testid="stMetricLabel"], [data-testid="stMetricValue"],
    div.stButton > button, div.stDownloadButton > button,
    div[data-testid="stTabs"] [data-testid="stTab"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }

    /* Estructura */
    #MainMenu, footer { visibility: hidden; }
    .block-container { padding-top: 5rem; padding-bottom: 2rem; max-width: 1500px; }

    /* Header: compacto pero funcional, sin tapar contenido */
    header[data-testid="stHeader"],
    [data-testid="stHeader"] {
        position: fixed !important;
        top: 0 !important;
        left: 0 !important;
        right: 0 !important;
        height: auto !important;
        min-height: 48px !important;
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        z-index: 1000 !important;
        overflow: visible !important;
        display: flex !important;
        align-items: center !important;
    }
    header[data-testid="stHeader"] .stAppToolbar,
    [data-testid="stToolbar"] {
        pointer-events: auto !important;
        background: transparent !important;
        display: flex !important;
        align-items: center !important;
        gap: 6px !important;
        padding: 8px !important;
    }
    header[data-testid="stHeader"] button,
    [data-testid="stHeader"] button {
        background: var(--sms-glass) !important;
        color: var(--sms-text) !important;
        border-radius: 10px !important;
        border: 1px solid var(--sms-glass-border) !important;
        backdrop-filter: blur(6px);
        -webkit-backdrop-filter: blur(6px);
        min-height: 36px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
    }
    header[data-testid="stHeader"] button:hover,
    [data-testid="stHeader"] button:hover {
        background: var(--sms-glass-strong) !important;
        border-color: rgba(255,255,255,0.35) !important;
    }
    header[data-testid="stHeader"] [data-testid="stIconMaterial"],
    header[data-testid="stHeader"] svg,
    [data-testid="stHeader"] [data-testid="stIconMaterial"],
    [data-testid="stHeader"] svg {
        color: var(--sms-text) !important;
    }

    h1, h2, h3, h4 { color: var(--sms-text); font-weight: 700; letter-spacing: -0.02em; }
    .stMarkdown, .stCaption, p { color: var(--sms-text); }
    hr { border-color: var(--sms-border); opacity: 0.6; }

    /* ===== Sidebar (cristal) ===== */
    div[data-testid="stSidebar"] {
        background: rgba(255,255,255,0.04);
        backdrop-filter: blur(14px) saturate(140%);
        -webkit-backdrop-filter: blur(14px) saturate(140%);
        border-right: 1px solid var(--sms-glass-border);
    }
    [data-testid="stSidebarCollapseButton"] > div > button {
        color: var(--sms-text) !important;
        background: transparent !important;
        border: none !important;
    }

    /* ===== Métricas (cristal + libertad de color del valor) ===== */
    [data-testid="stMetric"] {
        background: var(--sms-glass);
        backdrop-filter: blur(16px) saturate(150%);
        -webkit-backdrop-filter: blur(16px) saturate(150%);
        border: 1px solid var(--sms-glass-border);
        border-radius: 18px;
        padding: 18px 20px;
        box-shadow:
            0 8px 32px rgba(0,0,0,0.35),
            inset 0 1px 0 rgba(255,255,255,0.08);
    }
    [data-testid="stMetric"]:hover {
        background: var(--sms-glass-strong);
        border-color: rgba(255,255,255,0.18);
    }
    [data-testid="stMetricLabel"] {
        color: var(--sms-muted);
        font-size: 0.8rem;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.06em;
    }
    [data-testid="stMetricValue"] {
        color: var(--sms-text);
        font-size: 1.75rem;
        font-weight: 800;
    }

    /* ===== Botones (cristal) ===== alto especificidad */
    div[data-testid="stTabs"] button,
    div.stButton > button,
    div.stDownloadButton > button,
    div.stFormSubmitButton > button {
        border-radius: 14px !important;
        font-weight: 600;
        transition: all 0.2s ease;
        backdrop-filter: blur(8px);
        -webkit-backdrop-filter: blur(8px);
    }
    div.stButton > button {
        border: 1px solid var(--sms-glass-border);
        background: var(--sms-glass);
        color: var(--sms-text);
    }
    div.stButton > button:hover,
    div.stButton > button:active,
    div.stButton > button:focus {
        border-color: rgba(255,255,255,0.35);
        background: var(--sms-glass-strong);
        color: var(--sms-white);
    }
    div.stButton > button:active {
        border-color: rgba(255,255,255,0.5);
    }
    button[kind="primary"],
    div.stButton > button[data-testid="stBaseButton-primary"],
    div[data-testid="stForm"] button[data-testid="stBaseButton-primary"],
    button[data-testid="stBaseButton-primary"],
    div[data-testid="stMainBlockContainer"] button[kind="primary"],
    div[data-testid="stHorizontalBlock"] button[kind="primary"] {
        background: linear-gradient(135deg, #FFFFFF, #D9D9E3) !important;
        border: 1px solid rgba(255,255,255,0.7) !important;
        color: #000000 !important;
        box-shadow: 0 4px 16px rgba(0,0,0,0.25) !important;
    }
    button[kind="primary"] p,
    button[kind="primary"] span,
    button[kind="primary"] *,
    button[data-testid="stBaseButton-primary"] p,
    button[data-testid="stBaseButton-primary"] span,
    button[data-testid="stBaseButton-primary"] * {
        color: #000000 !important;
    }
    button[kind="primary"]:hover,
    button[kind="primary"]:active,
    button[kind="primary"]:focus,
    div.stButton > button[data-testid="stBaseButton-primary"]:hover,
    div.stButton > button[data-testid="stBaseButton-primary"]:active,
    div.stButton > button[data-testid="stBaseButton-primary"]:focus,
    button[data-testid="stBaseButton-primary"]:hover,
    button[data-testid="stBaseButton-primary"]:active,
    button[data-testid="stBaseButton-primary"]:focus {
        background: linear-gradient(135deg, #FFFFFF, #C9C9D6) !important;
        color: #000000 !important;
        border-color: rgba(255,255,255,0.9) !important;
    }
    button[kind="primary"]:hover *,
    button[kind="primary"]:active *,
    button[kind="primary"]:focus *,
    button[data-testid="stBaseButton-primary"]:hover *,
    button[data-testid="stBaseButton-primary"]:active *,
    button[data-testid="stBaseButton-primary"]:focus * {
        color: #000000 !important;
    }
    div.stDownloadButton > button {
        border: 1px solid rgba(255,255,255,0.35) !important;
        background: var(--sms-glass) !important;
        color: var(--sms-text) !important;
    }
    div.stDownloadButton > button:hover,
    div.stDownloadButton > button:active,
    div.stDownloadButton > button:focus {
        background: var(--sms-glass-strong) !important;
        color: var(--sms-text) !important;
    }

    /* ===== Tabs: píldoras de cristal ===== */
    /* Alto especificidad para ganarle a emotion cache de Streamlit */
    div[data-testid="stTabs"] {
        background: transparent !important;
        padding: 0 !important;
        border: none !important;
    }
    div[data-testid="stTabs"] [role="tablist"] {
        gap: 6px !important;
        background: transparent !important;
        border: none !important;
        padding: 0 !important;
        box-shadow: none !important;
        display: flex !important;
        align-items: center !important;
        overflow: visible !important;
    }
    div[data-testid="stTabs"] [role="tablist"] [data-testid="stTab"] {
        border-radius: 999px !important;
        color: var(--sms-muted) !important;
        font-weight: 500;
        padding: 8px 22px !important;
        cursor: pointer;
        transition: all 0.2s ease;
        background: var(--sms-glass) !important;
        border: 1px solid var(--sms-glass-border) !important;
        box-shadow: none !important;
        position: relative !important;
        outline: none !important;
        backdrop-filter: blur(8px);
        -webkit-backdrop-filter: blur(8px);
    }
    div[data-testid="stTabs"] [role="tablist"] [data-testid="stTab"]::before,
    div[data-testid="stTabs"] [role="tablist"] [data-testid="stTab"]::after {
        display: none !important;
        content: none !important;
        box-shadow: none !important;
        border: none !important;
        background: transparent !important;
    }
    div[data-testid="stTabs"] [role="tablist"] [data-testid="stTab"]:hover {
        color: var(--sms-text) !important;
        background: var(--sms-glass-strong) !important;
        border-color: rgba(255,255,255,0.28) !important;
    }
    div[data-testid="stTabs"] [role="tablist"] [data-testid="stTab"][data-selected="true"],
    div[data-testid="stTabs"] [role="tablist"] [data-testid="stTab"][aria-selected="true"] {
        color: var(--sms-black) !important;
        font-weight: 700;
        background: linear-gradient(135deg, #FFFFFF, #E4E4EE) !important;
        border: 1px solid rgba(255,255,255,0.7) !important;
        box-shadow: 0 6px 20px rgba(0,0,0,0.35);
        border-radius: 999px !important;
    }
    div[data-testid="stTabs"] [role="tablist"] [data-testid="stTab"][data-selected="true"]::before,
    div[data-testid="stTabs"] [role="tablist"] [data-testid="stTab"][data-selected="true"]::after,
    div[data-testid="stTabs"] [role="tablist"] [data-testid="stTab"][aria-selected="true"]::before,
    div[data-testid="stTabs"] [role="tablist"] [data-testid="stTab"][aria-selected="true"]::after {
        display: none !important;
        content: none !important;
    }
    div[data-testid="stTabs"] [role="tablist"] .react-aria-SelectionIndicator {
        display: none !important;
    }
    div[data-testid="stTabs"] [role="tablist"] [data-testid="stTab"] [data-testid="stMarkdownContainer"] p {
        margin: 0 !important;
        line-height: 1.2;
        white-space: nowrap;
        color: inherit !important;
    }

    /* ===== Tabla de datos (cristal) ===== */
    [data-testid="stDataFrame"] {
        border: 1px solid var(--sms-glass-border);
        border-radius: 16px;
        overflow: hidden;
        background: rgba(255,255,255,0.02);
        backdrop-filter: blur(8px);
        -webkit-backdrop-filter: blur(8px);
    }

    /* ===== Inputs / selectores (cristal) ===== alto especificidad */
    div[data-testid="stVerticalBlock"] .stTextInput input,
    div[data-testid="stVerticalBlock"] .stTextArea textarea,
    div[data-testid="stVerticalBlock"] .stSelectbox div[data-baseweb="select"] > div,
    div[data-testid="stVerticalBlock"] .stNumberInput input {
        color: var(--sms-text) !important;
        background: var(--sms-glass) !important;
        border-color: var(--sms-glass-border) !important;
        border-radius: 12px !important;
        backdrop-filter: blur(8px);
        -webkit-backdrop-filter: blur(8px);
    }
    div[data-testid="stVerticalBlock"] .stTextInput input:focus,
    div[data-testid="stVerticalBlock"] .stTextArea textarea:focus,
    div[data-testid="stVerticalBlock"] .stNumberInput input:focus,
    div[data-testid="stVerticalBlock"] .stSelectbox div[data-baseweb="select"] > div:focus-within {
        border-color: rgba(255,255,255,0.4) !important;
        box-shadow: 0 0 0 3px rgba(255,255,255,0.08) !important;
    }

    /* Selectbox: lista desplegable */
    div[data-testid="stVerticalBlock"] [data-baseweb="popover"] [role="listbox"],
    div[data-testid="stVerticalBlock"] [data-baseweb="menu"] {
        background: #1A1A22 !important;
        border: 1px solid var(--sms-glass-border) !important;
        border-radius: 12px !important;
    }
    div[data-testid="stVerticalBlock"] [data-baseweb="popover"] [role="listbox"] li,
    div[data-testid="stVerticalBlock"] [data-baseweb="menu"] li {
        color: var(--sms-text) !important;
        background: transparent !important;
    }
    div[data-testid="stVerticalBlock"] [data-baseweb="popover"] [role="listbox"] li[aria-selected="true"],
    div[data-testid="stVerticalBlock"] [data-baseweb="menu"] li[aria-selected="true"] {
        color: var(--sms-black) !important;
        background: #FFFFFF !important;
    }
    div[data-testid="stVerticalBlock"] [data-baseweb="popover"] [role="listbox"] li:hover,
    div[data-testid="stVerticalBlock"] [data-baseweb="menu"] li:hover {
        background: rgba(255,255,255,0.10) !important;
        color: var(--sms-text) !important;
    }

    /* File uploader (Importar) */
    div[data-testid="stVerticalBlock"] [data-testid="stFileUploader"],
    div[data-testid="stVerticalBlock"] section[data-testid="stFileUploader"] {
        background: var(--sms-glass) !important;
        border: 1px dashed rgba(255,255,255,0.3) !important;
        border-radius: 14px !important;
    }
    div[data-testid="stVerticalBlock"] [data-testid="stFileUploader"] button {
        color: var(--sms-text) !important;
        background: var(--sms-glass-strong) !important;
        border: 1px solid var(--sms-glass-border) !important;
        border-radius: 10px !important;
    }
    div[data-testid="stVerticalBlock"] [data-testid="stFileUploader"] button:hover,
    div[data-testid="stVerticalBlock"] [data-testid="stFileUploader"] button:active,
    div[data-testid="stVerticalBlock"] [data-testid="stFileUploader"] button:focus {
        color: var(--sms-black) !important;
        background: #FFFFFF !important;
    }
    div[data-testid="stVerticalBlock"] [data-testid="stFileUploader"] span,
    div[data-testid="stVerticalBlock"] [data-testid="stFileUploader"] small {
        color: var(--sms-text) !important;
    }

    /* Paneles de contenedor opcional */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background: var(--sms-glass);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid var(--sms-glass-border);
        border-radius: 18px;
    }

    /* ===== Tabla resumen del dashboard (cristal) ===== */
    .sms-resumen {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        border: 1px solid var(--sms-glass-border);
        border-radius: 16px;
        overflow: hidden;
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        background: var(--sms-glass);
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }
    .sms-resumen th {
        background: rgba(255,255,255,0.09);
        color: var(--sms-muted);
        font-size: 0.78rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        padding: 14px 18px;
        text-align: left;
        border-bottom: 1px solid var(--sms-glass-border);
    }
    .sms-resumen td {
        color: var(--sms-text);
        font-size: 0.92rem;
        font-weight: 500;
        padding: 13px 18px;
        border-bottom: 1px solid rgba(255,255,255,0.04);
    }
    .sms-resumen tr:last-child td {
        border-bottom: none;
    }
    .sms-resumen tr:hover td {
        background: rgba(255,255,255,0.04);
    }
    .sms-resumen .sms-total td {
        background: rgba(255,255,255,0.07);
        font-weight: 700;
        font-size: 0.95rem;
        border-top: 1px solid var(--sms-glass-border);
    }
    .sms-resumen .sms-num {
        text-align: right;
        font-variant-numeric: tabular-nums;
    }
    .sms-resumen .sms-dot {
        display: inline-block;
        width: 10px;
        height: 10px;
        border-radius: 50%;
        margin-right: 8px;
        vertical-align: middle;
    }

    /* ===== RESPONSIVE =====
       Al reducir resolución el contenido se compacta (no "se aleja"). */
    /* Pantalla grande (>=1200px): contenedor amplio */
    @media (min-width: 1200px) {
        .block-container { max-width: 1500px !important; padding-top: 5rem; }
        [data-testid="stMetricValue"] { font-size: 1.8rem; }
        [data-testid="stTab"] { padding: 8px 24px; }
    }
    /* Escritorio medio (>= 992px y < 1200px) */
    @media (min-width: 992px) and (max-width: 1199.98px) {
        .block-container { max-width: 1200px !important; padding-top: 4.5rem; }
        [data-testid="stMetricValue"] { font-size: 1.6rem; }
        [data-testid="stTab"] { padding: 8px 20px !important; }
    }
    /* Tablet (>= 768px y < 992px) — completamente responsive y táctil */
    @media (min-width: 768px) and (max-width: 991.98px) {
        .block-container { max-width: 100% !important; width: 100% !important;
                           padding: 4rem 1.25rem 2rem !important; }
        h1 { font-size: 1.7rem !important; }
        h2 { font-size: 1.4rem !important; }
        h3 { font-size: 1.15rem !important; }

        /* Métricas: 2 columnas (más grandes de tocar) en vez de 5 chicas */
        div[data-testid="stMetric"] { padding: 16px 18px !important; border-radius: 16px !important; }
        [data-testid="stMetricValue"] { font-size: 1.5rem !important; }
        [data-testid="stMetricLabel"] { font-size: 0.9rem !important; }
        [data-testid="stMetric"] { min-width: 140px !important; }

        /* Tabs: scroll horizontal suave para que no se apelen */
        [data-testid="stTabs"] { overflow-x: auto !important; }
        div[data-testid="stTabs"] [role="tablist"] { flex-wrap: nowrap !important; }
        [data-testid="stTab"] { padding: 9px 18px !important; font-size: 1rem !important; }

        /* Columnas: que apilen columnas de gráficos/métricas */
        .block-container [data-testid="column"] {
            min-width: 0 !important;
            word-break: break-word;
        }
        /* Force el apilado de sub-columnas agrupadas en 2 a 1 por renglón */
        div[data-testid="stHorizontalBlock"] {
            flex-wrap: wrap !important;
        }

        /* Botones a ancho completo para fácil toque */
        div.stButton > button,
        div.stDownloadButton > button,
        div.stFormSubmitButton > button,
        button[data-testid="stBaseButton-primary"] {
            min-height: 44px !important;
            width: 100% !important;
        }

        /* Inputs/selectores más altos para toque */
        div[data-testid="stVerticalBlock"] .stTextInput input,
        div[data-testid="stVerticalBlock"] .stTextArea textarea,
        div[data-testid="stVerticalBlock"] .stSelectbox div[data-baseweb="select"] > div,
        div[data-testid="stVerticalBlock"] .stNumberInput input {
            min-height: 44px !important;
        }

        /* Tabla: scroll */
        [data-testid="stDataFrame"] { overflow-x: auto !important; }

        /* Sidebar: ancho cómodo */
        header[data-testid="stHeader"] { min-height: 46px !important; }
        header[data-testid="stHeader"] button { min-height: 34px !important; }
        div[data-testid="stSidebar"] { width: 280px !important; }

        /* Checkboxes del módulo: más grandes de tocar y sin desbordar el texto */
        [data-testid="stCheckbox"] label {
            min-height: 38px !important;
            display: flex !important;
            align-items: center !important;
            word-break: break-word !important;
        }
        [data-testid="stCheckbox"] {
            padding: 2px 0 !important;
        }

        /* Buscador del módulo: que use todo el ancho */
        .block-container div[data-testid="stHorizontalBlock"] > div {
            min-width: 0 !important;
        }

        /* Expandir 'Vista previa de selección' ocupa ancho completo */
        [data-testid="stExpander"] { width: 100% !important; }

        /* Markdown de código del módulo: fuente legible */
        [data-testid="stMarkdownContainer"] h5 { font-size: 1.05rem !important; }

        /* Selectbox en táctil: el input interno es readonly para que NO abra el
           teclado automáticamente. Se conserva el dropdown (clic abre la lista). */
        div[data-testid="stVerticalBlock"] [data-baseweb="select"] input,
        [data-testid="stSelectbox"] input,
        [data-baseweb="select"] input {
            -webkit-user-select: none !important;
            user-select: none !important;
            pointer-events: none !important;
            caret-color: transparent !important;
        }
        div[data-testid="stVerticalBlock"] [data-baseweb="select"] > div,
        [data-baseweb="select"] > div {
            cursor: pointer !important;
        }
    }
    /* Móvil (< 768px): apilar y compactar */
    @media (max-width: 767.98px) {
        /* Selectbox en táctil: input readonly para que no abra el teclado */
        div[data-testid="stVerticalBlock"] [data-baseweb="select"] input,
        [data-testid="stSelectbox"] input,
        [data-baseweb="select"] input {
            -webkit-user-select: none !important;
            user-select: none !important;
            pointer-events: none !important;
            caret-color: transparent !important;
        }
        div[data-testid="stVerticalBlock"] [data-baseweb="select"] > div,
        [data-baseweb="select"] > div {
            cursor: pointer !important;
        }
        .block-container { max-width: 100% !important; width: 100% !important;
                           padding: 3.5rem 1rem 1.5rem !important; }
        h1 { font-size: 1.5rem !important; }
        h2 { font-size: 1.25rem !important; }
        h3 { font-size: 1.05rem !important; }
        [data-testid="stMetric"] { padding: 12px 14px; border-radius: 14px; }
        [data-testid="stMetricValue"] { font-size: 1.2rem; }
        [data-testid="stTab"] { padding: 8px 14px !important; font-size: 0.9rem !important; }
        [data-testid="stTabs"] { overflow-x: auto !important; }
        div[data-testid="stTabs"] [role="tablist"] { flex-wrap: nowrap !important; }
        header[data-testid="stHeader"] { min-height: 44px !important; }
        header[data-testid="stHeader"] button { min-height: 32px !important; }
        div[data-testid="stSidebar"] { width: 260px !important; }
        [data-testid="stDataFrame"] { overflow-x: auto !important; }
        .stButton > button, .stDownloadButton > button {
            width: 100% !important;
        }
    }
    /* Evitar que el contenido "se aleje": que las columnas colapsen suavemente */
    .block-container [data-testid="column"] {
        min-width: 0 !important;
        word-break: break-word;
    }
    .st-emotion-cache-13ln4jf, [data-testid="stMainBlockContainer"] {
        max-width: 100% !important;
    }
    </style>
    """, unsafe_allow_html=True)


# --- Modo solo lectura (dashboard puro) ---
import urllib.parse
SOLO = st.query_params.get('solo', '0') == '1'


db.init_db()
inyectar_css()

TALLAS = datos.TALLAS
ESTADOS = ['TELA X ING 3-SET', 'EN LAVANDER' + chr(205) + 'A', 'EN CORTE', 'EN COSTURA', 'EN ACABADOS', 'ENCAJADO']

CHART_COLORS = {'TELA X ING 3-SET': '#6B6B80', 'EN LAVANDER' + chr(205) + 'A': '#4ADE80',
                'EN CORTE': '#7C5CFC', 'EN COSTURA': '#FBBF24', 'EN ACABADOS': '#F97316',
                'ENCAJADO': '#EF4444'}


def fmt_fecha(f):
    try:
        return f.strftime('%d-%m-%Y')
    except Exception:
        return ''


def titulo_dashboard(hoja=None):
    """Devuelve el título del dashboard con cliente y PO, si están disponibles."""
    base = 'Dashboard de avance'
    datos_db = db.obtener_todos(hoja=hoja)
    po = next((d.get('po') for d in datos_db if d.get('po')), None)
    if not po:
        if hoja:
            return f'{base} · {hoja}'
        return base
    # El cliente se deriva del prefijo del PO (ej. 'SMS' en 'SMS-01432')
    prefijo = str(po).split('-')[0].strip()
    tipo = 'Muestra de venta' if prefijo.upper() == 'SMS' else prefijo
    return f'{base} ({tipo} · PO {po})'

# ---------- Estado de sesión ----------
if 'importado' not in st.session_state:
    st.session_state.importado = db.existe_datos()
if 'swal' not in st.session_state:
    st.session_state.swal = ''
if 'hoja_actual' not in st.session_state:
    hojas = db.obtener_hojas_disponibles()
    st.session_state.hoja_actual = hojas[0] if hojas else None


def hoja_actual():
    """Devuelve la hoja seleccionada actualmente (o None si no hay datos)."""
    hojas = db.obtener_hojas_disponibles()
    if not hojas:
        st.session_state.hoja_actual = None
        return None
    if st.session_state.hoja_actual not in hojas:
        st.session_state.hoja_actual = hojas[0]
    return st.session_state.hoja_actual


def toast(msg):
    st.session_state.swal = msg


# ---------- Barra lateral ----------
with st.sidebar:
    st.title('Seguimiento SMS')
    st.caption('Panel de control de producción')

    st.divider()

    if not SOLO:
        if st.session_state.importado:
            st.success('Datos cargados')
            if st.button('Recargar datos'):
                st.session_state.importado = db.existe_datos()
                st.rerun()
        else:
            st.warning('Aun no hay datos. Importa un Excel.')

        st.divider()

    if st.session_state.swal:
        st.toast(st.session_state.swal, icon=None)
        st.session_state.swal = ''

if not SOLO:
    tab_import, tab_seg, tab_dash, tab_export, tab_costura, tab_acabados = st.tabs([
        'Importar', 'Seguimiento', 'Dashboard', 'Exportar', 'Costura', 'Acabados'])
else:
    tab_import = tab_seg = tab_dash = tab_export = tab_costura = tab_acabados = None


@st.fragment(run_every=10)
def render_dashboard(hoja=None):
    """Renderiza el contenido del dashboard y se auto-refresca cada 10 segundos."""
    from collections import Counter
    import plotly.graph_objects as go
    import pandas as pd

    datos_db = db.obtener_todos(hoja=hoja)
    c_est = Counter(d['status'] or '(sin estado)' for d in datos_db)
    tp = sum(d['total_pedido'] or 0 for d in datos_db)
    tg = sum(d['total_prog'] or 0 for d in datos_db)
    n_estilos = len({d['style'] for d in datos_db})
    template = 'plotly_dark'

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric('Estilos', n_estilos)
    k2.metric('Variantes', len(datos_db))
    k3.metric('Total PEDIDO', tp)
    k4.metric('Total PROGRAMADO', tg)
    k5.metric('Avance', 'En Progreso')

    st.divider()
    c1, c2 = st.columns(2)

    with c1:
        st.subheader('Variantes por estado')
        order = ['TELA X ING 3-SET', 'EN LAVANDER' + chr(205) + 'A', 'EN CORTE', 'EN COSTURA', 'EN ACABADOS', 'ENCAJADO']
        ov = order + [k for k in c_est if k not in order and k != '(sin estado)'] + (['(sin estado)'] if '(sin estado)' in c_est else [])
        ov = [k for k in ov if k in c_est]
        fig = go.Figure(go.Bar(x=ov, y=[c_est[k] for k in ov],
                               marker_color=[CHART_COLORS.get(k, '#7C5CFC') for k in ov],
                               text=[c_est[k] for k in ov], textposition='outside'))
        fig.update_layout(margin=dict(t=30, b=0, l=0, r=0), height=320, template=template)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.subheader('Pedido vs Programado (Top 10)')
        estilos_map = {}
        for d in datos_db:
            e = estilos_map.setdefault(d['style'], {'ped': 0, 'prog': 0, 'name': d['name']})
            e['ped'] += d['total_pedido'] or 0
            e['prog'] += d['total_prog'] or 0
        top = sorted(estilos_map.items(), key=lambda kv: -kv[1]['prog'])[:10]
        tlabels = [k for k, _ in top]
        tped = [v['ped'] for _, v in top]
        tprog = [v['prog'] for _, v in top]
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(name='Programado', x=tlabels, y=tprog, marker_color='#7C5CFC'))
        fig2.add_trace(go.Bar(name='Pedido', x=tlabels, y=tped, marker_color='#4ADE80'))
        fig2.update_layout(barmode='group', margin=dict(t=30, b=0, l=0, r=0), height=320, template=template)
        st.plotly_chart(fig2, use_container_width=True)

    st.divider()
    st.subheader('Resumen por estado')
    resumen_rows = []
    for est in order:
        if est not in c_est:
            continue
        ped = sum(d['total_pedido'] or 0 for d in datos_db if (d['status'] or '(sin estado)') == est)
        prog = sum(d['total_prog'] or 0 for d in datos_db if (d['status'] or '(sin estado)') == est)
        resumen_rows.append({'Estado': est, 'Pedido': ped, 'Programado': prog})
    if resumen_rows:
        tot_ped = sum(r['Pedido'] for r in resumen_rows)
        tot_prog = sum(r['Programado'] for r in resumen_rows)
        resumen_rows.append({'Estado': 'TOTAL', 'Pedido': tot_ped, 'Programado': tot_prog, '_total': True})
        html_rows = ''
        for r in resumen_rows:
            es_total = r.get('_total', False)
            cls = ' class="sms-total"' if es_total else ''
            dot = '' if es_total else f'<span class="sms-dot" style="background:{CHART_COLORS.get(r["Estado"], "#7C5CFC")}"></span>'
            html_rows += f'<tr{cls}><td>{dot}{r["Estado"]}</td><td class="sms-num">{r["Pedido"]:,}</td><td class="sms-num">{r["Programado"]:,}</td></tr>'
        st.markdown(f'''<table class="sms-resumen">
        <thead><tr><th>Estado</th><th class="sms-num">Pedido</th><th class="sms-num">Programado</th></tr></thead>
        <tbody>{html_rows}</tbody></table>''', unsafe_allow_html=True)

    st.divider()
    st.subheader('Detalle de variantes')
    ver = st.toggle('Mostrar columnas de tallas', value=False)
    if ver:
        rows = []
        for d in datos_db:
            r = {
                'Style': d['style'], 'Color': d['color'], 'Prenda': d['name'], 'Tela': d['tela'],
                'Status': d['status'], 'Pedido': d['total_pedido'], 'Programado': d['total_prog'],
                'Obs': d['observaciones'] or ''
            }
            for i in range(len(TALLAS)):
                r[f'P {TALLAS[i]}'] = d['pedido'][i] if i < len(d['pedido']) else 0
                r[f'G {TALLAS[i]}'] = d['prog'][i] if i < len(d['prog']) else 0
            rows.append(r)
        df = pd.DataFrame(rows)
    else:
        df = pd.DataFrame([{
            'Style': d['style'], 'Color': d['color'], 'Prenda': d['name'], 'Tela': d['tela'],
            'Status': d['status'], 'Pedido': d['total_pedido'], 'Programado': d['total_prog'],
            'Obs': d['observaciones'] or ''
        } for d in datos_db])
    st.dataframe(df, use_container_width=True, hide_index=True)


def render_modulo(modo, hoja, hojas_disponibles):
    """Módulo de trabajo por estado: 'costura' -> EN COSTURA, 'acabados' -> EN ACABADOS.

    Lista todos los códigos (styles) con sus colores, con buscador.
    El usuario marca con check las prendas y guarda; se actualiza su estado.
    """
    estado_destino = 'EN COSTURA' if modo == 'costura' else 'EN ACABADOS'
    nombre = 'Costura' if modo == 'costura' else 'Acabados'

    import pandas as pd

    st.header(f'Módulo {nombre}')
    st.caption(f'Al guardar, las prendas seleccionadas pasarán a estado **{estado_destino}**.')
    st.write('Elige el proyecto sobre el que quieres trabajar:')
    # Selector propio de proyecto (independiente para cada módulo)
    idx = hojas_disponibles.index(hoja) if hoja in hojas_disponibles else 0
    hoja = st.selectbox('Proyecto', hojas_disponibles, index=idx, key=f'mod_hoja_{modo}')

    datos_db = db.obtener_todos(hoja=hoja)
    if not datos_db:
        st.info(f'El proyecto **{hoja}** no tiene datos. Impórtalo en la pestaña **Importar**.')
        return

    # Buscador
    col_b1, col_b2 = st.columns([3, 1])
    q = col_b1.text_input('Buscar por código o color', '', key=f'buscar_{modo}').strip().lower()
    # Filtro por estado actual (opcional)
    estados_list = sorted({d['status'] or '(sin estado)' for d in datos_db})
    filtro_est = col_b2.selectbox('Estado actual', ['Todos'] + estados_list, key=f'estado_filtro_{modo}')

    # Filtrar
    visibles = []
    for d in datos_db:
        if filtro_est != 'Todos' and (d['status'] or '(sin estado)') != filtro_est:
            continue
        if q:
            if q not in d['style'].lower() and q not in (d['color'] or '').lower():
                continue
        visibles.append(d)

    st.caption(f'Mostrando {len(visibles)} de {len(datos_db)} variantes')

    if not visibles:
        st.info('No hay resultados con el filtro actual.')
        return

    # Selección con checkboxes
    claves = {f'{d["style"]} — {d["color"]}': d for d in visibles}
    st.markdown('##### Marca las prendas para enviar a ' + estado_destino)
    sel = {}
    for k in claves:
        sel[k] = st.checkbox(k, key=f'mod_{modo}_{k}')

    st.divider()
    elegidas = [k for k, v in sel.items() if v]
    st.caption(f'{len(elegidas)} seleccionada(s)')

    col_guardar, col_info = st.columns([1, 3])
    if col_guardar.button('Guardar selección', type='primary', use_container_width=True, key=f'guardar_{modo}'):
        if not elegidas:
            col_info.warning('No has marcado ninguna prenda.')
        else:
            for k in elegidas:
                d = claves[k]
                db.actualizar_seguimiento(hoja, d['style'], d['color'], status=estado_destino)
            st.session_state.swal = (f'Actualizadas {len(elegidas)} prendas a {estado_destino}.')
            st.rerun()

    # Vista previa de lo que se va a cambiar (solo muestro marcas previas)
    if elegidas:
        with st.expander('Vista previa de selección'):
            prev = pd.DataFrame([{
                'Style': claves[k]['style'], 'Color': claves[k]['color'],
                'Prenda': claves[k]['name'], 'Estado actual': claves[k]['status'],
                'Nuevo estado': estado_destino
            } for k in elegidas])
            st.dataframe(prev, use_container_width=True, hide_index=True)


# --- Modo solo lectura: mostrar solo el dashboard y detener ---
if SOLO:
    hojas = db.obtener_hojas_disponibles()
    if not hojas:
        st.info('Sin datos para mostrar.')
        st.stop()
    # Selector de proyecto (con soporte de ?solo=1&hoja=...)
    q_hoja = st.query_params.get('hoja', None)
    idx = hojas.index(q_hoja) if q_hoja in hojas else 0
    sel = st.selectbox('Proyecto', hojas, index=idx, key='solo_hoja')
    st.header(titulo_dashboard(sel))
    st.caption('Se actualiza automaticamente cada 10 segundos')
    render_dashboard(sel)
    st.stop()

# ============================================================
# TAB IMPORTAR
# ============================================================
with tab_import:
    st.header('Importar datos desde Excel')
    st.write('Sube un archivo Excel con una hoja de seguimiento. La app extrae los datos y los guarda, '
             'y a partir de ahí podrás editar y dar seguimiento sin depender del Excel original.')

    uploaded = st.file_uploader('Archivo Excel (.xlsx)', type=['xlsx'])

    if uploaded is not None:
        temp_path = os.path.join(os.environ.get('TEMP', '/tmp'), 'sms_upload' + str(int(time.time())) + '.xlsx')
        with open(temp_path, 'wb') as fh:
            fh.write(uploaded.getbuffer())

        try:
            hojas = datos.obtener_hojas(temp_path)
            st.info(f'El archivo contiene {len(hojas)} hojas.')
            st.markdown('##### Elige qué hojas (proyectos) quieres importar')
            elegidas = []
            for h in hojas:
                if st.checkbox(h, key=f'hoja_import_{h}'):
                    elegidas.append(h)
            if st.button('Importar datos seleccionados', type='primary'):
                if not elegidas:
                    st.warning('No seleccionaste ninguna hoja.')
                else:
                    total_filas = 0
                    total_estilos = 0
                    for h in elegidas:
                        filas = datos.extractar_hoja(temp_path, h)
                        if not filas:
                            st.error(f'No se encontraron filas válidas en la hoja "{h}". No se importó.')
                            continue
                        db.reemplazar_datos(h, filas)
                        total_filas += len(filas)
                        total_estilos += len({f['style'] for f in filas})
                    if total_filas:
                        st.session_state.importado = True
                        # Actualizar la hoja actual al primero importado, si no hay una activa
                        if st.session_state.hoja_actual not in db.obtener_hojas_disponibles():
                            st.session_state.hoja_actual = elegidas[0]
                        st.session_state.swal = (f'Importadas {total_filas} variantes de '
                                                 f'{total_estilos} estilos en {len(elegidas)} hojas.')
                        st.rerun()
        except Exception as e:
            st.error(f'No se pudo leer el archivo: {e}')
        finally:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except PermissionError:
                    pass

    if st.session_state.importado:
        st.divider()
        st.warning('Reimportar una hoja reemplaza solo los datos de esa hoja/proyecto.',
                   icon=None)

# ============================================================
# TAB SEGUIMIENTO (EDICIÓN)
# ============================================================
with tab_seg:
    if not st.session_state.importado:
        st.info('Primero importa un Excel en la pestaña **Importar**.')
    else:
        cur_hoja = hoja_actual()
        hojas_disp = db.obtener_hojas_disponibles()
        if cur_hoja is None:
            st.info('No hay proyectos cargados. Importa un Excel en la pestaña **Importar**.')
        else:
            col_proy = st.columns(1)[0]
            cur_hoja = col_proy.selectbox('Proyecto', hojas_disp,
                                          index=hojas_disp.index(cur_hoja)
                                          if cur_hoja in hojas_disp else 0,
                                          key='seg_hoja')
            st.session_state.hoja_actual = cur_hoja

            datos_db = db.obtener_todos(hoja=cur_hoja)
            st.header('Seguimiento y edición')

        # Filtros
        col_f1, col_f2, col_f3 = st.columns(3)
        estilos_list = sorted({d['style'] for d in datos_db})
        estados_list = sorted({d['status'] for d in datos_db})
        filtro_style = col_f1.selectbox('Filtrar por estilo', ['Todos'] + estilos_list)
        filtro_estado = col_f2.selectbox('Filtrar por estado', ['Todos'] + estados_list)
        buscar = col_f3.text_input('Buscar (estilo o color)', '')

        def aplicar_filtros(ds):
            out = []
            for d in ds:
                if filtro_style != 'Todos' and d['style'] != filtro_style:
                    continue
                if filtro_estado != 'Todos' and d['status'] != filtro_estado:
                    continue
                if buscar:
                    q = buscar.lower()
                    if q not in d['style'].lower() and q not in d['color'].lower():
                        continue
                out.append(d)
            return out

        filtrados = aplicar_filtros(datos_db)
        st.caption(f'Mostrando {len(filtrados)} de {len(datos_db)} variantes')

        # ===== Acción en lote (selección múltiple) =====
        if filtrados:
            with st.expander('Cambiar estado en masa (selección múltiple)', expanded=False):
                opciones = [f'{d["style"]} — {d["color"]}'
                            + (f'  [{d["status"]}]' if d['status'] not in (None, '') else '')
                            for d in filtrados]
                # Atajos para seleccionar todo o limpiar la selección
                c_todo, c_limpiar = st.columns(2)
                if c_todo.button('Seleccionar todos', key='masiva_todo', use_container_width=True):
                    st.session_state['masiva_opciones'] = opciones
                    st.rerun()
                if c_limpiar.button('Limpiar', key='masiva_limpiar', use_container_width=True):
                    st.session_state['masiva_opciones'] = []
                    st.rerun()
                seleccion_masiva = st.multiselect('Selecciona las variantes a cambiar',
                                                  opciones, key='masiva_opciones')
                estado_masivo = st.selectbox('Nuevo estado para todas',
                                             ESTADOS + ['(sin estado)'],
                                             key='masiva_estado')
                if st.button(f'Aplicar cambio de estado ({len(seleccion_masiva)} variante(s))',
                             type='primary', key='masiva_aplicar'):
                    if not seleccion_masiva:
                        st.warning('No seleccionaste ninguna variante.')
                    else:
                        clave_normal = {o: d for o, d in zip(opciones, filtrados)}
                        for o in seleccion_masiva:
                            d = clave_normal[o]
                            db.actualizar_seguimiento(cur_hoja, d['style'], d['color'], status=estado_masivo)
                        st.session_state.swal = (f'Actualizadas {len(seleccion_masiva)} variantes '
                                                 f'a **{estado_masivo}**.')
                        st.rerun()

        if not filtrados:
            st.info('No hay resultados con los filtros actuales.')
        else:
            # Selector de variante
            claves = [f'{d["style"]} — {d["color"]}' for d in filtrados]
            seleccion = st.selectbox('Selecciona una variante para editar', claves)
            idx = claves.index(seleccion)
            act = filtrados[idx]

            st.divider()
            m1, m2, m3, m4 = st.metric('PO', act['po'] or '—'), st.metric('Prenda', act['name'] or '—'), \
                              st.metric('Tela', act['tela'] or '—'), st.metric('Ingreso Cost', act['ingreso'] or '—')
            col_izq, col_der = st.columns(2)

            with col_izq:
                st.subheader('Tallas — PEDIDO')
                pd_tab = {TALLAS[i]: act['pedido'][i] if i < len(act['pedido']) else 0 for i in range(len(TALLAS))}
                st.dataframe([pd_tab], hide_index=True, use_container_width=True)

            with col_der:
                st.subheader('Tallas — PROGRAMADO')
                # permitir editar las tallas programadas? por ahora solo muestra + total editable
                prog_lista = list(act['prog']) if len(act['prog']) == len(TALLAS) else [0]*len(TALLAS)
                def render_prog():
                    prog_tab = {TALLAS[i]: prog_lista[i] for i in range(len(TALLAS))}
                    st.dataframe([prog_tab], hide_index=True, use_container_width=True)
                render_prog()

            st.divider()
            st.subheader('Actualizar estado y observaciones')

            with st.form(f'form_{act["id"]}'):
                nuevo_estado = st.selectbox('Estado de tela', ESTADOS + ['(sin estado)'],
                                            index=(ESTADOS + ['(sin estado)']).index(act['status']) if act['status'] in ESTADOS + ['(sin estado)'] else 0)
                nuevas_obs = st.text_area('Observaciones', value=act['observaciones'] or '', height=100)
                guardar = st.form_submit_button('Guardar cambios', type='primary')

            if guardar:
                db.actualizar_seguimiento(cur_hoja, act['style'], act['color'], status=nuevo_estado,
                                          observaciones=nuevas_obs)
                toast(f'Guardado: {act["style"]} — {act["color"]} → {nuevo_estado}')
                st.rerun()

# ============================================================
# TAB DASHBOARD
# ============================================================
with tab_dash:
    if not st.session_state.importado:
        st.info('Primero importa un Excel en la pestaña **Importar**.')
    else:
        cur_hoja = hoja_actual()
        hojas_disp = db.obtener_hojas_disponibles()
        if cur_hoja is None:
            st.info('No hay proyectos cargados.')
        else:
            cur_hoja = st.selectbox('Proyecto', hojas_disp,
                                    index=hojas_disp.index(cur_hoja)
                                    if cur_hoja in hojas_disp else 0,
                                    key='dash_hoja')
            st.session_state.hoja_actual = cur_hoja
            st.header(titulo_dashboard(cur_hoja))
            st.caption('Se actualiza automáticamente cada 10 segundos')
            render_dashboard(cur_hoja)


# ============================================================
# TAB EXPORTAR
# ============================================================
with tab_export:
    if not st.session_state.importado:
        st.info('Primero importa un Excel en la pestaña **Importar**.')
    else:
        cur_hoja = hoja_actual()
        hojas_disp = db.obtener_hojas_disponibles()
        if cur_hoja is None:
            st.info('No hay proyectos cargados.')
        else:
            cur_hoja = st.selectbox('Proyecto', hojas_disp,
                                    index=hojas_disp.index(cur_hoja)
                                    if cur_hoja in hojas_disp else 0,
                                    key='export_hoja')
            st.session_state.hoja_actual = cur_hoja
            st.header('Exportar reporte')
            st.write('Genera un reporte del avance actual (con estados y observaciones ya editados).')

            cxl, cpdf = st.columns(2)
            if cxl.button('Exportar Excel', use_container_width=True):
                datos_db = db.obtener_todos(hoja=cur_hoja)
                b = reporte.exportar_excel(datos_db)
                st.download_button('Descargar Excel', data=b, file_name=f'seguimiento_{cur_hoja}_{datetime.date.today()}.xlsx',
                                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                   use_container_width=True)
            if cpdf.button('Exportar PDF', use_container_width=True):
                datos_db = db.obtener_todos(hoja=cur_hoja)
                b = reporte.exportar_pdf(datos_db)
                st.download_button('Descargar PDF', data=b, file_name=f'seguimiento_{cur_hoja}_{datetime.date.today()}.pdf',
                                   mime='application/pdf', use_container_width=True)

# ============================================================
# TAB MÓDULO COSTURA
# ============================================================
with tab_costura:
    if not st.session_state.importado:
        st.info('Primero importa un Excel en la pestaña **Importar**.')
    else:
        cur_hoja = hoja_actual()
        hojas_disp = db.obtener_hojas_disponibles()
        if cur_hoja is None or not hojas_disp:
            st.info('No hay proyectos cargados. Importa un Excel en la pestaña **Importar**.')
        else:
            render_modulo('costura', cur_hoja, hojas_disp)

# ============================================================
# TAB MÓDULO ACABADOS
# ============================================================
with tab_acabados:
    if not st.session_state.importado:
        st.info('Primero importa un Excel en la pestaña **Importar**.')
    else:
        cur_hoja = hoja_actual()
        hojas_disp = db.obtener_hojas_disponibles()
        if cur_hoja is None or not hojas_disp:
            st.info('No hay proyectos cargados. Importa un Excel en la pestaña **Importar**.')
        else:
            render_modulo('acabados', cur_hoja, hojas_disp)
