import streamlit as st
import pandas as pd
import os
from io import BytesIO
import matplotlib.pyplot as plt
import plotly.express as px
import re
import base64

from utils import *

# ==============================
# IMPORTS DAS PÁGINAS
# ==============================

from app_monitor import render_monitor
from app_devolucao import render_devolucao
from app_desempenho import render_desempenho
from app_tradeoff import render_tradeoff

# ==============================
# CONFIG STREAMLIT
# ==============================

st.set_page_config(
    page_title="BI Executivo - Monitor",
    layout="wide",
    page_icon="📊"
)

# ==============================
# ESTILO
# ==============================

st.markdown("""
<style>

    .titulo-painel {
        background: linear-gradient(90deg,#0f2a44,#1f4e79);
        color: white;
        font-size: 20px;
        font-weight: 700;
        text-align: center;
        padding: 12px;
        border-radius: 12px;
        margin-bottom: 20px;
    }

    .separador-colunas {
        border-right: 6px solid #1f4e79;
        padding-right: 25px;
    }

</style>

<style>

.stApp {
    background-color: #f4f6f9;
}

.header-box {
    background: linear-gradient(90deg, #0f2a44, #1f4e79);
    padding: 18px 24px;
    border-radius: 14px;
    display: flex;
    align-items: center;
    gap: 20px;
    margin-bottom: 20px;
}

.header-title {
    color: white;
    font-size: 26px;
    font-weight: 700;
    margin: 0;
}

.header-sub {
    color: white;
    opacity: 0.85;
    margin: 0;
    font-size: 14px;
}

.header-box img {
    max-width: 220px !important;
}

.data-title {
    font-size: 20px;
    font-weight: 700;
    color: #0f2a44;
    margin-top: 10px;
    margin-bottom: 10px;
}

.metric-small {
    font-size: 16px !important;
    font-weight: 600;
    color: #0f2a44;
}

button {
    background: linear-gradient(90deg, #0f2a44, #1f4e79) !important;
    color: white !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    border: none !important;
}

button:hover {
    background: linear-gradient(90deg, #1f4e79, #0f2a44) !important;
    color: white !important;
}

button[data-testid="collapsedControl"] {
    background-color: #0f2a44 !important;
    border-radius: 8px !important;
}

button[data-testid="collapsedControl"] svg {
    fill: white !important;
    color: white !important;
}

button[data-testid="collapsedControl"]:hover {
    background-color: #1f4e79 !important;
}

</style>
""", unsafe_allow_html=True)

# ==============================
# MENU LATERAL
# ==============================

pagina = st.sidebar.radio(
    "Painel",
    [
        "Monitor de Pedidos",
        "Indicador de Devolução",
        "Desempenho por Transportadora",
        "Trade-Off Logístico"
    ]
)

# ==============================
# HEADER
# ==============================

logo_html = ""
logo_base64 = ""

if os.path.exists(LOGO_PATH):

    with open(LOGO_PATH, "rb") as f:
        logo_base64 = base64.b64encode(
            f.read()
        ).decode()

    logo_html = (
        f'<img src="data:image/png;base64,{logo_base64}" width="120">'
    )

st.markdown(
    f"""
    <style>

    .stApp::before {{

        content: "";

        position: fixed;

        top: 0;
        left: 0;

        width: 100%;
        height: 100%;

        background-image:
        url("data:image/png;base64,{logo_base64}");

        background-repeat: no-repeat;

        background-position: center;

        background-size: 700px;

        opacity: 0.03;

        pointer-events: none;

        z-index: 0;
    }}

    </style>
    """,
    unsafe_allow_html=True
)

if pagina == "Monitor de Pedidos":

    titulo = "Monitor de Pedidos — BI Executivo"

    subtitulo = (
        "Análise de Risco Logístico • "
        "Transportadora • Status • Região • Cliente"
    )

elif pagina == "Indicador de Devolução":

    titulo = "Painel de Devolução"

    subtitulo = (
        "Devolução • Extravio • "
        "Avaria • Indicadores Logísticos"
    )

elif pagina == "Desempenho por Transportadora":

    titulo = "Desempenho por Transportadora - SLA"

    subtitulo = (
        "Eficiência de Entrega • "
        "Dentro vs Fora do Prazo"
    )

elif pagina == "Trade-Off Logístico":

    titulo = "Trade-Off Logístico"

    subtitulo = (
        "Simulação Estratégica • SLA • "
        "NFD • Frete • Similaridade CEP"
    )

st.markdown(
    f"""
    <div class="header-box">

        {logo_html}

        <div>

            <p class="header-title">
                {titulo}
            </p>

            <p class="header-sub">
                {subtitulo}
            </p>

        </div>

    </div>
    """,
    unsafe_allow_html=True
)

# ==============================
# ROTEAMENTO
# ==============================

if pagina == "Monitor de Pedidos":
    render_monitor()

elif pagina == "Indicador de Devolução":
    render_devolucao()

elif pagina == "Desempenho por Transportadora":
    render_desempenho()

elif pagina == "Trade-Off Logístico":
    render_tradeoff()