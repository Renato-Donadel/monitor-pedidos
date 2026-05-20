from app_monitor import render_monitor
from app_devolucao import render_devolucao
from app_desempenho import render_desempenho
from app_tradeoff import render_tradeoff
import streamlit as st
import pandas as pd
import os
from io import BytesIO
import matplotlib.pyplot as plt
import plotly.express as px
import re
import base64

# ==============================
# CONFIGS
# ==============================
BASE_DIR = os.path.dirname(__file__)
PASTA_DATA = os.path.join(BASE_DIR, "data")
PASTA_HIST = os.path.join(PASTA_DATA, "historico")
PASTA_MENSAL = os.path.join(PASTA_DATA, "mensal_status")
os.makedirs(PASTA_MENSAL, exist_ok=True)
ARQ_ATUAL = os.path.join(PASTA_DATA, "Monitor_Pedidos_Processado.xlsx")
ARQ_DEVOLUCAO = os.path.join(PASTA_DATA, "Base_Streamlit_Devolucao.xlsx")
ARQ_TRADEOFF = os.path.join(PASTA_DATA, "Base_Similaridade_Tarifarios.xlsx")
LOGO_PATH = os.path.join(PASTA_DATA, "logo_bravium.png")

TAMANHO_LOTE = 300

STATUS_DIARIOS = [
    "TSP - Pendente Transportes - Dados do Recebedor Solicitado",
    "TSP - Item Faltante",
    "TSP - Pendente Transportes - Aguardando Acareação",
    "TSP - Pendente Transportes - Acareação Solicitada",
    "TSP - Item Faltante Solicitado",
    "TSP - Aguardando Dados do Recebedor"
]

STATUS_Manuais = [
    "TSP - Críticos",
    "TSP - Reentrega",
    "TSP - Reentregar/Endereço correto",
    "Aguardando tratativa transportadora",
    "TSP - Aguardando Acareação",
    "TSP - Pendente Transportes - Acareação Solicitada",
    "TSP - Aguardando Dados do Recebedor",
    "TSP - Pendente Transportes - Dados do Recebedor Solicitado",
    "TSP - Item Faltante",
    "TSP - Item Faltante Solicitado",
    "TSP - Aguardando Avaliação de Problema de Coleta",
    "TSP - Coleta Realizada",
    "TSP - Aguardando Coleta",
    "TSP - Coleta Agendada",
    "TSP - Aguardando Envio de Guia de Retenção ao Fiscal"
]

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
.stApp { background-color: #f4f6f9; }

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

/* BOTÕES */
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

/* ============================= */
/* BOTÃO SIDEBAR (COR + ÍCONE) */
/* ============================= */

/* botão */
button[data-testid="collapsedControl"] {
    background-color: #0f2a44 !important;
    border-radius: 8px !important;
}

/* ícone (seta) */
button[data-testid="collapsedControl"] svg {
    fill: white !important;
    color: white !important;
}

/* hover */
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
        logo_base64 = base64.b64encode(f.read()).decode()
    logo_html = f'<img src="data:image/png;base64,{logo_base64}" width="120">'
    
st.markdown(
    f"""
    <style>
    .stApp {{
        background-color: #f4f6f9;
    }}

    .stApp::before {{
        content: "";
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background-image: url("data:image/png;base64,{logo_base64}");
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
    subtitulo = "Análise de Risco Logístico • Transportadora • Status • Região • Cliente"

elif pagina == "Indicador de Devolução":
    titulo = "Painel de Devolução"
    subtitulo = "Devolução • Extravio • Avaria • Indicadores Logísticos"

elif pagina == "Desempenho por Transportadora":
    titulo = "Desempenho por Transportadora - SLA"
    subtitulo = "Eficiência de Entrega • Dentro vs Fora do Prazo"
    
elif pagina == "Trade-Off Logístico":
    titulo = "Trade-Off Logístico"
    subtitulo = "Simulação Estratégica • SLA • NFD • Frete • Similaridade CEP"

st.markdown(f"""
<div class="header-box">
    {logo_html}
    <div>
        <p class="header-title">{titulo}</p>
        <p class="header-sub">
        {subtitulo}
        </p>
    </div>
</div>
""", unsafe_allow_html=True)

# ==============================
# FUNÇÕES
# ==============================

def normalizar_pedido(col):
    return (
        col.astype(str)
        .str.upper()
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )
    
def perc_transportes(x, venda_mes):
    if venda_mes == 0:
        return "0%"
    return f"{(x / venda_mes)*100:.2f}%".replace(".", ",")

def carregar_base_devolucao():
    return pd.read_excel(
        ARQ_DEVOLUCAO,
        sheet_name=[
            "vendas_mes",
            "vendas_mes_pedido",
            "vendas_transportadora",
            "potencial_triplo",
            "devolucao_processo",
            "retornando_transportes",
            "devolucao_atrasada",
            "nfd_mes",
            "nfd_coleta",
            "base",
            "retornando_detalhado",
            "dev_atrasada_detalhado"
        ]
    )
@st.cache_data(ttl=2)
def ler_base(path):
    if not os.path.exists(path):
        return pd.DataFrame()

    try:
        if path.endswith(".csv"):
            df = pd.read_csv(path, sep=";", encoding="utf-8-sig")
        else:
            df = pd.read_excel(path)
    except Exception:
        return pd.DataFrame()

    # 🔥 PADRONIZA NOMES DAS COLUNAS
    df.columns = df.columns.str.strip()

    # 🔥 GARANTE PEDIDO
    if "PedidoFormatado" in df.columns:
        df["PedidoFormatado"] = normalizar_pedido(df["PedidoFormatado"])

    # 🔥 GARANTE NIVEL IGOR (mesmo que venha com nome zoado)
    col_igor = [c for c in df.columns if c.strip().upper() == "NIVEL_IGOR"]

    if col_igor:
        df["NIVEL_IGOR"] = df[col_igor[0]].astype(str).str.strip().str.upper()
    else:
        df["NIVEL_IGOR"] = ""

    return df

def listar_dias():
    if not os.path.exists(PASTA_HIST):
        return []
    arquivos = os.listdir(PASTA_HIST)
    datas = set()
    for a in arquivos:
        m = re.match(r"(\d{2}-\d{2}-\d{4})_manha\.(xlsx|csv)$", a)
        if m:
            datas.add(m.group(1))
    return sorted(datas, key=lambda x: pd.to_datetime(x, format="%d-%m-%Y"))

def caminho(dia):
    caminho_csv = os.path.join(PASTA_HIST, f"{dia}_manha.csv")
    caminho_xlsx = os.path.join(PASTA_HIST, f"{dia}_manha.xlsx")

    if os.path.exists(caminho_csv):
        return caminho_csv

    return caminho_xlsx

def pizza(tratados, restantes, titulo):
    fig, ax = plt.subplots(figsize=(2.3, 2.3))
    total = tratados + restantes
    if total == 0:
        ax.text(0.5, 0.5, "0", ha="center", va="center")
    else:
        ax.pie([tratados, restantes], autopct="%1.0f%%", startangle=90)
    ax.set_title(titulo, fontsize=10)
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()
        

                    