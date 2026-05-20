import streamlit as st
import pandas as pd
import os
from io import BytesIO
import matplotlib.pyplot as plt
import re

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
            df = pd.read_csv(
                path,
                sep=";",
                encoding="utf-8-sig"
            )
        else:
            df = pd.read_excel(path)

    except Exception:
        return pd.DataFrame()

    df.columns = df.columns.str.strip()

    if "PedidoFormatado" in df.columns:
        df["PedidoFormatado"] = normalizar_pedido(
            df["PedidoFormatado"]
        )

    col_igor = [
        c for c in df.columns
        if c.strip().upper() == "NIVEL_IGOR"
    ]

    if col_igor:
        df["NIVEL_IGOR"] = (
            df[col_igor[0]]
            .astype(str)
            .str.strip()
            .str.upper()
        )
    else:
        df["NIVEL_IGOR"] = ""

    return df

def listar_dias():

    if not os.path.exists(PASTA_HIST):
        return []

    arquivos = os.listdir(PASTA_HIST)

    datas = set()

    for a in arquivos:

        m = re.match(
            r"(\d{2}-\d{2}-\d{4})_manha\.(xlsx|csv)$",
            a
        )

        if m:
            datas.add(m.group(1))

    return sorted(
        datas,
        key=lambda x: pd.to_datetime(
            x,
            format="%d-%m-%Y"
        )
    )

def caminho(dia):

    caminho_csv = os.path.join(
        PASTA_HIST,
        f"{dia}_manha.csv"
    )

    caminho_xlsx = os.path.join(
        PASTA_HIST,
        f"{dia}_manha.xlsx"
    )

    if os.path.exists(caminho_csv):
        return caminho_csv

    return caminho_xlsx

def pizza(tratados, restantes, titulo):

    fig, ax = plt.subplots(figsize=(2.3, 2.3))

    total = tratados + restantes

    if total == 0:
        ax.text(
            0.5,
            0.5,
            "0",
            ha="center",
            va="center"
        )

    else:
        ax.pie(
            [tratados, restantes],
            autopct="%1.0f%%",
            startangle=90
        )

    ax.set_title(titulo, fontsize=10)

    buf = BytesIO()

    fig.savefig(
        buf,
        format="png",
        bbox_inches="tight"
    )

    plt.close(fig)

    buf.seek(0)

    return buf.getvalue()