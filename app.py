import streamlit as st
import pandas as pd
import os
from io import BytesIO
import matplotlib.pyplot as plt
import re
import base64

# ==============================
# CONFIGS
# ==============================
BASE_DIR = os.path.dirname(__file__)
PASTA_DATA = os.path.join(BASE_DIR, "data")
PASTA_HIST = os.path.join(PASTA_DATA, "historico")
ARQ_ATUAL = os.path.join(PASTA_DATA, "Monitor_Pedidos_Processado.xlsx")
LOGO_PATH = os.path.join(PASTA_DATA, "logo_bravium.png")

TAMANHO_LOTE = 300

STATUS_DIARIOS = [
    "TSP - Pendente Transportes - Dados do Recebedor Solicitado",
    "TSP - Item Faltante",
    "TSP - Pendente Transportes - Aguardando Acareação",
    "TSP - Aguardando Dados do Recebedor"
]

st.set_page_config(
    page_title="BI Executivo - Monitor",
    layout="wide",
    page_icon="📊"
)

# ==============================
# HEADER
# ==============================
logo_html = ""
if os.path.exists(LOGO_PATH):
    with open(LOGO_PATH, "rb") as f:
        logo_base64 = base64.b64encode(f.read()).decode()
    logo_html = f'<img src="data:image/png;base64,{logo_base64}" width="120">'

st.markdown(f"""
<div style="background: linear-gradient(90deg, #0f2a44, #1f4e79);
padding: 18px 24px; border-radius: 14px; display: flex; align-items: center; gap: 20px; margin-bottom: 20px;">
    {logo_html}
    <div>
        <p style="color:white;font-size:26px;font-weight:700;margin:0;">
        Monitor de Pedidos — BI Executivo</p>
        <p style="color:white;opacity:0.85;margin:0;font-size:14px;">
        Análise de Risco Logístico • Transportadora • Status • Região • Cliente
        </p>
    </div>
</div>
""", unsafe_allow_html=True)

# ==============================
# FUNÇÕES
# ==============================
def ler_base(path):
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        df = pd.read_excel(path)
    except:
        return pd.DataFrame()

    if "PedidoFormatado" in df.columns:
        df["PedidoFormatado"] = (
            df["PedidoFormatado"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

    if "Status" in df.columns:
        df["Status"] = df["Status"].astype(str).str.strip()

    return df


def listar_dias():
    if not os.path.exists(PASTA_HIST):
        return []

    arquivos = os.listdir(PASTA_HIST)
    datas = set()

    for a in arquivos:
        m = re.match(r"(\d{2}-\d{2}-\d{4})_manha\.xlsx$", a)
        if m:
            datas.add(m.group(1))

    return sorted(datas, key=lambda x: pd.to_datetime(x, format="%d-%m-%Y"))


def caminho(dia):
    return os.path.join(PASTA_HIST, f"{dia}_manha.xlsx")


def pizza(tratados, restantes, titulo):
    fig, ax = plt.subplots(figsize=(2.3, 2.3))
    total = tratados + restantes

    if total == 0:
        ax.text(0.5, 0.5, "0", ha="center", va="center")
    else:
        ax.pie([tratados, restantes], autopct="%1.0f%%", startangle=90)

    ax.set_title(titulo, fontsize=9)
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()

# ==============================
# BI EXECUTIVO
# ==============================
dias = listar_dias()

if len(dias) < 2:
    st.warning("Histórico insuficiente.")
    st.stop()

dias = dias[-15:]

# ==============================
# 📈 STATUS DIÁRIOS (GRÁFICO PEQUENO)
# ==============================
st.markdown("### 📈 Status Diários")

contagem_status = []

for dia_hist in dias:
    df_temp = ler_base(caminho(dia_hist))
    if df_temp.empty:
        continue

    qtd = df_temp[df_temp["Status"].isin(STATUS_DIARIOS)].shape[0]
    contagem_status.append((dia_hist, qtd))

if contagem_status:

    df_graf = pd.DataFrame(contagem_status, columns=["Data", "Quantidade"])
    df_graf["Data"] = pd.to_datetime(df_graf["Data"], format="%d-%m-%Y")
    df_graf = df_graf.sort_values("Data")

    fig, ax = plt.subplots(figsize=(2.5, 1.4))  # metade do tamanho

    ax.plot(df_graf["Data"], df_graf["Quantidade"])

    ax.set_xticks(df_graf["Data"])
    ax.set_xticklabels(df_graf["Data"].dt.day, fontsize=7)

    ax.set_xlabel("Dia", fontsize=8)
    ax.set_ylabel("Qtde", fontsize=8)
    ax.set_title("Status Críticos", fontsize=9)

    st.pyplot(fig)
    plt.close(fig)

st.divider()

# ==============================
# LOOP ORIGINAL COMPLETO
# ==============================
for i in range(len(dias)-1, 0, -1):

    dia_atual = dias[i]
    dia_ant = dias[i-1]

    df_atual = ler_base(caminho(dia_atual))
    df_ant = ler_base(caminho(dia_ant))

    if df_atual.empty or df_ant.empty:
        continue

    st.markdown(f"### 📅 {dia_ant} ➜ {dia_atual}")

    col1, col2, col3 = st.columns(3)

    # TRIPLO
    with col1:
        if "Transportadora_Triplo" in df_atual.columns:

            atual = df_atual[df_atual["Transportadora_Triplo"]=="X"]
            ant = df_ant[df_ant["Transportadora_Triplo"]=="X"]

            tratados = ant[~ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
            restantes = ant[ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
            entrou = atual[~atual["PedidoFormatado"].isin(ant["PedidoFormatado"])]

            st.image(pizza(len(tratados), len(restantes), "Triplo"))

            st.markdown(f"Tratados: {len(tratados)} / {len(ant)}")
            st.markdown(f"Entraram: {len(entrou)}")

            buf = BytesIO()
            restantes.to_excel(buf, index=False)
            st.download_button(
                "Remanescentes Triplo",
                buf.getvalue(),
                file_name=f"remanescente_triplo_{dia_atual}.xlsx"
            )

    # STATUS 2X
    with col2:
        if "Status_Dobro" in df_atual.columns:

            atual = df_atual[df_atual["Status_Dobro"]=="X"]
            ant = df_ant[df_ant["Status_Dobro"]=="X"]

            tratados = ant[~ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
            restantes = ant[ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
            entrou = atual[~atual["PedidoFormatado"].isin(ant["PedidoFormatado"])]

            st.image(pizza(len(tratados), len(restantes), "Status 2x"))

            st.markdown(f"Tratados: {len(tratados)} / {len(ant)}")
            st.markdown(f"Entraram: {len(entrou)}")

            buf = BytesIO()
            restantes.to_excel(buf, index=False)
            st.download_button(
                "Remanescentes Status 2x",
                buf.getvalue(),
                file_name=f"remanescente_status_{dia_atual}.xlsx"
            )

    # REGIÃO 2X
    with col3:
        if "Regiao_Dobro" in df_atual.columns:

            atual = df_atual[df_atual["Regiao_Dobro"]=="X"]
            ant = df_ant[df_ant["Regiao_Dobro"]=="X"]

            tratados = ant[~ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
            restantes = ant[ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
            entrou = atual[~atual["PedidoFormatado"].isin(ant["PedidoFormatado"])]

            st.image(pizza(len(tratados), len(restantes), "Região 2x"))

            st.markdown(f"Tratados: {len(tratados)} / {len(ant)}")
            st.markdown(f"Entraram: {len(entrou)}")

            buf = BytesIO()
            restantes.to_excel(buf, index=False)
            st.download_button(
                "Remanescentes Região 2x",
                buf.getvalue(),
                file_name=f"remanescente_regiao_{dia_atual}.xlsx"
            )

    st.divider()