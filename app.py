import streamlit as st
import pandas as pd
import os
from io import BytesIO
import matplotlib.pyplot as plt
import re
from datetime import datetime

# ==============================
# CONFIGS
# ==============================
BASE_DIR = os.path.dirname(__file__)
PASTA_DATA = os.path.join(BASE_DIR, "data")
PASTA_HIST = os.path.join(PASTA_DATA, "historico")

st.set_page_config(page_title="BI Executivo - Monitor", layout="wide")

# ==============================
# FUNÇÕES BASE
# ==============================
def ler_base(path):
    if not os.path.exists(path):
        return pd.DataFrame()
    return pd.read_excel(path)

def listar_dias():
    arquivos = os.listdir(PASTA_HIST)
    datas = set()
    for a in arquivos:
        m = re.match(r"(\d{2}-\d{2}-\d{4})_manha\.xlsx$", a)
        if m:
            datas.add(m.group(1))
    return sorted(datas, key=lambda x: pd.to_datetime(x, format="%d-%m-%Y"))

def caminho(dia):
    return os.path.join(PASTA_HIST, f"{dia}_manha.xlsx")

def pizza(tratados, nao, titulo):
    fig, ax = plt.subplots(figsize=(2.5,2.5))
    if tratados + nao == 0:
        ax.text(0.5,0.5,"0",ha="center")
    else:
        ax.pie([tratados, nao], autopct="%1.0f%%", startangle=90)
    ax.set_title(titulo)
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()

# ==============================
# INÍCIO
# ==============================
st.title("📊 BI Executivo — Monitor de Risco")

dias = listar_dias()
if len(dias) < 2:
    st.warning("Histórico insuficiente.")
    st.stop()

dias = dias[-15:]

tend_triplo = []
datas_plot = []

# ==============================
# LOOP DIA A DIA
# ==============================
for i in range(len(dias)-1):

    dia_ant = dias[i]
    dia_atual = dias[i+1]

    df_ant = ler_base(caminho(dia_ant))
    df_atual = ler_base(caminho(dia_atual))

    if df_ant.empty or df_atual.empty:
        continue

    st.markdown(f"# 📅 {dia_ant} ➜ {dia_atual}")

    # =====================================================
    # 🔴 TRIPLO TRANSPORTADORA
    # =====================================================
    st.markdown("## 🔴 Triplo Transportadora")

    triplo_ant = df_ant[df_ant["Transportadora_Triplo"]=="X"]
    triplo_atual = df_atual[df_atual["Transportadora_Triplo"]=="X"]

    tratados = triplo_ant[
        ~triplo_ant["PedidoFormatado"].isin(triplo_atual["PedidoFormatado"])
    ]

    nao_tratados = triplo_ant[
        triplo_ant["PedidoFormatado"].isin(triplo_atual["PedidoFormatado"])
    ]

    entrou = triplo_atual[
        ~triplo_atual["PedidoFormatado"].isin(triplo_ant["PedidoFormatado"])
    ]

    persist_72 = nao_tratados[
        (pd.to_datetime(dia_atual, format="%d-%m-%Y") -
         pd.to_datetime(nao_tratados["DataÚltimoStatus"])
        ).dt.days >= 3
    ]

    c1,c2,c3 = st.columns(3)
    with c1:
        st.image(pizza(len(tratados), len(nao_tratados),
                       f"Tratados\n{len(tratados)}/{len(triplo_ant)}"))

    with c2:
        st.metric("Entraram no Triplo", len(entrou))

    with c3:
        st.metric("Persist ≥72h", len(persist_72))
        buffer = BytesIO()
        persist_72.to_excel(buffer,index=False)
        st.download_button(
            "Exportar Persistentes 72h",
            buffer.getvalue(),
            file_name=f"triplo_72h_{dia_atual}.xlsx"
        )

    tend_triplo.append(len(triplo_atual))
    datas_plot.append(dia_atual)

    # =====================================================
    # 🟡 STATUS 2x (USANDO FLAG)
    # =====================================================
    st.markdown("## 🟡 Status Específico 2x Prazo")

    status_ant = df_ant[df_ant["Status_Dobro"]=="X"]
    status_atual = df_atual[df_atual["Status_Dobro"]=="X"]

    persist_status = status_ant[
        status_ant["PedidoFormatado"].isin(status_atual["PedidoFormatado"])
    ]

    entrou_status = status_atual[
        ~status_atual["PedidoFormatado"].isin(status_ant["PedidoFormatado"])
    ]

    c1,c2,c3 = st.columns(3)
    with c1:
        st.metric("Críticos Status 2x", len(status_ant))
    with c2:
        st.metric("Entraram", len(entrou_status))
    with c3:
        st.metric("Persistentes", len(persist_status))

        buffer = BytesIO()
        persist_status.to_excel(buffer,index=False)
        st.download_button(
            "Exportar Status Persistente",
            buffer.getvalue(),
            file_name=f"status_2x_{dia_atual}.xlsx"
        )

    # =====================================================
    # 🔵 REGIÃO 2x (USANDO FLAG)
    # =====================================================
    st.markdown("## 🔵 Região 2x Prazo")

    reg_ant = df_ant[df_ant["Regiao_Dobro"]=="X"]
    reg_atual = df_atual[df_atual["Regiao_Dobro"]=="X"]

    persist_reg = reg_ant[
        reg_ant["PedidoFormatado"].isin(reg_atual["PedidoFormatado"])
    ]

    entrou_reg = reg_atual[
        ~reg_atual["PedidoFormatado"].isin(reg_ant["PedidoFormatado"])
    ]

    c1,c2,c3 = st.columns(3)
    with c1:
        st.metric("Críticos Região 2x", len(reg_ant))
    with c2:
        st.metric("Entraram", len(entrou_reg))
    with c3:
        st.metric("Persistentes", len(persist_reg))

        buffer = BytesIO()
        persist_reg.to_excel(buffer,index=False)
        st.download_button(
            "Exportar Região Persistente",
            buffer.getvalue(),
            file_name=f"regiao_2x_{dia_atual}.xlsx"
        )

    st.divider()

# ==============================
# 📈 TENDÊNCIA
# ==============================
st.markdown("# 📈 Tendência Triplo")

fig, ax = plt.subplots()
ax.plot(datas_plot, tend_triplo)
ax.set_title("Triplo ao longo do tempo")
ax.tick_params(axis='x', rotation=45)
st.pyplot(fig)
