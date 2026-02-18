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
ARQ_ATUAL = os.path.join(PASTA_DATA, "Monitor_Pedidos_Processado.xlsx")

TAMANHO_LOTE = 300

st.set_page_config(page_title="BI Executivo - Monitor", layout="wide")

# ==============================
# FUNÇÕES
# ==============================
def ler_base(path):
    if not os.path.exists(path):
        return pd.DataFrame()
    return pd.read_excel(path)

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

def pizza(tratados, nao, titulo):
    fig, ax = plt.subplots(figsize=(2.5,2.5))
    total = tratados + nao
    if total == 0:
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
st.title("📊 Monitor de Pedidos — Operacional & Executivo")

# =====================================================
# 📥 PARTE 1 — CARTEIRAS (300 EM 300)
# =====================================================
df_atual = ler_base(ARQ_ATUAL)

if not df_atual.empty and "Carteira" in df_atual.columns:

    st.markdown("## 📥 Carteiras — Download por Lote")

    if "Ranking" in df_atual.columns:
        df_atual = df_atual.sort_values("Ranking").reset_index(drop=True)

    if "offsets" not in st.session_state:
        st.session_state["offsets"] = {}

    carteiras = sorted(df_atual["Carteira"].dropna().unique())

    for carteira in carteiras:

        df_carteira = df_atual[df_atual["Carteira"] == carteira].reset_index(drop=True)
        total = len(df_carteira)

        offset = st.session_state["offsets"].get(carteira, 0)
        inicio = offset
        fim = min(offset + TAMANHO_LOTE, total)

        col1, col2 = st.columns([4,1])

        with col1:
            st.write(f"**{carteira}** — {inicio+1} a {fim} de {total}")

        with col2:
            if st.button("📥", key=f"btn_{carteira}"):

                df_lote = df_carteira.iloc[inicio:fim]

                if not df_lote.empty:
                    st.session_state["offsets"][carteira] = fim

                    buffer = BytesIO()
                    df_lote.to_excel(buffer, index=False)
                    buffer.seek(0)

                    st.download_button(
                        label="⬇️ Excel",
                        data=buffer,
                        file_name=f"Pedidos_{carteira}_{inicio+1}_a_{fim}.xlsx"
                    )

    st.divider()

# =====================================================
# 📊 PARTE 2 — BI EXECUTIVO
# =====================================================

dias = listar_dias()

if len(dias) < 2:
    st.warning("Histórico insuficiente.")
    st.stop()

dias = dias[-15:]

tend_triplo = []
datas_plot = []

# 🔥 LOOP DO MAIS NOVO PARA O MAIS ANTIGO
for i in range(len(dias)-1, 0, -1):

    dia_atual = dias[i]
    dia_ant = dias[i-1]

    df_atual = ler_base(caminho(dia_atual))
    df_ant = ler_base(caminho(dia_ant))

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

    c1,c2,c3 = st.columns(3)

    with c1:
        st.image(pizza(len(tratados), len(nao_tratados),
                       f"Tratados\n{len(tratados)}/{len(triplo_ant)}"))

    with c2:
        st.metric("Entraram no Triplo", len(entrou))

    with c3:
        st.metric("Persistentes", len(nao_tratados))
        buffer = BytesIO()
        nao_tratados.to_excel(buffer,index=False)
        st.download_button(
            "Exportar Persistentes",
            buffer.getvalue(),
            file_name=f"triplo_persist_{dia_atual}.xlsx"
        )

    tend_triplo.append(len(triplo_atual))
    datas_plot.append(dia_atual)

    # =====================================================
    # 🟡 STATUS 2x
    # =====================================================
    st.markdown("## 🟡 Status 2x Prazo")

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
        st.metric("Críticos", len(status_ant))
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
    # 🔵 REGIÃO 2x
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
        st.metric("Críticos Região", len(reg_ant))
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

# =====================================================
# 📈 TENDÊNCIA TRIPLO
# =====================================================
if tend_triplo:
    st.markdown("# 📈 Tendência Triplo")

    fig, ax = plt.subplots()
    ax.plot(datas_plot[::-1], tend_triplo[::-1])
    ax.set_title("Triplo ao longo do tempo")
    ax.tick_params(axis='x', rotation=45)
    st.pyplot(fig)
