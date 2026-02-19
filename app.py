import streamlit as st
import pandas as pd
import os
from io import BytesIO
import matplotlib.pyplot as plt
import re
from datetime import timedelta
import base64

# ==============================
# CONFIGS (PRIMEIRO DE TUDO)
# ==============================
BASE_DIR = os.path.dirname(__file__)
PASTA_DATA = os.path.join(BASE_DIR, "data")
PASTA_HIST = os.path.join(PASTA_DATA, "historico")
ARQ_ATUAL = os.path.join(PASTA_DATA, "Monitor_Pedidos_Processado.xlsx")
LOGO_PATH = os.path.join(PASTA_DATA, "logo_bravium.png")

TAMANHO_LOTE = 300

st.set_page_config(
    page_title="Monitor de Pedidos — BI Executivo",
    layout="wide",
    page_icon="📊"
)

# ==============================
# 🎨 CSS ESTILO BRAVIUM (COMPACTO)
# ==============================
st.markdown("""
<style>
.stApp {
    background-color: #f4f6f9;
}

h1 {font-size: 26px !important;}
h2 {font-size: 22px !important;}
h3 {font-size: 18px !important;}

.header-box {
    background: linear-gradient(90deg, #0f2a44, #1f4e79);
    padding: 18px 24px;
    border-radius: 18px;
    color: white;
    display: flex;
    align-items: center;
    gap: 18px;
    margin-bottom: 10px;
}

.stDownloadButton > button {
    background: linear-gradient(90deg, #0f2a44, #1f4e79);
    color: white;
    border-radius: 12px;
    font-weight: 700;
    height: 46px;
    width: 100%;
    border: none;
    font-size: 15px;
}

.stDownloadButton > button:hover {
    background: linear-gradient(90deg, #1f4e79, #2d6aa3);
}

div[data-testid="metric-container"] {
    background: white;
    border-radius: 12px;
    padding: 12px;
    border: 1px solid #e6e9ef;
}
</style>
""", unsafe_allow_html=True)

# ==============================
# HEADER COM LOGO DENTRO DA FAIXA AZUL (CORRIGIDO)
# ==============================
logo_base64 = ""
if os.path.exists(LOGO_PATH):
    with open(LOGO_PATH, "rb") as f:
        logo_base64 = base64.b64encode(f.read()).decode()

st.markdown(f"""
<div class="header-box">
    <img src="data:image/png;base64,{logo_base64}" width="120">
    <div>
        <div style="font-size:26px; font-weight:700;">
            Monitor de Pedidos — BI Executivo
        </div>
        <div style="opacity:0.85; font-size:14px;">
            Análise de Risco Logístico • Transportadora • Status • Região • Cliente
        </div>
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
    except Exception:
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

    # REGEX CORRIGIDO (SEM \\d)
    for a in arquivos:
        m = re.search(r"(\d{2}-\d{2}-\d{4})", a)
        if m:
            datas.add(m.group(1))

    return sorted(datas, key=lambda x: pd.to_datetime(x, format="%d-%m-%Y"))


def caminho(dia):
    return os.path.join(PASTA_HIST, f"{dia}_manha.xlsx")


def encontrar_dia_72h(dias, dia_atual):
    data_atual = pd.to_datetime(dia_atual, format="%d-%m-%Y")
    alvo = data_atual - timedelta(days=3)

    candidatos = [
        d for d in dias
        if pd.to_datetime(d, format="%d-%m-%Y") <= alvo
    ]

    if not candidatos:
        return None

    return max(candidatos, key=lambda d: pd.to_datetime(d, format="%d-%m-%Y"))


def pizza(tratados, nao, titulo):
    fig, ax = plt.subplots(figsize=(2.6, 2.6))
    total = tratados + nao

    if total == 0:
        ax.text(0.5, 0.5, "0", ha="center", va="center")
    else:
        ax.pie([tratados, nao], autopct="%1.0f%%", startangle=90)

    ax.set_title(titulo, fontsize=10)
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()

# ==============================
# 📥 EXPORTAÇÃO POR CARTEIRA (LOTE PROGRESSIVO CORRIGIDO)
# ==============================
st.markdown("## 📥 Exportação por Carteira (300 em 300)")

df_atual_base = ler_base(ARQ_ATUAL)

if not df_atual_base.empty and "Carteira" in df_atual_base.columns:

    if "Ranking" in df_atual_base.columns:
        df_atual_base = df_atual_base.sort_values("Ranking").reset_index(drop=True)

    # INICIALIZA CONTROLE DE LOTES (CRÍTICO)
    if "offsets" not in st.session_state:
        st.session_state["offsets"] = {}

    carteiras = sorted(df_atual_base["Carteira"].dropna().unique())

    for carteira in carteiras:

        if carteira not in st.session_state["offsets"]:
            st.session_state["offsets"][carteira] = 0

        offset = st.session_state["offsets"][carteira]

        df_carteira = df_atual_base[
            df_atual_base["Carteira"] == carteira
        ].reset_index(drop=True)

        total = len(df_carteira)

        # RESET AUTOMÁTICO SE ACABAR
        if offset >= total:
            offset = 0
            st.session_state["offsets"][carteira] = 0

        inicio = offset
        fim = min(offset + TAMANHO_LOTE, total)

        lote = df_carteira.iloc[inicio:fim]

        col1, col2 = st.columns([5, 2])

        with col1:
            st.markdown(
                f"**{carteira}** — Pedidos {inicio+1} até {fim} de {total}"
            )

        with col2:
            if not lote.empty:
                buffer = BytesIO()
                lote.to_excel(buffer, index=False)
                buffer.seek(0)

                # BOTÃO QUE AVANÇA O LOTE CORRETAMENTE
                clicked = st.download_button(
                    label=f"⬇️ {carteira}",
                    data=buffer,
                    file_name=f"{carteira}_{inicio+1}_a_{fim}.xlsx",
                    key=f"download_{carteira}_{offset}"
                )

                # ATUALIZA OFFSET APÓS CLIQUE (SOLUÇÃO DEFINITIVA)
                if clicked:
                    st.session_state["offsets"][carteira] = fim

st.divider()

# ==============================
# 📊 BI HISTÓRICO
# ==============================
dias = listar_dias()

if len(dias) < 2:
    st.warning("Histórico insuficiente na pasta data/historico.")
    st.stop()

dias = dias[-15:]
tend_triplo = []
datas_plot = []

st.markdown("## 📊 Comparação Histórica")

for i in range(len(dias) - 1, 0, -1):
    dia_atual = dias[i]
    dia_ant = dias[i - 1]

    df_atual = ler_base(caminho(dia_atual))
    df_ant = ler_base(caminho(dia_ant))

    if df_atual.empty or df_ant.empty:
        continue

    st.markdown(f"### 📅 {dia_ant} ➜ {dia_atual}")

    if "Transportadora_Triplo" not in df_atual.columns:
        continue

    triplo_atual = df_atual[df_atual["Transportadora_Triplo"] == "X"]
    triplo_ant = df_ant[df_ant["Transportadora_Triplo"] == "X"]

    tratados = triplo_ant[
        ~triplo_ant["PedidoFormatado"].isin(triplo_atual["PedidoFormatado"])
    ]

    entrou = triplo_atual[
        ~triplo_atual["PedidoFormatado"].isin(triplo_ant["PedidoFormatado"])
    ]

    dia_72 = encontrar_dia_72h(dias, dia_atual)

    if dia_72:
        df_72 = ler_base(caminho(dia_72))
        triplo_72 = df_72[df_72["Transportadora_Triplo"] == "X"]
        persist_72h = triplo_atual[
            triplo_atual["PedidoFormatado"].isin(triplo_72["PedidoFormatado"])
        ]
    else:
        persist_72h = pd.DataFrame()

    col1, col2, col3 = st.columns([2, 1, 1])

    with col1:
        st.image(
            pizza(
                len(tratados),
                len(persist_72h),
                f"Tratados {len(tratados)} / {len(triplo_ant)}"
            )
        )
        st.caption(f"Entraram no Triplo: {len(entrou)}")

    with col2:
        st.metric("Triplo > 72h", len(persist_72h))

    with col3:
        buffer = BytesIO()
        persist_72h.to_excel(buffer, index=False)
        st.download_button(
            "Exportar Triplo > 72h",
            buffer.getvalue(),
            file_name=f"triplo_72h_{dia_atual}.xlsx"
        )

    tend_triplo.append(len(triplo_atual))
    datas_plot.append(dia_atual)

# ==============================
# 📈 TENDÊNCIA (SEM BUG)
# ==============================
if len(tend_triplo) > 0:
    st.markdown("## 📈 Tendência Triplo Transportadora")

    fig, ax = plt.subplots(figsize=(8, 3))
    ax.plot(datas_plot[::-1], tend_triplo[::-1])
    ax.set_title("Triplo ao longo do tempo")
    ax.tick_params(axis='x', rotation=45)
    st.pyplot(fig)
