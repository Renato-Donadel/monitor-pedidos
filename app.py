import streamlit as st
import pandas as pd
import os
from io import BytesIO
import matplotlib.pyplot as plt
import re
from datetime import timedelta

# ==============================
# CONFIGS (SEMPRE PRIMEIRO)
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
# 🎨 TEMA VISUAL (ESTILO BRAVIUM)
# ==============================
st.markdown("""
<style>
.stApp {
    background-color: #f4f6f9;
}

/* Títulos menores (layout mais compacto) */
h1 { font-size: 28px !important; }
h2 { font-size: 22px !important; }
h3 { font-size: 18px !important; }

/* Cards métricos */
div[data-testid="metric-container"] {
    background: white;
    border-radius: 12px;
    padding: 14px;
    box-shadow: 0px 4px 14px rgba(0,0,0,0.05);
    border: 1px solid #e6e9ef;
}

/* Botões de download (meninas) */
.stDownloadButton > button {
    background: linear-gradient(90deg, #0b2c5f, #1f4e79);
    color: white;
    border-radius: 10px;
    font-weight: 700;
    height: 42px;
    width: 100%;
    border: none;
    font-size: 15px;
}

.stDownloadButton > button:hover {
    background: linear-gradient(90deg, #1f4e79, #2d6aa3);
    transform: scale(1.02);
}

/* Divisores */
hr {
    border: 1px solid #e6e9ef;
    margin-top: 25px;
    margin-bottom: 25px;
}
</style>
""", unsafe_allow_html=True)

# ==============================
# HEADER BRAVIUM (LOGO DENTRO DO AZUL)
# ==============================
col_header = st.container()

with col_header:
    col_logo, col_title = st.columns([1, 6])

    with col_logo:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=120)

    with col_title:
        st.markdown("""
        <div style="
            background: linear-gradient(90deg, #0b2c5f, #1f4e79);
            padding: 18px 24px;
            border-radius: 18px;
            color: white;
            height: 90px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        ">
            <div style="font-size:26px; font-weight:700;">
                Monitor de Pedidos — BI Executivo
            </div>
            <div style="font-size:14px; opacity:0.85;">
                Análise de Risco Logístico • Transportadora • Status • Região • Cliente
            </div>
        </div>
        """, unsafe_allow_html=True)

# ==============================
# FUNÇÕES BASE
# ==============================
def ler_base(path):
    if not os.path.exists(path):
        return pd.DataFrame()

    try:
        df = pd.read_excel(path)
    except Exception:
        return pd.DataFrame()

    # NORMALIZAÇÃO CRÍTICA (remove espaços bugados)
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
        m = re.match(r"(\\d{2}-\\d{2}-\\d{4})_manha\\.xlsx$", a)
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
        if pd.to_datetime  (d, format="%d-%m-%Y") <= alvo
    ]

    if not candidatos:
        return None

    return max(candidatos, key=lambda d: pd.to_datetime(d, format="%d-%m-%Y"))


def pizza(tratados, nao, titulo):
    fig, ax = plt.subplots(figsize=(2.8, 2.8))
    total = tratados + nao

    if total == 0:
        ax.text(0.5, 0.5, "0", ha="center", va="center")
    else:
        ax.pie([tratados, nao], autopct="%1.0f%%", startangle=90)

    ax.set_title(titulo, fontsize=11)
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()

# ==============================
# 📥 EXPORTAÇÃO POR CARTEIRA (300 EM 300)
# ==============================
st.markdown("## 📥 Exportação por Carteira (300 em 300)")

df_atual_base = ler_base(ARQ_ATUAL)

if not df_atual_base.empty and "Carteira" in df_atual_base.columns:

    if "Ranking" in df_atual_base.columns:
        df_atual_base = df_atual_base.sort_values("Ranking").reset_index(drop=True)

    if "offsets" not in st.session_state:
        st.session_state["offsets"] = {}

    carteiras = sorted(df_atual_base["Carteira"].dropna().unique())

    for carteira in carteiras:
        df_carteira = df_atual_base[
            df_atual_base["Carteira"] == carteira
        ].reset_index(drop=True)

        total = len(df_carteira)
        offset = st.session_state["offsets"].get(carteira, 0)
        inicio = offset
        fim = min(offset + TAMANHO_LOTE, total)

        lote = df_carteira.iloc[inicio:fim]

        if not lote.empty:
            buffer = BytesIO()
            lote.to_excel(buffer, index=False)
            buffer.seek(0)

            col1, col2 = st.columns([5, 2])

            with col1:
                st.markdown(f"**{carteira}** — Pedidos {inicio+1} até {fim} de {total}")

            with col2:
                if st.download_button(
                    label=f"⬇️ {carteira}",
                    data=buffer,
                    file_name=f"{carteira}_{inicio+1}_a_{fim}.xlsx",
                    key=f"download_{carteira}_{offset}"
                ):
                    st.session_state["offsets"][carteira] = fim

st.divider()

# ==============================
# 📊 BI EXECUTIVO (COMPARAÇÃO DIA A DIA)
# ==============================
dias = listar_dias()

if len(dias) < 2:
    st.warning("Histórico insuficiente na pasta data/historico.")
    st.stop()

dias = dias[-15:]

tend_triplo = []
datas_plot = []

st.markdown("## 📊 BI Executivo — Comparação Dia a Dia")

# LOOP DO MAIS NOVO PARA O MAIS ANTIGO
for i in range(len(dias) - 1, 0, -1):

    dia_atual = dias[i]
    dia_ant = dias[i - 1]

    df_atual = ler_base(caminho(dia_atual))
    df_ant = ler_base(caminho(dia_ant))

    if df_atual.empty or df_ant.empty:
        continue

    st.markdown(f"### 📅 {dia_ant} ➜ {dia_atual}")

    # ================= TRIPLO TRANSPORTADORA =================
    if "Transportadora_Triplo" in df_atual.columns:

        triplo_atual = df_atual[df_atual["Transportadora_Triplo"] == "X"]
        triplo_ant = df_ant[df_ant["Transportadora_Triplo"] == "X"]

        tratados = triplo_ant[
            ~triplo_ant["PedidoFormatado"].isin(triplo_atual["PedidoFormatado"])
        ]

        persist_d1 = triplo_ant[
            triplo_ant["PedidoFormatado"].isin(triplo_atual["PedidoFormatado"])
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

        # LAYOUT 3 COLUNAS COMPACTO (COMO VOCÊ PEDIU)
        col_triplo, col_status, col_regiao = st.columns(3)

        with col_triplo:
            st.markdown("#### 🔴 Triplo Transportadora")
            st.image(pizza(len(tratados), len(persist_d1), "Tratados vs Restantes"))
            st.metric("Entraram", len(entrou))

            buf = BytesIO()
            persist_72h.to_excel(buf, index=False)
            st.download_button(
                "⬇️ Triplo > 72h",
                buf.getvalue(),
                file_name=f"triplo_72h_{dia_atual}.xlsx",
                key=f"triplo_{dia_atual}"
            )

        # ================= STATUS 2X =================
        if "Status_Dobro" in df_atual.columns:
            status_atual = df_atual[df_atual["Status_Dobro"] == "X"]
            status_ant = df_ant[df_ant["Status_Dobro"] == "X"]

            persist_status = status_ant[
                status_ant["PedidoFormatado"].isin(status_atual["PedidoFormatado"])
            ]

            entrou_status = status_atual[
                ~status_atual["PedidoFormatado"].isin(status_ant["PedidoFormatado"])
            ]

            with col_status:
                st.markdown("#### 🟡 Status 2x Prazo")
                st.image(pizza(len(status_ant)-len(persist_status), len(persist_status), "Tratados vs Persistentes"))
                st.metric("Entraram", len(entrou_status))

                buf2 = BytesIO()
                persist_status.to_excel(buf2, index=False)
                st.download_button(
                    "⬇️ Status Persistente",
                    buf2.getvalue(),
                    file_name=f"status_2x_{dia_atual}.xlsx",
                    key=f"status_{dia_atual}"
                )

        # ================= REGIÃO 2X =================
        if "Regiao_Dobro" in df_atual.columns:
            reg_atual = df_atual[df_atual["Regiao_Dobro"] == "X"]
            reg_ant = df_ant[df_ant["Regiao_Dobro"] == "X"]

            persist_reg = reg_ant[
                reg_ant["PedidoFormatado"].isin(reg_atual["PedidoFormatado"])
            ]

            entrou_reg = reg_atual[
                ~reg_atual["PedidoFormatado"].isin(reg_ant["PedidoFormatado"])
            ]

            with col_regiao:
                st.markdown("#### 🔵 Região 2x Prazo")
                st.image(pizza(len(reg_ant)-len(persist_reg), len(persist_reg), "Tratados vs Persistentes"))
                st.metric("Entraram", len(entrou_reg))

                buf3 = BytesIO()
                persist_reg.to_excel(buf3, index=False)
                st.download_button(
                    "⬇️ Região Persistente",
                    buf3.getvalue(),
                    file_name=f"regiao_2x_{dia_atual}.xlsx",
                    key=f"regiao_{dia_atual}"
                )

        tend_triplo.append(len(triplo_atual))
        datas_plot.append(dia_atual)

    st.divider()

# ==============================
# 📈 TENDÊNCIA TRIPLO (SEM ERRO)
# ==============================
if len(tend_triplo) > 0:
    st.markdown("## 📈 Tendência — Triplo Transportadora")

    fig, ax = plt.subplots(figsize=(8, 3))
    ax.plot(datas_plot[::-1], tend_triplo[::-1])
    ax.set_title("Triplo ao longo do tempo (mais recente à direita)")
    ax.tick_params(axis='x', rotation=45)

    st.pyplot(fig)
