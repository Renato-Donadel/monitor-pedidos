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
    page_title="BI Executivo - Monitor",
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

/* HEADER AZUL COM LOGO DENTRO */
.header-box {
    background: linear-gradient(90deg, #0f2a44, #1f4e79);
    padding: 18px 24px;
    border-radius: 14px;
    display: flex;
    align-items: center;
    gap: 20px;
    margin-bottom: 10px;
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

/* Pizzas mais compactas */
img {
    max-width: 220px !important;
}

/* Botões bonitos */
.stDownloadButton > button {
    background: linear-gradient(90deg, #0f2a44, #1f4e79);
    color: white;
    border-radius: 10px;
    font-weight: 700;
    height: 42px;
    width: 100%;
    border: none;
    font-size: 14px;
}

.stDownloadButton > button:hover {
    background: linear-gradient(90deg, #1f4e79, #2d6aa3);
}

/* Título das datas menor */
.data-title {
    font-size: 20px;
    font-weight: 700;
    color: #0f2a44;
    margin-top: 10px;
    margin-bottom: 5px;
}
</style>
""", unsafe_allow_html=True)

# ==============================
# HEADER COM LOGO DENTRO DA FAIXA AZUL
# ==============================
logo_html = ""
if os.path.exists(LOGO_PATH):
    import base64
    with open(LOGO_PATH, "rb") as f:
        logo_base64 = base64.b64encode(f.read()).decode()
    logo_html = f'<img src="data:image/png;base64,{logo_base64}" width="120">'

st.markdown(f"""
<div class="header-box">
    {logo_html}
    <div>
        <p class="header-title">Monitor de Pedidos — BI Executivo</p>
        <p class="header-sub">
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
    except Exception:
        return pd.DataFrame()

    # NORMALIZAÇÃO CRÍTICA (resolve bug do espaço)
    if "PedidoFormatado" in df.columns:
        df["PedidoFormatado"] = (
            df["PedidoFormatado"]
            .astype(str)
            .str.strip()
            .str.upper()
        )
    return df


def listar_dias():
    if not os.path.exists(PASTA_HIST):
        return []

    arquivos = os.listdir(PASTA_HIST)
    datas = set()

    for a in arquivos:
        # REGEX CORRETA (SEM \\ DUPLO)
        m = re.match(r"(\d{2}-\d{2}-\d{4})_manha\.xlsx$", a)
        if m:
            datas.add(m.group(1))

    return sorted(datas, key=lambda x: pd.to_datetime(x, format="%d-%m-%Y"))


def caminho(dia):
    return os.path.join(PASTA_HIST, f"{dia}_manha.xlsx")


def pizza(tratados, restantes, titulo):
    fig, ax = plt.subplots(figsize=(2.4, 2.4))
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

# ==============================
# 📥 DOWNLOAD 300 EM 300 (CORRIGIDO)
# ==============================
st.markdown("### 📥 Exportação por Carteira (300 em 300)")

df_atual_base = ler_base(ARQ_ATUAL)

if "offsets" not in st.session_state:
    st.session_state["offsets"] = {}

if not df_atual_base.empty and "Carteira" in df_atual_base.columns:

    if "Ranking" in df_atual_base.columns:
        df_atual_base = df_atual_base.sort_values("Ranking").reset_index(drop=True)

    carteiras = sorted(df_atual_base["Carteira"].dropna().unique())

    for carteira in carteiras:
        df_carteira = df_atual_base[df_atual_base["Carteira"] == carteira].reset_index(drop=True)
        total = len(df_carteira)

        offset = st.session_state["offsets"].get(carteira, 0)
        inicio = offset
        fim = min(offset + TAMANHO_LOTE, total)

        lote = df_carteira.iloc[inicio:fim]

        if not lote.empty:
            buffer = BytesIO()
            lote.to_excel(buffer, index=False)
            buffer.seek(0)

            col1, col2 = st.columns([4, 2])

            with col1:
                st.write(f"**{carteira}** — {inicio+1} até {fim} de {total}")

            with col2:
                clicked = st.download_button(
                    label=f"⬇️ Baixar {carteira}",
                    data=buffer,
                    file_name=f"{carteira}_{inicio+1}_a_{fim}.xlsx",
                    key=f"dl_{carteira}_{offset}"
                )

                if clicked:
                    st.session_state["offsets"][carteira] = fim

st.divider()

# ==============================
# 📊 BI EXECUTIVO (3 PIZZAS LADO A LADO)
# ==============================
dias = listar_dias()

if len(dias) < 2:
    st.warning("Histórico insuficiente na pasta data/historico.")
    st.stop()

# últimos 15 dias
dias = dias[-15:]

for i in range(len(dias)-1, 0, -1):

    dia_atual = dias[i]
    dia_ant = dias[i-1]

    df_atual = ler_base(caminho(dia_atual))
    df_ant = ler_base(caminho(dia_ant))

    if df_atual.empty or df_ant.empty:
        continue

    st.markdown(f'<p class="data-title">📅 {dia_ant} ➜ {dia_atual}</p>', unsafe_allow_html=True)

    colA, colB, colC = st.columns(3)

    # ================= TRIPLO =================
    with colA:
        if "Transportadora_Triplo" in df_atual.columns:
            triplo_atual = df_atual[df_atual["Transportadora_Triplo"]=="X"]
            triplo_ant = df_ant[df_ant["Transportadora_Triplo"]=="X"]

            tratados = triplo_ant[
                ~triplo_ant["PedidoFormatado"].isin(triplo_atual["PedidoFormatado"])
            ]

            restantes = triplo_ant[
                triplo_ant["PedidoFormatado"].isin(triplo_atual["PedidoFormatado"])
            ]

            entrou = triplo_atual[
                ~triplo_atual["PedidoFormatado"].isin(triplo_ant["PedidoFormatado"])
            ]

            st.image(pizza(len(tratados), len(restantes), "Triplo Transportadora"))
            st.caption(f"Entraram: {len(entrou)}")

            buf = BytesIO()
            restantes.to_excel(buf, index=False)
            st.download_button(
                "Remanescentes Triplo",
                buf.getvalue(),
                file_name=f"remanescente_triplo_{dia_atual}.xlsx"
            )

    # ================= STATUS 2X =================
    with colB:
        if "Status_Dobro" in df_atual.columns:
            s_atual = df_atual[df_atual["Status_Dobro"]=="X"]
            s_ant = df_ant[df_ant["Status_Dobro"]=="X"]

            tratados = s_ant[
                ~s_ant["PedidoFormatado"].isin(s_atual["PedidoFormatado"])
            ]

            restantes = s_ant[
                s_ant["PedidoFormatado"].isin(s_atual["PedidoFormatado"])
            ]

            entrou = s_atual[
                ~s_atual["PedidoFormatado"].isin(s_ant["PedidoFormatado"])
            ]

            st.image(pizza(len(tratados), len(restantes), "Status Específico 2x"))
            st.caption(f"Entraram: {len(entrou)}")

            buf = BytesIO()
            restantes.to_excel(buf, index=False)
            st.download_button(
                "Remanescentes Status 2x",
                buf.getvalue(),
                file_name=f"remanescente_status_{dia_atual}.xlsx"
            )

    # ================= REGIÃO 2X =================
    with colC:
        if "Regiao_Dobro" in df_atual.columns:
            r_atual = df_atual[df_atual["Regiao_Dobro"]=="X"]
            r_ant = df_ant[df_ant["Regiao_Dobro"]=="X"]

            tratados = r_ant[
                ~r_ant["PedidoFormatado"].isin(r_atual["PedidoFormatado"])
            ]

            restantes = r_ant[
                r_ant["PedidoFormatado"].isin(r_atual["PedidoFormatado"])
            ]

            entrou = r_atual[
                ~r_atual["PedidoFormatado"].isin(r_ant["PedidoFormatado"])
            ]

            st.image(pizza(len(tratados), len(restantes), "Região 2x Prazo"))
            st.caption(f"Entraram: {len(entrou)}")

            buf = BytesIO()
            restantes.to_excel(buf, index=False)
            st.download_button(
                "Remanescentes Região 2x",
                buf.getvalue(),
                file_name=f"remanescente_regiao_{dia_atual}.xlsx"
            )

    st.divider()
