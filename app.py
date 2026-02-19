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

st.set_page_config(
    page_title="BI Executivo - Monitor",
    layout="wide",
    page_icon="📊"
)

# ==============================
# 🎨 ESTILO (MANTIDO IGUAL)
# ==============================
st.markdown("""
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

img {
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
    font-size: 16px;
    font-weight: 600;
    color: #0f2a44;
}

.stDownloadButton > button {
    background: linear-gradient(90deg, #0f2a44, #1f4e79);
    color: white;
    border-radius: 10px;
    font-weight: 700;
    height: 40px;
    width: 100%;
    border: none;
}
</style>
""", unsafe_allow_html=True)

# ==============================
# HEADER COM LOGO DENTRO DA FAIXA AZUL (MANTIDO)
# ==============================
logo_html = ""
if os.path.exists(LOGO_PATH):
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

    ax.set_title(titulo, fontsize=10)
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()

# ==============================
# 📥 DOWNLOAD POR CARTEIRA (SEQUENCIAL 300 EM 300 + IGOR)
# ==============================
st.markdown("### 📥 Exportação por Carteira (300 em 300)")

df_atual_base = ler_base(ARQ_ATUAL)

if "offsets_carteira" not in st.session_state:
    st.session_state["offsets_carteira"] = {}

if not df_atual_base.empty and "Carteira" in df_atual_base.columns:

    if "Ranking" in df_atual_base.columns:
        df_atual_base = df_atual_base.sort_values("Ranking").reset_index(drop=True)

    carteiras = sorted(df_atual_base["Carteira"].dropna().unique())

    # GARANTE QUE IGOR APAREÇA
    if "Igor" in df_atual_base["Carteira"].values and "Igor" not in carteiras:
        carteiras.append("Igor")

    for carteira in carteiras:

        df_carteira = df_atual_base[
            df_atual_base["Carteira"] == carteira
        ].reset_index(drop=True)

        total = len(df_carteira)

        offset = st.session_state["offsets_carteira"].get(carteira, 0)

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
                if st.download_button(
                    label=f"⬇️ Baixar {carteira}",
                    data=buffer,
                    file_name=f"{carteira}_{inicio+1}_a_{fim}.xlsx",
                    key=f"dl_{carteira}_{offset}"
                ):
                    st.session_state["offsets_carteira"][carteira] = fim

st.divider()

# ==============================
# 📊 BI EXECUTIVO (MANTIDO D-1)
# ==============================
dias = listar_dias()

if len(dias) < 2:
    st.warning("Histórico insuficiente na pasta data/historico.")
    st.stop()

dias = dias[-15:]

for i in range(len(dias)-1, 0, -1):

    dia_atual = dias[i]
    dia_ant = dias[i-1]

    df_atual = ler_base(caminho(dia_atual))
    df_ant = ler_base(caminho(dia_ant))

    if df_atual.empty or df_ant.empty:
        continue

    st.markdown(
        f'<p class="data-title">📅 {dia_ant} ➜ {dia_atual}</p>',
        unsafe_allow_html=True
    )

    col1, col2, col3 = st.columns(3)

    # ================= TRIPLO (PIZZA MANTIDA D-1) =================
    with col1:
        if "Transportadora_Triplo" in df_atual.columns:

            atual = df_atual[df_atual["Transportadora_Triplo"]=="X"]
            ant = df_ant[df_ant["Transportadora_Triplo"]=="X"]

            tratados = ant[~ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
            restantes = ant[ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
            entrou = atual[~atual["PedidoFormatado"].isin(ant["PedidoFormatado"])]

            st.image(pizza(len(tratados), len(restantes), "Triplo Transportadora"))

            st.markdown(
                f'<p class="metric-small">Tratados: {len(tratados)} / {len(ant)}</p>',
                unsafe_allow_html=True
            )
            st.markdown(
                f'<p class="metric-small">Entraram: {len(entrou)}</p>',
                unsafe_allow_html=True
            )

            # 🔥 REMANESCENTE NOVO (MATEMÁTICO, NÃO D-1)
            if (
                "DiasDesdeExpedicao" in atual.columns and
                "PrazoTransportadorDiasUteis" in atual.columns
            ):
                limite = (atual["PrazoTransportadorDiasUteis"] * 3) + 3
                remanescente_triplo = atual[
                    atual["DiasDesdeExpedicao"] > limite
                ].copy()
            else:
                remanescente_triplo = restantes.copy()

            buf = BytesIO()
            remanescente_triplo.to_excel(buf, index=False)
            st.download_button(
                "Remanescentes Triplo",
                buf.getvalue(),
                file_name=f"remanescente_triplo_{dia_atual}.xlsx"
            )

    # ================= STATUS 2X =================
    with col2:
        if "Status_Dobro" in df_atual.columns:

            atual = df_atual[df_atual["Status_Dobro"]=="X"]
            ant = df_ant[df_ant["Status_Dobro"]=="X"]

            tratados = ant[~ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
            restantes = ant[ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
            entrou = atual[~atual["PedidoFormatado"].isin(ant["PedidoFormatado"])]

            st.image(pizza(len(tratados), len(restantes), "Status Específico 2x"))

            st.markdown(
                f'<p class="metric-small">Tratados: {len(tratados)} / {len(ant)}</p>',
                unsafe_allow_html=True
            )
            st.markdown(
                f'<p class="metric-small">Entraram: {len(entrou)}</p>',
                unsafe_allow_html=True
            )

            # 🔥 REMANESCENTE MATEMÁTICO (2x + 1)
            if (
                "DiasDesdeUltimoStatus" in atual.columns and
                "Prazo_Status_Especifico" in atual.columns
            ):
                limite = (atual["Prazo_Status_Especifico"] * 2) + 1
                remanescente_status = atual[
                    atual["DiasDesdeUltimoStatus"] > limite
                ].copy()
            else:
                remanescente_status = restantes.copy()

            buf = BytesIO()
            remanescente_status.to_excel(buf, index=False)
            st.download_button(
                "Remanescentes Status 2x",
                buf.getvalue(),
                file_name=f"remanescente_status_{dia_atual}.xlsx"
            )

    # ================= REGIÃO 2X =================
    with col3:
        if "Regiao_Dobro" in df_atual.columns:

            atual = df_atual[df_atual["Regiao_Dobro"]=="X"]
            ant = df_ant[df_ant["Regiao_Dobro"]=="X"]

            tratados = ant[~ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
            restantes = ant[ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
            entrou = atual[~atual["PedidoFormatado"].isin(ant["PedidoFormatado"])]

            st.image(pizza(len(tratados), len(restantes), "Região 2x Prazo"))

            st.markdown(
                f'<p class="metric-small">Tratados: {len(tratados)} / {len(ant)}</p>',
                unsafe_allow_html=True
            )
            st.markdown(
                f'<p class="metric-small">Entraram: {len(entrou)}</p>',
                unsafe_allow_html=True
            )

            # 🔥 REMANESCENTE MATEMÁTICO (2x + 1)
            if (
                "DiasDesdeUltimoStatus" in atual.columns and
                "Prazo_Regiao" in atual.columns
            ):
                limite = (atual["Prazo_Regiao"] * 2) + 1
                remanescente_regiao = atual[
                    atual["DiasDesdeUltimoStatus"] > limite
                ].copy()
            else:
                remanescente_regiao = restantes.copy()

            buf = BytesIO()
            remanescente_regiao.to_excel(buf, index=False)
            st.download_button(
                "Remanescentes Região 2x",
                buf.getvalue(),
                file_name=f"remanescente_regiao_{dia_atual}.xlsx"
            )

    st.divider()
