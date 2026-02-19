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
# 🎨 ESTILO BRAVIUM
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
    font-size: 18px;
    font-weight: 700;
    color: #0f2a44;
    margin-top: 6px;
    margin-bottom: 6px;
}

.metric-small {
    font-size: 15px;
    font-weight: 600;
    color: #0f2a44;
    margin: 2px 0;
}

.stDownloadButton > button {
    background: linear-gradient(90deg, #0f2a44, #1f4e79);
    color: white;
    border-radius: 10px;
    font-weight: 700;
    height: 38px;
    width: 100%;
    border: none;
}
</style>
""", unsafe_allow_html=True)

# ==============================
# HEADER COM LOGO NA FAIXA AZUL
# ==============================
logo_html = ""
if os.path.exists(LOGO_PATH):
    with open(LOGO_PATH, "rb") as f:
        logo_base64 = base64.b64encode(f.read()).decode()
    logo_html = f'<img src="data:image/png;base64,{logo_base64}" width="130">'

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
    fig, ax = plt.subplots(figsize=(2.2, 2.2))
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
# 📥 DOWNLOAD POR CARTEIRA (SEQUENCIAL CORRETO)
# ==============================
st.markdown("### 📥 Exportação por Carteira (300 em 300)")

df_atual_base = ler_base(ARQ_ATUAL)

if "offsets_carteira" not in st.session_state:
    st.session_state["offsets_carteira"] = {}

if not df_atual_base.empty and "Carteira" in df_atual_base.columns:

    if "Ranking" in df_atual_base.columns:
        df_atual_base = df_atual_base.sort_values("Ranking").reset_index(drop=True)

    carteiras = sorted(df_atual_base["Carteira"].dropna().unique())

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
                clicou = st.download_button(
                    label=f"⬇️ Baixar {carteira}",
                    data=buffer,
                    file_name=f"{carteira}_{inicio+1}_a_{fim}.xlsx",
                    key=f"dl_{carteira}_{offset}"
                )

                if clicou:
                    st.session_state["offsets_carteira"][carteira] = fim

st.divider()

# ==============================
# 📊 BI EXECUTIVO (REMNESCENTE MATEMÁTICO)
# ==============================
dias = listar_dias()

if len(dias) < 1:
    st.warning("Nenhuma base encontrada na pasta data/historico.")
    st.stop()

# usa a base mais recente para cálculo matemático
dia_atual = dias[-1]
df_atual = ler_base(caminho(dia_atual))

if df_atual.empty:
    st.warning("Base mais recente vazia.")
    st.stop()

st.markdown(
    f'<p class="data-title">📅 Base Atual: {dia_atual}</p>',
    unsafe_allow_html=True
)

col1, col2, col3 = st.columns(3)

# ================= TRIPLO TRANSPORTADORA =================
with col1:
    if (
        "Transportadora_Triplo" in df_atual.columns and
        "DiasDesdeExpedicao" in df_atual.columns and
        "PrazoTransportadorDiasUteis" in df_atual.columns
    ):
        triplo = df_atual[df_atual["Transportadora_Triplo"] == "X"]

        limite_critico = (triplo["PrazoTransportadorDiasUteis"] * 3) + 3

        remanescente = triplo[
            triplo["DiasDesdeExpedicao"] > limite_critico
        ]

        tratados = len(triplo) - len(remanescente)

        st.image(pizza(tratados, len(remanescente), "Triplo Transportadora"))

        st.markdown(
            f'<p class="metric-small">Críticos: {len(remanescente)} / {len(triplo)}</p>',
            unsafe_allow_html=True
        )

        buf = BytesIO()
        remanescente.to_excel(buf, index=False)
        st.download_button(
            "Remanescente Triplo Crítico",
            buf.getvalue(),
            file_name=f"remanescente_triplo_critico_{dia_atual}.xlsx"
        )

# ================= STATUS ESPECÍFICO 2X =================
with col2:
    if (
        "Status_Dobro" in df_atual.columns and
        "DiasDesdeUltimoStatus" in df_atual.columns and
        "Prazo_Status_Especifico" in df_atual.columns
    ):
        status = df_atual[df_atual["Status_Dobro"] == "X"]

        limite_critico = (status["Prazo_Status_Especifico"] * 2) + 1

        remanescente = status[
            status["DiasDesdeUltimoStatus"] > limite_critico
        ]

        tratados = len(status) - len(remanescente)

        st.image(pizza(tratados, len(remanescente), "Status Específico 2x"))

        st.markdown(
            f'<p class="metric-small">Críticos: {len(remanescente)} / {len(status)}</p>',
            unsafe_allow_html=True
        )

        buf = BytesIO()
        remanescente.to_excel(buf, index=False)
        st.download_button(
            "Remanescente Status Crítico",
            buf.getvalue(),
            file_name=f"remanescente_status_critico_{dia_atual}.xlsx"
        )

# ================= REGIÃO 2X =================
with col3:
    if (
        "Regiao_Dobro" in df_atual.columns and
        "DiasDesdeUltimoStatus" in df_atual.columns and
        "Prazo_Regiao" in df_atual.columns
    ):
        regiao = df_atual[df_atual["Regiao_Dobro"] == "X"]

        limite_critico = (regiao["Prazo_Regiao"] * 2) + 1

        remanescente = regiao[
            regiao["DiasDesdeUltimoStatus"] > limite_critico
        ]

        tratados = len(regiao) - len(remanescente)

        st.image(pizza(tratados, len(remanescente), "Região 2x Prazo"))

        st.markdown(
            f'<p class="metric-small">Críticos: {len(remanescente)} / {len(regiao)}</p>',
            unsafe_allow_html=True
        )

        buf = BytesIO()
        remanescente.to_excel(buf, index=False)
        st.download_button(
            "Remanescente Região Crítico",
            buf.getvalue(),
            file_name=f"remanescente_regiao_critico_{dia_atual}.xlsx"
        )
