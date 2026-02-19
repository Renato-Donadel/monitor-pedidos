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
LOGO_PATH = os.path.join(PASTA_DATA, "logo_bravium.png")  # NOME CORRETO DO ARQUIVO

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

/* Títulos */
h1, h2, h3 {
    color: #0f2a44;
    font-weight: 700;
}

/* Cards de métricas */
div[data-testid="metric-container"] {
    background: white;
    border-radius: 14px;
    padding: 18px;
    box-shadow: 0px 6px 18px rgba(0,0,0,0.06);
    border: 1px solid #e6e9ef;
}

/* Botões principais (exportação das meninas) */
.stDownloadButton > button {
    background: linear-gradient(90deg, #0f2a44, #1f4e79);
    color: white;
    border-radius: 12px;
    font-weight: 700;
    height: 48px;
    width: 100%;
    border: none;
    font-size: 16px;
}

.stDownloadButton > button:hover {
    background: linear-gradient(90deg, #1f4e79, #2d6aa3);
    transform: scale(1.02);
}

/* Botões comuns */
.stButton > button {
    background: linear-gradient(90deg, #0f2a44, #1f4e79);
    color: white;
    border-radius: 10px;
    font-weight: 600;
    border: none;
    height: 42px;
    width: 100%;
}

.stButton > button:hover {
    background: linear-gradient(90deg, #1f4e79, #2d6aa3);
}

/* Sidebar estilo corporativo */
section[data-testid="stSidebar"] {
    background-color: #0f2a44;
}

/* Divisores */
hr {
    border: 1px solid #e6e9ef;
}
</style>
""", unsafe_allow_html=True)

# ==============================
# HEADER CORPORATIVO COM LOGO
# ==============================
col_logo, col_title = st.columns([1, 6])

with col_logo:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=140)

with col_title:
    st.markdown("""
    <div style="
        background: linear-gradient(90deg, #0f2a44, #1f4e79);
        padding: 22px;
        border-radius: 16px;
        color: white;
        box-shadow: 0px 8px 24px rgba(0,0,0,0.08);
    ">
        <h2 style="margin:0; font-size:28px;">
        Monitor de Pedidos — BI Executivo
        </h2>
        <p style="margin:6px 0 0 0; opacity:0.85; font-size:15px;">
        Análise de Risco Logístico • Transportadora • Status • Região • Cliente
        </p>
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

    # Normalização crítica (bug de espaço resolvido)
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
        if pd.to_datetime(d, format="%d-%m-%Y") <= alvo
    ]
    if not candidatos:
        return None
    return max(candidatos, key=lambda d: pd.to_datetime(d, format="%d-%m-%Y"))


def pizza(tratados, nao, titulo):
    fig, ax = plt.subplots(figsize=(3, 3))
    total = tratados + nao
    if total == 0:
        ax.text(0.5, 0.5, "0", ha="center", va="center")
    else:
        ax.pie([tratados, nao], autopct="%1.0f%%", startangle=90)
    ax.set_title(titulo)
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()

# ==============================
# 📥 CARTEIRAS (DOWNLOAD DIRETO)
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
                    label=f"⬇️ Baixar {carteira}",
                    data=buffer,
                    file_name=f"{carteira}_{inicio+1}_a_{fim}.xlsx",
                    key=f"download_{carteira}_{offset}"
                ):
                    st.session_state["offsets"][carteira] = fim

    st.divider()

# =====================================================
# 📊 PARTE 2 — BI EXECUTIVO (LAYOUT EXECUTIVO COMPACTO)
# =====================================================
dias = listar_dias()

if len(dias) < 2:
    st.warning("Histórico insuficiente na pasta data/historico.")
    st.stop()

dias = dias[-15:]

tend_triplo = []
datas_plot = []

st.markdown("## 📊 BI Executivo — Comparação Dia a Dia")

for i in range(len(dias) - 1, 0, -1):

    dia_atual = dias[i]
    dia_ant = dias[i - 1]

    df_atual = ler_base(caminho(dia_atual))
    df_ant = ler_base(caminho(dia_ant))

    if df_atual.empty or df_ant.empty:
        continue

    # ==============================
    # DATA (MENOR E MAIS ELEGANTE)
    # ==============================
    st.markdown(
        f"<h3 style='font-size:22px; margin-bottom:10px;'>📅 {dia_ant} ➜ {dia_atual}</h3>",
        unsafe_allow_html=True
    )

    # ==============================
    # BASES
    # ==============================
    triplo_atual = df_atual[df_atual.get("Transportadora_Triplo", "") == "X"]
    triplo_ant = df_ant[df_ant.get("Transportadora_Triplo", "") == "X"]

    status_atual = df_atual[df_atual.get("Status_Dobro", "") == "X"]
    status_ant = df_ant[df_ant.get("Status_Dobro", "") == "X"]

    reg_atual = df_atual[df_atual.get("Regiao_Dobro", "") == "X"]
    reg_ant = df_ant[df_ant.get("Regiao_Dobro", "") == "X"]

    # ==============================
    # CÁLCULOS
    # ==============================
    # TRIPLO
    tratados_triplo = triplo_ant[
        ~triplo_ant["PedidoFormatado"].isin(triplo_atual["PedidoFormatado"])
    ]
    persist_triplo_d1 = triplo_ant[
        triplo_ant["PedidoFormatado"].isin(triplo_atual["PedidoFormatado"])
    ]
    entrou_triplo = triplo_atual[
        ~triplo_atual["PedidoFormatado"].isin(triplo_ant["PedidoFormatado"])
    ]

    # 72H CORRETO
    dia_72 = encontrar_dia_72h(dias, dia_atual)
    if dia_72:
        df_72 = ler_base(caminho(dia_72))
        triplo_72 = df_72[df_72.get("Transportadora_Triplo", "") == "X"]
        persist_72h = triplo_atual[
            triplo_atual["PedidoFormatado"].isin(triplo_72["PedidoFormatado"])
        ]
    else:
        persist_72h = pd.DataFrame()

    # STATUS 2X
    tratados_status = status_ant[
        ~status_ant["PedidoFormatado"].isin(status_atual["PedidoFormatado"])
    ]
    persist_status = status_ant[
        status_ant["PedidoFormatado"].isin(status_atual["PedidoFormatado"])
    ]
    entrou_status = status_atual[
        ~status_atual["PedidoFormatado"].isin(status_ant["PedidoFormatado"])
    ]

    # REGIÃO 2X
    tratados_reg = reg_ant[
        ~reg_ant["PedidoFormatado"].isin(reg_atual["PedidoFormatado"])
    ]
    persist_reg = reg_ant[
        reg_ant["PedidoFormatado"].isin(reg_atual["PedidoFormatado"])
    ]
    entrou_reg = reg_atual[
        ~reg_atual["PedidoFormatado"].isin(reg_ant["PedidoFormatado"])
    ]

    # ==============================
    # LAYOUT HORIZONTAL (3 COLUNAS)
    # ==============================
    col_triplo, col_status, col_reg = st.columns(3)

    # ==============================
    # 🔴 COLUNA 1 — TRIPLO TRANSPORTADORA
    # ==============================
    with col_triplo:
        st.markdown("### 🔴 Triplo Transportadora")

        st.image(
            pizza(len(tratados_triplo), len(persist_triplo_d1),
                   f"Tratados\n{len(tratados_triplo)}/{len(triplo_ant)}")
        )

        st.markdown(
            f"<p style='font-size:18px; margin-top:5px;'><b>Entraram:</b> {len(entrou_triplo)}</p>",
            unsafe_allow_html=True
        )

        buffer = BytesIO()
        persist_72h.to_excel(buffer, index=False)
        st.download_button(
            "📥 Exportar Triplo > 72h",
            buffer.getvalue(),
            file_name=f"triplo_72h_{dia_atual}.xlsx",
            use_container_width=True
        )

        tend_triplo.append(len(triplo_atual))
        datas_plot.append(dia_atual)

    # ==============================
    # 🟡 COLUNA 2 — STATUS 2x PRAZO
    # ==============================
    with col_status:
        st.markdown("### 🟡 Status 2x Prazo")

        st.image(
            pizza(len(tratados_status), len(persist_status),
                   f"Tratados\n{len(tratados_status)}/{len(status_ant)}")
        )

        st.markdown(
            f"<p style='font-size:18px; margin-top:5px;'><b>Entraram:</b> {len(entrou_status)}</p>",
            unsafe_allow_html=True
        )

        buffer = BytesIO()
        persist_status.to_excel(buffer, index=False)
        st.download_button(
            "📥 Exportar Status Persistente",
            buffer.getvalue(),
            file_name=f"status_2x_{dia_atual}.xlsx",
            use_container_width=True
        )

    # ==============================
    # 🔵 COLUNA 3 — REGIÃO 2x PRAZO
    # ==============================
    with col_reg:
        st.markdown("### 🔵 Região 2x Prazo")

        st.image(
            pizza(len(tratados_reg), len(persist_reg),
                   f"Tratados\n{len(tratados_reg)}/{len(reg_ant)}")
        )

        st.markdown(
            f"<p style='font-size:18px; margin-top:5px;'><b>Entraram:</b> {len(entrou_reg)}</p>",
            unsafe_allow_html=True
        )

        buffer = BytesIO()
        persist_reg.to_excel(buffer, index=False)
        st.download_button(
            "📥 Exportar Região Persistente",
            buffer.getvalue(),
            file_name=f"regiao_2x_{dia_atual}.xlsx",
            use_container_width=True
        )

    st.divider()


# ==============================
# 📈 TENDÊNCIA
# ==============================
if tend_triplo:
    st.markdown("## 📈 Tendência de Pedidos em Triplo")

    fig, ax = plt.subplots()
    ax.plot(datas_plot[::-1], tend_triplo[::-1], marker="o")
    ax.set_title("Evolução do Triplo ao Longo do Tempo")
    ax.tick_params(axis='x', rotation=45)
    st.pyplot(fig)
