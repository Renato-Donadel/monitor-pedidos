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
PASTA_HIST_SITE = os.path.join(PASTA_DATA, "historico")

ARQ_ATUAL = os.path.join(PASTA_DATA, "Monitor_Pedidos_Processado.xlsx")
TAMANHO_LOTE = 300

st.set_page_config(page_title="Monitor de Pedidos Críticos", layout="wide")


# ==============================
# LOGIN (1x por sessão)
# ==============================
SENHA_APP = "8S15?w5fkP"

if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    st.title("🔒 Acesso restrito")
    senha = st.text_input("Digite a senha para acessar", type="password")

    if st.button("Entrar"):
        if senha == SENHA_APP:
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Senha incorreta.")
    st.stop()


# ==============================
# FUNÇÕES
# ==============================
def ler_base(path: str, nome: str) -> pd.DataFrame:
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        return pd.read_excel(path)
    except Exception as e:
        st.error(f"Erro ao ler **{nome}**: {e}")
        return pd.DataFrame()


def calcular_tratados(df_manha: pd.DataFrame, df_tarde: pd.DataFrame, filtro_manha=None):
    """
    TRATADO = pedido que estava de manhã e:
      - sumiu na tarde
      OU
      - continua, mas Status mudou
    """
    if df_manha.empty or df_tarde.empty:
        return None

    chave = "PedidoFormatado"
    if chave not in df_manha.columns or chave not in df_tarde.columns:
        return None
    if "Status" not in df_manha.columns or "Status" not in df_tarde.columns:
        return None

    if filtro_manha is not None:
        try:
            mask = filtro_manha(df_manha)
            df_manha = df_manha[mask].copy()
        except Exception:
            return None

    if df_manha.empty:
        return 0, 0, 0

    df_m = df_manha[[chave, "Status"]].copy()
    df_t = df_tarde[[chave, "Status"]].copy()

    m = df_m.merge(df_t, on=chave, how="left", suffixes=("_manha", "_tarde"))

    sumiu = m["Status_tarde"].isna()
    mudou = (~sumiu) & (m["Status_manha"] != m["Status_tarde"])
    tratados = sumiu | mudou

    total = len(m)
    qtd_tratados = int(tratados.sum())
    qtd_nao_tratados = total - qtd_tratados

    return total, qtd_tratados, qtd_nao_tratados


def pizza_tratados(titulo: str, total: int, tratados: int, nao_tratados: int):
    """
    Pizza MINIATURA e nítida (sem esticar no Streamlit)
    """
    if total == 0:
        st.caption(f"{titulo}: 0")
        return

    fig, ax = plt.subplots(figsize=(0.85, 0.85), dpi=200)

    # sem labels grandes
    ax.pie(
        [tratados, nao_tratados],
        labels=None,
        autopct=lambda p: f"{p:.0f}%",
        startangle=90,
        textprops={"fontsize": 5},
    )

    ax.set_title(titulo, fontsize=6, pad=2)

    # deixa bem "quadradinho" e sem margens
    ax.axis("equal")
    plt.tight_layout(pad=0.2)

    st.pyplot(fig)  # NÃO usar container_width


def listar_dias_historico():
    """
    Espera arquivos:
      data/historico/DD-MM-YYYY_manha.xlsx
      data/historico/DD-MM-YYYY_tarde.xlsx
    """
    if not os.path.exists(PASTA_HIST_SITE):
        return []

    arquivos = os.listdir(PASTA_HIST_SITE)

    datas = set()
    for a in arquivos:
        m = re.match(r"(\d{2}-\d{2}-\d{4})_(manha|tarde)\.xlsx$", a, flags=re.IGNORECASE)
        if m:
            datas.add(m.group(1))

    def chave_data(s):
        try:
            return pd.to_datetime(s, format="%d-%m-%Y")
        except Exception:
            return pd.Timestamp.min

    return sorted(list(datas), key=chave_data)


def caminho_hist(dia: str, periodo: str) -> str:
    return os.path.join(PASTA_HIST_SITE, f"{dia}_{periodo}.xlsx")


# ==============================
# TÍTULO
# ==============================
st.title("📦 Monitor de Pedidos Críticos")


# ==============================
# TOPO: BOTÕES (DOWNLOAD)
# ==============================
st.subheader("📥 Carteiras — Download (base atual)")

df_atual = ler_base(ARQ_ATUAL, "Monitor atual")

if df_atual.empty:
    st.error("Base atual não encontrada ou vazia (data/Monitor_Pedidos_Processado.xlsx).")
    st.stop()

if "Ranking" in df_atual.columns:
    df_atual = df_atual.sort_values("Ranking").reset_index(drop=True)

if "offsets" not in st.session_state:
    st.session_state["offsets"] = {}

carteiras = sorted(df_atual["Carteira"].dropna().unique())

COLS_POR_LINHA = 5
linhas = [carteiras[i:i + COLS_POR_LINHA] for i in range(0, len(carteiras), COLS_POR_LINHA)]

for grupo in linhas:
    cols = st.columns(COLS_POR_LINHA)
    for idx, carteira in enumerate(grupo):
        with cols[idx]:
            df_carteira = df_atual[df_atual["Carteira"] == carteira].reset_index(drop=True)
            total_carteira = len(df_carteira)

            offset_atual = st.session_state["offsets"].get(carteira, 0)
            inicio = offset_atual
            fim = min(offset_atual + TAMANHO_LOTE, total_carteira)

            st.caption(f"**{carteira}**")
            st.caption(f"{inicio+1}–{fim} / {total_carteira}")

            if st.button("📥", key=f"baixar_{carteira}"):
                df_lote = df_carteira.iloc[inicio:fim]

                if df_lote.empty:
                    st.warning("Fim da carteira.")
                    st.stop()

                st.session_state["offsets"][carteira] = fim

                nome_arquivo = f"Pedidos_{carteira}_{inicio+1}_a_{fim}.xlsx"

                buffer = BytesIO()
                df_lote.to_excel(buffer, index=False)
                buffer.seek(0)

                st.download_button(
                    label="⬇️ Excel",
                    data=buffer,
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{carteira}",
                )

st.divider()


# ==============================
# HISTÓRICO BI (dias lado a lado)
# ==============================
st.subheader("📊 BI — Histórico de Tratados (manhã x tarde)")
st.caption("Cada dia adiciona um novo bloco de 5 pizzas ao lado (histórico).")

dias = listar_dias_historico()

if not dias:
    st.info("Sem histórico ainda (pasta data/historico vazia).")
    st.stop()

ULTIMOS = 10
dias_exibir = dias[-ULTIMOS:]

cols_dias = st.columns(len(dias_exibir))

for i, dia in enumerate(dias_exibir):
    with cols_dias[i]:
        st.markdown(f"### 📅 {dia}")

        df_manha = ler_base(caminho_hist(dia, "manha"), f"{dia}_manha")
        df_tarde = ler_base(caminho_hist(dia, "tarde"), f"{dia}_tarde")

        if df_manha.empty or df_tarde.empty:
            st.warning("Sem manhã ou tarde")
            continue

        # 1) Geral
        r = calcular_tratados(df_manha, df_tarde)
        if r is None:
            st.warning("Erro geral")
            continue
        total, tratados, nao_tratados = r
        pizza_tratados(f"Geral\nT:{tratados}/{total}", total, tratados, nao_tratados)

        # 2) Triplo transportador (via texto)
        def filtro_triplo(df):
            if "DescricaoCriticidade" not in df.columns:
                return pd.Series([False] * len(df))
            return df["DescricaoCriticidade"].fillna("").str.contains("Triplo prazo transportador", case=False)

        t, tr, ntr = calcular_tratados(df_manha, df_tarde, filtro_manha=filtro_triplo)
        pizza_tratados(f"Triplo\nT:{tr}/{t}", t, tr, ntr)

        # 3) Status específico (via texto)
        def filtro_especifico(df):
            if "DescricaoCriticidade" not in df.columns:
                return pd.Series([False] * len(df))
            return df["DescricaoCriticidade"].fillna("").str.contains("Dobro prazo status específico", case=False)

        t, tr, ntr = calcular_tratados(df_manha, df_tarde, filtro_manha=filtro_especifico)
        pizza_tratados(f"Espec.\nT:{tr}/{t}", t, tr, ntr)

        # 4) Campanha peso 3
        def filtro_peso3(df):
            if "PesoCampanha" not in df.columns:
                return pd.Series([False] * len(df))
            return df["PesoCampanha"].fillna(0) == 3

        t, tr, ntr = calcular_tratados(df_manha, df_tarde, filtro_manha=filtro_peso3)
        pizza_tratados(f"Camp 3\nT:{tr}/{t}", t, tr, ntr)

        # 5) Região (via texto)
        def filtro_regiao(df):
            if "DescricaoCriticidade" not in df.columns:
                return pd.Series([False] * len(df))
            return df["DescricaoCriticidade"].fillna("").str.contains("Dobro prazo status por região", case=False)

        t, tr, ntr = calcular_tratados(df_manha, df_tarde, filtro_manha=filtro_regiao)
        pizza_tratados(f"Região\nT:{tr}/{t}", t, tr, ntr)
