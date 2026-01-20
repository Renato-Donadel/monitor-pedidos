import streamlit as st
import pandas as pd
import os
from io import BytesIO
import matplotlib.pyplot as plt

# ==============================
# CONFIGS
# ==============================
PASTA_DATA = os.path.join(os.path.dirname(__file__), "data")

ARQ_ATUAL = os.path.join(PASTA_DATA, "Monitor_Pedidos_Processado.xlsx")
ARQ_MANHA = os.path.join(PASTA_DATA, "Monitor_Pedidos_Processado_manha.xlsx")
ARQ_TARDE = os.path.join(PASTA_DATA, "Monitor_Pedidos_Processado_tarde.xlsx")

TAMANHO_LOTE = 300

st.set_page_config(page_title="Monitor de Pedidos Críticos", layout="wide")


# ==============================
# LOGIN (1x por sessão)
# ==============================
SENHA_APP = "SUA_SENHA_AQUI"

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
        st.warning(f"⚠️ Arquivo não encontrado: **{nome}**")
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
    Se filtro_manha for passado, ele filtra o df_manha antes de comparar.
    """
    if df_manha.empty or df_tarde.empty:
        return None

    chave = "PedidoFormatado"

    # garante colunas básicas
    for c in [chave, "Status"]:
        if c not in df_manha.columns or c not in df_tarde.columns:
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


def pizza_tratados(titulo: str, total: int, tratados: int, nao_tratados: int, tamanho=1.0):
    fig, ax = plt.subplots(figsize=(4.0 * tamanho, 4.0 * tamanho))
    ax.pie(
        [tratados, nao_tratados],
        labels=["Tratados", "Não tratados"],
        autopct="%1.0f%%",
        startangle=90,
    )
    ax.set_title(f"{titulo}\nTotal: {total} | Tratados: {tratados}")
    st.pyplot(fig)


# ==============================
# CARREGAR BASES
# ==============================
df_atual = ler_base(ARQ_ATUAL, "Monitor atual (botões)")
df_manha = ler_base(ARQ_MANHA, "Monitor manhã (gráficos)")
df_tarde = ler_base(ARQ_TARDE, "Monitor tarde (gráficos)")


# ==============================
# TÍTULO
# ==============================
st.title("📦 Monitor de Pedidos Críticos")


# ==============================
# TOPO: BOTÕES (HORIZONTAL)
# ==============================
st.subheader("📥 Carteiras — Download (base atual)")

if df_atual.empty:
    st.error("Base atual vazia ou não carregada. Verifique o arquivo Monitor_Pedidos_Processado.xlsx.")
    st.stop()

if "Ranking" in df_atual.columns:
    df_atual = df_atual.sort_values("Ranking").reset_index(drop=True)

if "offsets" not in st.session_state:
    st.session_state["offsets"] = {}

carteiras = sorted(df_atual["Carteira"].dropna().unique())

COLS_POR_LINHA = 4
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

            if st.button("📥 Baixar", key=f"baixar_{carteira}"):
                df_lote = df_carteira.iloc[inicio:fim]

                if df_lote.empty:
                    st.warning("✅ Já chegou no fim dessa carteira.")
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
# BI (EMBAIXO)
# ==============================
st.subheader("📊 BI — Tratados do dia (manhã x tarde)")
st.caption(f"📅 Data: **{pd.Timestamp.today().strftime('%d/%m/%Y')}**")

if df_manha.empty or df_tarde.empty:
    st.info("Os gráficos aparecem quando existirem os arquivos de **manhã** e **tarde** na pasta `data/`.")
    st.stop()

if "DescricaoCriticidade" not in df_manha.columns:
    st.warning("⚠️ Sua base da manhã não tem a coluna `DescricaoCriticidade`. Alguns gráficos podem não funcionar.")

with st.container(border=True):
    # 1) GERAL (em cima)
    r = calcular_tratados(df_manha, df_tarde)
    if r is None:
        st.error("Não foi possível calcular o gráfico geral (verifique colunas PedidoFormatado e Status).")
    else:
        total, tratados, nao_tratados = r
        pizza_tratados("Geral — pedidos tratados", total, tratados, nao_tratados, tamanho=0.85)

    # linha 1
    c1, c2 = st.columns(2)

    # 2) Triplo prazo transportador
    with c1:
        def filtro_triplo(df):
            if "DescricaoCriticidade" not in df.columns:
                return pd.Series([False] * len(df))
            return df["DescricaoCriticidade"].fillna("").str.contains(
                "Triplo prazo transportador", case=False
            )

        t, tr, ntr = calcular_tratados(df_manha, df_tarde, filtro_manha=filtro_triplo)
        pizza_tratados("Triplo prazo transportador (manhã)", t, tr, ntr, tamanho=0.75)

    # 3) Status específico
    with c2:
        def filtro_status_especifico(df):
            if "DescricaoCriticidade" not in df.columns:
                return pd.Series([False] * len(df))
            return df["DescricaoCriticidade"].fillna("").str.contains(
                "Dobro prazo status específico", case=False
            )

        t, tr, ntr = calcular_tratados(df_manha, df_tarde, filtro_manha=filtro_status_especifico)
        pizza_tratados("Status específico (manhã)", t, tr, ntr, tamanho=0.75)

    # linha 2
    c3, c4 = st.columns(2)

    # 4) Campanha peso 3
    with c3:
        def filtro_campanha_peso3(df):
            if "PesoCampanha" not in df.columns:
                return pd.Series([False] * len(df))
            return df["PesoCampanha"].fillna(0) == 3

        t, tr, ntr = calcular_tratados(df_manha, df_tarde, filtro_manha=filtro_campanha_peso3)
        pizza_tratados("Campanha prioritária (peso 3)", t, tr, ntr, tamanho=0.75)

    # 5) Por região
    with c4:
        def filtro_regiao(df):
            if "DescricaoCriticidade" not in df.columns:
                return pd.Series([False] * len(df))
            return df["DescricaoCriticidade"].fillna("").str.contains(
                "Dobro prazo status por região", case=False
            )

        t, tr, ntr = calcular_tratados(df_manha, df_tarde, filtro_manha=filtro_regiao)
        pizza_tratados("Fora do prazo por região (manhã)", t, tr, ntr, tamanho=0.75)
