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
        st.warning(f"⚠️ Arquivo não encontrado: **{nome}**")
        return pd.DataFrame()
    try:
        return pd.read_excel(path)
    except Exception as e:
        st.error(f"Erro ao ler **{nome}**: {e}")
        return pd.DataFrame()

def merge_manha_tarde(df_manha: pd.DataFrame, df_tarde: pd.DataFrame) -> pd.DataFrame:
    """
    Retorna tabela com Status_manha e Status_tarde por PedidoFormatado
    """
    chave = "PedidoFormatado"
    cols = [chave, "Status"]

    df_m = df_manha[cols].copy()
    df_t = df_tarde[cols].copy()

    m = df_m.merge(df_t, on=chave, how="left", suffixes=("_manha", "_tarde"))
    return m

def calcular_tratados(df_manha: pd.DataFrame, df_tarde: pd.DataFrame, filtro_manha=None):
    """
    Tratado = pedido que estava de manhã e:
      - sumiu na tarde
      OU
      - continua, mas Status mudou
    Se filtro_manha for passado, ele filtra o df_manha antes de comparar.
    """
    if df_manha.empty or df_tarde.empty:
        return None

    if filtro_manha is not None:
        df_manha = df_manha[filtro_manha(df_manha)].copy()

    if df_manha.empty:
        return 0, 0, 0

    chave = "PedidoFormatado"

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
    fig, ax = plt.subplots(figsize=(4.2 * tamanho, 4.2 * tamanho))
    ax.pie(
        [tratados, nao_tratados],
        labels=["Tratados", "Não tratados"],
        autopct="%1.0f%%",
        startangle=90,
    )
    ax.set_title(f"{titulo}\nTotal: {total} | Tratados: {tratados}")
    st.pyplot(fig, use_container_width=True)

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
# TOPO: BOTÕES (HORIZONTAL E MENOR)
# ==============================
st.subheader("📥 Carteiras — Download (base atual)")

if df_atual.empty:
    st.error("Base atual vazia ou não carregada. Verifique o arquivo Monitor_Pedidos_Processado.xlsx.")
    st.stop()

# ordena sempre
if "Ranking" in df_atual.columns:
    df_atual = df_atual.sort_values("Ranking").reset_index(drop=True)

if "offsets" not in st.session_state:
    st.session_state["offsets"] = {}

carteiras = sorted(df_atual["Carteira"].dropna().unique())

# cria botões em linhas com 4 colunas
COLS_POR_LINHA = 4
linhas = [carteiras[i:i + COLS_POR_LINHA] for i in range(0, len(carteiras), COLS_POR_LINHA)]

for grupo in linhas:
    cols = st.columns(COLS_POR_LINHA)
    for idx, carteira in enumerate(grupo):
        with cols[idx]:
            df_carteira = df_atual[df_atual["Carteira"] == carteira].reset_index(drop=True)
            total = len(df_carteira)

            offset_atual = st.session_state["offsets"].get(carteira, 0)
            inicio = offset_atual
            fim = min(offset_atual + TAMANHO_LOTE, total)

            st.caption(f"**{carteira}**")
            st.caption(f"{inicio+1}–{fim} / {total}")

            if st.button(f"📥 Baixar", key=f"baixar_{carteira}"):
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
# ABAIXO: QUADRO COM 5 GRÁFICOS
# ==============================
st.subheader("📊 BI — Tratados do dia (manhã x tarde)")

data_ref = None
if not df_tarde.empty and "DataPedido" in df_tarde.columns:
    # só pra tentar mostrar uma referência
    data_ref = pd.Timestamp.today().strftime("%d/%m/%Y")
else:
    data_ref = pd.Timestamp.today().strftime("%d/%m/%Y")

st.caption(f"📅 Data: **{data_ref}**")

if df_manha.empty or df_tarde.empty:
    st.info("Os gráficos aparecem quando existirem os arquivos de **manhã** e **tarde** na pasta `data/`.")
    st.stop()

with st.container(border=True):
    # 1) GERAL (GRANDE EM CIMA)
    total, tratados, nao_tratados = calcular_tratados(df_manha, df_tarde)
    pizza_tratados("Geral — pedidos tratados", total, tratados, nao_tratados, tamanho=1.25)

    # 2x2 embaixo + 1 extra (ficará 2 linhas)
    c1, c2 = st.columns(2)

    # 2) Triplo prazo transportador (manhã)
    with c1:
        def filtro_triplo(df):
            return df.get("TriploPrazoTransportador", False) == True
        t, tr, ntr = calcular_tratados(df_manha, df_tarde, filtro_manha=filtro_triplo)
        pizza_tratados("Triplo prazo transportador (manhã)", t, tr, ntr, tamanho=0.95)

    # 3) Status específico (manhã)
    with c2:
        def filtro_status_especifico(df):
            return df.get("DobroPrazoStatusEspecifico", False) == True
        t, tr, ntr = calcular_tratados(df_manha, df_tarde, filtro_manha=filtro_status_especifico)
        pizza_tratados("Status específico (manhã)", t, tr, ntr, tamanho=0.95)

    c3, c4 = st.columns(2)

    # 4) Campanha peso 3 (manhã)
    with c3:
        def filtro_campanha_peso3(df):
            return df.get("PesoCampanha", 0) == 3
        t, tr, ntr = calcular_tratados(df_manha, df_tarde, filtro_manha=filtro_campanha_peso3)
        pizza_tratados("Campanha prioritária (peso 3)", t, tr, ntr, tamanho=0.95)

    # 5) Fora do prazo por região (manhã)
    with c4:
        def filtro_regiao(df):
            return df.get("DobroPrazoStatusRegiao", False) == True
        t, tr, ntr = calcular_tratados(df_manha, df_tarde, filtro_manha=filtro_regiao)
        pizza_tratados("Fora do prazo por região (manhã)", t, tr, ntr, tamanho=0.95)
