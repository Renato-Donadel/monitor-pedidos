import streamlit as st
import pandas as pd
import os
from io import BytesIO

# ==============================
# CONFIGS
# ==============================
PASTA_DATA = os.path.join(os.path.dirname(__file__), "data")

ARQ_ATUAL = os.path.join(PASTA_DATA, "Monitor_Pedidos_Processado.xlsx")
ARQ_MANHA = os.path.join(PASTA_DATA, "Monitor_Pedidos_Processado_manha.xlsx")
ARQ_TARDE = os.path.join(PASTA_DATA, "Monitor_Pedidos_Processado_tarde.xlsx")

TAMANHO_LOTE = 300

st.set_page_config(page_title="Monitor de Pedidos Críticos", layout="centered")


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
        df = pd.read_excel(path)
        return df
    except Exception as e:
        st.error(f"Erro ao ler **{nome}**: {e}")
        return pd.DataFrame()


def calcular_tratados(df_manha: pd.DataFrame, df_tarde: pd.DataFrame):
    """
    Tratado = pedido que estava de manhã e:
      - sumiu da base da tarde
      OU
      - continua, mas o Status mudou
    """
    if df_manha.empty or df_tarde.empty:
        return None

    chave = "PedidoFormatado"

    df_m = df_manha[[chave, "Status"]].copy()
    df_t = df_tarde[[chave, "Status"]].copy()

    df_merge = df_m.merge(df_t, on=chave, how="left", suffixes=("_manha", "_tarde"))

    # sumiu
    sumiu = df_merge["Status_tarde"].isna()

    # mudou status (só faz sentido se não sumiu)
    mudou_status = (~sumiu) & (df_merge["Status_manha"] != df_merge["Status_tarde"])

    tratados = sumiu | mudou_status

    total = len(df_merge)
    qtd_tratados = int(tratados.sum())
    qtd_nao_tratados = total - qtd_tratados

    return total, qtd_tratados, qtd_nao_tratados


def pizza_tratados(titulo: str, total: int, tratados: int, nao_tratados: int):
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots()
    ax.pie(
        [tratados, nao_tratados],
        labels=["Tratados", "Não tratados"],
        autopct="%1.0f%%",
        startangle=90,
    )
    ax.set_title(f"{titulo}\nTotal: {total} | Tratados: {tratados}")
    st.pyplot(fig)


# ==============================
# LEITURA BASES
# ==============================
df_atual = ler_base(ARQ_ATUAL, "Monitor atual")
df_manha = ler_base(ARQ_MANHA, "Monitor manhã")
df_tarde = ler_base(ARQ_TARDE, "Monitor tarde")


# ==============================
# TÍTULO
# ==============================
st.title("📦 Monitor de Pedidos Críticos")

st.caption("✅ Botões usam a base **atual** | 📊 Gráficos comparam **manhã x tarde**")

st.divider()


# ==============================
# GRÁFICOS (BI)
# ==============================
st.subheader("📊 BI — Tratados no dia (manhã x tarde)")

resultado = calcular_tratados(df_manha, df_tarde)

if resultado is None:
    st.info("Os gráficos aparecem quando existirem os arquivos de **manhã** e **tarde** na pasta `data/`.")
else:
    total, qtd_tratados, qtd_nao_tratados = resultado
    pizza_tratados("Pedidos tratados", total, qtd_tratados, qtd_nao_tratados)

st.divider()


# ==============================
# BOTÕES (DOWNLOAD)
# ==============================
st.subheader("📥 Download por carteira (base atual)")

if df_atual.empty:
    st.error("Base atual vazia ou não carregada. Verifique o arquivo Monitor_Pedidos_Processado.xlsx.")
    st.stop()

# ordena pelo ranking sempre
if "Ranking" in df_atual.columns:
    df_atual = df_atual.sort_values("Ranking").reset_index(drop=True)

# offsets por carteira
if "offsets" not in st.session_state:
    st.session_state["offsets"] = {}

carteiras = sorted(df_atual["Carteira"].dropna().unique())

for carteira in carteiras:
    st.subheader(f"📁 {carteira}")

    df_carteira = df_atual[df_atual["Carteira"] == carteira].reset_index(drop=True)
    total = len(df_carteira)

    if total == 0:
        st.caption("Sem pedidos nessa carteira.")
        st.divider()
        continue

    offset_atual = st.session_state["offsets"].get(carteira, 0)
    inicio = offset_atual
    fim = min(offset_atual + TAMANHO_LOTE, total)

    st.caption(f"Total: {total} | Próximo lote: {inicio+1} até {fim}")

    if st.button(f"Baixar próximos 300 — {carteira}", key=f"baixar_{carteira}"):
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
            label="📥 Clique aqui para baixar o Excel",
            data=buffer,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.divider()
