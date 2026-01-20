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
        return pd.DataFrame()
    try:
        return pd.read_excel(path)
    except Exception as e:
        st.error(f"Erro ao ler **{nome}**: {e}")
        return pd.DataFrame()


def calcular_tratados(df_manha: pd.DataFrame, df_tarde: pd.DataFrame, filtro_manha=None):
    if df_manha.empty or df_tarde.empty:
        return None

    chave = "PedidoFormatado"
    if chave not in df_manha.columns or chave not in df_tarde.columns:
        return None
    if "Status" not in df_manha.columns or "Status" not in df_tarde.columns:
        return None

    if filtro_manha is not None:
        mask = filtro_manha(df_manha)
        df_manha = df_manha[mask].copy()

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


def fig_pizza_bytes(tratados: int, nao_tratados: int, titulo: str):
    """
    Renderiza a pizza como PNG em bytes (MUUUUITO menor e nítido no Streamlit).
    """
    fig, ax = plt.subplots(figsize=(0.55, 0.55), dpi=260)

    if tratados + nao_tratados == 0:
        ax.text(0.5, 0.5, "0", ha="center", va="center", fontsize=6)
    else:
        ax.pie(
            [tratados, nao_tratados],
            labels=None,
            startangle=90,
            autopct=lambda p: f"{p:.0f}%",
            textprops={"fontsize": 4},
        )

    ax.set_title(titulo, fontsize=5, pad=1)
    ax.axis("equal")
    plt.tight_layout(pad=0.01)

    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", pad_inches=0.01)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def listar_dias_historico():
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

COLS_POR_LINHA = 6
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
# BI HISTÓRICO (datas em VERTICAL / gráficos em HORIZONTAL)
# ==============================
st.subheader("📊 BI — Histórico de Tratados (manhã x tarde)")
st.caption("Datas em lista (vertical). Para cada data: 5 pizzas na horizontal.")

dias = listar_dias_historico()

if not dias:
    st.info("Sem histórico ainda (pasta data/historico vazia).")
    st.stop()

ULTIMOS = 15
dias_exibir = dias[-ULTIMOS:]

for dia in reversed(dias_exibir):  # mais recente em cima
    df_manha = ler_base(caminho_hist(dia, "manha"), f"{dia}_manha")
    df_tarde = ler_base(caminho_hist(dia, "tarde"), f"{dia}_tarde")

    st.markdown(f"### 📅 {dia}")

    if df_manha.empty or df_tarde.empty:
        st.warning("Sem manhã ou tarde")
        st.divider()
        continue

    # filtros por texto
    def filtro_triplo(df):
        if "DescricaoCriticidade" not in df.columns:
            return pd.Series([False] * len(df))
        return df["DescricaoCriticidade"].fillna("").str.contains("Triplo prazo transportador", case=False)

    def filtro_especifico(df):
        if "DescricaoCriticidade" not in df.columns:
            return pd.Series([False] * len(df))
        return df["DescricaoCriticidade"].fillna("").str.contains("Dobro prazo status específico", case=False)

    def filtro_regiao(df):
        if "DescricaoCriticidade" not in df.columns:
            return pd.Series([False] * len(df))
        return df["DescricaoCriticidade"].fillna("").str.contains("Dobro prazo status por região", case=False)

    def filtro_peso3(df):
        if "PesoCampanha" not in df.columns:
            return pd.Series([False] * len(df))
        return df["PesoCampanha"].fillna(0) == 3

    # calcula os 5
    r1 = calcular_tratados(df_manha, df_tarde)
    r2 = calcular_tratados(df_manha, df_tarde, filtro_manha=filtro_triplo)
    r3 = calcular_tratados(df_manha, df_tarde, filtro_manha=filtro_especifico)
    r4 = calcular_tratados(df_manha, df_tarde, filtro_manha=filtro_peso3)
    r5 = calcular_tratados(df_manha, df_tarde, filtro_manha=filtro_regiao)

    # render em 5 colunas
    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        t, tr, ntr = r1
        st.image(fig_pizza_bytes(tr, ntr, f"Geral\n{tr}/{t}"))

    with c2:
        t, tr, ntr = r2
        st.image(fig_pizza_bytes(tr, ntr, f"Triplo\n{tr}/{t}"))

    with c3:
        t, tr, ntr = r3
        st.image(fig_pizza_bytes(tr, ntr, f"Esp.\n{tr}/{t}"))

    with c4:
        t, tr, ntr = r4
        st.image(fig_pizza_bytes(tr, ntr, f"Camp 3\n{tr}/{t}"))

    with c5:
        t, tr, ntr = r5
        st.image(fig_pizza_bytes(tr, ntr, f"Região\n{tr}/{t}"))

    st.divider()
