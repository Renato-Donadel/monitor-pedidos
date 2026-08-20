import os
import sqlite3

import pandas as pd
import streamlit as st

# Banco gerado pelo pipeline horário (fluxo_devolucao/pipeline.py) — pasta irmã deste app.
FLUXO_DEVOLUCAO_DB = os.path.normpath(
    os.path.join(
        os.path.dirname(__file__),
        "..", "..",
        "fluxo_devolucao",
        "fluxo_devolucao.db",
    )
)

ESTADO_LABEL = {
    "em_devolucao": "Em Devolução",
    "devolvido_intelipost": "Devolvido (Intelipost)",
    "entregue": "Entregue",
}


def moeda(x):
    return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


@st.cache_data(ttl=60)
def carregar_pedidos():
    if not os.path.exists(FLUXO_DEVOLUCAO_DB):
        return pd.DataFrame()

    with sqlite3.connect(FLUXO_DEVOLUCAO_DB) as conn:
        return pd.read_sql("SELECT * FROM pedidos", conn)


@st.cache_data(ttl=60)
def carregar_ultima_execucao():
    if not os.path.exists(FLUXO_DEVOLUCAO_DB):
        return None

    with sqlite3.connect(FLUXO_DEVOLUCAO_DB) as conn:
        df = pd.read_sql(
            "SELECT * FROM pipeline_runs ORDER BY executado_em DESC LIMIT 1", conn
        )

    return None if df.empty else df.iloc[0]


def render_fluxo_devolucao():

    st.markdown(
        '<div class="titulo-painel">Fluxo de Devolução — TSP → Intelipost → Armazém → Entrega</div>',
        unsafe_allow_html=True,
    )

    if not os.path.exists(FLUXO_DEVOLUCAO_DB):
        st.warning(
            "Banco do fluxo de devolução ainda não foi gerado. "
            "Rode `fluxo_devolucao/pipeline.py` ao menos uma vez."
        )
        st.stop()

    df = carregar_pedidos()
    ultima_execucao = carregar_ultima_execucao()

    if ultima_execucao is not None:
        status_execucao = "✅" if ultima_execucao["status"] == "ok" else "⚠️"
        st.caption(
            f"{status_execucao} Última execução do pipeline: {ultima_execucao['executado_em']} "
            f"({ultima_execucao['linhas_processadas']} linha(s) processada(s))"
        )

    if df.empty:
        st.info("Nenhum pedido rastreado ainda.")
        st.stop()

    df["ValorNota"] = df["ValorNota"].fillna(0)

    # ==============================
    # ALERTAS — entregue com reenvio ativo
    # ==============================

    alertas = df[df["alerta_cancelamento_pendente"] == 1]

    if not alertas.empty:
        st.error(
            f"⚠️ {len(alertas)} pedido(s) foram **entregues** mas possuem **reenvio ativo** — "
            "necessário cancelamento."
        )
        with st.expander("Ver pedidos com alerta de cancelamento pendente", expanded=True):
            st.dataframe(
                alertas[
                    ["PedidoFormatado", "Transportadora", "ValorNota", "data_status_intelipost"]
                ],
                use_container_width=True,
            )
    else:
        st.success("Nenhum pedido entregue com reenvio ativo no momento.")

    # ==============================
    # RESUMO POR SEÇÃO (valor de nota)
    # ==============================

    st.markdown("### Valor de nota por seção")

    resumo = (
        df.groupby("estado_atual")["ValorNota"]
        .agg(["sum", "count"])
        .reindex(["em_devolucao", "devolvido_intelipost", "entregue"])
        .fillna(0)
    )

    cols = st.columns(3)

    for col, estado in zip(cols, resumo.index):
        valor_total = resumo.loc[estado, "sum"]
        qtd = int(resumo.loc[estado, "count"])

        with col:
            st.markdown(
                f"""
                <div style="
                    background:white;
                    padding:12px 16px;
                    border-radius:10px;
                    box-shadow:0 2px 6px rgba(0,0,0,0.06);
                    text-align:center;
                ">
                    <div style="font-size:13px;color:#6b7280;">{ESTADO_LABEL[estado]}</div>
                    <div style="font-size:20px;font-weight:800;color:#0f2a44;">
                        {moeda(valor_total)}
                    </div>
                    <div style="font-size:12px;color:#6b7280;">{qtd} pedido(s)</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    # ==============================
    # DETALHE POR SEÇÃO (abas)
    # ==============================

    aba_devolucao, aba_devolvido, aba_entregue = st.tabs(
        ["Em Devolução", "Devolvido (Intelipost)", "Entregue"]
    )

    colunas_base = ["PedidoFormatado", "Transportadora", "ValorNota", "data_entrada_tsp"]

    with aba_devolucao:
        sub = df[df["estado_atual"] == "em_devolucao"]
        st.dataframe(sub[colunas_base], use_container_width=True)

    with aba_devolvido:
        sub = df[df["estado_atual"] == "devolvido_intelipost"]
        colunas = colunas_base + ["data_status_intelipost", "confirmado_armazem"]
        st.dataframe(sub[colunas], use_container_width=True)

        pendentes = (sub["confirmado_armazem"] == "pendente").sum()
        if pendentes:
            st.caption(f"{pendentes} pedido(s) aguardando confirmação do armazém.")

    with aba_entregue:
        sub = df[df["estado_atual"] == "entregue"]
        colunas = colunas_base + ["data_status_intelipost", "tem_reenvio_ativo", "nfd_emitida"]
        st.dataframe(sub[colunas], use_container_width=True)
