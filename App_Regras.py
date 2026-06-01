import streamlit as st
import pandas as pd
import os
import plotly.graph_objects as go

from utils import PASTA_DATA

ARQ_REGRAS = os.path.join(PASTA_DATA, "Regras_Resultado.xlsx")


@st.cache_data(ttl=60)
def carregar(path):
    resultado  = pd.read_excel(path, sheet_name="Resultado")
    violacoes  = pd.read_excel(path, sheet_name="Violacoes")
    cotacoes   = pd.read_excel(path, sheet_name="Cotacoes_Raw")
    try:
        skus_novos = pd.read_excel(path, sheet_name="SKUs_Novos")
    except Exception:
        skus_novos = pd.DataFrame()
    try:
        regra26_cots = pd.read_excel(path, sheet_name="Regra26_Cotacoes")
    except Exception:
        regra26_cots = pd.DataFrame()
    return resultado, violacoes, cotacoes, skus_novos, regra26_cots


def render_regras():
    st.markdown("### 📋 Monitor de Regras PRW")

    if not os.path.exists(ARQ_REGRAS):
        st.error("Arquivo `Regras_Resultado.xlsx` não encontrado na pasta `data/`.")
        st.info("Rode o `Regras.py` na rede para gerar o arquivo antes de abrir esta página.")
        return

    df_res, df_viols, df_cot, df_skus_novos, df_r26_cots = carregar(ARQ_REGRAS)

    # Data da última atualização
    if "AtualizadoEm" in df_res.columns and not df_res["AtualizadoEm"].isna().all():
        ultima = df_res["AtualizadoEm"].iloc[0]
        st.caption(f"🕐 Última atualização: **{ultima}**  |  Rode o `Regras.py` para atualizar os dados.")

    # ── Aviso SKUs novos ──────────────────────────────────────
    if not df_skus_novos.empty:
        st.warning(
            f"⚠️ **{len(df_skus_novos)} SKU(s) novo(s) detectado(s)** — "
            f"A planilha de SKUs já foi atualizada automaticamente pelo ETL. "
            f"Valide os produtos abaixo antes de rodar novamente."
        )
        with st.expander(f"📦 Ver {len(df_skus_novos)} SKU(s) novo(s) detectado(s)"):
            st.info("Estes produtos foram identificados por palavra-chave e adicionados à planilha da rede. "
                    "Valide se são de fato da categoria correta ou se devem ser removidos.")
            cols_sku = [c for c in ["CODIGO_SKU","Descricao","Categoria"] if c in df_skus_novos.columns]
            st.dataframe(df_skus_novos[cols_sku].drop_duplicates(), use_container_width=True)

    # ── KPIs ──────────────────────────────────────────────────
    ok_count  = (df_res["Status"] == "OK").sum()
    nok_count = (df_res["Status"] == "VIOLADA").sum()
    sem_count = (df_res["Status"] == "SEM_VERIFICACAO").sum()
    verif     = ok_count + nok_count
    pct       = round(ok_count / verif * 100, 1) if verif > 0 else 0

    # Prejuízo da regra 26
    prejuizo = 0.0
    if "Prejuizo_Total" in df_res.columns:
        val = df_res[df_res["Regra"] == 26]["Prejuizo_Total"]
        if not val.empty and not pd.isna(val.iloc[0]):
            prejuizo = float(val.iloc[0])

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total regras ativas",     len(df_res))
    col2.metric("✅ Seguidas",             ok_count)
    col3.metric("❌ Violadas",             nok_count)
    col4.metric("⚪ Sem verificação",      sem_count)
    col5.metric("💸 Prejuízo Leilão",     f"R$ {prejuizo:,.2f}".replace(",","X").replace(".",",").replace("X","."))

    # ── Gauge ─────────────────────────────────────────────────
    if verif > 0:
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=pct,
            title={"text": "Conformidade (%)"},
            gauge={
                "axis": {"range": [0, 100]},
                "bar":  {"color": "#2a9d8f" if pct >= 80 else "#e63946"},
                "steps": [
                    {"range": [0,  60], "color": "#fde8e8"},
                    {"range": [60, 80], "color": "#fff3cd"},
                    {"range": [80,100], "color": "#d4edda"},
                ],
                "threshold": {
                    "line": {"color": "green", "width": 4},
                    "thickness": 0.75,
                    "value": 95,
                },
            }
        ))
        fig.update_layout(height=260, margin=dict(t=40, b=10, l=10, r=10))
        st.plotly_chart(fig, use_container_width=True)

    # ── Regras VIOLADAS ───────────────────────────────────────
    df_nok = df_res[df_res["Status"] == "VIOLADA"].sort_values("Violacoes", ascending=False)
    if not df_nok.empty:
        st.markdown("---")
        st.markdown("#### ❌ Regras com Violações")
        for _, row in df_nok.iterrows():
            n     = row["Regra"]
            desc  = row["Descricao"]
            viols = int(row["Violacoes"])
            total = int(row["TotalPedidos"])
            pct_v = round(viols / total * 100, 2) if total > 0 else 0

            # Título especial para regra 26 com prejuízo
            if n == 26 and prejuizo > 0:
                prej_fmt = f"R$ {prejuizo:,.2f}".replace(",","X").replace(".",",").replace("X",".")
                titulo = f"🔴 Regra **{n}** — {desc}  |  {viols} violação(ões) ({pct_v}%)  |  💸 Prejuízo: {prej_fmt}"
            else:
                titulo = f"🔴 Regra **{n}** — {desc}  |  {viols} violação(ões) ({pct_v}%)"

            with st.expander(titulo):
                if n == 26 and not df_r26_cots.empty:
                    # Regra 26: mostra todas as cotações dos pedidos violados
                    st.info("Todas as cotações dos pedidos que fugiram do mais barato. "
                            "A linha com CotacaoRow=1 é a mais barata disponível.")
                    pedidos_disp = sorted(df_r26_cots["PedidoID"].unique()) if "PedidoID" in df_r26_cots.columns else []
                    sel_pedido   = st.selectbox(f"Filtrar pedido (Regra 26)", ["Todos"] + [str(p) for p in pedidos_disp], key=f"sel26")
                    if sel_pedido != "Todos":
                        df_view26 = df_r26_cots[df_r26_cots["PedidoID"] == int(sel_pedido)]
                    else:
                        df_view26 = df_r26_cots.head(200)
                    st.dataframe(df_view26, use_container_width=True)
                else:
                    amostra = df_viols[df_viols["Regra"] == n]
                    if not amostra.empty:
                        cols_show = [c for c in amostra.columns if c != "Regra"]
                        st.dataframe(amostra[cols_show], use_container_width=True)
                    else:
                        st.info("Amostra não disponível.")

    # ── Regras OK ─────────────────────────────────────────────
    df_ok = df_res[df_res["Status"] == "OK"]
    if not df_ok.empty:
        st.markdown("---")
        st.markdown("#### ✅ Regras Sendo Seguidas")
        cols = st.columns(3)
        for i, (_, row) in enumerate(df_ok.iterrows()):
            cols[i % 3].success(f"Regra {int(row['Regra'])} — {str(row['Descricao'])[:55]}")

    # ── Sem verificação ───────────────────────────────────────
    df_sem = df_res[df_res["Status"] == "SEM_VERIFICACAO"]
    if not df_sem.empty:
        st.markdown("---")
        with st.expander(f"⚪ {len(df_sem)} regra(s) sem verificação automática"):
            for _, row in df_sem.iterrows():
                st.write(f"- **Regra {int(row['Regra'])}**: {row['Descricao']}")

    # ── Explorador de cotações ────────────────────────────────
    st.markdown("---")
    st.markdown("#### 🔍 Explorar Cotações")
    if not df_cot.empty:
        carriers = sorted(df_cot["Escolhida"].dropna().unique()) if "Escolhida" in df_cot.columns else []
        sel = st.multiselect("Filtrar transportadora", carriers, placeholder="Todas")
        df_view = df_cot[df_cot["Escolhida"].isin(sel)] if sel else df_cot
        st.dataframe(df_view.head(500), use_container_width=True)
    else:
        st.info("Sem cotações disponíveis.")