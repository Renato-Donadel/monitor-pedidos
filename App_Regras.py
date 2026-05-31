import streamlit as st
import pandas as pd
import os
import plotly.graph_objects as go
import plotly.express as px

from utils import PASTA_DATA

ARQ_REGRAS = os.path.join(PASTA_DATA, "Regras_Resultado.xlsx")


@st.cache_data(ttl=60)
def carregar(path):
    resultado  = pd.read_excel(path, sheet_name="Resultado")
    violacoes  = pd.read_excel(path, sheet_name="Violacoes")
    cotacoes   = pd.read_excel(path, sheet_name="Cotacoes_Raw")
    return resultado, violacoes, cotacoes


def render_regras():
    st.markdown("### 📋 Monitor de Regras PRW")

    if not os.path.exists(ARQ_REGRAS):
        st.error("Arquivo `Regras_Resultado.xlsx` não encontrado na pasta `data/`.")
        st.info("Rode o `Regras.py` na rede para gerar o arquivo antes de abrir esta página.")
        return

    df_res, df_viols, df_cot = carregar(ARQ_REGRAS)

    # Data da última atualização
    if "AtualizadoEm" in df_res.columns and not df_res["AtualizadoEm"].isna().all():
        ultima = df_res["AtualizadoEm"].iloc[0]
        st.caption(f"🕐 Última atualização: **{ultima}**  |  Rode o `Regras.py` para atualizar os dados.")

    # ── KPIs ──────────────────────────────────────────────────
    ok_count  = (df_res["Status"] == "OK").sum()
    nok_count = (df_res["Status"] == "VIOLADA").sum()
    sem_count = (df_res["Status"] == "SEM_VERIFICACAO").sum()
    verif     = ok_count + nok_count
    pct       = round(ok_count / verif * 100, 1) if verif > 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total regras ativas",     len(df_res))
    col2.metric("✅ Seguidas",             ok_count)
    col3.metric("❌ Violadas",             nok_count)
    col4.metric("⚪ Sem verificação auto.", sem_count)

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
            n    = row["Regra"]
            desc = row["Descricao"]
            viols = int(row["Violacoes"])
            total = int(row["TotalPedidos"])
            pct_v = round(viols / total * 100, 2) if total > 0 else 0

            with st.expander(f"🔴 Regra **{n}** — {desc}  |  {viols} violação(ões) ({pct_v}%)"):
                amostra = df_viols[df_viols["Regra"] == n]
                if not amostra.empty:
                    cols_show = [c for c in [
                        "ShipmentOrderId","CarrierName","CampaignCode",
                        "OriginZipCode","DestinationStateCode","TotalWeight","QuoteDate"
                    ] if c in amostra.columns]
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
        carriers = sorted(df_cot["CarrierName"].dropna().unique()) if "CarrierName" in df_cot.columns else []
        sel = st.multiselect("Filtrar transportadora", carriers, placeholder="Todas")
        df_view = df_cot[df_cot["CarrierName"].isin(sel)] if sel else df_cot
        st.dataframe(df_view.head(500), use_container_width=True)
    else:
        st.info("Sem cotações disponíveis.")