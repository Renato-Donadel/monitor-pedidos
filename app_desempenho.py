import streamlit as st
import pandas as pd
import os
from io import BytesIO
import plotly.graph_objects as go
import plotly.express as px

from utils import PASTA_DATA

COR_DENTRO  = "#2a9d8f"
COR_FORA    = "#e63946"
COR_NEUTRO  = "#1f4e79"

def render_desempenho():

    st.markdown("### 🚚 Desempenho por Transportadora")

    # ====================================================
    # CARREGAR BASE
    # ====================================================

    ARQ = os.path.join(PASTA_DATA, "Base_Pedidos_Codigo.xlsx")

    if not os.path.exists(ARQ):
        st.error("Base_Pedidos_Codigo.xlsx não encontrada. Rode o Trade-Off ETL primeiro.")
        return

    @st.cache_data(ttl=300)
    def carregar(path):
        df = pd.read_excel(path)
        df["DataFinal"] = pd.to_datetime(df["DataFinal"], errors="coerce")
        df["Mes"] = df["DataFinal"].dt.to_period("M").astype(str)
        df["DentroPrazo"] = df["DentroPrazo"].astype(bool)
        return df

    df = carregar(ARQ)

    if df.empty:
        st.warning("Base vazia.")
        return

    # ====================================================
    # FILTRO DE TRANSPORTADORA
    # ====================================================

    todas = sorted(df["Transportadora"].dropna().unique())

    selecionadas = st.multiselect(
        "🔍 Filtrar Transportadoras",
        todas,
        default=todas,
        placeholder="Selecione uma ou mais transportadoras..."
    )

    if not selecionadas:
        st.warning("Selecione ao menos uma transportadora.")
        return

    df = df[df["Transportadora"].isin(selecionadas)].copy()

    st.divider()

    # ====================================================
    # LOOP POR TRANSPORTADORA
    # ====================================================

    for transp in selecionadas:

        df_t = df[df["Transportadora"] == transp].copy()

        if df_t.empty:
            continue

        total  = len(df_t)
        dentro = df_t["DentroPrazo"].sum()
        fora   = total - dentro
        pct_dentro = round(dentro / total * 100, 1) if total > 0 else 0
        pct_fora   = round(fora   / total * 100, 1) if total > 0 else 0

        # --- Cabeçalho ---
        st.markdown(f"### 🚛 {transp}")

        col_info, col_dl = st.columns([5, 1])

        with col_info:
            st.markdown(
                f"""
                <div style="background:white;padding:10px 16px;border-radius:10px;
                            box-shadow:0 1px 4px rgba(0,0,0,0.08);display:flex;gap:32px;">
                    <div>
                        <div style="font-size:12px;color:#6b7280;">Total</div>
                        <div style="font-size:20px;font-weight:700;color:#0f2a44;">{total:,}</div>
                    </div>
                    <div>
                        <div style="font-size:12px;color:#6b7280;">Dentro do Prazo</div>
                        <div style="font-size:20px;font-weight:700;color:{COR_DENTRO};">
                            {dentro:,} ({pct_dentro}%)
                        </div>
                    </div>
                    <div>
                        <div style="font-size:12px;color:#6b7280;">Fora do Prazo</div>
                        <div style="font-size:20px;font-weight:700;color:{COR_FORA};">
                            {fora:,} ({pct_fora}%)
                        </div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

        with col_dl:
            df_fora_exp = df_t[~df_t["DentroPrazo"]].copy()
            buf = BytesIO()
            df_fora_exp.to_excel(buf, index=False)
            buf.seek(0)
            st.download_button(
                label="⬇️ Atrasados",
                data=buf,
                file_name=f"atrasados_{transp}.xlsx",
                key=f"dl_{transp}"
            )

        st.markdown("")

        # ====================================================
        # GRÁFICOS LADO A LADO
        # ====================================================

        g1, g2 = st.columns(2)

        # --- Gráfico 1: Evolução mensal ---
        with g1:

            mensal = (
                df_t.groupby("Mes")
                .agg(Total=("DentroPrazo","count"), Dentro=("DentroPrazo","sum"))
                .reset_index()
            )
            mensal["SLA%"] = (mensal["Dentro"] / mensal["Total"] * 100).round(1)
            mensal = mensal.sort_values("Mes")

            fig1 = go.Figure()

            fig1.add_trace(go.Scatter(
                x=mensal["Mes"],
                y=mensal["SLA%"],
                mode="lines+markers+text",
                text=mensal["SLA%"].apply(lambda v: f"{v:.1f}%"),
                textposition="top center",
                textfont=dict(size=11),
                line=dict(color=COR_NEUTRO, width=2),
                marker=dict(size=8, color=COR_NEUTRO),
                name="SLA%"
            ))

            fig1.add_hline(
                y=mensal["SLA%"].mean(),
                line_dash="dash",
                line_color="#f4a261",
                annotation_text=f"Média: {mensal['SLA%'].mean():.1f}%",
                annotation_font_size=11,
                annotation_position="top left"
            )

            fig1.update_layout(
                title="Evolução Mensal — SLA%",
                xaxis_title="Mês",
                yaxis_title="SLA (%)",
                height=300,
                plot_bgcolor="#f4f6f9",
                paper_bgcolor="#f4f6f9",
                margin=dict(t=40, b=40, l=40, r=20),
                font=dict(family="Arial", size=12),
                yaxis=dict(range=[
                    max(0, mensal["SLA%"].min() - 10),
                    min(100, mensal["SLA%"].max() + 10)
                ])
            )

            st.plotly_chart(fig1, use_container_width=True, key=f"g1_{transp}")

        # --- Gráfico 2: SLA por código tarifário ---
        with g2:

            por_cod = (
                df_t.groupby("CodigoTarifario")
                .agg(Total=("DentroPrazo","count"), Dentro=("DentroPrazo","sum"))
                .reset_index()
            )
            por_cod["SLA%"] = (por_cod["Dentro"] / por_cod["Total"] * 100).round(1)
            por_cod = por_cod.sort_values("SLA%", ascending=True)

            fig2 = go.Figure(go.Bar(
                x=por_cod["SLA%"],
                y=por_cod["CodigoTarifario"],
                orientation="h",
                marker_color=[
                    COR_DENTRO if v >= 90 else
                    "#f4a261"   if v >= 75 else
                    COR_FORA
                    for v in por_cod["SLA%"]
                ],
                text=por_cod["SLA%"].apply(lambda v: f"{v:.1f}%"),
                textposition="outside",
                textfont=dict(size=11),
            ))

            fig2.update_layout(
                title="SLA% por Código Tarifário",
                xaxis_title="SLA (%)",
                xaxis=dict(range=[0, 115]),
                yaxis=dict(tickfont=dict(size=10)),
                height=max(300, len(por_cod) * 28 + 60),
                plot_bgcolor="#f4f6f9",
                paper_bgcolor="#f4f6f9",
                margin=dict(t=40, b=40, l=10, r=60),
                font=dict(family="Arial", size=12),
            )

            st.plotly_chart(fig2, use_container_width=True, key=f"g2_{transp}")

        st.caption("🟢 SLA ≥ 90%  |  🟠 SLA entre 75–90%  |  🔴 SLA < 75%")

        # ====================================================
        # DETALHE: MÊS A MÊS POR CÓDIGO TARIFÁRIO
        # ====================================================

        with st.expander(f"📅 Ver mês a mês por código tarifário — {transp}"):

            codigos = sorted(df_t["CodigoTarifario"].dropna().unique())

            cod_sel = st.selectbox(
                "Código Tarifário",
                codigos,
                key=f"cod_{transp}"
            )

            df_cod = df_t[df_t["CodigoTarifario"] == cod_sel].copy()

            mensal_cod = (
                df_cod.groupby("Mes")
                .agg(Total=("DentroPrazo","count"), Dentro=("DentroPrazo","sum"))
                .reset_index()
            )
            mensal_cod["SLA%"] = (mensal_cod["Dentro"] / mensal_cod["Total"] * 100).round(1)
            mensal_cod["Fora%"] = (100 - mensal_cod["SLA%"]).round(1)
            mensal_cod = mensal_cod.sort_values("Mes")

            fig3 = go.Figure()

            fig3.add_trace(go.Bar(
                x=mensal_cod["Mes"],
                y=mensal_cod["SLA%"],
                name="Dentro do Prazo",
                marker_color=COR_DENTRO,
                text=mensal_cod["SLA%"].apply(lambda v: f"{v:.1f}%"),
                textposition="inside",
                textfont=dict(color="white", size=11),
            ))

            fig3.add_trace(go.Bar(
                x=mensal_cod["Mes"],
                y=mensal_cod["Fora%"],
                name="Fora do Prazo",
                marker_color=COR_FORA,
                text=mensal_cod["Fora%"].apply(lambda v: f"{v:.1f}%"),
                textposition="inside",
                textfont=dict(color="white", size=11),
            ))

            fig3.update_layout(
                barmode="stack",
                title=f"Mês a Mês — {cod_sel}",
                xaxis_title="Mês",
                yaxis_title="% Pedidos",
                yaxis=dict(range=[0, 110]),
                height=320,
                plot_bgcolor="#f4f6f9",
                paper_bgcolor="#f4f6f9",
                legend=dict(orientation="h", y=-0.25),
                margin=dict(t=40, b=60, l=40, r=20),
                font=dict(family="Arial", size=12),
            )

            st.plotly_chart(fig3, use_container_width=True, key=f"g3_{transp}_{cod_sel}")

            # Tabela resumo
            tabela = mensal_cod[["Mes","Total","Dentro","SLA%","Fora%"]].rename(columns={
                "Mes": "Mês", "Total": "Pedidos",
                "Dentro": "Dentro do Prazo",
                "SLA%": "SLA %", "Fora%": "Fora %"
            })
            st.dataframe(
                tabela.style.format({"SLA %": "{:.1f}%", "Fora %": "{:.1f}%"}),
                use_container_width=True,
                hide_index=True
            )

        st.divider()