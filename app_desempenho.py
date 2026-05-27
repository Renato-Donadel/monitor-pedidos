import streamlit as st
import pandas as pd
import os
from io import BytesIO
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta

from utils import PASTA_DATA

COR_DENTRO  = "#2a9d8f"
COR_FORA    = "#e63946"
COR_NEUTRO  = "#1f4e79"
COR_META    = "#e9c46a"

# Colunas a EXCLUIR nos exports
COLUNAS_EXCLUIR = ["ChaveNF", "Pedido de Venda", "CEP", "CepFinal", "PedidoFormatado"]

# Colunas que devem vir como texto (para evitar notação científica)
COLUNAS_TEXTO = ["ChaveNF", "PedidoFormatado"]

def preparar_export(df_exp: pd.DataFrame) -> pd.DataFrame:
    """Remove colunas desnecessárias e formata ChaveNF como texto."""
    df_exp = df_exp.copy()

    # Garante que ChaveNF apareça como texto (antes de remover)
    for col in COLUNAS_TEXTO:
        if col in df_exp.columns:
            df_exp[col] = df_exp[col].astype(str).str.strip()

    # Remove colunas indesejadas
    cols_remover = [c for c in COLUNAS_EXCLUIR if c in df_exp.columns]
    df_exp = df_exp.drop(columns=cols_remover)

    # Renomeia DataPrevista → Previsão Transportadora
    if "DataPrevista" in df_exp.columns:
        df_exp = df_exp.rename(columns={"DataPrevista": "Previsão Transportadora"})

    # Renomeia DataFinal → Data Final
    if "DataFinal" in df_exp.columns:
        df_exp = df_exp.rename(columns={"DataFinal": "Data Final"})

    return df_exp


def to_excel_texto(df_exp: pd.DataFrame) -> bytes:
    """Exporta para Excel garantindo que ChaveNF venha como texto."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_exp.to_excel(writer, index=False, sheet_name="Dados")
        ws = writer.sheets["Dados"]

        # Formata colunas de texto para evitar notação científica
        from openpyxl.styles import numbers as xl_numbers
        for col_idx, col_name in enumerate(df_exp.columns, start=1):
            if col_name in ["ChaveNF", "Previsão Transportadora", "Data Final"]:
                for row_idx in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.number_format = "@"  # força texto

    buf.seek(0)
    return buf.getvalue()


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
        df["DataFinal"]   = pd.to_datetime(df["DataFinal"],   errors="coerce")
        df["DataPrevista"] = pd.to_datetime(df["DataPrevista"], errors="coerce")
        df["Mes"]          = df["DataFinal"].dt.to_period("M").astype(str)
        df["DentroPrazo"]  = df["DentroPrazo"].astype(bool)
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

        total      = len(df_t)
        dentro     = df_t["DentroPrazo"].sum()
        fora       = total - dentro
        pct_dentro = round(dentro / total * 100, 1) if total > 0 else 0
        pct_fora   = round(fora   / total * 100, 1) if total > 0 else 0

        # --- Cabeçalho ---
        st.markdown(f"### 🚛 {transp}")

        col_info, col_dl_atr, col_dl_ok = st.columns([5, 1, 1])

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

        # --- Botão: Atrasados ---
        with col_dl_atr:
            df_fora_exp = preparar_export(df_t[~df_t["DentroPrazo"]].copy())
            st.download_button(
                label="⬇️ Atrasados",
                data=to_excel_texto(df_fora_exp),
                file_name=f"atrasados_{transp}.xlsx",
                key=f"dl_atr_{transp}"
            )

        # --- Botão: Dentro do prazo (últimos 30 dias) ---
        with col_dl_ok:
            corte_30d = pd.Timestamp(datetime.today() - timedelta(days=30))
            df_ok_30 = df_t[
                df_t["DentroPrazo"] &
                (df_t["DataFinal"] >= corte_30d)
            ].copy()
            df_ok_exp = preparar_export(df_ok_30)
            st.download_button(
                label="⬇️ No Prazo 30d",
                data=to_excel_texto(df_ok_exp),
                file_name=f"no_prazo_30d_{transp}.xlsx",
                key=f"dl_ok_{transp}"
            )

        st.markdown("")

        # ====================================================
        # GRÁFICO 1: Evolução mensal — largura total
        # ====================================================

        mensal = (
            df_t.groupby("Mes")
            .agg(Total=("DentroPrazo", "count"), Dentro=("DentroPrazo", "sum"))
            .reset_index()
        )
        mensal["SLA%"] = (mensal["Dentro"] / mensal["Total"] * 100).round(1)
        mensal = mensal.sort_values("Mes")

        # Formata eixo X como "Março 2026", "Abril 2026"...
        MESES_PT = {
            "01": "Janeiro", "02": "Fevereiro", "03": "Março",
            "04": "Abril",   "05": "Maio",       "06": "Junho",
            "07": "Julho",   "08": "Agosto",     "09": "Setembro",
            "10": "Outubro", "11": "Novembro",   "12": "Dezembro"
        }

        def formatar_mes(periodo_str):
            # periodo_str: "2026-03"
            partes = str(periodo_str).split("-")
            if len(partes) == 2:
                ano, mes = partes
                return f"{MESES_PT.get(mes, mes)} {ano}"
            return periodo_str

        mensal["MesLabel"] = mensal["Mes"].apply(formatar_mes)

        fig1 = go.Figure()

        fig1.add_trace(go.Scatter(
            x=mensal["MesLabel"],
            y=mensal["SLA%"],
            mode="lines+markers+text",
            text=mensal["SLA%"].apply(lambda v: f"{v:.1f}%"),
            textposition="top center",
            textfont=dict(size=12, color=COR_NEUTRO),
            line=dict(color=COR_NEUTRO, width=3),
            marker=dict(size=10, color=COR_NEUTRO),
            name="SLA%"
        ))

        # Linha de média
        media_sla = mensal["SLA%"].mean()
        fig1.add_hline(
            y=media_sla,
            line_dash="dash",
            line_color="#f4a261",
            annotation_text=f"Média: {media_sla:.1f}%",
            annotation_font_size=12,
            annotation_position="top left"
        )

        # Linha de meta 95%
        fig1.add_hline(
            y=95,
            line_dash="dot",
            line_color=COR_META,
            line_width=2,
            annotation_text="Meta: 95%",
            annotation_font_size=12,
            annotation_font_color=COR_META,
            annotation_position="top right"
        )

        y_min = max(0,  mensal["SLA%"].min() - 15)
        y_max = min(105, max(mensal["SLA%"].max() + 12, 97))

        fig1.update_layout(
            title="Evolução Mensal — SLA%",
            xaxis_title="Mês",
            yaxis_title="SLA (%)",
            height=340,
            plot_bgcolor="#f4f6f9",
            paper_bgcolor="#f4f6f9",
            margin=dict(t=50, b=50, l=50, r=80),
            font=dict(family="Arial", size=13),
            yaxis=dict(range=[y_min, y_max]),
            xaxis=dict(type="category")
        )

        st.plotly_chart(fig1, use_container_width=True, key=f"g1_{transp}")

        # ====================================================
        # GRÁFICO 2: SLA por código tarifário — horizontal com scroll
        # ====================================================

        por_cod = (
            df_t.groupby("CodigoTarifario")
            .agg(Total=("DentroPrazo", "count"), Dentro=("DentroPrazo", "sum"))
            .reset_index()
        )
        por_cod["SLA%"] = (por_cod["Dentro"] / por_cod["Total"] * 100).round(1)
        por_cod = por_cod.sort_values("SLA%", ascending=False)

        n_codigos = len(por_cod)
        # Largura mínima: 120px por código para garantir scroll horizontal
        largura_grafico = max(900, n_codigos * 90)

        fig2 = go.Figure(go.Bar(
            x=por_cod["CodigoTarifario"],
            y=por_cod["SLA%"],
            orientation="v",
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

        # Meta 95% no gráfico de barras também
        fig2.add_hline(
            y=95,
            line_dash="dot",
            line_color=COR_META,
            line_width=2,
            annotation_text="Meta: 95%",
            annotation_font_size=11,
            annotation_font_color=COR_META,
            annotation_position="top right"
        )

        fig2.update_layout(
            title="SLA% por Código Tarifário",
            yaxis_title="SLA (%)",
            yaxis=dict(range=[0, 115]),
            xaxis=dict(tickangle=-35, tickfont=dict(size=10)),
            height=420,
            width=largura_grafico,
            plot_bgcolor="#f4f6f9",
            paper_bgcolor="#f4f6f9",
            margin=dict(t=50, b=100, l=50, r=60),
            font=dict(family="Arial", size=12),
        )

        # Container com scroll horizontal
        st.markdown(
            f"""
            <div style="overflow-x:auto; border:1px solid #e5e7eb;
                        border-radius:10px; padding:8px; background:#f4f6f9;">
            """,
            unsafe_allow_html=True
        )
        st.plotly_chart(fig2, use_container_width=False, key=f"g2_{transp}")
        st.markdown("</div>", unsafe_allow_html=True)

        st.caption("🟢 SLA ≥ 90%  |  🟠 SLA entre 75–90%  |  🔴 SLA < 75%  |  🟡 Meta: 95%")

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
                .agg(Total=("DentroPrazo", "count"), Dentro=("DentroPrazo", "sum"))
                .reset_index()
            )
            mensal_cod["SLA%"]  = (mensal_cod["Dentro"] / mensal_cod["Total"] * 100).round(1)
            mensal_cod["Fora%"] = (100 - mensal_cod["SLA%"]).round(1)
            mensal_cod = mensal_cod.sort_values("Mes")
            mensal_cod["MesLabel"] = mensal_cod["Mes"].apply(formatar_mes)

            fig3 = go.Figure()

            fig3.add_trace(go.Bar(
                x=mensal_cod["MesLabel"],
                y=mensal_cod["SLA%"],
                name="Dentro do Prazo",
                marker_color=COR_DENTRO,
                text=mensal_cod["SLA%"].apply(lambda v: f"{v:.1f}%"),
                textposition="inside",
                textfont=dict(color="white", size=11),
            ))

            fig3.add_trace(go.Bar(
                x=mensal_cod["MesLabel"],
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
                legend=dict(orientation="h", y=-0.3),
                margin=dict(t=40, b=70, l=40, r=20),
                font=dict(family="Arial", size=12),
                xaxis=dict(type="category")
            )

            st.plotly_chart(fig3, use_container_width=True, key=f"g3_{transp}_{cod_sel}")

            # Tabela resumo
            tabela = mensal_cod[["MesLabel", "Total", "Dentro", "SLA%", "Fora%"]].rename(columns={
                "MesLabel": "Mês",
                "Total": "Pedidos",
                "Dentro": "Dentro do Prazo",
                "SLA%": "SLA %",
                "Fora%": "Fora %"
            })
            st.dataframe(
                tabela.style.format({"SLA %": "{:.1f}%", "Fora %": "{:.1f}%"}),
                use_container_width=True,
                hide_index=True
            )

        st.divider()