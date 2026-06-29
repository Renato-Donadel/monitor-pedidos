import streamlit as st
import pandas as pd
import os
import plotly.graph_objects as go

from utils import PASTA_DATA

ARQ_REGRAS  = os.path.join(PASTA_DATA, "Regras_Resultado.xlsx")
ARQ_IMPACTO = os.path.join(PASTA_DATA, "Impacto_Financeiro.csv")


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


@st.cache_data(ttl=3600)
def carregar_impacto(path):
    """
    Lê o CSV de impacto financeiro histórico.
    TTL alto pois cotações por dia são fixas — dados só mudam ao rodar o ETL.
    """
    if not os.path.exists(path):
        return pd.DataFrame()
    df = pd.read_csv(path, sep=";")
    df["Data"]   = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
    df["Regra"]  = pd.to_numeric(df["Regra"], errors="coerce").astype("Int64")
    df["NOrders"]       = pd.to_numeric(df["NOrders"],       errors="coerce").fillna(0).astype(int)
    df["Impacto_Total"] = pd.to_numeric(df["Impacto_Total"], errors="coerce").fillna(0.0)
    if "Categoria" not in df.columns:
        df["Categoria"] = "Não classificado"
    df["Categoria"] = df["Categoria"].fillna("Sem Categoria").str.strip()
    return df.dropna(subset=["Data"])


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

    # ── Impacto Financeiro por Regra ──────────────────────────
    st.markdown("---")
    st.markdown("#### 💰 Impacto Financeiro por Regra")
    st.caption(
        "Custo de compliance: quanto foi pago a mais ao seguir cada regra de bloqueio "
        "em comparação ao que teria custado usar a transportadora bloqueada, "
        "nos casos em que ela era mais barata."
    )

    df_imp = carregar_impacto(ARQ_IMPACTO)

    if df_imp.empty:
        st.info("Nenhum dado de impacto disponível. Rode o `Regras.py` para gerar.")
    else:
        datas_disp = sorted(df_imp["Data"].dt.date.unique())

        # ── Seletor de período ────────────────────────────────
        col_d1, col_d2 = st.columns(2)
        d_ini = col_d1.date_input(
            "De", value=datas_disp[0],
            min_value=datas_disp[0], max_value=datas_disp[-1],
            key="imp_d_ini",
        )
        d_fim = col_d2.date_input(
            "Até", value=datas_disp[-1],
            min_value=datas_disp[0], max_value=datas_disp[-1],
            key="imp_d_fim",
        )

        mask       = (df_imp["Data"].dt.date >= d_ini) & (df_imp["Data"].dt.date <= d_fim)
        df_periodo = df_imp[mask].copy()

        if df_periodo.empty:
            st.warning("Nenhum dado no período selecionado.")
        else:
            n_dias          = (d_fim - d_ini).days + 1
            total_impacto   = df_periodo["Impacto_Total"].sum()
            total_orders    = df_periodo["NOrders"].sum()
            media_dia       = total_impacto / n_dias if n_dias > 0 else 0

            def brl(v):
                return f"R$ {v:,.2f}".replace(",","X").replace(".",",").replace("X",".")

            total_categorizado = df_periodo[
                ~df_periodo["Categoria"].isin(["Sem Categoria","Não classificado","Nao classificado"])
            ]["Impacto_Total"].sum()
            pct_cat = round(total_categorizado / total_impacto * 100, 1) if total_impacto > 0 else 0

            col_a, col_b, col_c, col_d, col_e = st.columns(5)
            col_a.metric("💸 Impacto Total",      brl(total_impacto))
            col_b.metric("📦 Pedidos Impactados", f"{total_orders:,}")
            col_c.metric("📅 Dias Analisados",    n_dias)
            col_d.metric("📊 Média por Dia",       brl(media_dia))
            col_e.metric("🏷️ % Categorizado",      f"{pct_cat}%",
                         help="Percentual do impacto total que já tem categoria definida na planilha de regras")

            # ── Visão por Categoria ───────────────────────────
            ORDEM_CAT = [
                "Imediato",
                "Oportunidade",
                "Ações já feitas - Capturação futura",
                "Histórico",
                "Sem Categoria",
            ]
            CORES_CAT = {
                "Imediato":                              "#e63946",
                "Oportunidade":                          "#f4a261",
                "Ações já feitas - Capturação futura":   "#2a9d8f",
                "Histórico":                             "#457b9d",
                "Sem Categoria":                         "#adb5bd",
            }

            df_cat = (
                df_periodo
                .groupby("Categoria", as_index=False)["Impacto_Total"].sum()
            )
            df_cat["_ord"] = df_cat["Categoria"].apply(
                lambda x: ORDEM_CAT.index(x) if x in ORDEM_CAT else 99
            )
            df_cat = df_cat.sort_values("_ord")

            if not df_cat.empty:
                col_pie, col_cat = st.columns([1, 1])

                with col_pie:
                    fig_pie = go.Figure(go.Pie(
                        labels=df_cat["Categoria"],
                        values=df_cat["Impacto_Total"],
                        marker_colors=[CORES_CAT.get(c, "#adb5bd") for c in df_cat["Categoria"]],
                        hole=0.45,
                        textinfo="percent+label",
                        hovertemplate="%{label}<br>R$ %{value:,.2f}<extra></extra>",
                    ))
                    fig_pie.update_layout(
                        title="Distribuição por Categoria",
                        height=320,
                        margin=dict(t=50, b=10, l=10, r=10),
                        showlegend=False,
                    )
                    st.plotly_chart(fig_pie, use_container_width=True)

                with col_cat:
                    st.markdown("**Impacto por Categoria**")
                    for _, crow in df_cat.iterrows():
                        cor  = CORES_CAT.get(crow["Categoria"], "#adb5bd")
                        pct  = crow["Impacto_Total"] / total_impacto * 100 if total_impacto > 0 else 0
                        st.markdown(
                            f"<div style='margin-bottom:10px;'>"
                            f"<span style='font-size:12px;color:var(--color-text-secondary)'>{crow['Categoria']}</span><br>"
                            f"<span style='font-size:18px;font-weight:500;color:{cor}'>{brl(crow['Impacto_Total'])}</span>"
                            f"<span style='font-size:11px;color:var(--color-text-secondary)'> &nbsp;{pct:.1f}%</span>"
                            f"</div>",
                            unsafe_allow_html=True,
                        )

            st.markdown("---")

            # ── Donut + barra: impacto por categoria ─────────
            df_cat = (
                df_periodo[
                    (df_periodo["Impacto_Total"] > 0) &
                    (~df_periodo["Categoria"].isin(["Sem Categoria", "Não classificado", "Nao classificado"]))
                ]
                .groupby("Categoria", as_index=False)["Impacto_Total"].sum()
                .sort_values("Impacto_Total", ascending=False)
            )
            if not df_cat.empty:
                CORES_CAT = {
                    "Ganho Imediato":              "#e63946",
                    "Ganho Imediato - Já realizado": "#f4a261",
                    "Oportunidade":                "#2a9d8f",
                    "Regra operacional":           "#457b9d",
                    "Express":                     "#e9c46a",
                    "Default":                     "#6d6875",
                    "Não classificado":            "#adb5bd",
                }
                cores = [CORES_CAT.get(c, "#adb5bd") for c in df_cat["Categoria"]]
                col_donut, col_cat_bar = st.columns([1, 1])
                with col_donut:
                    fig_donut = go.Figure(go.Pie(
                        labels=df_cat["Categoria"],
                        values=df_cat["Impacto_Total"],
                        hole=0.55,
                        marker_colors=cores,
                        textinfo="label+percent",
                        hovertemplate="%{label}<br>R$ %{value:,.2f}<extra></extra>",
                    ))
                    fig_donut.update_layout(
                        title="Distribuição por Categoria",
                        height=320,
                        margin=dict(t=50, b=10, l=10, r=10),
                        showlegend=False,
                    )
                    st.plotly_chart(fig_donut, use_container_width=True)
                with col_cat_bar:
                    fig_cat = go.Figure(go.Bar(
                        x=df_cat["Impacto_Total"],
                        y=df_cat["Categoria"],
                        orientation="h",
                        marker_color=cores,
                        text=df_cat["Impacto_Total"].apply(brl),
                        textposition="outside",
                        hovertemplate="%{y}<br>%{text}<extra></extra>",
                    ))
                    fig_cat.update_layout(
                        title="Impacto por Categoria (R$)",
                        height=320,
                        margin=dict(t=50, b=10, l=10, r=20),
                        xaxis=dict(tickformat=",.2f", rangemode="tozero"),
                        yaxis=dict(autorange="reversed"),
                    )
                    st.plotly_chart(fig_cat, use_container_width=True)

            # ── Barras: impacto por regra ──────────────────────
            df_por_regra = (
                df_periodo
                .groupby(["Categoria", "Regra", "Descricao"], as_index=False)
                .agg(Impacto_Total=("Impacto_Total","sum"), NOrders=("NOrders","sum"))
            )
            df_por_regra = df_por_regra[df_por_regra["Impacto_Total"] > 0].sort_values(
                "Impacto_Total", ascending=True
            )

            if not df_por_regra.empty:
                labels = (
                    "Regra " + df_por_regra["Regra"].astype(str)
                    + " — " + df_por_regra["Descricao"].str[:45]
                )
                # Capa o eixo no P85 para não deixar outliers esmagar as outras barras
                vals = df_por_regra["Impacto_Total"]
                x_cap = float(vals.quantile(0.85)) * 1.4 if len(vals) > 3 else float(vals.max()) * 1.2
                x_cap = max(x_cap, float(vals.max()) * 0.15)  # garante visibilidade mínima

                fig_bar = go.Figure(go.Bar(
                    y=labels,
                    x=df_por_regra["Impacto_Total"],
                    orientation="h",
                    marker_color="#e63946",
                    text=df_por_regra["Impacto_Total"].apply(brl),
                    textposition="auto",
                    insidetextanchor="start",
                    cliponaxis=False,
                    customdata=df_por_regra["NOrders"],
                    hovertemplate=(
                        "<b>%{y}</b><br>"
                        "Impacto: %{text}<br>"
                        "Pedidos: %{customdata}<extra></extra>"
                    ),
                ))
                fig_bar.update_layout(
                    title=f"Custo de Compliance por Regra — {d_ini.strftime('%d/%m/%Y')} a {d_fim.strftime('%d/%m/%Y')}",
                    xaxis_title="R$", yaxis_title="",
                    height=max(300, len(df_por_regra) * 50 + 80),
                    margin=dict(t=50, b=30, l=10, r=20),
                    xaxis=dict(tickformat=",.2f", range=[0, x_cap]),
                )
                st.plotly_chart(fig_bar, use_container_width=True)

            # ── Linha temporal ─────────────────────────────────
            if len(datas_disp) > 1:
                df_por_dia = (
                    df_periodo
                    .groupby("Data", as_index=False)["Impacto_Total"].sum()
                    .sort_values("Data")
                )
                # Opcional: quebrar por regra no gráfico de linha
                regras_com_impacto = sorted(
                    df_periodo[df_periodo["Impacto_Total"] > 0]["Regra"].unique()
                )
                st.markdown("**Evolução diária do impacto financeiro**")
                modo_linha = st.radio(
                    "Visualizar por:", ["Total", "Por Regra"],
                    horizontal=True, key="imp_modo_linha",
                )

                if modo_linha == "Total":
                    fig_line = go.Figure(go.Scatter(
                        x=df_por_dia["Data"],
                        y=df_por_dia["Impacto_Total"],
                        mode="lines+markers",
                        line=dict(color="#e63946", width=2),
                        fill="tozeroy",
                        fillcolor="rgba(230,57,70,0.08)",
                        name="Total",
                        hovertemplate="%{x|%d/%m/%Y}<br>R$ %{y:,.2f}<extra></extra>",
                    ))
                else:
                    fig_line = go.Figure()
                    palette = [
                        "#e63946","#457b9d","#2a9d8f","#e9c46a",
                        "#f4a261","#264653","#6d6875","#b5838d",
                    ]
                    for i, reg in enumerate(regras_com_impacto):
                        df_r = df_periodo[df_periodo["Regra"] == reg].groupby("Data", as_index=False)["Impacto_Total"].sum()
                        desc_r = df_periodo[df_periodo["Regra"] == reg]["Descricao"].iloc[0]
                        fig_line.add_trace(go.Scatter(
                            x=df_r["Data"], y=df_r["Impacto_Total"],
                            mode="lines+markers",
                            name=f"R{reg}",
                            line=dict(color=palette[i % len(palette)], width=2),
                            hovertemplate=f"Regra {reg} — {desc_r[:30]}<br>%{{x|%d/%m/%Y}}<br>R$ %{{y:,.2f}}<extra></extra>",
                        ))

                fig_line.update_layout(
                    xaxis_title="Data", yaxis_title="R$",
                    height=320,
                    margin=dict(t=20, b=30, l=10, r=10),
                    legend=dict(orientation="h", y=-0.3),
                    yaxis=dict(tickformat=",.2f"),
                )
                st.plotly_chart(fig_line, use_container_width=True)

            # ── Tabela detalhada ───────────────────────────────
            with st.expander("📊 Detalhamento completo por regra"):
                df_det = (
                    df_periodo
                    .groupby(["Regra", "Descricao"], as_index=False)
                    .agg(
                        Pedidos_Impactados=("NOrders",       "sum"),
                        Impacto_Total=     ("Impacto_Total", "sum"),
                        Dias_com_Dados=    ("Data",          "nunique"),
                    )
                    .sort_values("Impacto_Total", ascending=False)
                )
                df_det["Impacto_Médio/Dia"] = (
                    df_det["Impacto_Total"] / df_det["Dias_com_Dados"]
                ).round(2)
                df_det["Impacto_Total"]     = df_det["Impacto_Total"].round(2)
                st.dataframe(df_det, use_container_width=True, hide_index=True)