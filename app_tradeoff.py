import os
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from utils import (
    ler_base,
    ARQ_ATUAL,
    STATUS_DIARIOS,
    STATUS_Manuais,
    listar_dias,
    PASTA_MENSAL,
    TAMANHO_LOTE,
    caminho
)

# ======================================================
# CORES
# ======================================================

COR_ORIGEM  = "#1f4e79"
COR_DESTINO = "#2e86ab"
COR_ALERTA  = "#e63946"
COR_OK      = "#2a9d8f"
COR_NEUTRO  = "#f4a261"

# ======================================================
# HELPERS
# ======================================================

def fmt_pct(v):
    return f"{v:.2f}%" if pd.notna(v) else "—"

def fmt_brl(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if pd.notna(v) else "—"

def delta_icon(val, positivo_bom=True):
    """Retorna ícone e cor de acordo com direção do delta."""
    if pd.isna(val) or val == 0:
        return "➡️", "#888888"
    if positivo_bom:
        return ("🟢 ▲", COR_OK) if val > 0 else ("🔴 ▼", COR_ALERTA)
    else:
        # Para custo: queda é boa
        return ("🟢 ▼", COR_OK) if val < 0 else ("🔴 ▲", COR_ALERTA)


# ======================================================
# RENDER PRINCIPAL
# ======================================================

def render_tradeoff():

    st.markdown("## 🚚 Trade-Off Logístico")
    st.caption(
        "Compare o desempenho real entre transportadoras por faixa de CEP — "
        "SLA, Custo (TM) e NFD lado a lado, com projeção de impacto operacional."
    )
    st.divider()

    # ====================================================
    # CARREGAR BASES
    # ====================================================

    if os.path.exists("data/Base_Pedidos_Codigo.xlsx"):
        df_trade = pd.read_excel("data/Base_Pedidos_Codigo.xlsx")
    else:
        df_trade = pd.DataFrame()

    if os.path.exists("data/Base_Similaridade_Tarifarios.xlsx"):
        df_sim_base = pd.read_excel("data/Base_Similaridade_Tarifarios.xlsx")
    else:
        df_sim_base = pd.DataFrame()

    if df_trade.empty or df_sim_base.empty:
        st.error(
            "⚠️ Bases de dados não encontradas. "
            "Execute o script `trade-off.py` para gerar os arquivos em `data/`."
        )
        return

    # ====================================================
    # FILTROS — LINHA 1
    # ====================================================

    col1, col2, col3 = st.columns(3)

    with col1:
        transportadoras_disponiveis = sorted(
            df_trade["Transportadora"].dropna().unique()
        )
        transportadora_origem = st.selectbox(
            "📦 Transportadora Atual",
            transportadoras_disponiveis
        )

    with col2:
        destinos_disponiveis = sorted(
            df_sim_base[
                df_sim_base["TransportadoraOrigem"] == transportadora_origem
            ]["TransportadoraDestino"].dropna().unique()
        )
        if not destinos_disponiveis:
            destinos_disponiveis = ["MAGALU", "IMILE"]

        transportadora_destino = st.selectbox(
            "🔄 Transportadora Simulada",
            destinos_disponiveis
        )

    with col3:
        periodo = st.selectbox(
            "📅 Período",
            ["30 dias", "60 dias", "90 dias"]
        )
        dias = {"30 dias": 30, "60 dias": 60, "90 dias": 90}[periodo]

    # Filtro de período
    df_trade["DataFinal"] = pd.to_datetime(df_trade["DataFinal"], errors="coerce")
    data_corte = pd.Timestamp.today() - pd.Timedelta(days=dias)
    df_trade_periodo = df_trade[df_trade["DataFinal"] >= data_corte].copy()

    # ====================================================
    # FILTROS — CÓDIGO TARIFÁRIO
    # ====================================================

    codigos_origem = sorted(
        df_trade_periodo[
            df_trade_periodo["Transportadora"] == transportadora_origem
        ]["CodigoTarifario"].dropna().unique()
    )

    if not codigos_origem:
        st.warning("Nenhum código tarifário encontrado para a transportadora e período selecionados.")
        return

    codigo_origem = st.selectbox("🏷️ Código Tarifário de Origem", codigos_origem)

    st.divider()

    # ====================================================
    # DADOS DE ORIGEM
    # ====================================================

    df_origem = df_trade_periodo[
        (df_trade_periodo["Transportadora"] == transportadora_origem) &
        (df_trade_periodo["CodigoTarifario"] == codigo_origem)
    ].copy()

    total_pedidos_origem = len(df_origem)

    sla_origem = round(
        df_origem["DentroPrazo"].astype(bool).mean() * 100, 2
    ) if total_pedidos_origem > 0 else 0

    tm_origem = round(
        df_origem["ValorFrete"].astype(float).mean(), 2
    ) if total_pedidos_origem > 0 else 0

    nfd_origem = round(
        df_origem["TemNFD"].astype(bool).mean() * 100, 2
    ) if total_pedidos_origem > 0 else 0

    gasto_total_origem = round(
        df_origem["ValorFrete"].astype(float).sum(), 2
    )

    # ====================================================
    # FILTRAR SIMILARIDADE
    # ====================================================

    df_sim_filtrado = df_sim_base[
        (df_sim_base["TransportadoraOrigem"] == transportadora_origem) &
        (df_sim_base["CodigoOrigem"] == codigo_origem) &
        (df_sim_base["TransportadoraDestino"] == transportadora_destino)
    ].copy()

    if df_sim_filtrado.empty:
        st.warning(
            f"⚠️ Nenhuma similaridade encontrada entre **{transportadora_origem} / {codigo_origem}** "
            f"e **{transportadora_destino}**. Verifique se o ETL foi executado com esses dados."
        )
        return

    # ====================================================
    # COBERTURA — PEDIDOS COM E SEM MATCH
    # ====================================================

    pedidos_com_cobertura = int(df_sim_filtrado["Pedidos"].sum()) \
        if "Pedidos" in df_sim_filtrado.columns else 0

    # Recalcula pedidos se coluna não veio pronta
    if pedidos_com_cobertura == 0 and "Percentual" in df_sim_filtrado.columns:
        pedidos_com_cobertura = int(
            (total_pedidos_origem * df_sim_filtrado["Percentual"] / 100).sum()
        )

    pedidos_sem_cobertura = total_pedidos_origem - pedidos_com_cobertura
    pct_cobertura = round((pedidos_com_cobertura / total_pedidos_origem) * 100, 2) \
        if total_pedidos_origem > 0 else 0
    pct_sem_cobertura = round(100 - pct_cobertura, 2)

    # ====================================================
    # MÉTRICAS POR CÓDIGO DESTINO
    # ====================================================

    # Busca TM, SLA, NFD reais do destino a partir da base de pedidos
    df_destino_real = df_trade_periodo[
        (df_trade_periodo["Transportadora"] == transportadora_destino) &
        (df_trade_periodo["CodigoTarifario"].isin(df_sim_filtrado["CodigoDestino"].unique()))
    ].copy()

    metricas_dest = df_destino_real.groupby("CodigoTarifario").agg(
        TM_Destino=("ValorFrete", "mean"),
        SLA_Destino=("DentroPrazo", "mean"),
        NFD_Destino=("TemNFD", "mean"),
        Pedidos_Reais_Dest=("ValorFrete", "count")
    ).reset_index().rename(columns={"CodigoTarifario": "CodigoDestino"})

    metricas_dest["SLA_Destino"] = metricas_dest["SLA_Destino"] * 100
    metricas_dest["NFD_Destino"] = metricas_dest["NFD_Destino"] * 100

    # Join com similaridade
    df_comp = df_sim_filtrado.merge(metricas_dest, on="CodigoDestino", how="left")

    # Preenche nulos
    df_comp["TM_Destino"]  = df_comp["TM_Destino"].fillna(0)
    df_comp["SLA_Destino"] = df_comp["SLA_Destino"].fillna(0)
    df_comp["NFD_Destino"] = df_comp["NFD_Destino"].fillna(0)

    # Pedidos simulados por código destino
    df_comp["PedidosSimulados"] = (
        total_pedidos_origem * (df_comp["Percentual"] / 100)
    ).round(0).astype(int)

    # Projeção financeira
    df_comp["FreteProjetado"] = df_comp["TM_Destino"] * df_comp["PedidosSimulados"]

    # Deltas vs origem
    df_comp["Delta_SLA"] = df_comp["SLA_Destino"] - sla_origem
    df_comp["Delta_TM"]  = df_comp["TM_Destino"]  - tm_origem
    df_comp["Delta_NFD"] = df_comp["NFD_Destino"]  - nfd_origem

    # ====================================================
    # KPIs PONDERADOS CONSOLIDADOS
    # ====================================================

    total_simulados = df_comp["PedidosSimulados"].sum()

    sla_destino_pond = round(
        (df_comp["SLA_Destino"] * df_comp["PedidosSimulados"]).sum() / total_simulados, 2
    ) if total_simulados > 0 else 0

    nfd_destino_pond = round(
        (df_comp["NFD_Destino"] * df_comp["PedidosSimulados"]).sum() / total_simulados, 2
    ) if total_simulados > 0 else 0

    tm_destino_pond = round(
        df_comp["FreteProjetado"].sum() / total_simulados, 2
    ) if total_simulados > 0 else 0

    gasto_projetado = round(df_comp["FreteProjetado"].sum(), 2)

    economia_projetada = round(
        gasto_total_origem * (pct_cobertura / 100) - gasto_projetado, 2
    )

    delta_sla_pond = round(sla_destino_pond - sla_origem, 2)
    delta_tm_pond  = round(tm_destino_pond  - tm_origem, 2)
    delta_nfd_pond = round(nfd_destino_pond - nfd_origem, 2)

    # ====================================================
    # AVISO SEM COBERTURA
    # ====================================================

    if pct_sem_cobertura > 0:
        st.error(
            f"⚠️ **{pct_sem_cobertura:.1f}% dos pedidos ({pedidos_sem_cobertura:,} pedidos) "
            f"NÃO possuem cobertura tarifária em {transportadora_destino}** para o código "
            f"**{codigo_origem}**. Esses pedidos não poderiam ser migrados."
        )
    else:
        st.success(
            f"✅ Cobertura total: **100%** dos pedidos de **{codigo_origem}** têm equivalência "
            f"em {transportadora_destino}."
        )

    # ====================================================
    # KPI ROW — COBERTURA
    # ====================================================

    st.markdown("### 📊 Visão Geral de Cobertura")

    k1, k2, k3, k4 = st.columns(4)

    with k1:
        st.metric("📦 Total de Pedidos (Origem)", f"{total_pedidos_origem:,}")
    with k2:
        st.metric(
            "✅ Pedidos com Cobertura",
            f"{pedidos_com_cobertura:,}",
            delta=f"{pct_cobertura:.1f}% dos pedidos"
        )
    with k3:
        cor_sem = "normal" if pedidos_sem_cobertura == 0 else "inverse"
        st.metric(
            "❌ Sem Cobertura",
            f"{pedidos_sem_cobertura:,}",
            delta=f"{pct_sem_cobertura:.1f}% sem match",
            delta_color=cor_sem
        )
    with k4:
        st.metric("🔢 Códigos Destino Mapeados", len(df_comp))

    st.divider()

    # ====================================================
    # KPI ROW — COMPARATIVO CONSOLIDADO
    # ====================================================

    st.markdown("### 📌 Comparativo Consolidado (Ponderado)")

    ka, kb, kc, kd, ke, kf = st.columns(6)

    icon_sla, cor_sla = delta_icon(delta_sla_pond, positivo_bom=True)
    icon_tm,  cor_tm  = delta_icon(delta_tm_pond,  positivo_bom=False)
    icon_nfd, cor_nfd = delta_icon(delta_nfd_pond, positivo_bom=False)

    with ka:
        st.metric("SLA Origem",  fmt_pct(sla_origem))
    with kb:
        st.metric(
            f"SLA {transportadora_destino}",
            fmt_pct(sla_destino_pond),
            delta=f"{delta_sla_pond:+.2f}pp"
        )
    with kc:
        st.metric("TM Origem",  fmt_brl(tm_origem))
    with kd:
        st.metric(
            f"TM {transportadora_destino}",
            fmt_brl(tm_destino_pond),
            delta=f"R$ {delta_tm_pond:+,.2f}",
            delta_color="inverse"
        )
    with ke:
        st.metric("NFD Origem",  fmt_pct(nfd_origem))
    with kf:
        st.metric(
            f"NFD {transportadora_destino}",
            fmt_pct(nfd_destino_pond),
            delta=f"{delta_nfd_pond:+.2f}pp",
            delta_color="inverse"
        )

    st.divider()

    # ====================================================
    # GRÁFICO DE BARRAS AGRUPADAS — MÉTRICA × CÓDIGO DESTINO
    # ====================================================

    st.markdown("### 📈 Comparativo por Código Tarifário Destino")

    metrica_selecionada = st.radio(
        "Selecione a métrica para comparar:",
        ["SLA (%)", "Ticket Médio (R$)", "NFD (%)"],
        horizontal=True
    )

    if metrica_selecionada == "SLA (%)":
        col_dest   = "SLA_Destino"
        val_origem = sla_origem
        titulo_y   = "SLA (%)"
        fmt        = lambda v: f"{v:.2f}%"
    elif metrica_selecionada == "Ticket Médio (R$)":
        col_dest   = "TM_Destino"
        val_origem = tm_origem
        titulo_y   = "Ticket Médio (R$)"
        fmt        = lambda v: f"R$ {v:,.2f}"
    else:
        col_dest   = "NFD_Destino"
        val_origem = nfd_origem
        titulo_y   = "NFD (%)"
        fmt        = lambda v: f"{v:.2f}%"

    codigos_dest_lista = df_comp["CodigoDestino"].tolist()
    vals_destino       = df_comp[col_dest].tolist()
    percentuais        = df_comp["Percentual"].tolist()

    fig = go.Figure()

    # Barra Origem (referência — linha repetida por código)
    fig.add_trace(go.Bar(
        name=f"{transportadora_origem} / {codigo_origem}",
        x=codigos_dest_lista,
        y=[val_origem] * len(codigos_dest_lista),
        marker_color=COR_ORIGEM,
        text=[fmt(val_origem)] * len(codigos_dest_lista),
        textposition="outside",
        opacity=0.75
    ))

    # Barra Destino
    cores_destino = [
        COR_OK if (
            (col_dest == "TM_Destino" and v <= val_origem) or
            (col_dest != "TM_Destino" and v >= val_origem)
        ) else COR_ALERTA
        for v in vals_destino
    ]

    fig.add_trace(go.Bar(
        name=f"{transportadora_destino} (simulado)",
        x=codigos_dest_lista,
        y=vals_destino,
        marker_color=cores_destino,
        text=[
            f"{fmt(v)}<br>({p:.1f}% dos pedidos)"
            for v, p in zip(vals_destino, percentuais)
        ],
        textposition="outside",
    ))

    fig.update_layout(
        barmode="group",
        yaxis_title=titulo_y,
        xaxis_title="Código Tarifário Destino",
        legend=dict(orientation="h", y=-0.25),
        height=480,
        plot_bgcolor="#f4f6f9",
        paper_bgcolor="#f4f6f9",
        font=dict(family="Arial", size=13),
        margin=dict(t=30, b=80),
        uniformtext_minsize=10
    )

    # Linha de referência origem
    fig.add_hline(
        y=val_origem,
        line_dash="dash",
        line_color=COR_NEUTRO,
        annotation_text=f"Origem: {fmt(val_origem)}",
        annotation_position="top left"
    )

    st.plotly_chart(fig, use_container_width=True)

    st.caption(
        "🟢 Verde = melhora em relação à origem &nbsp;&nbsp; "
        "🔴 Vermelho = piora em relação à origem &nbsp;&nbsp; "
        "Laranja tracejado = valor de referência da transportadora atual"
    )

    st.divider()

    # ====================================================
    # TABELA DETALHADA COM DELTAS
    # ====================================================

    st.markdown("### 🔄 Redistribuição Operacional — Detalhe por Código Destino")

    tabela = df_comp[[
        "CodigoDestino",
        "Percentual",
        "PedidosSimulados",
        "TM_Destino",
        "Delta_TM",
        "SLA_Destino",
        "Delta_SLA",
        "NFD_Destino",
        "Delta_NFD",
        "FreteProjetado"
    ]].copy()

    tabela = tabela.rename(columns={
        "CodigoDestino":    "Código Destino",
        "Percentual":       "Cobertura %",
        "PedidosSimulados": "Pedidos Simulados",
        "TM_Destino":       "TM Destino (R$)",
        "Delta_TM":         "Δ TM (R$)",
        "SLA_Destino":      "SLA Destino %",
        "Delta_SLA":        "Δ SLA (pp)",
        "NFD_Destino":      "NFD Destino %",
        "Delta_NFD":        "Δ NFD (pp)",
        "FreteProjetado":   "Frete Projetado (R$)"
    })

    for c in ["TM Destino (R$)", "Δ TM (R$)", "Frete Projetado (R$)"]:
        tabela[c] = tabela[c].round(2)
    for c in ["Cobertura %", "SLA Destino %", "Δ SLA (pp)", "NFD Destino %", "Δ NFD (pp)"]:
        tabela[c] = tabela[c].round(2)

    # Coloração condicional nos deltas
    def color_delta_sla(val):
        if val > 0:  return "background-color: #d4edda; color: #155724"
        if val < 0:  return "background-color: #f8d7da; color: #721c24"
        return ""

    def color_delta_tm(val):
        if val < 0:  return "background-color: #d4edda; color: #155724"
        if val > 0:  return "background-color: #f8d7da; color: #721c24"
        return ""

    def color_delta_nfd(val):
        if val < 0:  return "background-color: #d4edda; color: #155724"
        if val > 0:  return "background-color: #f8d7da; color: #721c24"
        return ""

    styled = (
        tabela.style
        .applymap(color_delta_sla, subset=["Δ SLA (pp)"])
        .applymap(color_delta_tm,  subset=["Δ TM (R$)"])
        .applymap(color_delta_nfd, subset=["Δ NFD (pp)"])
        .format({
            "TM Destino (R$)":      "R$ {:,.2f}",
            "Δ TM (R$)":            "{:+,.2f}",
            "SLA Destino %":        "{:.2f}%",
            "Δ SLA (pp)":           "{:+.2f}",
            "NFD Destino %":        "{:.2f}%",
            "Δ NFD (pp)":           "{:+.2f}",
            "Cobertura %":          "{:.2f}%",
            "Frete Projetado (R$)": "R$ {:,.2f}",
        })
    )

    st.dataframe(styled, use_container_width=True, hide_index=True)

    st.divider()

    # ====================================================
    # PROJEÇÃO FINANCEIRA
    # ====================================================

    st.markdown("### 💰 Projeção Financeira (sobre pedidos com cobertura)")

    f1, f2, f3 = st.columns(3)

    with f1:
        gasto_referencia = round(gasto_total_origem * (pct_cobertura / 100), 2)
        st.metric(
            "Gasto Atual (pedidos com cobertura)",
            fmt_brl(gasto_referencia)
        )
    with f2:
        st.metric(
            f"Gasto Projetado ({transportadora_destino})",
            fmt_brl(gasto_projetado),
        )
    with f3:
        st.metric(
            "Economia / Custo Projetado",
            fmt_brl(abs(economia_projetada)),
            delta=f"{'Economia' if economia_projetada >= 0 else 'Custo adicional'}: {fmt_brl(economia_projetada)}",
            delta_color="normal" if economia_projetada >= 0 else "inverse"
        )

    st.divider()

    # ====================================================
    # RESUMO EXECUTIVO
    # ====================================================

    st.markdown("### 📋 Resumo Executivo")

    sla_txt  = f"{'melhora' if delta_sla_pond >= 0 else 'piora'} de {abs(delta_sla_pond):.2f}pp"
    tm_txt   = f"{'redução' if delta_tm_pond <= 0 else 'aumento'} de R$ {abs(delta_tm_pond):,.2f} por pedido"
    nfd_txt  = f"{'redução' if delta_nfd_pond <= 0 else 'aumento'} de {abs(delta_nfd_pond):.2f}pp"
    eco_txt  = f"{'economia' if economia_projetada >= 0 else 'custo adicional'} de {fmt_brl(abs(economia_projetada))}"

    st.info(f"""
**Cenário simulado:** migrar os pedidos do código **{codigo_origem}** ({transportadora_origem}) para **{transportadora_destino}**
no período de **{periodo}**.

📦 **{total_pedidos_origem:,}** pedidos analisados &nbsp;|&nbsp; \
✅ **{pedidos_com_cobertura:,}** com cobertura ({pct_cobertura:.1f}%) &nbsp;|&nbsp; \
❌ **{pedidos_sem_cobertura:,}** sem cobertura ({pct_sem_cobertura:.1f}%)

Os **{pedidos_com_cobertura:,} pedidos** com cobertura seriam redistribuídos entre \
**{len(df_comp)} códigos tarifários** de {transportadora_destino}, com os seguintes impactos estimados:

- 🎯 **SLA:** {sla_txt} (de {sla_origem:.2f}% → {sla_destino_pond:.2f}%)
- 💲 **Ticket Médio:** {tm_txt} (de {fmt_brl(tm_origem)} → {fmt_brl(tm_destino_pond)})
- 📄 **NFD:** {nfd_txt} (de {nfd_origem:.2f}% → {nfd_destino_pond:.2f}%)
- 💰 **Impacto financeiro:** {eco_txt} sobre os pedidos migráveis
""")