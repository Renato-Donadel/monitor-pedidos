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
    if not pd.notna(v):
        return "—"
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

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
    # FILTROS
    # ====================================================

    col1, col2, col3 = st.columns(3)

    with col1:
        transportadoras_disponiveis = sorted(
            df_sim_base["TransportadoraOrigem"].dropna().unique()
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

    # Filtro de período na base de pedidos
    df_trade["DataFinal"] = pd.to_datetime(df_trade["DataFinal"], errors="coerce")
    data_corte = pd.Timestamp.today() - pd.Timedelta(days=dias)
    df_trade_periodo = df_trade[df_trade["DataFinal"] >= data_corte].copy()

    # ====================================================
    # CÓDIGO TARIFÁRIO
    # ====================================================

    codigos_origem = sorted(
        df_sim_base[
            (df_sim_base["TransportadoraOrigem"] == transportadora_origem) &
            (df_sim_base["TransportadoraDestino"] == transportadora_destino)
        ]["CodigoOrigem"].dropna().unique()
    )

    if not codigos_origem:
        st.warning("Nenhum código tarifário encontrado para essa combinação de transportadoras.")
        return

    codigo_origem = st.selectbox("🏷️ Código Tarifário de Origem", codigos_origem)

    st.divider()

    # ====================================================
    # FILTRAR SIMILARIDADE — já tem tudo calculado pelo ETL
    # ====================================================

    df_comp = df_sim_base[
        (df_sim_base["TransportadoraOrigem"] == transportadora_origem) &
        (df_sim_base["CodigoOrigem"] == codigo_origem) &
        (df_sim_base["TransportadoraDestino"] == transportadora_destino)
    ].copy()

    if df_comp.empty:
        st.warning(
            f"⚠️ Nenhuma similaridade encontrada entre **{transportadora_origem} / {codigo_origem}** "
            f"e **{transportadora_destino}**."
        )
        return

    # ====================================================
    # MÉTRICAS DE ORIGEM — da base de pedidos (período filtrado)
    # ====================================================

    df_origem_pedidos = df_trade_periodo[
        (df_trade_periodo["Transportadora"] == transportadora_origem) &
        (df_trade_periodo["CodigoTarifario"] == codigo_origem)
    ].copy()

    total_pedidos_origem = len(df_origem_pedidos)

    if total_pedidos_origem > 0:
        sla_origem = round(df_origem_pedidos["DentroPrazo"].mean() * 100, 2)
        tm_origem  = round(df_origem_pedidos["ValorFrete"].astype(float).mean(), 2)
        nfd_origem = round(df_origem_pedidos["TemNFD"].mean() * 100, 2)
        gasto_total_origem = round(df_origem_pedidos["ValorFrete"].astype(float).sum(), 2)
    else:
        # Fallback: usa o que está na base de similaridade
        sla_origem = round(df_comp["SLAOrigem"].iloc[0] * 100, 2)
        nfd_origem = round(df_comp["NFDOrigem"].iloc[0] * 100, 2)
        tm_origem  = round(
            (df_comp["ValorFreteOrigem"].sum() / df_comp["Pedidos"].sum()), 2
        ) if df_comp["Pedidos"].sum() > 0 else 0
        total_pedidos_origem = int(df_comp["Pedidos"].sum())
        gasto_total_origem   = round(df_comp["ValorFreteOrigem"].sum(), 2)

    # ====================================================
    # COBERTURA
    # ====================================================

    # Percentual já calculado no ETL — soma dos percentuais por código destino
    # Cada linha é um código destino com seu % de cobertura dos CEPs de origem
    pct_cobertura_total = min(df_comp["Percentual"].sum(), 100.0)
    pct_sem_cobertura   = round(100.0 - pct_cobertura_total, 2)

    pedidos_com_cobertura = int(
        total_pedidos_origem * (pct_cobertura_total / 100)
    )
    pedidos_sem_cobertura = total_pedidos_origem - pedidos_com_cobertura

    # ====================================================
    # PEDIDOS SIMULADOS POR CÓDIGO DESTINO
    # ====================================================

    df_comp["PedidosSimulados"] = (
        total_pedidos_origem * (df_comp["Percentual"] / 100)
    ).round(0).astype(int)

    # ====================================================
    # MÉTRICAS DESTINO (SLADestino, NFDDestino, TM_Destino)
    # já vêm do ETL — só converte escala onde necessário
    # SLAOrigem/NFDOrigem estão em 0-1, SLADestino/NFDDestino também
    # ====================================================

    df_comp["SLA_Dest_pct"] = (df_comp["SLADestino"].fillna(0) * 100).round(2)
    df_comp["NFD_Dest_pct"] = (df_comp["NFDDestino"].fillna(0) * 100).round(2)
    df_comp["TM_Dest"]      = df_comp["TM_Destino"].fillna(0).round(2)

    # Projeção financeira
    df_comp["FreteProjetado"] = (df_comp["TM_Dest"] * df_comp["PedidosSimulados"]).round(2)

    # Deltas por linha
    df_comp["Delta_SLA"] = (df_comp["SLA_Dest_pct"] - sla_origem).round(2)
    df_comp["Delta_TM"]  = (df_comp["TM_Dest"]      - tm_origem).round(2)
    df_comp["Delta_NFD"] = (df_comp["NFD_Dest_pct"]  - nfd_origem).round(2)

    # ====================================================
    # KPIs CONSOLIDADOS PONDERADOS
    # ====================================================

    total_simulados = df_comp["PedidosSimulados"].sum()

    def pond(col):
        if total_simulados == 0:
            return 0.0
        return (df_comp[col] * df_comp["PedidosSimulados"]).sum() / total_simulados

    sla_destino_pond = round(pond("SLA_Dest_pct"), 2)
    nfd_destino_pond = round(pond("NFD_Dest_pct"), 2)
    tm_destino_pond  = round(pond("TM_Dest"), 2)

    gasto_projetado    = round(df_comp["FreteProjetado"].sum(), 2)
    gasto_referencia   = round(gasto_total_origem * (pct_cobertura_total / 100), 2)
    economia_projetada = round(gasto_referencia - gasto_projetado, 2)

    delta_sla_pond = round(sla_destino_pond - sla_origem, 2)
    delta_tm_pond  = round(tm_destino_pond  - tm_origem, 2)
    delta_nfd_pond = round(nfd_destino_pond - nfd_origem, 2)

    # ====================================================
    # AVISO SEM COBERTURA
    # ====================================================

    if pct_sem_cobertura > 0.5:
        st.error(
            f"⚠️ **{pct_sem_cobertura:.1f}% dos pedidos (~{pedidos_sem_cobertura:,} pedidos) "
            f"NÃO possuem cobertura tarifária em {transportadora_destino}** para o código "
            f"**{codigo_origem}**. Esses pedidos não poderiam ser migrados."
        )
    else:
        st.success(
            f"✅ Cobertura total: **100%** dos pedidos de **{codigo_origem}** têm equivalência "
            f"em {transportadora_destino}."
        )

    # ====================================================
    # KPI — COBERTURA
    # ====================================================

    st.markdown("### 📊 Visão Geral de Cobertura")

    k1, k2, k3, k4 = st.columns(4)

    with k1:
        st.metric("📦 Total de Pedidos (Origem)", f"{total_pedidos_origem:,}")
    with k2:
        st.metric(
            "✅ Com Cobertura",
            f"{pedidos_com_cobertura:,}",
            delta=f"{pct_cobertura_total:.1f}% dos pedidos"
        )
    with k3:
        st.metric(
            "❌ Sem Cobertura",
            f"{pedidos_sem_cobertura:,}",
            delta=f"{pct_sem_cobertura:.1f}% sem match",
            delta_color="inverse" if pedidos_sem_cobertura > 0 else "normal"
        )
    with k4:
        st.metric("🔢 Códigos Destino Mapeados", len(df_comp))

    st.divider()

    # ====================================================
    # KPI — COMPARATIVO CONSOLIDADO
    # ====================================================

    st.markdown("### 📌 Comparativo Consolidado (Ponderado pelos pedidos simulados)")

    ka, kb, kc, kd, ke, kf = st.columns(6)

    with ka:
        st.metric("SLA Origem", fmt_pct(sla_origem))
    with kb:
        st.metric(
            f"SLA {transportadora_destino}",
            fmt_pct(sla_destino_pond),
            delta=f"{delta_sla_pond:+.2f}pp"
        )
    with kc:
        st.metric("TM Origem", fmt_brl(tm_origem))
    with kd:
        st.metric(
            f"TM {transportadora_destino}",
            fmt_brl(tm_destino_pond),
            delta=f"R$ {delta_tm_pond:+,.2f}",
            delta_color="inverse"
        )
    with ke:
        st.metric("NFD Origem", fmt_pct(nfd_origem))
    with kf:
        st.metric(
            f"NFD {transportadora_destino}",
            fmt_pct(nfd_destino_pond),
            delta=f"{delta_nfd_pond:+.2f}pp",
            delta_color="inverse"
        )

    st.divider()

    # ====================================================
    # GRÁFICO DE BARRAS AGRUPADAS
    # ====================================================

    st.markdown("### 📈 Comparativo por Código Tarifário Destino")

    metrica_sel = st.radio(
        "Selecione a métrica:",
        ["SLA (%)", "Ticket Médio (R$)", "NFD (%)"],
        horizontal=True
    )

    if metrica_sel == "SLA (%)":
        col_dest   = "SLA_Dest_pct"
        val_ref    = sla_origem
        titulo_y   = "SLA (%)"
        fmt_v      = lambda v: f"{v:.2f}%"
        melhor_alto = True
    elif metrica_sel == "Ticket Médio (R$)":
        col_dest   = "TM_Dest"
        val_ref    = tm_origem
        titulo_y   = "Ticket Médio (R$)"
        fmt_v      = lambda v: f"R$ {v:,.2f}"
        melhor_alto = False
    else:
        col_dest   = "NFD_Dest_pct"
        val_ref    = nfd_origem
        titulo_y   = "NFD (%)"
        fmt_v      = lambda v: f"{v:.2f}%"
        melhor_alto = False

    codigos_x    = df_comp["CodigoDestino"].tolist()
    vals_destino = df_comp[col_dest].tolist()
    percentuais  = df_comp["Percentual"].tolist()

    # Cor por código: verde se melhora, vermelho se piora
    cores_dest = []
    for v in vals_destino:
        if melhor_alto:
            cores_dest.append(COR_OK if v >= val_ref else COR_ALERTA)
        else:
            cores_dest.append(COR_OK if v <= val_ref else COR_ALERTA)

    fig = go.Figure()

    # Barra referência (origem)
    fig.add_trace(go.Bar(
        name=f"{transportadora_origem} — {codigo_origem} (atual)",
        x=codigos_x,
        y=[val_ref] * len(codigos_x),
        marker_color=COR_ORIGEM,
        opacity=0.7,
        text=[fmt_v(val_ref)] * len(codigos_x),
        textposition="outside",
    ))

    # Barra destino
    fig.add_trace(go.Bar(
        name=f"{transportadora_destino} (simulado)",
        x=codigos_x,
        y=vals_destino,
        marker_color=cores_dest,
        text=[
            f"{fmt_v(v)}<br>({p:.1f}% dos ped.)"
            for v, p in zip(vals_destino, percentuais)
        ],
        textposition="outside",
    ))

    fig.update_layout(
        barmode="group",
        yaxis_title=titulo_y,
        xaxis_title="Código Tarifário Destino",
        legend=dict(orientation="h", y=-0.28),
        height=500,
        plot_bgcolor="#f4f6f9",
        paper_bgcolor="#f4f6f9",
        font=dict(family="Arial", size=13),
        margin=dict(t=40, b=90),
    )

    fig.add_hline(
        y=val_ref,
        line_dash="dash",
        line_color=COR_NEUTRO,
        annotation_text=f"Referência origem: {fmt_v(val_ref)}",
        annotation_position="top left"
    )

    st.plotly_chart(fig, use_container_width=True)
    st.caption(
        "🟢 Verde = melhora vs origem  |  "
        "🔴 Vermelho = piora vs origem  |  "
        "Tracejado laranja = valor de referência da transportadora atual"
    )

    st.divider()

    # ====================================================
    # TABELA DETALHADA
    # ====================================================

    st.markdown("### 🔄 Redistribuição Operacional — Detalhe por Código Destino")

    tabela = df_comp[[
        "CodigoDestino",
        "Percentual",
        "PedidosSimulados",
        "TM_Dest",
        "Delta_TM",
        "SLA_Dest_pct",
        "Delta_SLA",
        "NFD_Dest_pct",
        "Delta_NFD",
        "FreteProjetado"
    ]].copy().rename(columns={
        "CodigoDestino":  "Código Destino",
        "Percentual":     "Cobertura %",
        "PedidosSimulados": "Pedidos Simulados",
        "TM_Dest":        "TM Destino (R$)",
        "Delta_TM":       "Δ TM (R$)",
        "SLA_Dest_pct":   "SLA Destino %",
        "Delta_SLA":      "Δ SLA (pp)",
        "NFD_Dest_pct":   "NFD Destino %",
        "Delta_NFD":      "Δ NFD (pp)",
        "FreteProjetado": "Frete Projetado (R$)"
    })

    def color_sla(v):
        if v > 0:  return "background-color:#d4edda;color:#155724"
        if v < 0:  return "background-color:#f8d7da;color:#721c24"
        return ""

    def color_tm(v):
        if v < 0:  return "background-color:#d4edda;color:#155724"
        if v > 0:  return "background-color:#f8d7da;color:#721c24"
        return ""

    def color_nfd(v):
        if v < 0:  return "background-color:#d4edda;color:#155724"
        if v > 0:  return "background-color:#f8d7da;color:#721c24"
        return ""

    styled = (
        tabela.style
        .map(color_sla, subset=["Δ SLA (pp)"])
        .map(color_tm,  subset=["Δ TM (R$)"])
        .map(color_nfd, subset=["Δ NFD (pp)"])
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

    st.markdown("### 💰 Projeção Financeira (pedidos com cobertura)")

    f1, f2, f3 = st.columns(3)

    with f1:
        st.metric("Gasto Atual (c/ cobertura)", fmt_brl(gasto_referencia))
    with f2:
        st.metric(f"Gasto Projetado ({transportadora_destino})", fmt_brl(gasto_projetado))
    with f3:
        label = "Economia Projetada" if economia_projetada >= 0 else "Custo Adicional"
        st.metric(
            label,
            fmt_brl(abs(economia_projetada)),
            delta=f"{fmt_brl(economia_projetada)}",
            delta_color="normal" if economia_projetada >= 0 else "inverse"
        )

    st.divider()

    # ====================================================
    # RESUMO EXECUTIVO
    # ====================================================

    st.markdown("### 📋 Resumo Executivo")

    sla_dir  = "melhora" if delta_sla_pond >= 0 else "piora"
    tm_dir   = "redução" if delta_tm_pond  <= 0 else "aumento"
    nfd_dir  = "redução" if delta_nfd_pond <= 0 else "aumento"
    eco_dir  = "economia" if economia_projetada >= 0 else "custo adicional"

    st.info(f"""
**Cenário simulado:** migrar os pedidos do código **{codigo_origem}** ({transportadora_origem}) \
para **{transportadora_destino}** no período de **{periodo}**.

📦 **{total_pedidos_origem:,}** pedidos analisados &nbsp;|&nbsp; \
✅ **{pedidos_com_cobertura:,}** com cobertura ({pct_cobertura_total:.1f}%) &nbsp;|&nbsp; \
❌ **{pedidos_sem_cobertura:,}** sem cobertura ({pct_sem_cobertura:.1f}%)

Os **{pedidos_com_cobertura:,}** pedidos migráveis seriam redistribuídos entre \
**{len(df_comp)} códigos tarifários** de {transportadora_destino}, com os seguintes impactos estimados:

- 🎯 **SLA:** {sla_dir} de **{abs(delta_sla_pond):.2f}pp** \
({sla_origem:.2f}% → {sla_destino_pond:.2f}%)
- 💲 **Ticket Médio:** {tm_dir} de **{fmt_brl(abs(delta_tm_pond))}** por pedido \
({fmt_brl(tm_origem)} → {fmt_brl(tm_destino_pond)})
- 📄 **NFD:** {nfd_dir} de **{abs(delta_nfd_pond):.2f}pp** \
({nfd_origem:.2f}% → {nfd_destino_pond:.2f}%)
- 💰 **Impacto financeiro:** {eco_dir} de **{fmt_brl(abs(economia_projetada))}** \
sobre os pedidos migráveis
""")
