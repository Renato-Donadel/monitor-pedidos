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
    # RANKING NFD — PIORES CÓDIGOS TARIFÁRIOS
    # ====================================================

    st.markdown("### 📊 Ranking de NFD por Código Tarifário")
    st.caption("Códigos com maior percentual de devolução — quanto mais alto, pior.")

    if not df_trade.empty and "TemNFD" in df_trade.columns:

        nfd_ranking = (
            df_trade.groupby(["Transportadora", "CodigoTarifario"])
            .agg(
                Pedidos=("TemNFD", "count"),
                NFD=("TemNFD", "sum")
            )
            .reset_index()
        )

        nfd_ranking = nfd_ranking[nfd_ranking["Pedidos"] >= 30]
        nfd_ranking["NFD%"] = (nfd_ranking["NFD"] / nfd_ranking["Pedidos"] * 100).round(2)
        nfd_ranking = nfd_ranking.sort_values("NFD%", ascending=False).reset_index(drop=True)

        if "pagina_nfd" not in st.session_state:
            st.session_state["pagina_nfd"] = 0

        por_pagina = 10
        inicio = st.session_state["pagina_nfd"] * por_pagina
        fim    = inicio + por_pagina
        total_paginas = (len(nfd_ranking) - 1) // por_pagina

        pagina_df = nfd_ranking.iloc[inicio:fim].copy()
        pagina_df.index = range(inicio + 1, inicio + len(pagina_df) + 1)

        fig_rank = go.Figure(go.Bar(
            x=pagina_df["NFD%"],
            y=pagina_df["Transportadora"] + " / " + pagina_df["CodigoTarifario"],
            orientation="h",
            marker_color=[
                COR_ALERTA if v >= 5 else COR_NEUTRO if v >= 2 else COR_OK
                for v in pagina_df["NFD%"]
            ],
            text=[f"{v:.2f}% ({p:,} ped.)" for v, p in zip(pagina_df["NFD%"], pagina_df["Pedidos"])],
            textposition="outside",
        ))

        fig_rank.update_layout(
            height=max(300, len(pagina_df) * 42),
            xaxis_title="NFD (%)",
            yaxis=dict(autorange="reversed", tickfont=dict(size=12)),
            plot_bgcolor="#f4f6f9",
            paper_bgcolor="#f4f6f9",
            margin=dict(l=10, r=80, t=20, b=40),
            font=dict(family="Arial", size=12),
        )

        st.plotly_chart(fig_rank, use_container_width=True)

        st.caption(
            "🔴 NFD ≥ 5%  |  🟠 NFD entre 2% e 5%  |  🟢 NFD < 2%  |  Mínimo 30 pedidos para aparecer no ranking"
        )

        nav1, nav2, nav3 = st.columns([1, 6, 1])

        with nav1:
            if st.session_state["pagina_nfd"] > 0:
                if st.button("◀ Anteriores", key="nfd_prev"):
                    st.session_state["pagina_nfd"] -= 1
                    st.rerun()

        with nav2:
            st.caption(
                f"Mostrando {inicio+1}–{min(fim, len(nfd_ranking))} de {len(nfd_ranking)} códigos  |  Página {st.session_state['pagina_nfd']+1} de {total_paginas+1}"
            )

        with nav3:
            if st.session_state["pagina_nfd"] < total_paginas:
                if st.button("Próximos ▶", key="nfd_next"):
                    st.session_state["pagina_nfd"] += 1
                    st.rerun()

    else:
        st.warning("Base de pedidos não encontrada ou sem coluna TemNFD.")

    st.divider()

    # ====================================================
    # SIMULAÇÃO — IMPACTO DE BLOQUEAR OS 10 PIORES CÓDIGOS
    # ====================================================

    st.markdown("### 🚫 Simulação: Impacto de Bloquear os 10 Piores Códigos Tarifários")
    st.caption(
        "Simulação retroativa nos últimos 90 dias — como teriam ficado SLA, TM, NFD e valor total "
        "se tivéssemos migrado os pedidos dos 10 piores códigos para as transportadoras alternativas."
    )

    if not df_trade.empty and not df_sim_base.empty and "TemNFD" in df_trade.columns:

        # Pega os 10 piores códigos (mínimo 30 pedidos)
        nfd_full = (
            df_trade.groupby(["Transportadora", "CodigoTarifario"])
            .agg(Pedidos=("TemNFD", "count"), NFD=("TemNFD", "sum"))
            .reset_index()
        )
        nfd_full = nfd_full[nfd_full["Pedidos"] >= 30].copy()
        nfd_full["NFD%"] = (nfd_full["NFD"] / nfd_full["Pedidos"] * 100).round(2)
        top10 = nfd_full.sort_values("NFD%", ascending=False).head(10)

        # Pedidos dos 10 piores códigos
        mask_piores = (
            df_trade[["Transportadora", "CodigoTarifario"]]
            .apply(tuple, axis=1)
            .isin(top10[["Transportadora", "CodigoTarifario"]].apply(tuple, axis=1))
        )
        df_piores = df_trade[mask_piores].copy()
        df_resto  = df_trade[~mask_piores].copy()

        # Para cada pedido dos piores, busca o destino ponderado no df_sim_base
        # (usa o código destino com maior Percentual para cada par origem/código)
        destino_map = (
            df_sim_base.sort_values("Percentual", ascending=False)
            .drop_duplicates(subset=["TransportadoraOrigem", "CodigoOrigem"])
            [[
                "TransportadoraOrigem", "CodigoOrigem",
                "TransportadoraDestino", "CodigoDestino",
                "SLADestino", "NFDDestino", "TM_Destino",
                "TM_Cotacao_Destino", "Pedidos_Com_Cotacao_Destino"
            ]]
            .copy()
        )

        df_piores_sim = df_piores.merge(
            destino_map,
            left_on=["Transportadora", "CodigoTarifario"],
            right_on=["TransportadoraOrigem", "CodigoOrigem"],
            how="left"
        )

        tem_destino    = df_piores_sim["TransportadoraDestino"].notna()
        df_migrados    = df_piores_sim[tem_destino].copy()
        df_sem_destino = df_piores_sim[~tem_destino].copy()

        # ── KPIs ANTES ──────────────────────────────────────────
        total_antes       = len(df_trade)
        sla_antes         = df_trade["DentroPrazo"].mean() * 100 if "DentroPrazo" in df_trade.columns else 0
        nfd_antes         = df_trade["TemNFD"].mean() * 100
        tm_antes          = df_trade["ValorFrete"].astype(float).mean() if "ValorFrete" in df_trade.columns else 0
        total_frete_antes = df_trade["ValorFrete"].astype(float).sum() if "ValorFrete" in df_trade.columns else 0

        # ── KPIs DEPOIS — TM HISTÓRICO ──────────────────────────
        # Pedidos migrados: usa SLA/NFD/TM do destino histórico
        # Pedidos sem destino mapeado: mantém como está (não bloqueia)
        # Pedidos fora dos piores: mantém como está

        sla_mig_hist  = df_migrados["SLADestino"].fillna(df_migrados["DentroPrazo"]).mean() * 100 if len(df_migrados) > 0 else 0
        nfd_mig_hist  = df_migrados["NFDDestino"].fillna(df_migrados["TemNFD"]).mean() * 100 if len(df_migrados) > 0 else 0
        tm_mig_hist   = df_migrados["TM_Destino"].fillna(df_migrados["ValorFrete"].astype(float)).mean() if len(df_migrados) > 0 else 0
        frete_mig_hist = (df_migrados["TM_Destino"].fillna(df_migrados["ValorFrete"].astype(float)) * 1).sum()

        n_mig    = len(df_migrados)
        n_resto  = len(df_resto) + len(df_sem_destino)

        sla_depois_hist = (
            (sla_mig_hist * n_mig + df_resto["DentroPrazo"].mean() * 100 * len(df_resto) if "DentroPrazo" in df_resto.columns else 0)
            / total_antes if total_antes > 0 else 0
        )
        sla_depois_hist = (
            sla_mig_hist * n_mig +
            (df_resto["DentroPrazo"].mean() * 100 * len(df_resto) if "DentroPrazo" in df_resto.columns else 0) +
            (df_sem_destino["DentroPrazo"].mean() * 100 * len(df_sem_destino) if "DentroPrazo" in df_sem_destino.columns else 0)
        ) / total_antes

        nfd_depois_hist = (
            nfd_mig_hist * n_mig +
            df_resto["TemNFD"].mean() * 100 * len(df_resto) +
            df_sem_destino["TemNFD"].mean() * 100 * len(df_sem_destino)
        ) / total_antes

        tm_depois_hist = (
            tm_mig_hist * n_mig +
            df_resto["ValorFrete"].astype(float).sum() +
            df_sem_destino["ValorFrete"].astype(float).sum()
        ) / total_antes if "ValorFrete" in df_resto.columns else 0

        total_frete_depois_hist = (
            frete_mig_hist +
            df_resto["ValorFrete"].astype(float).sum() +
            df_sem_destino["ValorFrete"].astype(float).sum()
        )

        # ── KPIs DEPOIS — TM COTAÇÃO ────────────────────────────
        tm_mig_cot = df_migrados["TM_Cotacao_Destino"].fillna(df_migrados["TM_Destino"].fillna(df_migrados["ValorFrete"].astype(float))).mean() if len(df_migrados) > 0 else 0
        frete_mig_cot = df_migrados["TM_Cotacao_Destino"].fillna(df_migrados["TM_Destino"].fillna(df_migrados["ValorFrete"].astype(float))).sum()

        tm_depois_cot = (
            tm_mig_cot * n_mig +
            df_resto["ValorFrete"].astype(float).sum() +
            df_sem_destino["ValorFrete"].astype(float).sum()
        ) / total_antes if "ValorFrete" in df_resto.columns else 0

        total_frete_depois_cot = (
            frete_mig_cot +
            df_resto["ValorFrete"].astype(float).sum() +
            df_sem_destino["ValorFrete"].astype(float).sum()
        )

        # ── RESUMO KPIs ─────────────────────────────────────────
        st.markdown("#### 📌 Impacto Consolidado da Migração")

        k1, k2, k3, k4 = st.columns(4)

        delta_sla = round(sla_depois_hist - sla_antes, 2)
        delta_nfd = round(nfd_depois_hist - nfd_antes, 2)
        delta_tm_hist = round(tm_depois_hist - tm_antes, 2)
        delta_frete_hist = round(total_frete_depois_hist - total_frete_antes, 2)

        with k1:
            st.metric(
                "SLA (antes → depois)",
                fmt_pct(sla_depois_hist),
                delta=f"{delta_sla:+.2f}pp",
                delta_color="normal" if delta_sla >= 0 else "inverse"
            )
        with k2:
            st.metric(
                "NFD (antes → depois)",
                fmt_pct(nfd_depois_hist),
                delta=f"{delta_nfd:+.2f}pp",
                delta_color="inverse" if delta_nfd > 0 else "normal"
            )
        with k3:
            st.metric(
                "TM Médio — Histórico",
                fmt_brl(tm_depois_hist),
                delta=f"R$ {delta_tm_hist:+,.2f}",
                delta_color="inverse" if delta_tm_hist > 0 else "normal"
            )
        with k4:
            st.metric(
                "TM Médio — Cotação",
                fmt_brl(tm_depois_cot),
                delta=fmt_brl(round(tm_depois_cot - tm_antes, 2)),
                delta_color="inverse" if tm_depois_cot > tm_antes else "normal"
            )

        st.caption(
            f"📦 {n_mig:,} pedidos migrados para transportadora alternativa  |  "
            f"{len(df_sem_destino):,} pedidos sem destino mapeado (mantidos)"
        )

        st.divider()

        # ── GRÁFICO 1 — TM HISTÓRICO ────────────────────────────
        st.markdown("#### 📈 Comparativo por KPI — Base Histórica de TM")

        categorias  = ["SLA (%)", "NFD (%)", "TM Médio (R$)", "Frete Total (R$)"]
        vals_antes  = [sla_antes, nfd_antes, tm_antes, total_frete_antes]
        vals_depois = [sla_depois_hist, nfd_depois_hist, tm_depois_hist, total_frete_depois_hist]

        fig_hist = go.Figure()

        fig_hist.add_trace(go.Bar(
            name="Antes do bloqueio",
            x=categorias,
            y=vals_antes,
            marker_color=COR_ORIGEM,
            text=[
                f"{sla_antes:.2f}%", f"{nfd_antes:.2f}%",
                fmt_brl(tm_antes), fmt_brl(total_frete_antes)
            ],
            textposition="outside",
        ))

        fig_hist.add_trace(go.Bar(
            name="Depois do bloqueio (TM histórico)",
            x=categorias,
            y=vals_depois,
            marker_color=[
                COR_OK if sla_depois_hist >= sla_antes else COR_ALERTA,
                COR_OK if nfd_depois_hist <= nfd_antes else COR_ALERTA,
                COR_OK if tm_depois_hist  <= tm_antes  else COR_ALERTA,
                COR_OK if total_frete_depois_hist <= total_frete_antes else COR_ALERTA,
            ],
            text=[
                f"{sla_depois_hist:.2f}%", f"{nfd_depois_hist:.2f}%",
                fmt_brl(tm_depois_hist), fmt_brl(total_frete_depois_hist)
            ],
            textposition="outside",
        ))

        fig_hist.update_layout(
            barmode="group",
            height=380,
            plot_bgcolor="#f4f6f9",
            paper_bgcolor="#f4f6f9",
            font=dict(family="Arial", size=12),
            legend=dict(orientation="h", y=-0.25),
            margin=dict(t=30, b=80, l=10, r=10),
            yaxis=dict(showticklabels=False),
        )

        st.plotly_chart(fig_hist, use_container_width=True)

        # ── GRÁFICO 2 — TM COTAÇÃO ──────────────────────────────
        st.markdown("#### 📈 Comparativo por KPI — Base de Cotação de Leilão")

        vals_depois_cot = [sla_depois_hist, nfd_depois_hist, tm_depois_cot, total_frete_depois_cot]

        fig_cot = go.Figure()

        fig_cot.add_trace(go.Bar(
            name="Antes do bloqueio",
            x=categorias,
            y=vals_antes,
            marker_color=COR_ORIGEM,
            text=[
                f"{sla_antes:.2f}%", f"{nfd_antes:.2f}%",
                fmt_brl(tm_antes), fmt_brl(total_frete_antes)
            ],
            textposition="outside",
        ))

        fig_cot.add_trace(go.Bar(
            name="Depois do bloqueio (TM cotação)",
            x=categorias,
            y=vals_depois_cot,
            marker_color=[
                COR_OK if sla_depois_hist >= sla_antes else COR_ALERTA,
                COR_OK if nfd_depois_hist <= nfd_antes else COR_ALERTA,
                COR_OK if tm_depois_cot   <= tm_antes  else COR_ALERTA,
                COR_OK if total_frete_depois_cot <= total_frete_antes else COR_ALERTA,
            ],
            text=[
                f"{sla_depois_hist:.2f}%", f"{nfd_depois_hist:.2f}%",
                fmt_brl(tm_depois_cot), fmt_brl(total_frete_depois_cot)
            ],
            textposition="outside",
        ))

        fig_cot.update_layout(
            barmode="group",
            height=380,
            plot_bgcolor="#f4f6f9",
            paper_bgcolor="#f4f6f9",
            font=dict(family="Arial", size=12),
            legend=dict(orientation="h", y=-0.25),
            margin=dict(t=30, b=80, l=10, r=10),
            yaxis=dict(showticklabels=False),
        )

        st.plotly_chart(fig_cot, use_container_width=True)

        st.caption(
            "🟢 Verde = melhora vs situação atual  |  "
            "🔴 Vermelho = piora vs situação atual  |  "
            f"SLA e NFD são iguais nos dois gráficos pois dependem do histórico real de entrega"
        )

    else:
        st.warning("Dados insuficientes para simulação.")

    st.divider()

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
    # 3 MINI GRÁFICOS LADO A LADO
    # ====================================================

    st.markdown("### 📈 Comparativo por Código Tarifário Destino")

    codigos_x   = df_comp["CodigoDestino"].tolist()
    percentuais = df_comp["Percentual"].tolist()

    def mini_grafico(col_dest, val_ref, titulo_y, fmt_v, maiores_melhores):
        vals = df_comp[col_dest].tolist()
        cores = [
            COR_OK if (maiores_melhores and v >= val_ref) or
                      (not maiores_melhores and v <= val_ref)
            else COR_ALERTA
            for v in vals
        ]
        fig = go.Figure()
        fig.add_trace(go.Bar(
            name=f"{transportadora_origem} (atual)",
            x=codigos_x,
            y=[val_ref] * len(codigos_x),
            marker_color=COR_ORIGEM,
            opacity=0.75,
            text=[fmt_v(val_ref)] * len(codigos_x),
            textposition="outside",
            textfont=dict(size=11),
        ))
        fig.add_trace(go.Bar(
            name=f"{transportadora_destino} (simulado)",
            x=codigos_x,
            y=vals,
            marker_color=cores,
            text=[
                f"{fmt_v(v)}<br>({p:.1f}%)"
                for v, p in zip(vals, percentuais)
            ],
            textposition="outside",
            textfont=dict(size=11),
        ))
        fig.update_layout(
            barmode="group",
            title=dict(text=titulo_y, font=dict(size=13), x=0),
            yaxis_title=titulo_y,
            xaxis=dict(tickfont=dict(size=11)),
            legend=dict(orientation="h", y=-0.35, font=dict(size=11)),
            height=320,
            plot_bgcolor="#f4f6f9",
            paper_bgcolor="#f4f6f9",
            font=dict(family="Arial", size=12),
            margin=dict(t=40, b=80, l=50, r=10),
        )
        fig.add_hline(
            y=val_ref,
            line_dash="dash",
            line_color=COR_NEUTRO,
            annotation_text=f"Ref: {fmt_v(val_ref)}",
            annotation_font_size=11,
            annotation_position="top left"
        )
        return fig

    g1, g2, g3 = st.columns(3)

    with g1:
        st.plotly_chart(
            mini_grafico("SLA_Dest_pct", sla_origem, "SLA (%)", lambda v: f"{v:.1f}%", True),
            use_container_width=True, key="graf_sla"
        )
    with g2:
        st.plotly_chart(
            mini_grafico("TM_Dest", tm_origem, "Ticket Médio (R$)", lambda v: f"R${v:.2f}", False),
            use_container_width=True, key="graf_tm"
        )
    with g3:
        st.plotly_chart(
            mini_grafico("NFD_Dest_pct", nfd_origem, "NFD (%)", lambda v: f"{v:.1f}%", False),
            use_container_width=True, key="graf_nfd"
        )

    st.caption(
        "🟢 Verde = melhora vs origem  |  "
        "🔴 Vermelho = piora vs origem  |  "
        "Tracejado laranja = referência da transportadora atual"
    )


    # ====================================================
    # COMPARATIVO POR COTAÇÃO DE LEILÃO
    # ====================================================

    st.markdown("### 🏷️ Comparativo por Cotação de Leilão")
    st.caption(
        "Baseado nos pedidos em que a transportadora destino participou do leilão — "
        "comparação direta entre o valor cotado por ela e o valor cotado pela origem, "
        "nos exatos mesmos pedidos."
    )

    total_com_cotacao = int(df_comp["Pedidos_Com_Cotacao_Destino"].fillna(0).sum())
    total_sem_cotacao = int(df_comp["PedidosSimulados"].sum()) - total_com_cotacao
    pct_com_cotacao   = round(total_com_cotacao / df_comp["PedidosSimulados"].sum() * 100, 1) if df_comp["PedidosSimulados"].sum() > 0 else 0

    if total_com_cotacao == 0:
        st.warning(
            f"⚠️ Nenhum pedido de **{codigo_origem}** ({transportadora_origem}) "
            f"teve cotação registrada de **{transportadora_destino}** no período. "
            "Verifique se a transportadora participa do leilão nessa região."
        )
    else:
        mask = df_comp["Pedidos_Com_Cotacao_Destino"] > 0
        df_com_cot = df_comp[mask].copy()

        tm_cot_dest_pond = round(
            (df_com_cot["TM_Cotacao_Destino"] * df_com_cot["Pedidos_Com_Cotacao_Destino"]).sum()
            / df_com_cot["Pedidos_Com_Cotacao_Destino"].sum(), 2
        ) if df_com_cot["Pedidos_Com_Cotacao_Destino"].sum() > 0 else 0

        tm_cot_orig_pond = round(
            (df_com_cot["TM_Cotacao_Origem"] * df_com_cot["Pedidos_Com_Cotacao_Destino"]).sum()
            / df_com_cot["Pedidos_Com_Cotacao_Destino"].sum(), 2
        ) if df_com_cot["Pedidos_Com_Cotacao_Destino"].sum() > 0 else 0

        val_total_cot_dest = round(df_com_cot["ValorTotal_Cotacao_Destino"].sum(), 2)
        val_total_cot_orig = round(df_com_cot["ValorTotal_Cotacao_Origem"].sum(), 2)
        economia_cot = round(val_total_cot_orig - val_total_cot_dest, 2)
        delta_cot    = round(tm_cot_dest_pond - tm_cot_orig_pond, 2)

        # Prazo médio ponderado
        if "Prazo_Cotacao_Destino" in df_com_cot.columns and "Prazo_Cotacao_Origem" in df_com_cot.columns:
            prazo_dest = round(
                (df_com_cot["Prazo_Cotacao_Destino"].fillna(0) * df_com_cot["Pedidos_Com_Cotacao_Destino"]).sum()
                / df_com_cot["Pedidos_Com_Cotacao_Destino"].sum(), 1
            ) if df_com_cot["Pedidos_Com_Cotacao_Destino"].sum() > 0 else 0

            prazo_orig = round(
                (df_com_cot["Prazo_Cotacao_Origem"].fillna(0) * df_com_cot["Pedidos_Com_Cotacao_Destino"]).sum()
                / df_com_cot["Pedidos_Com_Cotacao_Destino"].sum(), 1
            ) if df_com_cot["Pedidos_Com_Cotacao_Destino"].sum() > 0 else 0

            delta_prazo = round(prazo_dest - prazo_orig, 1)
        else:
            prazo_dest = prazo_orig = delta_prazo = None

        # Gasto projetado (histórico destino) nos pedidos com cotação
        gasto_proj_cot = round(
            (df_com_cot["TM_Dest"] * df_com_cot["Pedidos_Com_Cotacao_Destino"]).sum(), 2
        )
        gasto_atual_cot = round(
            (df_com_cot["TM_Cotacao_Origem"] * df_com_cot["Pedidos_Com_Cotacao_Destino"]).sum(), 2
        )
        economia_proj_cot = round(gasto_atual_cot - gasto_proj_cot, 2)

        # ---- 9 KPIs ----
        k1,k2,k3,k4,k5,k6,k7,k8,k9 = st.columns(9)

        with k1:
            st.metric("📋 Pedidos c/ Cotação", f"{total_com_cotacao:,}",
                      delta=f"{pct_com_cotacao:.1f}% dos simulados")
        with k2:
            st.metric("Gasto Atual (c/ cobertura)", fmt_brl(gasto_referencia))
        with k3:
            st.metric(f"Gasto Projetado ({transportadora_destino})",
                      fmt_brl(gasto_projetado))
        with k4:
            st.metric(f"Gasto Cotação ({transportadora_destino})",
                      fmt_brl(val_total_cot_dest))
        with k5:
            st.metric(f"TM Cotação {transportadora_origem}", fmt_brl(tm_cot_orig_pond))
        with k6:
            st.metric(f"TM Cotação {transportadora_destino}",
                      fmt_brl(tm_cot_dest_pond),
                      delta=f"R$ {delta_cot:+,.2f}", delta_color="inverse")
        with k7:
            st.metric(f"TM Projetado ({transportadora_destino})",
                      fmt_brl(tm_destino_pond))
        with k8:
            label_eco = "Economia Projetada" if economia_projetada >= 0 else "Custo Adicional Proj."
            st.metric(label_eco, fmt_brl(abs(economia_projetada)),
                      delta=fmt_brl(economia_projetada),
                      delta_color="normal" if economia_projetada >= 0 else "inverse")
        with k9:
            label_cot = "Economia na Cotação" if economia_cot >= 0 else "Custo Adicional Cot."
            st.metric(label_cot, fmt_brl(abs(economia_cot)),
                      delta=fmt_brl(economia_cot),
                      delta_color="normal" if economia_cot >= 0 else "inverse")

        # ---- Prazo ----
        if prazo_dest is not None:
            st.markdown("#### 📅 Ganho/Perda de Prazo na Cotação")
            p1, p2, p3 = st.columns(3)
            with p1:
                st.metric(f"Prazo Médio Cotação {transportadora_origem}", f"{prazo_orig:.1f} dias úteis")
            with p2:
                st.metric(f"Prazo Médio Cotação {transportadora_destino}", f"{prazo_dest:.1f} dias úteis",
                          delta=f"{delta_prazo:+.1f} dias",
                          delta_color="inverse" if delta_prazo > 0 else "normal")
            with p3:
                if delta_prazo < 0:
                    st.metric("Resultado", f"{abs(delta_prazo):.1f} dias mais rápido 🟢")
                elif delta_prazo > 0:
                    st.metric("Resultado", f"{delta_prazo:.1f} dias mais lento 🔴")
                else:
                    st.metric("Resultado", "Mesmo prazo ➡️")

        st.caption(
            f"📌 Comparação feita nos **{total_com_cotacao:,} pedidos** em que {transportadora_destino} "
            f"participou do leilão. Os outros **{total_sem_cotacao:,} pedidos** não tiveram cotação "
            f"registrada de {transportadora_destino}."
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
    # RESUMO EXECUTIVO
    # ====================================================

    st.markdown("### 📋 Resumo Executivo")

    sla_dir  = "melhora" if delta_sla_pond >= 0 else "piora"
    tm_dir   = "redução" if delta_tm_pond  <= 0 else "aumento"
    nfd_dir  = "redução" if delta_nfd_pond <= 0 else "aumento"
    eco_dir  = "economia" if economia_projetada >= 0 else "custo adicional"

    # Monta texto de cotação para o resumo
    if total_com_cotacao > 0:
        cot_txt = (
            f"\n\n**Comparativo por Cotação de Leilão** ({total_com_cotacao:,} pedidos com cotação registrada):\n"
            f"- 💲 **TM Cotação {transportadora_origem}:** {fmt_brl(tm_cot_orig_pond)} | "
            f"**TM Cotação {transportadora_destino}:** {fmt_brl(tm_cot_dest_pond)} "
            f"({'economia' if delta_cot <= 0 else 'custo adicional'} de {fmt_brl(abs(delta_cot))} por pedido)\n"
            f"- 💰 **Impacto financeiro na cotação:** "
            f"{'economia' if economia_cot >= 0 else 'custo adicional'} de **{fmt_brl(abs(economia_cot))}**"
        )
        if prazo_dest is not None and delta_prazo is not None:
            cot_txt += (
                f"\n- 📅 **Prazo:** {transportadora_origem} cotou {prazo_orig:.1f} dias vs "
                f"{transportadora_destino} cotou {prazo_dest:.1f} dias "
                f"({'mais rápido 🟢' if delta_prazo < 0 else 'mais lento 🔴' if delta_prazo > 0 else 'mesmo prazo ➡️'})"
            )
    else:
        cot_txt = f"\n\n⚠️ **Cotação:** {transportadora_destino} não participou do leilão para esse código no período."

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
- 💲 **Ticket Médio Histórico:** {tm_dir} de {fmt_brl(abs(delta_tm_pond))} por pedido \
({fmt_brl(tm_origem)} → {fmt_brl(tm_destino_pond)})
- 📄 **NFD:** {nfd_dir} de **{abs(delta_nfd_pond):.2f}pp** \
({nfd_origem:.2f}% → {nfd_destino_pond:.2f}%)
- 💰 **Impacto financeiro projetado:** {eco_dir} de **{fmt_brl(abs(economia_projetada))}** \
sobre os pedidos migráveis
{cot_txt}
""")