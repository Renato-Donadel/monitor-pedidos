import os
import streamlit as st
import pandas as pd
import plotly.graph_objects as go

COR_ORIGEM = "#1f4e79"
COR_ALERTA = "#e63946"
COR_OK     = "#2a9d8f"
COR_NEUTRO = "#f4a261"
COR_FUNDO  = "#f4f6f9"

def fmt_pct(v):
    if pd.isna(v): return "—"
    return f"{v:.2f}%"

def fmt_brl(v):
    if pd.isna(v): return "—"
    return f"R$ {v:,.2f}".replace(",","X").replace(".",",").replace("X",".")

def fmt_int(v):
    if pd.isna(v): return "—"
    return f"{int(v):,}"

def kpi_box(label, valor, cor_valor="#0f2a44", fundo="#ffffff", delta=None, delta_cor=None):
    delta_html = ""
    if delta is not None:
        dcor = delta_cor or ("#27ae60" if "+" in str(delta) else "#e74c3c")
        delta_html = f'<div style="font-size:11px;color:{dcor};margin-top:2px">{delta}</div>'
    return f"""<div style="background:{fundo};border-radius:10px;padding:12px 14px;
                border:1px solid #e0e4ea;text-align:center">
        <div style="font-size:11px;color:#888;margin-bottom:4px">{label}</div>
        <div style="font-size:18px;font-weight:700;color:{cor_valor}">{valor}</div>
        {delta_html}</div>"""

@st.cache_data(ttl=120)
def carregar_bases():
    arq_pedidos = "data/Base_Pedidos_Codigo.xlsx"
    arq_sim     = "data/Base_Similaridade_Tarifarios.xlsx"
    arq_nfd     = "data/Base_NFD_Real.xlsx"
    arq_intel   = "data/Base_Intelipost_Resumo.xlsx"

    df_trade = pd.read_excel(arq_pedidos) if os.path.exists(arq_pedidos) else pd.DataFrame()
    df_sim   = pd.read_excel(arq_sim)     if os.path.exists(arq_sim)     else pd.DataFrame()
    df_nfd   = pd.read_excel(arq_nfd)     if os.path.exists(arq_nfd)     else pd.DataFrame()
    df_intel = pd.read_excel(arq_intel)   if os.path.exists(arq_intel)   else pd.DataFrame()

    if not df_trade.empty:
        df_trade.columns = df_trade.columns.str.strip()
        if "DataFinal"      in df_trade.columns: df_trade["DataFinal"]   = pd.to_datetime(df_trade["DataFinal"],   errors="coerce")
        if "TemNFD"         in df_trade.columns: df_trade["TemNFD"]      = df_trade["TemNFD"].fillna(False).astype(bool)
        if "DentroPrazo"    in df_trade.columns: df_trade["DentroPrazo"] = df_trade["DentroPrazo"].fillna(False).astype(bool)
        if "ValorFrete"     in df_trade.columns: df_trade["ValorFrete"]  = pd.to_numeric(df_trade["ValorFrete"],  errors="coerce")
        if "ValorNota"      in df_trade.columns: df_trade["ValorNota"]   = pd.to_numeric(df_trade["ValorNota"],   errors="coerce")

    if not df_nfd.empty:
        df_nfd.columns = df_nfd.columns.str.strip()
        if "DataExpedicao"    in df_nfd.columns: df_nfd["DataExpedicao"]    = pd.to_datetime(df_nfd["DataExpedicao"], errors="coerce")
        if "DataDespacho"     in df_nfd.columns: df_nfd["DataDespacho"]     = pd.to_datetime(df_nfd["DataDespacho"], errors="coerce")
        if "ValorNota"        in df_nfd.columns: df_nfd["ValorNota"]        = pd.to_numeric(df_nfd["ValorNota"],      errors="coerce")
        if "TemNFD"           in df_nfd.columns: df_nfd["TemNFD"]           = df_nfd["TemNFD"].fillna(False).astype(bool)
        if "PassouRetirada"   in df_nfd.columns: df_nfd["PassouRetirada"]   = df_nfd["PassouRetirada"].fillna(False).astype(bool)
        if "TspMasNaoTsp"     in df_nfd.columns: df_nfd["TspMasNaoTsp"]     = df_nfd["TspMasNaoTsp"].fillna(False).astype(bool)
        if "CruzouIntelipost" in df_nfd.columns: df_nfd["CruzouIntelipost"] = df_nfd["CruzouIntelipost"].fillna(False).astype(bool)
        if "StatusNaoTsp"     in df_nfd.columns: df_nfd["StatusNaoTsp"]     = df_nfd["StatusNaoTsp"].fillna("").astype(str)

    return df_trade, df_sim, df_nfd, df_intel

def render_tradeoff():

    df_trade, df_sim, df_nfd_real, df_intel = carregar_bases()
    if df_trade.empty or df_sim.empty:
        st.error("Bases nao encontradas. Execute o ETL primeiro.")
        return

    # ====================================================
    # TOTAIS GERAIS (sem filtro — dados históricos completos)
    # ====================================================

    st.markdown("### Visao Geral da Empresa 2026 — Sem Filtro")

    # Total de vendas BRUTO (soma das planilhas do Joao, sem filtro nenhum)
    total_vendas_bruto = None
    if not df_intel.empty and "TotalVendasBruto" in df_intel.columns:
        total_vendas_bruto = df_intel["TotalVendasBruto"].iloc[0]
    # Fallback: se nao tiver TotalVendasBruto, usa TotalVendasR$ (filtrado)
    if not total_vendas_bruto and not df_intel.empty and "TotalVendasR$" in df_intel.columns:
        total_vendas_bruto = df_intel["TotalVendasR$"].iloc[0]

    # Total NFD bruto (soma da Rentabilidade 2026, sem filtro de motivo)
    total_nfd_bruto = None
    if not df_intel.empty and "TotalNFDBruto" in df_intel.columns:
        total_nfd_bruto = df_intel["TotalNFDBruto"].iloc[0]

    # Guardados para a secao Debug
    nfd_tsp_mas_nao = pd.DataFrame()
    pedidos_sem_valor = None
    if "PedidosSemValorNota" in df_intel.columns and not df_intel.empty:
        pedidos_sem_valor = df_intel["PedidosSemValorNota"].iloc[0]

    # Categorias de NFD
    if not df_nfd_real.empty:
        if "TspMasNaoTsp" in df_nfd_real.columns:
            mask_tsp_mas_nao = df_nfd_real["TspMasNaoTsp"] == True
        else:
            mask_tsp_mas_nao = pd.Series(False, index=df_nfd_real.index)

        nfd_tsp_mas_nao = df_nfd_real[mask_tsp_mas_nao].copy()

        # TSP "de verdade" = motivo TSP E NAO passou por status que descaracteriza
        df_nfd_tsp   = df_nfd_real[(df_nfd_real["TemNFD"] == True) & (~mask_tsp_mas_nao)].copy()

        if "CruzouIntelipost" in df_nfd_tsp.columns:
            nfd_cruzou     = df_nfd_tsp[df_nfd_tsp["CruzouIntelipost"] == True]
            nfd_nao_cruzou = df_nfd_tsp[df_nfd_tsp["CruzouIntelipost"] == False]
        else:
            nfd_cruzou     = df_nfd_tsp
            nfd_nao_cruzou = pd.DataFrame()

        total_nfd_tsp      = df_nfd_tsp["ValorNota"].sum()
        total_nfd_cruzou   = nfd_cruzou["ValorNota"].sum()
        total_nfd_nao_cruz = nfd_nao_cruzou["ValorNota"].sum()
        pct_nfd_tsp        = (total_nfd_tsp / total_vendas_bruto * 100) if total_vendas_bruto else None

        # NFD outros setores = total bruto - TSP - "tsp mas nao e tsp"
        total_nfd_outros = None
        if total_nfd_bruto is not None:
            total_tsp_mas_nao_val = nfd_tsp_mas_nao["ValorNota"].sum() if not nfd_tsp_mas_nao.empty else 0
            total_nfd_outros = total_nfd_bruto - total_nfd_tsp - total_tsp_mas_nao_val
    else:
        total_nfd_tsp = total_nfd_cruzou = total_nfd_nao_cruz = None
        nfd_nao_cruzou = pd.DataFrame()
        pct_nfd_tsp = total_nfd_outros = None

    # ── LINHA 1: Cards principais ────────────────────────────
    g1,g2,g3,g4,g5 = st.columns(5)
    g1.markdown(kpi_box("Total de Vendas 2026 (R$)",
        fmt_brl(total_vendas_bruto) if total_vendas_bruto else "sem dado"), unsafe_allow_html=True)
    g2.markdown(kpi_box("Total NFD 2026 (R$)",
        fmt_brl(total_nfd_bruto) if total_nfd_bruto else "sem dado",
        cor_valor=COR_ALERTA), unsafe_allow_html=True)
    g3.markdown(kpi_box("NFD TSP (R$)",
        fmt_brl(total_nfd_tsp) if total_nfd_tsp else "sem dado",
        cor_valor=COR_ALERTA), unsafe_allow_html=True)
    g4.markdown(kpi_box("NFD Outros Setores (R$)",
        fmt_brl(total_nfd_outros) if total_nfd_outros is not None else "sem dado",
        cor_valor=COR_NEUTRO), unsafe_allow_html=True)
    g5.markdown(kpi_box("% NFD TSP / Vendas",
        fmt_pct(pct_nfd_tsp) if pct_nfd_tsp else "sem dado",
        cor_valor=COR_ALERTA if (pct_nfd_tsp or 0)>=5 else COR_NEUTRO if (pct_nfd_tsp or 0)>=2 else COR_OK),
        unsafe_allow_html=True)

    # ── LINHA 2: Detalhe TSP (Intelipost) ────────────────────
    h1,h2 = st.columns(2)
    h1.markdown(kpi_box("NFD TSP no Intelipost (R$)",
        fmt_brl(total_nfd_cruzou) if total_nfd_cruzou else "sem dado"), unsafe_allow_html=True)
    h2.markdown(kpi_box("NFD TSP fora Intelipost (R$)",
        fmt_brl(total_nfd_nao_cruz) if total_nfd_nao_cruz else "sem dado",
        cor_valor=COR_ALERTA), unsafe_allow_html=True)

    st.caption("Notas com NFD TSP que passaram por status de retirada/endereco/ausente foram retiradas dos totais TSP (ver Debug no fim da pagina).")

    # Botao exportar NFD fora Intelipost
    if not nfd_nao_cruzou.empty:
        import io
        buf = io.BytesIO()
        nfd_nao_cruzou.to_excel(buf, index=False)
        buf.seek(0)
        st.caption("NFD TSP sem match no Intelipost:")
        st.download_button(
            "⬇️ Exportar lista",
            data=buf,
            file_name="nfd_fora_intelipost.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.divider()

    # ── PERÍODO (sem transportadora ainda) ──────────────
    col_per, _ = st.columns([1, 1])
    with col_per:
        periodo = st.selectbox("Periodo", ["30 dias", "60 dias", "90 dias"])
        dias = {"30 dias": 30, "60 dias": 60, "90 dias": 90}[periodo]

    data_corte = pd.Timestamp.today() - pd.Timedelta(days=dias)

    df_todas = df_trade[df_trade["DataFinal"] >= data_corte].copy()

    st.divider()

    # ── VISÃO POR TRANSPORTADORA ──────────────────────────
    st.markdown(f"### NFD por Transportadora | Ultimos {dias} dias")

    # VENDAS por transportadora: df_trade (Base_Pedidos_Codigo) filtrado por periodo
    vendas_tsp = (
        df_todas.groupby("Transportadora")
        .agg(
            Pedidos    = ("ValorNota", "count"),
            ValorVendas = ("ValorNota", "sum"),
        ).reset_index()
    )

    # NFD por transportadora: df_nfd_real (Base_NFD_Real) com Transportadora do Intelipost
    # Filtro: DataDespacho no periodo + TemNFD (motivo TSP) + exclui TspMasNaoTsp
    nfd_tsp_tabela = pd.DataFrame()
    nfd_sem_tsp_periodo = 0
    if not df_nfd_real.empty and "DataDespacho" in df_nfd_real.columns and "Transportadora" in df_nfd_real.columns:
        mask_periodo = df_nfd_real["DataDespacho"] >= data_corte
        mask_tsp     = df_nfd_real["TemNFD"] == True
        mask_real    = df_nfd_real.get("TspMasNaoTsp", pd.Series(False, index=df_nfd_real.index)) != True
        mask_intel   = df_nfd_real["Transportadora"].fillna("").str.strip() != "" if "Transportadora" in df_nfd_real.columns else df_nfd_real.get("CruzouIntelipost", pd.Series(True, index=df_nfd_real.index)) == True

        # NFD TSP real COM transportadora no periodo
        nfd_com_tsp = df_nfd_real[mask_periodo & mask_tsp & mask_real & mask_intel]
        # NFD TSP real SEM transportadora no periodo (para mostrar a contagem)
        nfd_sem_tsp_periodo = int((mask_periodo & mask_tsp & mask_real & ~mask_intel).sum())

        if not nfd_com_tsp.empty:
            nfd_tsp_tabela = (
                nfd_com_tsp.groupby("Transportadora")
                .agg(
                    PedidosNFD = ("ValorNota", "count"),
                    ValorNFD   = ("ValorNota", "sum"),
                ).reset_index()
            )

    # Merge vendas + NFD por transportadora
    resumo_tsp = vendas_tsp.merge(nfd_tsp_tabela, on="Transportadora", how="left").fillna(0)
    resumo_tsp["NFD_pct"] = (resumo_tsp["ValorNFD"] / resumo_tsp["ValorVendas"] * 100).round(2)
    resumo_tsp.loc[resumo_tsp["ValorVendas"] == 0, "NFD_pct"] = 0
    resumo_tsp = resumo_tsp.sort_values("NFD_pct", ascending=False)

    df_exib = resumo_tsp.rename(columns={
        "ValorVendas": "Valor Vendas (R$)",
        "PedidosNFD":  "Pedidos NFD",
        "ValorNFD":    "Valor NFD (R$)",
        "NFD_pct":     "% NFD",
    })[["Transportadora","Pedidos","Valor Vendas (R$)","Pedidos NFD","Valor NFD (R$)","% NFD"]]

    st.dataframe(
        df_exib.style
        .format({
            "Pedidos":           "{:,.0f}",
            "Valor Vendas (R$)": "R$ {:,.2f}",
            "Pedidos NFD":       "{:,.0f}",
            "Valor NFD (R$)":    "R$ {:,.2f}",
            "% NFD":             "{:.2f}%",
        })
        .background_gradient(subset=["% NFD"], cmap="RdYlGn_r"),
        use_container_width=True,
        hide_index=True
    )

    if nfd_sem_tsp_periodo > 0:
        st.caption(f"⚠️ {nfd_sem_tsp_periodo:,} notas NFD TSP no periodo sem transportadora identificada (nem Intelipost nem PRW).")

    st.divider()

    # ── FILTRO DE TRANSPORTADORA ─────────────────────────
    st.markdown("---")
    transportadoras = sorted(df_trade["Transportadora"].dropna().unique())
    transp_sel = st.selectbox("Selecione a Transportadora para analise detalhada", transportadoras)

    df_periodo = df_trade[
        (df_trade["Transportadora"] == transp_sel) &
        (df_trade["DataFinal"] >= data_corte)
    ].copy()

    # ── VISÃO GERAL DA TRANSPORTADORA SELECIONADA ─────────
    if df_periodo.empty:
        st.warning("Nenhum pedido encontrado para esse filtro.")
        return

    st.markdown(f"### Visao Geral — {transp_sel} (filtrado) | Ultimos {dias} dias")

    total_ped     = len(df_periodo)
    total_notas   = df_periodo["ValorNota"].sum()   if "ValorNota" in df_periodo.columns else None
    total_nfd_n   = int(df_periodo["TemNFD"].sum()) if "TemNFD"    in df_periodo.columns else 0
    nfd_pct_ger   = total_nfd_n / total_ped * 100   if total_ped > 0 else 0
    nfd_valor_ger = (total_notas * nfd_pct_ger / 100) if total_notas else None

    g1,g2,g3,g4,g5 = st.columns(5)
    g1.markdown(kpi_box("Total de Pedidos",    fmt_int(total_ped)), unsafe_allow_html=True)
    g2.markdown(kpi_box("Total de Vendas (R$)",fmt_brl(total_notas) if total_notas else "sem dado"), unsafe_allow_html=True)
    g3.markdown(kpi_box("Pedidos com NFD",     fmt_int(total_nfd_n)), unsafe_allow_html=True)
    g4.markdown(kpi_box("Valor NFD (R$)",      fmt_brl(nfd_valor_ger) if nfd_valor_ger else "sem dado", cor_valor=COR_ALERTA), unsafe_allow_html=True)
    g5.markdown(kpi_box("% NFD",               fmt_pct(nfd_pct_ger),
        cor_valor=COR_ALERTA if nfd_pct_ger>=5 else COR_NEUTRO if nfd_pct_ger>=2 else COR_OK),
        unsafe_allow_html=True)

    st.divider()

    # ── RANKING ──────────────────────────────────────────
    st.caption(f"Baseado nos ultimos {dias} dias | Minimo 30 pedidos")

    nfd_rank = (
        df_periodo.groupby("CodigoTarifario")
        .agg(
            Pedidos    = ("TemNFD",      "count"),
            NFD_n      = ("TemNFD",      "sum"),
            SLA_pct    = ("DentroPrazo", "mean"),
            TM         = ("ValorFrete",  "mean"),
            ValorNotas = ("ValorNota",   "sum"),
        ).reset_index()
    )
    nfd_rank = nfd_rank[nfd_rank["Pedidos"] >= 30].copy()
    nfd_rank["NFD_pct"] = (nfd_rank["NFD_n"] / nfd_rank["Pedidos"] * 100).round(2)
    nfd_rank["SLA_pct"] = (nfd_rank["SLA_pct"] * 100).round(2)
    nfd_rank = nfd_rank.sort_values("NFD_pct", ascending=False).reset_index(drop=True)

    if nfd_rank.empty:
        st.warning("Nenhum codigo com pedidos suficientes no periodo.")
        return

    chave_pag = f"pagina_nfd_{transp_sel}_{dias}"
    if chave_pag not in st.session_state:
        st.session_state[chave_pag] = 0

    por_pag    = 10
    total_pags = max(0, (len(nfd_rank) - 1) // por_pag)
    ini        = st.session_state[chave_pag] * por_pag
    fim        = ini + por_pag
    pagina_df  = nfd_rank.iloc[ini:fim].copy()

    cores = [COR_ALERTA if v >= 5 else COR_NEUTRO if v >= 2 else COR_OK for v in pagina_df["NFD_pct"]]
    fig = go.Figure(go.Bar(
        x=pagina_df["NFD_pct"], y=pagina_df["CodigoTarifario"],
        orientation="h", marker_color=cores,
        text=[f"{v:.2f}% ({p:,} ped.)" for v, p in zip(pagina_df["NFD_pct"], pagina_df["Pedidos"])],
        textposition="outside",
    ))
    fig.update_layout(
        height=max(300, len(pagina_df)*42), xaxis_title="NFD (%)",
        yaxis=dict(autorange="reversed", tickfont=dict(size=12)),
        plot_bgcolor=COR_FUNDO, paper_bgcolor=COR_FUNDO,
        margin=dict(l=10, r=100, t=20, b=30), font=dict(family="Arial", size=12),
    )
    st.plotly_chart(fig, use_container_width=True)
    st.caption("Vermelho: NFD >= 5%  |  Laranja: 2-5%  |  Verde: < 2%")

    nav1, nav2, nav3 = st.columns([1, 5, 1])
    with nav1:
        if st.session_state[chave_pag] > 0:
            if st.button("Anteriores"):
                st.session_state[chave_pag] -= 1
                st.rerun()
    with nav2:
        st.caption(f"Exibindo {ini+1}-{min(fim,len(nfd_rank))} de {len(nfd_rank)} | Pagina {st.session_state[chave_pag]+1} de {total_pags+1}")
    with nav3:
        if st.session_state[chave_pag] < total_pags:
            if st.button("Proximos"):
                st.session_state[chave_pag] += 1
                st.rerun()

    st.divider()

    # ── SELETOR ──────────────────────────────────────────
    st.markdown("### Selecione os codigos para analise detalhada")
    codigos_pag = pagina_df["CodigoTarifario"].tolist()
    codigos_sel = st.multiselect(
        "Codigos tarifarios (pagina atual)",
        options=codigos_pag,
        default=codigos_pag,
        help="Remova os que nao quer analisar. Ao virar pagina os novos 10 entram automaticamente."
    )

    if not codigos_sel:
        st.info("Selecione pelo menos um codigo acima.")
        return

    st.divider()

    # ── CARDS POR CODIGO ─────────────────────────────────
    st.markdown("### Detalhe por Codigo Tarifario")

    resumo_codigos  = []
    ganhos_positivos = []

    for codigo in codigos_sel:
        with st.expander(f"{transp_sel} / {codigo}", expanded=True):

            df_cod = df_periodo[df_periodo["CodigoTarifario"] == codigo].copy()

            pedidos_orig  = len(df_cod)
            sla_orig      = df_cod["DentroPrazo"].mean() * 100 if pedidos_orig > 0 else 0
            tm_orig       = df_cod["ValorFrete"].mean()         if pedidos_orig > 0 else 0
            nfd_orig_pct  = df_cod["TemNFD"].mean() * 100       if pedidos_orig > 0 else 0
            frete_orig    = df_cod["ValorFrete"].sum()
            valor_notas   = df_cod["ValorNota"].sum() if "ValorNota" in df_cod.columns and df_cod["ValorNota"].notna().any() else None
            nfd_orig_valor= (valor_notas * nfd_orig_pct / 100) if valor_notas else None

            st.markdown("**Situacao atual**")
            c1,c2,c3,c4,c5,c6 = st.columns(6)
            c1.markdown(kpi_box("Pedidos",    fmt_int(pedidos_orig)), unsafe_allow_html=True)
            c2.markdown(kpi_box("SLA",        fmt_pct(sla_orig),    cor_valor=COR_OK if sla_orig>=95 else COR_ALERTA), unsafe_allow_html=True)
            c3.markdown(kpi_box("TM Frete",   fmt_brl(tm_orig)),    unsafe_allow_html=True)
            c4.markdown(kpi_box("NFD %",      fmt_pct(nfd_orig_pct),cor_valor=COR_ALERTA if nfd_orig_pct>=5 else COR_NEUTRO if nfd_orig_pct>=2 else COR_OK), unsafe_allow_html=True)
            c5.markdown(kpi_box("NFD R$",     fmt_brl(nfd_orig_valor) if nfd_orig_valor else "sem dado", cor_valor=COR_ALERTA), unsafe_allow_html=True)
            c6.markdown(kpi_box("Valor Notas",fmt_brl(valor_notas) if valor_notas else "sem dado"), unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # Redistribuição — vem da cotação (novo Bloco 10)
            df_sim_cod = df_sim[
                (df_sim["TransportadoraOrigem"] == transp_sel) &
                (df_sim["CodigoOrigem"] == codigo)
            ].copy()

            if df_sim_cod.empty:
                st.warning("Sem dados de redistribuicao para esse codigo.")
                resumo_codigos.append({
                    "codigo":codigo, "pedidos_orig":pedidos_orig,
                    "sla_orig":sla_orig, "tm_orig":tm_orig,
                    "nfd_orig_pct":nfd_orig_pct, "frete_orig":frete_orig,
                    "valor_notas":valor_notas, "nfd_orig_valor":nfd_orig_valor,
                    "destinos":[]
                })
                continue

            # Total de pedidos redistribuídos por cotação
            total_redistribuido = df_sim_cod["Pedidos"].sum()

            rows_dest = []
            for tsp_dest, grp in df_sim_cod.groupby("TransportadoraDestino"):
                ped_sim  = int(grp["Pedidos"].sum())
                if ped_sim == 0: continue

                # Métricas históricas (ponderadas por pedidos)
                w        = grp["Pedidos"]
                sla_d    = (grp["SLA_Hist"].fillna(0)  * w).sum() / w.sum() * 100 if "SLA_Hist" in grp.columns else None
                nfd_d    = (grp["NFD_Hist"].fillna(0)  * w).sum() / w.sum() * 100 if "NFD_Hist" in grp.columns else None
                tm_hist  = (grp["TM_Hist"].fillna(0)   * w).sum() / w.sum()        if "TM_Hist"  in grp.columns else None

                # Frete histórico projetado
                frete_hist_proj = grp["ProjecaoFreteHist"].sum() if "ProjecaoFreteHist" in grp.columns else None

                # Frete por cotação
                cot_dest_total  = grp["ValorCotacaoDestTotal"].sum() if "ValorCotacaoDestTotal" in grp.columns else None
                cot_orig_total  = grp["ValorCotacaoOrigTotal"].sum()  if "ValorCotacaoOrigTotal" in grp.columns else None

                # Frete atual proporcional (pelos pedidos redistribuídos, não pelo período filtrado)
                frete_orig_prop = grp["ValorFreteOrigem"].sum() if "ValorFreteOrigem" in grp.columns else None

                # Deltas
                delta_frete_hist = (frete_hist_proj - frete_orig_prop) if (frete_hist_proj is not None and frete_orig_prop) else None
                delta_frete_cot  = (cot_dest_total  - cot_orig_total)  if (cot_dest_total  is not None and cot_orig_total)  else None

                pct_dist     = ped_sim / total_redistribuido * 100 if total_redistribuido > 0 else 0
                notas_dest   = (valor_notas * pct_dist / 100) if valor_notas else None
                nfd_val_dest = (notas_dest  * nfd_d   / 100)  if (notas_dest and nfd_d is not None) else None

                # TM cotação destino ponderado
                ped_cot  = grp["Pedidos"].sum()
                tm_cot   = cot_dest_total / ped_cot if (cot_dest_total and ped_cot > 0) else None

                rows_dest.append({
                    "TransportadoraDestino": tsp_dest,
                    "PedidosSimulados":  ped_sim,
                    "Pct_dist":          pct_dist,
                    "SLADestino":        sla_d,
                    "NFDDestino":        nfd_d,
                    "TM_Hist":           tm_hist,
                    "TM_Cotacao":        tm_cot,
                    "FreteHistProj":     frete_hist_proj,
                    "FreteCotDest":      cot_dest_total,
                    "FreteOrigProp":     frete_orig_prop,
                    "DeltaFreteHist":    delta_frete_hist,
                    "DeltaFreteCot":     delta_frete_cot,
                    "ValorNotasDest":    notas_dest,
                    "NFDValorDest":      nfd_val_dest,
                    "Pedidos_Com_Cotacao": ped_cot,
                })

            grp_dest = pd.DataFrame(rows_dest).sort_values("PedidosSimulados", ascending=False)

            # ── Ganho líquido por código tarifário ───────────────
            # GanhoLiquido = SUM(DeltaFreteCot) - SUM(NFDValorDest) + nfd_orig_valor
            delta_frete_total = grp_dest["DeltaFreteCot"].sum()   if "DeltaFreteCot"  in grp_dest.columns else None
            nfd_dest_total    = grp_dest["NFDValorDest"].sum()     if "NFDValorDest"   in grp_dest.columns else None
            if delta_frete_total is not None and nfd_dest_total is not None and nfd_orig_valor is not None:
                ganho_liquido = delta_frete_total - nfd_dest_total + nfd_orig_valor
                if ganho_liquido > 0:
                    ganhos_positivos.append({
                        "Transportadora":     transp_sel,
                        "CodigoTarifario":    codigo,
                        "Pedidos":            pedidos_orig,
                        "NFD_Atual_R$":       nfd_orig_valor,
                        "DeltaFrete_Cot_R$":  delta_frete_total,
                        "NFD_Destinos_R$":    nfd_dest_total,
                        "GanhoLiquido_R$":    ganho_liquido,
                    })

            st.markdown(f"**Para onde iriam os {total_redistribuido:,} pedidos (base cotacao)**")
            st.caption("Destino = mais barato no leilao excluindo a transportadora atual | Historico = dados reais de entrega")

            for _, row in grp_dest.iterrows():
                st.markdown(
                    f"<div style='background:#e8f0fe;border-radius:8px;padding:6px 12px;"
                    f"margin-bottom:6px;font-size:12px;font-weight:600;color:#1f4e79'>"
                    f"{row['TransportadoraDestino']} &nbsp; {int(row['PedidosSimulados']):,} pedidos ({row['Pct_dist']:.1f}%)</div>",
                    unsafe_allow_html=True
                )

                d1,d2,d3,d4,d5,d6,d7,d8 = st.columns(8)

                delta_sla_v = (row["SLADestino"] - sla_orig) if row["SLADestino"] is not None else None
                delta_nfd_v = (row["NFDDestino"] - nfd_orig_pct) if row["NFDDestino"] is not None else None
                dh = row["DeltaFreteHist"]
                dc = row["DeltaFreteCot"]
                notas_dest_val = row.get("ValorNotasDest")

                d1.markdown(kpi_box("Valor Dest.",
                    fmt_brl(notas_dest_val) if pd.notna(notas_dest_val) else "sem dado"), unsafe_allow_html=True)

                d2.markdown(kpi_box("SLA Hist.",
                    fmt_pct(row["SLADestino"]),
                    cor_valor=COR_OK if (row["SLADestino"] or 0)>=95 else COR_ALERTA,
                    delta=f"{delta_sla_v:+.1f}pp" if delta_sla_v is not None else None,
                    delta_cor=COR_OK if (delta_sla_v or 0)>=0 else COR_ALERTA), unsafe_allow_html=True)

                d3.markdown(kpi_box("NFD % Hist.",
                    fmt_pct(row["NFDDestino"]),
                    cor_valor=COR_ALERTA if (row["NFDDestino"] or 0)>=5 else COR_NEUTRO if (row["NFDDestino"] or 0)>=2 else COR_OK,
                    delta=f"{delta_nfd_v:+.1f}pp" if delta_nfd_v is not None else None,
                    delta_cor=COR_OK if (delta_nfd_v or 0)<=0 else COR_ALERTA), unsafe_allow_html=True)

                d4.markdown(kpi_box("TM Hist.", fmt_brl(row["TM_Hist"])), unsafe_allow_html=True)

                d5.markdown(kpi_box("TM Cotacao", fmt_brl(row["TM_Cotacao"]) if row["TM_Cotacao"] else "sem cot."), unsafe_allow_html=True)

                d6.markdown(kpi_box("Delta Frete Hist.",
                    fmt_brl(abs(dh)) if dh is not None else "—",
                    cor_valor=COR_OK if (dh or 0)<=0 else COR_ALERTA,
                    delta="ganho" if (dh or 0)<=0 else "perda",
                    delta_cor=COR_OK if (dh or 0)<=0 else COR_ALERTA), unsafe_allow_html=True)

                d7.markdown(kpi_box("Delta Frete Cot.",
                    fmt_brl(abs(dc)) if dc is not None else "—",
                    cor_valor=COR_OK if (dc or 0)<=0 else COR_ALERTA,
                    delta="ganho" if (dc or 0)<=0 else "perda",
                    delta_cor=COR_OK if (dc or 0)<=0 else COR_ALERTA), unsafe_allow_html=True)

                d8.markdown(kpi_box("NFD R$ Est.",
                    fmt_brl(row["NFDValorDest"]) if pd.notna(row.get("NFDValorDest")) else "sem dado"), unsafe_allow_html=True)

                st.markdown("<div style='margin-bottom:8px'></div>", unsafe_allow_html=True)

            resumo_codigos.append({
                "codigo":codigo, "pedidos_orig":pedidos_orig,
                "sla_orig":sla_orig, "tm_orig":tm_orig,
                "nfd_orig_pct":nfd_orig_pct, "frete_orig":frete_orig,
                "valor_notas":valor_notas, "nfd_orig_valor":nfd_orig_valor,
                "destinos":grp_dest.to_dict("records")
            })

    # ── EXPORTAR CÓDIGOS COM GANHO LÍQUIDO POSITIVO ──────
    st.divider()
    if ganhos_positivos:
        df_ganhos = pd.DataFrame(ganhos_positivos).sort_values("GanhoLiquido_R$", ascending=False)
        import io
        buf = io.BytesIO()
        df_ganhos.to_excel(buf, index=False)
        buf.seek(0)
        st.markdown(f"**✅ {len(df_ganhos)} código(s) com ganho líquido positivo no período ({periodo})**")
        st.download_button(
            "⬇️ Exportar códigos com ganho líquido",
            data=buf,
            file_name=f"ganho_liquido_{transp_sel}_{periodo.replace(' ','')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Nenhum código tarifário com ganho líquido positivo no período selecionado.")

    # ── RESUMO CONSOLIDADO ───────────────────────────────
    if not resumo_codigos:
        return

    st.divider()
    st.markdown("### Resumo Consolidado")

    total_pedidos_geral = sum(r["pedidos_orig"]  for r in resumo_codigos)
    total_frete_orig    = sum(r["frete_orig"]     for r in resumo_codigos if r["frete_orig"])
    total_notas         = sum(r["valor_notas"]    for r in resumo_codigos if r["valor_notas"])
    total_nfd_orig_val  = sum(r["nfd_orig_valor"] for r in resumo_codigos if r["nfd_orig_valor"])
    sla_orig_pond       = sum(r["sla_orig"] * r["pedidos_orig"] for r in resumo_codigos) / total_pedidos_geral if total_pedidos_geral else 0

    from collections import defaultdict
    dest_agg = defaultdict(lambda: {
        "pedidos":0, "frete_hist_proj":0, "frete_cot_dest":0,
        "frete_orig_prop":0, "sla_pond":0, "nfd_pond":0,
        "notas":0, "nfd_valor":0, "delta_hist":0, "delta_cot":0,
        "tem_cot": False
    })

    for r in resumo_codigos:
        for d in r["destinos"]:
            tsp = d["TransportadoraDestino"]
            dest_agg[tsp]["pedidos"]         += d["PedidosSimulados"]
            dest_agg[tsp]["frete_hist_proj"] += d.get("FreteHistProj")  or 0
            dest_agg[tsp]["frete_cot_dest"]  += d.get("FreteCotDest")   or 0
            dest_agg[tsp]["frete_orig_prop"] += d.get("FreteOrigProp")  or 0
            dest_agg[tsp]["sla_pond"]        += (d.get("SLADestino") or 0) * d["PedidosSimulados"]
            dest_agg[tsp]["nfd_pond"]        += (d.get("NFDDestino") or 0) * d["PedidosSimulados"]
            dest_agg[tsp]["notas"]           += d.get("ValorNotasDest") or 0
            dest_agg[tsp]["nfd_valor"]       += d.get("NFDValorDest")   or 0
            dest_agg[tsp]["delta_hist"]      += d.get("DeltaFreteHist") or 0
            dest_agg[tsp]["delta_cot"]       += d.get("DeltaFreteCot")  or 0
            if d.get("FreteCotDest"):
                dest_agg[tsp]["tem_cot"] = True

    rows_resumo = []
    for tsp, v in dest_agg.items():
        ped = v["pedidos"]
        if ped == 0: continue
        rows_resumo.append({
            "Transportadora":  tsp,
            "Pedidos":         ped,
            "Pct":             ped / total_pedidos_geral * 100,
            "SLA":             v["sla_pond"] / ped,
            "NFD_pct":         v["nfd_pond"] / ped,
            "FreteHistProj":   v["frete_hist_proj"],
            "FreteCotDest":    v["frete_cot_dest"] if v["tem_cot"] else None,
            "FreteOrigProp":   v["frete_orig_prop"],
            "NFDValor":        v["nfd_valor"],
            "DeltaFreteHist":  v["delta_hist"],
            "DeltaFreteCot":   v["delta_cot"] if v["tem_cot"] else None,
        })

    df_resumo = pd.DataFrame(rows_resumo).sort_values("Pedidos", ascending=False)

    total_frete_hist_proj = df_resumo["FreteHistProj"].sum()
    total_frete_cot_dest  = df_resumo["FreteCotDest"].dropna().sum()
    total_frete_orig_prop = df_resumo["FreteOrigProp"].sum()
    total_nfd_dest        = df_resumo["NFDValor"].sum()
    total_delta_hist      = df_resumo["DeltaFreteHist"].sum()
    total_delta_cot       = df_resumo["DeltaFreteCot"].dropna().sum()
    sla_dest_pond         = (df_resumo["SLA"] * df_resumo["Pedidos"]).sum() / df_resumo["Pedidos"].sum() if len(df_resumo) > 0 else 0

    # ganho positivo, perda negativo
    ganho_frete_hist = -total_delta_hist   # delta negativo = destino mais barato = ganho
    ganho_frete_cot  = -total_delta_cot    if total_delta_cot else None
    ganho_nfd        = total_nfd_orig_val - total_nfd_dest
    delta_sla        = sla_dest_pond - sla_orig_pond
    saldo            = ganho_nfd + ganho_frete_hist

    # KPIs gerais
    k1,k2,k3,k4,k5 = st.columns(5)
    k1.markdown(kpi_box("Total de Pedidos",     fmt_int(total_pedidos_geral)), unsafe_allow_html=True)
    k2.markdown(kpi_box("Valor Total das Notas",fmt_brl(total_notas) if total_notas else "sem dado"), unsafe_allow_html=True)
    k3.markdown(kpi_box("NFD Atual (R$)",       fmt_brl(total_nfd_orig_val) if total_nfd_orig_val else "sem dado", cor_valor=COR_ALERTA), unsafe_allow_html=True)
    k4.markdown(kpi_box("Frete Atual",          fmt_brl(total_frete_orig)), unsafe_allow_html=True)
    k5.markdown(kpi_box("SLA Atual",            fmt_pct(sla_orig_pond), cor_valor=COR_OK if sla_orig_pond>=95 else COR_ALERTA), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**Distribuicao por transportadora destino**")

    for _, row in df_resumo.iterrows():
        tsp = row["Transportadora"]
        ped = int(row["Pedidos"])
        pct = row["Pct"]
        dh  = row["DeltaFreteHist"]
        dc  = row["DeltaFreteCot"]

        st.markdown(
            f"<div style='font-size:13px;font-weight:600;color:#1f4e79;margin:8px 0 4px'>"
            f"{tsp} - {ped:,} pedidos ({pct:.1f}%)</div>", unsafe_allow_html=True)

        c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
        c1.markdown(kpi_box("SLA",         fmt_pct(row["SLA"])),     unsafe_allow_html=True)
        c2.markdown(kpi_box("NFD %",       fmt_pct(row["NFD_pct"])), unsafe_allow_html=True)
        c3.markdown(kpi_box("Frete Orig.", fmt_brl(row["FreteOrigProp"])), unsafe_allow_html=True)
        c4.markdown(kpi_box("Frete Hist.", fmt_brl(row["FreteHistProj"])), unsafe_allow_html=True)
        c5.markdown(kpi_box("Frete Cot.",  fmt_brl(row["FreteCotDest"]) if pd.notna(row["FreteCotDest"]) else "sem cot."), unsafe_allow_html=True)
        c6.markdown(kpi_box("Delta Hist.",
            fmt_brl(abs(dh)) if pd.notna(dh) else "—",
            cor_valor=COR_OK if (dh or 0)<=0 else COR_ALERTA,
            delta="ganho" if (dh or 0)<=0 else "perda",
            delta_cor=COR_OK if (dh or 0)<=0 else COR_ALERTA), unsafe_allow_html=True)
        c7.markdown(kpi_box("Delta Cot.",
            fmt_brl(abs(dc)) if pd.notna(dc) else "sem cot.",
            cor_valor=COR_OK if (dc or 0)<=0 else COR_ALERTA,
            delta="ganho" if (dc or 0)<=0 else "perda",
            delta_cor=COR_OK if (dc or 0)<=0 else COR_ALERTA), unsafe_allow_html=True)
        st.markdown("<div style='margin-bottom:4px'></div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.divider()

    # ── SALDO FINAL — 5 cartões ──────────────────────────
    st.markdown("### Saldo Final da Migracao")

    def cartao(titulo, valor_fmt, descricao, ganho):
        cor   = COR_OK if ganho else COR_ALERTA
        label = "ganho" if ganho else "perda"
        return f"""<div style="background:#fff;border-radius:12px;padding:16px;
            border-left:5px solid {cor};text-align:center;min-height:120px">
        <div style="font-size:11px;color:#888">{titulo}</div>
        <div style="font-size:20px;font-weight:700;color:{cor}">{valor_fmt}</div>
        <div style="font-size:11px;color:{cor};margin-top:4px;font-weight:600">{label}</div>
        <div style="font-size:10px;color:#aaa;margin-top:2px">{descricao}</div>
        </div>"""

    s1,s2,s3,s4,s5 = st.columns(5)

    s1.markdown(cartao(
        "Frete Historico",
        fmt_brl(abs(ganho_frete_hist)),
        f"Hist proj {fmt_brl(total_frete_hist_proj)} vs orig {fmt_brl(total_frete_orig_prop)}",
        ganho_frete_hist >= 0
    ), unsafe_allow_html=True)

    s2.markdown(cartao(
        "Frete Cotacao",
        fmt_brl(abs(ganho_frete_cot)) if ganho_frete_cot is not None else "sem dados",
        f"Cot dest {fmt_brl(total_frete_cot_dest)} vs orig {fmt_brl(total_frete_orig_prop)}",
        (ganho_frete_cot or 0) >= 0
    ), unsafe_allow_html=True)

    s3.markdown(cartao(
        "Impacto NFD (R$)",
        fmt_brl(abs(ganho_nfd)),
        f"NFD atual {fmt_brl(total_nfd_orig_val)} → dest {fmt_brl(total_nfd_dest)}",
        ganho_nfd >= 0
    ), unsafe_allow_html=True)

    cor_sla   = COR_OK if delta_sla >= 0 else COR_ALERTA
    label_sla = "ganho" if delta_sla >= 0 else "perda"
    s4.markdown(f"""<div style="background:#fff;border-radius:12px;padding:16px;
        border-left:5px solid {cor_sla};text-align:center;min-height:120px">
    <div style="font-size:11px;color:#888">Impacto SLA</div>
    <div style="font-size:20px;font-weight:700;color:{cor_sla}">{delta_sla:+.2f}pp</div>
    <div style="font-size:11px;color:{cor_sla};margin-top:4px;font-weight:600">{label_sla}</div>
    <div style="font-size:10px;color:#aaa;margin-top:2px">{fmt_pct(sla_orig_pond)} → {fmt_pct(sla_dest_pond)}</div>
    </div>""", unsafe_allow_html=True)

    cor_saldo = COR_OK if saldo >= 0 else COR_ALERTA
    label_sal = "GANHO LIQUIDO" if saldo >= 0 else "PERDA LIQUIDA"
    s5.markdown(f"""<div style="background:{cor_saldo}18;border-radius:12px;padding:16px;
        border:2px solid {cor_saldo};text-align:center;min-height:120px">
    <div style="font-size:11px;color:{cor_saldo};font-weight:700">{label_sal}</div>
    <div style="font-size:24px;font-weight:800;color:{cor_saldo}">{fmt_brl(abs(saldo))}</div>
    <div style="font-size:10px;color:{cor_saldo};margin-top:4px">Ganho NFD + Ganho Frete Hist.</div>
    </div>""", unsafe_allow_html=True)

    # ====================================================
    # PROJEÇÃO ATÉ FIM DO ANO
    # ====================================================

    st.divider()
    st.markdown("### 📅 Projeção até 31/12/2026")

    from datetime import date
    hoje        = date.today()
    fim_ano     = date(2026, 12, 31)
    dias_rest   = (fim_ano - hoje).days
    dias_period = dias

    def proj(valor, dias_base, dias_futuros):
        if not valor or dias_base == 0:
            return None
        return (valor / dias_base) * dias_futuros

    # ── BLOCO 1 — Com os filtros aplicados ───────────────
    st.markdown(f"#### Com os filtros aplicados — {transp_sel} | Codigos selecionados")
    st.caption(
        f"Base: ultimos {dias_period} dias ({dias_period} dias) → "
        f"media diaria × {dias_rest} dias restantes ate 31/12"
    )

    proj_ganho_nfd       = proj(ganho_nfd,        dias_period, dias_rest)
    proj_ganho_frete_hist= proj(ganho_frete_hist, dias_period, dias_rest)
    proj_ganho_frete_cot = proj(ganho_frete_cot,  dias_period, dias_rest) if ganho_frete_cot else None
    proj_delta_sla       = delta_sla  # SLA é percentual, não acumula
    proj_saldo           = (proj_ganho_nfd or 0) + (proj_ganho_frete_hist or 0)

    p1,p2,p3,p4,p5 = st.columns(5)

    def proj_cartao(col, titulo, valor, ganho, descricao=""):
        cor   = COR_OK if ganho else COR_ALERTA
        label = "ganho proj." if ganho else "perda proj."
        col.markdown(
            f"""<div style="background:#fff;border-radius:12px;padding:16px;
                border-left:5px solid {cor};text-align:center">
            <div style="font-size:11px;color:#888">{titulo}</div>
            <div style="font-size:20px;font-weight:700;color:{cor}">{valor}</div>
            <div style="font-size:11px;color:{cor};margin-top:3px;font-weight:600">{label}</div>
            <div style="font-size:10px;color:#aaa;margin-top:2px">{descricao}</div>
            </div>""",
            unsafe_allow_html=True
        )

    proj_cartao(p1, "NFD (R$)",
        fmt_brl(abs(proj_ganho_nfd)) if proj_ganho_nfd else "—",
        (proj_ganho_nfd or 0) >= 0,
        f"{fmt_brl(ganho_nfd / dias_period)}/dia" if ganho_nfd else ""
    )
    proj_cartao(p2, "Frete Historico",
        fmt_brl(abs(proj_ganho_frete_hist)) if proj_ganho_frete_hist else "—",
        (proj_ganho_frete_hist or 0) >= 0,
        f"{fmt_brl(ganho_frete_hist / dias_period)}/dia" if ganho_frete_hist else ""
    )
    proj_cartao(p3, "Frete Cotacao",
        fmt_brl(abs(proj_ganho_frete_cot)) if proj_ganho_frete_cot else "sem dados",
        (proj_ganho_frete_cot or 0) >= 0,
        f"{fmt_brl(ganho_frete_cot / dias_period)}/dia" if ganho_frete_cot else ""
    )

    cor_sla_p = COR_OK if proj_delta_sla >= 0 else COR_ALERTA
    p4.markdown(
        f"""<div style="background:#fff;border-radius:12px;padding:16px;
            border-left:5px solid {cor_sla_p};text-align:center">
        <div style="font-size:11px;color:#888">SLA</div>
        <div style="font-size:20px;font-weight:700;color:{cor_sla_p}">{proj_delta_sla:+.2f}pp</div>
        <div style="font-size:11px;color:{cor_sla_p};margin-top:3px;font-weight:600">
            {"ganho proj." if proj_delta_sla >= 0 else "perda proj."}</div>
        <div style="font-size:10px;color:#aaa;margin-top:2px">{fmt_pct(sla_orig_pond)} → {fmt_pct(sla_dest_pond)}</div>
        </div>""",
        unsafe_allow_html=True
    )

    cor_sp = COR_OK if proj_saldo >= 0 else COR_ALERTA
    label_sp = "GANHO PROJ." if proj_saldo >= 0 else "PERDA PROJ."
    p5.markdown(
        f"""<div style="background:{cor_sp}18;border-radius:12px;padding:16px;
            border:2px solid {cor_sp};text-align:center">
        <div style="font-size:11px;color:{cor_sp};font-weight:700">{label_sp}</div>
        <div style="font-size:22px;font-weight:800;color:{cor_sp}">{fmt_brl(abs(proj_saldo))}</div>
        <div style="font-size:10px;color:{cor_sp};margin-top:3px">NFD + Frete Hist. projetados</div>
        </div>""",
        unsafe_allow_html=True
    )

    st.divider()

    # ── BLOCO 2 — Empresa toda (10 piores de todas as TSPs) ──
    st.markdown("#### Empresa toda — 10 piores codigos (todas as transportadoras)")
    st.caption(
        f"Baseado nos ultimos 90 dias | "
        f"Media diaria × {dias_rest} dias restantes ate 31/12"
    )

    # Recalcula para todas as transportadoras, últimos 90 dias
    data_corte_90 = pd.Timestamp.today() - pd.Timedelta(days=90)
    df_emp = df_trade[df_trade["DataFinal"] >= data_corte_90].copy()

    nfd_emp = (
        df_emp.groupby(["Transportadora", "CodigoTarifario"])
        .agg(
            Pedidos    = ("TemNFD",      "count"),
            NFD_n      = ("TemNFD",      "sum"),
            SLA_pct    = ("DentroPrazo", "mean"),
            ValorNotas = ("ValorNota",   "sum"),
            ValorFrete = ("ValorFrete",  "sum"),
        ).reset_index()
    )
    nfd_emp = nfd_emp[nfd_emp["Pedidos"] >= 30].copy()
    nfd_emp["NFD_pct"]  = (nfd_emp["NFD_n"] / nfd_emp["Pedidos"] * 100).round(2)
    nfd_emp["SLA_pct"]  = (nfd_emp["SLA_pct"] * 100).round(2)
    nfd_emp["ValorNFD"] = nfd_emp["ValorNotas"] * nfd_emp["NFD_pct"] / 100
    top10_emp = nfd_emp.sort_values("NFD_pct", ascending=False).head(10)

    # Para cada código, busca o destino na base de similaridade
    total_nfd_emp       = 0
    total_frete_emp     = 0
    total_nfd_dest_emp  = 0
    total_fh_dest_emp   = 0
    total_fc_dest_emp   = 0
    total_fc_orig_emp   = 0
    sla_orig_emp_pond   = 0
    sla_dest_emp_pond   = 0
    total_ped_emp       = 0

    for _, r in top10_emp.iterrows():
        tsp_o = r["Transportadora"]
        cod_o = r["CodigoTarifario"]
        ped_o = r["Pedidos"]
        nfd_v = r["ValorNFD"]
        fr_o  = r["ValorFrete"]
        sla_o = r["SLA_pct"]

        total_nfd_emp   += nfd_v or 0
        total_frete_emp += fr_o  or 0
        sla_orig_emp_pond += (sla_o or 0) * ped_o
        total_ped_emp   += ped_o

        df_s = df_sim[
            (df_sim["TransportadoraOrigem"] == tsp_o) &
            (df_sim["CodigoOrigem"] == cod_o)
        ]
        if df_s.empty:
            total_nfd_dest_emp += nfd_v or 0
            total_fh_dest_emp  += fr_o  or 0
            sla_dest_emp_pond  += (sla_o or 0) * ped_o
            continue

        ped_sim_total = df_s["Pedidos"].sum()
        for _, sd in df_s.iterrows():
            w = sd["Pedidos"] / ped_sim_total if ped_sim_total > 0 else 0
            notas_p   = (r["ValorNotas"] or 0) * w
            nfd_d_val = sd["NFD_Hist"] if "NFD_Hist" in sd.index and pd.notna(sd["NFD_Hist"]) else 0
            nfd_d     = float(nfd_d_val) * 100
            total_nfd_dest_emp += notas_p * nfd_d / 100

            fh = sd["ProjecaoFreteHist"] if "ProjecaoFreteHist" in sd.index and pd.notna(sd["ProjecaoFreteHist"]) else 0
            total_fh_dest_emp  += float(fh)

            fc_dest = sd["ValorCotacaoDestTotal"] if "ValorCotacaoDestTotal" in sd.index and pd.notna(sd["ValorCotacaoDestTotal"]) else 0
            fc_orig = sd["ValorCotacaoOrigTotal"]  if "ValorCotacaoOrigTotal"  in sd.index and pd.notna(sd["ValorCotacaoOrigTotal"])  else 0
            total_fc_dest_emp  += float(fc_dest)
            total_fc_orig_emp  += float(fc_orig)

            sla_d_val = sd["SLA_Hist"] if "SLA_Hist" in sd.index and pd.notna(sd["SLA_Hist"]) else 0
            sla_dest_emp_pond  += float(sla_d_val) * 100 * sd["Pedidos"]
            total_ped_dest_emp = total_ped_dest_emp + sd["Pedidos"] if "total_ped_dest_emp" in dir() else sd["Pedidos"]

    total_ped_dest_emp = locals().get("total_ped_dest_emp", total_ped_emp)
    sla_orig_emp_m = sla_orig_emp_pond / total_ped_emp      if total_ped_emp      > 0 else 0
    sla_dest_emp_m = sla_dest_emp_pond / total_ped_dest_emp if total_ped_dest_emp > 0 else 0
    delta_sla_emp  = sla_dest_emp_m - sla_orig_emp_m

    ganho_nfd_emp  = total_nfd_emp   - total_nfd_dest_emp
    ganho_fh_emp   = (total_frete_emp - total_fh_dest_emp) if total_fh_dest_emp > 0 else None
    ganho_fc_emp   = total_fc_orig_emp - total_fc_dest_emp if total_fc_orig_emp > 0 else None
    saldo_emp      = (ganho_nfd_emp or 0) + (ganho_fh_emp or 0)

    # Projeções (base 90 dias)
    pp_nfd  = proj(ganho_nfd_emp,  90, dias_rest)
    pp_fh   = proj(ganho_fh_emp,   90, dias_rest)
    pp_fc   = proj(ganho_fc_emp,   90, dias_rest) if ganho_fc_emp else None
    pp_sal  = (pp_nfd or 0) + (pp_fh or 0)

    q1,q2,q3,q4,q5 = st.columns(5)

    proj_cartao(q1, "NFD (R$)",
        fmt_brl(abs(pp_nfd)) if pp_nfd else "—",
        (pp_nfd or 0) >= 0,
        f"{fmt_brl(ganho_nfd_emp/90)}/dia"
    )
    proj_cartao(q2, "Frete Historico",
        fmt_brl(abs(pp_fh)) if pp_fh else "—",
        (pp_fh or 0) >= 0,
        f"{fmt_brl(ganho_fh_emp/90)}/dia"
    )
    proj_cartao(q3, "Frete Cotacao",
        fmt_brl(abs(pp_fc)) if pp_fc else "sem dados",
        (pp_fc or 0) >= 0,
        f"{fmt_brl(ganho_fc_emp/90)}/dia" if ganho_fc_emp else ""
    )

    cor_sla_e = COR_OK if delta_sla_emp >= 0 else COR_ALERTA
    q4.markdown(
        f"""<div style="background:#fff;border-radius:12px;padding:16px;
            border-left:5px solid {cor_sla_e};text-align:center">
        <div style="font-size:11px;color:#888">SLA</div>
        <div style="font-size:20px;font-weight:700;color:{cor_sla_e}">{delta_sla_emp:+.2f}pp</div>
        <div style="font-size:11px;color:{cor_sla_e};margin-top:3px;font-weight:600">
            {"ganho proj." if delta_sla_emp >= 0 else "perda proj."}</div>
        <div style="font-size:10px;color:#aaa;margin-top:2px">{fmt_pct(sla_orig_emp_m)} → {fmt_pct(sla_dest_emp_m)}</div>
        </div>""",
        unsafe_allow_html=True
    )

    cor_se = COR_OK if pp_sal >= 0 else COR_ALERTA
    label_se = "GANHO PROJ." if pp_sal >= 0 else "PERDA PROJ."
    q5.markdown(
        f"""<div style="background:{cor_se}18;border-radius:12px;padding:16px;
            border:2px solid {cor_se};text-align:center">
        <div style="font-size:11px;color:{cor_se};font-weight:700">{label_se}</div>
        <div style="font-size:22px;font-weight:800;color:{cor_se}">{fmt_brl(abs(pp_sal))}</div>
        <div style="font-size:10px;color:{cor_se};margin-top:3px">NFD + Frete Hist. projetados</div>
        </div>""",
        unsafe_allow_html=True
    )

    st.caption(
        f"10 piores codigos: " +
        ", ".join(f"{r['Transportadora']}/{r['CodigoTarifario']}" for _, r in top10_emp.iterrows())
    )
    # ====================================================
    # SECAO GANHO LIQUIDO GLOBAL (todos os códigos)
    # ====================================================
    st.divider()
    st.markdown("## 💰 Exportar Ganho Líquido — Todos os Códigos Tarifários")
    st.caption(f"Calcula para todos os códigos da {transp_sel} no período de {periodo}. "
               f"Exporta apenas os que têm ganho líquido positivo.")

    if st.button("🔍 Calcular ganho líquido para todos os códigos"):
        todos_codigos = sorted(df_periodo["CodigoTarifario"].dropna().unique())
        ganhos_todos  = []

        prog = st.progress(0, text="Calculando...")
        for i, cod in enumerate(todos_codigos):
            prog.progress((i + 1) / len(todos_codigos), text=f"Processando {cod}...")

            df_c = df_periodo[df_periodo["CodigoTarifario"] == cod]
            pedidos_c     = len(df_c)
            nfd_pct_c     = df_c["TemNFD"].mean() * 100 if pedidos_c > 0 else 0
            valor_notas_c = df_c["ValorNota"].sum() if "ValorNota" in df_c.columns and df_c["ValorNota"].notna().any() else None
            nfd_orig_c    = (valor_notas_c * nfd_pct_c / 100) if valor_notas_c else None

            df_sim_c = df_sim[
                (df_sim["TransportadoraOrigem"] == transp_sel) &
                (df_sim["CodigoOrigem"] == cod)
            ]
            if df_sim_c.empty or nfd_orig_c is None:
                continue

            rows_c = []
            for tsp_dest, grp in df_sim_c.groupby("TransportadoraDestino"):
                ped_sim      = int(grp["Pedidos"].sum())
                if ped_sim == 0: continue
                w            = grp["Pedidos"]
                nfd_d        = (grp["NFD_Hist"].fillna(0) * w).sum() / w.sum() * 100 if "NFD_Hist" in grp.columns else None
                cot_dest     = grp["ValorCotacaoDestTotal"].sum() if "ValorCotacaoDestTotal" in grp.columns else None
                cot_orig     = grp["ValorCotacaoOrigTotal"].sum()  if "ValorCotacaoOrigTotal" in grp.columns else None
                delta_frete  = (cot_dest - cot_orig) if (cot_dest is not None and cot_orig is not None) else None
                pct_dist     = ped_sim / df_sim_c["Pedidos"].sum() * 100
                notas_dest   = (valor_notas_c * pct_dist / 100) if valor_notas_c else None
                nfd_val_dest = (notas_dest * nfd_d / 100) if (notas_dest and nfd_d is not None) else None
                if delta_frete is not None and nfd_val_dest is not None:
                    rows_c.append({"DeltaFreteCot": delta_frete, "NFDValorDest": nfd_val_dest})

            if not rows_c:
                continue

            df_rc         = pd.DataFrame(rows_c)
            delta_frete_t = df_rc["DeltaFreteCot"].sum()
            nfd_dest_t    = df_rc["NFDValorDest"].sum()
            ganho         = delta_frete_t - nfd_dest_t + nfd_orig_c

            if ganho > 0:
                ganhos_todos.append({
                    "Transportadora":    transp_sel,
                    "CodigoTarifario":   cod,
                    "Pedidos":           pedidos_c,
                    "NFD_Atual_R$":      nfd_orig_c,
                    "DeltaFrete_Cot_R$": delta_frete_t,
                    "NFD_Destinos_R$":   nfd_dest_t,
                    "GanhoLiquido_R$":   ganho,
                })

        prog.empty()

        if ganhos_todos:
            df_ganhos_todos = pd.DataFrame(ganhos_todos).sort_values("GanhoLiquido_R$", ascending=False)
            import io
            buf_todos = io.BytesIO()
            df_ganhos_todos.to_excel(buf_todos, index=False)
            buf_todos.seek(0)
            st.success(f"{len(df_ganhos_todos)} código(s) com ganho líquido positivo de {len(todos_codigos)} analisados.")
            st.download_button(
                "⬇️ Exportar todos os ganhos positivos",
                data=buf_todos,
                file_name=f"ganho_liquido_todos_{transp_sel}_{periodo.replace(' ','')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nenhum código tarifário com ganho líquido positivo no período selecionado.")

    # ====================================================
    # SECAO DEBUG (fim da pagina)
    # ====================================================
    st.divider()
    st.markdown("## 🔧 Debug")
    st.caption("Indicadores de apoio e qualidade de dados — nao entram nas analises principais.")

    # Recalcula o conjunto "TSP mas nao e TSP" (caso a base tenha a flag)
    if not df_nfd_real.empty and "TspMasNaoTsp" in df_nfd_real.columns:
        nfd_tsp_mas_nao = df_nfd_real[df_nfd_real["TspMasNaoTsp"] == True].copy()
    else:
        nfd_tsp_mas_nao = pd.DataFrame()

    total_tsp_mas_nao = nfd_tsp_mas_nao["ValorNota"].sum() if not nfd_tsp_mas_nao.empty else 0
    qtd_tsp_mas_nao   = len(nfd_tsp_mas_nao)

    dbg1, dbg2, dbg3 = st.columns(3)

    # Card: NFs sem Valor (movido do topo)
    dbg1.markdown(kpi_box(
        "NFs sem Valor (Intelipost)",
        fmt_int(pedidos_sem_valor) if pedidos_sem_valor is not None else "sem dado",
        cor_valor=COR_NEUTRO
    ), unsafe_allow_html=True)

    # Card: Categorizado como TSP mas nao e TSP
    dbg2.markdown(kpi_box(
        "Categorizado como TSP mas nao e TSP (R$)",
        fmt_brl(total_tsp_mas_nao) if total_tsp_mas_nao else "sem dado",
        cor_valor=COR_ALERTA
    ), unsafe_allow_html=True)

    # Card: NFD sem transportadora (nem Intelipost nem PRW identificou)
    nfd_sem_tsp_total = 0
    if not df_intel.empty and "NFDSemTransportadora" in df_intel.columns:
        nfd_sem_tsp_total = int(df_intel["NFDSemTransportadora"].iloc[0])
    nfd_prw_fallback = int(df_intel["NFDTranspPRWFallback"].iloc[0]) if not df_intel.empty and "NFDTranspPRWFallback" in df_intel.columns else 0
    dbg3.markdown(kpi_box(
        "NFD sem Transportadora",
        fmt_int(nfd_sem_tsp_total) if nfd_sem_tsp_total else "0",
        cor_valor=COR_NEUTRO
    ), unsafe_allow_html=True)
    if nfd_prw_fallback > 0:
        st.caption(f"ℹ️ {nfd_prw_fallback:,} NFDs tiveram transportadora identificada via PRW (fallback).")

    st.markdown(
        f"<div style='font-size:12px;color:#888;margin-top:6px'>"
        f"{qtd_tsp_mas_nao:,} notas tiveram NFD TSP mas passaram por status de retirada, "
        f"endereco, ausencia ou acareacao (21 status). Sao retiradas dos totais de TSP da Visao Geral.</div>",
        unsafe_allow_html=True
    )

    # Botao exportar — PedidoFormatado + valor (sem danfe, conforme LGPD)
    if not nfd_tsp_mas_nao.empty:
        import io
        cols_export = [c for c in ["PedidoFormatado", "ValorNota", "Transportadora",
                                   "DataDespacho", "MotivoDevolucao", "StatusNaoTsp", "CruzouIntelipost"]
                       if c in nfd_tsp_mas_nao.columns]
        df_export_debug = nfd_tsp_mas_nao[cols_export].copy()
        buf_dbg = io.BytesIO()
        df_export_debug.to_excel(buf_dbg, index=False)
        buf_dbg.seek(0)
        st.download_button(
            "⬇️ Exportar pedidos (categorizados como TSP mas nao sao)",
            data=buf_dbg,
            file_name="tsp_mas_nao_e_tsp.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )