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
    df_trade = pd.read_excel(arq_pedidos) if os.path.exists(arq_pedidos) else pd.DataFrame()
    df_sim   = pd.read_excel(arq_sim)     if os.path.exists(arq_sim)     else pd.DataFrame()
    if not df_trade.empty:
        df_trade.columns = df_trade.columns.str.strip()
        if "DataFinal"   in df_trade.columns: df_trade["DataFinal"]   = pd.to_datetime(df_trade["DataFinal"],   errors="coerce")
        if "TemNFD"      in df_trade.columns: df_trade["TemNFD"]      = df_trade["TemNFD"].fillna(False).astype(bool)
        if "DentroPrazo" in df_trade.columns: df_trade["DentroPrazo"] = df_trade["DentroPrazo"].fillna(False).astype(bool)
        if "ValorFrete"  in df_trade.columns: df_trade["ValorFrete"]  = pd.to_numeric(df_trade["ValorFrete"],  errors="coerce")
        if "ValorNota"   in df_trade.columns: df_trade["ValorNota"]   = pd.to_numeric(df_trade["ValorNota"],   errors="coerce")
    return df_trade, df_sim

def render_tradeoff():

    df_trade, df_sim = carregar_bases()
    if df_trade.empty or df_sim.empty:
        st.error("Bases nao encontradas. Execute o ETL primeiro.")
        return

    # ── FILTROS ──────────────────────────────────────────
    col1, col2 = st.columns(2)
    with col1:
        transportadoras = sorted(df_trade["Transportadora"].dropna().unique())
        transp_sel = st.selectbox("Transportadora", transportadoras)
    with col2:
        periodo = st.selectbox("Periodo", ["30 dias", "60 dias", "90 dias"])
        dias = {"30 dias": 30, "60 dias": 60, "90 dias": 90}[periodo]

    data_corte = pd.Timestamp.today() - pd.Timedelta(days=dias)
    df_periodo = df_trade[
        (df_trade["Transportadora"] == transp_sel) &
        (df_trade["DataFinal"] >= data_corte)
    ].copy()

    if df_periodo.empty:
        st.warning("Nenhum pedido encontrado para esse filtro.")
        return

    st.divider()

    # ── VISÃO GERAL DA TRANSPORTADORA ────────────────────
    st.markdown(f"### Visao Geral — {transp_sel} | Ultimos {dias} dias")

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
    g5.markdown(kpi_box("% NFD",              fmt_pct(nfd_pct_ger),
        cor_valor=COR_ALERTA if nfd_pct_ger>=5 else COR_NEUTRO if nfd_pct_ger>=2 else COR_OK),
        unsafe_allow_html=True)

    st.divider()

    # ── VISÃO POR TRANSPORTADORA ──────────────────────────
    st.markdown(f"### NFD por Transportadora | Ultimos {dias} dias")

    df_todas = df_trade[df_trade["DataFinal"] >= data_corte].copy()

    resumo_tsp = (
        df_todas.groupby("Transportadora")
        .agg(
            Pedidos    = ("TemNFD",    "count"),
            NFD_n      = ("TemNFD",    "sum"),
            ValorNotas = ("ValorNota", "sum"),
        ).reset_index()
    )
    resumo_tsp["NFD_pct"]  = (resumo_tsp["NFD_n"] / resumo_tsp["Pedidos"] * 100).round(2)
    resumo_tsp["ValorNFD"] = resumo_tsp["ValorNotas"] * resumo_tsp["NFD_pct"] / 100
    resumo_tsp = resumo_tsp.sort_values("NFD_pct", ascending=False)

    df_exib = resumo_tsp.copy()
    df_exib = df_exib.rename(columns={
        "Pedidos":    "Pedidos",
        "ValorNotas": "Valor Vendas (R$)",
        "NFD_n":      "Pedidos NFD",
        "ValorNFD":   "Valor NFD (R$)",
        "NFD_pct":    "% NFD",
    })[["Transportadora","Pedidos","Valor Vendas (R$)","Pedidos NFD","Valor NFD (R$)","% NFD"]]

    st.dataframe(
        df_exib.style
        .format({
            "Pedidos":          "{:,.0f}",
            "Valor Vendas (R$)":"R$ {:,.2f}",
            "Pedidos NFD":      "{:,.0f}",
            "Valor NFD (R$)":   "R$ {:,.2f}",
            "% NFD":            "{:.2f}%",
        })
        .background_gradient(subset=["% NFD"], cmap="RdYlGn_r"),
        use_container_width=True,
        hide_index=True
    )

    st.divider()

    # ── RANKING ──────────────────────────────────────────
    st.markdown("### Ranking — Piores Codigos Tarifarios por NFD")
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

    resumo_codigos = []

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
            c1,c2,c3,c4,c5 = st.columns(5)
            c1.markdown(kpi_box("Pedidos",    fmt_int(pedidos_orig)), unsafe_allow_html=True)
            c2.markdown(kpi_box("SLA",        fmt_pct(sla_orig),    cor_valor=COR_OK if sla_orig>=95 else COR_ALERTA), unsafe_allow_html=True)
            c3.markdown(kpi_box("TM Frete",   fmt_brl(tm_orig)),    unsafe_allow_html=True)
            c4.markdown(kpi_box("NFD %",      fmt_pct(nfd_orig_pct),cor_valor=COR_ALERTA if nfd_orig_pct>=5 else COR_NEUTRO if nfd_orig_pct>=2 else COR_OK), unsafe_allow_html=True)
            c5.markdown(kpi_box("Valor Notas",fmt_brl(valor_notas) if valor_notas else "sem dado"), unsafe_allow_html=True)

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

            st.markdown(f"**Para onde iriam os {total_redistribuido:,} pedidos (base cotacao)**")
            st.caption("Destino = mais barato no leilao excluindo a transportadora atual | Historico = dados reais de entrega")

            for _, row in grp_dest.iterrows():
                st.markdown(
                    f"<div style='background:#e8f0fe;border-radius:8px;padding:6px 12px;"
                    f"margin-bottom:6px;font-size:12px;font-weight:600;color:#1f4e79'>"
                    f"{row['TransportadoraDestino']} &nbsp; {int(row['PedidosSimulados']):,} pedidos ({row['Pct_dist']:.1f}%)</div>",
                    unsafe_allow_html=True
                )

                d1,d2,d3,d4,d5,d6,d7 = st.columns(7)

                delta_sla_v = (row["SLADestino"] - sla_orig) if row["SLADestino"] is not None else None
                delta_nfd_v = (row["NFDDestino"] - nfd_orig_pct) if row["NFDDestino"] is not None else None
                dh = row["DeltaFreteHist"]
                dc = row["DeltaFreteCot"]

                d1.markdown(kpi_box("SLA Hist.",
                    fmt_pct(row["SLADestino"]),
                    cor_valor=COR_OK if (row["SLADestino"] or 0)>=95 else COR_ALERTA,
                    delta=f"{delta_sla_v:+.1f}pp" if delta_sla_v is not None else None,
                    delta_cor=COR_OK if (delta_sla_v or 0)>=0 else COR_ALERTA), unsafe_allow_html=True)

                d2.markdown(kpi_box("NFD % Hist.",
                    fmt_pct(row["NFDDestino"]),
                    cor_valor=COR_ALERTA if (row["NFDDestino"] or 0)>=5 else COR_NEUTRO if (row["NFDDestino"] or 0)>=2 else COR_OK,
                    delta=f"{delta_nfd_v:+.1f}pp" if delta_nfd_v is not None else None,
                    delta_cor=COR_OK if (delta_nfd_v or 0)<=0 else COR_ALERTA), unsafe_allow_html=True)

                d3.markdown(kpi_box("TM Hist.", fmt_brl(row["TM_Hist"])), unsafe_allow_html=True)

                d4.markdown(kpi_box("TM Cotacao", fmt_brl(row["TM_Cotacao"]) if row["TM_Cotacao"] else "sem cot."), unsafe_allow_html=True)

                d5.markdown(kpi_box("Delta Frete Hist.",
                    fmt_brl(abs(dh)) if dh is not None else "—",
                    cor_valor=COR_OK if (dh or 0)<=0 else COR_ALERTA,
                    delta="ganho" if (dh or 0)<=0 else "perda",
                    delta_cor=COR_OK if (dh or 0)<=0 else COR_ALERTA), unsafe_allow_html=True)

                d6.markdown(kpi_box("Delta Frete Cot.",
                    fmt_brl(abs(dc)) if dc is not None else "—",
                    cor_valor=COR_OK if (dc or 0)<=0 else COR_ALERTA,
                    delta="ganho" if (dc or 0)<=0 else "perda",
                    delta_cor=COR_OK if (dc or 0)<=0 else COR_ALERTA), unsafe_allow_html=True)

                d7.markdown(kpi_box("NFD R$ Est.",
                    fmt_brl(row["NFDValorDest"]) if pd.notna(row.get("NFDValorDest")) else "sem dado"), unsafe_allow_html=True)

                st.markdown("<div style='margin-bottom:8px'></div>", unsafe_allow_html=True)

            resumo_codigos.append({
                "codigo":codigo, "pedidos_orig":pedidos_orig,
                "sla_orig":sla_orig, "tm_orig":tm_orig,
                "nfd_orig_pct":nfd_orig_pct, "frete_orig":frete_orig,
                "valor_notas":valor_notas, "nfd_orig_valor":nfd_orig_valor,
                "destinos":grp_dest.to_dict("records")
            })

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