import os
import streamlit as st
import pandas as pd
import plotly.graph_objects as go

# ======================================================
# CORES
# ======================================================

COR_ORIGEM = "#1f4e79"
COR_ALERTA = "#e63946"
COR_OK     = "#2a9d8f"
COR_NEUTRO = "#f4a261"
COR_FUNDO  = "#f4f6f9"

# ======================================================
# HELPERS
# ======================================================

def fmt_pct(v):
    if pd.isna(v):
        return "—"
    return f"{v:.2f}%"

def fmt_brl(v):
    if pd.isna(v):
        return "—"
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_int(v):
    if pd.isna(v):
        return "—"
    return f"{int(v):,}"

def kpi_box(label, valor, cor_valor="#0f2a44", fundo="#ffffff", delta=None, delta_cor=None):
    delta_html = ""
    if delta is not None:
        dcor = delta_cor or ("#27ae60" if "+" in str(delta) else "#e74c3c")
        delta_html = f'<div style="font-size:11px;color:{dcor};margin-top:2px">{delta}</div>'
    return f"""
    <div style="background:{fundo};border-radius:10px;padding:12px 14px;
                border:1px solid #e0e4ea;text-align:center">
        <div style="font-size:11px;color:#888;margin-bottom:4px">{label}</div>
        <div style="font-size:18px;font-weight:700;color:{cor_valor}">{valor}</div>
        {delta_html}
    </div>"""

# ======================================================
# CARREGAR DADOS
# ======================================================

@st.cache_data(ttl=120)
def carregar_bases():
    arq_pedidos = "data/Base_Pedidos_Codigo.xlsx"
    arq_sim     = "data/Base_Similaridade_Tarifarios.xlsx"

    df_trade = pd.read_excel(arq_pedidos) if os.path.exists(arq_pedidos) else pd.DataFrame()
    df_sim   = pd.read_excel(arq_sim)     if os.path.exists(arq_sim)     else pd.DataFrame()

    if not df_trade.empty:
        df_trade.columns = df_trade.columns.str.strip()

        if "DataFinal" in df_trade.columns:
            df_trade["DataFinal"] = pd.to_datetime(df_trade["DataFinal"], errors="coerce")

        if "TemNFD" in df_trade.columns:
            df_trade["TemNFD"] = df_trade["TemNFD"].fillna(False).astype(bool)

        if "DentroPrazo" in df_trade.columns:
            df_trade["DentroPrazo"] = df_trade["DentroPrazo"].fillna(False).astype(bool)

        if "ValorFrete" in df_trade.columns:
            df_trade["ValorFrete"] = pd.to_numeric(df_trade["ValorFrete"], errors="coerce")

        if "ValorNota" in df_trade.columns:
            df_trade["ValorNota"] = pd.to_numeric(df_trade["ValorNota"], errors="coerce")

    return df_trade, df_sim


# ======================================================
# RENDER PRINCIPAL
# ======================================================

def render_tradeoff():

    df_trade, df_sim = carregar_bases()

    if df_trade.empty or df_sim.empty:
        st.error("Bases nao encontradas. Execute o ETL primeiro.")
        return

    # ====================================================
    # FILTROS
    # ====================================================

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

    # ====================================================
    # SECAO 1 - RANKING
    # ====================================================

    st.markdown("### Ranking - Piores Codigos Tarifarios por NFD")
    st.caption(f"Baseado nos ultimos {dias} dias | Minimo 30 pedidos")

    nfd_rank = (
        df_periodo.groupby("CodigoTarifario")
        .agg(
            Pedidos    = ("TemNFD",      "count"),
            NFD_n      = ("TemNFD",      "sum"),
            SLA_pct    = ("DentroPrazo", "mean"),
            TM         = ("ValorFrete",  "mean"),
            ValorNotas = ("ValorNota",   "sum"),
        )
        .reset_index()
    )
    nfd_rank = nfd_rank[nfd_rank["Pedidos"] >= 30].copy()
    nfd_rank["NFD_pct"] = (nfd_rank["NFD_n"] / nfd_rank["Pedidos"] * 100).round(2)
    nfd_rank["SLA_pct"] = (nfd_rank["SLA_pct"] * 100).round(2)
    nfd_rank = nfd_rank.sort_values("NFD_pct", ascending=False).reset_index(drop=True)

    if nfd_rank.empty:
        st.warning("Nenhum codigo com pedidos suficientes no periodo.")
        return

    if "pagina_nfd" not in st.session_state:
        st.session_state["pagina_nfd"] = 0

    por_pag    = 10
    total_pags = max(0, (len(nfd_rank) - 1) // por_pag)
    ini        = st.session_state["pagina_nfd"] * por_pag
    fim        = ini + por_pag
    pagina_df  = nfd_rank.iloc[ini:fim].copy()

    cores = [
        COR_ALERTA if v >= 5 else COR_NEUTRO if v >= 2 else COR_OK
        for v in pagina_df["NFD_pct"]
    ]
    fig = go.Figure(go.Bar(
        x           = pagina_df["NFD_pct"],
        y           = pagina_df["CodigoTarifario"],
        orientation = "h",
        marker_color= cores,
        text        = [
            f"{v:.2f}% ({p:,} ped.)"
            for v, p in zip(pagina_df["NFD_pct"], pagina_df["Pedidos"])
        ],
        textposition= "outside",
    ))
    fig.update_layout(
        height       = max(300, len(pagina_df) * 42),
        xaxis_title  = "NFD (%)",
        yaxis        = dict(autorange="reversed", tickfont=dict(size=12)),
        plot_bgcolor = COR_FUNDO,
        paper_bgcolor= COR_FUNDO,
        margin       = dict(l=10, r=100, t=20, b=30),
        font         = dict(family="Arial", size=12),
    )
    st.plotly_chart(fig, use_container_width=True)
    st.caption("Vermelho: NFD >= 5%  |  Laranja: 2-5%  |  Verde: < 2%")

    nav1, nav2, nav3 = st.columns([1, 5, 1])
    with nav1:
        if st.session_state["pagina_nfd"] > 0:
            if st.button("Anteriores"):
                st.session_state["pagina_nfd"] -= 1
                st.rerun()
    with nav2:
        st.caption(
            f"Exibindo {ini+1}-{min(fim, len(nfd_rank))} de {len(nfd_rank)} codigos | "
            f"Pagina {st.session_state['pagina_nfd']+1} de {total_pags+1}"
        )
    with nav3:
        if st.session_state["pagina_nfd"] < total_pags:
            if st.button("Proximos"):
                st.session_state["pagina_nfd"] += 1
                st.rerun()

    st.divider()

    # ====================================================
    # SECAO 2 - SELETOR
    # ====================================================

    st.markdown("### Selecione os codigos para analise detalhada")

    codigos_pag = pagina_df["CodigoTarifario"].tolist()
    codigos_sel = st.multiselect(
        "Codigos tarifarios (pagina atual)",
        options  = codigos_pag,
        default  = codigos_pag[:1],
        help     = "Selecione um ou mais codigos. O resumo consolidado aparece no fim."
    )

    if not codigos_sel:
        st.info("Selecione pelo menos um codigo acima.")
        return

    st.divider()

    # ====================================================
    # SECAO 3 - CARDS POR CODIGO
    # ====================================================

    st.markdown("### Detalhe por Codigo Tarifario")

    resumo_codigos = []

    for codigo in codigos_sel:

        with st.expander(f"{transp_sel} / {codigo}", expanded=True):

            df_cod = df_periodo[df_periodo["CodigoTarifario"] == codigo].copy()

            pedidos_orig  = len(df_cod)
            sla_orig      = df_cod["DentroPrazo"].mean() * 100 if pedidos_orig > 0 else 0
            tm_orig       = df_cod["ValorFrete"].mean() if pedidos_orig > 0 else 0
            nfd_orig_pct  = df_cod["TemNFD"].mean() * 100 if pedidos_orig > 0 else 0
            frete_orig    = df_cod["ValorFrete"].sum()
            valor_notas   = df_cod["ValorNota"].sum() if "ValorNota" in df_cod.columns and df_cod["ValorNota"].notna().any() else None
            nfd_orig_valor= (valor_notas * nfd_orig_pct / 100) if valor_notas else None

            # Painel origem
            st.markdown(f"**Situacao atual**")
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.markdown(kpi_box("Pedidos",    fmt_int(pedidos_orig)), unsafe_allow_html=True)
            c2.markdown(kpi_box("SLA",        fmt_pct(sla_orig),
                cor_valor=COR_OK if sla_orig >= 95 else COR_ALERTA), unsafe_allow_html=True)
            c3.markdown(kpi_box("TM Frete",   fmt_brl(tm_orig)), unsafe_allow_html=True)
            c4.markdown(kpi_box("NFD %",      fmt_pct(nfd_orig_pct),
                cor_valor=COR_ALERTA if nfd_orig_pct >= 5 else COR_NEUTRO if nfd_orig_pct >= 2 else COR_OK),
                unsafe_allow_html=True)
            c5.markdown(kpi_box("Valor Notas",
                fmt_brl(valor_notas) if valor_notas else "sem dado"), unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # Redistribuicao
            df_sim_cod = df_sim[
                (df_sim["TransportadoraOrigem"] == transp_sel) &
                (df_sim["CodigoOrigem"] == codigo)
            ].copy()

            if df_sim_cod.empty:
                st.warning("Sem dados de redistribuicao para esse codigo.")
                resumo_codigos.append({
                    "codigo": codigo, "pedidos_orig": pedidos_orig,
                    "sla_orig": sla_orig, "tm_orig": tm_orig,
                    "nfd_orig_pct": nfd_orig_pct, "frete_orig": frete_orig,
                    "valor_notas": valor_notas, "nfd_orig_valor": nfd_orig_valor,
                    "destinos": []
                })
                continue

            df_sim_cod["PedidosSimulados"] = (
                pedidos_orig * df_sim_cod["Percentual"] / 100
            ).round(0).astype(int)

            # Agrega por transportadora destino (ponderado por PedidosSimulados)
            rows_dest = []
            for tsp_dest, grp in df_sim_cod.groupby("TransportadoraDestino"):
                ped_sim  = grp["PedidosSimulados"].sum()
                if ped_sim == 0:
                    continue
                w = grp["PedidosSimulados"]
                sla_d  = (grp["SLADestino"].fillna(0)  * w).sum() / w.sum() * 100
                nfd_d  = (grp["NFDDestino"].fillna(0)  * w).sum() / w.sum() * 100
                tm_d   = (grp["TM_Destino"].fillna(0)  * w).sum() / w.sum()

                ped_cot = grp["Pedidos_Com_Cotacao_Destino"].fillna(0).sum()
                if ped_cot > 0:
                    tm_cot = (grp["TM_Cotacao_Destino"].fillna(0) * grp["Pedidos_Com_Cotacao_Destino"].fillna(0)).sum() / ped_cot
                else:
                    tm_cot = None

                frete_proj = tm_d * ped_sim
                frete_cot  = tm_cot * ped_sim if tm_cot else None
                pct_dist   = ped_sim / pedidos_orig * 100

                notas_dest  = (valor_notas * pct_dist / 100) if valor_notas else None
                nfd_val_dest= (notas_dest  * nfd_d   / 100) if notas_dest  else None

                rows_dest.append({
                    "TransportadoraDestino": tsp_dest,
                    "PedidosSimulados":      ped_sim,
                    "Pct_dist":              pct_dist,
                    "SLADestino":            sla_d,
                    "NFDDestino":            nfd_d,
                    "TM_Destino":            tm_d,
                    "TM_Cotacao_Destino":    tm_cot,
                    "FreteProj":             frete_proj,
                    "FreteCot":              frete_cot,
                    "ValorNotasDest":        notas_dest,
                    "NFDValorDest":          nfd_val_dest,
                    "Pedidos_Com_Cotacao":   ped_cot,
                })

            grp_dest = pd.DataFrame(rows_dest).sort_values("PedidosSimulados", ascending=False)

            st.markdown(f"**Para onde iriam os {pedidos_orig:,} pedidos**")
            st.caption("TM e SLA/NFD historicos da transportadora destino | TM Cotacao = valor cotado nos leiloes em que ela participou")

            for _, row in grp_dest.iterrows():
                tsp_dest = row["TransportadoraDestino"]
                ped_dest = int(row["PedidosSimulados"])
                pct_dist = row["Pct_dist"]

                st.markdown(
                    f"<div style='background:#e8f0fe;border-radius:8px;padding:6px 12px;"
                    f"margin-bottom:6px;font-size:12px;font-weight:600;color:#1f4e79'>"
                    f"{tsp_dest} &nbsp; {ped_dest:,} pedidos ({pct_dist:.1f}%)</div>",
                    unsafe_allow_html=True
                )

                d1, d2, d3, d4, d5, d6 = st.columns(6)

                delta_sla_v = row["SLADestino"] - sla_orig
                delta_nfd_v = row["NFDDestino"] - nfd_orig_pct
                delta_tm_v  = row["TM_Destino"] - tm_orig

                d1.markdown(kpi_box("SLA Hist.",
                    fmt_pct(row["SLADestino"]),
                    cor_valor=COR_OK if row["SLADestino"] >= 95 else COR_ALERTA,
                    delta=f"{delta_sla_v:+.1f}pp",
                    delta_cor=COR_OK if delta_sla_v >= 0 else COR_ALERTA
                ), unsafe_allow_html=True)

                d2.markdown(kpi_box("TM Hist.",
                    fmt_brl(row["TM_Destino"]),
                    delta=f"R$ {delta_tm_v:+.2f}".replace(".", ","),
                    delta_cor=COR_OK if delta_tm_v <= 0 else COR_ALERTA
                ), unsafe_allow_html=True)

                d3.markdown(kpi_box("NFD % Hist.",
                    fmt_pct(row["NFDDestino"]),
                    cor_valor=COR_ALERTA if row["NFDDestino"] >= 5 else COR_NEUTRO if row["NFDDestino"] >= 2 else COR_OK,
                    delta=f"{delta_nfd_v:+.1f}pp",
                    delta_cor=COR_OK if delta_nfd_v <= 0 else COR_ALERTA
                ), unsafe_allow_html=True)

                d4.markdown(kpi_box("TM Cotacao",
                    fmt_brl(row["TM_Cotacao_Destino"]) if pd.notna(row["TM_Cotacao_Destino"]) else "Sem cot."
                ), unsafe_allow_html=True)

                d5.markdown(kpi_box("Frete Proj.",
                    fmt_brl(row["FreteProj"])
                ), unsafe_allow_html=True)

                d6.markdown(kpi_box("NFD R$ Est.",
                    fmt_brl(row["NFDValorDest"]) if pd.notna(row.get("NFDValorDest")) else "sem dado"
                ), unsafe_allow_html=True)

                st.markdown("<div style='margin-bottom:8px'></div>", unsafe_allow_html=True)

            resumo_codigos.append({
                "codigo":         codigo,
                "pedidos_orig":   pedidos_orig,
                "sla_orig":       sla_orig,
                "tm_orig":        tm_orig,
                "nfd_orig_pct":   nfd_orig_pct,
                "frete_orig":     frete_orig,
                "valor_notas":    valor_notas,
                "nfd_orig_valor": nfd_orig_valor,
                "destinos":       grp_dest.to_dict("records")
            })

    # ====================================================
    # SECAO 4 - RESUMO CONSOLIDADO
    # ====================================================

    if not resumo_codigos:
        return

    st.divider()
    st.markdown("### Resumo Consolidado")

    total_pedidos_geral = sum(r["pedidos_orig"]   for r in resumo_codigos)
    total_frete_orig    = sum(r["frete_orig"]      for r in resumo_codigos if r["frete_orig"])
    total_notas         = sum(r["valor_notas"]     for r in resumo_codigos if r["valor_notas"])
    total_nfd_orig_val  = sum(r["nfd_orig_valor"]  for r in resumo_codigos if r["nfd_orig_valor"])

    from collections import defaultdict
    dest_agg = defaultdict(lambda: {
        "pedidos": 0, "frete_proj": 0, "frete_cot": 0,
        "sla_pond": 0, "nfd_pond": 0, "notas": 0,
        "nfd_valor": 0, "ped_com_cot": 0
    })

    for r in resumo_codigos:
        for d in r["destinos"]:
            tsp = d["TransportadoraDestino"]
            dest_agg[tsp]["pedidos"]    += d["PedidosSimulados"]
            dest_agg[tsp]["frete_proj"] += d.get("FreteProj") or 0
            dest_agg[tsp]["frete_cot"]  += d.get("FreteCot")  or 0
            dest_agg[tsp]["sla_pond"]   += (d.get("SLADestino") or 0) * d["PedidosSimulados"]
            dest_agg[tsp]["nfd_pond"]   += (d.get("NFDDestino") or 0) * d["PedidosSimulados"]
            dest_agg[tsp]["notas"]      += d.get("ValorNotasDest") or 0
            dest_agg[tsp]["nfd_valor"]  += d.get("NFDValorDest")   or 0
            dest_agg[tsp]["ped_com_cot"]+= d.get("Pedidos_Com_Cotacao") or 0

    rows_resumo = []
    for tsp, v in dest_agg.items():
        ped = v["pedidos"]
        if ped == 0:
            continue
        rows_resumo.append({
            "Transportadora": tsp,
            "Pedidos":        ped,
            "Pct":            ped / total_pedidos_geral * 100,
            "SLA":            v["sla_pond"] / ped,
            "NFD_pct":        v["nfd_pond"] / ped,
            "FreteProj":      v["frete_proj"],
            "FreteCot":       v["frete_cot"] if v["ped_com_cot"] > 0 else None,
            "NFDValor":       v["nfd_valor"],
        })

    df_resumo = pd.DataFrame(rows_resumo).sort_values("Pedidos", ascending=False)

    total_frete_proj = df_resumo["FreteProj"].sum()
    total_nfd_dest   = df_resumo["NFDValor"].sum()
    delta_frete      = total_frete_proj - total_frete_orig
    delta_nfd        = total_nfd_dest   - total_nfd_orig_val
    saldo            = delta_nfd - delta_frete

    # KPIs gerais
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi_box("Total de Pedidos",    fmt_int(total_pedidos_geral)), unsafe_allow_html=True)
    k2.markdown(kpi_box("Valor Total das Notas", fmt_brl(total_notas) if total_notas else "sem dado"), unsafe_allow_html=True)
    k3.markdown(kpi_box("NFD Atual (R$)",      fmt_brl(total_nfd_orig_val) if total_nfd_orig_val else "sem dado", cor_valor=COR_ALERTA), unsafe_allow_html=True)
    k4.markdown(kpi_box("Frete Atual",         fmt_brl(total_frete_orig)), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**Distribuicao por transportadora destino**")

    for _, row in df_resumo.iterrows():
        tsp  = row["Transportadora"]
        ped  = int(row["Pedidos"])
        pct  = row["Pct"]
        sla  = row["SLA"]
        nfd  = row["NFD_pct"]
        fp   = row["FreteProj"]
        fc   = row["FreteCot"]
        nfdv = row["NFDValor"]

        delta_fp   = fp   - (total_frete_orig    * pct / 100)
        delta_nfdv = nfdv - (total_nfd_orig_val  * pct / 100) if total_nfd_orig_val else None

        st.markdown(
            f"<div style='font-size:13px;font-weight:600;color:#1f4e79;margin:8px 0 4px'>"
            f"{tsp} - {ped:,} pedidos ({pct:.1f}%)</div>",
            unsafe_allow_html=True
        )

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.markdown(kpi_box("SLA",         fmt_pct(sla)), unsafe_allow_html=True)
        c2.markdown(kpi_box("NFD %",       fmt_pct(nfd)), unsafe_allow_html=True)
        c3.markdown(kpi_box("Frete Proj.", fmt_brl(fp)),  unsafe_allow_html=True)
        c4.markdown(kpi_box("Frete Cot.",  fmt_brl(fc) if fc else "sem cot."), unsafe_allow_html=True)
        c5.markdown(kpi_box("Delta Frete",
            fmt_brl(delta_fp),
            cor_valor=COR_OK if delta_fp <= 0 else COR_ALERTA
        ), unsafe_allow_html=True)
        c6.markdown(kpi_box("NFD R$ Est.",
            fmt_brl(nfdv),
            delta=fmt_brl(delta_nfdv) if delta_nfdv is not None else None,
            delta_cor=COR_OK if delta_nfdv is not None and delta_nfdv <= 0 else COR_ALERTA
        ), unsafe_allow_html=True)

        st.markdown("<div style='margin-bottom:4px'></div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.divider()

    # Saldo final
    st.markdown("### Saldo Final da Migracao")

    cor_frete = COR_OK    if delta_frete <= 0 else COR_ALERTA
    cor_nfd   = COR_OK    if delta_nfd   <= 0 else COR_ALERTA
    cor_saldo = COR_OK    if saldo       >= 0 else COR_ALERTA
    emoji     = "GANHO LIQUIDO" if saldo >= 0 else "PERDA LIQUIDA"

    s1, s2, s3 = st.columns(3)

    s1.markdown(
        f"""<div style="background:#fff;border-radius:12px;padding:18px;
            border-left:5px solid {cor_frete};text-align:center">
        <div style="font-size:12px;color:#888">Impacto em Frete</div>
        <div style="font-size:22px;font-weight:700;color:{cor_frete}">{fmt_brl(delta_frete)}</div>
        <div style="font-size:11px;color:#888;margin-top:4px">
            {'Economia' if delta_frete <= 0 else 'Custo adicional'} vs situacao atual
        </div></div>""",
        unsafe_allow_html=True
    )

    s2.markdown(
        f"""<div style="background:#fff;border-radius:12px;padding:18px;
            border-left:5px solid {cor_nfd};text-align:center">
        <div style="font-size:12px;color:#888">Impacto em NFD (R$)</div>
        <div style="font-size:22px;font-weight:700;color:{cor_nfd}">{fmt_brl(delta_nfd)}</div>
        <div style="font-size:11px;color:#888;margin-top:4px">
            {'Reducao' if delta_nfd <= 0 else 'Aumento'} vs NFD atual ({fmt_brl(total_nfd_orig_val)})
        </div></div>""",
        unsafe_allow_html=True
    )

    s3.markdown(
        f"""<div style="background:{cor_saldo}18;border-radius:12px;padding:18px;
            border:2px solid {cor_saldo};text-align:center">
        <div style="font-size:13px;color:{cor_saldo};font-weight:700">{emoji}</div>
        <div style="font-size:28px;font-weight:800;color:{cor_saldo}">{fmt_brl(abs(saldo))}</div>
        <div style="font-size:11px;color:{cor_saldo};margin-top:4px">Economia NFD - Custo extra frete</div>
        </div>""",
        unsafe_allow_html=True
    )

    st.markdown("<br>", unsafe_allow_html=True)
    st.info(
        f"**Como foi calculado:**\n\n"
        f"- **{total_pedidos_geral:,} pedidos** de {transp_sel} nos ultimos {dias} dias\n"
        f"- Valor total das notas: **{fmt_brl(total_notas)}**\n"
        f"- NFD atual em R$: **{fmt_brl(total_nfd_orig_val)}** (% NFD x valor das notas)\n"
        f"- Frete projetado pos-migracao: **{fmt_brl(total_frete_proj)}** vs atual **{fmt_brl(total_frete_orig)}** "
        f"-> {'economia' if delta_frete <= 0 else 'custo'} de **{fmt_brl(abs(delta_frete))}**\n"
        f"- NFD estimado pos-migracao: **{fmt_brl(total_nfd_dest)}** "
        f"-> {'reducao' if delta_nfd <= 0 else 'aumento'} de **{fmt_brl(abs(delta_nfd))}**\n"
        f"- **Saldo liquido: {fmt_brl(saldo)}**"
    )