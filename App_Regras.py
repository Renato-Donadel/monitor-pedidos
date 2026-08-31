import streamlit as st
import pandas as pd
import os
import plotly.graph_objects as go

from utils import PASTA_DATA

ARQ_REGRAS  = os.path.join(PASTA_DATA, "Regras_Resultado.xlsx")
ARQ_IMPACTO     = os.path.join(PASTA_DATA, "Impacto_Financeiro.csv")
ARQ_DETALHE     = os.path.join(PASTA_DATA, "Impacto_Detalhe.csv")
ARQ_CATEGORIAS  = os.path.join(PASTA_DATA, "Regras_Categorias.csv")
ARQ_ENTRADAS    = os.path.join(PASTA_DATA, "Novas_Transportadoras.csv")
ARQ_MCPOR       = os.path.join(PASTA_DATA, "MCPOR_Economia.csv")
MCPOR_CAP_PCT   = 999  # deve ficar igual ao MCPOR_CAP_PCT do Regras.py — só pra exibir na legenda
ARQ_ECONOMIA_EBB = os.path.join(PASTA_DATA, "Economia_EBB.csv")


@st.cache_data(ttl=3600)
def carregar_cat_map(path):
    """Mapa Regra -> Categoria a partir do CSV de categorias."""
    try:
        if os.path.exists(path):
            df_cat = pd.read_csv(path, sep=";")
            df_cat["Regra"] = pd.to_numeric(df_cat["Regra"], errors="coerce")
            return dict(zip(df_cat["Regra"].astype("Int64"),
                            df_cat["Categoria"].fillna("Sem Categoria")))
    except Exception:
        pass
    return {}


@st.cache_data(ttl=3600)
def carregar_detalhe(path):
    """
    Detalhe por EVENTO de cotação (v3). Colunas: Data;DataHora;PedidoID;
    ShipmentID;NF;CotacaoID;N_Tsp;Valor_Escolhida;Menor_Valor;Impacto;Regras
    """
    if not os.path.exists(path):
        return pd.DataFrame()
    df = pd.read_csv(path, sep=";", dtype={"Regras": str})
    if "N_Tsp" not in df.columns:      # formato antigo → força fallback
        return pd.DataFrame()
    df["Data"]     = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
    df["DataHora"] = pd.to_datetime(df["DataHora"], errors="coerce")
    for c in ["Impacto", "Valor_Escolhida", "Menor_Valor", "N_Tsp"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Impacto"] = df["Impacto"].fillna(0.0)
    df["Regras"]  = df["Regras"].fillna("")
    return df.dropna(subset=["Data", "DataHora"])


@st.cache_data(ttl=3600)
def carregar_entradas(path):
    """What-if das entradas de transportadoras (Novas_Transportadoras.csv)."""
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        df = pd.read_csv(path, sep=";")
    except Exception:
        return pd.DataFrame()
    if "Entrada" not in df.columns:
        return pd.DataFrame()
    df["Data"]     = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
    df["DataHora"] = pd.to_datetime(df["DataHora"], errors="coerce")
    for c in ["Valor_Escolhida", "Valor_Segunda", "Economia"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Campanha"] = df["Campanha"].astype(str)
    return df.dropna(subset=["Data", "DataHora", "Economia"])


@st.cache_data(ttl=3600)
def carregar_mcpor(path):
    """
    What-if MCPOR: regra ANTIGA (mais rápida, até 999% mais cara que a
    mais barata do leilão) x regra ATUAL (sempre a mais barata).
    Colunas: Data;DataHora;PedidoID;ShipmentID;NF;CotacaoID;Tsp_Novo;
    Valor_Novo;Prazo_Novo;Tsp_Antigo;Valor_Antigo;Prazo_Antigo;Economia;
    Campanha;ARM
    """
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        df = pd.read_csv(path, sep=";")
    except Exception:
        return pd.DataFrame()
    if "Economia" not in df.columns:
        return pd.DataFrame()
    df["Data"]     = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
    df["DataHora"] = pd.to_datetime(df["DataHora"], errors="coerce")
    for c in ["Valor_Novo", "Valor_Antigo", "Prazo_Novo", "Prazo_Antigo", "Economia"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df.dropna(subset=["Data", "DataHora", "Economia"])


@st.cache_data(ttl=3600)
def carregar_economia_ebb(path):
    """
    Economia estimada por NÃO usar mais a EBB (Todo Brasil/Dominalog no
    Sul/Sudeste e campanha HEIST). Diferente das demais seções desta
    página: não é uma 2ª cotação real do MESMO leilão — é uma ESTIMATIVA
    baseada na curva histórica de preço da EBB por (UF de destino x faixa
    de peso). Gerado por
    Projeto_CTE_Completo/scripts_temp/economia_ebb_todobrasil_dominalog_heist.py
    Colunas: Data;DataHora;Grupo;Transportadora;ChaveNFe;NumeroNF;
    ShipmentOrderID;UF;FaixaPeso;PesoRealKg;ValorCobrado;ValorEBBEstimado;Economia
    """
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        df = pd.read_csv(path, sep=";")
    except Exception:
        return pd.DataFrame()
    if "Economia" not in df.columns:
        return pd.DataFrame()
    df["Data"]     = pd.to_datetime(df["Data"], errors="coerce")
    df["DataHora"] = pd.to_datetime(df["DataHora"], errors="coerce")
    for c in ["PesoRealKg", "ValorCobrado", "ValorEBBEstimado", "Economia"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df.dropna(subset=["Data", "Economia"])


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
    df_base = df.dropna(subset=["Data"])
    # Aplica categorias atuais do CSV de mapeamento (substitui o que está salvo no histórico)
    try:
        if os.path.exists(ARQ_CATEGORIAS):
            df_cat_map = pd.read_csv(ARQ_CATEGORIAS, sep=";")
            df_cat_map["Regra"] = pd.to_numeric(df_cat_map["Regra"], errors="coerce")
            cat_dict = dict(zip(df_cat_map["Regra"], df_cat_map["Categoria"].fillna("Não classificado")))
            df_base["Categoria"] = df_base["Regra"].map(cat_dict).fillna("Não classificado")
    except Exception:
        pass
    return df_base


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
        "Diferença entre o frete contratado e a cotação mais barata do mesmo embarque "
        "(igual ao 'valor perdido' do relatório de pedidos). Cada embarque conta uma única vez; "
        "as regras servem apenas para alocar/explicar o impacto."
    )

    df_imp = carregar_impacto(ARQ_IMPACTO)
    df_det = carregar_detalhe(ARQ_DETALHE)
    cat_map = carregar_cat_map(ARQ_CATEGORIAS)

    if df_imp.empty:
        st.info("Nenhum dado de impacto disponível. Rode o `Regras.py` para gerar.")
    else:
        if df_det.empty:
            st.warning("`Impacto_Detalhe.csv` não encontrado — rode o `Regras.py` atualizado. "
                       "Exibindo valores agregados antigos (podem conter dupla contagem entre regras).")
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

        df_det_p = pd.DataFrame()
        if not df_det.empty:
            mask_d   = (df_det["Data"].dt.date >= d_ini) & (df_det["Data"].dt.date <= d_fim)
            df_det_p = df_det[mask_d].copy()

        if df_periodo.empty:
            st.warning("Nenhum dado no período selecionado.")
        else:
            n_dias = (d_fim - d_ini).days + 1

            def brl(v):
                return f"R$ {v:,.2f}".replace(",","X").replace(".",",").replace("X",".")

            # Categorias por embarque (um embarque pode tocar 2+ categorias)
            def _cats(regras_str):
                cats = set()
                for r in str(regras_str).split("|"):
                    r = r.strip()
                    if r:
                        try:
                            cats.add(cat_map.get(int(float(r)), "Sem Categoria"))
                        except (ValueError, TypeError):
                            pass
                cats.discard(None)
                return frozenset(cats) if cats else frozenset({"Sem Categoria"})

            if not df_det_p.empty:
                # ═════ REGRAS DO PERÍODO (v3) ═════════════════════
                # Só eventos DENTRO do período contam; cotações da
                # mesma NF antes/depois são irrelevantes.
                ev = df_det_p.sort_values("DataHora")

                # último evento (qualquer) de cada NF no período
                ult = ev.drop_duplicates(subset=["ShipmentID"], keep="last")[
                    ["ShipmentID", "N_Tsp"]].rename(columns={"N_Tsp": "N_Ult"})

                # leilões válidos: 2+ transportadoras e a escolhida cotou
                mult = ev[(ev["N_Tsp"] >= 2) & ev["Valor_Escolhida"].notna()]
                # o leilão que VALE: último do período por NF
                principal = mult.drop_duplicates(subset=["ShipmentID"], keep="last")
                # leilões anteriores da mesma NF no período (excedentes)
                extras = mult.loc[~mult.index.isin(principal.index)]
                extras_pos = extras[extras["Impacto"] > 0]

                principal = principal.merge(ult, on="ShipmentID", how="left")
                tampada = principal["N_Ult"] == 1   # re-cotação só com a escolhida por cima

                df_main = principal[(~tampada) & (principal["Impacto"] > 0)].copy()
                df_b1   = principal[(tampada)  & (principal["Impacto"] > 0)].copy()

                # ═════ KPIs (card principal) ══════════════════════
                df_main["Cats"] = df_main["Regras"].apply(_cats)
                total_impacto = df_main["Impacto"].sum()
                total_orders  = (df_main["PedidoID"].nunique()
                                 if df_main["PedidoID"].notna().any()
                                 else df_main["ShipmentID"].nunique())
                total_categorizado = df_main[
                    df_main["Cats"] != frozenset({"Sem Categoria"})
                ]["Impacto"].sum()
            else:
                df_main = pd.DataFrame()
                df_b1 = pd.DataFrame(); extras_pos = pd.DataFrame()
                # Fallback (CSV antigo, com dupla contagem entre regras)
                total_impacto = df_periodo["Impacto_Total"].sum()
                total_orders  = df_periodo["NOrders"].sum()
                total_categorizado = df_periodo[
                    ~df_periodo["Categoria"].isin(["Sem Categoria","Não classificado","Nao classificado"])
                ]["Impacto_Total"].sum()

            media_dia = total_impacto / n_dias if n_dias > 0 else 0
            pct_cat   = round(total_categorizado / total_impacto * 100, 1) if total_impacto > 0 else 0

            col_a, col_b, col_c, col_d, col_e = st.columns(5)
            col_a.metric("💸 Impacto Total",      brl(total_impacto),
                         help="Para cada NF: preço da escolhida no ÚLTIMO leilão do período "
                              "(cotação com 2+ transportadoras) menos a mais barata do mesmo leilão. "
                              "Cada NF conta uma única vez. Cotações fora do período são ignoradas.")
            col_b.metric("📦 Pedidos Impactados", f"{total_orders:,}")
            col_c.metric("📅 Dias Analisados",    n_dias)
            col_d.metric("📊 Média por Dia",       brl(media_dia))
            col_e.metric("🏷️ % Categorizado",      f"{pct_cat}%",
                         help="Percentual do impacto total que já tem categoria definida na planilha de regras")

            # ═════ Cards informativos (fora do Impacto Total) ═════
            if not df_det_p.empty:
                n_b1, v_b1 = len(df_b1), df_b1["Impacto"].sum()
                n_ex = extras_pos["ShipmentID"].nunique() if not extras_pos.empty else 0
                v_ex = extras_pos["Impacto"].sum() if not extras_pos.empty else 0.0

                col_i1, col_i2 = st.columns(2)
                col_i1.metric("🔁 Re-cotação com transportadora única",
                              f"{n_b1:,} NFs  ·  {brl(v_b1)}",
                              help="NFs cuja cotação mais recente do período tem SÓ a transportadora "
                                   "escolhida (re-cotação tampando o leilão). O valor é o impacto do "
                                   "último leilão com concorrência do período. NÃO entra no Impacto Total.")
                col_i2.metric("📑 2+ cotações no mesmo período",
                              f"{n_ex:,} NFs  ·  {brl(v_ex)}",
                              help="NFs com mais de um leilão dentro do período: só a diferença do "
                                   "último entra no Impacto Total; este card soma a dos leilões "
                                   "anteriores, apenas informativo.")
                st.caption("Os dois cards acima são informativos e **não** somam no Impacto Total — "
                           "evitam contar a mesma NF duas vezes.")

            # ── Visão por Categoria (com sobreposição explícita) ──
            CORES_CAT = {
                "Ganho Imediato":                "#e63946",
                "Ganho Imediato - Já realizado": "#f4a261",
                "Oportunidade":                  "#2a9d8f",
                "Regra operacional":             "#457b9d",
                "Express":                       "#e9c46a",
                "Default":                       "#6d6875",
                "Sem Categoria":                 "#adb5bd",
                "Não classificado":              "#adb5bd",
            }
            PRIORIDADE_CAT = ["Ganho Imediato", "Ganho Imediato - Já realizado",
                              "Oportunidade", "Regra operacional", "Express",
                              "Default", "Sem Categoria"]

            if not df_main.empty:
                # Por categoria: Total (pedidos que tocam a categoria),
                # Exclusivo (pedidos SÓ dessa categoria) e Sobreposto
                todas_cats = sorted({c for s in df_main["Cats"] for c in s})
                linhas_cat = []
                for c in todas_cats:
                    toca      = df_main[df_main["Cats"].apply(lambda s: c in s)]
                    exclusivo = toca[toca["Cats"].apply(lambda s: len(s) == 1)]["Impacto"].sum()
                    total_c   = toca["Impacto"].sum()
                    linhas_cat.append({
                        "Categoria":  c,
                        "Total":      round(total_c, 2),
                        "Exclusivo":  round(exclusivo, 2),
                        "Sobreposto": round(total_c - exclusivo, 2),
                        "Pedidos":    toca["ShipmentID"].nunique(),
                    })
                df_cat = (pd.DataFrame(linhas_cat)
                          .sort_values("Total", ascending=False))
                df_cat_plot = df_cat[df_cat["Categoria"] != "Sem Categoria"]

                # Donut: cada pedido alocado a UMA categoria (por prioridade),
                # para os percentuais somarem 100% do impacto total
                def _cat_primaria(cats):
                    for c in PRIORIDADE_CAT:
                        if c in cats:
                            return c
                    return sorted(cats)[0]
                df_main["Cat_Primaria"] = df_main["Cats"].apply(_cat_primaria)
                df_donut = (df_main.groupby("Cat_Primaria", as_index=False)["Impacto"].sum()
                            .rename(columns={"Cat_Primaria": "Categoria", "Impacto": "Impacto_Total"})
                            .sort_values("Impacto_Total", ascending=False))

                col_donut, col_cat_bar = st.columns([1, 1])
                with col_donut:
                    fig_donut = go.Figure(go.Pie(
                        labels=df_donut["Categoria"],
                        values=df_donut["Impacto_Total"],
                        hole=0.55,
                        marker_colors=[CORES_CAT.get(c, "#adb5bd") for c in df_donut["Categoria"]],
                        textinfo="label+percent",
                        hovertemplate="%{label}<br>R$ %{value:,.2f}<extra></extra>",
                    ))
                    fig_donut.update_layout(
                        title="Distribuição por Categoria",
                        height=340,
                        margin=dict(t=50, b=10, l=10, r=10),
                        showlegend=False,
                    )
                    st.plotly_chart(fig_donut, use_container_width=True)
                    st.caption("No donut cada pedido entra em uma única categoria "
                               "(a de maior prioridade), por isso soma 100% do impacto total.")

                with col_cat_bar:
                    df_cp = df_cat_plot.copy()
                    df_cp["Pct_Sobre"] = (df_cp["Sobreposto"] / df_cp["Total"] * 100).fillna(0)
                    df_cp["Rotulo"] = df_cp.apply(
                        lambda r: f"{brl(r['Total'])}  ·  {r['Pct_Sobre']:.0f}% sobreposto"
                                  if r["Sobreposto"] > 0 else brl(r["Total"]),
                        axis=1,
                    )
                    fig_cat = go.Figure(go.Bar(
                        y=df_cp["Categoria"],
                        x=df_cp["Total"],
                        orientation="h",
                        marker=dict(color=[CORES_CAT.get(c, "#adb5bd") for c in df_cp["Categoria"]]),
                        text=df_cp["Rotulo"],
                        textposition="inside",
                        insidetextanchor="start",
                        textfont=dict(color="white", size=12),
                        textangle=0,
                        constraintext="none",
                        cliponaxis=False,
                        customdata=list(zip(df_cp["Exclusivo"].apply(brl),
                                            df_cp["Sobreposto"].apply(brl),
                                            df_cp["Pct_Sobre"].round(1))),
                        hovertemplate="<b>%{y}</b><br>"
                                      "Total: %{x:,.2f}<br>"
                                      "Exclusivo da categoria: %{customdata[0]}<br>"
                                      "Sobreposto (2+ categorias): %{customdata[1]} (%{customdata[2]}%)"
                                      "<extra></extra>",
                    ))
                    fig_cat.update_layout(
                        title="Impacto por Categoria (R$)",
                        height=max(300, len(df_cp) * 60 + 100),
                        margin=dict(t=50, b=10, l=10, r=10),
                        xaxis=dict(tickformat=",.2f", rangemode="tozero"),
                        yaxis=dict(autorange="reversed"),
                        showlegend=False,
                    )
                    st.plotly_chart(fig_cat, use_container_width=True)
                    st.caption("O % sobreposto é a fatia do valor da categoria que também pertence a "
                               "outra categoria (pedido atingido por regras de categorias diferentes). "
                               "Por isso a soma das barras pode exceder o Impacto Total.")
            else:
                # Fallback sem detalhe: visão antiga por categoria (agregado)
                df_cat = (
                    df_periodo[df_periodo["Impacto_Total"] > 0]
                    .groupby("Categoria", as_index=False)["Impacto_Total"].sum()
                    .sort_values("Impacto_Total", ascending=False)
                )
                if not df_cat.empty:
                    cores = [CORES_CAT.get(c, "#adb5bd") for c in df_cat["Categoria"]]
                    col_donut, col_cat_bar = st.columns([1, 1])
                    with col_donut:
                        fig_donut = go.Figure(go.Pie(
                            labels=df_cat["Categoria"], values=df_cat["Impacto_Total"],
                            hole=0.55, marker_colors=cores, textinfo="label+percent",
                            hovertemplate="%{label}<br>R$ %{value:,.2f}<extra></extra>",
                        ))
                        fig_donut.update_layout(title="Distribuição por Categoria", height=320,
                                                margin=dict(t=50, b=10, l=10, r=10), showlegend=False)
                        st.plotly_chart(fig_donut, use_container_width=True)
                    with col_cat_bar:
                        fig_cat = go.Figure(go.Bar(
                            x=df_cat["Impacto_Total"], y=df_cat["Categoria"], orientation="h",
                            marker_color=cores, text=df_cat["Impacto_Total"].apply(brl),
                            textposition="inside", insidetextanchor="start",
                            textfont=dict(color="white", size=12),
                        textangle=0,
                            constraintext="none", cliponaxis=False,
                            hovertemplate="%{y}<br>%{text}<extra></extra>",
                        ))
                        fig_cat.update_layout(title="Impacto por Categoria (R$)",
                                              height=max(280, len(df_cat) * 60 + 80),
                                              margin=dict(t=50, b=10, l=10, r=10),
                                              xaxis=dict(tickformat=",.2f", rangemode="tozero"),
                                              yaxis=dict(autorange="reversed"))
                        st.plotly_chart(fig_cat, use_container_width=True)

            st.markdown("---")

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
                    textposition="inside",
                    insidetextanchor="start",
                    textfont=dict(color="white", size=12),
                        textangle=0,
                    constraintext="none",
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
                st.caption("Um pedido atingido por mais de uma regra aparece em todas elas — "
                           "a soma das barras pode exceder o Impacto Total, que é deduplicado por pedido.")

            # ── Linha temporal ─────────────────────────────────
            if len(datas_disp) > 1:
                if not df_main.empty:
                    df_por_dia = (
                        df_main
                        .groupby("Data", as_index=False)["Impacto"].sum()
                        .rename(columns={"Impacto": "Impacto_Total"})
                        .sort_values("Data")
                    )
                else:
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


    # ══════════════════════════════════════════════════════════
    # 🚚 ENTRADAS DE TRANSPORTADORAS — what-if de economia
    # ══════════════════════════════════════════════════════════
    st.markdown("---")
    st.markdown("### 🚚 Entradas de Transportadoras — quanto elas economizam")
    st.caption(
        "Para cada frete vencido por uma transportadora nova, comparamos com a "
        "**2ª opção mais barata do mesmo leilão** — quem levaria o frete se a "
        "entrada não existisse. Economia líquida = soma de (2ª opção − escolhida); "
        "valores negativos entram na conta (fretes em que a entrada era mais cara)."
    )

    df_ent = carregar_entradas(ARQ_ENTRADAS)
    if df_ent.empty:
        st.info("Sem dados ainda — rode o `Regras.py` atualizado para gerar o "
                "`Novas_Transportadoras.csv` na pasta `data/`.")
    else:
        def _brl(v):
            return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        e1, e2, e3 = st.columns([1, 1, 2])
        ent_ini = e1.date_input("De",  value=df_ent["Data"].min().date(),
                                key="ent_d_ini")
        ent_fim = e2.date_input("Até", value=df_ent["Data"].max().date(),
                                key="ent_d_fim")
        sem_heist = e3.toggle("Excluir campanha Heist", value=True, key="ent_heist",
                              help="Heist tem direcionamento próprio e não conta "
                                   "como economia da entrada.")

        mask_e = (df_ent["Data"].dt.date >= ent_ini) & (df_ent["Data"].dt.date <= ent_fim)
        df_e = df_ent[mask_e].copy()
        if sem_heist:
            df_e = df_e[~df_e["Campanha"].str.upper().str.contains("HEIST", na=False)]

        # 1 NF conta uma vez: último evento do período por ShipmentID
        df_e = (df_e.sort_values("DataHora")
                    .drop_duplicates(subset=["ShipmentID"], keep="last"))

        if df_e.empty:
            st.warning("Nenhum frete das entradas no período selecionado.")
        else:
            # Economia real (ganhou no preço) ≠ direcionamento (regra
            # escolheu a entrada mesmo mais cara). São perguntas diferentes.
            df_e["Tipo"] = df_e["Economia"].apply(
                lambda v: "Ganhou no preço" if v >= 0 else "Direcionada (mais cara)")

            entradas = sorted(df_e["Entrada"].unique())
            ESPERADAS = ["JADLOG", "GOL (SPO LOG)", "LSP",
                         "IMILE (VGP)", "ASAP (Casas Bahia)"]
            faltando = [e for e in ESPERADAS if e not in entradas]
            if faltando:
                st.caption("⚠️ Sem fretes no período para: **" + ", ".join(faltando) +
                           "** — confira o nome/armazém no `NOVAS_TSP` do Regras.py.")

            cols = st.columns(len(entradas)) if len(entradas) <= 6 else st.columns(6)
            for i, ent in enumerate(entradas):
                d    = df_e[df_e["Entrada"] == ent]
                pos  = d[d["Economia"] >= 0]
                neg  = d[d["Economia"] < 0]
                cols[i % len(cols)].metric(
                    f"🚚 {ent}", _brl(pos["Economia"].sum()),
                    delta=f"{pos['ShipmentID'].nunique():,} fretes ganhos no preço",
                    delta_color="off",
                    help=f"Economia real: soma de (2ª opção − {ent}) apenas nos "
                         f"fretes em que {ent} era a mais barata. Fora do número: "
                         f"{neg['ShipmentID'].nunique():,} fretes direcionados "
                         f"(custo de {_brl(-neg['Economia'].sum())}).")

            resumo_ent = (df_e.groupby(["Entrada", "Tipo"])
                              .agg(Fretes=("ShipmentID", "nunique"),
                                   Valor=("Economia", "sum"))
                              .reset_index()
                              .pivot(index="Entrada", columns="Tipo",
                                     values=["Fretes", "Valor"]))
            eco_tot = df_e.loc[df_e["Economia"] >= 0, "Economia"].sum()
            dir_tot = -df_e.loc[df_e["Economia"] < 0, "Economia"].sum()
            c_ok, c_dir = st.columns(2)
            c_ok.success(f"**💰 Economia real (ganhou no preço): {_brl(eco_tot)}**")
            c_dir.warning(f"**🎯 Custo de direcionamento (regra escolheu mesmo "
                          f"mais cara): {_brl(dir_tot)}**")
            with st.expander("📊 Composição economia × direcionamento por entrada"):
                st.dataframe(resumo_ent.round(2), use_container_width=True)

            # ── Evolução semanal e mensal (só economia real) ───
            df_pos = df_e[df_e["Economia"] >= 0].copy()
            df_pos["Semana"] = df_pos["Data"].dt.to_period("W").dt.start_time
            df_pos["Mes"]    = df_pos["Data"].dt.to_period("M").dt.to_timestamp()

            piv_s = (df_pos.groupby(["Semana", "Entrada"])["Economia"]
                           .sum().reset_index())
            fig_sem = go.Figure()
            for ent in entradas:
                d = piv_s[piv_s["Entrada"] == ent]
                fig_sem.add_trace(go.Bar(x=d["Semana"], y=d["Economia"], name=ent))
            fig_sem.update_layout(
                barmode="group", height=360,
                title="Economia real por SEMANA (fretes ganhos no preço)",
                yaxis=dict(title="Economia (R$)", tickformat=",.0f"),
                xaxis=dict(tickformat="%d/%m"),
                legend=dict(orientation="h", y=1.15),
            )
            st.plotly_chart(fig_sem, use_container_width=True)

            piv_m = (df_pos.groupby(["Mes", "Entrada"])["Economia"]
                           .sum().reset_index())
            fig_mes = go.Figure()
            for ent in entradas:
                d = piv_m[piv_m["Entrada"] == ent]
                fig_mes.add_trace(go.Bar(x=d["Mes"], y=d["Economia"], name=ent))
            fig_mes.update_layout(
                barmode="group", height=360,
                title="Economia real por MÊS (fretes ganhos no preço)",
                yaxis=dict(title="Economia (R$)", tickformat=",.0f"),
                xaxis=dict(dtick="M1", tickformat="%b/%Y"),
                legend=dict(orientation="h", y=1.15),
            )
            st.plotly_chart(fig_mes, use_container_width=True)

            # ── Quem herdaria os fretes ────────────────────────
            with st.expander("🔍 Detalhe por entrada — quem herdaria os fretes"):
                ent_sel = st.selectbox("Entrada", entradas, key="ent_sel")
                d = df_e[df_e["Entrada"] == ent_sel]

                c_h1, c_h2 = st.columns(2)
                herd = (d.groupby("Segunda_Tsp")
                          .agg(Fretes=("ShipmentID", "nunique"),
                               Economia=("Economia", "sum"),
                               Media=("Economia", "mean"))
                          .round(2).sort_values("Economia", ascending=False))
                c_h1.markdown(f"**Para quem iriam os fretes da {ent_sel}:**")
                c_h1.dataframe(herd, use_container_width=True)

                top_nf = (d.nlargest(15, "Economia")
                            [["NF", "Data", "Tsp_Escolhida", "Valor_Escolhida",
                              "Segunda_Tsp", "Valor_Segunda", "Economia", "Campanha"]])
                top_nf["Data"] = top_nf["Data"].dt.strftime("%d/%m/%Y")
                c_h2.markdown("**Top 15 maiores economias:**")
                c_h2.dataframe(top_nf, use_container_width=True, hide_index=True)

            # ── Exportar ───────────────────────────────────────
            st.download_button(
                "⬇️ Exportar composição (CSV)",
                df_e.to_csv(index=False, sep=";").encode("utf-8-sig"),
                file_name=f"Entradas_{ent_ini:%d%m%Y}_{ent_fim:%d%m%Y}.csv",
                mime="text/csv", key="ent_dl")

    # ══════════════════════════════════════════════════════════
    # 🥤 MCPOR — economia com a troca de regra (mais barato x mais rápido)
    # ══════════════════════════════════════════════════════════
    st.markdown("---")
    st.markdown("### 🥤 MCPOR — Economia com a troca de regra de leilão")
    st.caption(
        "Comparação por leilão (2+ transportadoras elegíveis cotando) da campanha MCPOR: "
        "regra **antiga** — escolhia a **mais rápida**, desde que o valor não passasse de "
        f"**{MCPOR_CAP_PCT}% mais caro** que a mais barata do leilão — contra a regra "
        "**atual** — escolhe sempre a **mais barata**. "
        "Economia = valor da regra antiga − valor da regra atual (positivo = economizado). "
        "EBB é excluída dos dois lados por já ser bloqueada em MCPOR pela Regra 394."
    )

    df_mcp = carregar_mcpor(ARQ_MCPOR)
    if df_mcp.empty:
        st.info("Sem dados ainda — rode o `Regras.py` atualizado para gerar o "
                "`MCPOR_Economia.csv` na pasta `data/`.")
    else:
        def _brl2(v):
            return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        m1, m2 = st.columns(2)
        mc_ini = m1.date_input("De",  value=df_mcp["Data"].min().date(), key="mcp_d_ini")
        mc_fim = m2.date_input("Até", value=df_mcp["Data"].max().date(), key="mcp_d_fim")

        mask_m = (df_mcp["Data"].dt.date >= mc_ini) & (df_mcp["Data"].dt.date <= mc_fim)
        df_m = df_mcp[mask_m].copy()

        # 1 leilão conta uma vez: último evento do período por ShipmentID+CotacaoID
        df_m = (df_m.sort_values("DataHora")
                    .drop_duplicates(subset=["ShipmentID", "CotacaoID"], keep="last"))

        if df_m.empty:
            st.warning("Nenhum leilão de MCPOR no período selecionado.")
        else:
            df_m["Trocou_Tsp"] = df_m["Tsp_Novo"] != df_m["Tsp_Antigo"]
            n_dias_m       = (mc_fim - mc_ini).days + 1
            economia_total = df_m["Economia"].sum()
            n_leiloes      = len(df_m)
            n_trocou       = int(df_m["Trocou_Tsp"].sum())

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("💰 Economia total", _brl2(economia_total),
                      help="Soma de (valor da regra antiga − valor da regra atual) em todos os leilões do período.")
            c2.metric("📦 Leilões MCPOR", f"{n_leiloes:,}")
            c3.metric("🔀 Trocaram de transportadora", f"{n_trocou:,}",
                      help="Leilões em que a regra atual (mais barata) escolheria uma "
                           "transportadora diferente da regra antiga (mais rápida).")
            c4.metric("📊 Média por dia", _brl2(economia_total / n_dias_m if n_dias_m > 0 else 0))

            # ── Evolução semanal e mensal ───────────────────────
            df_m["Semana"] = df_m["Data"].dt.to_period("W").dt.start_time
            df_m["Mes"]    = df_m["Data"].dt.to_period("M").dt.to_timestamp()

            piv_s = df_m.groupby("Semana", as_index=False)["Economia"].sum()
            fig_sem = go.Figure(go.Bar(x=piv_s["Semana"], y=piv_s["Economia"],
                                       marker_color="#2a9d8f"))
            fig_sem.update_layout(
                height=320, title="Economia MCPOR por SEMANA",
                yaxis=dict(title="Economia (R$)", tickformat=",.0f"),
                xaxis=dict(tickformat="%d/%m"),
                margin=dict(t=50, b=10, l=10, r=10),
            )
            st.plotly_chart(fig_sem, use_container_width=True)

            piv_m = df_m.groupby("Mes", as_index=False)["Economia"].sum()
            fig_mes = go.Figure(go.Bar(x=piv_m["Mes"], y=piv_m["Economia"],
                                       marker_color="#2a9d8f"))
            fig_mes.update_layout(
                height=320, title="Economia MCPOR por MÊS",
                yaxis=dict(title="Economia (R$)", tickformat=",.0f"),
                xaxis=dict(dtick="M1", tickformat="%b/%Y"),
                margin=dict(t=50, b=10, l=10, r=10),
            )
            st.plotly_chart(fig_mes, use_container_width=True)

            # ── Quem ganhava x quem ganha hoje ──────────────────
            with st.expander("🔍 Quem ganhava (regra antiga) x quem ganha hoje (regra atual)"):
                col_g1, col_g2 = st.columns(2)
                ganhava = (df_m.groupby("Tsp_Antigo")
                              .agg(Leiloes=("ShipmentID", "nunique"))
                              .sort_values("Leiloes", ascending=False))
                ganha = (df_m.groupby("Tsp_Novo")
                            .agg(Leiloes=("ShipmentID", "nunique"))
                            .sort_values("Leiloes", ascending=False))
                col_g1.markdown("**Regra antiga (mais rápida elegível):**")
                col_g1.dataframe(ganhava, use_container_width=True)
                col_g2.markdown("**Regra atual (mais barata):**")
                col_g2.dataframe(ganha, use_container_width=True)

                top_nf_m = (df_m.nlargest(15, "Economia")
                              [["NF", "Data", "Tsp_Antigo", "Valor_Antigo", "Prazo_Antigo",
                                "Tsp_Novo", "Valor_Novo", "Prazo_Novo", "Economia", "Campanha"]])
                top_nf_m["Data"] = top_nf_m["Data"].dt.strftime("%d/%m/%Y")
                st.markdown("**Top 15 maiores economias:**")
                st.dataframe(top_nf_m, use_container_width=True, hide_index=True)

            # ── Exportar ─────────────────────────────────────────
            st.download_button(
                "⬇️ Exportar MCPOR (CSV)",
                df_m.to_csv(index=False, sep=";").encode("utf-8-sig"),
                file_name=f"MCPOR_{mc_ini:%d%m%Y}_{mc_fim:%d%m%Y}.csv",
                mime="text/csv", key="mcp_dl")

    # ══════════════════════════════════════════════════════════
    # 🚛 EBB — Economia estimada por não usar mais (histórico)
    # ══════════════════════════════════════════════════════════
    st.markdown("---")
    st.markdown("### 🚛 EBB — Economia por não usar mais (estimado)")
    st.caption(
        "⚠️ **Diferente das seções acima**: aqui não comparamos com a 2ª cotação real do "
        "mesmo leilão — a EBB foi retirada da operação, então estimamos quanto ela cobraria "
        "hoje pela **curva histórica de preço dela** (média por UF de destino x faixa de peso, "
        "com base nos próprios fretes que ela já levou). Peso = peso real (físico x cubagem) via "
        "Intelipost. Economia = valor estimado da EBB − valor realmente cobrado "
        "(positivo = economizado por não mandar mais pela EBB, já que ela foi tirada por sair mais cara)."
    )

    df_ebb = carregar_economia_ebb(ARQ_ECONOMIA_EBB)
    if df_ebb.empty:
        st.info("Sem dados ainda — rode "
                "`Projeto_CTE_Completo/scripts_temp/economia_ebb_todobrasil_dominalog_heist.py` "
                "para gerar o `Economia_EBB.csv` na pasta `data/`.")
    else:
        def _brl3(v):
            return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        b1, b2 = st.columns(2)
        ebb_ini = b1.date_input("De",  value=df_ebb["Data"].min().date(), key="ebb_d_ini")
        ebb_fim = b2.date_input("Até", value=df_ebb["Data"].max().date(), key="ebb_d_fim")

        mask_b = (df_ebb["Data"].dt.date >= ebb_ini) & (df_ebb["Data"].dt.date <= ebb_fim)
        df_b = df_ebb[mask_b].copy()

        if df_b.empty:
            st.warning("Nenhum pedido no período selecionado.")
        else:
            grupos = sorted(df_b["Grupo"].unique())
            cols_g = st.columns(len(grupos))
            for i, g in enumerate(grupos):
                d = df_b[df_b["Grupo"] == g]
                cols_g[i].metric(
                    f"💰 {g}", _brl3(d["Economia"].sum()),
                    delta=f"{len(d):,} pedido(s)", delta_color="off",
                    help="Soma de (valor estimado da EBB − valor cobrado) nos pedidos "
                         "com faixa de peso/UF que tiveram histórico da EBB pra comparar.",
                )

            st.markdown("**Economia estimada por UF de destino**")
            por_uf_ebb = (df_b.groupby(["UF", "Grupo"], as_index=False)["Economia"].sum())
            fig_uf = go.Figure()
            for g in grupos:
                d = por_uf_ebb[por_uf_ebb["Grupo"] == g].sort_values("UF")
                fig_uf.add_trace(go.Bar(x=d["UF"], y=d["Economia"], name=g))
            fig_uf.update_layout(
                barmode="group", height=340,
                yaxis=dict(title="Economia (R$)", tickformat=",.0f"),
                legend=dict(orientation="h", y=1.15),
                margin=dict(t=30, b=10, l=10, r=10),
            )
            st.plotly_chart(fig_uf, use_container_width=True)

            with st.expander("🔍 Detalhe — maiores economias e amostra completa"):
                g_sel = st.selectbox("Grupo", grupos, key="ebb_grupo_sel")
                d_sel = df_b[df_b["Grupo"] == g_sel]
                top_ebb = (d_sel.nlargest(15, "Economia")
                           [["NumeroNF", "Data", "Transportadora", "UF", "FaixaPeso",
                             "ValorCobrado", "ValorEBBEstimado", "Economia"]])
                top_ebb["Data"] = top_ebb["Data"].dt.strftime("%d/%m/%Y")
                st.markdown(f"**Top 15 maiores economias — {g_sel}:**")
                st.dataframe(top_ebb, use_container_width=True, hide_index=True)

            st.download_button(
                "⬇️ Exportar Economia EBB (CSV)",
                df_b.to_csv(index=False, sep=";").encode("utf-8-sig"),
                file_name=f"Economia_EBB_{ebb_ini:%d%m%Y}_{ebb_fim:%d%m%Y}.csv",
                mime="text/csv", key="ebb_dl")