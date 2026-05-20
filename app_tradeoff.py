import os
import streamlit as st
import pandas as pd
from io import BytesIO
import plotly.express as px
import matplotlib.pyplot as plt
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

def render_tradeoff():

# ==============================================
# TRADE-OFF LOGÍSTICO
# ==============================================

    st.markdown("## 🚚 Trade-Off Logístico")

    st.caption(
        "Simulação operacional entre transportadoras baseada em similaridade de CEP"
    )

    st.divider()

    # =====================================================
    # CARREGAR BASES
    # =====================================================

    if os.path.exists("data/Base_Pedidos_Codigo.xlsx"):

        df_trade = pd.read_excel(
            "data/Base_Pedidos_Codigo.xlsx"
        )

    else:

        df_trade = pd.DataFrame()

    if os.path.exists("data/Base_Similaridade_Tarifarios.xlsx"):

        df_sim = pd.read_excel(
            "data/Base_Similaridade_Tarifarios.xlsx"
        )

    else:

        df_sim = pd.DataFrame()

    # =====================================================
    # FILTROS
    # =====================================================

    col1, col2, col3 = st.columns(3)

    with col1:

        transportadoras = sorted(

            df_trade["Transportadora"]
            .dropna()
            .unique()
        )

        transportadora_origem = st.selectbox(
            "Transportadora Atual",
            transportadoras
        )

    with col2:

        transportadora_destino = st.selectbox(
            "Transportadora Simulada",
            [
                "MAGALU",
                "IMILE"
            ]
        )

    with col3:

        periodo = st.selectbox(
            "Período",
            [
                "30 dias",
                "60 dias",
                "90 dias"
            ]
        )
        
        if periodo == "30 dias":
            dias = 30

        elif periodo == "60 dias":
            dias = 60

        else:
            dias = 90

        df_trade["DataFinal"] = pd.to_datetime(
            df_trade["DataFinal"],
            errors="coerce"
        )

        data_corte = (
            pd.Timestamp.today()
            - pd.Timedelta(days=dias)
        )

        df_trade = df_trade[
            df_trade["DataFinal"] >= data_corte
        ]

    # =====================================================
    # CÓDIGOS TARIFÁRIOS
    # =====================================================

    codigos = sorted(

        df_trade[

            df_trade["Transportadora"]
            == transportadora_origem

        ]["CodigoTarifario"]

        .dropna()
        .unique()
    )

    codigo_origem = st.selectbox(
        "Código Tarifário",
        codigos
    )

    st.divider()

    # =====================================================
    # FILTRAR SIMILARIDADE
    # =====================================================

    df_sim_filtrado = df_sim[

        (df_sim["TransportadoraOrigem"] == transportadora_origem)

        &

        (df_sim["CodigoOrigem"] == codigo_origem)

        &

        (
            df_sim["TransportadoraDestino"]
            == transportadora_destino
        )

    ].copy()

    # ordena
    df_sim_filtrado = df_sim_filtrado.sort_values(
        "Percentual",
        ascending=False
    )
    
    # =====================================================
    # BASE DESTINO
    # =====================================================

    codigos_destino = (

        df_sim_filtrado["CodigoDestino"]
        .dropna()
        .unique()
    )

    df_destino = df_trade[

        (df_trade["Transportadora"] == transportadora_destino)

        &

        (df_trade["CodigoTarifario"].isin(codigos_destino))

    ].copy()
    
    # =====================================================
    # TM DESTINO POR CÓDIGO
    # =====================================================

    tm_destino = (

        df_destino.groupby("CodigoTarifario")

        .agg(

            TM_Destino=("ValorFrete", "mean"),

            SLA_Destino=("DentroPrazo", "mean"),

            NFD_Destino=("TemNFD", "mean")

        )

        .reset_index()
    )

    # =====================================================
    # PEDIDOS ORIGEM
    # =====================================================

    df_origem = df_trade[

        (df_trade["Transportadora"] == transportadora_origem)

        &

        (df_trade["CodigoTarifario"] == codigo_origem)

    ].copy()
    
    # =====================================================
    # SLA
    # =====================================================

    sla_origem = round(

        (
            df_origem["DentroPrazo"]
            .astype(bool)
            .mean()
        ) * 100,

        2
    )
    
    
    
    nfd_origem = round(

        (
            df_origem["TemNFD"]
            .astype(bool)
            .mean()
        ) * 100,

        2
    )

    nfd_destino = round(

        (
            df_destino["TemNFD"]
            .astype(bool)
            .mean()
        ) * 100,

        2
    )

    total_pedidos = len(df_origem)

    # =====================================================
    # PEDIDOS SIMULADOS
    # =====================================================

    df_sim_filtrado["Pedidos"] = (

        total_pedidos

        *

        (df_sim_filtrado["Percentual"] / 100)

    ).round(0)
    
    # =====================================================
    # JOIN COM TM DESTINO
    # =====================================================

    df_sim_filtrado = df_sim_filtrado.merge(

        tm_destino,

        left_on="CodigoDestino",

        right_on="CodigoTarifario",

        how="left"
    )

    df_sim_filtrado["TM_Destino"] = (
        df_sim_filtrado["TM_Destino"]
        .fillna(0)
    )

    df_sim_filtrado["SLA_Destino"] = (
        df_sim_filtrado["SLA_Destino"]
        .fillna(0)
    )

    df_sim_filtrado["NFD_Destino"] = (
        df_sim_filtrado["NFD_Destino"]
        .fillna(0)
    )
    
    sla_destino = round(

        (
            (
                df_sim_filtrado["SLA_Destino"]

                *

                df_sim_filtrado["Pedidos"]

            ).sum()

            /

            df_sim_filtrado["Pedidos"].sum()

        ) * 100,

        2
    ) if df_sim_filtrado["Pedidos"].sum() > 0 else 0

    nfd_destino = round(

        (
            (
                df_sim_filtrado["NFD_Destino"]

                *

                df_sim_filtrado["Pedidos"]

            ).sum()

            /

            df_sim_filtrado["Pedidos"].sum()

        ) * 100,

        2
    ) if df_sim_filtrado["Pedidos"].sum() > 0 else 0

    # =====================================================
    # KPIs
    # =====================================================

    similaridade_media = round(

        df_sim_filtrado["Percentual"].mean(),

        2
    )

    pedidos_simulados = int(

        df_sim_filtrado["Pedidos"].sum()

    )
    
    frete_medio_origem = round(

        df_origem["ValorFrete"]
        .astype(float)
        .mean(),

        2
    )


    gasto_total = round(

        df_origem["ValorFrete"]
        .astype(float)
        .sum(),

        2
    )
    
    # =====================================================
    # PROJEÇÃO PONDERADA
    # =====================================================

    df_sim_filtrado["FreteProjetado"] = (

        df_sim_filtrado["TM_Destino"]

        *

        df_sim_filtrado["Pedidos"]

    )

    gasto_projetado = round(

        df_sim_filtrado["FreteProjetado"].sum(),

        2
    )

    economia_projetada = round(

        gasto_total

        -

        gasto_projetado,

        2
    )
    
    tm_destino_ponderado = round(

        gasto_projetado

        /

        pedidos_simulados,

        2
    ) if pedidos_simulados > 0 else 0

    k1, k2, k3, k4, k5, k6, k7, k8, k9, k10, k11 = st.columns(11)

    with k1:

        st.metric(
            "Pedidos Origem",
            f"{total_pedidos:,}"
        )

    with k2:

        st.metric(
            "Similaridade Média",
            f"{similaridade_media}%"
        )

    with k3:

        st.metric(
            "Pedidos Simulados",
            f"{pedidos_simulados:,}"
        )
        
    with k4:

        st.metric(
            "TM Origem",
            f"R$ {frete_medio_origem:,.2f}"
        )

    with k5:

        st.metric(
            "Gasto Total",
            f"R$ {gasto_total:,.2f}"
        )
        
    with k6:

        st.metric(
            "SLA Origem",
            f"{sla_origem}%"
        )

    with k7:

        st.metric(
            "SLA Destino",
            f"{sla_destino}%"
        )
        
    with k8:

        st.metric(
            "TM Destino",
            f"R$ {tm_destino_ponderado:,.2f}"
        )

    with k9:

        st.metric(
            "Economia Projetada",
            f"R$ {economia_projetada:,.2f}"
        )
        
    with k10:

        st.metric(
            "NFD Origem",
            f"{nfd_origem}%"
        )

    with k11:

        st.metric(
            "NFD Destino",
            f"{nfd_destino}%"
        )

    st.divider()

    # =====================================================
    # TABELA
    # =====================================================

    st.markdown("### 🔄 Redistribuição Operacional")

    tabela = df_sim_filtrado[
        [
            "CodigoDestino",
            "Percentual",
            "Pedidos",
            "TM_Destino",
            "SLA_Destino",
            "NFD_Destino",
            "FreteProjetado"
        ]
    ].copy()

    tabela["TM_Destino"] = tabela["TM_Destino"].round(2)

    tabela["SLA_Destino"] = (
        tabela["SLA_Destino"] * 100
    ).round(2)

    tabela["NFD_Destino"] = (
        tabela["NFD_Destino"] * 100
    ).round(2)

    tabela["FreteProjetado"] = (
        tabela["FreteProjetado"]
    ).round(2)

    tabela = tabela.rename(columns={

        "CodigoDestino": "Código Destino",

        "Percentual": "Similaridade %",

        "Pedidos": "Pedidos Simulados",

        "TM_Destino": "TM Destino",

        "SLA_Destino": "SLA Destino %",

        "NFD_Destino": "NFD Destino %",

        "FreteProjetado": "Valor Projetado"

    })

    st.dataframe(
        tabela,
        use_container_width=True,
        hide_index=True
    )

    st.divider()

    # =====================================================
    # RESUMO EXECUTIVO
    # =====================================================

    st.markdown("### 📌 Resumo Executivo")

    st.info(f"""

    Se os pedidos do código tarifário:

    • {codigo_origem}

    da transportadora:

    • {transportadora_origem}

    fossem redistribuídos para:

    • {transportadora_destino}

    aproximadamente:

    • {pedidos_simulados:,} pedidos

    seriam redistribuídos entre os códigos equivalentes encontrados.

    """)
        