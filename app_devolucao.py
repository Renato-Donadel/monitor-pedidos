import streamlit as st
import pandas as pd
import os
from io import BytesIO
import plotly.express as px
import matplotlib.pyplot as plt

from utils import (
    carregar_base_devolucao,
    perc_transportes,
    ARQ_DEVOLUCAO
)

def render_devolucao():

    # ==============================
    # INDICADOR DE DEVOLUÇÃO
    # ==============================


    if not os.path.exists(ARQ_DEVOLUCAO):
        st.warning("Base de devolução ainda não foi gerada.")
        st.stop()

    try:

        bases = carregar_base_devolucao()
        dev_atrasada_det = bases["dev_atrasada_detalhado"]
        retornando_det = bases["retornando_detalhado"]
        
        # ==============================
        # AJUSTE Sem_Data_Coleta
        # ==============================

        for df in [retornando_det, dev_atrasada_det]:
            if "Mes" in df.columns:
                df["Mes"] = df["Mes"].replace("Sem_Data_Coleta", "1977-07-01")
        

        vendas_mes = bases["vendas_mes"]
        vendas_mes_pedido = bases["vendas_mes_pedido"]
        vendas_transp = bases["vendas_transportadora"]
        potencial = bases["potencial_triplo"]
        devolucao_proc = bases["devolucao_processo"]
        retornando_transp = bases["retornando_transportes"]
        if "Mes" in retornando_transp.columns:
            retornando_transp["Mes"] = retornando_transp["Mes"].replace("Sem_Data_Coleta", "1977-07-01")
        devolucao_atras = bases["devolucao_atrasada"]
        nfd_mes = bases["nfd_mes"]
        nfd_coleta = bases["nfd_coleta"]

    except Exception as e:
        st.error(f"Erro ao ler base de devolução: {e}")
        st.stop()

    # PADRONIZAR TRANSPORTADORAS PARA O MERGE
    vendas_transp["Transportadora"] = (
        vendas_transp["Transportadora"]
        .astype(str)
        .str.strip()
        .str.upper()
    )
    
    potencial["Transportadora"] = (
        potencial["Transportadora"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # ordenar os dados
    vendas_mes = vendas_mes.sort_values("Mes")
    vendas_transp = vendas_transp.sort_values(["Mes", "Transportadora"])
    
    # ==============================
    # FORMATAR NOME DOS MESES
    # ==============================

    vendas_transp["Mes"] = pd.to_datetime(vendas_transp["Mes"].astype(str))
    vendas_transp["Mes"] = vendas_transp["Mes"].dt.strftime("%B %Y").str.capitalize()
    
    potencial["Transportadora"] = (
        potencial["Transportadora"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    potencial["Mes"] = pd.to_datetime(potencial["Mes"].astype(str))
    potencial["Mes"] = potencial["Mes"].dt.strftime("%B %Y").str.capitalize()
    
    devolucao_proc["Transportadora"] = (
        devolucao_proc["Transportadora"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    devolucao_proc["Mes"] = pd.to_datetime(devolucao_proc["Mes"].astype(str))
    devolucao_proc["Mes"] = devolucao_proc["Mes"].dt.strftime("%B %Y").str.capitalize()
    
    devolucao_atras["Transportadora"] = (
        devolucao_atras["Transportadora"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    devolucao_atras["Mes"] = pd.to_datetime(devolucao_atras["Mes"].astype(str))
    devolucao_atras["Mes"] = devolucao_atras["Mes"].dt.strftime("%B %Y").str.capitalize()
    
    if "Transportadora" in nfd_mes.columns:
        nfd_mes["Transportadora"] = (
            nfd_mes["Transportadora"]
            .astype(str)
            .str.strip()
            .str.upper()
        )
    
    nfd_mes["Mes_NFD"] = pd.to_datetime(nfd_mes["Mes_NFD"].astype(str))
    nfd_mes["Mes_NFD"] = nfd_mes["Mes_NFD"].dt.strftime("%B %Y").str.capitalize()
    
    if "Transportadora" in nfd_coleta.columns:
        nfd_coleta["Transportadora"] = (
            nfd_coleta["Transportadora"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

    nfd_coleta["Mes_Coleta"] = pd.to_datetime(nfd_coleta["Mes_Coleta"].astype(str))
    nfd_coleta["Mes_Coleta"] = nfd_coleta["Mes_Coleta"].dt.strftime("%B %Y").str.capitalize()
    
    
    # ==============================
    # FILTROS
    # ==============================

    meses = sorted(vendas_transp["Mes"].dropna().unique())
    transportadoras = sorted(vendas_transp["Transportadora"].unique())
    
    # ==============================
    # CONTROLE DE ABERTURA DOS FILTROS
    # ==============================

    if "abrir_filtro_mes" not in st.session_state:
        st.session_state["abrir_filtro_mes"] = False

    if "abrir_filtro_transp" not in st.session_state:
        st.session_state["abrir_filtro_transp"] = False

    col_space1, col_btn1, col_btn2, col_space2 = st.columns([2,1,1,2])

    with col_btn1:
        if st.button("Mês"):
            st.session_state["abrir_filtro_mes"] = not st.session_state["abrir_filtro_mes"]

    with col_btn2:
        if st.button("Transportadora"):
            st.session_state["abrir_filtro_transp"] = not st.session_state["abrir_filtro_transp"]


    # ==============================
    # FILTRO MÊS
    # ==============================

    if st.session_state["abrir_filtro_mes"]:

        filtro_mes = st.multiselect(
            "Filtrar mês",
            options=meses,
            default=meses
        )

        if st.button("Aplicar filtro mês"):
            st.session_state["abrir_filtro_mes"] = False

    else:
        filtro_mes = meses


    # ==============================
    # FILTRO TRANSPORTADORA
    # ==============================

    if st.session_state["abrir_filtro_transp"]:

        filtro_transportadora = st.multiselect(
            "Filtrar transportadora",
            options=transportadoras,
            default=transportadoras
        )

        if st.button("Aplicar filtro transportadora"):
            st.session_state["abrir_filtro_transp"] = False

    else:
        filtro_transportadora = transportadoras
        
    # filtros já existem aqui

    retornando_det["ValorNota"] = retornando_det["ValorNota"].fillna(0)

    retornando_det["DataColeta"] = pd.to_datetime(retornando_det["DataColeta"], errors="coerce")

    retornando_det["Mes"] = (
        retornando_det["DataColeta"]
        .dt.strftime("%B %Y")
        .str.capitalize()
    )

    retornando_total = retornando_det[
        retornando_det["Mes"].isin(filtro_mes) &
        retornando_det["Transportadora"].isin(filtro_transportadora)
    ]["ValorNota"].sum()   
    
          
    # ==============================
    # TOTAIS DOS INDICADORES
    # ==============================

    venda_total = vendas_transp[
        vendas_transp["Mes"].isin(filtro_mes) &
        vendas_transp["Transportadora"].isin(filtro_transportadora)
    ]["ValorVenda"].sum()

    devolucao_total = devolucao_proc[
        devolucao_proc["Mes"].isin(filtro_mes) &
        devolucao_proc["Transportadora"].isin(filtro_transportadora)
    ]["Devolucao_Processo"].sum()

    potencial_total = potencial[
        potencial["Mes"].isin(filtro_mes) &
        potencial["Transportadora"].isin(filtro_transportadora)
    ]["Potencial"].sum()

    devolucao_atras_total = devolucao_atras[
        devolucao_atras["Mes"].isin(filtro_mes) &
        devolucao_atras["Transportadora"].isin(filtro_transportadora)
    ]["Devolucao_Atrasada"].sum()


    # percentuais
    perc_devolucao = devolucao_total / venda_total if venda_total > 0 else 0
    perc_potencial = potencial_total / venda_total if venda_total > 0 else 0
    perc_atrasada = devolucao_atras_total / venda_total if venda_total > 0 else 0
    
    # ==============================
    # FILTRO NAS NFD POR TRANSPORTADORA
    # ==============================

    nfd_mes = nfd_mes[
        nfd_mes["Mes_NFD"].isin(filtro_mes)
    ]

    nfd_coleta = nfd_coleta[
        nfd_coleta["Mes_Coleta"].isin(filtro_mes)
    ]


    # formatação
    def moeda(x):
        return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
    col_transp, col_brav = st.columns([2,2])
    
    hoje = pd.Timestamp.today().normalize()
    fim_mes = hoje + pd.offsets.MonthEnd(0)
    dias_restantes_mes = (fim_mes - hoje).days
              
    # =====================================
    # DEVOLUÇÃO - PAINEL TRANSPORTES
    # =====================================

    with col_transp:
    
            st.markdown(
                '<div class="titulo-painel">Devolução - Painel Transportes</div>',
                unsafe_allow_html=True
            )
        
            # base completa
            ret = retornando_transp.copy()
            ret["Mes"] = ret["Mes"].replace("Sem_Data_Coleta", "1977-07-01")
            base_dev_transportes = bases["base"].copy()
            base_dev_bravium = bases["base"].copy()
            base_dev = base_dev_transportes

            ret["Mes"] = pd.to_datetime(ret["Mes"].astype(str))
            ret["Mes"] = ret["Mes"].dt.strftime("%B %Y").str.capitalize()

            ret = ret[
                ret["Mes"].isin(filtro_mes) &
                ret["Transportadora"].isin(filtro_transportadora)
            ]
            
            venda_mes = vendas_transp[
                vendas_transp["Mes"].isin(filtro_mes) &
                vendas_transp["Transportadora"].isin(filtro_transportadora)
            ]["ValorVenda"].sum()

            nfd_total = nfd_coleta["Valor_NFD"].sum()
            
            # ===== ADICIONE ESTAS 3 LINHAS =====
            indice_nfd = nfd_total
            atrasado = devolucao_atras_total
            indice_atrasado = nfd_total + atrasado
                        
            
            # converter datas
            base_dev["DataColeta"] = pd.to_datetime(base_dev["DataColeta"], errors="coerce")
            base_dev["DataÚltimoStatus"] = pd.to_datetime(base_dev["DataÚltimoStatus"], errors="coerce")

            # mês baseado na coleta (Intelipost)
            base_dev["MesFiltro"] = base_dev["DataColeta"].dt.strftime("%B %Y").str.capitalize()

            # aplicar filtro de mês
            base_dev = base_dev[
                base_dev["MesFiltro"].isin(filtro_mes)
            ]
            # =====================================
            # PRAZO DEVOLUÇÃO
            # =====================================

            base_dev["PrazoFinal"] = base_dev["DataÚltimoStatus"] + pd.Timedelta(days=30)

            base_dev["DiasRestantesPrazo"] = (
                base_dev["PrazoFinal"] - hoje
            ).dt.days

            # =====================================
            # IMPACTO RETORNANDO POR MÊS DE COLETA
            # =====================================

            impacto_retornando = retornando_transp.copy()
            impacto_retornando["Mes"] = impacto_retornando["Mes"].replace("Sem_Data_Coleta", "1977-07-01")
            
            impacto_retornando["Mes"] = pd.to_datetime(impacto_retornando["Mes"].astype(str))
            impacto_retornando["Mes"] = impacto_retornando["Mes"].dt.strftime("%B %Y").str.capitalize()

            impacto_retornando = impacto_retornando[
                impacto_retornando["Mes"].isin(filtro_mes) &
                impacto_retornando["Transportadora"].isin(filtro_transportadora)
            ]

            impacto_retornando = impacto_retornando.rename(
                columns={"Mes": "MesColeta"}
            )  

            impacto_retornando = (
                impacto_retornando
                .groupby("MesColeta")["Impacto"]
                .sum()
                .reset_index()
            )
                      
            # ==============================
            # NFD POR MÊS DE COLETA
            # ==============================

            nfd_por_mes = (
                nfd_coleta
                .groupby("Mes_Coleta")["Valor_NFD"]
                .sum()
                .reset_index()
            )

            nfd_por_mes = nfd_por_mes.rename(
                columns={
                    "Mes_Coleta": "Mes",
                    "Valor_NFD": "NFD"
                }
            )

            # =====================================
            # POTENCIAL TRIPLO
            # =====================================

            triplo_real = potencial["Potencial"].sum()
            
            # ==============================
            # IMPACTO POR MÊS DE COLETA
            # ==============================

            impacto_mes = (
                potencial
                .groupby("Mes")["Potencial"]
                .sum()
                .reset_index()
            )

            impacto_mes = impacto_mes.rename(
                columns={
                    "Mes": "MesColeta",
                    "Potencial": "ValorNota"
                }
            )   

            # =====================================
            # FUNÇÕES FORMATAÇÃO
            # =====================================

            # formatação
            def moeda(x):
                return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            
            # =====================================
            # EXIBIÇÃO
            # =====================================

            st.markdown("### Venda total")
            
            st.markdown(
                f"""
                <div style="
                    background:white;
                    padding:8px 12px;
                    border-radius:10px;
                    box-shadow:0 2px 6px rgba(0,0,0,0.06);
                    text-align:center;
                    display:inline-block;
                    min-width:220px;
                    margin-bottom:30px;
                ">
                    <div style="font-size:12px;color:#6b7280;">Venda Total</div>
                    <div style="font-size:18px;font-weight:800;color:#0f2a44;">
                        {moeda(venda_mes)}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )
                    
            # ==============================
            # GRÁFICO VENDAS POR MÊS - TRANSPORTES
            # ==============================

            graf_vendas = vendas_mes.copy()

            graf_vendas["Mes"] = pd.to_datetime(graf_vendas["Mes"].astype(str))

            graf_vendas = graf_vendas.sort_values("Mes")

            fig = px.line(
                graf_vendas,
                x="Mes",
                y="ValorVenda",
                markers=True,
                title="Venda mensal - Transportadoras"
            )

            fig.update_layout(
                xaxis_title="Mês",
                yaxis_title="Valor",
                height=500
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )

            st.markdown("### NFD gerada")
            st.write(f"{moeda(nfd_total)} | {perc_transportes(indice_nfd, venda_mes)}")

            st.markdown("### Atrasado")
            st.write(f"{moeda(atrasado)} | {perc_transportes(indice_atrasado, venda_mes)}")

            st.markdown("### Retornando")
            
            st.write(f"Total: {moeda(retornando_total)}")

            for _, row in impacto_retornando.iterrows():

                mes = row["MesColeta"]
                impacto_valor = row["Impacto"]

                venda_mes_base = graf_vendas[
                    graf_vendas["Mes"].dt.strftime("%B %Y").str.capitalize() == mes
                ]["ValorVenda"].sum()

                nfd_mes_base = nfd_por_mes[
                    nfd_por_mes["Mes"] == mes
                ]["NFD"].sum()

                indice_atual = nfd_mes_base
                indice_novo = nfd_mes_base + impacto_valor

                st.write(
                    f"Impacto {mes}: "
                    f"{moeda(impacto_valor)} | "
                    f"{perc_transportes(indice_atual, venda_mes_base)} → "
                    f"{perc_transportes(indice_novo, venda_mes_base)}"
                )
            
            st.markdown("### Potencial Triplo Prazo")

            st.write(f"Total potencial: {moeda(triplo_real)}")
            
            for _, row in impacto_mes.iterrows():

                mes = row["MesColeta"]
                impacto_valor = row["ValorNota"]

                venda_mes_base = graf_vendas[
                    graf_vendas["Mes"].dt.strftime("%B %Y").str.capitalize() == mes
                ]["ValorVenda"].sum()

                nfd_mes_base = nfd_por_mes[
                    nfd_por_mes["Mes"] == mes
                ]["NFD"].sum()

                indice_atual = nfd_mes_base
                indice_novo = nfd_mes_base + impacto_valor

                st.write(
                    f"Impacto {mes}: "
                    f"{moeda(impacto_valor)} | "
                    f"{perc_transportes(indice_atual, venda_mes_base)} → "
                    f"{perc_transportes(indice_novo, venda_mes_base)}"
                )

            st.markdown("</div>", unsafe_allow_html=True)
            
            # ==============================
            # GRÁFICO IMPACTO POTENCIAL
            # ==============================

            graf_indice = vendas_mes.copy()

            graf_indice["Mes"] = pd.to_datetime(graf_indice["Mes"].astype(str))

            graf_indice = graf_indice.sort_values("Mes")

            # transformar mes em periodo
            graf_indice["MesPeriodo"] = graf_indice["Mes"].dt.to_period("M")

            # mapa impacto por mes de coleta
            impacto_map = dict(
                zip(impacto_mes["MesColeta"], impacto_mes["ValorNota"])
            )

            graf_indice["Impacto"] = graf_indice["MesPeriodo"].map(impacto_map).fillna(0)

            # índice atual (NFD daquele mês / venda daquele mês)
            
            graf_indice["Mes"] = graf_indice["Mes"].astype(str)
            nfd_por_mes["Mes"] = nfd_por_mes["Mes"].astype(str)
            graf_indice = graf_indice.merge(
                nfd_por_mes,
                on="Mes",
                how="left"
            )

            graf_indice["NFD"] = graf_indice["NFD"].fillna(0)

            graf_indice["IndiceAtual"] = graf_indice["NFD"] / graf_indice["ValorVenda"]

            # índice considerando potencial
            graf_indice["IndicePotencial"] = (
                (graf_indice["NFD"] + graf_indice["Impacto"]) /
                graf_indice["ValorVenda"]
            )

            # ==============================
            # PLOT
            # ==============================

            fig = px.line(
                graf_indice,
                x="Mes",
                y=["IndiceAtual", "IndicePotencial"],
                markers=True,
                title="Impacto Potencial Triplo no Índice de Devolução"
            )

            fig.update_layout(
                yaxis_title="Índice %",
                xaxis_title="Mês",
                height=450
            )

            fig.update_yaxes(
                tickformat=".2%"
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )
            
    # =====================================
    # DEVOLUÇÃO - PAINEL BRAVIUM
    # =====================================

    with col_brav:

        st.markdown(
            '<div class="titulo-painel">Devolução - Painel Bravium</div>',
            unsafe_allow_html=True
        )

        # ==============================
        # BASE NFD (empresa)
        # ==============================

        base_nfd = retornando_det.copy()
        
        base_nfd["DataÚltimoStatus"] = pd.to_datetime(base_nfd["DataÚltimoStatus"], errors="coerce")
        base_nfd["DataColeta"] = pd.to_datetime(base_nfd["DataColeta"], errors="coerce")

        hoje = pd.Timestamp.today().normalize()
        fim_mes = hoje + pd.offsets.MonthEnd(0)
        dias_restantes_mes = (fim_mes - hoje).days

        # ==============================
        # FILTRO (POR COLETA)
        # ==============================

        base_nfd["MesFiltro"] = base_nfd["DataColeta"].dt.strftime("%B %Y").str.capitalize()

        base_nfd = base_nfd[
            (base_nfd["MesFiltro"].isin(filtro_mes)) &
            (base_nfd["Transportadora"].isin(filtro_transportadora))
        ]

        # ==============================
        # DIAS NO STATUS (BASE DA REGRA)
        # ==============================

        base_nfd["DiasNoStatus"] = (
            hoje - base_nfd["DataÚltimoStatus"]
        ).dt.days

        # ==============================
        # CLASSIFICAÇÃO CORRETA
        # ==============================

        # atrasado (já existente)
        base_atrasado = devolucao_atras.copy()

        base_atrasado = base_atrasado[
            base_atrasado["Mes"].isin(filtro_mes) &
            base_atrasado["Transportadora"].isin(filtro_transportadora)
        ]

        atrasado_brav = base_atrasado["Devolucao_Atrasada"].sum()
        
        # PROVÁVEL
        provavel_brav = base_nfd[
            (base_nfd["DiasNoStatus"] >= 20) &
            (dias_restantes_mes >= 10)
        ]["ValorNota"].sum()

        # IMPROVÁVEL
        improv_brav = base_nfd[
            (base_nfd["DiasNoStatus"] < 10) &
            (dias_restantes_mes <= 10)
        ]["ValorNota"].sum()

        # POSSÍVEL = resto
        poss_brav = base_nfd[
            ~(
                ((base_nfd["DiasNoStatus"] >= 20) & (dias_restantes_mes >= 10)) |
                ((base_nfd["DiasNoStatus"] < 10) & (dias_restantes_mes <= 10))
            )
        ]["ValorNota"].sum()

        # ==============================
        # NFD EMPRESA
        # ==============================

        vendas_mes_pedido["Mes_Pedido"] = pd.to_datetime(
            vendas_mes_pedido["Mes_Pedido"].astype(str)
        )
        vendas_mes_pedido["Mes_Pedido"] = vendas_mes_pedido["Mes_Pedido"].dt.strftime("%B %Y").str.capitalize()

        venda_mes = vendas_mes_pedido[
            vendas_mes_pedido["Mes_Pedido"].isin(filtro_mes)
        ]["ValorVenda"].sum()

        nfd_empresa = nfd_mes["Valor_NFD"].sum()

        # ==============================
        # INDICES
        # ==============================

        indice_brav_nfd = nfd_empresa

        indice_brav_atras = nfd_empresa + atrasado_brav

        indice_brav_prov = nfd_empresa + atrasado_brav + provavel_brav

        indice_brav_poss = nfd_empresa + atrasado_brav + provavel_brav + poss_brav

        indice_brav_improv = nfd_empresa + atrasado_brav + provavel_brav + poss_brav + improv_brav

        potencial_brav = potencial[
            potencial["Mes"].isin(filtro_mes)
        ]["Potencial"].sum()

        indice_brav_potencial_1 = (
            nfd_empresa
            + atrasado_brav
            + potencial_brav
        )

        indice_brav_potencial_2 = (
            nfd_empresa
            + atrasado_brav
            + provavel_brav
            + potencial_brav
        )

        indice_brav_potencial_poss = (
            nfd_empresa
            + atrasado_brav
            + provavel_brav
            + poss_brav
            + potencial_brav
        )
        # ==============================
        # FUNÇÃO PERCENTUAL BRAVIUM
        # ==============================

        def perc_bravium(x):
            if venda_mes == 0:
                return "0%"
            return f"{(x / venda_mes)*100:.2f}%".replace(".", ",")

        # ==============================
        # EXIBIÇÃO
        # ==============================
        
        st.markdown("### Venda total(Empresa)")

        st.markdown(
            f"""
            <div style="
                background:white;
                padding:8px 12px;
                border-radius:10px;
                box-shadow:0 2px 6px rgba(0,0,0,0.06);
                text-align:center;
                display:inline-block;
                min-width:220px;
                margin-bottom:30px;
            ">
                <div style="font-size:12px;color:#6b7280;">Venda Total</div>
                <div style="font-size:18px;font-weight:800;color:#0f2a44;">
                    {moeda(venda_mes)}
                </div>
            </div>
            """,
                
        )
        
        # ==============================
        # GRÁFICO VENDAS POR MÊS - BRAVIUM
        # ==============================
        
        graf_bravium = bases["vendas_mes_pedido"].copy()

        graf_bravium["Mes_Pedido"] = pd.to_datetime(
            graf_bravium["Mes_Pedido"].astype(str)
        )

        graf_bravium = graf_bravium.sort_values("Mes_Pedido")

        fig = px.line(
            graf_bravium,
            x="Mes_Pedido",
            y="ValorVenda",
            markers=True,
            title="Venda mensal - Bravium"
        )

        fig.update_layout(
            xaxis_title="Mês",
            yaxis_title="Valor",
            height=500
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

        st.markdown("### NFD gerada (Empresa)")
        st.write(f"{moeda(nfd_empresa)} | {perc_bravium(indice_brav_nfd)}")

        st.markdown("### Atrasado")
        st.write(f"{moeda(atrasado_brav)} | {perc_bravium(indice_brav_atras)}")

        st.markdown("### Retornando")

        st.write(f"Total: {moeda(retornando_total)}")

        st.write(
            f"Provável: {moeda(provavel_brav)} | "
            f"{perc_bravium(indice_brav_atras)} → {perc_bravium(indice_brav_prov)}"
        )

        st.write(
            f"Possível: {moeda(poss_brav)} | "
            f"{perc_bravium(indice_brav_prov)} → {perc_bravium(indice_brav_poss)}"
        )

        st.write(
            f"Improvável: {moeda(improv_brav)} | "
            f"{perc_bravium(indice_brav_poss)} → {perc_bravium(indice_brav_improv)}"
        )
       

        st.markdown("### Potencial no mês (Bravium)")

        # 1. Potencial puro
        st.write(
            f"Cenário Potencial: {moeda(potencial_brav)} | "
            f"{perc_bravium(indice_brav_atras)} → {perc_bravium(indice_brav_potencial_1)}"
        )

        # 2. Potencial + provável
        st.write(
            f"Cenário Potencial + Provável: {moeda(potencial_brav)} | "
            f"{perc_bravium(indice_brav_potencial_1)} → {perc_bravium(indice_brav_potencial_2)}"
        )

        # 3. Potencial + provável + possível
        st.write(
            f"Cenário Potencial + Provável + Possível: {moeda(potencial_brav)} | "
            f"{perc_bravium(indice_brav_potencial_2)} → {perc_bravium(indice_brav_potencial_poss)}"
        )