import streamlit as st
import pandas as pd
from io import BytesIO
import plotly.express as px
import matplotlib.pyplot as plt

def render_monitor():

    

    # ==============================
    # DOWNLOAD POR CARTEIRA (COM ABA EXTRA)
    # ==============================
    st.markdown("### 📥 Exportação por Carteira (300 em 300)")
    if st.button("🔄 Atualizar dados"):
        st.cache_data.clear()
        st.rerun()

    df_atual = ler_base(ARQ_ATUAL)
    # 🔥 NORMALIZA CARTEIRA (CORRIGE IGOR SUMINDO)
    if "Carteira" in df_atual.columns:
        df_atual["Carteira"] = (
            df_atual["Carteira"]
            .astype(str)
            .str.strip()
        )
    # 🔥 GARANTE DIAS COMO NUMÉRICO (IGOR)
    if "DiasDesdeUltimoStatus" in df_atual.columns:
        df_atual["DiasDesdeUltimoStatus"] = pd.to_numeric(
            df_atual["DiasDesdeUltimoStatus"],
            errors="coerce"
        )
    # 🔥 NORMALIZA STATUS
    if "Status" in df_atual.columns:
        df_atual["Status"] = (
            df_atual["Status"]
            .astype(str)
            .str.upper()
            .str.strip()
        )
   
    # ==============================
    # 🚚 KPI GERAL - PRAZO TRANSPORTADORA
    # ==============================

    if not df_atual.empty and "PrazoTransportadorDiasUteis" in df_atual.columns:

        df_tmp = df_atual.copy()

        # 🚫 EXCLUIR Status fora de desempenho
        status_devolucao = [
            "TSP - Aguardando Confirmar Devolução",
            "TSP - Trânsferência para Devolução",
            "TSP - Rota de Devolução",
            "TSP - Reentrega",
            "TSP - Aguardando Tratativa Transportadora",
            "TSP - Coleta Realizada",
            "TSP - Item Faltante",
            "TSP - REENTREGAR/ENDERECO CORRETO"
        ]

        df_tmp = df_tmp[~df_tmp["Status"].isin(status_devolucao)]

        # 🚫 EXCLUIR AGUARDANDO EXPEDIÇÃO
        df_tmp = df_tmp[
            ~df_tmp["Status"].str.contains("AGUARDANDO EXPED", na=False)
        ]

        # garante tipo numérico
        df_tmp["DiasDesdeExpedicao"] = pd.to_numeric(df_tmp["DiasDesdeExpedicao"], errors="coerce")
        df_tmp["PrazoTransportadorDiasUteis"] = pd.to_numeric(df_tmp["PrazoTransportadorDiasUteis"], errors="coerce")

        # dentro / fora baseado SÓ na transportadora
        dentro = len(df_tmp[
            df_tmp["DiasDesdeExpedicao"] <= df_tmp["PrazoTransportadorDiasUteis"]
        ])

        fora = len(df_tmp[
            df_tmp["DiasDesdeExpedicao"] > df_tmp["PrazoTransportadorDiasUteis"]
        ])

        total = dentro + fora

        perc_dentro = (dentro / total * 100) if total > 0 else 0
        perc_fora = (fora / total * 100) if total > 0 else 0

        st.markdown("### 🚚 Prazo Transportadora (Geral)")

        st.markdown(
            f"""
            <div style="
                background:white;
                padding:14px;
                border-radius:12px;
                box-shadow:0 2px 6px rgba(0,0,0,0.06);
                margin-bottom:20px;
            ">
                <div style="font-size:14px;color:#6b7280;">Transportadora</div>
                <div style="font-size:18px;font-weight:700;color:#0f2a44;">
                    Dentro: {dentro} ({perc_dentro:.1f}%) &nbsp;&nbsp;|&nbsp;&nbsp;
                    Fora: {fora} ({perc_fora:.1f}%)
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
    if "offsets_carteira" not in st.session_state:
        st.session_state["offsets_carteira"] = {}

    if not df_atual.empty and "Carteira" in df_atual.columns:

        if "Ranking" in df_atual.columns:
            df_atual = df_atual.sort_values("Ranking").reset_index(drop=True)

        carteiras = sorted(df_atual["Carteira"].dropna().unique())

        # 🔥 REMOVE RENATO DO FRONT
        carteiras = [c for c in carteiras if c != "Renato"]
        carteiras = [c for c in carteiras if c != "Augusto"]

        if "Igor" in df_atual["Carteira"].values and "Igor" not in carteiras:
            carteiras.append("Igor")

        for carteira in carteiras:
        
            
            if carteira == "Renato":
                continue

            if f"next_{carteira}" not in st.session_state:
                st.session_state[f"next_{carteira}"] = False

            df_carteira = df_atual[
                df_atual["Carteira"] == carteira
            ].copy()
            
            if carteira == "Igor" and df_carteira.empty:
                st.write("Igor sem dados — verificar base")
            
            # REMOVE DUPLICIDADE POR PEDIDO
            if "PedidoFormatado" in df_carteira.columns:
                df_carteira = df_carteira.drop_duplicates(subset=["PedidoFormatado"])

            # ==============================
            # 🎯 REGRA POR CARTEIRA
            # ==============================

            if carteira == "Igor":
                df_dentro_prazo = df_carteira[
                df_carteira["Nivel_Igor_30d"] == "Dentro"
                ]

                df_fora_prazo = df_carteira[
                    df_carteira["Nivel_Igor_30d"] == "Fora"
                ].reset_index(drop=True)
                   
            else:

                for col in ["Cliente_Dentro","Transportadora_Dentro","Status_Dentro","Regiao_Dentro"]:
                    df_carteira[col] = df_carteira[col].astype(str).str.strip().str.upper()

                df_dentro_prazo = df_carteira[
                    (df_carteira["Cliente_Dentro"] == "X") &
                    (df_carteira["Transportadora_Dentro"] == "X") &
                    (df_carteira["Status_Dentro"] == "X") &
                    (df_carteira["Regiao_Dentro"] == "X")
                ]

                df_fora_prazo = df_carteira[
                    ~(
                        (df_carteira["Cliente_Dentro"] == "X") &
                        (df_carteira["Transportadora_Dentro"] == "X") &
                        (df_carteira["Status_Dentro"] == "X") &
                        (df_carteira["Regiao_Dentro"] == "X")
                    )
                ].reset_index(drop=True)

            total = len(df_fora_prazo)
            total_dentro = len(df_dentro_prazo)

            offset = st.session_state["offsets_carteira"].get(carteira, 0)
            if offset >= total:
                offset = 0
                st.session_state["offsets_carteira"][carteira] = 0

            if st.session_state[f"next_{carteira}"]:
                offset = offset + TAMANHO_LOTE
                st.session_state["offsets_carteira"][carteira] = offset
                st.session_state[f"next_{carteira}"] = False

            inicio = offset
            fim = min(offset + TAMANHO_LOTE, total)

            lote = df_fora_prazo.iloc[inicio:fim]

            if not lote.empty or carteira == "Igor":
                buffer = BytesIO()

                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    lote.to_excel(writer, index=False, sheet_name="Lote")

                    df_status = df_carteira[
                        df_carteira["Status"].isin(STATUS_DIARIOS)
                    ]

                    if not df_status.empty:
                        df_status.to_excel(
                            writer,
                            index=False,
                            sheet_name="Status Diários"
                        )

                buffer.seek(0)

                perc_fora = (total / (total + total_dentro)) * 100 if (total + total_dentro) > 0 else 0
                barra = int(perc_fora // 5)

                col1, col2 = st.columns([4, 2])

                total_geral = total + total_dentro

                perc_dentro = (total_dentro / total_geral) * 100 if total_geral > 0 else 0
                perc_fora = (total / total_geral) * 100 if total_geral > 0 else 0

                with col1:
                    st.write(f"**{carteira}** — {inicio+1} até {fim} de {total}")
                    st.write(
                        f"Dentro do prazo: {total_dentro} ({perc_dentro:.1f}%) | "
                        f"Fora do prazo: {total} ({perc_fora:.1f}%)"
                    )
                   
                with col2:
                    if st.download_button(
                        label=f"⬇️ Baixar {carteira}",
                        data=buffer,
                        file_name=f"{carteira}_{inicio+1}_a_{fim}.xlsx",
                        key=f"dl_{carteira}_{offset}"
                    ):
                        st.session_state[f"next_{carteira}"] = True
                        st.rerun()
                                         
                  
    # ==============================
    # 🧑 CARTEIRA AUGUSTO (NOVA)
    # ==============================

    if "Carteira" in df_atual.columns:

        df_augusto = df_atual[
            df_atual["Carteira"] == "Augusto"
        ].copy()

    else:
        df_augusto = pd.DataFrame()

    # REMOVE DUPLICIDADE POR PEDIDO
    if "PedidoFormatado" in df_augusto.columns:
        df_augusto = df_augusto.drop_duplicates(subset=["PedidoFormatado"])

    total_augusto = len(df_augusto)

    col1, col2 = st.columns([4,2])

    with col1:
        st.write(f"**Augusto** — 1 até {total_augusto} de {total_augusto}")
        df_augusto["Cliente_Dentro"] = df_augusto["Cliente_Dentro"].astype(str).str.strip().str.upper()
        df_augusto["Transportadora_Dentro"] = df_augusto["Transportadora_Dentro"].astype(str).str.strip().str.upper()
        df_augusto["Status_Dentro"] = df_augusto["Status_Dentro"].astype(str).str.strip().str.upper()
        df_augusto["Regiao_Dentro"] = df_augusto["Regiao_Dentro"].astype(str).str.strip().str.upper()

        dentro = len(df_augusto[
            (df_augusto["Cliente_Dentro"] == "X") &
            (df_augusto["Transportadora_Dentro"] == "X") &
            (df_augusto["Status_Dentro"] == "X") &
            (df_augusto["Regiao_Dentro"] == "X")
        ])

        fora = len(df_augusto) - dentro

        total = dentro + fora

        perc_dentro = (dentro / total * 100) if total > 0 else 0
        perc_fora = (fora / total * 100) if total > 0 else 0

        st.write(
            f"Dentro do prazo: {dentro} ({perc_dentro:.1f}%) | "
            f"Fora do prazo: {fora} ({perc_fora:.1f}%)"
        )

    with col2:

        buffer = BytesIO()
        df_augusto.to_excel(buffer, index=False)
        buffer.seek(0)

        st.download_button(
            label="⬇️ Baixar Augusto",
            data=buffer,
            file_name="augusto.xlsx",
            key="dl_augusto"
        )

    st.divider()

    # ==============================
    # 🚚 EXPEDIÇÃO — 3+ DIAS NO STATUS
    # ==============================
    st.markdown("### 🚚 Expedição (3+ dias no status)")

    if not df_atual.empty:

        if (
            "Status" in df_atual.columns and
            "DiasDesdeUltimoStatus" in df_atual.columns
        ):

            df_expedicao = df_atual[
                (df_atual["Status"].str.contains("AGUARDANDO EXPED", na=False)) &
                (df_atual["DiasDesdeUltimoStatus"] >= 3)
            ].copy()

            total = len(df_expedicao)

            col1, col2 = st.columns([4, 2])

            with col1:
                st.write(
                    f"Pedidos aguardando expedição há ≥ 3 dias úteis: **{total}**"
                )

            with col2:
                if total > 0:

                    colunas_exportar = [
                        c for c in [
                            "PedidoFormatado",
                            "NotaFiscal",
                            "Armazém",
                            "Logistica",
                            "DiasDesdeUltimoStatus"
                        ] if c in df_expedicao.columns
                    ]

                    df_export = df_expedicao[colunas_exportar].copy()

                    df_export = df_export.rename(columns={
                        "DiasDesdeUltimoStatus": "Dias_Parado_no_Status"
                    })

                    df_export = df_export.sort_values(
                        "Dias_Parado_no_Status",
                        ascending=False
                    )

                    buffer = BytesIO()
                    df_export.to_excel(buffer, index=False)
                    buffer.seek(0)

                    st.download_button(
                        label="⬇️ Baixar Expedição (3+ dias)",
                        data=buffer,
                        file_name="expedicao_parada_3_dias_ou_mais.xlsx",
                        key="download_expedicao_3dias"
                    )
                else:
                    st.write("Nenhum pedido elegível.")
                    
            # ==============================
            # 🚨 STATUS DIVERGENTES
            # ==============================
            st.markdown("### 🚨 Status Divergentes PRW x Intelipost")

            if not df_atual.empty:

                if "Divergencia_Intelipost" in df_atual.columns:

                    df_divergente = df_atual[
                        df_atual["Divergencia_Intelipost"] == True
                    ].copy()

                    total_divergente = len(df_divergente)

                    col1, col2 = st.columns([4, 2])

                    with col1:

                        st.write(
                            f"Pedidos divergentes entre PRW e Intelipost: **{total_divergente}**"
                        )

                    with col2:

                        if total_divergente > 0:

                            colunas_exportar = [
                                c for c in [
                                    "PedidoFormatado",
                                    "NotaFiscal",
                                    "Status",
                                    "Status Transportador",
                                    "DataÚltimoStatus",
                                    "Data do Último Status",
                                    "Logistica"
                                ]
                                if c in df_divergente.columns
                            ]

                            df_export = df_divergente[
                                colunas_exportar
                            ].copy()

                            buffer = BytesIO()

                            df_export.to_excel(
                                buffer,
                                index=False
                            )

                            buffer.seek(0)

                            st.download_button(
                                label="⬇️ Baixar Divergentes",
                                data=buffer,
                                file_name="status_divergentes.xlsx",
                                key="download_divergentes"
                            )

                        else:

                            st.write("Nenhum pedido divergente.")                

    # ==============================
    # BI EXECUTIVO
    # ==============================
    dias = listar_dias()

    if len(dias) < 2:
        st.warning("Histórico insuficiente na pasta data/historico.")
        st.stop()

    st.markdown("### 📈 Status Diários por Mês")

    contagem = []

    mes_atual = pd.Timestamp.today().to_period("M")

    for dia_hist in dias:
    
        data = pd.to_datetime(dia_hist, format="%d-%m-%Y")

        data = pd.to_datetime(dia_hist, format="%d-%m-%Y")
        mes = data.to_period("M")

        arquivo_mes = os.path.join(PASTA_MENSAL, f"{mes}.xlsx")

        # Se o mês já estiver congelado
        if mes != mes_atual and os.path.exists(arquivo_mes):

            df_mes = pd.read_excel(arquivo_mes)

            for _, row in df_mes.iterrows():
                contagem.append((row["Data"], row["Qtd"]))

            continue

        # Caso contrário calcula normalmente
        path = caminho(dia_hist)
        df_temp = ler_base(path)

        if df_temp.empty:
            continue

        df_temp["Status"] = (
            df_temp["Status"]
            .astype(str)
            .str.upper()
            .str.strip()
        )

        def limpar_status(s):
            return (
                str(s)
                .upper()
                .strip()
                .replace("Ç", "C")
                .replace("Ã", "A")
                .replace("Á", "A")
            )

        status_validos = [limpar_status(s) for s in STATUS_DIARIOS]

        df_temp["Status"] = df_temp["Status"].apply(limpar_status)

        qtd = df_temp[
            df_temp["Status"].isin(status_validos)
        ].shape[0]

        contagem.append((data, qtd))

    if contagem:

        df_graf = pd.DataFrame(contagem, columns=["Data", "Qtd"])
        df_graf["Data"] = pd.to_datetime(df_graf["Data"], format="%d-%m-%Y")
        df_graf = df_graf.sort_values("Data")
        
        # ==============================
        # SALVAR MÊS FECHADO AUTOMATICAMENTE
        # ==============================

        mes_atual = pd.Timestamp.today().to_period("M")

        for mes in df_graf["Data"].dt.to_period("M").unique():

            if mes == mes_atual:
                continue

            arquivo_mes = os.path.join(PASTA_MENSAL, f"{mes}.xlsx")

            # se ainda não existir, salva
            if not os.path.exists(arquivo_mes):

                df_mes = df_graf[
                    df_graf["Data"].dt.to_period("M") == mes
                ].copy()
                
                # ajuste especial fevereiro 2026
                if mes.month == 2 and mes.year == 2026:
                    df_mes = df_mes[df_mes["Data"].dt.day >= 18]

                df_mes.to_excel(arquivo_mes, index=False)

        # Agrupa por Ano + Mês
        df_graf["AnoMes"] = df_graf["Data"].dt.to_period("M")

        meses = df_graf["AnoMes"].unique()

        colunas = st.columns(len(meses))

        for i, mes in enumerate(meses):

            df_mes = df_graf[df_graf["AnoMes"] == mes]
        
            # Se for fevereiro de 2026 (ajuste o ano se necessário)
            if mes.month == 2 and mes.year == 2026:
                df_mes = df_mes[df_mes["Data"].dt.day >= 18]

            with colunas[i]:
            
                nome_mes = df_mes["Data"].dt.strftime("%B").iloc[0].capitalize()
                ano = df_mes["Data"].dt.year.iloc[0]

                fig = px.line(
                    df_mes,
                    x="Data",
                    y="Qtd",
                    title=f"{nome_mes}/{ano}"
                )

                fig.update_layout(
                    height=220,
                    margin=dict(l=10, r=10, t=40, b=10),
                    xaxis_title="",
                    yaxis_title=""
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True
                )
                
    # ==============================
    # 📌 STATUS MANUAIS (POR MÊS - IGUAL DIÁRIOS)
    # ==============================

    st.markdown("### 📌 Status Manuais")

    contagem = []
    mes_atual = pd.Timestamp.today().to_period("M")

    for dia_hist in dias:
    
        data = pd.to_datetime(dia_hist, format="%d-%m-%Y")

        if data < pd.Timestamp("2026-05-01"):
            continue

        data = pd.to_datetime(dia_hist, format="%d-%m-%Y")
        mes = data.to_period("M")

        arquivo_mes = os.path.join(PASTA_MENSAL, f"manuais_{mes}.xlsx")

        # 🔒 usa mês já salvo
        if (
            mes != mes_atual
            and os.path.exists(arquivo_mes)
            and mes >= pd.Period("2026-05", freq="M")
        ):

            df_mes = pd.read_excel(arquivo_mes)

            for _, row in df_mes.iterrows():
                contagem.append((
                row["Data"],
                row["Total"],
                row.get("Entrou", 0),
                row.get("Tratados", 0)
            ))

            continue

        # 🔄 calcula mês atual
        df_temp = ler_base(caminho(dia_hist))

        if df_temp.empty:
            continue

        def limpar_status(s):
            return (
                str(s)
                .upper()
                .strip()
                .replace("Ç", "C")
                .replace("Ã", "A")
                .replace("Á", "A")
            )

        df_temp["Status"] = df_temp["Status"].apply(limpar_status)
        if "PedidoFormatado" in df_temp.columns:
            df_temp = df_temp.drop_duplicates(subset=["PedidoFormatado"])
        status_validos = [limpar_status(s) for s in STATUS_Manuais]

        atuais = df_temp[
            df_temp["Status"].isin(status_validos)
        ].copy()
        # 🔥 remover duplicidade
        if "PedidoFormatado" in atuais.columns:
            atuais = atuais.drop_duplicates(subset=["PedidoFormatado"])

        # 🔥 carregar dia anterior
        idx = dias.index(dia_hist)

        if idx > 0:
            dia_ant = dias[idx - 1]
            df_ant = ler_base(caminho(dia_ant))

            df_ant["Status"] = df_ant["Status"].apply(limpar_status)

            anteriores = df_ant[
                df_ant["Status"].isin(status_validos)
            ].copy()

            if "PedidoFormatado" in anteriores.columns:
                anteriores = anteriores.drop_duplicates(subset=["PedidoFormatado"])

        else:
            anteriores = pd.DataFrame(columns=atuais.columns)

        # 🔥 métricas
        total = len(atuais)
        tratados = len(anteriores[~anteriores["PedidoFormatado"].isin(atuais["PedidoFormatado"])])
        entrou = len(atuais[~atuais["PedidoFormatado"].isin(anteriores["PedidoFormatado"])])

        contagem.append((data, total, entrou, tratados))

    if contagem:

        df_graf = pd.DataFrame(contagem, columns=["Data", "Total", "Entrou", "Tratados"])
        df_graf["Data"] = pd.to_datetime(df_graf["Data"])
        df_graf = df_graf.sort_values("Data")

        # 🔒 SALVAR MÊS FECHADO
        for mes in df_graf["Data"].dt.to_period("M").unique():

            if mes == mes_atual:
                continue

            arquivo_mes = os.path.join(PASTA_MENSAL, f"manuais_{mes}.xlsx")

            if not os.path.exists(arquivo_mes):

                df_mes = df_graf[
                    df_graf["Data"].dt.to_period("M") == mes
                ].copy()

                df_mes.to_excel(arquivo_mes, index=False)

        # 📊 GRÁFICO POR MÊS (IGUAL DIÁRIOS)
        df_graf["AnoMes"] = df_graf["Data"].dt.to_period("M")

        meses = df_graf["AnoMes"].unique()
        colunas = st.columns(len(meses))

        for i, mes in enumerate(meses):

            df_mes = df_graf[df_graf["AnoMes"] == mes]

            with colunas[i]:

                df_plot = df_mes.melt(
                    id_vars="Data",
                    value_vars=["Total", "Entrou", "Tratados"],
                    var_name="Tipo",
                    value_name="Quantidade"
                )
                
                nome_mes = df_mes["Data"].dt.strftime("%B").iloc[0].capitalize()
                ano = df_mes["Data"].dt.year.iloc[0]

                fig = px.line(
                    df_plot,
                    x="Data",
                    y="Quantidade",
                    color="Tipo",
                    title=f"{nome_mes}/{ano}"
                )

                fig.update_layout(
                    height=250,
                    margin=dict(l=10, r=10, t=40, b=10),
                    xaxis_title="",
                    yaxis_title=""
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True
                )
        
    # ==============================
    # LOOP ORIGINAL COMPLETO (PIZZAS)
    # ==============================

    dias_pizza = dias[-7:]

    if len(dias_pizza) < 2:
        st.warning("Histórico insuficiente para gráfico de pizza.")
    else:
        for i in range(len(dias_pizza)-1, 0, -1):

            dia_atual = dias_pizza[i]
            dia_ant = dias_pizza[i-1]

            df_hist_atual = ler_base(caminho(dia_atual))
            df_hist_ant = ler_base(caminho(dia_ant))
            
            df_hist_atual = df_hist_atual[
                df_hist_atual["Logistica"] != "ATRUS INTERMEDIACAO"
            ]

            df_hist_ant = df_hist_ant[
                df_hist_ant["Logistica"] != "ATRUS INTERMEDIACAO"
            ]

            if df_hist_atual.empty or df_hist_ant.empty:
                continue

            st.markdown(
                f'<p class="data-title">📅 {dia_ant} ➜ {dia_atual}</p>',
                unsafe_allow_html=True
            )

            col1, col2, col3, col4 = st.columns(4)

            # TRIPLO
            with col1:
                if "Transportadora_Triplo" in df_hist_atual.columns:
                
                    atual = df_hist_atual[
                        df_hist_atual["Transportadora_Triplo"]
                        .astype(str)
                        .str.strip()
                        .str.upper() == "X"
                    ]

                    ant = df_hist_ant[
                        df_hist_ant["Transportadora_Triplo"]
                        .astype(str)
                        .str.strip()
                        .str.upper() == "X"
                    ]
                    restantes = ant[ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
                    entrou = atual[~atual["PedidoFormatado"].isin(ant["PedidoFormatado"])]
                    tratados = ant[~ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]

                    valor_entrou = entrou["ValorNota"].sum() if "ValorNota" in entrou.columns else 0
                    valor_restantes = restantes["ValorNota"].sum() if "ValorNota" in restantes.columns else 0

                    valor_entrou_fmt = f"{valor_entrou:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    valor_restantes_fmt = f"{valor_restantes:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

                    st.image(pizza(len(tratados), len(restantes), "Triplo Transportadora"))

                    st.markdown(
                        f'<p class="metric-small">Tratados: {len(tratados)} / {len(ant)}</p>',
                        unsafe_allow_html=True
                    )

                    st.markdown(
                        f'<p class="metric-small">Entraram: {len(entrou)} | R$ {valor_entrou_fmt}</p>',
                        unsafe_allow_html=True
                    )

                    st.markdown(
                        f'<p class="metric-small">Remanescentes: {len(restantes)} | R$ {valor_restantes_fmt}</p>',
                        unsafe_allow_html=True
                    )

                    buf = BytesIO()
                    restantes.to_excel(buf, index=False)
                    st.download_button(
                        "Remanescentes Triplo",
                        buf.getvalue(),
                        file_name=f"remanescente_triplo_{dia_atual}.xlsx"
                    )

            # STATUS 2X
            with col2:
                if "Status_Dobro" in df_hist_atual.columns:

                    df_hist_atual["Status_Dobro"] = df_hist_atual["Status_Dobro"].astype(str).str.strip().str.upper()
                    df_hist_ant["Status_Dobro"] = df_hist_ant["Status_Dobro"].astype(str).str.strip().str.upper()

                    atual = df_hist_atual[
                        (df_hist_atual["Status_Dobro"] == "X") |
                        (df_hist_atual["Status_Triplo"] == "X")
                    ]

                    ant = df_hist_ant[
                        df_hist_ant["Status_Dobro"] == "X"
                    ]

                    tratados = ant[~ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
                    restantes = ant[ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
                    entrou = atual[~atual["PedidoFormatado"].isin(ant["PedidoFormatado"])]
                    tratados = ant[~ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]

                    st.image(pizza(len(tratados), len(restantes), "Status Específico 2x"))

                    st.markdown(
                        f'<p class="metric-small">Tratados: {len(tratados)} / {len(ant)}</p>',
                        unsafe_allow_html=True
                    )

                    st.markdown(
                        f'<p class="metric-small">Entraram: {len(entrou)}</p>',
                        unsafe_allow_html=True
                    )

                    buf = BytesIO()
                    restantes.to_excel(buf, index=False)
                    st.download_button(
                        "Remanescentes Status 2x",
                        buf.getvalue(),
                        file_name=f"remanescente_status_{dia_atual}.xlsx"
                    )

            # REGIÃO 2X
            with col3:
                if "Regiao_Dobro" in df_hist_atual.columns:

                    atual = df_hist_atual[
                        df_hist_atual["Regiao_Dobro"]
                        .astype(str)
                        .str.strip()
                        .str.upper() == "X"
                    ]

                    ant = df_hist_ant[
                        df_hist_ant["Regiao_Dobro"]
                        .astype(str)
                        .str.strip()
                        .str.upper() == "X"
                    ]

                    tratados = ant[~ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
                    restantes = ant[ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
                    entrou = atual[~atual["PedidoFormatado"].isin(ant["PedidoFormatado"])]

                    st.image(pizza(len(tratados), len(restantes), "Região 2x Prazo"))

                    st.markdown(
                        f'<p class="metric-small">Tratados: {len(tratados)} / {len(ant)}</p>',
                        unsafe_allow_html=True
                    )

                    st.markdown(
                        f'<p class="metric-small">Entraram: {len(entrou)}</p>',
                        unsafe_allow_html=True
                    )

                    buf = BytesIO()
                    restantes.to_excel(buf, index=False)
                    st.download_button(
                        "Remanescentes Região 2x",
                        buf.getvalue(),
                        file_name=f"remanescente_regiao_{dia_atual}.xlsx"
                    )
                    
            # CARTEIRA (TOP 300 DE CADA)

            with col4:

                if "Carteira" in df_hist_atual.columns:

                    # ordenar se tiver ranking
                    if "Ranking" in df_atual.columns:
                        df_hist_atual = df_hist_atual.sort_values("Ranking")
                        df_hist_ant = df_hist_ant.sort_values("Ranking")

                    # pegar carteiras
                    carteiras = df_hist_atual["Carteira"].dropna().unique()

                    tratados_total = 0
                    restantes_total = 0
                    entrou_total = 0

                    for c in carteiras:

                        atual_carteira = df_hist_atual[df_hist_atual["Carteira"] == c].head(300)
                        ant_carteira = df_hist_ant[df_hist_ant["Carteira"] == c].head(300)

                        set_atual = set(atual_carteira["PedidoFormatado"])
                        set_ant = set(ant_carteira["PedidoFormatado"])

                        tratados = set_ant - set_atual
                        restantes = set_ant & set_atual
                        entrou = set_atual - set_ant

                        tratados_total += len(tratados)
                        restantes_total += len(restantes)
                        entrou_total += len(entrou)
                    # gráfico
                    st.image(pizza(tratados_total, restantes_total, "Carteiras (Top 300)"))
                    # métricas corretas
                    st.markdown(
                        f'<p class="metric-small">Tratados: {tratados_total} / {tratados_total + restantes_total}</p>',
                        unsafe_allow_html=True
                    )

                    st.markdown(
                        f'<p class="metric-small">Entraram: {entrou_total}</p>',
                        unsafe_allow_html=True
                    )

                    # exportação correta
                    restantes_df = pd.DataFrame()

                    for c in carteiras:

                        atual_carteira = df_hist_atual[df_hist_atual["Carteira"] == c].head(300)
                        ant_carteira = df_hist_ant[df_hist_ant["Carteira"] == c].head(300)

                        restantes_ids = set(ant_carteira["PedidoFormatado"]) & set(atual_carteira["PedidoFormatado"])

                        restantes_df = pd.concat([
                            restantes_df,
                            ant_carteira[ant_carteira["PedidoFormatado"].isin(restantes_ids)]
                        ])

                    buf = BytesIO()
                    restantes_df.to_excel(buf, index=False)

                    st.download_button(
                        "Remanescentes Carteiras",
                        buf.getvalue(),
                        file_name=f"remanescente_carteiras_{dia_atual}.xlsx"
                    )

            st.divider()