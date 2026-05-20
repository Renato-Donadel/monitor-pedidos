import streamlit as st
import pandas as pd
import os
from io import BytesIO
import plotly.express as px
import matplotlib.pyplot as plt

from utils import PASTA_DATA

def render_desempenho():


    # ==============================
    # DESEMPENHO POR TRANSPORTADORA
    # ==============================


    st.markdown("### 🚚 Desempenho por Transportadora")

    ARQ_TRANSP = os.path.join(PASTA_DATA, "Base_Transportadora.xlsx")
    
    if not os.path.exists(ARQ_TRANSP):
        st.error(f"Arquivo não encontrado: {ARQ_TRANSP}")
        st.stop()
        
    @st.cache_data(ttl=300)
    def carregar_base_transportadora(path):
        return pd.read_excel(path)

    try:
        df = carregar_base_transportadora(ARQ_TRANSP)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        st.stop()

    if df.empty:
        st.warning("Sem dados.")
        st.stop()

    # ==============================
    # TRATAMENTO
    # ==============================

    df["DataExpedicao"] = pd.to_datetime(df["DataExpedicao"], errors="coerce")
    df["DataFinal"] = pd.to_datetime(df["DataFinal"], errors="coerce")
    
    df["DataFinal"] = pd.to_datetime(df["DataFinal"], errors="coerce")
    df["DataPrevista"] = pd.to_datetime(df["DataPrevista"], errors="coerce")

    df = df.dropna(subset=["Transportadora", "DataFinal", "DataPrevista"])

    df["DentroPrazo"] = df["DataFinal"].dt.date <= df["DataPrevista"].dt.date
    
    df["Mes"] = df["DataFinal"].dt.to_period("M").astype(str)

    mensal = (
        df.groupby(["Mes", "Transportadora"])
        .agg(
            Total=("DentroPrazo", "count"),
            Dentro=("DentroPrazo", "sum")
        )
        .reset_index()
    )

    mensal["% Dentro"] = (
        mensal["Dentro"] / mensal["Total"]
    ) * 100
    

    # ==============================
    # REGRA PRAZO
    # ==============================

    resumo = (
        df.groupby("Transportadora")
        .agg(
            Total=("DentroPrazo", "count"),
            Dentro=("DentroPrazo", "sum")
        )
        .reset_index()
    )

    resumo["Fora"] = resumo["Total"] - resumo["Dentro"]

    resumo["% Dentro"] = (resumo["Dentro"] / resumo["Total"]) * 100
    resumo["% Fora"] = (resumo["Fora"] / resumo["Total"]) * 100

    resumo = resumo.sort_values("% Fora", ascending=False)
    
    st.markdown("## 📈 Evolução Mensal")

    transportadoras_graf = sorted(
        df["Transportadora"].unique()
    )

    transp_sel = st.selectbox(
        "Transportadora",
        transportadoras_graf
    )

    df_graf = mensal[
        mensal["Transportadora"] == transp_sel
    ].copy()

    df_graf = df_graf.sort_values("Mes")

    df_graf["Mes"] = pd.to_datetime(df_graf["Mes"])

    df_graf = df_graf.sort_values("Mes")

    fig = px.line(
        df_graf,
        x="Mes",
        y="% Dentro",
        markers=True,
        title=f"Desempenho Mensal - {transp_sel}"
    )

    fig.update_layout(
        xaxis_title="Mês",
        yaxis_title="% Dentro do Prazo",
        height=450
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

    # ==============================
    # VISUAL (MESMO PADRÃO)
    # ==============================

    for _, row in resumo.iterrows():
    
            transportadora = row["Transportadora"]

            df_transp = df[
                df["Transportadora"] == transportadora
            ].copy()

            df_fora = df_transp[
                df_transp["DentroPrazo"] == False
            ].copy()

            st.write(f"### {transportadora}")

            st.write(
                f"Dentro: {int(row['Dentro'])} ({row['% Dentro']:.1f}%) | "
                f"Fora: {int(row['Fora'])} ({row['% Fora']:.1f}%)"
            )
            
            if not df_fora.empty:

                df_export = df_fora.copy()

                # ==============================
                # STATUS FINAL
                # ==============================

                mapa_status = {
                    5: "ENTREGUE",
                    15: "ENTREGA SAT",
                    25: "ENTREGUE SUJEITA REABERTURA",
                    35: "ENTREGA SAT PRODUTO NAO POSTADO",
                    105: "DEVOLVIDA",
                    107: "EM TRANSFERENCIA",
                    118: "DEVOLUCAO EM ROTA",
                    119: "TRANSFERENCIA PARA DEVOLUCAO",
                    182: "EM PROCESSO DE DEVOLUCAO",
                    183: "EM ROTA DE DEVOLUCAO",
                    185: "REENTREGA",
                    573: "CANCELADO",
                    611: "CANCELADO FRAUDE",
                    847: "AGUARDANDO RETORNO CLIENTE DEVOLVIDO",
                    957: "EXTRAVIO",
                    977: "DEVOLUCAO ENV SAP",
                    978: "DEVOLUCAO RET SAP",
                    979: "DEVOLUCAO NF",
                    983: "CANCELADO ENV SAP",
                    986: "CANCELADO APOS ANALISE FRAUDE",
                    987: "PEDIDO ENCERRADO"
                }

                if "idNumStatus" in df_export.columns:
                    df_export["StatusFinal"] = (
                        df_export["idNumStatus"]
                        .map(mapa_status)
                        .fillna(df_export["idNumStatus"].astype(str))
                    )
                else:
                    df_export["StatusFinal"] = ""

                # ==============================
                # RENOMEAR DATAS
                # ==============================

                if "DataPrevista" in df_export.columns:
                    df_export = df_export.rename(columns={
                        "DataPrevista": "DataEntregaPrevistaCotacao"
                    })

                if "DataFinal" in df_export.columns:
                    df_export = df_export.rename(columns={
                        "DataFinal": "DataStatusFinal"
                    })

                # ==============================
                # COLUNAS FINAIS
                # ==============================

                colunas_exportar = [
                    c for c in [
                        "PedidoFormatado",
                        "Campanha",
                        "Transportadora",
                        "DataEntregaPrevistaCotacao",
                        "DataStatusFinal",
                        "StatusFinal"
                    ]
                    if c in df_export.columns
                ]

                df_export = df_export[colunas_exportar].copy()

                # ==============================
                # FORMATAÇÃO DATA
                # ==============================

                for col in [
                    "DataEntregaPrevistaCotacao",
                    "DataStatusFinal"
                ]:

                    if col in df_export.columns:

                        df_export[col] = pd.to_datetime(
                            df_export[col],
                            errors="coerce"
                        ).dt.strftime("%d/%m/%Y")

                # ==============================
                # EXPORTAÇÃO
                # ==============================

                buffer = BytesIO()

                df_export.to_excel(
                    buffer,
                    index=False
                )

                buffer.seek(0)

                st.download_button(
                    label=f"⬇️ Baixar atrasados - {transportadora}",
                    data=buffer,
                    file_name=f"atrasados_{transportadora}.xlsx",
                    key=f"download_atrasados_{transportadora}"
                )