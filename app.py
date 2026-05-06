import streamlit as st
import pandas as pd
import os
from io import BytesIO
import matplotlib.pyplot as plt
import re
import base64
import matplotlib.dates as mdates

# ==============================
# CONFIGS
# ==============================
BASE_DIR = os.path.dirname(__file__)
PASTA_DATA = os.path.join(BASE_DIR, "data")
PASTA_HIST = os.path.join(PASTA_DATA, "historico")
PASTA_MENSAL = os.path.join(PASTA_DATA, "mensal_status")
os.makedirs(PASTA_MENSAL, exist_ok=True)
ARQ_ATUAL = os.path.join(PASTA_DATA, "Monitor_Pedidos_Processado.xlsx")
ARQ_DEVOLUCAO = os.path.join(PASTA_DATA, "Base_Streamlit_Devolucao.xlsx")
LOGO_PATH = os.path.join(PASTA_DATA, "logo_bravium.png")

TAMANHO_LOTE = 300

STATUS_DIARIOS = [
    "TSP - Pendente Transportes - Dados do Recebedor Solicitado",
    "TSP - Item Faltante",
    "TSP - Pendente Transportes - Aguardando Acareação",
    "TSP - Pendente Transportes - Acareação Solicitada",
    "TSP - Item Faltante Solicitado",
    "TSP - Aguardando Dados do Recebedor"
]

STATUS_Manuais = [
    "TSP - Críticos",
    "TSP - Reentrega",
    "TSP - Reentregar/Endereço correto",
    "Aguardando tratativa transportadora",
    "TSP - Aguardando Acareação",
    "TSP - Pendente Transportes - Acareação Solicitada",
    "TSP - Aguardando Dados do Recebedor",
    "TSP - Pendente Transportes - Dados do Recebedor Solicitado",
    "TSP - Item Faltante",
    "TSP - Item Faltante Solicitado",
    "TSP - Aguardando Avaliação de Problema de Coleta",
    "TSP - Coleta Realizada",
    "TSP - Aguardando Coleta",
    "TSP - Coleta Agendada",
    "TSP - Aguardando Envio de Guia de Retenção ao Fiscal"
]

st.set_page_config(
    page_title="BI Executivo - Monitor",
    layout="wide",
    page_icon="📊"
)

# ==============================
# ESTILO
# ==============================

  
st.markdown("""
<style>

    .titulo-painel {
        background: linear-gradient(90deg,#0f2a44,#1f4e79);
        color: white;
        font-size: 20px;
        font-weight: 700;
        text-align: center;
        padding: 12px;
        border-radius: 12px;
        margin-bottom: 20px;
    }

    .separador-colunas {
        border-right: 6px solid #1f4e79;
        padding-right: 25px;
    }

    </style>
<style>
.stApp { background-color: #f4f6f9; }

.header-box { 
background: linear-gradient(90deg, #0f2a44, #1f4e79);
padding: 18px 24px; 
border-radius: 14px; 
display: flex;
align-items: center; 
gap: 20px; 
margin-bottom: 20px; 
}

.header-title { 
color: white; 
font-size: 26px; 
font-weight: 700; 
margin: 0; 
}

.header-sub { 
color: white; 
opacity: 0.85; 
margin: 0; 
font-size: 14px; 
}

.header-box img {
max-width: 220px !important; 
}

.data-title { 
font-size: 20px; 
font-weight: 700; 
color: #0f2a44;
margin-top: 10px; 
margin-bottom: 10px; 
}

.metric-small { 
font-size: 16px !important; 
font-weight: 600; 
color: #0f2a44; 
}

/* BOTÕES */
button {
background: linear-gradient(90deg, #0f2a44, #1f4e79) !important;
color: white !important;
border-radius: 10px !important;
font-weight: 700 !important;
border: none !important;
}

button:hover {
background: linear-gradient(90deg, #1f4e79, #0f2a44) !important;
color: white !important;
}

/* ============================= */
/* BOTÃO SIDEBAR (COR + ÍCONE) */
/* ============================= */

/* botão */
button[data-testid="collapsedControl"] {
    background-color: #0f2a44 !important;
    border-radius: 8px !important;
}

/* ícone (seta) */
button[data-testid="collapsedControl"] svg {
    fill: white !important;
    color: white !important;
}

/* hover */
button[data-testid="collapsedControl"]:hover {
    background-color: #1f4e79 !important;
}

</style>
""", unsafe_allow_html=True)

# ==============================
# MENU LATERAL
# ==============================

pagina = st.sidebar.selectbox(
    "Painel",
    [
        "Monitor de Pedidos",
        "Indicador de Devolução"
    ]
)

# ==============================
# HEADER
# ==============================
logo_html = ""
logo_base64 = ""
if os.path.exists(LOGO_PATH):
    with open(LOGO_PATH, "rb") as f:
        logo_base64 = base64.b64encode(f.read()).decode()
    logo_html = f'<img src="data:image/png;base64,{logo_base64}" width="120">'
    
st.markdown(
    f"""
    <style>
    .stApp {{
        background-color: #f4f6f9;
    }}

    .stApp::before {{
        content: "";
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background-image: url("data:image/png;base64,{logo_base64}");
        background-repeat: no-repeat;
        background-position: center;
        background-size: 700px;
        opacity: 0.03;
        pointer-events: none;
        z-index: 0;
    }}
    </style>
    """,
    unsafe_allow_html=True
)

if pagina == "Monitor de Pedidos":
    titulo = "Monitor de Pedidos — BI Executivo"
    subtitulo = "Análise de Risco Logístico • Transportadora • Status • Região • Cliente"

elif pagina == "Indicador de Devolução":
    titulo = "Painel de Devolução"
    subtitulo = "Devolução • Extravio • Avaria • Indicadores Logísticos"

st.markdown(f"""
<div class="header-box">
    {logo_html}
    <div>
        <p class="header-title">{titulo}</p>
        <p class="header-sub">
        {subtitulo}
        </p>
    </div>
</div>
""", unsafe_allow_html=True)

# ==============================
# FUNÇÕES
# ==============================

def normalizar_pedido(col):
    return (
        col.astype(str)
        .str.upper()
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )
    
def perc_transportes(x, venda_mes):
    if venda_mes == 0:
        return "0%"
    return f"{(x / venda_mes)*100:.2f}%".replace(".", ",")

def carregar_base_devolucao():
    return pd.read_excel(
        ARQ_DEVOLUCAO,
        sheet_name=[
            "vendas_mes",
            "vendas_mes_pedido",
            "vendas_transportadora",
            "potencial_triplo",
            "devolucao_processo",
            "retornando_transportes",
            "devolucao_atrasada",
            "nfd_mes",
            "nfd_coleta",
            "base",
            "retornando_detalhado",
            "dev_atrasada_detalhado"
        ]
    )
@st.cache_data
def ler_base(path):
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        df = pd.read_excel(path)
    except Exception:
        return pd.DataFrame()
    if "PedidoFormatado" in df.columns:
        df["PedidoFormatado"] = normalizar_pedido(df["PedidoFormatado"])
    return df

def listar_dias():
    if not os.path.exists(PASTA_HIST):
        return []
    arquivos = os.listdir(PASTA_HIST)
    datas = set()
    for a in arquivos:
        m = re.match(r"(\d{2}-\d{2}-\d{4})_manha\.xlsx$", a)
        if m:
            datas.add(m.group(1))
    return sorted(datas, key=lambda x: pd.to_datetime(x, format="%d-%m-%Y"))

def caminho(dia):
    return os.path.join(PASTA_HIST, f"{dia}_manha.xlsx")

def pizza(tratados, restantes, titulo):
    fig, ax = plt.subplots(figsize=(2.3, 2.3))
    total = tratados + restantes
    if total == 0:
        ax.text(0.5, 0.5, "0", ha="center", va="center")
    else:
        ax.pie([tratados, restantes], autopct="%1.0f%%", startangle=90)
    ax.set_title(titulo, fontsize=10)
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()
    
if pagina == "Monitor de Pedidos":

    # ==============================
    # DOWNLOAD POR CARTEIRA (COM ABA EXTRA)
    # ==============================
    st.markdown("### 📥 Exportação por Carteira (300 em 300)")

    df_atual = ler_base(ARQ_ATUAL)

    if "offsets_carteira" not in st.session_state:
        st.session_state["offsets_carteira"] = {}

    if not df_atual.empty and "Carteira" in df_atual.columns:

        if "Ranking" in df_atual.columns:
            df_atual = df_atual.sort_values("Ranking").reset_index(drop=True)

        carteiras = sorted(df_atual["Carteira"].dropna().unique())
        carteiras = [c for c in carteiras if c != "Augusto"]

        if "Igor" in df_atual["Carteira"].values and "Igor" not in carteiras:
            carteiras.append("Igor")

        for carteira in carteiras:

            df_carteira = df_atual[
                df_atual["Carteira"] == carteira
            ].copy()

            df_fora_prazo = df_carteira[
                (df_carteira["Nivel_Cliente"] != "Dentro") |
                (df_carteira["Nivel_Transportadora"] != "Dentro") |
                (df_carteira["Nivel_Status_Especifico"] != "Dentro") |
                (df_carteira["Nivel_Regiao"] != "Dentro")
            ].reset_index(drop=True)

            df_dentro_prazo = df_carteira[
                (df_carteira["Nivel_Cliente"] == "Dentro") &
                (df_carteira["Nivel_Transportadora"] == "Dentro") &
                (df_carteira["Nivel_Status_Especifico"] == "Dentro") &
                (df_carteira["Nivel_Regiao"] == "Dentro")
            ]

            total = len(df_fora_prazo)
            total_dentro = len(df_dentro_prazo)
            offset = st.session_state["offsets_carteira"].get(carteira, 0)

            inicio = offset
            fim = min(offset + TAMANHO_LOTE, total)

            lote = df_fora_prazo.iloc[inicio:fim]

            if not lote.empty:
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

                col1, col2 = st.columns([4, 2])

                perc_fora = (total / (total + total_dentro)) * 100 if (total + total_dentro) > 0 else 0
                barra = int(perc_fora // 5)

                st.markdown(f"""
                <div style="
                    background: white;
                    padding: 18px;
                    border-radius: 16px;
                    box-shadow: 0 3px 10px rgba(0,0,0,0.08);
                    margin-bottom: 12px;
                ">

                    <div style="font-size:18px;font-weight:700;color:#0f2a44;">
                        {carteira}
                    </div>

                    <div style="font-size:13px;color:#6b7280;margin-bottom:10px;">
                        Lote {inicio+1} → {fim}
                    </div>

                    <div style="display:flex;justify-content:space-between;">
                        <div style="font-size:26px;font-weight:800;color:#b91c1c;">
                            🚨 {total}
                        </div>
                        <div style="font-size:16px;font-weight:700;color:#b91c1c;">
                            {perc_fora:.1f}%
                        </div>
                    </div>

                    <div style="font-size:13px;color:#6b7280;">
                        Fora do prazo
                    </div>

                    <div style="margin-top:8px;font-family:monospace;color:#b91c1c;">
                        {"█"*barra}{"░"*(20-barra)}
                    </div>

                    <div style="font-size:13px;font-weight:600;color:#047857;margin-top:8px;">
                        ✅ {total_dentro} dentro do prazo
                    </div>

                </div>
                """, unsafe_allow_html=True)

                with col2:
                    if st.download_button(
                        label=f"⬇️ Baixar {carteira}",
                        data=buffer,
                        file_name=f"{carteira}_{inicio+1}_a_{fim}.xlsx",
                        key=f"dl_{carteira}_{offset}"
                    ):
                        st.session_state["offsets_carteira"][carteira] = fim
                        
    # ==============================
    # 🧑 CARTEIRA AUGUSTO (NOVA)
    # ==============================

    df_augusto = df_atual[
        df_atual["Carteira"] == "Augusto"
    ]

    total_augusto = len(df_augusto)

    col1, col2 = st.columns([4,2])

    with col1:
        st.write(f"**Augusto** — 1 até {total_augusto} de {total_augusto}")

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
                (df_atual["Status"] == "TSP - Aguardando Expedição") &
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

        qtd = df_temp[df_temp["Status"].isin(STATUS_DIARIOS)].shape[0]

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

                fig, ax = plt.subplots(figsize=(3.2, 1.2))

                ax.plot(df_mes["Data"], df_mes["Qtd"])

                ax.set_xticks(df_mes["Data"])
                ax.set_xticklabels(df_mes["Data"].dt.day, fontsize=7)

                ax.set_xlabel("Dia", fontsize=7)
                ax.set_ylabel("Qtde", fontsize=7)

                nome_mes = df_mes["Data"].dt.strftime("%B").iloc[0].capitalize()
                ano = df_mes["Data"].dt.year.iloc[0]

                ax.set_title(f"{nome_mes}/{ano}", fontsize=9)

                st.pyplot(fig)
                plt.close(fig)
                
    # ==============================
    # 📌 STATUS MANUAIS (3 CURVAS)
    # ==============================

    st.markdown("### 📌 Status Manuais")

    ARQ_CACHE = os.path.join(PASTA_MENSAL, "manuais_cache.xlsx")

    if os.path.exists(ARQ_CACHE):

        df_graf = pd.read_excel(ARQ_CACHE)

        if not df_graf.empty:

            df_graf["Data"] = pd.to_datetime(df_graf["Data"])
            df_graf = df_graf.sort_values("Data")

            fig, ax = plt.subplots(figsize=(8, 3))

            ax.plot(df_graf["Data"], df_graf["Total"], label="Total")
            ax.plot(df_graf["Data"], df_graf["Entraram"], label="Entraram")
            ax.plot(df_graf["Data"], df_graf["Sairam"], label="Saíram")

            ax.set_title("Status Manuais")
            ax.set_xlabel("Dia")
            ax.set_ylabel("Quantidade")
            ax.legend()

            ax.xaxis.set_major_formatter(mdates.DateFormatter('%d'))

            st.pyplot(fig, use_container_width=True)
            plt.close(fig)

    else:
        st.info("Cache de status manuais ainda não foi gerado.")
        
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
            df_hist_ant = historico[dia_ant]
            
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

                    atual = df_hist_atual[df_hist_atual["Transportadora_Triplo"] == "X"]
                    ant = df_hist_ant[df_hist_ant["Transportadora_Triplo"] == "X"]

                    tratados = ant[~ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
                    restantes = ant[ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
                    entrou = atual[~atual["PedidoFormatado"].isin(ant["PedidoFormatado"])]

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

                    atual = df_hist_atual[df_hist_atual["Status_Dobro"] == "X"]
                    ant = df_hist_ant[df_hist_ant["Status_Dobro"] == "X"]

                    tratados = ant[~ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
                    restantes = ant[ant["PedidoFormatado"].isin(atual["PedidoFormatado"])]
                    entrou = atual[~atual["PedidoFormatado"].isin(ant["PedidoFormatado"])]

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

                    atual = df_hist_atual[df_hist_atual["Regiao_Dobro"] == "X"]
                    ant = df_hist_ant[df_hist_ant["Regiao_Dobro"] == "X"]

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

# ==============================
# INDICADOR DE DEVOLUÇÃO
# ==============================

elif pagina == "Indicador de Devolução":

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

            fig, ax = plt.subplots(figsize=(20,8))

            ax.plot(
                graf_vendas["Mes"],
                graf_vendas["ValorVenda"],
                marker="o"
            )

            ax.set_title("Venda mensal - Transportadoras")
            ax.set_xlabel("Mês")
            ax.set_ylabel("Valor")

            ax.xaxis.set_major_formatter(mdates.DateFormatter('%b/%Y'))

            plt.xticks(rotation=45)

            st.pyplot(fig, use_container_width=True)
            plt.close(fig)

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

            fig, ax = plt.subplots(figsize=(8,4))

            ax.plot(
                graf_indice["Mes"],
                graf_indice["IndiceAtual"] * 100,
                label="Índice atual"
            )

            ax.plot(
                graf_indice["Mes"],
                graf_indice["IndicePotencial"] * 100,
                label="Índice com potencial"
            )

            ax.set_title("Impacto Potencial Triplo no Índice de Devolução")
            ax.set_ylabel("Índice %")

            ax.legend()

            ax.xaxis.set_major_formatter(mdates.DateFormatter('%b/%Y'))

            plt.xticks(rotation=45)

            st.pyplot(fig, use_container_width=True)

            plt.close(fig)
            
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
            unsafe_allow_html=True
        )
        
        # ==============================
        # GRÁFICO VENDAS POR MÊS - BRAVIUM
        # ==============================
        
        graf_bravium = bases["vendas_mes_pedido"].copy()

        graf_bravium["Mes_Pedido"] = pd.to_datetime(
            graf_bravium["Mes_Pedido"].astype(str)
        )

        graf_bravium = graf_bravium.sort_values("Mes_Pedido")

        fig, ax = plt.subplots(figsize=(20,8))

        ax.plot(
            graf_bravium["Mes_Pedido"],
            graf_bravium["ValorVenda"],
            marker="o"
        )

        ax.set_title("Venda mensal - Bravium")
        ax.set_xlabel("Mês")
        ax.set_ylabel("Valor")

        ax.xaxis.set_major_formatter(mdates.DateFormatter('%b/%Y'))

        plt.xticks(rotation=45)

        st.pyplot(fig, use_container_width=True)
        plt.close(fig)

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
                    