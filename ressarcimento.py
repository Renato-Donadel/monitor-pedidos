"""
Central de Ressarcimento — controle de pagamentos recebidos das transportadoras referentes a
pedidos extraviados que já geraram NFD (bucket "Devolvido" de extravio_detalhado: NFD emitida
depois do extravio, sem entrega depois — ver Indicador_de_Devolucao.py pro detalhe da regra).

PERSISTÊNCIA — por que Google Sheets e não o próprio repositório git (pedido do usuário,
09/2026): o site roda no Streamlit Community Cloud, sem acesso ao drive Z: da empresa, e
qualquer escrita no disco local do site se perde a cada reinício/nova publicação. Gravar direto
no repositório git exigiria guardar uma credencial do GitHub com permissão de ESCREVER no
repositório — como o repositório é PÚBLICO, se essa credencial vazasse daria pra reescrever
QUALQUER arquivo dele, código incluso, não só a lista de pagamentos. Uma Service Account do
Google com acesso só a ESSA planilha específica reduz o estrago máximo de um vazamento a
"alguém mexe na lista de pagamentos" — nunca no código do site.

O `Indicador_de_Devolucao.py` (que roda localmente, com acesso ao Z:) também lê essa mesma
planilha e salva uma cópia na pasta do projeto — ver seção "CONTROLE DE RESSARCIMENTO (GOOGLE
SHEETS)" lá.

CONFIGURAÇÃO (Streamlit Cloud → app → Settings → Secrets — colar como texto TOML, com os
valores copiados do arquivo .json baixado no Google Cloud Console):

    ressarcimento_sheet_id = "<ID da planilha, parte da URL entre /d/ e /edit>"

    [gcp_service_account]
    type = "service_account"
    project_id = "..."
    private_key_id = "..."
    private_key = "..."
    client_email = "..."
    client_id = "..."
    auth_uri = "https://accounts.google.com/o/oauth2/auth"
    token_uri = "https://oauth2.googleapis.com/token"
    auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
    client_x509_cert_url = "..."
"""

import streamlit as st
import pandas as pd

COLUNAS_PLANILHA = ["Pedido Formatado", "Valor Pago"]


def _acesso_configurado():
    return "gcp_service_account" in st.secrets and "ressarcimento_sheet_id" in st.secrets


def _abrir_planilha():
    import gspread  # import local: só é uma dependência real quando o acesso existir
    from google.oauth2.service_account import Credentials

    escopos = ["https://www.googleapis.com/auth/spreadsheets"]
    credenciais = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=escopos
    )
    cliente = gspread.authorize(credenciais)
    return cliente.open_by_key(st.secrets["ressarcimento_sheet_id"]).sheet1


@st.cache_data(ttl=30)
def carregar_pagamentos():
    """DataFrame com PedidoFormatado/ValorPago — tudo que já foi importado até agora. Cache
    curto (30s) só pra não bater na API do Google a cada interação da página."""
    if not _acesso_configurado():
        return pd.DataFrame(columns=["PedidoFormatado", "ValorPago"])

    try:
        aba = _abrir_planilha()
        # UNFORMATTED_VALUE (não o padrão FORMATTED_VALUE do gspread): achado real (09/2026,
        # testado contra a planilha de verdade) — a planilha usa formato BR (vírgula decimal),
        # então o valor "formatado" de 4837.8 vem como a STRING "4837,8"; convertendo isso pro
        # Python, a vírgula era descartada e virava 48378 (10x maior). UNFORMATTED_VALUE traz o
        # número de verdade (4837.8), sem depender de nenhuma conversão de texto/locale.
        #
        # numericise_ignore=["all"] — achado real separado: o PRÓPRIO gspread (não o Google
        # Sheets) tenta "adivinhar" e converter qualquer texto só de dígitos pra número dentro
        # de get_all_records(), o que reintroduz o mesmo problema pro Pedido Formatado (perde
        # zero à esquerda, ex.: "00760000714423" virava 760000714423) mesmo com a célula
        # armazenada corretamente como texto. Desligado pra tudo; ValorPago é convertido pra
        # número manualmente logo abaixo com pd.to_numeric.
        registros = aba.get_all_records(
            value_render_option="UNFORMATTED_VALUE",
            numericise_ignore=["all"],
        )
    except Exception as exc:
        st.warning(f"Não consegui ler a planilha de ressarcimento: {exc}")
        return pd.DataFrame(columns=["PedidoFormatado", "ValorPago"])

    if not registros:
        return pd.DataFrame(columns=["PedidoFormatado", "ValorPago"])

    df = pd.DataFrame(registros).rename(
        columns={"Pedido Formatado": "PedidoFormatado", "Valor Pago": "ValorPago"}
    )
    df["PedidoFormatado"] = df["PedidoFormatado"].astype(str).str.strip().str.upper()
    df["ValorPago"] = pd.to_numeric(df["ValorPago"], errors="coerce").fillna(0)
    return df[["PedidoFormatado", "ValorPago"]]


def _normalizar_pedido(serie: pd.Series) -> pd.Series:
    """Pedido Formatado em texto, maiúsculo, sem '.0' de número lido do Excel. Pedido só de
    dígitos tem SEMPRE 14 caracteres (confirmado nos 16.305 pedidos de extravio_detalhado; os
    com letra de reenvio têm 16) — se vier mais curto, o Excel do usuário comeu zeros à
    esquerda (célula numérica), então restauramos com zfill(14)."""
    s = serie.astype(str).str.strip().str.upper().str.replace(r"\.0$", "", regex=True)
    so_digitos = s.str.fullmatch(r"\d+")
    s = s.where(~so_digitos, s.str.zfill(14))
    return s


def _parse_valor(v):
    """Converte 'Valor Pago' em float aceitando número do Excel (4837.8) e texto BR ('R$
    4.837,80', '4837,8') ou US ('4837.80'). Devolve NaN se vazio/inválido — quem chama decide
    (nunca vira 0 em silêncio)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return float("nan")
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace("R$", "").replace(" ", "")
    if not t:
        return float("nan")
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return float("nan")


def registrar_pagamentos(df_novo: pd.DataFrame) -> tuple[int, int]:
    """Recebe o Excel/CSV importado pelo usuário (colunas 'Pedido Formatado'/'Valor Pago') e
    ACRESCENTA na planilha só os pedidos que ainda não estavam lá — evita duplicar valor se o
    mesmo arquivo for importado de novo por engano. Devolve (qtd_inseridos, qtd_ja_existentes).
    Levanta ValueError se as colunas não baterem, RuntimeError se o acesso não estiver
    configurado."""
    if not _acesso_configurado():
        raise RuntimeError(
            "Acesso à planilha de ressarcimento não configurado "
            "(Settings > Secrets do app no Streamlit Cloud)."
        )

    df_novo = df_novo.rename(columns=lambda c: str(c).strip())

    faltando = set(COLUNAS_PLANILHA) - set(df_novo.columns)
    if faltando:
        raise ValueError(
            f"A planilha importada precisa ter as colunas {COLUNAS_PLANILHA} — "
            f"faltando: {sorted(faltando)}."
        )

    df_novo = df_novo[COLUNAS_PLANILHA].copy()
    df_novo = df_novo.dropna(subset=["Pedido Formatado"])
    df_novo["Pedido Formatado"] = _normalizar_pedido(df_novo["Pedido Formatado"])
    df_novo = df_novo[df_novo["Pedido Formatado"] != ""]

    valores = df_novo["Valor Pago"].map(_parse_valor)
    if valores.isna().any():
        raise ValueError(
            f"{int(valores.isna().sum())} linha(s) com 'Valor Pago' vazio ou inválido "
            f"(ex.: {df_novo.loc[valores.isna(), 'Valor Pago'].head(3).tolist()}). "
            "Nada foi importado — corrija a planilha e envie de novo."
        )
    df_novo["Valor Pago"] = valores

    existentes = set(carregar_pagamentos()["PedidoFormatado"])
    novos = df_novo[~df_novo["Pedido Formatado"].isin(existentes)]
    ja_existiam = len(df_novo) - len(novos)

    if not novos.empty:
        aba = _abrir_planilha()
        # RAW (não USER_ENTERED): achado real (09/2026, testado contra a planilha de verdade)
        # — USER_ENTERED deixa o Google Sheets "interpretar" o valor como se alguém tivesse
        # digitado, e isso quebrou os dois lados: um Pedido Formatado começando com zero
        # ("00760000714423") virou número e perdeu os zeros à esquerda ("760000714423"), e um
        # Valor Pago com decimal ("4837.80") teve o "." tratado como separador de milhar da
        # planilha (locale) e virou "48378". RAW grava o valor exatamente como foi passado, sem
        # nenhuma dessas conversões.
        aba.append_rows(novos.values.tolist(), value_input_option="RAW")
        carregar_pagamentos.clear()

    return len(novos), ja_existiam
