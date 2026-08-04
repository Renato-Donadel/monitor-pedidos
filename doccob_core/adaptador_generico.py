"""
Adaptador GENERICO: para planilhas de transportadoras que ainda nao tem um
adaptador dedicado (ver csv_transportadora.py para o formato "Composicao dos
Fretes" da Carvalima/OTC/FAV, e xlsx_lsp.py para o formato "Remessas" da
LSP/JADLOG).

Estrategia: como o layout de cada transportadora varia (~20 transportadoras,
cada uma com sua planilha), em vez de mapear coluna por coluna manualmente,
este modulo tenta descobrir automaticamente:

  1. A coluna com a CHAVE DE ACESSO do CT-e (44 digitos) - isso e' universal
     (todo CT-e tem, nao importa a transportadora) e e' o dado mais critico
     (foi a causa do bug de reconhecimento corrigido no adaptador da
     Carvalima). Uma vez achada essa coluna, CNPJ do emissor, numero e serie
     do CT-e sao extraidos dela com a MESMA tecnica ja validada.
  2. As demais colunas (valor, datas, UFs, CNPJ remetente/destinatario,
     numero da fatura) por casamento de palavras-chave no nome da coluna.

Isso e' um "melhor esforco" - NAO tem a mesma confianca dos adaptadores
dedicados (Carvalima e LSP), que foram validados byte a byte contra
arquivos reais. Todo campo que nao for encontrado com confianca gera um
aviso explicito - nunca falha silenciosamente.

Quando uma transportadora nova aparecer com frequencia, o ideal e' criar um
adaptador dedicado pra ela (como os outros dois), usando este modulo so
como ponto de partida/fallback.
"""

import re
import unicodedata
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation

import pandas as pd

from .doccob import CTeCobranca, DocumentoCobranca, LoteCobranca


def _normalizar(txt: str) -> str:
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode("ascii")
    return txt.upper().strip()


def _cnpj_da_chave(chave) -> str | None:
    chave = "".join(ch for ch in str(chave) if ch.isdigit())
    if len(chave) != 44:
        return None
    return chave[6:20]


def _serie_da_chave(chave) -> str | None:
    chave = "".join(ch for ch in str(chave) if ch.isdigit())
    if len(chave) != 44:
        return None
    return chave[22:25]


def _numero_cte_da_chave(chave) -> str | None:
    chave = "".join(ch for ch in str(chave) if ch.isdigit())
    if len(chave) != 44:
        return None
    return str(int(chave[25:34]))


def encontrar_coluna_chave_cte(df: pd.DataFrame) -> str | None:
    """Acha a coluna cujos valores parecem chave de acesso de CT-e (44
    digitos), testando uma amostra de ate 30 linhas nao vazias."""
    melhor_col, melhor_taxa = None, 0.0
    for col in df.columns:
        amostra = df[col].dropna().astype(str).head(30)
        if amostra.empty:
            continue
        digitos = amostra.map(lambda v: "".join(ch for ch in v if ch.isdigit()))
        taxa = (digitos.str.len() == 44).mean()
        if taxa > melhor_taxa:
            melhor_col, melhor_taxa = col, taxa
    return melhor_col if melhor_taxa >= 0.8 else None


# Palavras-chave (ja normalizadas: maiusculas, sem acento) para achar cada
# conceito pelo nome da coluna. Ordem importa - primeiro match ganha.
PALAVRAS_CHAVE = {
    "valor_frete": ["VALOR A RECEBER", "VAL RECEBER", "VALOR DO FRETE", "VALOR FRETE", "FRETE PESO", "VALOR"],
    "dt_emissao": ["DATA EMISSAO", "DT EMISSAO", "DATA"],
    "cnpj_remetente": ["CNPJ REMETENTE", "CPF/CNPJ REMETENTE"],
    "cnpj_destinatario": ["CNPJ DESTINATARIO", "CPF/CNPJ DESTINATARIO"],
    "uf_origem": ["UF INICIO", "UF REMETENTE", "UF ORIGEM", "UF EXPEDIDOR"],
    "uf_destino": ["UF FIM", "UF DESTINATARIO", "UF DESTINO", "UF ENTREGA", "ESTADO DEST"],
    "numero_fatura": ["NUMERO DA FATURA", "NUMERO FATURA", "NR FATURA", "FATURA"],
    "nome_tomador": ["NOME TOMADOR"],
    "nome_transportadora": ["EMISSOR", "TRANSPORTADORA", "NOME TRANSPORTADORA"],
    "tipo_documento": ["TIPO DOCUMENTO", "TIPO"],
    "numero_pedido": ["PEDIDO", "NUMERO DO PEDIDO"],
}


def _achar_coluna(df: pd.DataFrame, conceito: str) -> str | None:
    candidatos = PALAVRAS_CHAVE[conceito]
    colunas_norm = {col: _normalizar(col) for col in df.columns}
    for alvo in candidatos:
        for col, norm in colunas_norm.items():
            if norm == alvo:
                return col
    for alvo in candidatos:
        for col, norm in colunas_norm.items():
            if alvo in norm:
                return col
    return None


def _parse_valor(v) -> Decimal | None:
    if pd.isna(v):
        return None
    if isinstance(v, (int, float)):
        return Decimal(str(v))
    txt = str(v).strip().replace(".", "").replace(",", ".")
    try:
        return Decimal(txt)
    except InvalidOperation:
        return None


def _parse_data(v) -> date | None:
    if pd.isna(v):
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.date()
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(v).strip(), fmt).date()
        except ValueError:
            continue
    return None


def montar_lotes(
    caminho: str,
    sheet_name=0,
    header: int = 0,
    e_devolucao: bool = False,
    dt_emissao_fatura: date = None,
    dt_vencimento_fatura: date = None,
    tomador_nome: str = "BRAVIUM S.A",
    transportadora_nome: str = None,
    transportadora_cnpj: str = None,
) -> list[tuple[LoteCobranca, list[str]]]:
    """Tenta montar uma LoteCobranca por numero de fatura encontrado, a
    partir de QUALQUER planilha (csv ou xlsx), descobrindo as colunas por
    palavra-chave. Devolve lista de (lote, avisos) - SEMPRE conferir os
    avisos antes de considerar o resultado confiavel.
    """
    avisos_gerais = []

    if caminho.lower().endswith(".csv"):
        df = pd.read_csv(caminho, sep=None, engine="python", encoding="latin-1", dtype=str)
    else:
        df = pd.read_excel(caminho, sheet_name=sheet_name, header=header)

    col_chave = encontrar_coluna_chave_cte(df)
    if not col_chave:
        raise ValueError(
            "Nao encontrei nenhuma coluna com chave de acesso de CT-e (44 digitos) "
            "nesta planilha - preciso dela para saber CNPJ/numero/serie do CT-e com confianca."
        )

    col_valor = _achar_coluna(df, "valor_frete")
    col_dt_emissao = _achar_coluna(df, "dt_emissao")
    col_cnpj_rem = _achar_coluna(df, "cnpj_remetente")
    col_cnpj_dest = _achar_coluna(df, "cnpj_destinatario")
    col_uf_origem = _achar_coluna(df, "uf_origem")
    col_uf_destino = _achar_coluna(df, "uf_destino")
    col_fatura = _achar_coluna(df, "numero_fatura")
    col_nome_tomador = _achar_coluna(df, "nome_tomador")
    col_nome_transp = _achar_coluna(df, "nome_transportadora")
    col_tipo_doc = _achar_coluna(df, "tipo_documento")
    col_pedido = _achar_coluna(df, "numero_pedido")

    for nome_conceito, col in [
        ("valor do frete", col_valor), ("data de emissao", col_dt_emissao),
        ("CNPJ remetente", col_cnpj_rem), ("CNPJ destinatario", col_cnpj_dest),
        ("UF origem", col_uf_origem), ("UF destino", col_uf_destino),
        ("numero da fatura", col_fatura),
    ]:
        if not col:
            avisos_gerais.append(f"Nao achei coluna para '{nome_conceito}' - vai faltar ou ficar em branco.")

    if not col_fatura:
        df = df.copy()
        df["__fatura_unica__"] = "1"
        col_fatura = "__fatura_unica__"
        avisos_gerais.append("Sem coluna de numero de fatura identificada - tratando a planilha inteira como uma unica fatura.")

    resultados = []
    for numero_fatura, grupo in df.groupby(col_fatura):
        avisos = list(avisos_gerais)
        grupo = grupo.reset_index(drop=True)

        cnpjs = grupo[col_chave].map(_cnpj_da_chave).dropna()
        cnpj_emissor_geral = transportadora_cnpj or (cnpjs.mode().iloc[0] if not cnpjs.empty else None)
        if not cnpj_emissor_geral:
            avisos.append("CNPJ do emissor nao pode ser extraido da chave do CT-e.")
            cnpj_emissor_geral = "00000000000000"

        nome_transp = transportadora_nome
        if not nome_transp and col_nome_transp:
            vals = grupo[col_nome_transp].dropna()
            nome_transp = vals.iloc[0] if not vals.empty else None
        if not nome_transp:
            nome_transp = "DESCONHECIDA"
            avisos.append("Nome da transportadora nao encontrado - usando placeholder.")

        nome_tomador_final = tomador_nome
        if col_nome_tomador:
            vals = grupo[col_nome_tomador].dropna()
            if not vals.empty:
                nome_tomador_final = vals.iloc[0]

        if col_dt_emissao:
            datas = grupo[col_dt_emissao].map(_parse_data).dropna()
        else:
            datas = pd.Series([], dtype="object")

        emissao = dt_emissao_fatura or (min(datas) if len(datas) else date.today())
        if not dt_emissao_fatura:
            avisos.append(f"Data de emissao da fatura estimada como {emissao}.")

        vencimento = dt_vencimento_fatura or (emissao + timedelta(days=15))
        if not dt_vencimento_fatura:
            avisos.append(f"Data de vencimento estimada como {vencimento} (emissao + 15 dias) - confirmar com a transportadora.")

        ctes = []
        for _, row in grupo.iterrows():
            chave = row[col_chave]
            dt_emi = _parse_data(row[col_dt_emissao]) if col_dt_emissao else None
            valor = _parse_valor(row[col_valor]) if col_valor else None
            tipo_doc = _normalizar(row[col_tipo_doc]) if col_tipo_doc and pd.notna(row[col_tipo_doc]) else ""

            ctes.append(CTeCobranca(
                filial="",
                numero_doc=_numero_cte_da_chave(chave) or "",
                serie=_serie_da_chave(chave) or "001",
                valor_frete=valor if valor is not None else Decimal("0"),
                dt_emissao=dt_emi or emissao,
                cnpj_rem_nfe=str(row[col_cnpj_rem]) if col_cnpj_rem and pd.notna(row[col_cnpj_rem]) else "",
                cnpj_dest_nfe=str(row[col_cnpj_dest]) if col_cnpj_dest and pd.notna(row[col_cnpj_dest]) else "",
                cnpj_emissor_cte=_cnpj_da_chave(chave) or cnpj_emissor_geral,
                uf_embarcador=str(row[col_uf_origem]) if col_uf_origem and pd.notna(row[col_uf_origem]) else "",
                uf_emissor_cte=str(row[col_uf_origem]) if col_uf_origem and pd.notna(row[col_uf_origem]) else "",
                uf_destino=str(row[col_uf_destino]) if col_uf_destino and pd.notna(row[col_uf_destino]) else "",
                cte_devolucao="S" if (e_devolucao or "DEVOLUCAO" in tipo_doc) else "N",
                cod_iva="",
                numero_pedido=str(row[col_pedido]) if col_pedido and pd.notna(row[col_pedido]) else "",
            ))

        if any(c.valor_frete == 0 for c in ctes):
            avisos.append("Um ou mais CT-e ficaram com valor de frete zerado/nao encontrado - conferir manualmente.")

        valor_total = sum((c.valor_frete for c in ctes), Decimal("0"))

        documento = DocumentoCobranca(
            filial="",
            numero_doc_cobranca=str(numero_fatura),
            dt_emissao=emissao,
            dt_vencimento=vencimento,
            valor_total=valor_total,
            ctes=ctes,
        )

        lote = LoteCobranca(
            transportadora_nome=nome_transp,
            transportadora_cnpj=cnpj_emissor_geral,
            tomador_nome=nome_tomador_final,
            data_hora=datetime.combine(emissao, datetime.min.time()),
            documentos=[documento],
        )
        resultados.append((lote, avisos))

    return resultados
