"""
Deteccao de formato + montagem do(s) lote(s) de cobranca a partir de uma
planilha de transportadora (CSV "COMPOSICAO DOS FRETES" ou xlsx "RELATORIO
DE REMESSAS"), com deteccao de PENDENCIAS (campos que a planilha nao trouxe
e que precisam de confirmacao humana antes de gerar o DOCCOB definitivo).

Copia enxuta de billing/processar_planilhas.py (projeto Z:\\9. Transportes\\
9.10 CTEs\\Projeto_CTE_Completo) - so' a parte de deteccao/montagem, sem a
parte de varrer pasta (nao se aplica aqui: este app e' upload/download, nao
uma pasta monitorada).
"""

import re
from decimal import Decimal

import pandas as pd

from . import csv_transportadora
from . import xlsx_lsp
from . import adaptador_generico
from .doccob import DocumentoCobranca, LoteCobranca

# Avisos que sinalizam um campo ESTIMADO/CHUTADO (nao veio na planilha) que
# precisa de confirmacao humana ANTES de gerar o DOCCOB definitivo.
PADROES_PENDENCIA = [
    ("dt_emissao_fatura", re.compile(r"Data de emissao da fatura nao informada")),
    ("dt_vencimento_fatura", re.compile(r"Data de vencimento da fatura nao informada")),
    ("filial_552", re.compile(r"Codigo de filial \(registro 552\)")),
    ("transportadora_nome", re.compile(r"Nome/razao social de")),
]


def identificar_pendencias(avisos: list) -> list:
    """Devolve as pendencias (campos estimados que a planilha nao trouxe e
    que precisam de confirmacao humana) encontradas na lista de avisos -
    [{"campo": ..., "aviso": ...}, ...]. No maximo UMA pendencia por
    `campo` (esses campos sao de nivel-arquivo/fatura - se o arquivo foi
    dividido em varios lotes por CNPJ emissor, o mesmo aviso aparece
    repetido em cada lote, mas so' precisa ser perguntado uma vez)."""
    pendencias = []
    campos_vistos = set()
    for aviso in avisos:
        for campo, padrao in PADROES_PENDENCIA:
            if padrao.search(aviso):
                if campo not in campos_vistos:
                    pendencias.append({"campo": campo, "aviso": aviso})
                    campos_vistos.add(campo)
                break
    return pendencias


def detectar_formato(caminho: str) -> str:
    """Devolve 'csv_composicao_fretes', 'xlsx_lsp' ou 'desconhecido'."""
    nome = caminho.lower()
    if nome.endswith(".csv"):
        with open(caminho, "r", encoding="latin-1") as f:
            primeira_linha = f.readline()
        if "COMPOSICAO DOS FRETES" in primeira_linha.upper():
            return "csv_composicao_fretes"
        return "desconhecido"
    if nome.endswith(".xlsx"):
        try:
            planilha = pd.ExcelFile(caminho)
        except Exception:
            return "desconhecido"
        if "Remessas" in planilha.sheet_names:
            return "xlsx_lsp"
        return "desconhecido"
    return "desconhecido"


def _dividir_lote_por_cnpj_emissor(lote: LoteCobranca, avisos: list) -> list:
    """Regra de negocio (independente do formato da planilha de origem):
    ANTES de gerar, verifica se todos os CT-e de um lote tem o MESMO CNPJ
    emissor completo (14 digitos - inclui filial, nao so' a raiz). Se sim,
    devolve o lote como esta' (1 arquivo). Se houver N CNPJs diferentes,
    divide em N lotes/arquivos - um por CNPJ - MESMO que sejam filiais da
    mesma transportadora (cada CNPJ completo e' uma "transportadora" pra
    fins de DOCCOB, nao a raiz).

    Numeracao: o primeiro CNPJ (na ordem em que aparece no arquivo) mantem
    o numero de fatura original; os seguintes ganham um prefixo numerico
    (1, 2, 3...) pra nao colidir - ex.: fatura 3313783 com 3 CNPJs vira
    3313783, 13313783, 23313783."""
    doc = lote.documentos[0]

    ctes_por_cnpj = {}
    ordem_cnpjs = []
    for cte in doc.ctes:
        cnpj = cte.cnpj_emissor_cte
        if cnpj not in ctes_por_cnpj:
            ctes_por_cnpj[cnpj] = []
            ordem_cnpjs.append(cnpj)
        ctes_por_cnpj[cnpj].append(cte)

    if len(ordem_cnpjs) <= 1:
        return [(lote, avisos)]

    resultado = []
    for indice, cnpj in enumerate(ordem_cnpjs):
        ctes_do_cnpj = ctes_por_cnpj[cnpj]
        numero_fatura = doc.numero_doc_cobranca if indice == 0 else f"{indice}{doc.numero_doc_cobranca}"

        novo_doc = DocumentoCobranca(
            filial=doc.filial,
            numero_doc_cobranca=numero_fatura,
            dt_emissao=doc.dt_emissao,
            dt_vencimento=doc.dt_vencimento,
            valor_total=sum((c.valor_frete for c in ctes_do_cnpj), Decimal("0")),
            ctes=ctes_do_cnpj,
        )
        novo_lote = LoteCobranca(
            transportadora_nome=lote.transportadora_nome,
            transportadora_cnpj=cnpj,
            tomador_nome=lote.tomador_nome,
            data_hora=lote.data_hora,
            documentos=[novo_doc],
            identificador_intercambio=lote.identificador_intercambio,
            identificador_intercambio_550=lote.identificador_intercambio_550,
        )
        avisos_novo = list(avisos) + [
            f"Fatura original {doc.numero_doc_cobranca} dividida em {len(ordem_cnpjs)} DOCCOB(s) "
            f"porque os CT-e tinham CNPJ emissor diferente - esta parte (CNPJ {cnpj}) "
            f"ficou com o numero {numero_fatura}."
        ]
        resultado.append((novo_lote, avisos_novo))

    return resultado


def montar_lotes_do_arquivo(caminho: str, formato: str, respostas: dict) -> list:
    """Chama o adaptador certo pro formato, repassando as respostas (se
    houver) como parametros explicitos - isso faz o adaptador NAO gerar o
    aviso de campo estimado, porque o valor foi informado de verdade.
    Devolve sempre uma lista de (lote, avisos) - inclusive pro formato de
    lote unico (csv_composicao_fretes), que ainda pode virar mais de um se
    tiver CNPJ emissor variado (ver `_dividir_lote_por_cnpj_emissor`)."""
    if formato == "csv_composicao_fretes":
        lote, avisos = csv_transportadora.montar_lote(caminho_csv=caminho, **respostas)
        lotes_brutos = [(lote, avisos)]
    elif formato == "xlsx_lsp":
        e_devolucao = "devolu" in caminho.lower()
        lotes_brutos = xlsx_lsp.montar_lotes(caminho, e_devolucao=e_devolucao, **respostas)
    else:
        e_devolucao = "devolu" in caminho.lower()
        lotes_brutos = adaptador_generico.montar_lotes(caminho, e_devolucao=e_devolucao, **respostas)

    lotes_finais = []
    for lote, avisos in lotes_brutos:
        lotes_finais.extend(_dividir_lote_por_cnpj_emissor(lote, avisos))
    return lotes_finais
