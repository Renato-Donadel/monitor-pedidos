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

import pandas as pd

from . import csv_transportadora
from . import xlsx_lsp
from . import adaptador_generico

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
    [{"campo": ..., "aviso": ...}, ...]."""
    pendencias = []
    for aviso in avisos:
        for campo, padrao in PADROES_PENDENCIA:
            if padrao.search(aviso):
                pendencias.append({"campo": campo, "aviso": aviso})
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


def montar_lotes_do_arquivo(caminho: str, formato: str, respostas: dict) -> list:
    """Chama o adaptador certo pro formato, repassando as respostas (se
    houver) como parametros explicitos - isso faz o adaptador NAO gerar o
    aviso de campo estimado, porque o valor foi informado de verdade.
    Devolve sempre uma lista de (lote, avisos), mesmo pro formato de lote
    unico (csv_composicao_fretes)."""
    if formato == "csv_composicao_fretes":
        lote, avisos = csv_transportadora.montar_lote(caminho_csv=caminho, **respostas)
        return [(lote, avisos)]

    if formato == "xlsx_lsp":
        e_devolucao = "devolu" in caminho.lower()
        return xlsx_lsp.montar_lotes(caminho, e_devolucao=e_devolucao, **respostas)

    # Formato desconhecido - tenta o adaptador generico (best-effort).
    e_devolucao = "devolu" in caminho.lower()
    return adaptador_generico.montar_lotes(caminho, e_devolucao=e_devolucao, **respostas)
