"""
Geracao e leitura do arquivo DOCCOB (padrao Proceda 5.0) - Etapa 6.

Regra de negocio: um unico CNPJ tomador e um unico CNPJ emissor de CT-e por
arquivo -> por isso este modulo trabalha com exatamente 1 registro 550/551/552
por arquivo (o caso real observado nos 82 exemplos fornecidos), com N
registros 555 (um por CT-e).

Posicoes de campo confirmadas contra a especificacao oficial
(docs/5-doccob (3).pdf) E contra 82 arquivos reais em producao (todas as
linhas tem 280 posicoes). Duas divergencias entre a spec e a realidade,
usadas aqui:
  - Registro 552, campo "serie do documento" (pos 15, tam 3) e campo
    "tipo de cobranca" (pos 59, tam 3): a spec marca como opcional/riscado
    ou sugere BCO/CAR, mas os 82 exemplos reais SEMPRE usam "FAT" nos dois.
  - Registro 556 ("Notas Fiscais"): a spec diz que nao e' necessario
    ("CTe ja e' informado no arquivo") e os exemplos reais o usam mesmo
    assim, mas sem tabela de campos documentada. Por decisao do usuario,
    este gerador NAO produz o registro 556 - so 000/550/551/552/555/559.

Todas as linhas tem exatamente 280 posicoes, preenchidas com espaco a
direita quando o conteudo e' menor que o campo.
"""

from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP

TAMANHO_LINHA = 280

# Constantes observadas como fixas em 82 arquivos reais (nao documentadas
# como tal na spec, mas nunca variam nos exemplos fornecidos):
CONTA_RAZAO_PADRAO = "422619"   # registro 555, pos 102 (10) - conta contabil
COD_IVA_PADRAO = "Z3"           # registro 555, pos 112 (2)
ACAO_DOC_PADRAO = "I"           # registro 552, pos 156 (1) - "I" em 100% dos exemplos


# ---------------------------------------------------------------------------
# Helpers de formatacao fixed-width
# ---------------------------------------------------------------------------

def _texto(valor, tamanho: int) -> str:
    """Alinha a esquerda, corta ou completa com espaco."""
    valor = "" if valor is None else str(valor)
    return valor[:tamanho].ljust(tamanho)


def _numero(valor, tamanho: int) -> str:
    """Alinha a direita, completa com zero a esquerda."""
    valor = "" if valor is None else str(valor)
    return valor[:tamanho].rjust(tamanho, "0")


def _data(d: date, tamanho: int = 8) -> str:
    """Formato DDMMAAAA."""
    return d.strftime("%d%m%Y") if len(f"{d:%Y}") == 4 else d.strftime("%d%m%y")


def _valor(v, tamanho: int = 15) -> str:
    """15 posicoes: 13 inteiras + 2 decimais, sem separador, zero a esquerda."""
    centavos = int((Decimal(str(v)) * 100).to_integral_value(rounding=ROUND_HALF_UP))
    return str(centavos).rjust(tamanho, "0")


def _linha(*partes: str) -> str:
    linha = "".join(partes)
    if len(linha) > TAMANHO_LINHA:
        raise ValueError(f"Linha excedeu {TAMANHO_LINHA} posicoes ({len(linha)}): {linha!r}")
    return linha.ljust(TAMANHO_LINHA)


# ---------------------------------------------------------------------------
# Modelo de dados
# ---------------------------------------------------------------------------

@dataclass
class CTeCobranca:
    filial: str            # registro 555, pos 4 (10) - codigo interno da filial (ex.: "0102")
    numero_doc: str         # pos 19 (12) - numero do CT-e
    valor_frete: Decimal    # pos 31 (15)
    dt_emissao: date        # pos 46 (8)
    cnpj_rem_nfe: str       # pos 54 (14) - CNPJ remetente da NF-e (normalmente a propria Bravium)
    cnpj_dest_nfe: str      # pos 68 (14) - CNPJ destinatario da NF-e
    cnpj_emissor_cte: str   # pos 82 (14) - CNPJ da transportadora emissora do CT-e
    uf_embarcador: str      # pos 96 (2)
    uf_emissor_cte: str     # pos 98 (2)
    uf_destino: str         # pos 100 (2)
    cte_devolucao: str = "N"  # pos 194 (1) - "S"/"N"
    serie: str = "002"        # pos 14 (5) - serie do CT-e (usar a serie real, extraida da chave - ver csv_transportadora.py)
    cod_iva: str = COD_IVA_PADRAO  # pos 112 (2) - NAO e' universal: confirmado "Z3" para J&T, mas em branco para Carvalima/OTC
    numero_pedido: str = ""   # pos 114 (20, campo "num_romaneio") - NAO e' universal: em branco para J&T, mas a Carvalima usa o numero do pedido aqui


@dataclass
class DocumentoCobranca:
    filial: str               # registro 552, pos 4 (10) - label da filial do emitente (ex.: "J&T - SP /")
    numero_doc_cobranca: str   # pos 18 (10)
    dt_emissao: date           # pos 28 (8)
    dt_vencimento: date        # pos 36 (8)
    valor_total: Decimal       # pos 44 (15)
    ctes: list = field(default_factory=list)  # list[CTeCobranca]


@dataclass
class LoteCobranca:
    """Um arquivo DOCCOB = um lote: 1 transportadora (CNPJ emissor de CT-e),
    1 tomador (cliente), N documentos de cobranca."""
    transportadora_nome: str      # registro 000, pos 4 (35) / registro 551, pos 18 (50)
    transportadora_cnpj: str      # registro 551, pos 4 (14)
    tomador_nome: str             # registro 000, pos 39 (35)
    data_hora: datetime           # registro 000, pos 74 (data) + 80 (hora)
    documentos: list = field(default_factory=list)  # list[DocumentoCobranca]
    identificador_intercambio: str = None  # registro 000 - se None, gerado a partir de data_hora
    identificador_intercambio_550: str = None  # registro 550 - se None, usa o mesmo do 000


def _gerar_identificador(data_hora: datetime) -> str:
    # Convencao propria (a spec so sugere um formato, exemplos reais usam
    # convencoes proprias de cada vendor) - "COB" + DDMMHHMM, 11 caracteres.
    return "COB" + data_hora.strftime("%d%m%H%M")


# ---------------------------------------------------------------------------
# Geracao
# ---------------------------------------------------------------------------

def gerar_doccob(lote: LoteCobranca) -> str:
    id_intercambio = lote.identificador_intercambio or _gerar_identificador(lote.data_hora)

    linhas = []

    # 000 - cabecalho de intercambio
    linhas.append(_linha(
        "000",
        _texto(lote.transportadora_nome, 35),
        _texto(lote.tomador_nome, 35),
        lote.data_hora.strftime("%d%m%y"),
        lote.data_hora.strftime("%H%M"),
        _texto(id_intercambio, 12),
    ))

    # 550 - cabecalho de documento
    id_550 = lote.identificador_intercambio_550 or id_intercambio
    linhas.append(_linha("550", _texto(id_550, 14)))

    # 551 - dados da transportadora
    linhas.append(_linha(
        "551",
        _numero(lote.transportadora_cnpj, 14),
        _texto(lote.transportadora_nome, 50),
    ))

    for doc in lote.documentos:
        # 552 - documento de cobranca
        linhas.append(_linha(
            "552",
            _texto(doc.filial, 10),
            "0",                                # tipo do documento: 0 = Fatura de NF
            _texto("FAT", 3),                    # "serie" - observado sempre "FAT" nos exemplos reais
            _numero(doc.numero_doc_cobranca, 10),
            _data(doc.dt_emissao),
            _data(doc.dt_vencimento),
            _valor(doc.valor_total),
            _texto("FAT", 3),                    # tipo de cobranca - observado sempre "FAT"
            _numero(0, 4),                       # pct_multa - sempre zero nos exemplos
            _texto("", 15),                      # valor_juros_dia - sempre vazio
            _data(doc.dt_vencimento),             # dt_limite_desc - sempre igual a dt_vencimento
            _valor(0),                           # valor_desc - sempre zero
            _numero(0, 5),                       # cod_agente - sempre zero
            _texto("", 30),                      # nome_agente - sempre vazio
            _numero(0, 4),                       # num_agencia - sempre zero
            _texto(" ", 1),                      # dv_agencia - sempre em branco
            _numero(0, 10),                      # num_cc - sempre zero
            _texto("  ", 2),                     # dv_cc - sempre em branco
            _texto(ACAO_DOC_PADRAO, 1),           # acao_doc - sempre "I"
            _texto("", 10),                      # id_prefat - sempre vazio
            _texto("", 20),                      # id_add_prefat - sempre vazio
            _texto("", 5),                       # cfop - sempre vazio
            _numero(int(doc.numero_doc_cobranca), 9),  # cod_nfe - mesmo numero do doc de cobranca (sem zeros a esquerda antes de reformatar)
            _texto("", 45),                      # chave_acesso_dv - sempre vazio
            _texto("", 15),                      # num_protocolo - sempre vazio
        ))

        # 555 - um por CT-e
        for cte in doc.ctes:
            linhas.append(_linha(
                "555",
                _texto(cte.filial, 10),
                _texto(cte.serie, 5),
                _texto(cte.numero_doc, 12),
                _valor(cte.valor_frete),
                _data(cte.dt_emissao),
                _numero(cte.cnpj_rem_nfe, 14),
                _numero(cte.cnpj_dest_nfe, 14),
                _numero(cte.cnpj_emissor_cte, 14),
                _texto(cte.uf_embarcador, 2),
                _texto(cte.uf_emissor_cte, 2),
                _texto(cte.uf_destino, 2),
                _texto(CONTA_RAZAO_PADRAO, 10),
                _texto(cte.cod_iva, 2),
                _texto(cte.numero_pedido, 20),    # num_romaneio (Carvalima usa; J&T deixa vazio)
                _texto("", 60),                  # num_sap1/2/3 - sempre vazios
                _texto(cte.cte_devolucao, 1),
            ))

    # 559 - totais (1 por registro 550 - aqui 1 unico 550 por arquivo)
    qtde_docs = len(lote.documentos)
    valor_total_docs = sum((d.valor_total for d in lote.documentos), Decimal("0"))
    linhas.append(_linha(
        "559",
        _numero(qtde_docs, 4),
        _valor(valor_total_docs),
    ))

    return "\r\n".join(linhas) + "\r\n"


# ---------------------------------------------------------------------------
# Leitura (quando a transportadora ja envia o DOCCOB pronto)
# ---------------------------------------------------------------------------

def _campo(linha: str, pos: int, tamanho: int) -> str:
    """pos e' 1-based, igual a especificacao."""
    return linha[pos - 1:pos - 1 + tamanho]


def _parse_data(txt: str) -> date | None:
    txt = txt.strip()
    if not txt or txt == "0" * len(txt):
        return None
    return datetime.strptime(txt, "%d%m%Y").date()


def _parse_valor(txt: str) -> Decimal:
    return Decimal(txt) / 100


def parse_doccob(texto: str) -> LoteCobranca:
    """Le um arquivo DOCCOB (recebido pronto da transportadora) e devolve a
    mesma estrutura usada por gerar_doccob. Ignora registros 553/556
    (nao usados por este projeto - ver decisao no topo do arquivo)."""
    linhas = [l for l in texto.splitlines() if l.strip()]

    l000 = next(l for l in linhas if l[:3] == "000")
    l550 = next(l for l in linhas if l[:3] == "550")
    l551 = next(l for l in linhas if l[:3] == "551")

    dt_str = _campo(l000, 74, 6)
    hr_str = _campo(l000, 80, 4)
    data_hora = datetime.strptime(dt_str + hr_str, "%d%m%y%H%M")

    lote = LoteCobranca(
        transportadora_nome=_campo(l000, 4, 35).strip(),
        transportadora_cnpj=_campo(l551, 4, 14).strip(),
        tomador_nome=_campo(l000, 39, 35).strip(),
        data_hora=data_hora,
        identificador_intercambio=_campo(l000, 84, 12).strip(),
        identificador_intercambio_550=_campo(l550, 4, 14).strip(),
    )

    doc_atual = None
    for l in linhas:
        tipo = l[:3]
        if tipo == "552":
            doc_atual = DocumentoCobranca(
                filial=_campo(l, 4, 10).strip(),
                numero_doc_cobranca=_campo(l, 18, 10).strip(),
                dt_emissao=_parse_data(_campo(l, 28, 8)),
                dt_vencimento=_parse_data(_campo(l, 36, 8)),
                valor_total=_parse_valor(_campo(l, 44, 15)),
            )
            lote.documentos.append(doc_atual)
        elif tipo == "555":
            if doc_atual is None:
                raise ValueError("Registro 555 encontrado antes de um registro 552.")
            doc_atual.ctes.append(CTeCobranca(
                filial=_campo(l, 4, 10).strip(),
                serie=_campo(l, 14, 5).strip(),
                numero_doc=_campo(l, 19, 12).strip(),
                valor_frete=_parse_valor(_campo(l, 31, 15)),
                dt_emissao=_parse_data(_campo(l, 46, 8)),
                cnpj_rem_nfe=_campo(l, 54, 14).strip(),
                cnpj_dest_nfe=_campo(l, 68, 14).strip(),
                cnpj_emissor_cte=_campo(l, 82, 14).strip(),
                uf_embarcador=_campo(l, 96, 2).strip(),
                uf_emissor_cte=_campo(l, 98, 2).strip(),
                uf_destino=_campo(l, 100, 2).strip(),
                cte_devolucao=_campo(l, 194, 1).strip() or "N",
            ))
        # 000/550/551/559/553/556 ja tratados ou ignorados de proposito

    return lote


# ---------------------------------------------------------------------------
# Upload ao PRW - EM ABERTO
# ---------------------------------------------------------------------------
# O PDF (PROC001) descreve o fluxo hoje como uma tela ("Integracao EDI
# PROCEDA") que consulta Financeiro_Movimento/FATURAMENTO_CTE/cad_cliente e
# manda os dados para um webservice, que devolve o arquivo. Falta confirmar
# qual e' esse webservice (endpoint, autenticacao) para automatizar o envio
# real - nao implementado ainda, nao enviar nada de verdade sem confirmar.
