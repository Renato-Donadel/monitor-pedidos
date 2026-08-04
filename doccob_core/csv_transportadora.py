"""
Adaptador: CSV de composicao de fretes da transportadora -> LoteCobranca
(estrutura usada por gerar_doccob, ver doccob.py).

Formato confirmado com o CSV real "COMPOSICAO DOS FRETES DOS CTRCS"
(observado em arquivos de OTC e FAV - mesmo layout, 112 colunas,
separador ";", encoding latin-1, 1a linha = metadados, 2a linha = cabecalho,
ultima linha = "9;" (rodape)).

Campos que o CSV NAO traz (data de emissao/vencimento da fatura, codigo de
filial) sao parametros obrigatorios desta funcao - em uso real, quem estiver
operando o agente deve perguntar essas informacoes ao usuario (ex.: a data
de vencimento vem no e-mail da transportadora, nao no CSV).
"""

import csv
from collections import Counter
from datetime import date, datetime
from decimal import Decimal

from .doccob import CTeCobranca, DocumentoCobranca, LoteCobranca
from . import cadastro_transportadoras as cadastro

# indices das colunas relevantes (0-based) no CSV real de 112 colunas
COL_NUMERO_CTE = 10
COL_PRACA_EXPEDIDORA = 12      # usado como "filial" do registro 555 (confirmado contra referencia real)
COL_CODIGO_BARRAS_DACTE = 13   # chave de acesso do CT-e (44 digitos)
COL_TIPO_DOCUMENTO = 14
COL_DATA_EMISSAO = 16
COL_CNPJ_REMETENTE = 19
COL_UF_REMETENTE = 23
COL_UF_EXPEDIDOR = 29
COL_CNPJ_DESTINATARIO = 31
COL_UF_ENTREGA = 42
COL_NUMERO_PEDIDO = 102
COL_VAL_RECEBER = 94
COL_NUMERO_FATURA = 99


def _cnpj_da_chave_cte(chave: str) -> str | None:
    """Extrai o CNPJ do emissor a partir da chave de acesso do CT-e
    (posicoes 7-20, formato padrao NFe/CTe: cUF(2)+AAMM(4)+CNPJ(14)+...)."""
    chave = "".join(ch for ch in chave if ch.isdigit())
    if len(chave) != 44:
        return None
    return chave[6:20]


def _numero_cte_da_chave(chave: str) -> str | None:
    """Extrai o numero do CT-e (nCT, 9 digitos) a partir da chave de acesso
    (posicoes 26-34). O CSV da transportadora traz "NUMERO CT-E" como
    serie+numero concatenados (12 digitos) - o campo do DOCCOB (registro
    555) espera so o numero, sem a serie (confirmado contra os 82 exemplos
    reais, onde o campo e' um numero puro tipo "359420430")."""
    chave = "".join(ch for ch in chave if ch.isdigit())
    if len(chave) != 44:
        return None
    return str(int(chave[25:34]))


def _serie_da_chave(chave: str) -> str | None:
    """Extrai a serie do CT-e (3 digitos, posicoes 23-25) da chave de acesso.
    CRITICO: uma referencia real (DOCCOB gerado pelo PRW para a fatura
    3304380) mostrou que usar uma serie fixa/errada aqui faz o sistema nao
    reconhecer o CT-e - a serie tem que ser a verdadeira, extraida da chave,
    nao um valor generico."""
    chave = "".join(ch for ch in chave if ch.isdigit())
    if len(chave) != 44:
        return None
    return chave[22:25]


def _limpar(txt: str) -> str:
    return txt.strip().replace("\xa0", "").strip()


def _valor_br(txt: str) -> Decimal:
    txt = _limpar(txt).replace(".", "").replace(",", ".")
    return Decimal(txt or "0")


def _data_br(txt: str) -> date:
    return datetime.strptime(_limpar(txt), "%d/%m/%Y").date()


def ler_csv_transportadora(caminho: str, encoding: str = "latin-1") -> tuple[str, list[dict]]:
    """Le o CSV e devolve (codigo_transportadora, lista_de_registros) - um
    registro por CT-e, ja com os campos relevantes limpos e convertidos.
    Ignora a linha de metadados, a de cabecalho e o rodape ("9;...")."""
    with open(caminho, "r", encoding=encoding) as f:
        linhas = f.readlines()

    meta = linhas[0].rstrip("\r\n").split(";")
    codigo_transportadora = meta[1].strip() if len(meta) > 1 else "DESCONHECIDA"

    registros = []
    for linha in linhas[2:]:
        campos = linha.rstrip("\r\n").split(";")
        if len(campos) <= COL_NUMERO_FATURA:
            continue  # linha de rodape ou incompleta
        registros.append({
            "numero_cte": _limpar(campos[COL_NUMERO_CTE]),
            "praca_expedidora": _limpar(campos[COL_PRACA_EXPEDIDORA]),
            "chave_cte": _limpar(campos[COL_CODIGO_BARRAS_DACTE]),
            "tipo_documento": _limpar(campos[COL_TIPO_DOCUMENTO]),
            "dt_emissao": _data_br(campos[COL_DATA_EMISSAO]),
            "cnpj_remetente": _limpar(campos[COL_CNPJ_REMETENTE]),
            "uf_remetente": _limpar(campos[COL_UF_REMETENTE]),
            "uf_expedidor": _limpar(campos[COL_UF_EXPEDIDOR]),
            "cnpj_destinatario": _limpar(campos[COL_CNPJ_DESTINATARIO]),
            "uf_entrega": _limpar(campos[COL_UF_ENTREGA]),
            "val_receber": _valor_br(campos[COL_VAL_RECEBER]),
            "numero_fatura": _limpar(campos[COL_NUMERO_FATURA]),
            "numero_pedido": _limpar(campos[COL_NUMERO_PEDIDO]),
        })
    return codigo_transportadora, registros


def identificar_cnpj_raiz(caminho_csv: str) -> str | None:
    """Le o CSV e devolve a raiz do CNPJ (8 primeiros digitos) da
    transportadora, extraida das chaves de acesso dos CT-e (maioria, caso
    haja ruido) - a mesma logica usada em `montar_lote`. Usado por quem
    precisa identificar a transportadora (chave do cadastro) SEM montar o
    lote inteiro."""
    _codigo, registros = ler_csv_transportadora(caminho_csv)
    cnpjs_da_chave = Counter(
        c for r in registros if (c := _cnpj_da_chave_cte(r["chave_cte"]))
    )
    cnpj_extraido = cnpjs_da_chave.most_common(1)[0][0] if cnpjs_da_chave else None
    return cadastro.raiz_cnpj(cnpj_extraido) if cnpj_extraido else None


def montar_lote(
    caminho_csv: str,
    tomador_nome: str = "BRAVIUM S.A",
    transportadora_cnpj: str = None,
    transportadora_nome: str = None,
    dt_emissao_fatura: date = None,
    dt_vencimento_fatura: date = None,
    filial_552: str = None,
    filial_555: str = None,
    data_hora: datetime = None,
) -> tuple[LoteCobranca, list[str]]:
    """Monta o LoteCobranca a partir do CSV da transportadora. Devolve
    (lote, avisos) - avisos lista o que foi estimado/assumido e precisa de
    confirmacao humana antes de considerar o arquivo definitivo.

    O CNPJ da transportadora e' extraido automaticamente da chave de acesso
    do CT-e (nao precisa ser informado). Nome e codigos de filial vem do
    cadastro persistente (data/cobranca/cadastro_transportadoras.json) se ja
    tiverem sido informados antes; senao ficam como placeholder e um aviso e'
    gerado. dt_emissao/dt_vencimento da fatura, se nao informados, sao
    estimados (emissao = menor data de emissao dos CT-e do arquivo,
    vencimento = emissao + 15 dias) - sempre gera aviso, pois normalmente vem
    do e-mail da transportadora, nao do CSV.
    """
    avisos = []
    codigo, registros = ler_csv_transportadora(caminho_csv)
    if not registros:
        raise ValueError(f"Nenhum registro de CT-e encontrado em {caminho_csv!r}.")

    numeros_fatura = {r["numero_fatura"] for r in registros}
    if len(numeros_fatura) > 1:
        raise ValueError(
            f"O CSV traz mais de uma fatura ({numeros_fatura}) - "
            f"a regra e' um arquivo DOCCOB por fatura/tomador/transportadora."
        )
    numero_fatura = numeros_fatura.pop()

    # CNPJ da transportadora: extraido da chave de acesso (maioria dos CT-e,
    # caso haja ruido). O cadastro e' consultado pela RAIZ do CNPJ (8
    # primeiros digitos) - NUNCA pelo "codigo" (rotulo do arquivo, tipo
    # "OTC") - o mesmo rotulo pode nao identificar a mesma empresa entre
    # arquivos diferentes, mas o CNPJ nao mente (ver cadastro_transportadoras.py).
    cnpjs_da_chave = Counter(
        c for r in registros if (c := _cnpj_da_chave_cte(r["chave_cte"]))
    )
    cnpj_extraido = cnpjs_da_chave.most_common(1)[0][0] if cnpjs_da_chave else None
    cnpj_raiz = cadastro.raiz_cnpj(cnpj_extraido) if cnpj_extraido else None

    if cnpj_raiz:
        entrada_cadastro = cadastro.obter_ou_criar(cnpj_raiz, cnpj_completo=cnpj_extraido, codigo_visto=codigo)
    else:
        entrada_cadastro = {}
        avisos.append(f"Nao foi possivel extrair o CNPJ da transportadora '{codigo}' das chaves de CT-e - cadastro nao pode ser consultado.")

    cnpj_final = transportadora_cnpj or cnpj_extraido
    if not cnpj_final:
        avisos.append(f"CNPJ da transportadora '{codigo}' nao pode ser determinado.")
        cnpj_final = "00000000000000"

    nome_final = transportadora_nome or entrada_cadastro.get("nome")
    if not nome_final:
        nome_final = codigo
        avisos.append(f"Nome/razao social de '{codigo}' (CNPJ raiz {cnpj_raiz}) nao confirmado - usando o codigo '{codigo}' como placeholder.")

    filial_552_final = filial_552 or entrada_cadastro.get("filial_552")
    if not filial_552_final:
        filial_552_final = f"{codigo} /"
        avisos.append(f"Codigo de filial (registro 552) de '{codigo}' (CNPJ raiz {cnpj_raiz}) nao confirmado - usando placeholder {filial_552_final!r}.")

    if not dt_emissao_fatura:
        dt_emissao_fatura = min(r["dt_emissao"] for r in registros)
        avisos.append(f"Data de emissao da fatura nao informada - estimada como {dt_emissao_fatura} (menor data de emissao entre os CT-e).")

    if not dt_vencimento_fatura:
        from datetime import timedelta
        dt_vencimento_fatura = dt_emissao_fatura + timedelta(days=15)
        avisos.append(f"Data de vencimento da fatura nao informada - estimada como {dt_vencimento_fatura} (emissao + 15 dias). Confirmar com o e-mail da transportadora.")

    ctes = [
        CTeCobranca(
            # confirmado contra referencia real (fatura 3304380): so os 3
            # primeiros caracteres da praca expedidora (ex.: CSV traz "SAOI",
            # DOCCOB usa "SAO").
            filial=(r["praca_expedidora"] or (filial_555 or "000"))[:3],
            numero_doc=_numero_cte_da_chave(r["chave_cte"]) or r["numero_cte"],
            serie=_serie_da_chave(r["chave_cte"]) or "002",
            valor_frete=r["val_receber"],
            dt_emissao=r["dt_emissao"],
            cnpj_rem_nfe=r["cnpj_remetente"],
            cnpj_dest_nfe=r["cnpj_destinatario"],
            # CNPJ por CT-e (nao um "CNPJ mais comum do arquivo"): uma
            # referencia real (fatura 3304381, CT-e de devolucao) mostrou que
            # o CNPJ emissor pode variar entre CT-e da MESMA fatura/mesma
            # transportadora (ex.: filial diferente para devolucao).
            cnpj_emissor_cte=_cnpj_da_chave_cte(r["chave_cte"]) or cnpj_final,
            uf_embarcador=r["uf_expedidor"],
            uf_emissor_cte=r["uf_remetente"],
            uf_destino=r["uf_entrega"],
            cte_devolucao="S" if r["tipo_documento"].upper() == "DEVOLUCAO" else "N",
            # confirmado contra referencia real: cod_iva fica em branco para
            # este formato de CSV (diferente do padrao "Z3" visto nos
            # exemplos da J&T).
            cod_iva="",
            numero_pedido=r["numero_pedido"],
        )
        for r in registros
    ]

    valor_total = sum((c.valor_frete for c in ctes), Decimal("0"))

    documento = DocumentoCobranca(
        filial=filial_552_final,
        numero_doc_cobranca=numero_fatura,
        dt_emissao=dt_emissao_fatura,
        dt_vencimento=dt_vencimento_fatura,
        valor_total=valor_total,
        ctes=ctes,
    )

    lote = LoteCobranca(
        transportadora_nome=nome_final,
        transportadora_cnpj=cnpj_final,
        tomador_nome=tomador_nome,
        data_hora=data_hora or datetime.combine(dt_emissao_fatura, datetime.min.time()),
        documentos=[documento],
    )
    return lote, avisos
