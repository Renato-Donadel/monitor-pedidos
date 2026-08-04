"""
Adaptador: planilha .xlsx "RELATORIO DE REMESSAS POR PERIODO" (formato LSP)
-> LoteCobranca (estrutura usada por gerar_doccob, ver doccob.py).

O formato do arquivo DOCCOB gerado e' sempre o mesmo (um so padrao,
confirmado byte a byte contra a fatura 3304380 da Carvalima apos subir com
sucesso no PRW - ver csv_transportadora.py). O que muda de transportadora
para transportadora e' so a ORIGEM dos dados (o layout da planilha recebida)
- por isso este adaptador usa exatamente as mesmas convencoes ja validadas
(serie extraida da chave sem alterar zeros, CNPJ do emissor por CT-e
extraido da chave, cod_iva em branco, numero do pedido preenchido), so
trocando os nomes das colunas de origem para os desta planilha.

Diferencas estruturais desta planilha em relacao ao CSV da Carvalima:
  - E' um .xlsx com 3 abas ("Resumo", "Resumo Bravium", "Remessas") - os
    dados por CT-e ficam na aba "Remessas". A primeira linha da planilha e'
    um titulo, a segunda linha e' o cabecalho de verdade (header=1).
  - Um UNICO arquivo cobre VARIAS faturas ao mesmo tempo (uma quinzena
    inteira) - por isso `montar_lotes` devolve uma lista, uma LoteCobranca
    por numero de fatura encontrado (a regra de 1 CNPJ tomador + 1 CNPJ
    emissor por arquivo DOCCOB continua valendo, um arquivo final por
    fatura).
  - Nao ha campo que diga se a linha e' "devolucao" - isso so se sabe pelo
    arquivo em si (o nome do arquivo recebido tras isso, ex.:
    "..._Descritivo_Devolucao.xlsx" vs "..._Descritivo_Envio.xlsx") - por
    isso e' parametro da funcao.
  - CNPJ do emissor/tomador tambem existem como colunas ("CNPJ
    Emitente"/"CNPJ Tomador"), mas o pandas LE ESSAS COLUNAS COMO NUMERO E
    PERDE ZEROS A ESQUERDA (ex.: 1336140000874 em vez de 01336140000874) -
    por isso o CNPJ do emissor e' extraido da chave de acesso do CT-e
    (coluna "Dacte"), que e' texto e nao perde digitos.
"""

from datetime import date, datetime
from decimal import Decimal

import pandas as pd

from .doccob import CTeCobranca, DocumentoCobranca, LoteCobranca


def _cnpj_da_chave(chave) -> str | None:
    chave = "".join(ch for ch in str(chave) if ch.isdigit())
    if len(chave) != 44:
        return None
    return chave[6:20]


def _serie_da_chave(chave) -> str | None:
    """Mesma tecnica validada em csv_transportadora.py: serie tal como esta
    na chave (3 digitos, sem remover zeros a esquerda)."""
    chave = "".join(ch for ch in str(chave) if ch.isdigit())
    if len(chave) != 44:
        return None
    return chave[22:25]


def _numero_cte_da_chave(chave) -> str | None:
    chave = "".join(ch for ch in str(chave) if ch.isdigit())
    if len(chave) != 44:
        return None
    return str(int(chave[25:34]))


def ler_remessas(caminho_xlsx: str) -> pd.DataFrame:
    """Le a aba 'Remessas' do relatorio LSP (pula a linha de titulo)."""
    return pd.read_excel(caminho_xlsx, sheet_name="Remessas", header=1)


def montar_lotes(
    caminho_xlsx: str,
    e_devolucao: bool = False,
    dt_emissao_fatura: date = None,
    dt_vencimento_fatura: date = None,
    tomador_nome: str = None,
    transportadora_nome: str = None,
) -> list[tuple[LoteCobranca, list[str]]]:
    """Monta uma LoteCobranca para CADA numero de fatura encontrado na
    planilha (um arquivo LSP pode cobrir varias faturas de uma quinzena).
    Devolve lista de (lote, avisos).
    """
    df = ler_remessas(caminho_xlsx)
    resultados = []

    for numero_fatura, grupo in df.groupby("Número da fatura"):
        avisos = []
        grupo = grupo.reset_index(drop=True)

        cnpjs_emissor = grupo["Dacte"].map(_cnpj_da_chave).dropna()
        cnpj_emissor_geral = cnpjs_emissor.mode().iloc[0] if not cnpjs_emissor.empty else None
        if not cnpj_emissor_geral:
            avisos.append("CNPJ do emissor nao pode ser extraido da chave do CT-e (coluna Dacte).")
            cnpj_emissor_geral = "00000000000000"

        nome_transportadora = transportadora_nome or (grupo["Emissor"].dropna().iloc[0] if not grupo["Emissor"].dropna().empty else None)
        if not nome_transportadora:
            nome_transportadora = "DESCONHECIDA"
            avisos.append("Nome da transportadora (coluna Emissor) nao encontrado.")

        nome_tomador = tomador_nome or (grupo["Nome Tomador"].dropna().iloc[0] if not grupo["Nome Tomador"].dropna().empty else "BRAVIUM S.A")

        datas_emissao = pd.to_datetime(grupo["Data"], format="%d/%m/%Y", errors="coerce").dropna()

        emissao = dt_emissao_fatura or (datas_emissao.min().date() if not datas_emissao.empty else None)
        if not dt_emissao_fatura:
            avisos.append(f"Data de emissao da fatura nao informada - estimada como {emissao} (menor data de emissao das remessas).")

        if dt_vencimento_fatura:
            vencimento = dt_vencimento_fatura
        else:
            from datetime import timedelta
            vencimento = emissao + timedelta(days=15)
            avisos.append(f"Data de vencimento da fatura nao informada - estimada como {vencimento} (emissao + 15 dias). Confirmar com o e-mail da transportadora.")

        ctes = []
        for _, row in grupo.iterrows():
            chave = row["Dacte"]
            dt_emi = pd.to_datetime(row["Data"], format="%d/%m/%Y", errors="coerce")
            ctes.append(CTeCobranca(
                # mesma ideia da Carvalima (praca/cidade de origem, 3
                # caracteres) - aqui usando a cidade de inicio da prestacao
                filial=str(row.get("Cidade Inicio Prestação", "") or "")[:3].upper(),
                numero_doc=_numero_cte_da_chave(chave) or str(row.get("Cte", "")),
                serie=_serie_da_chave(chave) or "001",
                valor_frete=Decimal(str(row["Valor"])),
                dt_emissao=dt_emi.date() if pd.notna(dt_emi) else emissao,
                cnpj_rem_nfe=str(row.get("Cpf/CNPJ Remetente", "")),
                cnpj_dest_nfe=str(row.get("Cpf/CNPJ Destinatario", "")),
                cnpj_emissor_cte=_cnpj_da_chave(chave) or cnpj_emissor_geral,
                uf_embarcador=str(row.get("UF Inicio Prestação", "")),
                uf_emissor_cte=str(row.get("UF Inicio Prestação", "")),
                uf_destino=str(row.get("UF Fim Prestação", "")),
                cte_devolucao="S" if e_devolucao else "N",
                cod_iva="",  # mesma convencao validada com a Carvalima
                numero_pedido=str(row.get("Pedido", "") or ""),
            ))

        valor_total = sum((c.valor_frete for c in ctes), Decimal("0"))

        documento = DocumentoCobranca(
            filial=nome_transportadora[:10],
            numero_doc_cobranca=str(int(numero_fatura)),
            dt_emissao=emissao,
            dt_vencimento=vencimento,
            valor_total=valor_total,
            ctes=ctes,
        )

        lote = LoteCobranca(
            transportadora_nome=nome_transportadora,
            transportadora_cnpj=cnpj_emissor_geral,
            tomador_nome=nome_tomador,
            data_hora=datetime.combine(emissao, datetime.min.time()),
            documentos=[documento],
        )
        resultados.append((lote, avisos))

    return resultados
