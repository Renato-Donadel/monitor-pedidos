"""
Pagina "Gerador de DOCCOB": ferramenta de FALLBACK do time de cobranca.

Contexto: no fluxo automatico (projeto separado de monitoramento de CT-e),
um agente busca nos e-mails o arquivo DOCCOB que a transportadora deveria
mandar pronto (padrao Proceda 5.0). Quando a transportadora NAO manda esse
arquivo (ou manda errado), alguem do time usa esta pagina pra gerar o
DOCCOB manualmente a partir da planilha de composicao de fretes que a
transportadora manda em vez disso (CSV "COMPOSICAO DOS FRETES" ou xlsx
"RELATORIO DE REMESSAS").

Codigo proprio e autocontido: toda a logica de adaptacao/geracao do DOCCOB
vive em `doccob_core/` (pacote local deste repositorio, nao depende de
nenhum outro projeto/drive de rede) - pode rodar tanto localmente (.bat)
quanto no site hospedado (Streamlit Community Cloud).

So' upload -> download: o arquivo enviado fica so' num temp local (apagado
ao processar outro ou fechar a sessao), nada e' salvo permanentemente.

Excecao: a confirmacao de codigo de filial/nome de transportadora fica
salva em data/doccob_cadastro_transportadoras.json (deste repositorio) pra
nao perguntar de novo da proxima vez. No site hospedado isso e' efemero
(some quando o container reinicia) - so' funciona de forma persistente de
verdade quando rodado localmente.
"""

import os
import re
import tempfile
from datetime import date, datetime

import streamlit as st

from doccob_core.orquestrador import detectar_formato, montar_lotes_do_arquivo, identificar_pendencias
from doccob_core.doccob import gerar_doccob
from doccob_core import csv_transportadora
from doccob_core import cadastro_transportadoras as cadastro

ROTULOS_CAMPO = {
    "dt_emissao_fatura": "Data de emissao da fatura",
    "dt_vencimento_fatura": "Data de vencimento da fatura",
    "filial_552": "Codigo de filial (registro 552)",
    "transportadora_nome": "Nome/razao social da transportadora",
}
CAMPOS_DATA = {"dt_emissao_fatura", "dt_vencimento_fatura"}


def _sugestao_data(aviso: str) -> date:
    m = re.search(r"estimada como (\d{4}-\d{2}-\d{2})", aviso)
    if m:
        return datetime.strptime(m.group(1), "%Y-%m-%d").date()
    return date.today()


def _sugestao_texto(aviso: str) -> str:
    m = re.search(r"placeholder '([^']*)'", aviso) or re.search(r"usando o codigo '([^']*)'", aviso)
    return m.group(1) if m else ""


def _limpar_estado():
    caminho = st.session_state.get("doccob_caminho_temp")
    if caminho and os.path.exists(caminho):
        try:
            os.remove(caminho)
        except OSError:
            pass
    st.session_state.doccob_nome_original = None
    st.session_state.doccob_caminho_temp = None
    st.session_state.doccob_resultado = None
    st.session_state.doccob_respostas = {}
    st.session_state.doccob_cadastro_salvo = False


def render_doccob():
    st.markdown('<div class="data-title">Gerador de DOCCOB (Proceda 5.0)</div>', unsafe_allow_html=True)
    st.caption(
        "Ferramenta manual - use quando a transportadora nao enviar (ou enviar errado) o arquivo "
        "DOCCOB pronto. So' upload e download - nada fica salvo no servidor."
    )

    for chave, padrao in [
        ("doccob_nome_original", None), ("doccob_caminho_temp", None),
        ("doccob_resultado", None), ("doccob_respostas", {}), ("doccob_cadastro_salvo", False),
    ]:
        if chave not in st.session_state:
            st.session_state[chave] = padrao

    arquivo = st.file_uploader("Planilha da transportadora (CSV ou XLSX)", type=["csv", "xlsx"], key="doccob_upload")

    if arquivo is None:
        _limpar_estado()
        return

    if st.session_state.doccob_nome_original != arquivo.name:
        _limpar_estado()
        sufixo = os.path.splitext(arquivo.name)[1]
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=sufixo)
        tmp.write(arquivo.getbuffer())
        tmp.close()
        st.session_state.doccob_caminho_temp = tmp.name
        st.session_state.doccob_nome_original = arquivo.name

    caminho = st.session_state.doccob_caminho_temp

    if st.session_state.doccob_resultado is None:
        with st.spinner("Analisando planilha..."):
            try:
                formato = detectar_formato(caminho)
                formato_dispatch = "generico (best-effort)" if formato == "desconhecido" else formato
                lotes = montar_lotes_do_arquivo(caminho, formato_dispatch, st.session_state.doccob_respostas)
                todos_avisos = [a for _, avisos in lotes for a in avisos]
                st.session_state.doccob_resultado = {
                    "formato": formato,
                    "lotes": lotes,
                    "pendencias": identificar_pendencias(todos_avisos),
                    "erro": None,
                }
            except Exception as e:
                st.session_state.doccob_resultado = {"erro": f"{type(e).__name__}: {e}"}

    resultado = st.session_state.doccob_resultado

    if resultado.get("erro"):
        st.error(f"Erro ao processar a planilha: {resultado['erro']}")
        return

    if resultado["pendencias"]:
        st.warning(
            "Essa planilha nao traz algumas informacoes necessarias pro DOCCOB. "
            "Responda abaixo (normalmente vem do e-mail da transportadora):"
        )
        with st.form("form_pendencias_doccob"):
            respostas_form = {}
            for pergunta in resultado["pendencias"]:
                campo = pergunta["campo"]
                aviso = pergunta["aviso"]
                rotulo = ROTULOS_CAMPO.get(campo, campo)
                if campo in CAMPOS_DATA:
                    respostas_form[campo] = st.date_input(rotulo, value=_sugestao_data(aviso), key=f"doccob_campo_{campo}")
                else:
                    valor = st.text_input(rotulo, value=_sugestao_texto(aviso), key=f"doccob_campo_{campo}")
                    if valor:
                        respostas_form[campo] = valor
                st.caption(aviso)
            enviado = st.form_submit_button("Confirmar e gerar DOCCOB")
        if enviado:
            st.session_state.doccob_respostas = respostas_form
            st.session_state.doccob_resultado = None
            st.rerun()
        return

    st.success("DOCCOB pronto para download.")
    respostas = st.session_state.doccob_respostas
    for lote, avisos in resultado["lotes"]:
        texto = gerar_doccob(lote)
        numero_fatura = lote.documentos[0].numero_doc_cobranca
        nome_arquivo = f"{int(numero_fatura):09d}_DOCCOB_BRAVIUM.TXT"
        st.download_button(
            label=f"Baixar {nome_arquivo} (fatura {numero_fatura})",
            data=texto.encode("latin-1"),
            file_name=nome_arquivo,
            mime="text/plain",
            key=f"doccob_download_{nome_arquivo}",
        )
        if avisos:
            with st.expander(f"{len(avisos)} aviso(s) da fatura {numero_fatura}"):
                for a in avisos:
                    st.write(f"- {a}")

    # Confirmacao de filial/nome fica salva no cadastro compartilhado - so'
    # uma vez por resultado (nao a cada rerun da pagina).
    if not st.session_state.doccob_cadastro_salvo and resultado["formato"] == "csv_composicao_fretes" \
            and ("filial_552" in respostas or "transportadora_nome" in respostas):
        cnpj_raiz = csv_transportadora.identificar_cnpj_raiz(caminho)
        if cnpj_raiz:
            campos = {"confirmado": True}
            if "filial_552" in respostas:
                campos["filial_552"] = respostas["filial_552"]
            if "transportadora_nome" in respostas:
                campos["nome"] = respostas["transportadora_nome"]
            cadastro.atualizar(cnpj_raiz, **campos)
        st.session_state.doccob_cadastro_salvo = True

    if st.button("Processar outra planilha", key="doccob_reset"):
        _limpar_estado()
        st.rerun()
