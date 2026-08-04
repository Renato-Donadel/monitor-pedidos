"""
Cadastro persistente de transportadoras (nome, codigos de filial), para nao
precisar re-perguntar/re-descobrir a cada arquivo processado.

CHAVE DO CADASTRO = RAIZ DO CNPJ (8 primeiros digitos), NUNCA o codigo/nome
que aparece no arquivo recebido (ex.: "OTC", "FAV") - esse rotulo e' so' do
formato do arquivo/template de quem enviou, nao identifica a empresa de
verdade (a mesma transportadora pode mandar arquivos com rotulos
diferentes, e o rotulo sozinho nao e' garantia de nada).

Guardado em data/doccob_cadastro_transportadoras.json - simples o
suficiente para editar a mao quando alguem confirmar um dado (ex.: a razao
social oficial, ou o codigo de filial correto).

ATENCAO (deploy no Streamlit Community Cloud): o sistema de arquivos la e'
EFEMERO - esse JSON some quando o app reinicia/redeploya. Localmente (rodando
o .bat) a persistencia funciona normal; no site hospedado, o cadastro so'
"lembra" enquanto o container atual estiver de pe.
"""

import json
import os

CAMINHO_CADASTRO = os.path.normpath(os.path.join(
    os.path.dirname(__file__), "..", "data", "doccob_cadastro_transportadoras.json"
))


def raiz_cnpj(cnpj: str) -> str | None:
    digitos = "".join(ch for ch in str(cnpj) if ch.isdigit())
    return digitos[:8] if len(digitos) >= 8 else None


def carregar() -> dict:
    if not os.path.exists(CAMINHO_CADASTRO):
        return {}
    with open(CAMINHO_CADASTRO, "r", encoding="utf-8") as f:
        return json.load(f)


def salvar(cadastro: dict):
    os.makedirs(os.path.dirname(CAMINHO_CADASTRO), exist_ok=True)
    tmp = CAMINHO_CADASTRO + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cadastro, f, ensure_ascii=False, indent=2, sort_keys=True)
    os.replace(tmp, CAMINHO_CADASTRO)


def obter_ou_criar(cnpj_raiz: str, cnpj_completo: str = None, codigo_visto: str = None) -> dict:
    """`cnpj_raiz` = 8 primeiros digitos do CNPJ da transportadora - a
    chave real do cadastro. Se ja existir, devolve o que ja se sabe
    (podendo ter campos None ainda nao confirmados). Se nao existir, cria
    uma entrada nova. `cnpj_completo`/`codigo_visto` sao so' guardados como
    informacao (quais CNPJs completos e rotulos de arquivo ja apareceram
    pra essa raiz) - nunca usados pra identificar."""
    cadastro = carregar()
    if cnpj_raiz not in cadastro:
        cadastro[cnpj_raiz] = {
            "nome": None,
            "filial_552": None,
            "filial_555": None,
            "confirmado": False,
            "codigos_vistos": [],
            "cnpjs_completos_vistos": [],
        }

    entrada = cadastro[cnpj_raiz]
    mudou = False
    if codigo_visto and codigo_visto not in entrada.setdefault("codigos_vistos", []):
        entrada["codigos_vistos"].append(codigo_visto)
        mudou = True
    if cnpj_completo and cnpj_completo not in entrada.setdefault("cnpjs_completos_vistos", []):
        entrada["cnpjs_completos_vistos"].append(cnpj_completo)
        mudou = True
    if mudou:
        salvar(cadastro)
    return entrada


def atualizar(cnpj_raiz: str, **campos):
    """Atualiza/confirma campos de uma transportadora ja cadastrada (chave
    = raiz do CNPJ), ex.:
    atualizar('33070814', nome='CARVALIMA TRANSPORTES LTDA', confirmado=True)."""
    cadastro = carregar()
    cadastro.setdefault(cnpj_raiz, {})
    cadastro[cnpj_raiz].update(campos)
    salvar(cadastro)
