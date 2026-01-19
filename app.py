import streamlit as st
import pandas as pd
import os
from io import BytesIO
import hashlib

# ==============================
# LOGIN SIMPLES (1x por sessão)
# ==============================
SENHA_APP = "8S15?w5fkP"

if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    st.title("🔒 Acesso restrito")
    senha = st.text_input("Digite a senha para acessar", type="password")

    if st.button("Entrar"):
        if senha == SENHA_APP:
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Senha incorreta.")
    st.stop()



# ==============================
# CONFIGURAÇÕES
# ==============================
BASE_MONITOR = os.path.join(
    os.path.dirname(__file__),
    "data",
    "Monitor_Pedidos_Processado.xlsx"
)

TAMANHO_LOTE = 300

st.set_page_config(
    page_title="Monitor de Pedidos Críticos",
    layout="centered"
)

# ==============================
# FUNÇÕES
# ==============================
def file_hash(path: str) -> str:
    """Gera um hash do arquivo para detectar atualização."""
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()

# ==============================
# LEITURA BASE
# ==============================
if not os.path.exists(BASE_MONITOR):
    st.error("Arquivo Monitor_Pedidos_Processado.xlsx não encontrado.")
    st.stop()

# hash para detectar se o Excel foi atualizado
arquivo_hash = file_hash(BASE_MONITOR)

df = pd.read_excel(BASE_MONITOR)
df = df.sort_values("Ranking").reset_index(drop=True)

# ==============================
# SESSION STATE (offset + versão do arquivo)
# ==============================
if "offsets" not in st.session_state:
    st.session_state["offsets"] = {}

if "arquivo_hash" not in st.session_state:
    st.session_state["arquivo_hash"] = arquivo_hash

# se o arquivo mudou, zera offsets
if st.session_state["arquivo_hash"] != arquivo_hash:
    st.session_state["offsets"] = {}
    st.session_state["arquivo_hash"] = arquivo_hash
    st.info("📌 Base atualizada! Os lotes foram reiniciados do começo.")

# ==============================
# INTERFACE
# ==============================

senha = st.text_input("🔒 Digite a senha para acessar", type="password")

if senha != SENHA_APP:
    st.warning("Acesso restrito. Digite a senha correta.")
    st.stop()

st.title("📦 Monitor de Pedidos Críticos")
st.write(
    "Clique no botão da sua carteira para baixar **300 pedidos por vez**, "
    "em ordem de criticidade."
)

carteiras = sorted(df["Carteira"].dropna().unique())

st.divider()

for carteira in carteiras:
    st.subheader(f"📁 {carteira}")

    df_carteira = df[df["Carteira"] == carteira].reset_index(drop=True)

    total = len(df_carteira)

    if total == 0:
        st.caption("Sem pedidos nessa carteira.")
        st.divider()
        continue

    offset_atual = st.session_state["offsets"].get(carteira, 0)

    inicio = offset_atual
    fim = min(offset_atual + TAMANHO_LOTE, total)

    st.caption(f"Total pedidos: {total} | Próximo lote: {inicio+1} até {fim}")

    if st.button(f"Baixar próximos 300 — {carteira}", key=f"baixar_{carteira}"):
        df_lote = df_carteira.iloc[inicio:fim]

        # se já chegou no fim, trava (não volta pro início)
        if df_lote.empty:
            st.warning("✅ Você já baixou todos os pedidos dessa carteira. Aguarde a próxima atualização da base.")
            st.stop()

        # atualiza offset para o próximo clique
        novo_offset = fim
        st.session_state["offsets"][carteira] = novo_offset

        nome_arquivo = f"Pedidos_{carteira}_{inicio+1}_a_{fim}.xlsx"

        buffer = BytesIO()
        df_lote.to_excel(buffer, index=False)
        buffer.seek(0)

        st.download_button(
            label="📥 Clique aqui para baixar o Excel",
            data=buffer,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.divider()
