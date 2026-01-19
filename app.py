import streamlit as st
import pandas as pd
import os
from io import BytesIO

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
# LEITURA BASE
# ==============================
if not os.path.exists(BASE_MONITOR):
    st.error("Arquivo Monitor_Pedidos_Processado.xlsx não encontrado.")
    st.stop()

df = pd.read_excel(BASE_MONITOR)

# garante ordenação absoluta
df = df.sort_values("Ranking").reset_index(drop=True)

# ==============================
# INTERFACE
# ==============================
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

    if st.button(f"Baixar próximos 300 — {carteira}", key=carteira):
        df_lote = df_carteira.head(TAMANHO_LOTE)

        nome_arquivo = f"Pedidos_{carteira}.xlsx"

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
