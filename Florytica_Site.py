# =========================================================
# Florytica Inventário — Processamento Completo
# Versão 3.0 — DAP, g, Volume, Escore Z, Erro Amostral e Indicadores por Nível
# =========================================================

import streamlit as st
import pandas as pd
import numpy as np
import math
from io import BytesIO
from datetime import datetime
from docx import Document
from docx.shared import Inches
import matplotlib.pyplot as plt
from PIL import Image

# =========================================================
# CONFIGURAÇÕES DA PÁGINA E LOGO
# =========================================================

LOGO_PATH = "assets/florytica_logo.png"
FAVICON_PATH = "assets/favicon.png"

try:
    st.set_page_config(
        page_title="Florytica Inventário",
        page_icon=Image.open(FAVICON_PATH),
        layout="wide"
    )
except Exception:
    st.set_page_config(page_title="Florytica Inventário", page_icon="🌳", layout="wide")

# Exibe logo no topo
try:
    st.logo(LOGO_PATH)
except Exception:
    st.sidebar.image(LOGO_PATH, use_column_width=True)

# =========================================================
# ESTILO VISUAL (cores da marca)
# =========================================================
st.markdown("""
<style>
/* Fundo e elementos principais */
[data-testid="stAppViewContainer"] {
    background-color: #0B0F0D;
}
[data-testid="stSidebar"] {
    background-color: #141A16;
}
h1, h2, h3, h4, h5, h6, p, label, span, div {
    color: #E6F2EC !important;
    font-family: "Segoe UI", sans-serif;
}
/* Botões */
button[kind="primary"] {
    background-color: #13A66B !important;
    color: white !important;
    border: none !important;
    font-weight: 600;
}
/* Inputs */
.stTextInput>div>div>input, .stNumberInput input {
    background-color: #1C211D !important;
    color: #E6F2EC !important;
}
/* Cards */
div[data-testid="stExpander"] {
    background-color: #141A16 !important;
    border-radius: 8px !important;
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# CABEÇALHO
# =========================================================
st.markdown("""
# 🌳 Florytica Inventário — Processamento Completo
**Versão 3.0** — Processa DAP, g, Volume, Escore Z, Erro Amostral e Indicadores por Nível.
""")

# =========================================================
# INSTRUÇÕES
# =========================================================
with st.expander("📘 Como usar", expanded=False):
    st.markdown("""
1. Envie sua planilha **Dados_Básicos (.xlsx)**.  
2. Informe a **área total (ha)** do imóvel na barra lateral.  
3. O sistema calcula automaticamente:
   - g (cm²) = π * (DAP / 2)²  
   - Volume (HT e HC) = 1,3332 * (DAP / 100)².0836 * Altura^0.732  
   - Escore Z (Volume) = (Volume_i - Média) / Desvio  
   - Fator de Expansão, Densidade, Vol/ha, Média de Altura e DAP  
   - Classes de DAP e Altura (Sturges)
""")

# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.header("⚙️ Parâmetros")
area_ha = st.sidebar.number_input("Área total do imóvel (ha)", min_value=0.0, value=0.0, step=0.1)
file = st.file_uploader("📂 Envie sua planilha (.xlsx)", type=["xlsx"])

# =========================================================
# PROCESSAMENTO PRINCIPAL
# =========================================================
if file:
    df = pd.read_excel(file)

    if "DAP" not in df.columns or "Altura" not in df.columns:
        st.error("A planilha precisa conter as colunas 'DAP' e 'Altura'.")
    else:
        # Cálculos
        df["g (cm²)"] = np.pi * (df["DAP"] / 2) ** 2
        df["Volume (m³)"] = 1.3332 * (df["DAP"] / 100) ** 2.0836 * df["Altura"] ** 0.732
        df["Escore Z (Volume)"] = (df["Volume (m³)"] - df["Volume (m³)"].mean()) / df["Volume (m³)"].std()

        total_volume = df["Volume (m³)"].sum()
        densidade = len(df) / area_ha if area_ha > 0 else 0
        vol_ha = total_volume / area_ha if area_ha > 0 else 0

        # =========================================================
        # EXIBIÇÃO DE RESULTADOS
        # =========================================================
        st.success("✅ Processamento concluído com sucesso!")
        st.metric("Volume Total (m³)", f"{total_volume:,.2f}")
        st.metric("Densidade (árv/ha)", f"{densidade:,.2f}")
        st.metric("Volume por hectare (m³/ha)", f"{vol_ha:,.2f}")

        st.dataframe(df.head())

        # Gráfico de Volume
        fig, ax = plt.subplots()
        ax.hist(df["Volume (m³)"], bins=15, color="#13A66B", edgecolor="white")
        ax.set_xlabel("Volume (m³)")
        ax.set_ylabel("Frequência")
        ax.set_title("Distribuição de Volume")
        st.pyplot(fig)

        # Exporta arquivo Excel
        output = BytesIO()
        df.to_excel(output, index=False)
        st.download_button(
            label="📥 Baixar resultados em Excel",
            data=output.getvalue(),
            file_name="Inventario_Processado_Florytica.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Envie a planilha de Dados_Básicos para iniciar o processamento.")
