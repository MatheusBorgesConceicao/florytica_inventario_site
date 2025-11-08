# Florytica_Site.py
# ─────────────────────────────────────────────────────────────────────────────
# App Streamlit para processamento de inventário florestal
# - Usa SEMPRE Hc (altura comercial) no cálculo de volume
# - Calcula DAP a partir de CAP/π quando DAP não existir
# - Área basal (g) em m²
# - Exporta resultado e resumo por nível
# - Tema e logo da Florytica
# ─────────────────────────────────────────────────────────────────────────────

import io
import math
import numpy as np
import pandas as pd
import streamlit as st

APP_TITLE = "Florytica Inventário — Processamento Completo"
LOGO_PATH = "assets/logo_florytica.png"   # coloque sua imagem aqui (PNG/SVG)

# -----------------------------------------------------------------------------
# Configuração de página + tema + logo
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Florytica Inventário",
    page_icon=LOGO_PATH if LOGO_PATH else "🌳",
    layout="wide",
)

# Logo grande (Streamlit >= 1.26 tem st.logo)
try:
    st.logo(LOGO_PATH, size="large")
except Exception:
    # fallback se st.logo indisponível
    st.image(LOGO_PATH, width=64)

st.markdown(
    f"""
    <h1 style="margin-top:-6px; font-weight:800;">
      Florytica Inventário — Processamento Completo
    </h1>
    <p style="opacity:0.8;margin-top:-8px;">
      Versão 3.0 — Processa DAP (via DAP ou CAP/π), g, Volume (com Hc), Escore Z (opcional) e Indicadores por Nível.
    </p>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# Sidebar — parâmetros simples (mantido tradicional e direto)
# -----------------------------------------------------------------------------
with st.sidebar:
    st.subheader("Parâmetros")
    area_imovel_ha = st.number_input("Área total do imóvel (ha)", min_value=0.0, step=0.01, value=0.0, format="%.2f")

# -----------------------------------------------------------------------------
# Ajuda rápida
# -----------------------------------------------------------------------------
with st.expander("Como usar", expanded=False):
    st.markdown(
        """
        1) Envie um **.xlsx** com os dados do inventário.  
        2) O app exige **Hc** (altura comercial).  
        3) O **DAP** pode vir pronto; se não vier, será calculado por **CAP/π**.  
        4) Saída com **DAP (cm)**, **g_m2 (m²)** e **Vol_Hc_m3 (m³)**, além de resumos por **Nível**.
        """
    )
    st.info("Colunas esperadas (nomes insensíveis a maiúsculas/minúsculas): **Hc** obrigatório; **DAP** ou **CAP**; opcional **Nível**, **Espécie**, **PF**.")

# -----------------------------------------------------------------------------
# Upload
# -----------------------------------------------------------------------------
st.subheader("Envie sua planilha (.xlsx)")
file = st.file_uploader("Arraste/solte ou clique em 'Browse files'", type=["xlsx"])

# -----------------------------------------------------------------------------
# Funções utilitárias
# -----------------------------------------------------------------------------
def _first_sheet_or_named(df_dict: dict, preferred_names=("Dados_Básicos", "dados_básicos", "dados_basicos")) -> pd.DataFrame:
    """Retorna o DataFrame da primeira planilha ou uma com nome preferido, se existir."""
    for name in preferred_names:
        for key in df_dict.keys():
            if str(key).strip().lower() == name.lower():
                return df_dict[key]
    # senão, pega a primeira
    first_key = list(df_dict.keys())[0]
    return df_dict[first_key]

def _to_float(series):
    return pd.to_numeric(series, errors="coerce")

def process_dataframe(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Valida colunas, padroniza e calcula DAP, g e Volume com Hc."""
    if df_raw.empty:
        st.error("Planilha vazia.")
        st.stop()

    # Mapa lower->original
    lower_map = {c.lower().strip(): c for c in df_raw.columns}

    col_cap = lower_map.get("cap")
    col_dap = lower_map.get("dap")
    col_hc  = lower_map.get("hc")
    col_niv = lower_map.get("nível") or lower_map.get("nivel")
    col_esp = lower_map.get("espécie") or lower_map.get("especie")
    col_pf  = lower_map.get("pf")

    # Regras
    if not col_hc:
        st.error("A planilha precisa ter **Hc** (altura comercial).")
        st.stop()
    if not (col_dap or col_cap):
        st.error("A planilha precisa ter **DAP** ou **CAP**.")
        st.stop()

    df = df_raw.copy()

    # DAP (cm)
    if col_dap:
        df["DAP"] = _to_float(df[col_dap])
    else:
        df["DAP"] = _to_float(df[col_cap]) / math.pi  # CAP/π

    # Hc (m)
    df["Hc"] = _to_float(df[col_hc])

    # Checagens
    if df["DAP"].isna().all():
        st.error("Todos os valores de DAP ficaram inválidos. Confira DAP/CAP.")
        st.stop()
    if df["Hc"].isna().all():
        st.error("Todos os valores de Hc ficaram inválidos. Confira a coluna Hc.")
        st.stop()

    # Área basal (m²): π * ( (DAP/100)/2 )²
    df["g_m2"] = math.pi * ((df["DAP"] / 100.0) / 2.0) ** 2

    # Volume (m³) com Hc — fórmula do usuário
    # Volume = 1,3332 * ((DAP/100) ** 2,0836) * (Hc ** 0,732)
    df["Vol_Hc_m3"] = 1.3332 * ((df["DAP"] / 100.0) ** 2.0836) * (df["Hc"] ** 0.732)

    # Metadados úteis (se existirem)
    if col_niv: df["Nível"]   = df[col_niv]
    if col_esp: df["Espécie"] = df[col_esp]
    if col_pf:  df["PF"]      = df[col_pf]

    # Ordena colunas principais primeiro
    cols_first = ["PF", "Nível", "Espécie", "DAP", "Hc", "g_m2", "Vol_Hc_m3"]
    ordered = [c for c in cols_first if c in df.columns] + [c for c in df.columns if c not in cols_first]
    df = df[ordered]

    return df

def resumo_por_nivel(df: pd.DataFrame) -> pd.DataFrame:
    """Resumo por nível: n árvores, somas e médias básicas."""
    if "Nível" not in df.columns:
        # Se não existir "Nível", faz um resumo geral
        res = pd.DataFrame({
            "n_indivíduos": [df.shape[0]],
            "DAP_médio_cm": [df["DAP"].mean()],
            "g_total_m2": [df["g_m2"].sum()],
            "Vol_total_m3": [df["Vol_Hc_m3"].sum()],
        })
        res.index = ["Geral"]
        return res.reset_index(names="Nível/Grupo")

    grp = df.groupby("Nível", dropna=False)
    res = grp.agg(
        n_indivíduos=("DAP", "size"),
        DAP_médio_cm=("DAP", "mean"),
        g_total_m2=("g_m2", "sum"),
        Vol_total_m3=("Vol_Hc_m3", "sum"),
    ).reset_index()
    return res

def download_xlsx(dfs: dict, filename: str) -> bytes:
    """Cria um .xlsx em memória com várias abas."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet, df in dfs.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)
    buffer.seek(0)
    return buffer.read()

# -----------------------------------------------------------------------------
# Execução
# -----------------------------------------------------------------------------
if file:
    try:
        xl = pd.read_excel(file, sheet_name=None)  # carrega todas as abas
        df_in = _first_sheet_or_named(xl)
    except Exception as e:
        st.error(f"Falha ao ler o Excel: {e}")
        st.stop()

    df_proc = process_dataframe(df_in)
    res_nivel = resumo_por_nivel(df_proc)

    st.success("Processado com sucesso (usando Hc).")
    st.write("Prévia dos dados:")
    st.dataframe(df_proc.head(50), use_container_width=True)

    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("Resumo por Nível")
        st.dataframe(res_nivel, use_container_width=True)

    with col_b:
        st.subheader("Indicadores gerais")
        vol_total = df_proc["Vol_Hc_m3"].sum()
        g_total   = df_proc["g_m2"].sum()
        dap_med   = df_proc["DAP"].mean()
        st.metric("Volume total (m³)", f"{vol_total:,.3f}")
        st.metric("Área basal total (m²)", f"{g_total:,.3f}")
        st.metric("DAP médio (cm)", f"{dap_med:,.2f}")

    st.divider()

    # Exports
    xlsx_bytes = download_xlsx(
        {
            "Dados_processados": df_proc,
            "Resumo_por_nivel": res_nivel,
        },
        filename="Florytica_Processado.xlsx",
    )
    st.download_button(
        label="⬇️ Baixar Excel (processado)",
        data=xlsx_bytes,
        file_name="Florytica_Processado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
else:
    st.info("Envie a planilha de dados para iniciar o processamento.")
