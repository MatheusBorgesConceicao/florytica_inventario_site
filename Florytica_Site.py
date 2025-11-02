
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import math
from datetime import datetime
from docx import Document
from docx.shared import Inches
import matplotlib.pyplot as plt

st.set_page_config(page_title="Florytica Inventário — Etapa 3", layout="wide")

st.title("🌳 Florytica Inventário — Processamento Completo")
st.write("Versão 3.0 — Processa DAP, g, Volume, Escore Z, Erro Amostral e Indicadores por Nível.")

with st.expander("📘 Como usar", expanded=False):
    st.markdown("""
1. Envie sua planilha **Dados_Básicos** (.xlsx).  
2. Informe a **área total (ha)** do imóvel na barra lateral.  
3. O sistema processa automaticamente:
   - **DAP (cm)** = CAP / π  
   - **g (m²)** = π × (DAP/200)²  
   - **Volumes (HT e HC)** = 1.3332 × (DAP/100)^2.0836 × Altura^0.732  
   - **Escore Z (Volume)** = (Média(volume) − Volume_linha) / Soma(volume)  
   - Fator de Expansão, Ind.ha⁻¹, g.ha⁻¹, Vol m³.ha  
   - Classes de DAP e Altura (Sturges)
    """)

# =============== Sidebar ===============
st.sidebar.header("⚙️ Parâmetros")
area_ha = st.sidebar.number_input("Área total do imóvel (ha)", min_value=0.0, value=0.0, step=0.1)

file = st.file_uploader("📂 Envie sua planilha (.xlsx)", type=["xlsx"])

if file:
    df = pd.read_excel(file)

    if "Nivel" not in df.columns or "CAP" not in df.columns:
        st.error("As colunas 'Nivel' e 'CAP' são obrigatórias.")
        st.stop()

    df["Nivel"] = df["Nivel"].astype(str).str.strip().str.upper()
    df["CAP"] = pd.to_numeric(df["CAP"], errors="coerce")

    # =============== Cálculos principais ===============
    def vol_formula(dap_cm, altura):
        if pd.notna(dap_cm) and pd.notna(altura) and altura > 0:
            return 1.3332 * ((float(dap_cm) / 100.0) ** 2.0836) * (float(altura) ** 0.732)
        return np.nan

    df["DAP"] = df["CAP"] / np.pi
    df["g (m²)"] = np.pi * (df["DAP"] / 200.0) ** 2

    if "HT" in df.columns:
        df["Vol. HT"] = df.apply(lambda r: vol_formula(r["DAP"], r["HT"]), axis=1)
    else:
        df["Vol. HT"] = np.nan

    if "HC" in df.columns:
        df["Vol. HC"] = df.apply(lambda r: vol_formula(r["DAP"], r["HC"]), axis=1)
    else:
        df["Vol. HC"] = np.nan

    df["Vol_base"] = df["Vol. HT"].where(df["Vol. HT"].notna(), df["Vol. HC"])

    # =============== Fator de Expansão ===============
    AREA_PARC = {"S2": 10*50, "S1": 10*10, "R3": 5*5, "R2": 2*2, "R1": 1*1}
    n_parc = df.groupby("Nivel")["PF"].nunique() if "PF" in df.columns else df["Nivel"].value_counts()

    def fator_exp(nivel):
        ap = AREA_PARC.get(nivel, None)
        n = n_parc.get(nivel, 0)
        if ap and n > 0:
            return 10000.0 / (ap * n)
        return np.nan

    df["FE"] = df["Nivel"].map(fator_exp)

    # =============== Índices por hectare ===============
    col_cont_regen = next((c for c in ["Nº de Ind.", "Num_Ind", "Qtde", "Quantidade"] if c in df.columns), None)

    def ind_ha(row):
        nv = row["Nivel"]
        fe = row["FE"]
        if not np.isfinite(fe):
            return np.nan
        if nv in ("R1", "R2", "R3") and col_cont_regen:
            q = pd.to_numeric(row.get(col_cont_regen), errors="coerce")
            return (q if pd.notna(q) else 0) * fe
        return fe

    df["Ind.ha-1"] = df.apply(ind_ha, axis=1)
    df["g (m²).ha-1"] = df["g (m²)"] * df["FE"]
    df["Vol m³.ha"] = df["Vol_base"] * df["FE"]

    # =============== Classes DAP e Altura (Sturges) ===============
    mask_s2 = df["Nivel"] == "S2"

    def class_sturges(serie, minimo):
        serie = pd.to_numeric(serie, errors="coerce").dropna()
        if len(serie) < 2:
            return None
        k = max(1, round(1 + 3.322 * math.log10(len(serie))))
        xmin, xmax = max(minimo, float(serie.min())), float(serie.max())
        if xmax <= xmin:
            xmax = xmin + 1e-6
        width = (xmax - xmin) / k
        bins = np.arange(xmin, xmax + width, width)
        return bins

    dap_bins = class_sturges(df.loc[mask_s2, "DAP"], minimo=10)
    alt_col = "HT" if "HT" in df.columns else "HC"
    ht_bins = class_sturges(df.loc[mask_s2, alt_col], minimo=1.5)

    df["Classe_DAP"] = pd.cut(df["DAP"], bins=dap_bins, include_lowest=True) if dap_bins is not None else "-"
    df["Classe_H"] = pd.cut(df[alt_col], bins=ht_bins, include_lowest=True) if ht_bins is not None else "-"

    # =============== Escore Z (base Volume) ===============
    vol = pd.to_numeric(df["Vol_base"], errors="coerce")
    media_vol = vol.mean(skipna=True)
    soma_vol = vol.sum(skipna=True)
    df["Escore Z"] = (media_vol - vol) / soma_vol if soma_vol not in (0, np.nan) else np.nan

    # =============== Indicadores por nível ===============
    niveis = df["Nivel"].unique()
    linhas = []

    for nv in sorted(niveis):
        sub = df[df["Nivel"] == nv]
        fe = sub["FE"].dropna().unique()
        fe_val = fe[0] if len(fe) > 0 else np.nan

        if nv in ["R1", "R2", "R3"] and col_cont_regen:
            n_amostrado = pd.to_numeric(sub[col_cont_regen], errors="coerce").fillna(0).sum()
        else:
            n_amostrado = len(sub)

        g_amostrada = sub["g (m²)"].sum(skipna=True)
        vol_amostrado = sub["Vol_base"].sum(skipna=True)

        ind_ha = n_amostrado * fe_val if np.isfinite(fe_val) else np.nan
        g_ha = g_amostrada * fe_val if np.isfinite(fe_val) else np.nan
        vol_ha = vol_amostrado * fe_val if np.isfinite(fe_val) else np.nan

        ind_total = ind_ha * area_ha if (area_ha > 0 and np.isfinite(ind_ha)) else np.nan
        g_total = g_ha * area_ha if (area_ha > 0 and np.isfinite(g_ha)) else np.nan
        vol_total = vol_ha * area_ha if (area_ha > 0 and np.isfinite(vol_ha)) else np.nan

        linhas.append({
            "Nível": nv,
            "Parcelas": int(n_parc.get(nv, 0)),
            "Fator (ha⁻¹)": round(fe_val, 4) if np.isfinite(fe_val) else "-",
            "Ind.ha⁻¹": round(ind_ha, 2) if np.isfinite(ind_ha) else "-",
            "g.ha⁻¹ (m²)": round(g_ha, 4) if np.isfinite(g_ha) else "-",
            "Vol m³.ha": round(vol_ha, 4) if np.isfinite(vol_ha) else "-",
            "Ind. estimado (área)": round(ind_total, 0) if np.isfinite(ind_total) else "-",
            "g estimado (m² área)": round(g_total, 2) if np.isfinite(g_total) else "-",
            "Vol estimado (m³ área)": round(vol_total, 2) if np.isfinite(vol_total) else "-"
        })

    tabela_niveis = pd.DataFrame(linhas)
    st.subheader("📊 Indicadores por nível")
    st.dataframe(tabela_niveis, use_container_width=True)

    # =============== Gráfico DAP (S2) ===============
    img_buf = None
    if mask_s2.any() and df.loc[mask_s2, "DAP"].notna().any():
        fig, ax = plt.subplots()
        ax.hist(df.loc[mask_s2, "DAP"].dropna(), bins=10, color="forestgreen", edgecolor="black")
        ax.set_xlabel("DAP (cm)")
        ax.set_ylabel("Frequência")
        ax.set_title("Distribuição de DAP — Nível S2")
        st.pyplot(fig)
        img_buf = BytesIO()
        fig.savefig(img_buf, format="png", bbox_inches="tight", dpi=150)
        img_buf.seek(0)

    # =============== Exportação ===============
    excel_buf = BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Processado")
        tabela_niveis.to_excel(w, index=False, sheet_name="Indicadores_Nivel")
    st.download_button("💾 Baixar Excel Processado", excel_buf.getvalue(), "Inventario_Processado.xlsx")

    # =============== Relatório Word ===============
    def gerar_relatorio(df, grafico_buf):
        doc = Document()
        doc.add_heading("Relatório de Processamento Florestal", level=1)
        doc.add_paragraph(f"Data de geração: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        doc.add_paragraph("Sistema: Florytica Inventário Automático")
        doc.add_paragraph(" ")

        doc.add_heading("Resumo (amostra)", level=2)
        tabela = doc.add_table(rows=1, cols=6)
        hdr = tabela.rows[0].cells
        hdr[0].text = "PF"
        hdr[1].text = "Nível"
        hdr[2].text = "DAP (cm)"
        hdr[3].text = "g (m²)"
        hdr[4].text = "Vol (m³)"
        hdr[5].text = "Escore Z"

        for _, r in df.head(30).iterrows():
            row = tabela.add_row().cells
            row[0].text = str(r.get("PF", ""))
            row[1].text = str(r.get("Nivel", ""))
            row[2].text = f"{r.get('DAP', np.nan):.2f}" if pd.notna(r.get("DAP")) else "-"
            row[3].text = f"{r.get('g (m²)', np.nan):.4f}" if pd.notna(r.get('g (m²)')) else "-"
            row[4].text = f"{r.get('Vol_base', np.nan):.4f}" if pd.notna(r.get('Vol_base')) else "-"
            row[5].text = f"{r.get('Escore Z', np.nan):.4f}" if pd.notna(r.get('Escore Z')) else "-"

        doc.add_paragraph(" ")
        doc.add_heading("Gráfico de Distribuição de DAP (S2)", level=2)
        if grafico_buf:
            doc.add_picture(grafico_buf, width=Inches(5.5))
        else:
            doc.add_paragraph("Gráfico não disponível.")

        out = BytesIO()
        doc.save(out)
        return out.getvalue()

    if st.button("📄 Gerar Relatório Word"):
        word_bytes = gerar_relatorio(df, img_buf)
        st.download_button("⬇️ Baixar Relatório Word", word_bytes, "Relatorio_Florytica.docx")

else:
    st.info("Envie a planilha de Dados_Básicos para iniciar o processamento.")
