import streamlit as st
import pandas as pd
import PyPDF2
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

st.set_page_config(page_title="AuditAI - IPEM/RJ", layout="wide", page_icon="🛡️")

def extrair_dados_pdf(file):
    reader = PyPDF2.PdfReader(file)
    paginas_texto = []
    for page in reader.pages:
        content = page.extract_text()
        if content:
            paginas_texto.append(content)
    
    texto_completo = " ".join(paginas_texto)
    texto_limpo = " ".join(texto_completo.split())

    # --- FUNÇÃO DE BUSCA POR CONTEXTO (PÁGINA A PÁGINA) ---
    def buscar_sei_especifico(termos_chave):
        for texto_pagina in paginas_texto:
            # Verifica se algum dos termos (Ex: 'Nota Fiscal', 'Fatura') está na página
            if any(termo.lower() in texto_pagina.lower() for termo in termos_chave):
                # Busca o código verificador nesta página específica
                match = re.search(r"verificador\s+(\d{8,10})", texto_pagina, re.IGNORECASE)
                if match:
                    return match.group(1)
        return "Não identificado"

    # --- EXTRAÇÃO ITEM 1 (FINANCEIRO) ---
    id_nl = "Não encontrada"
    id_ne = "Não encontrada"
    match_nl = re.search(r"202\dNL\d{5}", texto_limpo)
    if match_nl: id_nl = match_nl.group(0)
    match_ne = re.search(r"202\dNE\d{5}", texto_limpo)
    if match_ne: id_ne = match_ne.group(0)

    # --- MAPEAMENTO INDIVIDUAL DOS DOCUMENTOS ---
    return {
        "processo": re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo).group(0) if re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo) else "Não encontrado",
        "empenho": id_ne,
        "liquidacao": id_nl,
        # Item 2: Nota Fiscal
        "sei_nf": buscar_sei_especifico(["Nota Fiscal", "DANFE", "Fatura"]),
        # Item 3: Certidão Federal
        "sei_federal": buscar_sei_especifico(["Receita Federal", "Créditos Tributários Federais", "Dívida Ativa da União"]),
        # Item 4: FGTS
        "sei_fgts": buscar_sei_especifico(["FGTS", "Fundo de Garantia", "CRF"]),
        # Item 5: Trabalhista
        "sei_trabalhista": buscar_sei_especifico(["Trabalhista", "CNDT", "Justiça do Trabalho"]),
        # Item 13: Atesto
        "sei_atesto": buscar_sei_especifico(["Atesto", "Atestamos", "fatura foi conferida"])
    }

# --- INTERFACE ---
st.title("🛡️ AuditAI - IPEM/RJ")
uploaded_file = st.sidebar.file_uploader("Upload do Processo (PDF)", type="pdf")

if uploaded_file:
    d = extrair_dados_pdf(uploaded_file)
    
    obs_1 = f"{d['empenho']} - Gerando a {d['liquidacao']}"
    
    checklist = [
        {"ITEM": 1, "EVENTO": "Nota de empenho e demonstrativo de saldo", "S/N/NA": "S", "OBSERVAÇÕES": obs_1},
        {"ITEM": 2, "EVENTO": "Nota Fiscal / Fatura", "S/N/NA": "S", "OBSERVAÇÕES": f"Documento SEI {d['sei_nf']}"},
        {"ITEM": 3, "EVENTO": "Certidão Federal e Dívida Ativa", "S/N/NA": "S", "OBSERVAÇÕES": f"Documento SEI {d['sei_federal']}"},
        {"ITEM": 4, "EVENTO": "Certidão de FGTS", "S/N/NA": "S", "OBSERVAÇÕES": f"Documento SEI {d['sei_fgts']}"},
        {"ITEM": 5, "EVENTO": "Certidão de Justiça do Trabalho", "S/N/NA": "S", "OBSERVAÇÕES": f"Documento SEI {d['sei_trabalhista']}"},
        {"ITEM": 13, "EVENTO": "Atestado do Gestor", "S/N/NA": "S", "OBSERVAÇÕES": f"Documento SEI {d['sei_atesto']}"},
    ]
    
    st.table(pd.DataFrame(checklist))
    st.download_button("📥 Baixar Checklist", io.BytesIO().getvalue(), "checklist.xlsx")
