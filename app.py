import streamlit as st
import pandas as pd
import PyPDF2
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

st.set_page_config(page_title="AuditAI - IPEM/RJ", layout="wide", page_icon="🛡️")

def extrair_dados_pdf(file):
    texto_completo = ""
    reader = PyPDF2.PdfReader(file)
    paginas = []
    for page in reader.pages:
        content = page.extract_text()
        if content:
            texto_completo += content
            paginas.append(content)
    
    texto_limpo = " ".join(texto_completo.split())

    # --- FUNÇÃO INTERNA PARA BUSCAR SEI POR PALAVRA-CHAVE ---
    def buscar_sei_por_contexto(termo_chave):
        for p in paginas:
            if termo_chave.lower() in p.lower():
                match = re.search(r"(?:verificador|Documento\s+SEI)\s+(\d{8,10})", p, re.IGNORECASE)
                if match: return match.group(1)
        return "Verificar no SEI"

    # --- EXTRAÇÃO DE DADOS FINANCEIROS (ITEM 1) ---
    id_nl = "Não encontrada"
    id_ne = "Não encontrada"
    match_nl = re.search(r"202\dNL\d{5}", texto_limpo)
    if match_nl: id_nl = match_nl.group(0)
    match_ne = re.search(r"202\dNE\d{5}", texto_limpo)
    if match_ne: id_ne = match_ne.group(0)

    # --- EXTRAÇÃO INDIVIDUAL DE DOCUMENTOS SEI ---
    return {
        "processo": re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo).group(0) if re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo) else "Não encontrado",
        "empenho": id_ne,
        "liquidacao": id_nl,
        "sei_nota_fiscal": buscar_sei_por_contexto("Nota Fiscal"),
        "sei_federal": buscar_sei_por_contexto("Certidão Negativa de Débitos Relativos a Créditos Tributários Federais"),
        "sei_fgts": buscar_sei_por_contexto("Certificado de Regularidade do FGTS"),
        "sei_trabalhista": buscar_sei_por_contexto("Certidão Negativa de Débitos Trabalhistas"),
        "sei_gestor": buscar_sei_por_contexto("Atesto")
    }

# --- INTERFACE ---
st.title("🛡️ AuditAI - IPEM/RJ")
uploaded_file = st.sidebar.file_uploader("Upload do Processo (PDF)", type="pdf")

if uploaded_file:
    d = extrair_dados_pdf(uploaded_file)
    
    obs_1 = f"{d['empenho']} - Gerando a {d['liquidacao']}"
    
    checklist = [
        {"ITEM": 1, "EVENTO": "Nota de empenho e demonstrativo de saldo", "S/N/NA": "S", "OBSERVAÇÕES": obs_1},
        {"ITEM": 2, "EVENTO": "Nota Fiscal / Fatura", "S/N/NA": "S", "OBSERVAÇÕES": f"Documento SEI {d['sei_nota_fiscal']}"},
        {"ITEM": 3, "EVENTO": "Certidão Federal e Dívida Ativa", "S/N/NA": "S", "OBSERVAÇÕES": f"Documento SEI {d['sei_federal']}"},
        {"ITEM": 4, "EVENTO": "Certidão de FGTS", "S/N/NA": "S", "OBSERVAÇÕES": f"Documento SEI {d['sei_fgts']}"},
        {"ITEM": 5, "EVENTO": "Certidão de Justiça do Trabalho", "S/N/NA": "S", "OBSERVAÇÕES": f"Documento SEI {d['sei_trabalhista']}"},
        {"ITEM": 13, "EVENTO": "Atestado do Gestor", "S/N/NA": "S", "OBSERVAÇÕES": f"Documento SEI {d['sei_gestor']}"},
    ]
    
    df = pd.DataFrame(checklist)
    st.table(df)
    st.download_button("📥 Baixar Checklist", io.BytesIO().getvalue(), "checklist.xlsx")
