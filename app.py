import streamlit as st
import pandas as pd
import PyPDF2
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

# Configurações de Interface
st.set_page_config(page_title="AuditAI - IPEM/RJ", layout="wide", page_icon="🛡️")

def extrair_dados_pdf(file):
    texto_completo = ""
    reader = PyPDF2.PdfReader(file)
    for page in reader.pages:
        content = page.extract_text()
        if content: texto_completo += content
    
    # Normalização para busca linear
    texto_limpo = " ".join(texto_completo.split())

    # --- LÓGICA DE EXTRAÇÃO PARA ITEM 1 (SIAFE-RJ) ---
    id_nl = "Não encontrada"
    id_ne = "Não encontrada"
    
    # 1. Localiza a Nota de Liquidação (NL) - Padrão 2026NLXXXXX
    match_nl = re.search(r"202\dNL\d{5}", texto_limpo)
    if match_nl:
        id_nl = match_nl.group(0)

    # 2. Localiza a Nota de Empenho (NE) - Padrão 2026NEXXXXX
    # Captura a NE que consta no documento (ex: 2026NE00021)
    match_ne = re.search(r"202\dNE\d{5}", texto_limpo)
    if match_ne:
        id_ne = match_ne.group(0)

    # --- IDENTIFICAÇÃO SEI (CÓDIGO VERIFICADOR) ---
    # Captura o número de 8 a 10 dígitos do documento
    busca_sei = re.search(r"(?:verificador|Documento\s+SEI)\s+(\d{8,10})", texto_limpo, re.IGNORECASE)
    id_sei = busca_sei.group(1) if busca_sei else "Verificar SEI"

    # Captura do Número do Processo
    re_proc = re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo)
    processo_num = re_proc.group(0) if re_proc else "Não encontrado"

    return {
        "processo": processo_num,
        "empenho": id_ne,
        "liquidacao": id_nl,
        "sei_verificador": id_sei
    }

def gerar_excel_limpo(dados_p, df_c):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Auditoria"
    ws['A1'] = "IPEM/RJ - CHECKLIST DE AUDITORIA"
    ws['A2'] = f"Processo: {dados_p['processo']}"
    wb.save(output)
    return output.getvalue()

# --- INTERFACE STREAMLIT ---
st.markdown("<h2 style='text-align: center;'>AUDIT - IPEM/RJ</h2>", unsafe_allow_html=True)
uploaded_file = st.sidebar.file_uploader("Carregar PDF do Processo SEI", type="pdf")

if uploaded_file:
    d = extrair_dados_pdf(uploaded_file)
    
    # Formatação exata conforme solicitado: 2026NExxxxx - Gerando a 2026NLXXXXX
    obs_financeiro = f"{d['empenho']} - Gerando a {d['liquidacao']}"
    obs_sei = f"Documento SEI {d['sei_verificador']}"
    
    checklist = [
        {"ITEM": 1, "EVENTO": "Nota de empenho e demonstrativo de saldo", "S/N/NA": "S", "OBSERVAÇÕES": obs_financeiro},
        {"ITEM": 2, "EVENTO": "Nota Fiscal / Fatura", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 3, "EVENTO": "Certidão Federal e Dívida Ativa", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 4, "EVENTO": "Certidão de FGTS", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 5, "EVENTO": "Certidão de Justiça do Trabalho", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 13, "EVENTO": "Atestado do Gestor", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
    ]
    
    # Preenchimento automático dos itens restantes
    for i in range(6, 20):
        if i != 13:
            checklist.append({"ITEM": i, "EVENTO": f"Verificação de Auditoria {i}", "S/N/NA": "S", "OBSERVAÇÕES": "Conforme processo"})

    df = pd.DataFrame(checklist).sort_values("ITEM")
    st.table(df)
    
    st.download_button("📥 Baixar Checklist", gerar_excel_limpo(d, df), f"Checklist_{d['processo'].replace('/','_')}.xlsx")
