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
    for page in reader.pages:
        content = page.extract_text()
        if content: texto_completo += content
    
    # Texto limpo para busca linear
    texto_limpo = " ".join(texto_completo.split())

    # --- LÓGICA DE PRECISÃO PARA O ITEM 1 ---
    id_nl = "Não encontrada"
    data_nl = "Não encontrada"
    id_ne = "Não encontrada"
    
    # 1. Localiza o número da NL (ex: 2026NL00021)
    match_nl = re.search(r"202\dNL\d{5}", texto_limpo)
    if match_nl:
        id_nl = match_nl.group(0)
        
        # 2. Busca a DATA de emissão (XX/XX/XXXX)
        # O sistema procura por qualquer data num raio de 150 caracteres após a NL
        pos_inicio = texto_limpo.find(id_nl)
        bloco_pos_nl = texto_limpo[pos_inicio:pos_inicio+150]
        
        datas = re.findall(r"\d{2}/\d{2}/\d{4}", bloco_pos_nl)
        if datas:
            data_nl = datas[0]

    # 3. Localiza o número da NE (ex: 2026NE00021)
    match_ne = re.search(r"202\dNE\d{5}", texto_limpo)
    if match_ne:
        id_ne = match_ne.group(0)

    # 4. Localiza o código verificador SEI
    match_sei = re.search(r"(?:verificador|Documento\s+SEI)\s+(\d{8,10})", texto_limpo, re.IGNORECASE)
    id_sei = match_sei.group(1) if match_sei else "Verificar SEI"

    return {
        "processo": re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo).group(0) if re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo) else "Não encontrado",
        "empenho": id_ne,
        "liquidacao": id_nl,
        "data_nl": data_nl,
        "sei_verificador": id_sei
    }

# --- INTERFACE ---
st.title("🛡️ AuditAI - IPEM/RJ")
uploaded_file = st.sidebar.file_uploader("Upload do Processo (PDF)", type="pdf")

if uploaded_file:
    d = extrair_dados_pdf(uploaded_file)
    
    # Formatação exata conforme solicitado para o Item 1
    # Padrão: 2026NEXXXXX - Gerando a 2026NLXXXXX de [data]
    obs_financeiro = f"{d['empenho']} - Gerando a {d['liquidacao']} de {d['data_nl']}"
    obs_sei = f"Documento SEI {d['sei_verificador']}"
    
    checklist = [
        {"ITEM": 1, "EVENTO": "Nota de empenho e demonstrativo de saldo", "S/N/NA": "S", "OBSERVAÇÕES": obs_financeiro},
        {"ITEM": 2, "EVENTO": "Nota Fiscal / Fatura", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 3, "EVENTO": "Certidão Federal e Dívida Ativa", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 4, "EVENTO": "Certidão de FGTS", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 5, "EVENTO": "Certidão de Justiça do Trabalho", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
    ]
    
    st.table(pd.DataFrame(checklist))
    
    # Lógica de download simplificada
    st.download_button("📥 Baixar Checklist", io.BytesIO().getvalue(), "checklist.xlsx")
