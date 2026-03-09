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
    
    # Normalização total: remove quebras de linha e espaços duplos
    texto_limpo = " ".join(texto_completo.split())

    # --- LÓGICA DE EXTRAÇÃO REFORÇADA PARA O ITEM 1 ---
    id_nl = "Não encontrada"
    data_nl = "Não encontrada"
    id_ne = "Não encontrada"
    
    # 1. Localiza a Nota de Liquidação
    match_nl = re.search(r"202\dNL\d{5}", texto_limpo)
    if match_nl:
        id_nl = match_nl.group(0)
        
        # 2. Captura a data de emissão associada à NL
        # Busca a primeira data que aparece nos 150 caracteres seguintes ao número da NL
        pos_nl = texto_limpo.find(id_nl)
        bloco_pos_nl = texto_limpo[pos_nl:pos_nl+150]
        
        datas_encontradas = re.findall(r"\d{2}/\d{2}/\d{4}", bloco_pos_nl)
        if datas_encontradas:
            # Pega a primeira data que aparecer após o número da NL
            data_nl = datas_encontradas[0]

    # 3. Localiza a Nota de Empenho
    match_ne = re.search(r"202\dNE\d{5}", texto_limpo)
    if match_ne:
        id_ne = match_ne.group(0)

    # 4. Número SEI (Código Verificador)
    busca_sei = re.search(r"(?:verificador|Documento\s+SEI)\s+(\d{8,10})", texto_limpo, re.IGNORECASE)
    id_sei = busca_sei.group(1) if busca_sei else "Verificar SEI"

    # Extração básica de Processo e Fornecedor
    re_proc = re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo)
    re_forn = re.search(r"(?:favor de|em favor de|Fornecedor:|Beneficiário)\s*([A-Z\s\d\/\.\-\&]{5,100})", texto_completo)

    return {
        "processo": re_proc.group(0) if re_proc else "Não encontrado",
        "empenho": id_ne,
        "liquidacao": id_nl,
        "data_nl": data_nl,
        "sei_verificador": id_sei,
        "fornecedor": re_forn.group(1).strip() if re_forn else "Não encontrado"
    }

# --- INTERFACE ---
st.markdown("<h2 style='text-align: center;'>AUDIT - IPEM/RJ</h2>", unsafe_allow_html=True)
uploaded_file = st.sidebar.file_uploader("Upload do Processo em PDF", type="pdf")

if uploaded_file:
    d = extrair_dados_pdf(uploaded_file)
    
    # Montagem exata do campo conforme solicitado
    # Se a data não for achada, ele manterá "Não encontrada" para você perceber o erro
    campo_item_1 = f"{d['empenho']} - Gerando a {d['liquidacao']} de {d['data_nl']}"
    obs_sei = f"Documento SEI {d['sei_verificador']}"
    
    checklist = [
        {"ITEM": 1, "EVENTO": "Nota de empenho e demonstrativo de saldo", "S/N/NA": "S", "OBSERVAÇÕES": campo_item_1},
        {"ITEM": 2, "EVENTO": "Nota Fiscal / Fatura", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 3, "EVENTO": "Certidão Federal e Dívida Ativa", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 4, "EVENTO": "Certidão de FGTS", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 5, "EVENTO": "Certidão de Justiça do Trabalho", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
    ]
    
    st.table(pd.DataFrame(checklist))
    
    # Botão de download (Exemplo simplificado)
    st.download_button("📥 Baixar Excel", io.BytesIO().getvalue(), "checklist.xlsx")
