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
    
    # Texto limpo remove quebras de linha que separam o número da NL da sua data
    texto_limpo = " ".join(texto_completo.split())

    # --- LÓGICA DE CAPTURA DA DATA DE EMISSÃO (Vínculo com a NL) ---
    id_nl = "Não encontrada"
    data_emissao_nl = "Não encontrada"
    
    # 1. Busca o padrão da NL (Ex: 2026NL00021)
    match_nl = re.search(r"202\dNL\d{5}", texto_limpo)
    
    if match_nl:
        id_nl = match_nl.group(0)
        # 2. O SEGREDO: Pega a posição onde termina o número da NL
        posicao_fim_nl = texto_limpo.find(id_nl) + len(id_nl)
        # 3. Analisa apenas os 30 caracteres seguintes (onde a data DEVE estar no SIAFE)
        trecho_proximo = texto_limpo[posicao_fim_nl:posicao_fim_nl+30]
        
        match_data = re.search(r"(\d{2}/\d{2}/\d{4})", trecho_proximo)
        if match_data:
            data_emissao_nl = match_data.group(1)

    # --- CAPTURA DA NOTA DE EMPENHO (NE) ---
    match_ne = re.search(r"202\dNE\d{5}", texto_limpo)
    id_ne = match_ne.group(0) if match_ne else "Não encontrada"

    # --- CAPTURA DO NÚMERO SEI (Itens 2, 3, 4, 5 e 13) ---
    busca_sei = re.search(r"(?:verificador|Documento\s+SEI)\s+(\d{8,10})", texto_limpo, re.IGNORECASE)
    id_sei = busca_sei.group(1) if busca_sei else "Verificar SEI"

    dados = {
        "processo": re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo).group(0) if re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo) else "Não encontrado",
        "cnpj": re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto_completo).group(0) if re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto_completo) else "Não encontrado",
        "contrato": re.search(r"(?:99\d{8}|2100\d{4})", texto_completo).group(0) if re.search(r"(?:99\d{8}|2100\d{4})", texto_completo) else "Não encontrado",
        "empenho": id_ne,
        "liquidacao": id_nl,
        "data_nl": data_emissao_nl,
        "sei_verificador": id_sei
    }
    
    # Fornecedor
    re_forn = re.search(r"(?:favor de|em favor de|Fornecedor:|Beneficiário)\s*([A-Z\s\d\/\.\-\&]{5,100})", texto_completo)
    dados["fornecedor"] = re_forn.group(1).replace('\n', ' ').strip() if re_forn else "Não encontrado"
    
    return dados

# Interface do Streamlit mantida com a formatação das observações
st.title("AuditAI - IPEM/RJ")
uploaded_file = st.sidebar.file_uploader("Upload PDF do Processo", type="pdf")

if uploaded_file:
    d = extrair_dados_pdf(uploaded_file)
    
    # Montagem do campo Observações do Item 1 exatamente como solicitado
    obs_item_1 = f"{d['empenho']} (Gerando a {d['liquidacao']} de {d['data_nl']})"
    obs_sei = f"Documento SEI {d['sei_verificador']}"
    
    checklist = [
        {"ITEM": 1, "EVENTO": "Nota de empenho e demonstrativo de saldo", "S/N/NA": "S", "OBSERVAÇÕES": obs_item_1},
        {"ITEM": 2, "EVENTO": "Nota Fiscal / Fatura", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 3, "EVENTO": "Certidão Federal", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 4, "EVENTO": "Certidão FGTS", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 5, "EVENTO": "Certidão Trabalhista", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
    ]
    
    st.table(pd.DataFrame(checklist))
    
    # Lógica de download do Excel mantida...
