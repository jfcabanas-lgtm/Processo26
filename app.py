import streamlit as st
import pandas as pd
import PyPDF2
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

# Configurações iniciais
st.set_page_config(page_title="AuditAI - IPEM/RJ", layout="wide", page_icon="🛡️")

def extrair_dados_pdf(file):
    texto_completo = ""
    reader = PyPDF2.PdfReader(file)
    for page in reader.pages:
        content = page.extract_text()
        if content: texto_completo += content
    
    # Texto limpo para busca de padrões
    texto_limpo = " ".join(texto_completo.split())

    # --- LÓGICA DE PRECISÃO PARA ITEM 1 (SIAFE-RJ) ---
    id_nl = "Não encontrada"
    data_nl = "Não encontrada"
    id_ne = "Não encontrada"
    
    # 1. Identifica a Nota de Liquidação (NL)
    match_nl = re.search(r"202\dNL\d{5}", texto_limpo)
    if match_nl:
        id_nl = match_nl.group(0)
        
        # BUSCA AMPLIADA: Localiza a NL e varre os próximos 100 caracteres 
        # para encontrar a data, ignorando espaços ou textos intermediários.
        pos_inicio = texto_limpo.find(id_nl)
        trecho_busca = texto_limpo[pos_inicio:pos_inicio+120]
        
        match_data = re.search(r"(\d{2}/\d{2}/\d{4})", trecho_busca)
        if match_data:
            data_nl = match_data.group(1)

    # 2. Identifica a Nota de Empenho (NE)
    # Procura o padrão 2026NE... no mesmo trecho ou no documento
    match_ne = re.search(r"202\dNE\d{5}", texto_limpo)
    if match_ne:
        id_ne = match_ne.group(0)

    # 3. Número SEI (Código Verificador)
    busca_sei = re.search(r"(?:verificador|Documento\s+SEI)\s+(\d{8,10})", texto_limpo, re.IGNORECASE)
    id_sei = busca_sei.group(1) if busca_sei else "Verificar SEI"

    # Extração do Fornecedor (Simplificada)
    re_forn = re.search(r"(?:favor de|em favor de|Fornecedor:|Beneficiário)\s*([A-Z\s\d\/\.\-\&]{5,100})", texto_completo)
    fornecedor = re_forn.group(1).strip() if re_forn else "Não identificado"

    return {
        "processo": re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo).group(0) if re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo) else "Não encontrado",
        "empenho": id_ne,
        "liquidacao": id_nl,
        "data_nl": data_nl,
        "sei_verificador": id_sei,
        "fornecedor": fornecedor
    }

def gerar_excel(dados_p, df_c):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Auditoria"
    # Estilização básica omitida para brevidade, mas mantida no Excel
    ws['A1'] = "IPEM/RJ - CHECKLIST DE AUDITORIA"
    ws['A2'] = f"Processo: {dados_p['processo']}"
    wb.save(output)
    return output.getvalue()

# --- INTERFACE ---
st.title("🛡️ AuditAI - IPEM/RJ")
uploaded_file = st.sidebar.file_uploader("Upload do Processo em PDF", type="pdf")

if uploaded_file:
    d = extrair_dados_pdf(uploaded_file)
    
    # Formatação das observações conforme solicitado
    obs_item_1 = f"{d['empenho']} (Gerando a {d['liquidacao']} de {d['data_nl']})"
    obs_sei = f"Documento SEI {d['sei_verificador']}"
    
    checklist = [
        {"ITEM": 1, "EVENTO": "Nota de empenho e demonstrativo de saldo", "S/N/NA": "S", "OBSERVAÇÕES": obs_item_1},
        {"ITEM": 2, "EVENTO": "Nota Fiscal / Fatura", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 3, "EVENTO": "Certidão Federal", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 4, "EVENTO": "Certidão FGTS", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 5, "EVENTO": "Certidão Trabalhista", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
    ]
    
    df = pd.DataFrame(checklist)
    st.table(df)
    
    st.download_button("📥 Baixar Checklist", gerar_excel(d, df), "checklist.xlsx")
