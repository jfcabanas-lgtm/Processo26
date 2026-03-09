import streamlit as st
import pandas as pd
import PyPDF2
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

# Configuração da Página
st.set_page_config(page_title="AuditAI - IPEM/RJ", layout="wide", page_icon="🛡️")

def extrair_dados_pdf(file):
    texto_completo = ""
    reader = PyPDF2.PdfReader(file)
    for page in reader.pages:
        content = page.extract_text()
        if content: texto_completo += content
    
    # Normalização do texto para busca linear
    texto_limpo = " ".join(texto_completo.split())

    # --- LÓGICA DE PRECISÃO PARA O ITEM 1 ---
    id_nl = "Não encontrada"
    data_nl = "Não encontrada"
    id_ne = "Não encontrada"
    
    # 1. Localiza a Nota de Liquidação (NL)
    match_nl = re.search(r"202\dNL\d{5}", texto_limpo)
    if match_nl:
        id_nl = match_nl.group(0)
        
        # ANCORAGEM POSICIONAL: 
        # Localiza a posição da NL no texto e busca a data nos 100 caracteres seguintes.
        # Isso permite "pular" informações de conta bancária ou convênio que ficam no meio.
        pos_inicio = texto_limpo.find(id_nl)
        janela_busca = texto_limpo[pos_inicio:pos_inicio+120]
        
        match_data = re.search(r"(\d{2}/\d{2}/\d{4})", janela_busca)
        if match_data:
            data_nl = match_data.group(1)

    # 2. Localiza a Nota de Empenho (NE)
    # Busca o padrão 2026NE no documento
    match_ne = re.search(r"202\dNE\d{5}", texto_limpo)
    if match_ne:
        id_ne = match_ne.group(0)

    # --- NÚMERO SEI (Itens 2, 3, 4, 5 e 13) ---
    busca_sei = re.search(r"(?:verificador|Documento\s+SEI)\s+(\d{8,10})", texto_limpo, re.IGNORECASE)
    id_sei = busca_sei.group(1) if busca_sei else "Verificar SEI"

    # Fornecedor e Processo
    processo = re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo)
    re_forn = re.search(r"(?:favor de|em favor de|Fornecedor:|Beneficiário)\s*([A-Z\s\d\/\.\-\&]{5,100})", texto_completo)
    
    return {
        "processo": processo.group(0) if processo else "Não encontrado",
        "empenho": id_ne,
        "liquidacao": id_nl,
        "data_nl": data_nl,
        "sei_verificador": id_sei,
        "fornecedor": re_forn.group(1).strip() if re_forn else "Não encontrado"
    }

def gerar_excel(dados_p, df_c):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Checklist"
    # Estilização
    ws['A1'] = "IPEM/RJ - AUDITORIA INTERNA"; ws['A1'].font = Font(bold=True)
    ws['A2'] = f"Processo: {dados_p['processo']}"
    # ... (lógica de preenchimento da tabela)
    wb.save(output)
    return output.getvalue()

# --- INTERFACE STREAMLIT ---
st.markdown("<h2 style='text-align: center;'>AUDIT - IPEM/RJ</h2>", unsafe_allow_html=True)
uploaded_file = st.sidebar.file_uploader("Submeter PDF do Processo SEI", type="pdf")

if uploaded_file:
    d = extrair_dados_pdf(uploaded_file)
    
    # Formatação exata do Item 1 conforme solicitado
    obs_item_1 = f"{d['empenho']} - Gerando a {d['liquidacao']} de {d['data_nl']}"
    obs_sei = f"Documento SEI {d['sei_verificador']}"
    
    checklist = [
        {"ITEM": 1, "EVENTO": "Nota de empenho e demonstrativo de saldo", "S/N/NA": "S", "OBSERVAÇÕES": obs_item_1},
        {"ITEM": 2, "EVENTO": "Nota Fiscal / Fatura", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 3, "EVENTO": "Certidão Federal e Dívida Ativa", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 4, "EVENTO": "Certidão de FGTS", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 5, "EVENTO": "Certidão de Justiça do Trabalho", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 13, "EVENTO": "Atestado do Gestor", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
    ]
    
    # Itens automáticos para visualização
    for i in range(6, 20):
        if i != 13:
            checklist.append({"ITEM": i, "EVENTO": f"Verificação {i}", "S/N/NA": "S", "OBSERVAÇÕES": "Conforme processo"})

    df = pd.DataFrame(checklist).sort_values("ITEM")
    st.table(df)
    
    st.download_button("📥 Baixar Checklist Oficial", gerar_excel(d, df), f"Audit_{d['processo'].replace('/','_')}.xlsx")
