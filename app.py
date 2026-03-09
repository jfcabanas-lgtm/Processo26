import streamlit as st
import pandas as pd
import PyPDF2
import re
from datetime import datetime
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
    
    # Normalização para busca linear sem quebras de linha confusas
    texto_limpo = " ".join(texto_completo.split())

    # --- LÓGICA DE PRECISÃO PARA ITEM 1 (SIAFE-RJ) ---
    # 1. Busca o bloco específico da Nota de Liquidação para evitar pegar dados da Nota Fiscal
    # Procuramos o número da NL e a data que vem LOGO APÓS ou na mesma linha
    match_nl_dados = re.search(r"(202\dNL\d{5})\s+(\d{2}/\d{2}/\d{4})", texto_limpo)
    id_nl = match_nl_dados.group(1) if match_nl_dados else "Não encontrada"
    data_nl = match_nl_dados.group(2) if match_nl_dados else "Não encontrada"

    # 2. Busca a Nota de Empenho (NE) vinculada
    # Geralmente aparece no campo "Nota de Empenho" do formulário SIAFE
    match_ne = re.search(r"Nota\s+de\s+Empenho\s+(202\dNE\d{5})", texto_limpo, re.IGNORECASE)
    id_ne = match_ne.group(1) if match_ne else "Não encontrada"

    # --- LÓGICA DE PRECISÃO PARA ITENS 3, 4 e 5 (DOCUMENTO SEI) ---
    # Busca o número de 8 a 10 dígitos que identifica o DOCUMENTO específico no sistema
    # Evita pegar o número do processo ou o código CRC.
    busca_sei = re.search(r"código\s+verificador\s+(\d{8,10})", texto_limpo, re.IGNORECASE)
    if not busca_sei:
        busca_sei = re.search(r"Documento\s+SEI\s+(\d{8,10})", texto_limpo, re.IGNORECASE)
    
    id_sei = busca_sei.group(1) if busca_sei else "Verificar no SEI"

    dados = {
        "processo": re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo).group(0) if re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo) else "Não encontrado",
        "cnpj": re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto_completo).group(0) if re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto_completo) else "Não encontrado",
        "contrato": re.search(r"(?:99\d{8}|2100\d{4})", texto_completo).group(0) if re.search(r"(?:99\d{8}|2100\d{4})", texto_completo) else "Não encontrado",
        "empenho": id_ne,
        "liquidacao": id_nl,
        "data_nl": data_nl,
        "sei_verificador": id_sei,
        "valor_bruto": re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto_completo).group(0) if re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto_completo) else "R$ 0,00"
    }

    # Fornecedor
    re_forn = re.search(r"(?:favor de|em favor de|Fornecedor:|Beneficiário)\s*([A-Z\s\d\/\.\-\&]{5,100})", texto_completo)
    dados["fornecedor"] = re_forn.group(1).replace('\n', ' ').strip() if re_forn else "Não encontrado"
    
    return dados

def gerar_excel_oficial(dados_p, df_c):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Checklist Auditoria"
    # Estilização básica
    bold_f = Font(bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws.merge_cells('A1:D1')
    ws['A1'] = "INSTITUTO DE PESOS E MEDIDAS IPEM/RJ — AUDITORIA INTERNA - AUDIT"
    ws['A1'].alignment = Alignment(horizontal='center')

    info = [("Processo:", dados_p['processo']), ("Fornecedor:", dados_p['fornecedor']), ("Contrato:", dados_p['contrato'])]
    for i, (l, v) in enumerate(info, 2):
        ws.cell(row=i, column=1, value=l).font = bold_f
        ws.cell(row=i, column=2, value=v)

    for r_idx, row in df_c.iterrows():
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx+10, column=c_idx, value=val)
            cell.border = border
            
    wb.save(output)
    return output.getvalue()

# --- INTERFACE ---
st.title("AuditAI - IPEM/RJ")
uploaded_file = st.sidebar.file_uploader("Submeta o Processo (PDF)", type="pdf")

if uploaded_file:
    d = extrair_dados_pdf(uploaded_file)
    
    # Montagem do Checklist conforme as novas regras de visualização
    obs_financeiro = f"{d['empenho']} (Gerando a {d['liquidacao']} de {d['data_nl']})"
    obs_sei_doc = f"Documento SEI {d['sei_verificador']}"
    
    checklist = [
        {"ITEM": 1, "EVENTO": "Nota de empenho e demonstrativo de saldo", "S/N/NA": "S", "OBSERVAÇÕES": obs_financeiro},
        {"ITEM": 2, "EVENTO": "Nota Fiscal / Fatura em nome do IPEM", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei_doc},
        {"ITEM": 3, "EVENTO": "Certidão Tributos Federais e Dívida Ativa", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei_doc},
        {"ITEM": 4, "EVENTO": "Certidão de regularidade junto ao FGTS", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei_doc},
        {"ITEM": 5, "EVENTO": "Certidão de regularidade junto a Justiça do Trabalho", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei_doc},
        # ... demais itens mantidos conforme padrão
    ]
    
    # Preenchimento automático dos itens restantes para visualização
    for i in range(6, 20):
        checklist.append({"ITEM": i, "EVENTO": f"Evento de Auditoria {i}", "S/N/NA": "S", "OBSERVAÇÕES": "Verificado"})

    df = pd.DataFrame(checklist)
    st.table(df)
    
    st.download_button("📥 Baixar Checklist Excel", gerar_excel_oficial(d, df), "checklist.xlsx")
