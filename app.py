import streamlit as st
import pandas as pd
import PyPDF2
import re
from datetime import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

# 1. Configurações Iniciais
st.set_page_config(page_title="AuditAI - IPEM/RJ", layout="wide", page_icon="🛡️")

def extrair_dados_pdf(file):
    texto = ""
    reader = PyPDF2.PdfReader(file)
    for page in reader.pages:
        content = page.extract_text()
        if content: texto += content
    
    # Extração de dados com Regex
    dados = {
        "processo": re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto).group(0) if re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto) else "Não encontrado",
        "cnpj": re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto).group(0) if re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto) else "Não encontrado",
        "contrato": re.search(r"2100\d{4}", texto).group(0) if re.search(r"2100\d{4}", texto) else "Não encontrado",
        "empenho": re.search(r"202\dNE\d{5}", texto).group(0) if re.search(r"202\dNE\d{5}", texto) else "2026NEXXXXX",
        "liquidacao": re.search(r"202\dNL\d{5}", texto).group(0) if re.search(r"202\dNL\d{5}", texto) else "2026NLXXXXX",
        "valor_bruto": re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto).group(0) if re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto) else "R$ 0,00",
        "sei_verificador": re.search(r"verificador\s(\d{9})", texto).group(1) if re.search(r"verificador\s(\d{9})", texto) else "Não identificado"
    }

    # Datas
    datas = re.findall(r"(\d{2}/\d{2}/\d{4})", texto)
    dados["data_proc"] = datas[0] if datas else datetime.now().strftime('%d/%m/%Y')
    dados["validade"] = max(datas, key=lambda d: datetime.strptime(d, '%d/%m/%Y')) if datas else "Verificar"

    # Fornecedor e Gestor
    re_forn = re.search(r"(?:favor de|em favor de|Fornecedor:)\s*([A-Z\s\d\/\.\-\&]{5,100})", texto)
    dados["fornecedor"] = re_forn.group(1).replace('\n', ' ').strip() if re_forn else "Não encontrado"
    
    re_gestor = re.search(r"assinado eletronicamente por\s*([A-Za-z\s]+),", texto)
    dados["gestor"] = re_gestor.group(1).strip() if re_gestor else "Não identificado"
    
    return dados

def gerar_excel_oficial(dados_p, df_c):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Checklist Auditoria"
    
    bold_f = Font(bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Título do Relatório
    ws.merge_cells('A1:D1')
    ws['A1'] = "INSTITUTO DE PESOS E MEDIDAS IPEM/RJ — AUDITORIA INTERNA - AUDIT"
    ws['A1'].font = Font(bold=True, size=11)
    ws['A1'].alignment = center

    # Cabeçalho de Dados
    info_header = [
        ("Processo SEI:", dados_p['processo']),
        ("Fornecedor:", dados_p['fornecedor']),
        ("Contrato:", dados_p['contrato']),
        ("CNPJ:", dados_p['cnpj']),
        ("Valor Bruto:", dados_p['valor_bruto']),
        ("Gestor:", dados_p['gestor'])
    ]
    for idx, (label, val) in enumerate(info_header, start=2):
        ws.cell(row=idx, column=1, value=label).font = bold_f
        ws.cell(row=idx, column=2, value=val)

    # Cabeçalho da Tabela
    cols = ["ITEM", "EVENTO A SER VERIFICADO", "S/N/NA", "OBSERVAÇÕES"]
    for c_idx, text in enumerate(cols, start=1):
        cell = ws.cell(row=9, column=c_idx, value=text)
        cell.font = bold_f
        cell.border = border
        cell.alignment = center

    # Preenchimento dos Itens
    for r_idx, row in df_c.iterrows():
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx+10, column=c_idx, value=value)
            cell.border = border
            if c_idx != 2: cell.alignment = center

    wb.save(output)
    return output.getvalue()

# --- INTERFACE STREAMLIT ---
st.markdown("<h2 style='text-align: center; color: #1f4e78;'>AUDIT - IPEM/RJ</h2>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center;'>Gerador de Checklist (19 Itens)</h4>", unsafe_allow_html=True)
st.divider()

uploaded_file = st.sidebar.file_uploader("Upload PDF do Processo SEI", type="pdf")

if uploaded_file:
    d = extrair_dados_pdf(uploaded_file)
    
    # Painel Visual
    c1, c2 = st.columns(2)
    with c1:
        st.info(f"**Processo:** {d['processo']}\n\n**Fornecedor:** {d['fornecedor']}")
    with c2:
        st.success(f"**Valor Bruto:** {d['valor_bruto']}\n\n**Gestor:** {d['gestor']}")

    st.divider()

    # Definição das Observações Específicas
    obs_1 = f"{d['empenho']} (Gerando a {d['liquidacao']} de {d['data_proc']})"
    obs_sei = f"Documento SEI {d['sei_verificador']}"
    
    checklist_19 = [
        {"ITEM": 1, "EVENTO": "Nota de empenho e demonstrativo de saldo", "S/N/NA": "S", "OBSERVAÇÕES": obs_1},
        {"ITEM": 2, "EVENTO": "Nota Fiscal / Fatura em nome do IPEM", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 3, "EVENTO": "Certidão Tributos Federais e Dívida Ativa", "S/N/NA": "S", "OBSERVAÇÕES": f"Válida até: {d['validade']}"},
        {"ITEM": 4, "EVENTO": "Certidão de regularidade - CRF (FGTS)", "S/N/NA": "S", "OBSERVAÇÕES": f"Válida até: {d['validade']}"},
        {"ITEM": 5, "EVENTO": "Certidão de regularidade - CNDT", "S/N/NA": "S", "OBSERVAÇÕES": f"Válida até: {d['validade']}"},
        {"ITEM": 6, "EVENTO": "Certidão de Regularidade Estadual (ICMS)", "S/N/NA": "S", "OBSERVAÇÕES": "Verificada"},
        {"ITEM": 7, "EVENTO": "Certidão de Regularidade Municipal (ISS)", "S/N/NA": "S", "OBSERVAÇÕES": "Verificada"},
        {"ITEM": 8, "EVENTO": "Consulta ao CADIN Estadual", "S/N/NA": "S", "OBSERVAÇÕES": "Nada consta"},
        {"ITEM": 9, "EVENTO": "Consulta de Sanções (CEIS/CNEP)", "S/N/NA": "S", "OBSERVAÇÕES": "Nada consta"},
        {"ITEM": 10, "EVENTO": "Incidência de tributos retidos na fonte?", "S/N/NA": "S", "OBSERVAÇÕES": "Verificado na NL"},
        {"ITEM": 11, "EVENTO": "Comprovação de não incidência de tributos?", "S/N/NA": "NA", "OBSERVAÇÕES": ""},
        {"ITEM": 12, "EVENTO": "Portaria de Nomeação de Fiscalização", "S/N/NA": "S", "OBSERVAÇÕES": "Portaria GAPRE"},
        {"ITEM": 13, "EVENTO": "Atestado do Gest
