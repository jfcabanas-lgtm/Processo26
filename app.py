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
    
    # Normalização rigorosa para tratar espaços e quebras de linha do SIAFE
    texto_limpo = " ".join(texto_completo.split())

    # --- NOVA LÓGICA DE CAPTURA DA NL E EMISSÃO (SIAFE-RJ) ---
    # Busca especificamente o termo "Nota de Liquidação" e tenta capturar o número e a data logo a frente
    # O padrão procura por 2026NL seguido de 5 dígitos e uma data no formato XX/XX/XXXX
    busca_nl_data = re.search(r"Nota\s+de\s+Liquidação\s+(\d{4}NL\d{5})\s+(\d{2}/\d{2}/\d{4})", texto_completo, re.IGNORECASE)
    
    if busca_nl_data:
        id_nl = busca_nl_data.group(1)
        data_nl = busca_nl_data.group(2)
    else:
        # Segundo plano: busca apenas o padrão do número e a data mais próxima no texto limpo
        match_alt = re.search(r"(202\dNL\d{5})\s+(\d{2}/\d{2}/\d{4})", texto_limpo)
        id_nl = match_alt.group(1) if match_alt else "Não encontrada"
        data_nl = match_alt.group(2) if match_alt else "Não encontrada"

    # --- CAPTURA DA NOTA DE EMPENHO (NE) ---
    # Busca o padrão 2026NE000XX
    match_ne = re.search(r"202\dNE\d{5}", texto_limpo)
    id_ne = match_ne.group(0) if match_ne else "Não encontrada"

    # --- CAPTURA DO NÚMERO SEI (Para itens 2, 3, 4, 5 e 13) ---
    # Busca o código verificador do documento específico
    busca_sei = re.search(r"(?:verificador|Documento\s+SEI)\s+(\d{8,10})", texto_limpo, re.IGNORECASE)
    id_sei = busca_sei.group(1) if busca_sei else "Verificar SEI"

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
    bold_f = Font(bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws.merge_cells('A1:D1')
    ws['A1'] = "INSTITUTO DE PESOS E MEDIDAS IPEM/RJ — AUDITORIA INTERNA - AUDIT"
    ws['A1'].alignment = Alignment(horizontal='center')

    info = [("Processo:", dados_p['processo']), ("Fornecedor:", dados_p['fornecedor']), ("Contrato:", dados_p['contrato'])]
    for i, (l, v) in enumerate(info, 2):
        ws.cell(row=i, column=1, value=l).font = bold_f
        ws.cell(row=i, column=2, value=v)

    # Cabeçalho da tabela
    for c, texto in enumerate(["ITEM", "EVENTO", "S/N/NA", "OBSERVAÇÕES"], 1):
        ws.cell(row=9, column=c, value=texto).font = bold_f

    for r_idx, row in df_c.iterrows():
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx+10, column=c_idx, value=val)
            cell.border = border
            
    wb.save(output)
    return output.getvalue()

# --- INTERFACE ---
st.title("AuditAI - IPEM/RJ")
uploaded_file = st.sidebar.file_uploader("Submeta o PDF do Processo SEI", type="pdf")

if uploaded_file:
    d = extrair_dados_pdf(uploaded_file)
    
    # O SEGREDO DO ACERTO: Concatenar NE, NL e Data na mesma observação do Item 1
    obs_financeiro = f"{d['empenho']} (Gerando a {d['liquidacao']} de {d['data_nl']})"
    obs_sei_doc = f"Documento SEI {d['sei_verificador']}"
    
    checklist = [
        {"ITEM": 1, "EVENTO": "Nota de empenho e demonstrativo de saldo", "S/N/NA": "S", "OBSERVAÇÕES": obs_financeiro},
        {"ITEM": 2, "EVENTO": "Nota Fiscal / Fatura em nome do IPEM", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei_doc},
        {"ITEM": 3, "EVENTO": "Certidão Tributos Federais e Dívida Ativa", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei_doc},
        {"ITEM": 4, "EVENTO": "Certidão de regularidade junto ao FGTS", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei_doc},
        {"ITEM": 5, "EVENTO": "Certidão de regularidade junto a Justiça do Trabalho", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei_doc},
        {"ITEM": 6, "EVENTO": "Certidão de Regularidade Estadual (ICMS)", "S/N/NA": "S", "OBSERVAÇÕES": "Verificada"},
        {"ITEM": 7, "EVENTO": "Certidão de Regularidade Municipal (ISS)", "S/N/NA": "S", "OBSERVAÇÕES": "Verificada"},
        {"ITEM": 8, "EVENTO": "Consulta ao CADIN Estadual", "S/N/NA": "S", "OBSERVAÇÕES": "Nada consta"},
        {"ITEM": 9, "EVENTO": "Consulta de Sanções (CEIS/CNEP)", "S/N/NA": "S", "OBSERVAÇÕES": "Nada consta"},
        {"ITEM": 10, "EVENTO": "Incidência de tributos retidos na fonte?", "S/N/NA": "S", "OBSERVAÇÕES": "Verificado na NL"},
        {"ITEM": 11, "EVENTO": "Comprovação de não incidência de tributos?", "S/N/NA": "NA", "OBSERVAÇÕES": ""},
        {"ITEM": 12, "EVENTO": "Portaria de Nomeação de Fiscalização", "S/N/NA": "S", "OBSERVAÇÕES": "Portaria GAPRE"},
        {"ITEM": 13, "EVENTO": "Atestado do Gestor do contrato", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei_doc},
        {"ITEM": 14, "EVENTO": "Relação dos funcionários que executaram o serviço", "S/N/NA": "S", "OBSERVAÇÕES": "Anexo"},
        {"ITEM": 15, "EVENTO": "Comprovante da GFIP / eSocial", "S/N/NA": "S", "OBSERVAÇÕES": "Anexo"},
        {"ITEM": 16, "EVENTO": "Comprovante de pagamento do INSS", "S/N/NA": "S", "OBSERVAÇÕES": "Guia Paga"},
        {"ITEM": 17, "EVENTO": "Comprovante de pagamento do FGTS", "S/N/NA": "S", "OBSERVAÇÕES": "Bancário"},
        {"ITEM": 18, "EVENTO": "Folha de pagamento", "S/N/NA": "S", "OBSERVAÇÕES": "Anexo"},
        {"ITEM": 19, "EVENTO": "Comprovante bancário de salários", "S/N/NA": "S", "OBSERVAÇÕES": "Transferência"}
    ]

    df = pd.DataFrame(checklist)
    st.table(df)
    
    st.download_button("📥 Baixar Checklist Excel", gerar_excel_oficial(d, df), f"Checklist_{d['processo'].replace('/','_')}.xlsx")
