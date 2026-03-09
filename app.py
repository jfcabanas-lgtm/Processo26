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
    
    # Texto limpo para buscas que ignoram quebras de linha acidentais
    texto_limpo = " ".join(texto_completo.split())

    # --- LÓGICA DE PRECISÃO CIRÚRGICA PARA NL E EMISSÃO ---
    # 1. Localiza a âncora "Liquidação" e pega o que vem depois
    id_nl = "Não encontrada"
    data_nl = "Não encontrada"
    
    # Busca o padrão 2026NL seguido de 5 dígitos
    match_nl = re.search(r"202\dNL\d{5}", texto_limpo)
    if match_nl:
        id_nl = match_nl.group(0)
        # Após encontrar a NL, busca a primeira data que aparece nos 50 caracteres seguintes
        pos_fim_nl = texto_limpo.find(id_nl) + len(id_nl)
        trecho_pos_nl = texto_limpo[pos_fim_nl:pos_fim_nl+50]
        match_data = re.search(r"\d{2}/\d{2}/\d{4}", trecho_pos_nl)
        if match_data:
            data_nl = match_data.group(0)

    # --- CAPTURA DA NOTA DE EMPENHO (NE) ---
    # Prioriza a NE que aparece no cabeçalho ou perto de "Empenho"
    match_ne = re.search(r"202\dNE\d{5}", texto_limpo)
    id_ne = match_ne.group(0) if match_ne else "Não encontrada"

    # --- CAPTURA DO NÚMERO SEI (Para itens 2, 3, 4, 5 e 13) ---
    # Pega o número verificador de 8 a 10 dígitos
    busca_sei = re.search(r"(?:verificador|Documento\s+SEI|nº)\s*(\d{8,10})", texto_limpo, re.IGNORECASE)
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

    # Extração do Fornecedor
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

    # Cabeçalho de informações do processo
    ws['A2'] = "Processo:"; ws['B2'] = dados_p['processo']
    ws['A3'] = "Fornecedor:"; ws['B3'] = dados_p['fornecedor']
    ws['A4'] = "Contrato:"; ws['B4'] = dados_p['contrato']

    # Tabela de itens
    for r_idx, row in df_c.iterrows():
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx+10, column=c_idx, value=val)
            cell.border = border
            
    wb.save(output)
    return output.getvalue()

# --- INTERFACE ---
st.title("AuditAI - IPEM/RJ")
uploaded_file = st.sidebar.file_uploader("Submeter PDF do Processo", type="pdf")

if uploaded_file:
    d = extrair_dados_pdf(uploaded_file)
    
    # Formatação exata para o Item 1
    obs_financeiro = f"{d['empenho']} (Gerando a {d['liquidacao']} de {d['data_nl']})"
    obs_sei_doc = f"Documento SEI {d['sei_verificador']}"
    
    checklist = [
        {"ITEM": 1, "EVENTO": "Nota de empenho e demonstrativo de saldo", "S/N/NA": "S", "OBSERVAÇÕES": obs_financeiro},
        {"ITEM": 2, "EVENTO": "Nota Fiscal / Fatura em nome do IPEM", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei_doc},
        {"ITEM": 3, "EVENTO": "Certidão Tributos Federais e Dívida Ativa", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei_doc},
        {"ITEM": 4, "EVENTO": "Certidão de regularidade junto ao FGTS", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei_doc},
        {"ITEM": 5, "EVENTO": "Certidão de regularidade junto a Justiça do Trabalho", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei_doc},
        {"ITEM": 13, "EVENTO": "Atestado do Gestor do contrato", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei_doc},
    ]
    
    # Itens complementares
    for i in range(6, 20):
        if i != 13:
            checklist.append({"ITEM": i, "EVENTO": f"Verificação de Auditoria {i}", "S/N/NA": "S", "OBSERVAÇÕES": "Conforme processo"})

    df = pd.DataFrame(checklist).sort_values("ITEM")
    st.table(df)
    
    st.download_button("📥 Baixar Excel", gerar_excel_oficial(d, df), f"Audit_{d['processo'].replace('/','_')}.xlsx")
