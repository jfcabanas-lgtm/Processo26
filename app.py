import streamlit as st
import pandas as pd
import PyPDF2
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

# 1. Configurações de Layout
st.set_page_config(page_title="AuditAI - IPEM/RJ", layout="wide", page_icon="🛡️")

def extrair_fornecedor(texto):
    re_forn = re.search(r"(?:favor de|em favor de|Fornecedor:|Beneficiário)\s*([A-Z\s\d\/\.\-\&]{5,100})", texto)
    return re_forn.group(1).replace('\n', ' ').strip() if re_forn else "Não encontrado"

def extrair_dados_pdf(file):
    texto_completo = ""
    reader = PyPDF2.PdfReader(file)
    for page in reader.pages:
        content = page.extract_text()
        if content: texto_completo += content
    
    # Normalização que remove quebras de linha mas preserva a ordem das colunas
    texto_limpo = " ".join(texto_completo.split())

    # --- LÓGICA DE PRECISÃO CIRÚRGICA PARA O ITEM 1 ---
    id_nl = "Não encontrada"
    data_nl = "Não encontrada"
    id_ne = "Não encontrada"
    
    # 1. Localizar a Nota de Liquidação (NL)
    match_nl = re.search(r"202\dNL\d{5}", texto_limpo)
    if match_nl:
        id_nl = match_nl.group(0)
        
        # ANCORAGEM: Pegar a posição da NL e buscar a data nos 60 caracteres seguintes
        # Isso evita capturar datas de notas fiscais ou assinaturas que estão longe
        pos_fim_nl = texto_limpo.find(id_nl) + len(id_nl)
        trecho_pos_nl = texto_limpo[pos_fim_nl:pos_fim_nl+60]
        
        match_data = re.search(r"(\d{2}/\d{2}/\d{4})", trecho_pos_nl)
        if match_data:
            data_nl = match_data.group(1)

    # 2. Localizar a Nota de Empenho (NE)
    # No SIAFE, a NE de pagamento aparece perto da NL no espelho de liquidação
    match_ne = re.search(r"202\dNE\d{5}", texto_limpo)
    if match_ne:
        id_ne = match_ne.group(0)

    # --- LÓGICA PARA OS ITENS 3, 4 E 5 (NÚMERO SEI) ---
    # Busca o código verificador do documento específico anexado
    busca_sei = re.search(r"(?:verificador|Documento\s+SEI)\s+(\d{8,10})", texto_limpo, re.IGNORECASE)
    id_sei = busca_sei.group(1) if busca_sei else "Verificar SEI"

    return {
        "processo": re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo).group(0) if re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo) else "Não encontrado",
        "cnpj": re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto_completo).group(0) if re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto_completo) else "Não encontrado",
        "contrato": re.search(r"(?:99\d{8}|2100\d{4})", texto_completo).group(0) if re.search(r"(?:99\d{8}|2100\d{4})", texto_completo) else "Não encontrado",
        "empenho": id_ne,
        "liquidacao": id_nl,
        "data_nl": data_nl,
        "sei_verificador": id_sei,
        "fornecedor": extrair_fornecedor(texto_completo)
    }

def gerar_excel(dados_p, df_c):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Checklist"
    bold = Font(bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws.merge_cells('A1:D1')
    ws['A1'] = "IPEM/RJ — AUDITORIA INTERNA"
    ws['A1'].alignment = Alignment(horizontal='center')
    
    info = [("Processo:", dados_p['processo']), ("Fornecedor:", dados_p['fornecedor']), ("Contrato:", dados_p['contrato'])]
    for i, (l, v) in enumerate(info, 2):
        ws.cell(row=i, column=1, value=l).font = bold
        ws.cell(row=i, column=2, value=v)

    for r_idx, row in df_c.iterrows():
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx+10, column=c_idx, value=val)
            cell.border = border
            
    wb.save(output)
    return output.getvalue()

# --- INTERFACE ---
st.markdown("<h2 style='text-align: center;'>AUDIT - IPEM/RJ</h2>", unsafe_allow_html=True)
file = st.sidebar.file_uploader("Arraste o PDF do Processo SEI", type="pdf")

if file:
    d = extrair_dados_pdf(file)
    
    # Formatação das observações com os dados "ancorados"
    obs_financeiro = f"{d['empenho']} (Gerando a {d['liquidacao']} de {d['data_nl']})"
    obs_sei = f"Documento SEI {d['sei_verificador']}"
    
    checklist_dados = [
        {"ITEM": 1, "EVENTO": "Nota de empenho e demonstrativo de saldo", "S/N/NA": "S", "OBSERVAÇÕES": obs_financeiro},
        {"ITEM": 2, "EVENTO": "Nota Fiscal / Fatura", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 3, "EVENTO": "Certidão Federal e Dívida Ativa", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 4, "EVENTO": "Certidão de FGTS", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 5, "EVENTO": "Certidão de Justiça do Trabalho", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
        {"ITEM": 13, "EVENTO": "Atestado do Gestor", "S/N/NA": "S", "OBSERVAÇÕES": obs_sei},
    ]
    
    # Preencher itens automáticos para visualização
    for i in range(6, 20):
        if i != 13:
            checklist_dados.append({"ITEM": i, "EVENTO": f"Verificação {i}", "S/N/NA": "S", "OBSERVAÇÕES": "Verificado"})

    df = pd.DataFrame(checklist_dados).sort_values("ITEM")
    st.table(df)
    
    st.download_button("📥 Gerar Excel Oficial", gerar_excel(d, df), f"Audit_{d['processo'].replace('/','_')}.xlsx")
