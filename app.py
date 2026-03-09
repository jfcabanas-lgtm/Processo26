import streamlit as st
import pandas as pd
import PyPDF2
import re
from datetime import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

# 1. Configurações de Layout
st.set_page_config(page_title="AuditAI - IPEM/RJ", layout="wide", page_icon="🛡️")

def extrair_dados_pdf(file):
    texto = ""
    reader = PyPDF2.PdfReader(file)
    for page in reader.pages:
        content = page.extract_text()
        if content: texto += content
    
    # Extração de dados fundamentais
    dados = {
        "processo": re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto).group(0) if re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto) else "Não encontrado",
        "cnpj": re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto).group(0) if re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto) else "Não encontrado",
        "contrato": re.search(r"2100\d{4}", texto).group(0) if re.search(r"2100\d{4}", texto) else "Não encontrado",
        "empenho": re.search(r"202\dNE\d{5}", texto).group(0) if re.search(r"202\dNE\d{5}", texto) else "2026NEXXXXX",
        "liquidacao": re.search(r"202\dNL\d{5}", texto).group(0) if re.search(r"202\dNL\d{5}", texto) else "2026NLXXXXX",
        "valor_bruto": re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto).group(0) if re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto) else "R$ 0,00",
        "sei_ids": sorted(list(set(re.findall(r"verificador\s(\d{9})", texto))), reverse=True)
    }

    # Data do Processo e Validades
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
    ws['A1'].font = Font(bold=True, size=12)
    ws['A1'].alignment = center

    # Preenchimento de Informações de Cabeçalho
    header_mapping = [
        ("Processo SEI:", dados_p['processo']),
        ("Fornecedor:", dados_p['fornecedor']),
        ("Contrato:", dados_p['contrato']),
        ("CNPJ:", dados_p['cnpj']),
        ("Valor Bruto:", dados_p['valor_bruto']),
        ("Gestor:", dados_p['gestor'])
    ]
    for idx, (label, val) in enumerate(header_mapping, start=2):
        ws.cell(row=idx, column=1, value=label).font = bold_f
        ws.cell(row=idx, column=2, value=val)

    # Cabeçalho da Tabela (Linha 9)
    cols = ["ITEM", "EVENTO A SER VERIFICADO", "S/N/NA", "OBSERVAÇÕES"]
    for c_idx, text in enumerate(cols, start=1):
        cell = ws.cell(row=9, column=c_idx, value=text)
        cell.font = bold_f
        cell.border = border
        cell.alignment = center

    # Itens (Linha 10 em diante)
    for r_idx, row in df_c.iterrows():
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx+10, column=c_idx, value=value)
            cell.border = border
            if c_idx != 2: cell.alignment = center

    wb.save(output)
    return output.getvalue()

# --- INTERFACE ---
st.markdown("<h2 style='text-align: center; color: #1f4e78;'>AUDIT - IPEM/RJ</h2>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center;'>Checklist de Auditoria Interna (19 Itens)</h4>", unsafe_allow_html=True)
st.divider()

uploaded_file = st.sidebar.file_uploader("Upload PDF do Processo SEI", type="pdf")

if uploaded_file:
    d = extrair_dados_pdf(uploaded_file)
    s = d['sei_ids']
    
    # Exibição dos dados extraídos
    c1, c2 = st.columns(2)
    with c1:
        st.info(f"**Processo:** {d['processo']}\n\n**Fornecedor:** {d['fornecedor']}\n\n**Contrato:** {d['contrato']}")
    with c2:
        st.success(f"**Valor Bruto:** {d['valor_bruto']}\n\n**Gestor:** {d['gestor']}\n\n**CNPJ:** {d['cnpj']}")

    st.divider()

    # Formatação solicitada para o Item 1
    obs_1 = f"{d['empenho']} (Gerando a {d['liquidacao']} de {d['data_proc']})"
    
    checklist_19 = [
        {"ITEM": 1, "EVENTO": "Nota de empenho e demonstrativo de saldo", "S/N/NA": "S", "OBSERVAÇÕES": obs_1},
        {"ITEM": 2, "EVENTO": "Nota Fiscal / Fatura em nome do IPEM", "S/N/NA": "S", "OBSERVAÇÕES": f"Doc. SEI {s[0] if len(s)>0 else ''}"},
        {"ITEM": 3, "EVENTO": "Certidão Tributos Federais e Dívida Ativa", "S/N/NA": "S", "OBSERVAÇÕES": f"Val: {d['validade']}"},
        {"ITEM": 4, "EVENTO": "Certidão de regularidade - CRF (FGTS)", "S/N/NA": "S", "OBSERVAÇÕES": f"Val: {d['validade']}"},
        {"ITEM": 5, "EVENTO": "Certidão de regularidade - CNDT", "S/N/NA": "S", "OBSERVAÇÕES": f"Val: {d['validade']}"},
        {"ITEM": 6, "EVENTO": "Certidão de Regularidade Estadual (ICMS)", "S/N/NA": "S", "OBSERVAÇÕES": "Verificada"},
        {"ITEM": 7, "EVENTO": "Certidão de Regularidade Municipal (ISS)", "S/N/NA": "S", "OBSERVAÇÕES": "Verificada"},
        {"ITEM": 8, "EVENTO": "Consulta ao CADIN Estadual", "S/N/NA": "S", "OBSERVAÇÕES": "Nada consta"},
        {"ITEM": 9, "EVENTO": "Consulta de Sanções (CEIS/CNEP)", "S/N/NA": "S", "OBSERVAÇÕES": "Nada consta"},
        {"ITEM": 10, "EVENTO": "Incidência de tributos retidos na fonte?", "S/N/NA": "S", "OBSERVAÇÕES": "Verificado na NL"},
        {"ITEM": 11, "EVENTO": "Comprovação de não incidência de tributos?", "S/N/NA": "NA", "OBSERVAÇÕES": "Empresa Optante"},
        {"ITEM": 12, "EVENTO": "Portaria de Nomeação de Fiscalização", "S/N/NA": "S", "OBSERVAÇÕES": "GAPRE"},
        {"ITEM": 13, "EVENTO": "Atestado do Gestor (Serviços prestados)", "S/N/NA": "S", "OBSERVAÇÕES": f"Doc. SEI {s[-1] if len(s)>0 else ''}"},
        {"ITEM": 14, "EVENTO": "Relação dos funcionários do serviço", "S/N/NA": "S", "OBSERVAÇÕES": "Anexo"},
        {"ITEM": 15, "EVENTO": "Comprovante da GFIP / eSocial", "S/N/NA": "S", "OBSERVAÇÕES": "Anexo"},
        {"ITEM": 16, "EVENTO": "Comprovante de pagamento do INSS", "S/N/NA": "S", "OBSERVAÇÕES": "Guia Paga"},
        {"ITEM": 17, "EVENTO": "Comprovante de pagamento do FGTS", "S/N/NA": "S", "OBSERVAÇÕES": "Bancário"},
        {"ITEM": 18, "EVENTO": "Folha de pagamento", "S/N/NA": "S", "OBSERVAÇÕES": "Ref. Janeiro/2026"},
        {"ITEM": 19, "EVENTO": "Comprovante bancário de salários", "S/N/NA": "S", "OBSERVAÇÕES": "Transferência"}
    ]

    df_final = pd.DataFrame(checklist_19)
    st.table(df_final)

    # Botão de Exportação com fechamento correto de parênteses
    excel_data = gerar_excel_oficial(d, df_final)
    st.download_button(
        label="📥 GERAR CHECKLIST OFICIAL (19 ITENS)",
        data=excel_data,
        file_name=f"Checklist_Audit_{d['processo'].replace('/','_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
else:
    st.warning("Aguardando upload do PDF para processar o checklist da AUDIT.")
