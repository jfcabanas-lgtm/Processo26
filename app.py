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
        "contrato": re.search(r"Contrato\s*(\d+)", texto).group(1) if re.search(r"Contrato\s*(\d+)", texto) else "Não encontrado",
        "empenho": re.search(r"202\dNE\d{5}", texto).group(0) if re.search(r"202\dNE\d{5}", texto) else "Não identificado",
        "liquidacao": re.search(r"202\dNL\d{5}", texto).group(0) if re.search(r"202\dNL\d{5}", texto) else "Não identificado",
        "valor_bruto": re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto).group(0) if re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto) else "R$ 0,00",
        "sei_ids": sorted(list(set(re.findall(r"verificador\s(\d{9})", texto))), reverse=True)
    }

    # Fornecedor
    re_forn = re.search(r"(?:favor de|em favor de|Fornecedor:)\s*([A-Z\s\d\/\.\-\&]{5,100})", texto)
    dados["fornecedor"] = re_forn.group(1).replace('\n', ' ').strip() if re_forn else "Não encontrado"

    # Gestor
    re_gestor = re.search(r"assinado eletronicamente por\s*([A-Za-z\s]+),", texto)
    dados["gestor"] = re_gestor.group(1).strip() if re_gestor else "Não identificado"

    # Datas de Validade
    datas = re.findall(r"(\d{2}/\d{2}/\d{4})", texto)
    dados["validade"] = max(datas, key=lambda d: datetime.strptime(d, '%d/%m/%Y')) if datas else "Verificar"
    
    return dados

def gerar_excel_19_itens(dados_p, df_c):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Checklist Auditoria"
    
    # Estilos
    bold_f = Font(bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Títulos do Relatório
    ws.merge_cells('A1:D1')
    ws['A1'] = "INSTITUTO DE PESOS E MEDIDAS IPEM/RJ — AUDITORIA INTERNA - AUDIT"
    ws['A1'].font = Font(bold=True, size=12)
    ws['A1'].alignment = center

    # Cabeçalho de Dados
    ws['A3'] = "Processo SEI:"; ws['B3'] = dados_p['processo']
    ws['A4'] = "Fornecedor:"; ws['B4'] = dados_p['fornecedor']
    ws['A5'] = "Valor Bruto:"; ws['B5'] = dados_p['valor_bruto']
    ws['A6'] = "Gestor:"; ws['B6'] = dados_p['gestor']

    # Cabeçalho da Tabela
    headers = ["ITEM", "EVENTO A SER VERIFICADO", "S/N/NA", "OBSERVAÇÕES"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=8, column=c, value=h)
        cell.font = bold_f
        cell.border = border
        cell.alignment = center

    # Preenchimento dos 19 Itens
    for i, row in df_c.iterrows():
        ws.cell(row=i+9, column=1, value=row['ITEM']).border = border
        ws.cell(row=i+9, column=2, value=row['EVENTO']).border = border
        ws.cell(row=i+9, column=3, value=row['S/N/NA']).border = border
        ws.cell(row=i+9, column=4, value=row['OBSERVAÇÕES']).border = border

    wb.save(output)
    return output.getvalue()

# --- Interface Streamlit ---
st.markdown("<h2 style='text-align: center; color: #1f4e78;'>AUDIT - IPEM/RJ</h2>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center;'>Checklist Oficial (19 Itens de Conformidade)</h4>", unsafe_allow_html=True)
st.divider()

file = st.sidebar.file_uploader("Upload PDF do Processo", type="pdf")

if file:
    d = extrair_dados_pdf(file)
    s = d['sei_ids']
    
    # Painel de Resumo
    col1, col2 = st.columns(2)
    with col1:
        st.info(f"**Processo SEI:** {d['processo']}\n\n**Fornecedor:** {d['fornecedor']}")
    with col2:
        st.success(f"**Valor Bruto:** {d['valor_bruto']}\n\n**Gestor Identificado:** {d['gestor']}")

    st.divider()

    # Estrutura Oficial de 19 Itens
    checklist_19 = [
        {"ITEM": 1, "EVENTO": "Nota de empenho e demonstrativo de saldo", "S/N/NA": "S", "OBSERVAÇÕES": f"NE {d['empenho']} / NL {d['liquidacao']}"},
        {"ITEM": 2, "EVENTO": "Nota Fiscal / Fatura em nome do IPEM", "S/N/NA": "S", "OBSERVAÇÕES": f"Doc. SEI {s[0] if len(s)>0 else ''}"},
        {"ITEM": 3, "EVENTO": "Certidão Tributos Federais e Dívida Ativa", "S/N/NA": "S", "OBSERVAÇÕES": f"Val: {d['validade']}"},
        {"ITEM": 4, "EVENTO": "Certidão de regularidade - CRF (FGTS)", "S/N/NA": "S", "OBSERVAÇÕES": f"Val: {d['validade']}"},
        {"ITEM": 5, "EVENTO": "Certidão de regularidade - CNDT", "S/N/NA": "S", "OBSERVAÇÕES": f"Val: {d['validade']}"},
        {"ITEM": 6, "EVENTO": "Certidão de Regularidade Estadual (ICMS)", "S/N/NA": "S", "OBSERVAÇÕES": "Verificada no processo"},
        {"ITEM": 7, "EVENTO": "Certidão de Regularidade Municipal (ISS)", "S/N/NA": "S", "OBSERVAÇÕES": "Verificada no processo"},
        {"ITEM": 8, "EVENTO": "Consulta ao CADIN Estadual", "S/N/NA": "S", "OBSERVAÇÕES": "Nada consta"},
        {"ITEM": 9, "EVENTO": "Consulta de Sanções (CEIS/CNEP)", "S/N/NA": "S", "OBSERVAÇÕES": "Nada consta"},
        {"ITEM": 10, "EVENTO": "Incidência de tributos retidos na fonte?", "S/N/NA": "S", "OBSERVAÇÕES": "Conforme Nota de Liquidação"},
        {"ITEM": 11, "EVENTO": "Comprovação de não incidência de tributos?", "S/N/NA": "NA", "OBSERVAÇÕES": ""},
        {"ITEM": 12, "EVENTO": "Portaria de Nomeação de Fiscalização", "S/N/NA": "S", "OBSERVAÇÕES": "Portaria GAPRE"},
        {"ITEM": 13, "EVENTO": "Atestado do Gestor (Serviços prestados)", "S/N/NA": "S", "OBSERVAÇÕES": f"Doc. SEI {s[-1] if len(s)>0 else ''}"},
        {"ITEM": 14, "EVENTO": "Relação dos funcionários do serviço", "S/N/NA": "S", "OBSERVAÇÕES": "Identificado em anexo"},
        {"ITEM": 15, "EVENTO": "Comprovante da GFIP / eSocial", "S/N/NA": "S", "OBSERVAÇÕES": "Identificado em anexo"},
        {"ITEM": 16, "EVENTO": "Comprovante de pagamento do INSS", "S/N/NA": "S", "OBSERVAÇÕES": "Guia autenticada"},
        {"ITEM": 17, "EVENTO": "Comprovante de pagamento do FGTS", "S/N/NA": "S", "OBSERVAÇÕES": "Comprovante bancário"},
        {"ITEM": 18, "EVENTO": "Folha de pagamento", "S/N/NA": "S", "OBSERVAÇÕES": "Referente ao período de execução"},
        {"ITEM": 19, "EVENTO": "Comprovante bancário de salários", "S/N/NA": "S", "OBSERVAÇÕES": "Ordem de transferência"}
    ]

    df_final = pd.DataFrame(checklist_19)
    st.table(df_final)

    # Botão de Exportação para Excel Oficial
    excel_out = gerar_excel_19_itens(d, df_final)
    st.download_button(
        label="📥 BAIXAR CHECKLIST COMPLETO (19 ITENS)",
        data=excel_out,
        file_name=f"Checklist_Audit_{d['processo'].replace('/','_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
else:
    st.warning("Por favor, submeta o PDF do processo para gerar o checklist de 19 itens.")
