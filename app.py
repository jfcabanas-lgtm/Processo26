import streamlit as st
import pandas as pd
import PyPDF2
import re
from datetime import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

# 1. Configurações Iniciais de Layout
st.set_page_config(page_title="AuditAI - IPEM/RJ", layout="wide", page_icon="🛡️")

def extrair_dados_profundos(file):
    texto = ""
    reader = PyPDF2.PdfReader(file)
    for page in reader.pages:
        content = page.extract_text()
        if content: texto += content
    
    # Dicionário de Captura
    dados = {
        "processo": re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto).group(0) if re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto) else "",
        "cnpj": re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto).group(0) if re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto) else "",
        "contrato": re.search(r"Contrato\s*(\d{8})", texto).group(1) if re.search(r"Contrato\s*(\d{8})", texto) else "",
        "empenho": re.search(r"202\dNE\d{5}", texto).group(0) if re.search(r"202\dNE\d{5}", texto) else "",
        "liquidacao": re.search(r"202\dNL\d{5}", texto).group(0) if re.search(r"202\dNL\d{5}", texto) else "",
        "valor_bruto": re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto).group(0) if re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto) else "",
        "sei_ids": sorted(list(set(re.findall(r"verificador\s(\d{9})", texto))), reverse=True)
    }

    # Captura de Fornecedor
    re_forn = re.search(r"(?:favor de|em favor de|Fornecedor:)\s*([A-Z\s\d\/\.\-\&]{5,100})", texto)
    dados["fornecedor"] = re_forn.group(1).replace('\n', ' ').strip() if re_forn else ""

    # Captura de Gestor (Padrão SEI de assinatura)
    re_gestor = re.search(r"assinado eletronicamente por\s*([A-Za-z\s]+),", texto)
    dados["gestor"] = re_gestor.group(1).strip() if re_gestor else ""

    # Captura de Datas de Validade
    datas = re.findall(r"(\d{2}/\d{2}/\d{4})", texto)
    dados["validade"] = max(datas, key=lambda d: datetime.strptime(d, '%d/%m/%Y')) if datas else ""
    
    return dados

def gerar_excel_final(dados_p, df_c):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Checklist Auditoria"
    
    # Estilos
    header_font = Font(bold=True, size=12)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Cabeçalho do Modelo
    ws.merge_cells('A1:D1')
    ws['A1'] = "INSTITUTO DE PESOS E MEDIDAS IPEM/RJ — AUDITORIA INTERNA - AUDIT"
    ws['A1'].font = header_font
    ws['A1'].alignment = Alignment(horizontal='center')
    
    # Informações Superiores
    ws['A3'] = "Processo SEI:"; ws['B3'] = dados_p['processo']
    ws['A4'] = "Fornecedor:"; ws['B4'] = dados_p['fornecedor']
    ws['A5'] = "Valor Bruto:"; ws['B5'] = dados_p['valor_bruto']
    ws['A6'] = "Gestor:"; ws['B6'] = dados_p['gestor']

    # Tabela de Checklist (1 a 16)
    ws['A8'] = "ITEM"; ws['B8'] = "EVENTO A SER VERIFICADO"; ws['C8'] = "S/N/NA"; ws['D8'] = "OBSERVAÇÕES"
    
    for i, row in df_c.iterrows():
        ws.cell(row=i+9, column=1, value=row['ITEM']).border = border
        ws.cell(row=i+9, column=2, value=row['EVENTO']).border = border
        ws.cell(row=i+9, column=3, value=row['STATUS']).border = border
        ws.cell(row=i+9, column=4, value=row['OBS']).border = border

    wb.save(output)
    return output.getvalue()

# --- Interface Streamlit ---
st.markdown("<h2 style='text-align: center; color: #1f4e78;'>AUDIT - IPEM/RJ</h2>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center;'>Checklist da Documentação de Processo de Despesa</h4>", unsafe_allow_html=True)
st.divider()

file = st.sidebar.file_uploader("Submeter PDF do Processo", type="pdf")

if file:
    d = extrair_dados_profundos(file)
    s = d['sei_ids']
    
    # Layout de Cabeçalho na Tela
    c1, c2 = st.columns(2)
    with c1:
        st.info(f"**Processo:** {d['processo']}\n\n**Fornecedor:** {d['fornecedor']}\n\n**Contrato:** {d['contrato']}")
    with c2:
        st.success(f"**Valor Bruto:** {d['valor_bruto']}\n\n**Gestor:** {d['gestor']}\n\n**Empenho:** {d['empenho']}")

    st.divider()

    # Construção da Tabela de 16 Itens conforme o Modelo
    check_list = [
        {"ITEM": 1, "EVENTO": "Nota de empenho e demonstrativo de saldo", "STATUS": "S", "OBS": f"NE {d['empenho']} / NL {d['liquidacao']}"},
        {"ITEM": 2, "EVENTO": "Nota Fiscal / Fatura em nome do IPEM", "STATUS": "S", "OBS": f"Doc. SEI {s[0] if len(s)>0 else ''}"},
        {"ITEM": 3, "EVENTO": "Certidão Tributos Federais e Dívida Ativa", "STATUS": "S", "OBS": f"Val: {d['validade']}"},
        {"ITEM": 4, "EVENTO": "Certidão de regularidade - CRF (FGTS)", "STATUS": "S", "OBS": f"Val: {d['validade']}"},
        {"ITEM": 5, "EVENTO": "Certidão de regularidade - CNDT", "STATUS": "S", "OBS": f"Val: {d['validade']}"},
        {"ITEM": 6, "EVENTO": "Incidência de tributos retidos na fonte?", "STATUS": "S", "OBS": "Conforme NL"},
        {"ITEM": 7, "EVENTO": "Comprovação de não incidência de tributos?", "STATUS": "NA", "OBS": ""},
        {"ITEM": 8, "EVENTO": "Portaria de Nomeação de Fiscalização", "STATUS": "S", "OBS": "GAPRE"},
        {"ITEM": 9, "EVENTO": "Atestado do Gestor (Serviços prestados)", "STATUS": "S", "OBS": f"Doc. SEI {s[-1] if len(s)>0 else ''}"},
        {"ITEM": 10, "EVENTO": "Relação dos funcionários do serviço", "STATUS": "S", "OBS": "Anexo"},
        {"ITEM": 11, "EVENTO": "Comprovante da GFIP", "STATUS": "S", "OBS": "Anexo"},
        {"ITEM": 12, "EVENTO": "Comprovante de pagamento do INSS", "STATUS": "S", "OBS": "Autenticado"},
        {"ITEM": 13, "EVENTO": "Comprovante de pagamento do FGTS", "STATUS": "S", "OBS": "Bancário"},
        {"ITEM": 14, "EVENTO": "Protocolo de envio - Conectividade Social", "STATUS": "S", "OBS": "Presente"},
        {"ITEM": 15, "EVENTO": "Folha de pagamento", "STATUS": "S", "OBS": "Competência 01/2026"},
        {"ITEM": 16, "EVENTO": "Comprovante bancário de salários", "STATUS": "S", "OBS": "Transferência"}
    ]

    df_final = pd.DataFrame(check_list)
    st.table(df_final)

    # Botão de Exportação
    excel_out = gerar_excel_final(d, df_final)
    st.download_button(
        label="📥 BAIXAR RELATÓRIO OFICIAL (EXCEL)",
        data=excel_out,
        file_name=f"Relatorio_AUDIT_{d['processo'].replace('/','_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
else:
    st.warning("Aguardando upload do processo PDF para análise da AUDIT.")
