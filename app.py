import streamlit as st
import pandas as pd
import PyPDF2
import re
from datetime import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

# Configuração da página
st.set_page_config(page_title="AuditAI - IPEM/RJ", layout="wide", page_icon="🛡️")

def extrair_dados_pdf(file):
    texto = ""
    reader = PyPDF2.PdfReader(file)
    for page in reader.pages:
        content = page.extract_text()
        if content: texto += content
    
    # Extração de dados com Regex
    dados = {
        "processo": re.search(r"\d{6}/\d{6}/\d{4}", texto).group(0) if re.search(r"\d{6}/\d{6}/\d{4}", texto) else "",
        "cnpj": re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto).group(0) if re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto) else "",
        "empenho": re.search(r"202\dNE\d{5}", texto).group(0) if re.search(r"202\dNE\d{5}", texto) else "",
        "liquidacao": re.search(r"202\dNL\d{5}", texto).group(0) if re.search(r"202\dNL\d{5}", texto) else "",
        "valor": re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto).group(0) if re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto) else "",
        "sei_doc": re.search(r"verificador\s(\d{9})", texto).group(1) if re.search(r"verificador\s(\d{9})", texto) else ""
    }

    re_forn = re.search(r"(?:empresa|favor da empresa|Credor|Favorecido|Fornecedor):\s*([A-Z\s\d\/\.\-\&]{5,100})", texto, re.IGNORECASE)
    dados["fornecedor"] = re_forn.group(1).replace('\n', ' ').strip() if re_forn else ""

    datas = re.findall(r"(\d{2}/\d{2}/\d{4})", texto)
    dados["validade"] = max(datas, key=lambda d: datetime.strptime(d, '%d/%m/%Y')) if datas else ""
    
    return dados

def gerar_excel_oficial(dados_identificados, df_checklist):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Checklist Audit"

    # Cabeçalho formatado
    ws.merge_cells('A1:D1')
    ws['A1'] = "CHECKLIST DA DOCUMENTAÇÃO APRESENTADA DE PROCESSO DE DESPESA"
    ws['A1'].font = Font(bold=True, size=12)
    ws['A1'].alignment = Alignment(horizontal='center')

    # Dados do Processo
    ws['A3'] = "PROCESSO SEI:"; ws['B3'] = dados_identificados['processo']
    ws['A4'] = "FORNECEDOR:"; ws['B4'] = dados_identificados['fornecedor']
    ws['A5'] = "CNPJ:"; ws['B5'] = dados_identificados['cnpj']
    ws['A6'] = "VALOR:"; ws['B6'] = dados_identificados['valor']

    # Tabela
    headers = ["ITEM", "EVENTO A SER VERIFICADO", "S/N/NA", "OBSERVAÇÕES"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=8, column=c, value=h).font = Font(bold=True)

    for i, row in df_checklist.iterrows():
        ws.cell(row=i+9, column=1, value=row['ITEM'])
        ws.cell(row=i+9, column=2, value=row['EVENTO A SER VERIFICADO'])
        ws.cell(row=i+9, column=3, value=row['S/N/NA'])
        ws.cell(row=i+9, column=4, value=row['OBSERVAÇÕES'])

    wb.save(output)
    return output.getvalue()

# --- INTERFACE ---
st.title("🛡️ AUDIT - IPEM/RJ")
st.markdown("### Checklist Oficial de Despesa")

uploaded_file = st.sidebar.file_uploader("Upload Processo SEI (PDF)", type="pdf")

if uploaded_file:
    dados = extrair_dados_pdf(uploaded_file)
    
    # Exibição de Informações Gerais em Tópicos
    st.markdown("#### 📌 Informações Extraídas")
    c1, c2 = st.columns(2)
    with c1:
        st.write(f"**Processo:** {dados['processo']}")
        st.write(f"**Fornecedor:** {dados['fornecedor']}")
    with c2:
        st.write(f"**CNPJ:** {dados['cnpj']}")
        st.write(f"**Empenho:** {dados['empenho']}")

    st.divider()

    # --- MONTAGEM DO CHECKLIST IGUAL AO EXCEL ---
    st.markdown("#### 📋 Checklist de Auditoria")
    
    lista_itens = [
        [1, "Nota de empenho e demonstrativo de saldo", "S", f"NE {dados['empenho']}"],
        [2, "Nota Fiscal de acordo com o empenho", "S", f"Doc SEI {dados['sei_doc']}"],
        [3, "Certidão Tributos Federais / Receita Federal", "S", f"Válida até: {dados['validade']}"],
        [4, "Certidão de regularidade junto ao FGTS", "S", f"Válida até: {dados['validade']}"],
        [5, "Certidão junto a Justiça do Trabalho", "S", f"Válida até: {dados['validade']}"],
        [6, "Incidência de tributos retidos na fonte", "NA", ""],
        [7, "Comprovação de não incidência de tributos", "NA", ""],
        [8, "Portaria de Nomeação de Fiscalização", "S", "Portaria IPEM/GAPRE Nº"],
        [9, "Atestado do Gestor (Serviço Prestado)", "S", f"Doc SEI {dados['sei_doc']}"],
        [10, "Relação dos funcionários que executaram serviço", "S", ""],
        [11, "Comprovante da GFIP", "S", ""],
        [12, "Comprovante de pagamento do INSS", "S", ""],
        [13, "Comprovante de pagamento do FGTS", "S", ""],
        [14, "Protocolo de envio - Conectividade Social", "S", ""],
        [15, "Folha de pagamento", "S", ""],
        [16, "Comprovante bancário de pagamento de salários", "S", ""]
    ]

    df_checklist = pd.DataFrame(lista_itens, columns=["ITEM", "EVENTO A SER VERIFICADO", "S/N/NA", "OBSERVAÇÕES"])
    
    # Exibe a tabela na tela para conferência
    st.table(df_checklist)

    # Botão para gerar o documento oficial
    st.divider()
    excel_file = gerar_excel_oficial(dados, df_checklist)
    
    st.download_button(
        label="📥 GERAR RELATÓRIO OFICIAL (EXCEL)",
        data=excel_file,
        file_name=f"Checklist_Audit_{dados['processo'].replace('/','_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

else:
    st.info("Aguardando upload do processo PDF para gerar o checklist.")
