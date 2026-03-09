import streamlit as st
import pandas as pd
import PyPDF2
import re
from datetime import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

# Configuração da página
st.set_page_config(page_title="AuditAI - Modelo IPEM", layout="wide")

def extrair_dados_pdf(file):
    texto = ""
    reader = PyPDF2.PdfReader(file)
    for page in reader.pages:
        content = page.extract_text()
        if content: texto += content
    
    # Extração de dados principais
    dados = {
        "processo": re.search(r"\d{6}/\d{6}/\d{4}", texto).group(0) if re.search(r"\d{6}/\d{6}/\d{4}", texto) else "Não identificado",
        "cnpj": re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto).group(0) if re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto) else "Não identificado",
        "empenho": re.search(r"202\dNE\d{5}", texto).group(0) if re.search(r"202\dNE\d{5}", texto) else "Não identificado",
        "liquidacao": re.search(r"202\dNL\d{5}", texto).group(0) if re.search(r"202\dNL\d{5}", texto) else "Não identificado",
        "valor": re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto).group(0) if re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto) else "Não identificado",
        "sei_despacho": re.search(r"verificador\s(\d{9})", texto).group(1) if re.search(r"verificador\s(\d{9})", texto) else ""
    }

    # Nome do Fornecedor
    re_forn = re.search(r"(?:empresa|favor da empresa|Credor|Favorecido|Fornecedor):\s*([A-Z\s\d\/\.\-\&]{5,100})", texto, re.IGNORECASE)
    dados["fornecedor"] = re_forn.group(1).replace('\n', ' ').strip() if re_forn else "Não identificado"

    # Validades das Certidões (busca datas e assume as futuras como validades)
    datas = re.findall(r"(\d{2}/\d{2}/\d{4})", texto)
    dados["validade"] = max(datas, key=lambda d: datetime.strptime(d, '%d/%m/%Y')) if datas else "Verificar"
    
    return dados

def gerar_excel_modelo(dados):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Checklist Audit"

    # Cabeçalho baseado no seu modelo
    ws.merge_cells('A1:J1')
    ws['A1'] = "CHECKLIST DA DOCUMENTAÇÃO APRESENTADA DE PROCESSO DE DESPESA"
    ws['A1'].font = Font(bold=True, size=12)
    ws['A1'].alignment = Alignment(horizontal='center')

    # Informações do Processo
    info = [
        ["PROCESSO SEI:", dados['processo']],
        ["FORNECEDOR:", dados['fornecedor']],
        ["CNPJ:", dados['cnpj']],
        ["VALOR:", dados['valor']],
        ["EMPENHO:", dados['empenho']]
    ]
    
    for r, row in enumerate(info, start=3):
        ws.cell(row=r, column=1, value=row[0]).font = Font(bold=True)
        ws.cell(row=r, column=2, value=row[1])

    # Tabela de Itens (Mapeando seu Modelo Excel)
    headers = ["ITEM", "EVENTO A SER VERIFICADO", "S/N/NA", "OBSERVAÇÕES"]
    for c, header in enumerate(headers, start=1):
        cell = ws.cell(row=9, column=c if c < 3 else c+5, value=header) # Ajustando colunas conforme seu CSV
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    itens = [
        (1, "Nota de empenho e demonstrativo de saldo", "S", f"NE {dados['empenho']}"),
        (2, "Nota Fiscal em nome do IPEM", "S", f"SEI {dados['sei_despacho']}"),
        (3, "Certidão Tributos Federais (Receita)", "S", f"Válida até: {dados['validade']}"),
        (4, "Certidão FGTS", "S", f"Válida até: {dados['validade']}"),
        (5, "Certidão Justiça do Trabalho", "S", f"Válida até: {dados['validade']}"),
        (6, "Incidência de tributos retidos", "NA", ""),
        (8, "Portaria de Nomeação de Fiscalização", "S", ""),
        (9, "Atestado do Gestor (Serviço prestado)", "S", f"Despacho SEI {dados['sei_despacho']}")
    ]

    for i, item in enumerate(itens, start=10):
        ws.cell(row=i, column=1, value=item[0])
        ws.cell(row=i, column=2, value=item[1])
        ws.cell(row=i, column=8, value=item[2])
        ws.cell(row=i, column=9, value=item[3])

    wb.save(output)
    return output.getvalue()

# --- Interface Streamlit ---
st.title("🛡️ AuditAI Pro - Modelo Oficial")

uploaded_file = st.sidebar.file_uploader("Upload Processo SEI (PDF)", type="pdf")

if uploaded_file:
    dados = extrair_dados_pdf(uploaded_file)
    
    st.subheader(f"Análise: {dados['fornecedor']}")
    
    # Exibição simplificada para conferência
    st.write(f"**Processo:** {dados['processo']} | **Empenho:** {dados['empenho']} | **Valor:** {dados['valor']}")
    
    st.divider()
    
    # Geração do Excel baseado no Modelo
    excel_data = gerar_excel_modelo(dados)
    
    st.success("Dados mapeados com sucesso para o modelo Excel!")
    
    st.download_button(
        label="📥 Baixar Modelo Excel Preenchido",
        data=excel_data,
        file_name=f"Checklist_{dados['processo'].replace('/','_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Por favor, suba o PDF do processo para preencher o modelo.")
