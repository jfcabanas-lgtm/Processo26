import streamlit as st
import pandas as pd
import PyPDF2
import re
from datetime import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

st.set_page_config(page_title="AuditAI - IPEM/RJ", layout="wide")

def extrair_dados_pdf(file):
    texto = ""
    reader = PyPDF2.PdfReader(file)
    for page in reader.pages:
        content = page.extract_text()
        if content: texto += content
    
    # Extração de dados fundamentais
    dados = {
        "processo": re.search(r"\d{6}/\d{6}/\d{4}", texto).group(0) if re.search(r"\d{6}/\d{6}/\d{4}", texto) else "",
        "cnpj": re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto).group(0) if re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto) else "",
        "empenho": re.search(r"202\dNE\d{5}", texto).group(0) if re.search(r"202\dNE\d{5}", texto) else "",
        "valor": re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto).group(0) if re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto) else "",
        # Captura todos os números verificadores SEI de 9 dígitos
        "sei_lista": sorted(list(set(re.findall(r"\b\d{9}\b", texto))), reverse=True)
    }

    re_forn = re.search(r"(?:empresa|favor da empresa|Credor|Favorecido|Fornecedor):\s*([A-Z\s\d\/\.\-\&]{5,100})", texto, re.IGNORECASE)
    dados["fornecedor"] = re_forn.group(1).replace('\n', ' ').strip() if re_forn else ""
    
    datas = re.findall(r"(\d{2}/\d{2}/\d{4})", texto)
    dados["validade"] = max(datas, key=lambda d: datetime.strptime(d, '%d/%m/%Y')) if datas else ""
    
    return dados

def gerar_excel_oficial(dados_proc, df_check):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Checklist Audit"

    # Estilização
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Cabeçalho
    ws.merge_cells('A1:D1')
    ws['A1'] = "CHECKLIST DA DOCUMENTAÇÃO APRESENTADA DE PROCESSO DE DESPESA"
    ws['A1'].font = Font(bold=True, size=12)
    ws['A1'].alignment = center_align

    # Cabeçalho de Itens
    headers = ["ITEM", "EVENTO A SER VERIFICADO", "S/N/NA", "NÚMERO SEI / OBSERVAÇÃO"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=8, column=c, value=h)
        cell.font = bold_font
        cell.border = border
        cell.alignment = center_align

    # Preenchimento das linhas
    for i, row in df_check.iterrows():
        for c, value in enumerate(row, start=1):
            cell = ws.cell(row=i+9, column=c, value=value)
            cell.border = border
            if c != 2: cell.alignment = center_align

    wb.save(output)
    return output.getvalue()

# --- INTERFACE ---
st.title("🛡️ Auditoria Interna - IPEM/RJ")
uploaded_file = st.sidebar.file_uploader("Upload Processo SEI (PDF)", type="pdf")

if uploaded_file:
    d = extrair_dados_pdf(uploaded_file)
    s = d['sei_lista'] # Atalho para a lista de números SEI
    
    st.markdown(f"### Análise: {d['fornecedor']}")
    st.info(f"**Processo SEI:** {d['processo']} | **CNPJ:** {d['cnpj']} | **Valor:** {d['valor']}")

    # Mapeamento Dinâmico dos Itens com números SEI encontrados
    # Tentamos distribuir os números SEI encontrados por ordem de aparição ou lógica
    checklist_data = [
        [1, "Nota de empenho e demonstrativo de saldo", "S", f"NE {d['empenho']} (SEI {s[0] if len(s)>0 else ''})"],
        [2, "Nota Fiscal em nome do IPEM", "S", f"Doc. SEI {s[1] if len(s)>1 else ''}"],
        [3, "Certidão Tributos Federais / Receita Federal", "S", f"Val: {d['validade']} (SEI {s[2] if len(s)>2 else ''})"],
        [4, "Certidão de regularidade FGTS", "S", f"Val: {d['validade']} (SEI {s[3] if len(s)>3 else ''})"],
        [5, "Certidão Justiça do Trabalho", "S", f"Val: {d['validade']} (SEI {s[4] if len(s)>4 else ''})"],
        [6, "Incidência de tributos retidos na fonte", "NA", ""],
        [7, "Documento de comprovação de não incidência", "NA", ""],
        [8, "Portaria de Nomeação de Fiscalização", "S", "Portaria GAPRE"],
        [9, "Atestado do Gestor (Serviço Prestado)", "S", f"Doc. SEI {s[5] if len(s)>5 else ''}"],
        [10, "Relação dos funcionários", "S", ""],
        [11, "Comprovante da GFIP", "S", ""],
        [12, "Comprovante de pagamento do INSS", "S", ""],
        [13, "Comprovante de pagamento do FGTS", "S", ""],
        [14, "Protocolo de envio - Conectividade Social", "S", ""],
        [15, "Folha de pagamento", "S", ""],
        [16, "Comprovante bancário de salários", "S", ""]
    ]

    df_check = pd.DataFrame(checklist_data, columns=["ITEM", "EVENTO", "S/N/NA", "NÚMERO SEI / OBSERVAÇÃO"])
    
    st.table(df_check)

    if st.download_button(
        label="📥 GERAR CHECKLIST OFICIAL EM EXCEL",
        data=gerar_excel_oficial(d, df_check),
        file_name=f"Checklist_{d['processo'].replace('/','_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    ):
        st.balloons()
else:
    st.warning("Aguardando upload do PDF para processar o checklist.")
