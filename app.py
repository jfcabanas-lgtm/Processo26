import streamlit as st
import pandas as pd
import PyPDF2
import re
from datetime import datetime
from fpdf import FPDF
import io

# Configuração da página
st.set_page_config(page_title="AuditAI Pro - IPEM/RJ", layout="wide", page_icon="🛡️")

def limpar_texto_pdf(txt):
    """Remove caracteres que causam erro na geração do PDF (latin-1)"""
    if not txt: return ""
    return str(txt).encode('latin-1', 'replace').decode('latin-1').replace('?', '')

def validar_data(texto_data):
    """Compara a data extraída com a data atual"""
    try:
        data_obj = datetime.strptime(texto_data, '%d/%m/%Y')
        if data_obj >= datetime.now():
            return "✅ Válida", "Normal"
        else:
            return "❌ Vencida", "Critico"
    except:
        return "⚠️ Não identificada", "Alerta"

def extrair_dados_pdf(file):
    texto = ""
    reader = PyPDF2.PdfReader(file)
    for page in reader.pages:
        content = page.extract_text()
        if content: texto += content
    
    # Extração de Dados com Regex
    re_processo = re.search(r"\d{6}/\d{6}/\d{4}", texto)
    re_cnpj = re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto)
    re_empenho = re.search(r"202\dNE\d{5}", texto)
    re_liquidacao = re.search(r"202\dNL\d{5}", texto)
    re_valor = re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto)
    
    # Busca Fornecedor (Nome longo e limpeza de quebra de linha)
    re_forn = re.search(r"(?:empresa|favor da empresa|Credor|Favorecido|Fornecedor):\s*([A-Z\s\d\/\.\-\&]{5,100})", texto, re.IGNORECASE)
    nome_fornecedor = re_forn.group(1).replace('\n', ' ').strip() if re_forn else "Não identificado"

    # Busca todas as datas e tenta isolar a de validade (geralmente a maior/futura)
    datas = re.findall(r"(\d{2}/\d{2}/\d{4})", texto)
    validade = "Não encontrada"
    if datas:
        datas_obj = []
        for d in datas:
            try: datas_obj.append(datetime.strptime(d, '%d/%m/%Y'))
            except: continue
        if datas_obj:
            validade = max(datas_obj).strftime('%d/%m/%Y')

    return {
        "processo": re_processo.group(0) if re_processo else "Não identificado",
        "fornecedor": nome_fornecedor,
        "cnpj": re_cnpj.group(0) if re_cnpj else "Não identificado",
        "empenho": re_empenho.group(0) if re_empenho else "Não identificado",
        "liquidacao": re_liquidacao.group(0) if re_liquidacao else "Não identificado",
        "valor_bruto": re_valor.group(0) if re_valor else "Não identificado",
        "validade": validade,
        "sei_docs": list(set(re.findall(r"verificador\s(\d{9})", texto)))
    }, texto

def gerar_pdf_bytes(dados, df):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(190, 10, "IPEM/RJ - AUDITORIA INTERNA (AUDIT)", ln=True, align='C')
    pdf.set_font("Arial", '', 11)
    pdf.ln(5)
    
    # Cabeçalho do Relatório
    pdf.cell(190, 7, f"Processo: {limpar_texto_pdf(dados['processo'])}", ln=True)
    pdf.cell(190, 7, f"Fornecedor: {limpar_texto_pdf(dados['fornecedor'])}", ln=True)
    pdf.cell(190, 7, f"CNPJ: {dados['cnpj']}", ln=True)
    pdf.cell(190, 7, f"Data da Auditoria: {datetime.now().strftime('%d/%m/%Y')}", ln=True)
    pdf.ln(5)
    
    # Tabela de Itens
    pdf.set_fill_color(240, 240, 240)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(85, 8, "Item Verificado", 1, 0, 'C', True)
    pdf.cell(30, 8, "Status", 1, 0, 'C', True)
    pdf.cell(75, 8, "Dados Identificados", 1, 1, 'C', True)
    
    pdf.set_font("Arial", '', 9)
    for _, row in df.iterrows():
        # Remove emojis para o PDF
        status_txt = str(row['Status']).replace("✅", "SIM").replace("❌", "NAO").replace("⚠️", "ALERTA")
        pdf.cell(85, 8, limpar_texto_pdf(row['Criterio']), 1)
        pdf.cell(30, 8, status_txt, 1, 0, 'C')
        pdf.cell(75, 8, limpar_texto_pdf(row['Evidencia']), 1, 1)

    return pdf.output()

# --- Interface Streamlit ---
st.title("🛡️ AuditAI Pro")
st.sidebar.header("Análise de Processo")
uploaded_file = st.sidebar.file_uploader("Upload PDF SEI", type="pdf")

if uploaded_file:
    dados, texto_completo = extrair_dados_pdf(uploaded_file)
    status_val, _ = validar_data(dados['validade'])
    
    st.header(f"🏢 {dados['fornecedor']}")
    
    # Resumo Superior
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Empenho", dados['empenho'])
    c2.metric("Liquidação", dados['liquidacao'])
    c3.metric("Validade Certidão", dados['validade'], delta=status_val, delta_color="normal" if "Válida" in status_val else "inverse")
    c4.metric("Valor NF", dados['valor_bruto'])

    st.divider()

    # Checklist
    df_check = pd.DataFrame([
        {"Criterio": "Nota de Empenho", "Status": "✅" if dados['empenho'] != "Não identificado" else "❌", "Evidencia": dados['empenho']},
        {"Criterio": "Regularidade (Validade)", "Status": status_val, "Evidencia": f"Vencimento: {dados['validade']}"},
        {"Criterio": "Nota Fiscal / Valor", "Status": "✅" if dados['valor_bruto'] != "Não identificado" else "❌", "Evidencia": dados['valor_bruto']},
        {"Criterio": "Nota de Liquidação", "Status": "✅" if dados['liquidacao'] != "Não identificado" else "❓", "Evidencia": dados['liquidacao']}
    ])
    
    st.subheader("📋 Checklist de Auditoria")
    st.table(df_check)

    # Exportação
    st.subheader("📥 Exportar Resultados")
    col_pdf, col_excel = st.columns(2)
    
    # PDF Download
    try:
        pdf_data = gerar_pdf_bytes(dados, df_check)
        col_pdf.download_button("📄 Baixar Relatório PDF", data=bytes(pdf_data), file_name=f"Audit_{dados['processo'].replace('/','_')}.pdf", mime="application/pdf")
    except Exception as e:
        col_pdf.error(f"Erro ao gerar PDF: {e}")

    # Excel Download
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_check.to_excel(writer, index=False, sheet_name='Checklist')
    col_excel.download_button("📊 Baixar Tabela Excel", data=output.getvalue(), file_name=f"Audit_{dados['processo'].replace('/','_')}.xlsx", mime="application/vnd.ms-excel")

else:
    st.info("Submeta o PDF do processo para iniciar.")
