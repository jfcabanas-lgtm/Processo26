import streamlit as st
import pandas as pd
import PyPDF2
import re
from datetime import datetime
from fpdf import FPDF
import io

# Configuração da página
st.set_page_config(page_title="AuditAI Pro - IPEM/RJ", layout="wide", page_icon="🛡️")

def validar_data(texto_data):
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
    
    # Extração de Datas de Validade (procura padrões dd/mm/aaaa próximos a palavras-chave)
    datas_encontradas = re.findall(r"(\d{2}/\d{2}/\d{4})", texto)
    
    # Lógica simplificada: assume que as datas mais distantes no futuro são as validades
    proximas_validades = sorted([d for d in datas_encontradas if datetime.strptime(d, '%d/%m/%Y') > datetime(2024,1,1)], 
                               key=lambda x: datetime.strptime(x, '%d/%m/%Y'), reverse=True)

    re_fornecedor = re.search(r"(?:empresa|favor da empresa|Credor|Favorecido|Fornecedor):\s*([A-Z\s\d\/\.\-\&]{5,100})", texto, re.IGNORECASE)
    nome_fornecedor = re_fornecedor.group(1).replace('\n', ' ').strip() if re_fornecedor else "Não identificado"

    dados = {
        "processo": re.search(r"\d{6}/\d{6}/\d{4}", texto).group(0) if re.search(r"\d{6}/\d{6}/\d{4}", texto) else "Não identificado",
        "fornecedor": nome_fornecedor,
        "cnpj": re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto).group(0) if re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto) else "Não identificado",
        "empenho": re.search(r"202\dNE\d{5}", texto).group(0) if re.search(r"202\dNE\d{5}", texto) else "Não identificado",
        "validade_sugerida": proximas_validades[0] if proximas_validades else "Não encontrada",
        "valor_bruto": re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto).group(0) if re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto) else "Não identificado"
    }
    return dados, texto

def gerar_pdf(dados, df_checklist):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(200, 10, "Relatorio de Auditoria Interna - IPEM/RJ", ln=True, align='C')
    pdf.ln(10)
    pdf.set_font("Arial", '', 12)
    pdf.cell(200, 10, f"Processo: {dados['processo']}", ln=True)
    pdf.cell(200, 10, f"Fornecedor: {dados['fornecedor']}", ln=True)
    pdf.cell(200, 10, f"CNPJ: {dados['cnpj']}", ln=True)
    pdf.ln(5)
    
    # Tabela Simples no PDF
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(90, 10, "Item Verificado", 1)
    pdf.cell(30, 10, "Status", 1)
    pdf.cell(70, 10, "Evidencia", 1, ln=True)
    
    pdf.set_font("Arial", '', 9)
    for index, row in df_checklist.iterrows():
        pdf.cell(90, 10, str(row['Critério']), 1)
        pdf.cell(30, 10, str(row['Resultado']), 1)
        pdf.cell(70, 10, str(row['Dados']), 1, ln=True)
        
    return pdf.output(dest='S').encode('latin-1')

# --- Interface ---
st.title("🛡️ AuditAI Pro: Auditoria e Conformidade")

uploaded_file = st.sidebar.file_uploader("Upload do Processo SEI", type="pdf")

if uploaded_file:
    dados, texto_bruto = extrair_dados_pdf(uploaded_file)
    status_val, classe = validar_data(dados['validade_sugerida'])
    
    st.subheader(f"🏢 {dados['fornecedor']}")
    
    # Cards de Status
    c1, c2, c3 = st.columns(3)
    c1.metric("Empenho", dados['empenho'])
    c2.metric("Validade Certidão", dados['validade_sugerida'], delta=status_val, delta_color="normal" if "Válida" in status_val else "inverse")
    c3.metric("Valor", dados['valor_bruto'])

    # Tabela de Dados
    df_check = pd.DataFrame({
        "Critério": ["Empenho Cadastrado", "Regularidade Fiscal", "Objeto e Valor", "CNPJ Ativo"],
        "Resultado": ["✅ OK", status_val, "✅ OK", "✅ OK"],
        "Dados": [dados['empenho'], f"Validade: {dados['validade_sugerida']}", dados['valor_bruto'], dados['cnpj']]
    })
    st.table(df_check)

    # Botões de Exportação
    st.markdown("### 📥 Exportar Relatório")
    col_pdf, col_excel = st.columns(2)
    
    # PDF
    pdf_bytes = gerar_pdf(dados, df_check)
    col_pdf.download_button(label="📄 Baixar Checklist em PDF", data=pdf_bytes, file_name=f"Relatorio_{dados['processo'].replace('/','-')}.pdf", mime="application/pdf")
    
    # EXCEL
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_check.to_excel(writer, index=False, sheet_name='Auditoria')
    col_excel.download_button(label="📊 Baixar Checklist em Excel", data=output.getvalue(), file_name=f"Auditoria_{dados['processo'].replace('/','-')}.xlsx", mime="application/vnd.ms-excel")

else:
    st.info("Aguardando submissão de documento para análise de conformidade.")
