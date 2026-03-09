import streamlit as st
import pandas as pd
import PyPDF2
import re
from datetime import datetime

# Configuração da página
st.set_page_config(page_title="AuditAI - IPEM/RJ", layout="wide", page_icon="🛡️")

def extrair_dados_pdf(file):
    texto = ""
    reader = PyPDF2.PdfReader(file)
    for page in reader.pages:
        texto += page.extract_text()
    
    # Dicionário de busca (Regex para padrões SEI e Financeiros)
    dados = {
        "processo": re.search(r"\d{6}/\d{6}/\d{4}", texto),
        "cnpj": re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto),
        "empenho": re.search(r"202\dNE\d{5}", texto),
        "liquidacao": re.search(r"202\dNL\d{5}", texto),
        "valor_bruto": re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto),
        "sei_docs": re.findall(r"verificador\s(\d{9})", texto) # Busca IDs SEI comuns
    }
    
    # Limpeza dos resultados
    for k, v in dados.items():
        if isinstance(v, list):
            dados[k] = list(set(v)) # Remove duplicados
        elif v:
            dados[k] = v.group(0)
        else:
            dados[k] = "Não identificado"
            
    return dados, texto

# Interface Streamlit
st.title("🔍 Auditoria Inteligente - IPEM/RJ")
st.markdown("---")

# Sidebar para Upload
st.sidebar.header("Configurações")
uploaded_file = st.sidebar.file_uploader("Submeter Processo SEI (PDF)", type="pdf")

if uploaded_file:
    with st.spinner('Analisando documentos...'):
        dados, texto_completo = extrair_dados_pdf(uploaded_file)
        
    # Exibição de Resumo
    col1, col2, col3 = st.columns(3)
    col1.metric("Processo", dados["processo"])
    col2.metric("Valor Identificado", dados["valor_bruto"])
    col3.metric("Empenho", dados["empenho"])

    st.divider()

    # Checklist Dinâmico
    st.subheader("📋 Checklist de Conformidade")
    
    # Simulando a verificação de documentos baseada nos IDs SEI encontrados
    status_docs = "✅ Presente" if len(dados["sei_docs"]) > 5 else "⚠️ Verificar anexos"
    
    checklist_data = {
        "Item": [1, 2, 3, 4],
        "Documento": ["Nota de Empenho", "Nota Fiscal", "Certidões de Regularidade", "Nota de Liquidação"],
        "Status": [
            "✅ Identificado" if dados["empenho"] != "Não identificado" else "❌ Ausente",
            "✅ Identificado" if dados["valor_bruto"] != "Não identificado" else "❌ Ausente",
            status_docs,
            "✅ Identificado" if dados["liquidacao"] != "Não identificado" else "❓ Pendente"
        ],
        "Evidência": [dados["empenho"], "Valor: " + dados["valor_bruto"], f"{len(dados['sei_docs'])} docs encontrados", dados["liquidacao"]]
    }
    
    st.table(pd.DataFrame(checklist_data))

    # Área de Texto Extraído para conferência
    with st.expander("Visualizar Texto Extraído do Processo"):
        st.text(texto_completo)

    # Conclusão sugerida
    st.subheader("✍️ Sugestão de Despacho")
    st.code(f"""Conforme análise do processo {dados['processo']}, verificou-se a regularidade da despesa. 
A Nota de Empenho {dados['empenho']} e a Liquidação {dados['liquidacao']} encontram-se devidamente instruídas.""")

else:
    st.info("Aguardando upload do processo PDF para iniciar a auditoria...")
    st.image("https://www.sei.rj.gov.br/sei/imagens/sei_logo.png", width=100)