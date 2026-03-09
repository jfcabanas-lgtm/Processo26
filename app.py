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
    
    # Regex para capturar o Fornecedor (busca por nomes em maiúsculo próximos a indicadores de empresa)
    # Tenta capturar o nome que vem antes de "CNPJ" ou após "favor da empresa"
    re_fornecedor = re.search(r"(?:empresa|Favor do Credor|Favorecido):\s?([A-Z\s]{5,50})", texto, re.IGNORECASE)
    if not re_fornecedor:
        # Alternativa: capturar linhas que contenham LTDA ou EIRELI
        re_fornecedor = re.search(r"([A-Z\s\d.-]+(?:LTDA|EIRELI|S\.A|S/A|ME|EPP))", texto)

    dados = {
        "processo": re.search(r"\d{6}/\d{6}/\d{4}", texto),
        "fornecedor": re_fornecedor,
        "cnpj": re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto),
        "empenho": re.search(r"202\dNE\d{5}", texto),
        "liquidacao": re.search(r"202\dNL\d{5}", texto),
        "valor_bruto": re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto),
        "sei_docs": re.findall(r"verificador\s(\d{9})", texto)
    }
    
    # Limpeza dos resultados
    for k, v in dados.items():
        if isinstance(v, list):
            dados[k] = list(set(v))
        elif v and hasattr(v, 'group'):
            dados[k] = v.group(1).strip() if k == "fornecedor" else v.group(0).strip()
        else:
            dados[k] = "Não identificado"
            
    return dados, texto

# Interface Streamlit
st.title("🔍 Auditoria Inteligente - IPEM/RJ")
st.markdown("---")

st.sidebar.header("Configurações")
uploaded_file = st.sidebar.file_uploader("Submeter Processo SEI (PDF)", type="pdf")

if uploaded_file:
    with st.spinner('Analisando documentos...'):
        dados, texto_completo = extrair_dados_pdf(uploaded_file)
        
    # Exibição de Resumo com o nome do Fornecedor em destaque
    st.subheader(f"📄 Análise: {dados['fornecedor']}")
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Processo", dados["processo"])
    col2.metric("CNPJ", dados["cnpj"])
    col3.metric("Valor Bruto", dados["valor_bruto"])
    col4.metric("Empenho", dados["empenho"])

    st.divider()

    # Checklist Dinâmico
    st.subheader("📋 Checklist de Conformidade")
    
    status_docs = "✅ Presente" if len(dados["sei_docs"]) > 5 else "⚠️ Verificar anexos"
    
    checklist_data = {
        "Item": [1, 2, 3, 4, 5],
        "Documento": ["Nota de Empenho", "Identificação do Fornecedor", "Nota Fiscal", "Certidões de Regularidade", "Nota de Liquidação"],
        "Status": [
            "✅ Identificado" if dados["empenho"] != "Não identificado" else "❌ Ausente",
            "✅ Identificado" if dados["fornecedor"] != "Não identificado" else "⚠️ Conferir Manuscrito",
            "✅ Identificado" if dados["valor_bruto"] != "Não identificado" else "❌ Ausente",
            status_docs,
            "✅ Identificado" if dados["liquidacao"] != "Não identificado" else "❓ Pendente"
        ],
        "Evidência": [
            dados["empenho"], 
            dados["fornecedor"], 
            "Valor: " + dados["valor_bruto"], 
            f"{len(dados['sei_docs'])} IDs SEI encontrados", 
            dados["liquidacao"]
        ]
    }
    
    st.table(pd.DataFrame(checklist_data))

    with st.expander("Visualizar Texto Extraído"):
        st.text(texto_completo)

    st.subheader("✍️ Sugestão de Despacho")
    st.code(f"""Conforme análise do processo {dados['processo']}, referente ao fornecedor {dados['fornecedor']}, 
verificou-se a regularidade da despesa. A Nota de Empenho {dados['empenho']} e a 
Liquidação {dados['liquidacao']} encontram-se devidamente instruídas.""")

else:
    st.info("Aguardando upload do processo PDF para iniciar a auditoria...")
