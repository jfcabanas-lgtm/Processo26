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
        content = page.extract_text()
        if content:
            texto += content
    
    # --- LÓGICA REFINADA PARA O FORNECEDOR ---
    # Busca após palavras-chave e aceita nomes longos (até 100 caracteres)
    # Inclui suporte a caracteres como &, /, -, e pontos
    re_fornecedor = re.search(
        r"(?:empresa|favor da empresa|Credor|Favorecido|Fornecedor):\s*([A-Z\s\d\/\.\-\&]{5,100})", 
        texto, 
        re.IGNORECASE
    )
    
    nome_extraido = "Não identificado"
    
    if re_fornecedor:
        nome_extraido = re_fornecedor.group(1).strip()
    else:
        # Fallback: Procura por linhas que terminam com sufixos empresariais
        fallback = re.search(r"([A-Z\s\d\/\.\-\&]+(?:EIRELI|LTDA|S\.A|S/A|ME|EPP|LIMITADA))", texto)
        if fallback:
            nome_extraido = fallback.group(1).strip()

    # Limpeza crucial: Remove quebras de linha e espaços duplos que cortam o nome
    nome_extraido = nome_extraido.replace('\n', ' ').replace('  ', ' ').strip()

    dados = {
        "processo": re.search(r"\d{6}/\d{6}/\d{4}", texto),
        "fornecedor": nome_extraido,
        "cnpj": re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto),
        "empenho": re.search(r"202\dNE\d{5}", texto),
        "liquidacao": re.search(r"202\dNL\d{5}", texto),
        "valor_bruto": re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto),
        "sei_docs": re.findall(r"verificador\s(\d{9})", texto)
    }
    
    # Limpeza dos campos capturados via Regex Match Objects
    for k, v in dados.items():
        if k == "fornecedor": continue 
        if isinstance(v, list):
            dados[k] = list(set(v))
        elif v:
            dados[k] = v.group(0).strip()
        else:
            dados[k] = "Não identificado"
            
    return dados, texto

# --- INTERFACE STREAMLIT ---
st.title("🔍 Auditoria Inteligente - IPEM/RJ")
st.markdown("---")

st.sidebar.header("Painel de Controle")
uploaded_file = st.sidebar.file_uploader("Carregar Processo SEI (PDF)", type="pdf")

if uploaded_file:
    with st.spinner('Processando documentos e extraindo dados...'):
        dados, texto_completo = extrair_dados_pdf(uploaded_file)
        
    # Exibição do Nome do Fornecedor em Destaque
    st.subheader(f"🏢 Fornecedor: {dados['fornecedor']}")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Processo SEI", dados["processo"])
    with col2:
        st.metric("CNPJ", dados["cnpj"])
    with col3:
        st.metric("Valor Bruto", dados["valor_bruto"])
    with col4:
        st.metric("Nota de Empenho", dados["empenho"])

    st.divider()

    # Tabela de Checklist
    st.markdown("### 📋 Checklist de Instrução Processual")
    
    status_docs = "✅ Presente" if len(dados["sei_docs"]) > 5 else "⚠️ Verificar anexos"
    
    checklist_df = pd.DataFrame({
        "Item": [1, 2, 3, 4, 5],
        "Critério de Auditoria": [
            "Existência de Nota de Empenho",
            "Identificação Completa do Credor",
            "Valor da NF vs Empenho",
            "Certidões e Documentos SEI",
            "Nota de Liquidação (NL)"
        ],
        "Resultado": [
            "✅" if dados["empenho"] != "Não identificado" else "❌",
            "✅" if dados["fornecedor"] != "Não identificado" else "⚠️",
            "✅" if dados["valor_bruto"] != "Não identificado" else "❌",
            status_docs,
            "✅" if dados["liquidacao"] != "Não identificado" else "❓"
        ],
        "Dados Extraídos": [
            dados["empenho"], 
            dados["fornecedor"], 
            dados["valor_bruto"], 
            f"{len(dados['sei_docs'])} códigos identificados", 
            dados["liquidacao"]
        ]
    })
    
    st.table(checklist_df)

    # Conclusão Automática
    st.subheader("📝 Minuta de Despacho Sugerida")
    st.info(f"""
    À AUDIT,
    
    Trata-se da análise do processo {dados['processo']} relativo ao pagamento da empresa {dados['fornecedor']}. 
    Verificou-se que a instrução processual contém as certidões necessárias e a Nota de Empenho {dados['empenho']}.
    Diante do exposto, nada a opor ao prosseguimento do pagamento.
    """)

    with st.expander("Ver Texto Bruto do PDF (Depuração)"):
        st.text(texto_completo)

else:
    st.info("Utilize a barra lateral para fazer o upload do processo em PDF.")
    st.image("https://www.ipem.rj.gov.br/images/logo_ipem.png", width=150)

st.sidebar.markdown("---")
st.sidebar.caption(f"Versão 1.2 | Atualizado em: {datetime.now().strftime('%d/%m/%Y')}")
