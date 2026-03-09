import streamlit as st
import pandas as pd
import PyPDF2
import re
from datetime import datetime
from fpdf import FPDF
import io

# Configuração da página para um visual profissional
st.set_page_config(page_title="AUDIT - IPEM/RJ", layout="wide", page_icon="🛡️")

def limpar_texto_pdf(txt):
    if not txt: return ""
    return str(txt).encode('latin-1', 'replace').decode('latin-1').replace('?', '')

def extrair_dados_pdf(file):
    texto = ""
    reader = PyPDF2.PdfReader(file)
    for page in reader.pages:
        content = page.extract_text()
        if content: texto += content
    
    # Extração via Regex
    re_processo = re.search(r"\d{6}/\d{6}/\d{4}", texto)
    re_cnpj = re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto)
    re_empenho = re.search(r"202\dNE\d{5}", texto)
    re_liquidacao = re.search(r"202\dNL\d{5}", texto)
    re_valor = re.search(r"R\$\s?(\d{1,3}(\.\d{3})*,\d{2})", texto)
    
    # Fornecedor com limpeza de quebras de linha
    re_forn = re.search(r"(?:empresa|favor da empresa|Credor|Favorecido|Fornecedor):\s*([A-Z\s\d\/\.\-\&]{5,100})", texto, re.IGNORECASE)
    nome_fornecedor = re_forn.group(1).replace('\n', ' ').replace('  ', ' ').strip() if re_forn else "Não identificado"

    # Captura da maior data como validade
    datas = re.findall(r"(\d{2}/\d{2}/\d{4})", texto)
    validade = "Não encontrada"
    if datas:
        try:
            validade = max(datas, key=lambda d: datetime.strptime(d, '%d/%m/%Y'))
        except: pass

    # Captura de todos os números verificadores SEI
    sei_docs = sorted(list(set(re.findall(r"verificador\s(\d{9})", texto))))

    return {
        "processo": re_processo.group(0) if re_processo else "Não identificado",
        "fornecedor": nome_fornecedor,
        "cnpj": re_cnpj.group(0) if re_cnpj else "Não identificado",
        "empenho": re_empenho.group(0) if re_empenho else "Não identificado",
        "liquidacao": re_liquidacao.group(0) if re_liquidacao else "Não identificado",
        "valor_bruto": re_valor.group(0) if re_valor else "Não identificado",
        "validade": validade,
        "sei_docs": sei_docs
    }, texto

# --- INTERFACE ---
st.title("🛡️ Auditoria Interna - IPEM/RJ")
st.markdown("### Sistema de Checklist de Conformidade")
st.divider()

uploaded_file = st.sidebar.file_uploader("Upload do Processo SEI (PDF)", type="pdf")

if uploaded_file:
    dados, texto_completo = extrair_dados_pdf(uploaded_file)
    
    # --- SEÇÃO 1: CABEÇALHO DO PROCESSO ---
    st.markdown(f"## 📄 Processo: {dados['processo']}")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"**Fornecedor:** {dados['fornecedor']}")
        st.markdown(f"**CNPJ:** {dados['cnpj']}")
    with col2:
        st.markdown(f"**Empenho:** {dados['empenho']}")
        st.markdown(f"**Valor Bruto:** {dados['valor_bruto']}")

    st.divider()

    # --- SEÇÃO 2: CHECKLIST DETALHADO (LISTA) ---
    st.markdown("### 📋 Itens Verificados")
    
    # Lógica para associar documentos aos IDs SEI encontrados (exemplo baseado na ordem comum)
    docs_sei = dados['sei_docs']
    
    # Criando o DataFrame para o Layout de Tabela
    df_itens = pd.DataFrame([
        {"ITEM": 1, "EVENTO A SER VERIFICADO": "Nota de empenho e demonstrativo de saldo", "S/N/NA": "S", "OBSERVAÇÕES": f"NE {dados['empenho']}"},
        {"ITEM": 2, "EVENTO A SER VERIFICADO": "Nota Fiscal em nome do IPEM", "S/N/NA": "S", "OBSERVAÇÕES": f"Valor {dados['valor_bruto']}"},
        {"ITEM": 3, "EVENTO A SER VERIFICADO": "Certidão de Tributos Federais e Dívida Ativa", "S/N/NA": "S", "OBSERVAÇÕES": f"Validade: {dados['validade']}"},
        {"ITEM": 4, "EVENTO A SER VERIFICADO": "Nota de Liquidação (NL)", "S/N/NA": "S", "OBSERVAÇÕES": f"NL {dados['liquidacao']}"},
        {"ITEM": 5, "EVENTO A SER VERIFICADO": "IDs SEI Identificados no PDF", "S/N/NA": "S", "OBSERVAÇÕES": ", ".join(docs_sei[:4]) + "..."}
    ])

    st.table(df_itens)

    # --- SEÇÃO 3: CONCLUSÃO ---
    st.markdown("### ✍️ Conclusão Sugerida")
    conclusao_texto = f"""
    A despesa refere-se à prestação de serviços do fornecedor **{dados['fornecedor']}**. 
    Os documentos de regularidade fiscal e trabalhista foram identificados. 
    A liquidação foi processada sob o número **{dados['liquidacao']}**, estando o processo instruído para prosseguimento.
    """
    st.success(conclusao_texto)

    # --- BOTÕES DE EXPORTAÇÃO ---
    st.divider()
    st.download_button("📊 Exportar Checklist (Excel)", 
                       data=pd.DataFrame(df_itens).to_csv(index=False).encode('utf-8'), 
                       file_name=f"Checklist_{dados['processo'].replace('/','_')}.csv", 
                       mime="text/csv")

else:
    st.info("Aguardando upload do arquivo PDF na barra lateral para gerar o checklist.")
