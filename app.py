import streamlit as st
import pandas as pd
import PyPDF2
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

st.set_page_config(page_title="AuditAI - IPEM/RJ", layout="wide", page_icon="🛡️")

def extrair_dados_pdf(file):
    reader = PyPDF2.PdfReader(file)
    # Armazenamos o texto de cada página individualmente para análise de contexto
    paginas_texto = []
    for page in reader.pages:
        content = page.extract_text()
        if content:
            paginas_texto.append(content)
    
    texto_completo = " ".join(paginas_texto)
    texto_limpo = " ".join(texto_completo.split())

    # --- FUNÇÃO DE BUSCA POR PÁGINA ALVO ---
    def buscar_sei_especifico(termos_chave):
        """
        Percorre página por página. Se encontrar o termo (ex: 'Nota Fiscal'),
        extrai o código verificador que está NAQUELA página.
        """
        for texto_pagina in paginas_texto:
            # Verifica se a página contém as palavras-chave do documento buscado
            if any(termo.lower() in texto_pagina.lower() for termo in termos_chave):
                # Busca o padrão 'verificador XXXXXXXX' nesta página específica
                match = re.search(r"verificador\s+(\d{8,10})", texto_pagina, re.IGNORECASE)
                if match:
                    return match.group(1)
        return "Verificar no SEI"

    # --- EXTRAÇÃO DO ITEM 1 (FINANCEIRO) ---
    id_nl = "Não encontrada"
    id_ne = "Não encontrada"
    match_nl = re.search(r"202\dNL\d{5}", texto_limpo)
    if match_nl: id_nl = match_nl.group(0)
    match_ne = re.search(r"202\dNE\d{5}", texto_limpo)
    if match_ne: id_ne = match_ne.group(0)

    # --- MAPEAMENTO DOS ITENS DO CHECKLIST ---
    return {
        "processo": re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo).group(0) if re.search(r"SEI-\d{6}/\d{6}/\d{4}", texto_completo) else "Não encontrado",
        "empenho": id_ne,
        "liquidacao": id_nl,
        # Busca específica para cada item do checklist
        "sei_item_2": buscar_sei_especifico(["Nota Fiscal", "DANFE", "NFS-e", "Fatura"]),
        "sei_item_3": buscar_sei_especifico(["Certidão Negativa", "Créditos Tributários Federais", "Receita Federal"]),
        "sei_item_4": buscar_sei_especifico(["FGTS", "Fundo de Garantia", "CRF"]),
        "sei_item_5": buscar_sei_especifico(["Trabalhista", "CNDT", "Débitos Trabalhistas"]),
        "sei_item_13": buscar_sei_especifico(["Atesto", "Atestamos", "fatura foi conferida"])
    }

# --- INTERFACE ---
st.markdown("<h2 style='text-align: center;'>AUDIT - IPEM/RJ</h2>", unsafe_allow_html=True)
uploaded_file = st.sidebar.file_uploader("Submeter PDF do Processo SEI", type="pdf")

if uploaded_file:
    d = extrair_dados_pdf(uploaded_file)
    
    # Formatação do Item 1 (conforme sua última solicitação: sem data)
    obs_financeiro = f"{d['empenho']} - Gerando a {d['liquidacao']}"
    
    checklist = [
        {"ITEM": 1, "EVENTO": "Nota de empenho e demonstrativo de saldo", "S/N/NA": "S", "OBSERVAÇÕES": obs_financeiro},
        {"ITEM": 2, "EVENTO": "Nota Fiscal / Fatura", "S/N/NA": "S", "OBSERVAÇÕES": f"Documento SEI {d['sei_item_2']}"},
        {"ITEM": 3, "EVENTO": "Certidão Federal e Dívida Ativa", "S/N/NA": "S", "OBSERVAÇÕES": f"Documento SEI {d['sei_item_3']}"},
        {"ITEM": 4, "EVENTO": "Certidão de FGTS", "S/N/NA": "S", "OBSERVAÇÕES": f"Documento SEI {d['sei_item_4']}"},
        {"ITEM": 5, "EVENTO": "Certidão de Justiça do Trabalho", "S/N/NA": "S", "OBSERVAÇÕES": f"Documento SEI {d['sei_item_5']}"},
        {"ITEM": 13, "EVENTO": "Atestado do Gestor", "S/N/NA": "S", "OBSERVAÇÕES": f"Documento SEI {d['sei_item_13']}"},
    ]
    
    df = pd.DataFrame(checklist)
    st.table(df)
    
    # Download do arquivo Excel
    st.download_button("📥 Gerar Checklist Oficial", io.BytesIO().getvalue(), f"Audit_{d['processo'].replace('/','_')}.xlsx")
