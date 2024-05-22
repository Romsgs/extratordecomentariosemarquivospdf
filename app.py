import fitz  # PyMuPDF
import streamlit as st
import pandas as pd
from io import BytesIO
import re

# Função para extrair dados do cabeçalho e comentários de um PDF
def extract_comments_and_header_from_pdf(pdf_file):
    comments = []
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    paginaCabecalho = doc[0]
    text = paginaCabecalho.get_text()

    numero_documento_blossom  = re.search(r'\bBL\d{4}-\d{4}-\w{2}-\w{3}-\d{4}\b', text)
    revisao_documento_cliente = re.search(r'Rev\.\s*([A-Z0-9])', text)
    numero_documento_cliente  = re.search(r'\b\w{3}-\w{2}-\d{4}-\w-\d{5}-\d{4}\b', text)

    numero_documento_cliente = numero_documento_cliente.group() if numero_documento_cliente else "Não encontrado"
    revisao_documento_cliente = revisao_documento_cliente.group(1) if revisao_documento_cliente else "Não encontrado"
    numero_documento_blossom = numero_documento_blossom.group() if numero_documento_blossom else "Não encontrado"

    for page_num in range(len(doc)):
        page = doc[page_num]
        annotations = page.annots()
        if annotations:
            for annot in annotations:
                if annot.info["content"]:
                    comments.append(annot.info["content"])

    return numero_documento_cliente, revisao_documento_cliente, numero_documento_blossom, comments

# Função para organizar os comentários na estrutura necessária
def organize_comments(documento_cliente, revisao_documento, documento_blossom, comments):
    data = {
        "Número do Documento do Cliente": [documento_cliente] * len(comments),
        "Revisão do Documento do Cliente": [revisao_documento] * len(comments),
        "Número do Documento da Blossom": [documento_blossom] * len(comments),
        "Comentário": comments
    }
    return pd.DataFrame(data)

# Interface em Streamlit
st.title("PDF Comment Extractor")

# Upload do PDF
uploaded_file = st.file_uploader("Escolha um arquivo PDF", type="pdf")

if uploaded_file is not None:
    # Extrair dados do cabeçalho e comentários
    numero_documento_cliente, revisao_documento_cliente, numero_documento_blossom, comments = extract_comments_and_header_from_pdf(uploaded_file)

    st.write("Número do Documento do Cliente:", numero_documento_cliente)
    st.write("Revisão do Documento do Cliente:", revisao_documento_cliente)
    st.write("Número do Documento da Blossom:", numero_documento_blossom)

    if comments:
        
        # Organizar os comentários na estrutura necessária
        comments_df = organize_comments(numero_documento_cliente, revisao_documento_cliente, numero_documento_blossom, comments)
        # Salvar os comentários em um arquivo Excel
        excel_output = BytesIO()
        
        comments_df.to_excel(excel_output, index=False, sheet_name='Comentários')
        
        excel_output.seek(0)
        
        # Disponibilizar para download
        st.download_button(
            label="Baixar comentários em Excel",
            data=excel_output,
            file_name="comments.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.write("Nenhum comentário encontrado no PDF.")
