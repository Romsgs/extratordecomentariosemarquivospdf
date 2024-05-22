import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
import re

st.set_page_config(page_title="Extração de Comentários")

def extract_revisoes_table(text):
    revisions_start = re.search(r'Revisões:', text)
    if not revisions_start:
        return pd.DataFrame()

    table_text = text[revisions_start.end():]
    row_pattern = re.compile(r'(\w+)\s+(\w+)\s+(.*?)\s+(\w+)\s+(\w+)\s+(\w+)\s+(\w+)\s+(\d{2}/\d{2}/\d{4})')
    rows = row_pattern.findall(table_text)
    columns = ['Rev.', 'T.E', 'Descrição', 'Elab.', 'Ver.', 'Apr.', 'Aut.', 'Data']
    table_data = pd.DataFrame([dict(zip(columns, row)) for row in rows])
    return table_data

def extract_comments_from_pdf(pdf_file):
    comments = []
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    paginaCabecalho = doc[0]
    text = paginaCabecalho.get_text()

    numero_documento_blossom = re.search(r'\bBL\d{4}-\d{4}-\w{2}-\w{3}-\d{4}\b', text)
    numero_documento_cliente = re.search(r'\b\w{3}-\w{2}-\d{4}-\w-\d{5}-\d{4}\b', text)

    numero_documento_cliente = numero_documento_cliente.group() if numero_documento_cliente else "Não encontrado"
    numero_documento_blossom = numero_documento_blossom.group() if numero_documento_blossom else "Não encontrado"

    tabela_revisoes = extract_revisoes_table(text)
    df_revisoes = tabela_revisoes if not tabela_revisoes.empty else pd.DataFrame(columns=["Rev.", "Tipo", "Descrição", "Elab.", "Ver.", "Apr.", "Aut.", "Data"])

    for page_num in range(len(doc)):
        page = doc[page_num]
        annotations = page.annots()
        if annotations:
            for annot in annotations:
                if annot.info["content"]:
                    comments.append(annot.info["content"])

    return numero_documento_cliente, numero_documento_blossom, comments, df_revisoes

def organize_comments(documento_cliente, documento_blossom, comments):
    data = {
        "Número do Documento do Cliente": [documento_cliente] * len(comments),
        "Número do Documento da Blossom": [documento_blossom] * len(comments),
        "Comentário": comments
    }
    return pd.DataFrame(data)

def save_to_excel(tables, filename):
    with BytesIO() as buffer:
        with pd.ExcelWriter(buffer) as writer:
            for i, table in enumerate(tables):
                table.to_excel(writer, index=False, sheet_name=f"Tabela {i+1}")
            writer.save()
        st.download_button(label="Download Excel", data=buffer.getvalue(), file_name=f"{filename}.xlsx")

uploaded_files = st.file_uploader("Faça o upload de arquivos", type=['pdf', 'dwg'], accept_multiple_files=True)

if uploaded_files:
    all_comments = []
    all_revisoes = []

    for uploaded_file in uploaded_files:
        numero_documento_cliente, numero_documento_blossom, comments, df_revisoes = extract_comments_from_pdf(uploaded_file)
        
        if comments:
            table = organize_comments(numero_documento_cliente, numero_documento_blossom, comments)
            table["Revisão do Documento do Cliente"] = df_revisoes['Rev.'].iloc[-1] if not df_revisoes.empty else "Não encontrado"
            all_comments.append(table)
            all_revisoes.append(df_revisoes)

    if all_comments:
        combined_comments = pd.concat(all_comments, ignore_index=True)
        combined_revisoes = pd.concat(all_revisoes, ignore_index=True)
        
        st.write("Comentários extraídos:")
        st.dataframe(combined_comments)

        st.write("Tabela de Revisões:")
        st.dataframe(combined_revisoes)

        if st.button("Exportar para Excel"):
            filename = st.text_input("Digite o nome do arquivo:", value="Comentarios")
            if filename:
                save_to_excel([combined_comments, combined_revisoes], filename)
            else:
                st.warning("Por favor, forneça um nome para o arquivo.")
    else:
        st.warning("Nenhum comentário encontrado nos arquivos. Verifique a exatidão dos arquivos.")

if st.button("Resetar"):
    st.experimental_rerun()


# import streamlit as st
# import pandas as pd
# import fitz  # PyMuPDF
# from io import BytesIO
# import re

# st.set_page_config(page_title="Extração de Comentários")

# def extract_revisoes_table(text):
#     # Find the start of the revisions table
#     revisions_start = re.search(r'Revisões:', text)
#     if not revisions_start:
#         return [], []

#     # Extract the table content after 'Revisões:'
#     table_text = text[revisions_start.end():]

#     # Define the pattern to match rows of the table
#     row_pattern = re.compile(r'(\w+)\s+(\w+)\s+(.*?)\s+(\w+)\s+(\w+)\s+(\w+)\s+(\w+)\s+(\d{2}/\d{2}/\d{4})')
    
#     # Find all rows that match the pattern
#     rows = row_pattern.findall(table_text)

#     # Extract the column names from the first row
#     columns = ['Rev.', 'T.E', 'Descrição', 'Elab.', 'Ver.', 'Apr.', 'Aut.', 'Data']

#     # Convert the rows to a list of dictionaries
#     table_data = pd.DataFrame([dict(zip(columns, row)) for row in rows])
#     return table_data

# def extract_comments_from_pdf(pdf_file):
#     comments = []
#     doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
#     paginaCabecalho = doc[0]
#     text = paginaCabecalho.get_text()

#     # Extraindo informações do cabeçalho
#     numero_documento_blossom = re.search(r'\bBL\d{4}-\d{4}-\w{2}-\w{3}-\d{4}\b', text)
#     numero_documento_cliente = re.search(r'\b\w{3}-\w{2}-\d{4}-\w-\d{5}-\d{4}\b', text)

#     numero_documento_cliente = numero_documento_cliente.group() if numero_documento_cliente else "Não encontrado"
#     numero_documento_blossom = numero_documento_blossom.group() if numero_documento_blossom else "Não encontrado"

#     # Extraindo a tabela de revisões
#     tabela_revisoes = extract_revisoes_table(text)
#     df_revisoes = pd.DataFrame(tabela_revisoes, columns=["Rev.", "Tipo", "Descrição", "Elab.", "Ver.", "Apr.", "Aut.", "Data"])

#     for page_num in range(len(doc)):
#         page = doc[page_num]
#         annotations = page.annots()
#         if annotations:
#             for annot in annotations:
#                 if annot.info["content"]:
#                     comments.append(annot.info["content"])

#     return numero_documento_cliente, numero_documento_blossom, comments, df_revisoes

# def organize_comments(documento_cliente, documento_blossom, comments):
#     data = {
#         "Número do Documento do Cliente": [documento_cliente] * len(comments),
#         "Número do Documento da Blossom": [documento_blossom] * len(comments),
#         "Comentário": comments
#     }
#     return pd.DataFrame(data)

# def save_to_excel(tables, filename):
#     with BytesIO() as buffer:
#         with pd.ExcelWriter(buffer) as writer:
#             for i, table in enumerate(tables):
#                 table.to_excel(writer, index=False, sheet_name=f"Tabela {i+1}")
#             writer.save()
#         st.download_button(label="Download Excel", data=buffer.getvalue(), file_name=f"{filename}.xlsx")

# uploaded_files = st.file_uploader("Faça o upload de um arquivo", type=['pdf', 'dwg'], accept_multiple_files=True)

# if uploaded_files:
#     if len(uploaded_files) > 1:
#         st.warning("Por favor, faça o upload de apenas um arquivo por vez.")
#     else:
#         uploaded_file = uploaded_files[0]
#         numero_documento_cliente, numero_documento_blossom, comments, df_revisoes = extract_comments_from_pdf(uploaded_file)
        
#         if comments:
#             table = organize_comments(numero_documento_cliente, numero_documento_blossom, comments)
#             table["Revisão do Documento do Cliente"] = df_revisoes['Rev.'].iloc[-1]
#             st.write("Comentários extraídos:")
#             st.dataframe(table)

#             st.write("Tabela de Revisões:")
#             st.dataframe(df_revisoes)

#             if st.button("Exportar para Excel"):
#                 filename = st.text_input("Digite o nome do arquivo:", value="Comentarios")
#                 if filename:
#                     save_to_excel([table, df_revisoes], filename)
#                 else:
#                     st.warning("Por favor, forneça um nome para o arquivo.")
#         else:
#             st.warning("Nenhum comentário encontrado no arquivo. Verifique a exatidão do arquivo.")

# if st.button("Resetar"):
#     st.experimental_rerun()

