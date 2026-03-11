import streamlit as st
from generar_notas import generar_notas

st.set_page_config(page_title="Notas Estados Financieros")

st.title("Generador de Notas a Estados Financieros")

st.write("Suba su archivo Excel con el balance.")

archivo = st.file_uploader("Subir archivo Excel", type=["xlsx"])

if archivo is not None:

    with open("balance_temp.xlsx", "wb") as f:
        f.write(archivo.read())

    st.success("Archivo cargado correctamente")

    if st.button("Generar Notas"):

        archivo_word = generar_notas("balance_temp.xlsx")

        with open(archivo_word, "rb") as f:

            st.download_button(
                label="Descargar archivo Word",
                data=f,
                file_name="notas_estados_financieros.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
