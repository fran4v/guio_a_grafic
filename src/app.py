import streamlit as st
from io import StringIO
from rtf_conversion import convert_graph, convert_summary
from striprtf.striprtf import rtf_to_text
import unidecode
from zipfile import ZipFile

def run_app():
    container = st.container()
    container.write("""
    # Crea un gràfic a partir d'un guió
    """)

    uploaded_file = container.file_uploader("Selecciona l'arxiu .rtf", ['rtf'])
    project_name = container.text_input('Nom del projecte')

    result = container.empty()

    result.button("Crea", disabled=True, key='1')
    include_summary = container.checkbox('Inclou resum')

    if uploaded_file is not None:
        result.button("Crea", on_click=convert_file_to_graph, args=(uploaded_file, container, project_name, include_summary), key='2')


def convert_file_to_graph(file, container, project_name, include_summary):   
    if file is not None:
        # To read file as bytes:
        bytes_data = file.getvalue()

        #To convert to a string based IO:
        stringio = StringIO(bytes_data.decode("utf-8"))

        #To read file as string:
        string_data = stringio.read()

        if file.type == 'text/rtf':
            string_data = rtf_to_text(string_data, errors="ignore")
        
        convert_graph(string_data, project_name)

        if include_summary:
            convert_summary(string_data)

        download(project_name, container, include_summary)
        

def download(project_name, container, include_summary):
    if not include_summary:
        with open('grafic_output.xlsx', 'rb') as output_file:
            container.download_button(
                label='Descarrega',
                data=output_file,
                file_name=f'grafic {unidecode.unidecode(project_name)}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        write_zip()
        with open("output.zip", "rb") as output_file:
            container.download_button(
                label='Descarrega',
                data=output_file,
                file_name=f'{unidecode.unidecode(project_name)}.zip',
                mime='application/zip')


def write_zip():
    zipObj = ZipFile("output.zip", "w")
    zipObj.write("grafic_output.xlsx")
    zipObj.write("summary.rtf")
    zipObj.close()