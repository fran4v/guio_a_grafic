import streamlit as st
from io import StringIO
from rtf_conversion import convert_graph, convert_summary
from striprtf.striprtf import rtf_to_text
import unidecode
from zipfile import ZipFile
import docx2txt

def run_app():
    container = st.container()
    container.write("""
    # Crea un gràfic a partir d'un guió
    """)

    uploaded_file = container.file_uploader("Selecciona l'arxiu .rtf, .txt o .docx", ['rtf', 'txt', 'docx'])
    project_name = container.text_input('Nom del projecte')

    result = container.empty()

    result.button("Crea", disabled=True, key='1')
    include_summary = container.checkbox('Inclou resum (WIP)')

    if uploaded_file is not None:
        result.button("Crea", on_click=convert_file_to_graph, args=(uploaded_file, container, project_name, include_summary), key='2')


def convert_file_to_graph(file, container, project_name, include_summary):   
    if file is not None:
        bytes_data = file.getvalue() # Read file as bytes

        stringio = StringIO(bytes_data.decode('ISO-8859-1')) # Convert bytes to a string based IO

        string_data = stringio.read() # Read file as string

        # Check if its a .rtf file
        if file.type == 'text/rtf':
            string_data = rtf_to_text(string_data, errors="ignore") # Convert .rtf file to clean string, ignorning special characters
        
        # Check if its a .docx file
        elif file.type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
            string_data = docx2txt.process(file)  # Convert to txt

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