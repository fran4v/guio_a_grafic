import streamlit as st
from io import StringIO
from txt_conversion import convert
from striprtf.striprtf import rtf_to_text
import unidecode

def run_app():
    container = st.container()
    container.write("""
    # Crea un gràfic a partir d'un guió
    """)

    uploaded_file = container.file_uploader("Selecciona l'arxiu .rtf", ['rtf'])
    project_name = container.text_input('Nom del projecte')

    result = container.empty()

    result.button("Crea", disabled=True, key='1')
    if uploaded_file is not None:
        result.button("Crea", on_click=convert_file_to_graph, args=(uploaded_file, container, project_name), key='2')


def convert_file_to_graph(file, container, project_name):   
    if file is not None:
        # To read file as bytes:
        bytes_data = file.getvalue()

        #To convert to a string based IO:
        stringio = StringIO(bytes_data.decode("utf-8"))

        #To read file as string:
        string_data = stringio.read()

        if file.type == 'text/rtf':
            string_data = rtf_to_text(string_data, errors="ignore")
        
        convert(string_data, project_name)

        with open('grafic_output.xlsx', 'rb') as output_file:
            container.download_button(
                label='Descarrega',
                data=output_file,
                file_name=f'grafic {unidecode.unidecode(project_name)}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            