import streamlit as st
from io import StringIO
from txt_conversion import convert
from striprtf.striprtf import rtf_to_text

def run_app():
    container = st.container()
    container.write("""
    # Crea un gràfic a partir d'un guió
    """)

    uploaded_file = container.file_uploader("Selecciona l'arxiu .rtf o .txt", ['rtf', 'txt'])
        
    result = container.empty()
    result.button("Crea", disabled=True, key='1')
    if uploaded_file is not None:
        result.button("Crea", on_click=convert_file_to_graph, args=(uploaded_file, container), key='2')


def convert_file_to_graph(file, container):   
    if file is not None:
        # To read file as bytes:
        bytes_data = file.getvalue()

        #To convert to a string based IO:
        stringio = StringIO(file.getvalue().decode("utf-8"))

        #To read file as string:
        string_data = stringio.read()

        if file.type == 'text/rtf':
            string_data = rtf_to_text(string_data)
        
        convert(string_data)

        container.progress(30)