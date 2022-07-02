import streamlit as st
from io import StringIO

def run_app():
    st.write("""
    # Crea un gràfic a partir d'un guió
    """)

    uploaded_file = st.file_uploader("Selecciona l'arxiu .txt", ['txt'])
        
    result = st.empty()
    result.button("Crea", disabled=True, key='1')
    if uploaded_file is not None:
        result.button("Crea", on_click=convert_txt_to_graph, args=(uploaded_file,), key='2')


def convert_txt_to_graph(txt_file):   
    if txt_file is not None:
        #To convert to a string based IO:
        stringio = StringIO(txt_file.getvalue().decode("utf-8"))

        #To read file as string:
        string_data = stringio.read()
        st.write(string_data)
