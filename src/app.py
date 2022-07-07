import streamlit as st
from io import StringIO
from text_conversion import convert_graph, convert_summary, update_files
from striprtf.striprtf import rtf_to_text
import unidecode
from zipfile import ZipFile
import docx2txt

def run_app():
    hide_menu()

    if 'pushed_crea_button' not in st.session_state:
        st.session_state['pushed_crea_button'] = False
    if 'characters_total_takes' not in st.session_state:
        st.session_state['characters_total_takes'] = {}
    if 'updated_actors' not in st.session_state:
        st.session_state['updated_actors'] = False
    if 'project_name' not in st.session_state:
        st.session_state['project_name'] = ''
    if 'include_summary' not in st.session_state:
        st.session_state['include_summary'] = False
    if 'num_takes' not in st.session_state:
        st.session_state['num_takes'] = 0
    if 'characters_actors_dict' not in st.session_state:
            st.session_state['characters_actors_dict'] = {}

    container = st.container()
    container.write("""
    # Crea un gràfic a partir d'un guió
    """)

    uploaded_file = container.file_uploader("Selecciona l'arxiu .rtf, .txt o .docx", ['rtf', 'txt', 'docx'])
    project_name = container.text_input('Nom del projecte')
    st.session_state['project_name'] = project_name

    result = container.empty()

    result.button("Crea", disabled=True, key='1')
    
    if not st.session_state['pushed_crea_button']:
        include_summary = container.checkbox('Inclou resum')
    else: 
        include_summary = container.checkbox('Inclou resum', disabled=True)


    if include_summary:
        st.session_state['include_summary'] = True
    else:
        st.session_state['include_summary'] = False

    if uploaded_file is not None:
        result.button("Crea", on_click=convert_file_to_graph, args=(uploaded_file, project_name, include_summary), key='2')
    
    if uploaded_file is None:
        st.session_state['pushed_crea_button'] = False
    
    if st.session_state['pushed_crea_button'] and uploaded_file is not None:
        download(project_name, container, include_summary)

        container.text("")
        container.text("")
        container.write('Previsualització (els personatges assenyalats amb "⚠" només apareixen en una take)')
        
        characters_total_takes = st.session_state['characters_total_takes']
        preview_characters(characters_total_takes, container)

        container.write("*En cas que es detecti algun error, cal modificar l'arxiu que s'ha penjat i tornar-ho a intentar")

def convert_file_to_graph(file, project_name, include_summary):
    st.session_state['pushed_crea_button'] = True

    if file is not None:
        bytes_data = file.getvalue() # Read file as bytes

        stringio = StringIO(bytes_data.decode('ISO-8859-1')) # Convert bytes to a string based IO

        string_data = stringio.read() # Read file as string

        # Check if its a .rtf file
        if file.name.endswith('.rtf'):
            string_data = rtf_to_text(string_data, errors="ignore") # Convert .rtf file to clean string, ignorning special characters
        # Check if its a .docx file
        elif file.name.endswith('.docx'):
            string_data = docx2txt.process(file)  # Convert to txt

        characters_total_takes = convert_graph(string_data, project_name)

        st.session_state['characters_total_takes'] = characters_total_takes

        if include_summary:
            convert_summary(characters_total_takes)
            
def preview_characters(characters_total_takes, container):
    copy_characters_total_takes = dict(characters_total_takes)
    for i in list(copy_characters_total_takes):
        if copy_characters_total_takes[i] == 1 and '⚠' not in i:
            copy_characters_total_takes.pop(i, None)
            copy_characters_total_takes[f'{i}\t⚠'] = 1
    sorted_dict = {key: value for key, value in sorted(copy_characters_total_takes.items())}

    voice_actors = ('(buit)', 'Actor1', 'Actor2', 'Actor3', 'Actor4') # WIP: Afegir txt amb llistat d'actors de doblatge
        
    form = container.form(key='voice_actors_form', clear_on_submit=False)
    for i in range(len(characters_total_takes)):
        character = list(sorted_dict.keys())[i]
        total_takes = list(sorted_dict.values())[i]
        form.selectbox(label=f'{character} ({total_takes} takes)', options=voice_actors, key=character.replace('\t⚠', ''))
    submitted = form.form_submit_button("Afegeix actors")

    if submitted:
        project_name = st.session_state['project_name']
        include_summary = st.session_state['include_summary']
        st.session_state['updated_actors'] = True
        update_files(characters_total_takes)
        download_updated(project_name, container, include_summary)

def download(project_name, container, include_summary):
    if not include_summary:
        with open('grafic.xlsx', 'rb') as output_file:
            container.download_button(
                label='Descarrega sense actors',
                data=output_file,
                file_name=f'grafic {unidecode.unidecode(project_name)}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key = 1)
    else:
        write_zip()
        if project_name == '':
            file_name = 'grafic.zip'
        else:
            file_name = f'{unidecode.unidecode(project_name)}.zip'
        with open("output.zip", "rb") as output_file:
            container.download_button(
                label='Descarrega sense actors',
                data=output_file,
                file_name=file_name,
                mime='application/zip',
                key = 1)

def download_updated(project_name, container, include_summary):
    if not include_summary:
        with open('grafic_actors.xlsx', 'rb') as output_file:
            container.download_button(
                label='Descarrega amb actors',
                data=output_file,
                file_name=f'grafic {unidecode.unidecode(project_name)}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key = 2)
    else:
        write_zip()
        if project_name == '':
            file_name = 'grafic.zip'
        else:
            file_name = f'{unidecode.unidecode(project_name)}.zip'
        with open("output_updated.zip", "rb") as output_file:
            container.download_button(
                label='Descarrega amb actors',
                data=output_file,
                file_name=file_name,
                mime='application/zip',
                key = 2)

def write_zip():
    if st.session_state['updated_actors'] == False:
        zipObj = ZipFile("output.zip", "w")
        zipObj.write("grafic.xlsx")
        zipObj.write("resum.txt")
        zipObj.close()
    else:
        zipObj = ZipFile("output_updated.zip", "w")
        zipObj.write("grafic_actors.xlsx")
        zipObj.write("resum_actors.txt")
        zipObj.close()

def hide_menu():
    hide_menu_style = """
        <style>
        #MainMenu {visibility: hidden; }
        footer {visibility: hidden;}
        </style>
        """
    st.markdown(hide_menu_style, unsafe_allow_html=True)
