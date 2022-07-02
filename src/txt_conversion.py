import re
import openpyxl
import datetime 

def convert(string_data):
    
    project_name = string_data.partition('\n')[0]

    takes_list = string_data.split('------------------------------')
    num_takes = len(takes_list)

    actors_in_takes_list = []

    for take in takes_list:
        actors = detect_actors_in_take(take)
        actors_in_takes_list.append(actors)
    
    template = openpyxl.load_workbook('data/template.xlsx')

    for i in range(num_takes//50 + 1):
        write_xlsx_page(i, actors_in_takes_list, project_name, template)


def detect_actors_in_take(take):
    actors = re.findall(re.escape('*')+"(.*?)"+re.escape('*'), take)
    return list(set(actors))

def write_xlsx_page(page, actors_in_takes_list, project_name, template):
    sheets = template.sheetnames
    sheet = template[sheets[page]]

    sheet['D3'] = project_name # Nom del projecte
    sheet['AW3'] = datetime.datetime.now().strftime("%d/%m/%Y") # Data d'avui en format dd/mm/yyyy

    characters = get_characters_of_page(page, actors_in_takes_list)
    print(characters)

    write_characters(sheet, characters)
    write_total_takes(sheet)
    write_partial_takes(sheet)

    template.save('grafic_output.xlsx')

def get_characters_of_page(page, actors_in_takes_list):
    actors_in_page = []
    for i in range(50*page, min(50*page + min(50, 50 + len(actors_in_takes_list) - 50*page), 50*page + len(actors_in_takes_list) - 50*page)):
        for actor in actors_in_takes_list[i]:
            actors_in_page.append(actor)
    return list(set(actors_in_page))

def write_characters(sheet, characters):
    for index, character in enumerate(characters):
        cell = f'D{8+index}'
        sheet[cell] = character

def write_total_takes(sheet):
    pass

def write_partial_takes(sheet):
    pass