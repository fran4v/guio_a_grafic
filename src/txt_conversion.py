import re
import openpyxl
import datetime 

def convert(string_data):
    
    project_name = string_data.partition('\n')[0]

    takes_list = string_data.split('------------------------------')
    num_takes = len(takes_list)

    characters_in_takes_list = []

    for take in takes_list:
        characters = detect_characters_in_take(take)
        characters_in_takes_list.append(characters)
    
    template = openpyxl.load_workbook('data/template.xlsx')

    for i in range(num_takes//50 + 1):
        write_xlsx_page(i, characters_in_takes_list, project_name, template)


def detect_characters_in_take(take):
    characters = re.findall(re.escape('*')+"(.*?)"+re.escape('*'), take)
    return list(set(characters))

def write_xlsx_page(page, characters_in_takes_list, project_name, template):
    sheets = template.sheetnames
    sheet = template[sheets[page]]

    sheet['D3'] = project_name # Nom del projecte
    sheet['AW3'] = datetime.datetime.now().strftime("%d/%m/%Y") # Data d'avui en format dd/mm/yyyy

    characters = get_characters_in_page(page, characters_in_takes_list)
    characters_total_takes = get_total_num_takes_of_characters(characters, characters_in_takes_list)
    characters_partial_takes = get_partial_num_takes_of_characters(characters, characters_in_takes_list, page)

    write_characters(sheet, characters)
    write_total_takes(sheet, characters, characters_total_takes)
    write_partial_takes(sheet, characters, characters_partial_takes)

    template.save('grafic_output.xlsx')

def get_characters_in_page(page, characters_in_takes_list):
    characters_in_page = []
    for i in range(50*page + 1, min(50*page + 50 + 1, len(characters_in_takes_list))):
        for character in characters_in_takes_list[i]:
            characters_in_page.append(character)
    return list(set(characters_in_page))

def get_total_num_takes_of_characters(characters, characters_in_takes_list):
    character_total_dict = {}
    for character in characters:
        for take in characters_in_takes_list:
            try:
                character_total_dict[character] += take.count(character)
            except:
                character_total_dict[character] = 0
    return character_total_dict

def get_partial_num_takes_of_characters(characters, characters_in_takes_list, page):
    character_partial_dict = {}
    for character in characters:
        for take in characters_in_takes_list[50*page + 1 : min(50*page + 50 + 1, len(characters_in_takes_list))]:
            try:
                character_partial_dict[character] += take.count(character)
            except:
                character_partial_dict[character] = 0
    return character_partial_dict

def write_characters(sheet, characters):
    for index, character in enumerate(characters):
        cell = f'D{8+index}'
        sheet[cell] = character

def write_total_takes(sheet, characters, characters_total_takes):
    for index, character in enumerate(characters):
        cell = f'A{8+index}'
        sheet[cell] = characters_total_takes[character]

def write_partial_takes(sheet, characters, characters_partial_takes):
    for index, character in enumerate(characters):
        cell = f'B{8+index}'
        sheet[cell] = characters_partial_takes[character]