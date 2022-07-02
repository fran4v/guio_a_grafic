import re

def convert(string_data):
    
    takes_list = string_data.split('------------------------------')
    num_takes = len(takes_list)

    actors_in_takes_list = []

    for take in takes_list:
        actors = detect_actors_in_take(take)
        actors_in_takes_list.append(actors)
    
    for i in range(num_takes//50):
        write_xlsx_page(i, actors_in_takes_list)
    

    

def detect_actors_in_take(take):
    actors = re.findall(re.escape('*')+"(.*?)"+re.escape('*'), take)
    return list(set(actors))

def write_xlsx_page(page, actors_in_takes_list):
    