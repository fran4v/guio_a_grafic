import re

def convert(string_data):
    
    takes_list = string_data.split('------------------------------')
    actors_in_takes_list = []

    for take in takes_list:
        actors = detect_actors_in_take(take)
        actors_in_takes_list.append(actors)
    print(actors_in_takes_list)

def detect_actors_in_take(take):
    actors = re.findall(re.escape('*')+"(.*?)"+re.escape('*'),take)
    return list(set(actors))