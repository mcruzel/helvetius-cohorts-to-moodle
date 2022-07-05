# Export Helvetius student file to Moodle
# V. 1.0 (2022, June)
# Maxime Cruzel

import json
import pandas as pd # XLS to CSV conversion
import csv
import OleFileIO_PL # XLS file "uncorrupting"
import glob # files listing

# arrays for CSV making
lastnames_from_csv = []
firstnames_from_csv = []
emails_from_csv = []
csv_final = []


# Configuration file reading
with open('config.json') as json_data:
    json_data = json_data.read()
    json_data = json.loads(json_data)

# Menu build
menu_content = []
for i in json_data:
   menu_content.append(i)

# Menu functions
list_files_array = []
def menu_csv():
    global csv_file
    global num_file
    global list_files
    list_files = glob.glob(r"*.xls") # xls file listing
    num_file = 1 # reset in each menu call
    print('---- MENU liste fichiers -----')
    for i_files in list_files:
       
        print(str(num_file)+" -> "+i_files)
        list_files_array.append([num_file, i_files]) # matching between 
        num_file += 1
    print('------------------------------') 
    csv_file = input('Choisir le fichier à importer. Tapez 0 si STOP.')
    if csv_file == "0":
        export_csv()
    else:
        csv_file = int(csv_file) - 1 # id of an array starts at 0
        csv_file = list_files_array[int(csv_file)][1]
        reading_csv()

# choice cohort menu (from json file)
def menu_cohort():
    global menu_cht_choice
    print(csv_file)
    print('---- MENU choix cohorte -----')
    for ma in menu_content:
        print(str(ma['id'])+" -> "+ma['name'])
    menu_cht_choice = input('Quel est votre choix : (tapez 0 si STOP)')
    print('------------------------------')
    if menu_cht_choice == "0":
        menu_csv()
    else:
        cohort_choice_matching()

# matching between cohort choice (on menu) and json file
def cohort_choice_matching():
    global cohort_choice
    with open('config.json') as json_data:
        json_data = json.load(json_data)
        for i in json_data:
            if str(i['id']) == menu_cht_choice:
                cohort_choice = i['cohort']
        print(cohort_choice)

# Helvetius xls reading
def reading_csv():
    menu_cohort() # cohort choice before
    global lastnames_from_csv
    global firstnames_from_csv
    global emails_from_csv
    global csv_final
    print('lecture en cours')
    path = csv_file
    with open(path,'rb') as file:
        ole = OleFileIO_PL.OleFileIO(file) # "uncorrupting" XLS file
        if ole.exists('Workbook'): # XLS file reading
            d = ole.openstream('Workbook')
            x=pd.read_excel(d,engine='xlrd')   
    x.to_csv("a_supprimer.csv", index = None, header = True, sep=";") # csv "working file" generating
    df = pd.DataFrame(pd.read_csv("a_supprimer.csv", sep=";"))
    
    with open('a_supprimer.csv', newline='') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=';', quotechar='|')
        for row in csvreader:
            if row[2] == "NOM": # header skipping
                print()
            else:
                lastname = row[2]
                lastname = lastname.capitalize()
                lastnames_from_csv.append((row[2]).capitalize()) # capitalize : first letter on caps
                firstnames_from_csv.append(row[3])
                emails_from_csv.append(row[11])
                csv_final.append([row[11], lastname, row[3], row[11], cohort_choice, 'oauth2'])
    print(csv_final)
    menu_csv()

# CSV making function
def export_csv():
    print('export en cours')
    header_csv_moodle = ['username', 'lastname', 'firstname', 'email', 'cohort1', 'auth']
    with open('moodle_csv.csv', 'w', encoding='utf-8', newline='') as f_csv_moodle:
        writer = csv.writer(f_csv_moodle, delimiter=';')
        writer.writerow(header_csv_moodle)
        writer.writerows(csv_final)
    f_csv_moodle.close()
    print('Travail terminé')
    
menu_csv()
