# Export Helvetius student file to Moodle
# V. 2.0 (2023, June)
# Maxime Cruzel

import requests
import time
import json
import urllib.parse
import olefile
import sendgrid
import os
from sendgrid.helpers.mail import *
import glob
import pandas as pd # XLS to CSV conversion
import xlrd

class config:
    login_ws = ''
    password_ws = ''
    url = 'https://my-moodle.com/'
    sendgrid_api_key = ''
    sendgrid_sender_email = 'mail@mail.com'
    # Edit the mail content in the mail_function() function
    # Edit idnumber of new accounts in the create_user() function
    
class FilesToRead:
    files_to_read = []
    

def request_ws(ws, param_request):
    login = config.login_ws
    password = config.password_ws
    password = urllib.parse.quote(password)
    url_moodle = config.url
    time.sleep(0.005)
    request_url = url_moodle+"webservice/rest/simpleserver.php?wsusername="+str(login)+"&wspassword="+str(password)+"&moodlewsrestformat=json&wsfunction="+ws+"&"
    webservice_reponse_content = requests.get(request_url, params=param_request)
    webservice_reponse_content_formated = json.loads(webservice_reponse_content.text)
    return webservice_reponse_content_formated

def mail_function(email, firstname, url):
    sg = sendgrid.SendGridAPIClient(config.sendgrid_api_key)
    from_email = Email(config.sendgrid_sender_email)
    to_email = To(email)
    subject = "Your account was created"
    txt_mail = f'''
                    Bonjour {firstname}, \n \n
                    Votre compte Moodle a été créé. \n \n
                    Vous pouvez vous connecter à l'adresse suivante : {url} \n \n
                    Pour vous connecter, cliquez sur le bouton "Microsoft" et utilisez vos identifiants Office365
                    qui vous ont été fournis par SMS. \n \n
                    Cordialement, \n
                    Le support Moodle
                    
                '''
    content = Content("text/plain", txt_mail)
    mail = Mail(from_email, to_email, subject, content)
    response = sg.client.mail.send.post(request_body=mail.get())
    print(response.status_code)
    print(response.body)
    print(response.headers)
    
def check_account(student_email):
    ws = 'core_user_get_users'
    param_request = {'criteria[0][key]':'email', 'criteria[0][value]':student_email}
    response = request_ws(ws, param_request)
    if len(response['users']) == 0:
        return False # account doesn't exist
    else:
        return True

# Take the first cohort of Helvetius file for indications in menu
def xls_reading_to_dict(file):
    with open(file, 'rb') as file_read:
        ole = olefile.OleFileIO(file_read)
        #ole = OleFileIO_PL.OleFileIO(file) # "uncorrupting" XLS file
        if ole.exists('Workbook'): # XLS file reading
            d = ole.openstream('Workbook')
            x=pd.read_excel(d, engine='xlrd')
        
        content_in_dict = x.to_dict('records')
        return content_in_dict


def get_xls_headers(file_content):
    produit = None
    formation = None
    site = None
    for row in file_content:
        produit = row['PRODUIT_LIBELLE']
        formation = row['DOMAINE']
        site = row['LIEU_DE_RATTACHEMENT']
        break
    return produit, formation, site

def read_students_in_xls(file_content):
    list_of_students = []
    for row in file_content:
        firstname = row['PRENOM']
        lastname = row['NOM']
        email = row['EMAIL_ECOLE_STD']
        student_id = row['NUMERO_ETU_ECOLE']
        list_of_students.append([firstname, lastname, email, student_id])
    return list_of_students

def create_user(firstname, lastname, email, student_id):
    ws = 'core_user_create_users'
    student_id_year = str(student_id)[1:4] # 1-2-3 first digits of student_id for the year
    student_id_number = str(student_id)[-4:] # 4 last digits of student_id for the studentnumber
    student_idnumber = student_id_year + student_id_number
    param_request = {
        'users[0][username]':email, 
        'users[0][firstname]':firstname, 
        'users[0][lastname]':lastname, 
        'users[0][email]':email, 
        'users[0][idnumber]':student_idnumber,
        'users[0][auth]':'oauth2'}
    response = request_ws(ws, param_request)
    print(response)
    if 'debuginfo' in response:
        if response['debuginfo'] == f'Username already exists: {email}':
            print('User already exists')
            return False
        else:
            print('Fatal error during account creation')
            return False
    return response[0]['id']

def enrol_user_in_cohort(userid, produit, formation, site):
    cohortname = None
    cohort_to_merge = None
    with open('config.json') as json_data:
        json_data = json_data.read()
        json_data = json.loads(json_data)
        for cohort_in_json in json_data:
            if produit == cohort_in_json['helvetius_produit'] and site == cohort_in_json['helvetius_site'] and formation == cohort_in_json['helvetius_formation']:
                cohortname = cohort_in_json['moodle_cohort_name']
                cohort_to_merge = cohort_in_json['moodle_cohort_merge_to']
                print(f'cohortname on json->{cohortname} for site = {site} and formation = {formation} and produit = {produit}')
         
                break
    ws = 'core_cohort_search_cohorts'
    param_request = {
        'query':cohortname,
        'context[contextid]':1,
        'context[contextlevel]': '',
        'context[instanceid]':0,
        # Obtained WITH SQL request on Moodle database
        # SELECT * FROM prefix_cohort ch 
        # INNER JOIN prefix_context ct ON ct.id = ch.contextid
        # 'context[contextlevel]': 10, sur beta 4.2
        }
    response = request_ws(ws, param_request)
    cohortid = response['cohorts'][0]['id']
    cohortname = response['cohorts'][0]['name']
    print(f'FOUND cohortname on Moodle->{cohortname})')
   
    ws = 'core_cohort_add_cohort_members'
    param_request = {
        'members[0][cohorttype][type]': 'id',
        'members[0][cohorttype][value]': cohortid,
        'members[0][usertype][type]': 'id',
        'members[0][usertype][value]': userid
    }
    response = request_ws(ws, param_request)
    
    if cohort_to_merge != None:
        ws = 'core_cohort_search_cohorts'
        param_request = {
            'query':cohort_to_merge,
            'context[contextid]':1,
            'context[contextlevel]': '',
            'context[instanceid]':0,
            # Obtained WITH SQL request on Moodle database
            # SELECT * FROM prefix_cohort ch 
            # INNER JOIN prefix_context ct ON ct.id = ch.contextid
            # 'context[contextlevel]': 10, sur beta 4.2
            }
        response = request_ws(ws, param_request)
        cohortid = response['cohorts'][0]['id']
        
        ws = 'core_cohort_add_cohort_members'
        param_request = {
        'members[0][cohorttype][type]': 'id',
        'members[0][cohorttype][value]': cohortid,
        'members[0][usertype][type]': 'id',
        'members[0][usertype][value]': userid
            }
        response = request_ws(ws, param_request)
    
    
        
def main():
    # Configuration file reading
    with open('config.json') as json_data:
        json_data = json_data.read()
        json_data = json.loads(json_data)
        
    # Menu build
    menu_content = []
    for i in json_data:
        menu_content.append(i)

    # Menu functions
    
    menu()
    
def menu():
    list_files_array = []
    list_files = glob.glob(r"*.xls") # xls file listing
    num_file = 1 # reset in each menu call
    print('---- MENU liste fichiers -----')
    for i_files in list_files:
        xls_data = xls_reading_to_dict(i_files)
        xls_header_data = get_xls_headers(xls_data)
        produit = xls_header_data[0]
        formation = xls_header_data[1]
        site = xls_header_data[2]
        print(f'{num_file} {produit} [{formation}] à {site} -> {i_files}')
        list_files_array.append([num_file, i_files]) # matching between 
        num_file += 1
    print('------------------------------') 
    csv_file = input('Choisir le fichier à importer. Tapez 0 si STOP > ')
    if csv_file == "0":
        for file in FilesToRead.files_to_read:
            xls_content = xls_reading_to_dict(file)
            
            xls_header_data = get_xls_headers(xls_content)
            produit = xls_header_data[0]
            formation = xls_header_data[1]
            site = xls_header_data[2]
            
            list_of_students = read_students_in_xls(xls_content)
            for student in list_of_students:
                firstname = student[0]
                lastname = student[1]
                lastname = lastname.capitalize() # student lastname not in uppercase
                email = student[2]
                idnumber = student[3]
                student_exist_test = check_account(email)
                if student_exist_test == True:
                    print(f'Student {firstname} {lastname} with email {email} already exists')
                else:
                    userid = create_user(firstname, lastname, email, idnumber)
                    if userid != False: # if user doesn't exist
                        enrol_user_in_cohort(userid, produit, formation, site)
                        mail_function(email, firstname, config.url)
                
                
    else:
        xls_file = int(csv_file) - 1 # id of an array starts at 0
        extraction_file_to_read = list_files_array[int(xls_file)][1]
        FilesToRead.files_to_read.append(extraction_file_to_read)
        clear = lambda: os.system('cls')
        clear()
        print('Liste des fichiers à importer :')
        print(FilesToRead.files_to_read)
        menu()

main()
