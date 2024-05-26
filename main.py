import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)
client = gspread.authorize(creds)

sheet = client.open('MudaSG Project Sign-Ups [TRAINED VOLUNTEERS] (Responses)')

def project_Grabber():
    #Grabs Sheet Names and creates an Array with all names.
    Projects = []

    sheet_instance = sheet.worksheets()

    for x in sheet_instance:
        try:
            team = (str(x).split("<Worksheet '")[1].split("'")[0])
            Projects.append(team)
        except Exception as e:
            print(x)
            print(e)

    print(Projects)

projects = ['Bazaar Bestari', "IM'VN SG collab", 'FFTH x MudaSG', 'Willing Hearts collab', 'Lion Befrienders collab']
memberList = []
tempNumbers = []

for project in projects:
    sheet_instance = sheet.worksheet(project)
    members = sheet_instance.get_all_values()[1:]
    for member in members:
        #formatted_member = member[0] + ':' + member[1]
        if member[1] in tempNumbers:
            pass
        elif member[1] == '88008330':
            pass
        else:
            tempNumbers.append(member[1])
            memberList.append([member[0], member[1]])


print(memberList)

driver = webdriver.Chrome()
driver.get('https://web.whatsapp.com')
input('Enter when logged in')
for projectName in projects:
    try:
        driver.find_element_by_xpath('//span[@data-icon="menu"]').click()
    except Exception as e:
        input('Error Clicking Menu icon')
    try:
        driver.find_element_by_xpath('//div[@aria-label="New group"]').click()
    except Exception as e:
        input('Error Clicking New group')
    print(f'{projectName} - Starting creation...')
    textBox = driver.find_element_by_xpath('//input[@placeholder="Type contact name"]')
    sheet_instance = sheet.worksheet(projectName)
    members = sheet_instance.get_all_values()[1:]
    for member in members:
        try:
            textBox.send_keys(member[0])
            driver.find_element_by_xpath(f'//span[@title="{member[0]}"]').click()
            print(f'{projectName} - Added {member[0]}')
        except Exception as e:
            input(f'Error - Adding {member[0]}')

    time.sleep(1)
    try:
        driver.find_element_by_xpath('//span[@data-icon="arrow-forward"]').click()
    except Exception as e:
        input('Error Clicking Arrow Forward')
        print(e)
    try:
        driver.find_element_by_xpath('//*[@id="app"]/div[1]/div[1]/div[2]/div[1]/span/div[1]/span/div[1]/div/div/div[2]/div/div[2]/div/div/div[2]').send_keys(projectName)
    except Exception as e:
        input('Error Typing Group Name')
    time.sleep(1)
    try:
        driver.find_element_by_xpath('//span[@data-icon="checkmark-medium"]').click()
    except Exception as e:
        input('Error Clicking create button')
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f'//span[@title="{projectName}"]'))
        )
        print(f'{projectName} - Group Created Successfully! ')
    except:
        input(f'{projectName} - Error Creating Group!')

def create_contact_list_sheet(memberList, sheet):
    sheet = sheet.worksheet('All Names & Numbers')
    row = [[member[0], member[1]] for member in memberList]
    print(row)
    sheet.insert_rows(row)

