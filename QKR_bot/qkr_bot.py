# Selenium inports
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

#Normal Imports
import time
from QKR_bot import secret
import os
import pyautogui
import shutil
import pandas as pd
from xls2xlsx import XLS2XLSX

# --------------------------SETUP-------------------------- #
dir_path = os.path.dirname(os.path.realpath(__file__))
downloads = dir_path + '\Downloads\Waiting' # Create downloads path for files to be downloaded using selenium

view_button_array = []

class Browser: # Selenium Browser Configuration
    browser, service = None, None

    # Initialise the webdriver with the path to chromedriver.exe
    def __init__(self, driver: str):
        self.service = Service(driver)
        
        options = Options()
        options.add_argument("start-maximized")
        options.add_experimental_option("prefs", {"download.default_directory": downloads})
        
        
        self.browser = webdriver.Chrome(service=self.service, chrome_options=options)
    
    def open_page(self, url: str):
        self.browser.get(url)

    def close_browser(self):
        self.browser.close()
        
    def add_input(self, by: By, value: str, text: str):
        field = self.browser.find_element(by=by, value=value)
        field.send_keys(text)
        time.sleep(1)
        
    def click_button(self, by: By, value: str):
        button = self.browser.find_element(by=by, value=value)
        button.click()
        time.sleep(1)
        
    def login_qkr(self, username: str, password: str):
        self.add_input(by=By.NAME, value='username', text=username)
        self.add_input(by=By.NAME, value='password', text=password)
        self.click_button(by=By.CLASS_NAME, value='btn-success')

# Main Function
class Functions:
    def rename_download():
        files = os.listdir('Downloads\Waiting')
        paths = [os.path.join('Downloads\Waiting', basename) for basename in files]
        name = max(paths, key=os.path.getctime)
        
        name_new = name[17:name.find('-')] # Trim string
        name_new = 'Downloads' + name_new + '.xls'
        
        
        os.replace(name, name_new)
        return name_new 
    
    def process_download():
        print('insert code here')



def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    try:
        shutil.rmtree('Downloads\Waiting')
    except:
        try:
            os.mkdir('Downloads\Waiting')
        except:
            os.mkdir('Downloads')
            os.mkdir('Downloads\Waiting')
            
            
    browser = Browser('drivers/chromedriver')

    browser.open_page('https://qkr-mss.qkrschool.com/qkr_mss/index.html')
    time.sleep(2)
    browser.login_qkr(secret.username, secret.password)
    browser.open_page('https://qkr-mss.qkrschool.com/qkr_mss/app/storeFront#/forms')
    time.sleep(2)
    pyautogui.scroll(-330)
    time.sleep(2)

    # Append coordinate to array for every location of a QKR view list button to the view_button_array array
    for view_button_pos in pyautogui.locateAllOnScreen('Assets/view_list.png'): # Create an array of values for the Y location of the QKR view list buttons so i can click them later with pyautogui
        view_button_array.append(view_button_pos[1])

    
    i = 0
    for x in view_button_array: # All code needs to go in this loop
        if i != 0:
            pyautogui.scroll(-330)
        time.sleep(.5)
        pyautogui.moveTo(1266, view_button_array[i])
        pyautogui.click()
        time.sleep(3)
        browser.click_button(by=By.CLASS_NAME, value='btn-primary')
        time.sleep(2)
        pyautogui.moveTo(1901, 1048)
        pyautogui.click()
        browser.open_page('https://qkr-mss.qkrschool.com/qkr_mss/app/storeFront#/forms')
        pyautogui.hotkey('ctrl', 'home')
        time.sleep(3)
        i = i + 1
                
        
        files = os.listdir('Downloads\Waiting')
        paths = [os.path.join('Downloads\Waiting', basename) for basename in files]
        name = max(paths, key=os.path.getctime)

        # Rename the file 
        name_trimmed = name[17:name.find('-')] # Trim string
        name_new = 'Downloads' + name_trimmed + '.xls'
        xlsxfile = 'Downloads' + name_trimmed + '.xlsx'


        # Copy/Move the file to the downloads directory, only uncomment one of these
        shutil.copyfile(name, name_new)
        #os.replace(name, name_new)

        # Convert the file to xlsx and remove the xls version
        XLS2XLSX(name_new).to_xlsx(xlsxfile)
        os.remove(name_new)


        qkr_df = pd.read_excel(xlsxfile) 

        # Split students name into first and last. Added an exception for local excursions that have a different format
        try:
            qkr_df[['First Name','Last Name']] = qkr_df['Students Full Name:'].loc[qkr_df['Students Full Name:'].str.split().str.len() == 2].str.split(expand=True) # Split students name into first and last
            qkr_df['First Name'].fillna(qkr_df['Students Full Name:'],inplace=True)
        except: # Change this later to explicitly work for local excursion forms
            qkr_df[['First Name','Last Name']] = qkr_df['Student Name:'].loc[qkr_df['Student Name:'].str.split().str.len() == 2].str.split(expand=True) # Split students name into first and last
            qkr_df['First Name'].fillna(qkr_df['Student Name:'],inplace=True)

        try:
            qkr_df = qkr_df[["First Name", "Last Name", "Parent/Carer's Full Name:" ,"Parent/Carer's business hours number:"]]
            qkr_df = qkr_df.rename(columns={"Parent/Carer's Full Name:": "Guardian's Name",
                                    "Parent/Carer's business hours number:": "Contact Number"})
        except:
            qkr_df = qkr_df[["First Name", "Last Name", "Parent/Carer's Name:" ,"Phone Number 1:", "Name:", "Relationship to student:", "Phone Number:"]]
            qkr_df = qkr_df.rename(columns={"Parent/Carer's Name:": "Guardian's Name",
                                    "Parent/Carer's business hours number:": "Contact Number",
                                    "Name:": "Emergency Contact Name",
                                    "Relationship to student:": "Relationship",
                                    "Phone Number:": "Contact Number"})



        try:

            masterfile = 'Excursions' + name_trimmed + '.xlsx'
            master_df = pd.read_excel(masterfile)

            frames = [master_df, qkr_df]
            result = pd.concat(frames)
            result = result.drop_duplicates()
            result.to_excel('Excursions' + name_trimmed + '.xlsx', index=False)  
            
            
        except:
            qkr_df.to_excel('Excursions' + name_trimmed + '.xlsx', index=False)  

        os.remove(xlsxfile) 