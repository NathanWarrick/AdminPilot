# Selenium inports
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

import os
import pyautogui
import pandas as pd
from xls2xlsx import XLS2XLSX
import time
import cv2

from workspace.bots import secret


# --------------------------SETUP-------------------------- #
dir_path = os.path.dirname(os.path.realpath(__file__))
downloads = dir_path + r"\downloads" # Create downloads path for files to be downloaded using selenium
driverlocation = r'workspace\assets\drivers\chromedriver.exe'


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
        
class Browser_Headless: # Selenium Browser Configuration
    browser, service = None, None

    # Initialise the webdriver with the path to chromedriver.exe
    def __init__(self, driver: str):
        self.service = Service(driver)
        
        options = Options()
        #options.add_argument("start-maximized")
        options.headless = True
        options.add_argument('window-size=1920x1080');
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
    
    def page_height(self):
        S = lambda X: self.browser.execute_script('return document.body.parentNode.scroll'+X)
        self.browser.set_window_size(S('Width'),S('Height')) # May need manual adjustment
        self.browser.find_element(by=By.TAG_NAME, value='body').screenshot('web_screenshot.png')
        im = cv2.imread('web_screenshot.png')
        h, w, _ = im.shape
        os.remove('web_screenshot.png')       
        return h


def main():   
    
    browser = Browser_Headless(driverlocation)
    browser.open_page('https://qkr-mss.qkrschool.com/qkr_mss/index.html')
    time.sleep(2)
    browser.login_qkr(secret.qkr_username, secret.qkr_password)
    browser.open_page('https://qkr-mss.qkrschool.com/qkr_mss/app/storeFront#/forms')
    time.sleep(2)
    print('Page height is: ' + str(browser.page_height()))    
       
    browser = Browser(driverlocation)
    browser.open_page('https://qkr-mss.qkrschool.com/qkr_mss/index.html')
    time.sleep(2)
    browser.login_qkr(secret.qkr_username, secret.qkr_password)
    browser.open_page('https://qkr-mss.qkrschool.com/qkr_mss/app/storeFront#/forms')
    time.sleep(2)
    pyautogui.scroll(-330)
    time.sleep(2)

    # Append coordinate to array for every location of a QKR view list button to the view_button_array array

    for view_button_pos in pyautogui.locateAllOnScreen(r"workspace\bots\qkr_bot\assets\view_list.png"): # Create an array of values for the Y location of the QKR view list buttons so i can click them later with pyautogui
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
        x, y =pyautogui.locateCenterOnScreen(r"workspace\bots\qkr_bot\assets\download_exit.png", region=(0,700,1920,1080))
        pyautogui.moveTo(x, y, .2)
        pyautogui.click()
        browser.open_page('https://qkr-mss.qkrschool.com/qkr_mss/app/storeFront#/forms')
        pyautogui.hotkey('ctrl', 'home')
        time.sleep(3)
        i = i + 1
                
        
        # Determine the newest file added to the directory
        files = os.listdir(r"workspace\bots\qkr_bot\Downloads")
        paths = [os.path.join(r"workspace\bots\qkr_bot\Downloads", basename) for basename in files]
        name = max(paths, key=os.path.getctime)


        # Rename the file 
        excursionname = name[32:name.find('-')] # Trim string to just the name of the form
        #name_new = r"workspace\bots\qkr_bot\Downloads" + excursionname + '.xls'
        xlsxfile = r"workspace\bots\qkr_bot\Downloads" + excursionname + '.xlsx'


        # Copy/Move the file to the downloads directory
        #os.replace(name, name_new)

        # Convert the file to xlsx and remove the xls version
        XLS2XLSX(name).to_xlsx(xlsxfile)
        os.remove(name)
        qkr_df = pd.read_excel(xlsxfile) 
        
        
        # Split students name into first and last. Added an exception for local excursions that have a different format
        if excursionname == '\LocalExcursionForm':
            print("Local Excursion Form Detected")
            qkr_df[['First Name','Last Name']] = qkr_df['Student Name:'].loc[qkr_df['Student Name:'].str.split().str.len() == 2].str.split(expand=True) # Split students name into first and last
            qkr_df['First Name'].fillna(qkr_df['Student Name:'],inplace=True)
            qkr_df = qkr_df[["First Name", "Last Name", "Parent/Carer's Name:" ,"Phone Number 1:", "Name:", "Relationship to student:", "Phone Number:"]]
            qkr_df = qkr_df.rename(columns={"Parent/Carer's Name:": "Guardian's Name",
                                    "Parent/Carer's business hours number:": "Contact Number",
                                    "Name:": "Emergency Contact Name",
                                    "Relationship to student:": "Relationship",
                                    "Phone Number:": "Contact Number"})

        else:
            qkr_df[['First Name','Last Name']] = qkr_df['Students Full Name:'].loc[qkr_df['Students Full Name:'].str.split().str.len() == 2].str.split(expand=True) # Split students name into first and last
            qkr_df['First Name'].fillna(qkr_df['Students Full Name:'],inplace=True)

            try:
                qkr_df = qkr_df[["First Name", "Last Name", "Parent/Carer's Full Name:" ,"Parent/Carer's business hours number:",
                                "Swimming Ability:",
                                "Tick all relevant conditions", "If other, include any other diagnosed physical or mental health conditions",
                                "Please tick if you are allergic to any of the following", "If other please list all other allergies",
                                "What special care is recommended for these allergies?", "Is your child taking any medicine(s)?",
                                "If yes, provide the name of medication, dose and describe when and how it is to be taken.",
                                "Is there anything else about your child’s health and wellbeing or medical history that is important for us to know?"]]
                qkr_df = qkr_df.rename(columns={"Parent/Carer's Full Name:": "Guardian's Name",
                                        "Parent/Carer's business hours number:": "Contact Number",
                                        "Swimming Ability:": "Swimming Ability",
                                        "Tick all relevant conditions": "Conditions", 
                                        "If other, include any other diagnosed physical or mental health conditions": "Other Conditions",
                                        "Please tick if you are allergic to any of the following": "Allergies",
                                        "If other please list all other allergies": "Other Allergies",
                                        "What special care is recommended for these allergies?": "Allergy Care",
                                        "Is your child taking any medicine(s)?": "Medicines",
                                        "If yes, provide the name of medication, dose and describe when and how it is to be taken.": "Medicine Instructions",
                                        "Is there anything else about your child’s health and wellbeing or medical history that is important for us to know?": "Additional Medical Details"})
            except:
                qkr_df = qkr_df[["First Name", "Last Name", "Parent/Carer's Full Name:" ,"Parent/Carer's business hours number:",
                                "Tick all relevant conditions", "If other, include any other diagnosed physical or mental health conditions",
                                "Please tick if you are allergic to any of the following", "If other please list all other allergies",
                                "What special care is recommended for these allergies?", "Is your child taking any medicine(s)?",
                                "If yes, provide the name of medication, dose and describe when and how it is to be taken.",
                                "Is there anything else about your child’s health and wellbeing or medical history that is important for us to know?"]]
                qkr_df = qkr_df.rename(columns={"Parent/Carer's Full Name:": "Guardian's Name",
                                        "Parent/Carer's business hours number:": "Contact Number",
                                        "Tick all relevant conditions": "Conditions", 
                                        "If other, include any other diagnosed physical or mental health conditions": "Other Conditions",
                                        "Please tick if you are allergic to any of the following": "Allergies",
                                        "If other please list all other allergies": "Other Allergies",
                                        "What special care is recommended for these allergies?": "Allergy Care",
                                        "Is your child taking any medicine(s)?": "Medicines",
                                        "If yes, provide the name of medication, dose and describe when and how it is to be taken.": "Medicine Instructions",
                                        "Is there anything else about your child’s health and wellbeing or medical history that is important for us to know?": "Additional Medical Details"})



        # Compare master file in /Excursions to downloaded and formatted file, exception for if file does not exist
        try:
            masterfile = 'Excursions' + excursionname + '.xlsx'
            master_df = pd.read_excel(masterfile)

            frames = [master_df, qkr_df]
            result = pd.concat(frames)
            result = result.drop_duplicates()
            result.to_excel('Excursions' + excursionname + '.xlsx', index=False)  
            
            
        except:
            qkr_df.to_excel("Excursions" + excursionname + '.xlsx', index=False)  

        os.remove(xlsxfile) 
        return('done')
        