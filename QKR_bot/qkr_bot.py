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


    for view_button_pos in pyautogui.locateAllOnScreen('Assets/view_list.png'): # Create an array of values for the Y location of the QKR view list buttons so i can click them later with pyautogui
        view_button_array.append(view_button_pos[1])

    
    i = 0
    for x in view_button_array:
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
        Functions.rename_download()