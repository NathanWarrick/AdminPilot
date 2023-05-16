import win32com.client
import os
import re
import PyPDF2
import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By



# ------------------------ SETUP --------------------------------
driverlocation = r'workspace\assets\drivers\chromedriver.exe'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the Inbox

qkrfolder = {'Jan': '1. January','Feb': '2. February','Mar': '3. March','Apr': '4. April','May': '5. May','Jun': '6. June','July': '7. Jul','Aug': '8. August','Sep': '9. September','Oct': '10. October','Nov': '11. November','Dec': '12. Decemmber'}

folders = {}    

class Browser: # Selenium Browser Configuration
    browser, service = None, None

    # Initialise the webdriver with the path to chromedriver.exe
    def __init__(self, driver: str):
        self.service = Service(driver)
        
        options = Options()
        options.add_argument("start-maximized")
        
        
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
        self.add_input(by=By.XPATH, value='//*[@id="content"]/div[2]/div/div[2]/form/input[2]', text=username)
        self.add_input(by=By.XPATH, value='//*[@id="content"]/div[2]/div/div[2]/form/input[3]', text=password)
        self.click_button(by=By.XPATH, value='//*[@id="signIn"]')
    

# This funcion assigns all folders indexes to a dictionary
# To retrieve a index use folders['Folder Name']
for f in range(100):
    try:
        fdr = inbox.Folders(f)
        name = fdr.Name
        folders.update({name: f})
        #print(f, name)
    except:
        pass


def excursion_check():
    browser = Browser(driverlocation)
    for msg in inbox.Folders(folders['1 Excursions']).Items:
        if re.match('Excursion Approved - .+', msg.Subject) and msg.FlagStatus == 2: # Grab flagged excursion approvals
            subject = msg.Subject
            #print('Subject: ' + subject)

            match = 0
            i = 1
            while match == 0:
                if not re.search('Medical Form', str(msg.Attachments.Item(i))) and not re.search('.png', str(msg.Attachments.Item(i))): # Check that the attachment doesn't contain the terms medical form or .png in the title
                    attachment = msg.Attachments.Item(i)
                    #print('Path: ' + os.getcwd() + r'\workspace\clients\email_client\downloads' + '\\' + attachment.FileName)
                    attachment.SaveAsFile(os.getcwd() + r'\workspace\clients\email_client\downloads' + '\\' + attachment.FileName)
                    match = 1
                    
                    # Read the PDF
                    form = PyPDF2.PdfReader(os.getcwd() + r'\workspace\clients\email_client\downloads' + '\\' + attachment.FileName)
                    data = form.pages[0].extract_text()
                    #print(data)
                    
                    # ************************************************ Data Extraction ************************************************
                    
                    # Excursion Name
                    excursionnamestart = subject.find('Excursion Approved')
                    excursionnameend = subject.find('\n')
                    excursionname = subject[excursionnamestart+21:]
                    print('Name: ' + excursionname)               
                    
                    # Excursion Date
                    dateloc = data.find('Date /s')
                    datenl = data.find('\n', dateloc)
                    date = data[dateloc+8:datenl]
                    print('Date: ' + date)
                    
                    # Excursion Cost
                    costloc = data.find('Details of Cost')
                    costnl = data.find('\n', costloc)   
                    cost = data[costloc+16:costnl]     
                    print('Cost: ' + cost)
                    print('\n')
                    
                    # ************************************************ Begin QKR ************************************************
                    
                    browser.open_page('https://qkr-mss.qkrschool.com/qkr_mss/index.html')
                    time.sleep(2)
                    browser.login_qkr('nathan.warrick@education.vic.gov.au', 'B@yed133')
                    time.sleep(1)
                    browser.open_page('https://qkr-mss.qkrschool.com/qkr_mss/app/storeFront#/inventory')
                    time.sleep(2)
                    browser.click_button(by=By.XPATH, value='//*[@id="addProduct"]/span')
                    time.sleep(1)
                    browser.add_input(by=By.XPATH, value='//*[@id="productname"]', text=excursionname)
                    time.sleep(.1)
                    browser.add_input(by=By.XPATH, value='//*[@id="displaycategoryName1"]', text='Excursions')
                    time.sleep(.1)
                    browser.add_input(by=By.XPATH, value='//*[@id="shortDescription"]', text=date)
                    time.sleep(.1)
                    browser.add_input(by=By.XPATH, value='//*[@id="amount"]', text=cost)
                    time.sleep(.1)
                    browser.click_button(by=By.XPATH, value='//*[@id="saveBtn"]')
                else:
                    i = i + 1  # Try the next attachment
                    
            os.remove(os.getcwd() + r'\workspace\clients\email_client\downloads' + '\\' + attachment.FileName) # Remove file upon completion
            
def camp_check():
    browser = Browser(driverlocation)
    for msg in inbox.Folders(folders['1 Excursions']).Items:
        if re.match('Camp Approved - .+', msg.Subject) and msg.FlagStatus == 2: # Grab flagged excursion approvals
            subject = msg.Subject
            #print('Subject: ' + subject)

            match = 0
            i = 1
            while match == 0:
                if not re.search('Medical Form', str(msg.Attachments.Item(i))) and not re.search('.png', str(msg.Attachments.Item(i))): # Check that the attachment doesn't contain the terms medical form or .png in the title
                    attachment = msg.Attachments.Item(i)
                    #print('Path: ' + os.getcwd() + r'\workspace\clients\email_client\downloads' + '\\' + attachment.FileName)
                    attachment.SaveAsFile(os.getcwd() + r'\workspace\clients\email_client\downloads' + '\\' + attachment.FileName)
                    match = 1
                    
                    # Read the PDF
                    form = PyPDF2.PdfReader(os.getcwd() + r'\workspace\clients\email_client\downloads' + '\\' + attachment.FileName)
                    data = form.pages[0].extract_text()
                    #print(data)
                    
                    # ************************************************ Data Extraction ************************************************
                    
                    # Excursion Name
                    excursionnamestart = subject.find('Camp Approved')
                    excursionnameend = subject.find('\n')
                    excursionname = subject[excursionnamestart+21:]
                    print('Name: ' + excursionname)               
                    
                    # Excursion Date
                    dateloc = data.find('Date /s')
                    datenl = data.find('\n', dateloc)
                    date = data[dateloc+8:datenl]
                    print('Date: ' + date)
                    
                    # Excursion Cost
                    costloc = data.find('Details of Cost')
                    costnl = data.find('\n', costloc)   
                    cost = data[costloc+16:costnl]     
                    print('Cost: ' + cost)
                    print('\n')
                    
                    # ************************************************ Begin QKR ************************************************
                    
                    browser.open_page('https://qkr-mss.qkrschool.com/qkr_mss/index.html')
                    time.sleep(2)
                    browser.login_qkr('nathan.warrick@education.vic.gov.au', 'B@yed133')
                    time.sleep(1)
                    browser.open_page('https://qkr-mss.qkrschool.com/qkr_mss/app/storeFront#/inventory')
                    time.sleep(2)
                    browser.click_button(by=By.XPATH, value='//*[@id="addProduct"]/span')
                    time.sleep(1)
                    browser.add_input(by=By.XPATH, value='//*[@id="productname"]', text=excursionname)
                    time.sleep(.1)
                    browser.add_input(by=By.XPATH, value='//*[@id="displaycategoryName1"]', text='Camps')
                    time.sleep(.1)
                    browser.add_input(by=By.XPATH, value='//*[@id="shortDescription"]', text=date)
                    time.sleep(.1)
                    browser.add_input(by=By.XPATH, value='//*[@id="amount"]', text=cost)
                    time.sleep(.1)
                    browser.click_button(by=By.XPATH, value='//*[@id="saveBtn"]')
                else:
                    i = i + 1  # Try the next attachment
                    
            os.remove(os.getcwd() + r'\workspace\clients\email_client\downloads' + '\\' + attachment.FileName) # Remove file upon completion
            
def qkr_cases_report():
    for msg in inbox.Folders(folders['QKR To Process']).Items:
        if re.match('Qkr! Accounting System.+', msg.Subject) and msg.FlagStatus == 2: # Grab flagged excursion approvals
            subject = msg.Subject
            #print('Subject: ' + subject)

            match = 0
            i = 1
            while match == 0:
                if re.search('.csv', str(msg.Attachments.Item(i))): # Check if the file is a csv
                    attachment = msg.Attachments.Item(i)
                    month = qkrfolder.get(attachment.FileName[21:24]) # Extract month from attachment name and reference with dictionary for file path translation
                    year = attachment.FileName[17:21] # Extract year from attachment name
                    attachment.SaveAsFile(r'U:\PUBLIC\Finance\.QKR Files' + '\\' + year + '\\' + month + '\\' + attachment.FileName) # Save .csv to the U drive
                    #return(r'U:\PUBLIC\Finance\.QKR Files' + '\\' + year + '\\' + month + '\\' + attachment.FileName)
                    match = 1
                    print(r'U:\PUBLIC\Finance\.QKR Files' + '\\' + year + '\\' + month + '\\' + attachment.FileName)
                          
                else:
                    i = i + 1  # Try the next attachment
                            
def qkr_transaction_report():
    for msg in inbox.Folders(folders['QKR To Process']).Items:
        if re.match('Qkr! Report Transaction Details Report', msg.Subject) and msg.FlagStatus == 2: # Grab flagged excursion approvals
            subject = msg.Subject
            print('Subject: ' + subject)

            match = 0
            i = 1
            while match == 0:
                if re.search('.xls', str(msg.Attachments.Item(i))): # Check if the file is a csv
                    attachment = msg.Attachments.Item(i)
                    #print('Path: ' + os.getcwd() + r'\workspace\clients\email_client\downloads' + '\\' + attachment.FileName)
                    attachment.SaveAsFile(os.getcwd() + r'\workspace\clients\email_client\downloads' + '\\' + attachment.FileName)
                    match = 1           
                else:
                    i = i + 1  # Try the next attachment
                    
            #os.remove(os.getcwd() + r'\workspace\clients\email_client\downloads' + '\\' + attachment.FileName) # Remove file upon completion
            
print(qkr_cases_report())
#qkr_transaction_report()