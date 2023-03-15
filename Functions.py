import win32com.client
from datetime import date, datetime
import pyautogui
from time import sleep

import FunctionsAdvanced as functionsadvanced

def click(path):
    img = pyautogui.locateCenterOnScreen(path)
    click_x = img[0]
    click_y = img[1]
    pyautogui.leftClick(x=click_x, y=click_y)
    sleep(.2)

# General
def attendance_update(name, date_str, time_str, returning, reason, collected):
    
    #If date was blank insert todays date
    if date_str == "":
        date_str = date.today().strftime("%d/%m/%Y")
        
    #If time was blank insert current time    
    if time_str == "":
        time_str = datetime.now().strftime("%H:%M")
    
    #If returning was blank insert No
    if returning == "":
        returning = "No"

    
    #Email is created and processed
    ol = win32com.client.Dispatch('Outlook.Application')
    mailItem = ol.CreateItem(0)
    mailItem.BodyFormat = 1
    mailItem.To = 'absences@horsham-college.vic.edu.au'
    mailItem.Subject = "Attendance Update"
    if collected == "":
        mailItem.htmlBody = ('''
            <h1>
            Name:
            '''
            + str(name) + 
            '''
            <br><br>
            
            Date:
            '''
            + str(date_str) + 
            '''
            <br><br>
            
            Time:
            '''
            + str(time_str) + 
            '''
            <br><br>
            
            Returning? If so, what time?
            '''
            + str(returning) + 
            '''
            <br><br>
            
            Reason:
            '''
            + str(reason) + 
            '''
            <br><br>
            
            
            </h1>

            <style>
            h1 {
                text-shadow: 1px 1px;
                text-align: left;
                font-family: sans-serif;
                font-size: 20px;
                color: black;
                }
            </style>
            ''')
    else:
        mailItem.htmlBody = ('''
            <h1>
            Name:
            '''
            + str(name) + 
            '''
            <br><br>
            
            Date:
            '''
            + str(date_str) + 
            '''
            <br><br>
            
            Time:
            '''
            + str(time_str) + 
            '''
            <br><br>
            
            Returning? If so, what time?
            '''
            + str(returning) + 
            '''
            <br><br>
            
            Reason:
            '''
            + str(reason) + 
            '''
            <br><br>
            
            Who Collected:
            '''
            + str(collected) + 
            '''
            
            
            </h1>

            <style>
            h1 {
                text-shadow: 1px 1px;
                text-align: left;
                font-family: sans-serif;
                font-size: 20px;
                color: black;
                }
            </style>
            ''')
    # mailItem.Display() #email is displayed prior to sending
    mailItem.Send() #email is sent

# Accounts Receivable
def Centerpay(student_code, receipt_date, payment_total, fee_total):
    print("Centerpay Code Here")
    # Centerpay is going to be a bit of a challenge to do well. 
    # Leave Centerpay for last
    
def BPAY():
    click('Assets/Financial/Families/Families.png')
    click('Assets/Financial/Families/Process BPAY Receipts.png')
    click('Assets/Financial/Families/BPAY Receipts.png')
    pyautogui.press("Enter")
    sleep(2)
    pyautogui.press("Enter")
    sleep(2)
    pyautogui.press("Enter")
    sleep(1)
    # If there are no records close the BPAY menu
    if str(pyautogui.locateOnScreen('Assets/Financial/Errors/There are no records to generate the batch with.png')) != "None":
        print("No BPAY!")
        pyautogui.press("Enter")
        pyautogui.hotkey('alt','f4')
    
def QKR_Canteen(total, receipt_date):
    print("QKR Canteen Code Here")
    click('Assets/Financial/General Ledger/General Ledger.png')
    click('Assets/Financial/General Ledger/Process Receipts.png')
    click('Assets/Financial/General Ledger/General Ledger Receipt.png')
    pyautogui.press("Enter")
    sleep(2)
    pyautogui.press("Enter")
    sleep(2)
    pyautogui.press("Tab")
    pyautogui.typewrite("CANTEEN")
    pyautogui.press("TAB")
    pyautogui.press("TAB")
    pyautogui.press("TAB")
    pyautogui.press("TAB")
    pyautogui.press("TAB")
    pyautogui.press("TAB")
    pyautogui.press("TAB")
    pyautogui.typewrite(total)
    pyautogui.press("TAB")
    pyautogui.typewrite('QKR Canteen ')
    pyautogui.typewrite(receipt_date)
    pyautogui.press("TAB")
    pyautogui.typewrite("EF")
    pyautogui.press("TAB")
    
    
    
def Canteen(cash_total, eft1_total, eft2_total, receipt_date):
    print("Canteen Code Here")