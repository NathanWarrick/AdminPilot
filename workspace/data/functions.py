import os
from datetime import date, datetime
from time import sleep
import pandas as pd
import customtkinter
import pyautogui
import win32com.client

import workspace.data.functionsadvanced as functionsadv
from workspace.ui import popups as guis

def click(path):
    img = pyautogui.locateCenterOnScreen(path, confidence=.98)
    click_x = img[0]
    click_y = img[1]
    pyautogui.leftClick(x=click_x, y=click_y)
    sleep(.3)

def cases_check():
    try:
        functionsadv.focus('CASES21')
        sleep(.5)
    except:
        print("Can't find the application")    

def cases_find(name, tries):
    cases_check()
    sleep(.3)
    pyautogui.hotkey('ctrl', 'f')
    sleep(.3)
    pyautogui.typewrite(name)
    sleep(.3)
    i = 0
    while i < tries:
        pyautogui.press("ENTER")
        sleep(.3)
        i = i + 1
    sleep(.3)
    pyautogui.press("esc")
    sleep(.3)
    pyautogui.press("ENTER")

def print_bank_deposit(): 
    print("Print Bank Deposit")
    click(r'workspace/assets/general/Print.png')
    click(r'workspace/assets/general/Bank Deposit Slip.png')
    sleep(10)
    if str(pyautogui.locateOnScreen(r'workspace/assets/general/Print Job Notification.png')) != "None":
        print("Found")
        while str(pyautogui.locateOnScreen(r'workspace/assets/general/Print Job Notification zAdmininstration.png')) == "None":
            print("Inorrect Account")
            pyautogui.typewrite("z")
        else:
            print("Correct Account")
            pyautogui.press("ENTER")
            
    else:
        print("Can't find Papercut notification")
    sleep(3)

def print_bank_deposit_fake(): 
    print("Fake Bank Deposit")
    click(r'workspace/assets/general/Print.png')
    click(r'workspace/assets/general/Bank Deposit Slip.png')
    sleep(8)
    if str(pyautogui.locateOnScreen(r'workspace/assets/general/Print Job Notification.png')) != "None":
        print("Found")
        pyautogui.hotkey('alt','f4')
            
    else:
        print("Can't find Papercut notification")

def print_online_print(): 
    print("Online Print")
    click(r'workspace/assets/general/Print.png')
    click(r'workspace/assets/general/Online Print.png')
    sleep(10)

def print_audit_trail():
    print("Print Audit Trail")
    batch = functionsadv.batch_report()
    print("Batch Number")
    print(batch)
    click(r'workspace/assets/general/Print.png')
    click(r'workspace/assets/general/Audit Trail.png')
    sleep(10)
    if str(pyautogui.locateOnScreen(r'workspace/assets/general/Filename_Dark.png')) != "None":
        click(r'workspace/assets/general/Filename_Dark.png')
    else:
        if str(pyautogui.locateOnScreen(r'workspace/assets/general/Filename2.png')) != "None":
            click(r'workspace/assets/general/Filename_Light.png')
    sleep(1)
    pyautogui.typewrite(batch)
    pyautogui.typewrite(" by AdminPilot")
    pyautogui.press("Enter")
    sleep(3)
    click(r'workspace/assets/general/Batch Print Yes.png')
    sleep(2)
    pyautogui.hotkey('alt','f4')

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
            <br><br>
            <br><br>
            
            </h1>
            <p class="adminpilot">
                Sent with AdminPilot
                <br><br>
                https://github.com/NathanWarrick/AdminPilot
            </p>
            <style>
            h1 {
                text-shadow: 1px 1px;
                text-align: left;
                font-family: sans-serif;
                font-size: 20px;
                color: black;
                }
            .adminpilot {
                text-shadow: 0px 0px;
                text-align: left;
                font-family: sans-serif;
                font-size: 15px;
                color: black;
                font-style: italic;
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
            <br><br>
            <br><br>
            <br><br>
            </h1>
            <p class="adminpilot">
                Sent with AdminPilot
                <br><br>
                https://github.com/NathanWarrick/AdminPilot
            </p>

            <style>
            h1 {
                text-shadow: 1px 1px;
                text-align: left;
                font-family: sans-serif;
                font-size: 20px;
                color: black;
                }
            .adminpilot {
                text-shadow: 0px 0px;
                text-align: left;
                font-famsily: sans-serif;
                font-size: 15px;
                color: black;
                font-style: italic;
                }
            </style>
            ''')
    # mailItem.Display() #email is displayed prior to sending
    mailItem.Send() #email is sent

def student_ID(name):
      
    #Email is created and processed
    ol = win32com.client.Dispatch('Outlook.Application')
    mailItem = ol.CreateItem(0)
    mailItem.BodyFormat = 1
    mailItem.To = '8818-helpdesk@schools.vic.edu.au' # enter IT email here
    mailItem.Subject = "Student ID Request"
    mailItem.htmlBody = ('''
        <p>
        Hi IT,
        <br><br>
        Can i please get a Student IT card made up for the following student as they have paid their $5 fee.
        </p>       
        <p class="bolded">
        '''
        +str(name)+ 
        '''
        </p>      
        <p>
        <br><br>
        Thank you!
        <br><br>   
        <br><br> 
        <br><br>  
        </p>
    
        <p class="adminpilot">
        Sent with AdminPilot
        <br><br>
        https://github.com/NathanWarrick/AdminPilot
        </p>
        
        <style>
        p {
            text-shadow: 0px 0px;
            text-align: left;
            font-family: sans-serif;
            font-size: 18px;
            color: black;
            }
        .bolded {
            font-weight: bold;
            text-shadow: 0px 0px;
            text-align: left;
            font-family: sans-serif;
            font-size: 20px;
            color: black;
            }
        <style>
        .adminpilot {
            text-shadow: 0px 0px;
            text-align: left;
            font-family: sans-serif;
            font-size: 15px;
            color: black;
            font-style: italic;
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
    cases_find('DF31062', 1)
    pyautogui.press("Enter")
    sleep(4)
    pyautogui.press("Enter")
    sleep(2)
    pyautogui.press("Enter")
    pyautogui.moveTo(10,10)
    sleep(3)
    
    # If there are no records close the BPAY menu
    if str(pyautogui.locateOnScreen(r'workspace/assets/financial/Errors/There are no records to generate the batch with.png')) != "None":
        print("No BPAY!")
        pyautogui.press("Enter")
        pyautogui.hotkey('alt','f4')
        return()
    print_bank_deposit()
    sleep(4)
    print_audit_trail()

def QKR_Canteen(total, receipt_date):
    print("Processing QKR Canteen")
    cases_find('GL31061', 1)
    sleep(4)
    pyautogui.press("Enter")
    pyautogui.moveTo(10,10)
    sleep(4) 
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
    click(r'workspace/assets/general/Save.png')
    print_bank_deposit_fake()
    sleep(3)
    print_audit_trail()

def Canteen(cash_total, eft1_total, eft2_total, receipt_date):
    print("Processing Canteen Payments")
    
    if receipt_date != "":
        receipt_date = date.today().strftime("%d/%m/%Y")
    cashdone = 0
    eft1done = 0
    eft2done = 0
    
    # Canteen Cash
    if cash_total != "":
        cases_find('GL31061', 1)
        sleep(4)
        pyautogui.press("Enter")
        pyautogui.moveTo(10,10)
        sleep(3)
        pyautogui.press("Tab")
        pyautogui.typewrite("CANTEEN")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.typewrite(cash_total)
        pyautogui.press("TAB")
        pyautogui.typewrite('Canteen ')
        pyautogui.typewrite(receipt_date)
        pyautogui.typewrite(' CSH ')
        pyautogui.press("TAB")
        pyautogui.typewrite("CA")
        pyautogui.press("TAB")
        cash_gl = functionsadv.reference_report()
        click(r'workspace/assets/general/Save.png')
        print_online_print()
        print_bank_deposit()
        print_audit_trail()
        cashdone = 1
        
        sleep(5)
    
    # Canteen Eft 1
    if eft1_total != "":
        if cashdone != 1: 
            cases_find('GL31061', 1)
        else:
            pyautogui.press("Enter")
        sleep(4)
        pyautogui.press("Enter")
        pyautogui.moveTo(10,10)
        sleep(3)
        pyautogui.press("Tab")
        pyautogui.typewrite("CANTEEN")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.typewrite(eft1_total)
        pyautogui.press("TAB")
        pyautogui.typewrite('Canteen ')
        pyautogui.typewrite(receipt_date)
        pyautogui.typewrite(' EFT1 ')
        pyautogui.press("TAB")
        pyautogui.typewrite("EF")
        pyautogui.press("TAB")
        eft1_gl = functionsadv.reference_report()
        click(r'workspace/assets/general/Save.png')
        print_online_print()
        print_bank_deposit()
        print_audit_trail()
        eft1done = 1
        
        sleep(5)
    
    # Canteen Eft 2
    if eft2_total != "":
        if cashdone != 1:
            if eft1done != 1:
                cases_find('GL31061', 1)
            else:
                pyautogui.press("Enter")
        else:
            pyautogui.press("Enter")
        sleep(4)
        pyautogui.press("Enter")
        pyautogui.moveTo(10,10)
        sleep(3)
        pyautogui.press("Tab")
        pyautogui.typewrite("CANTEEN")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        pyautogui.typewrite(eft2_total)
        pyautogui.press("TAB")
        pyautogui.typewrite('Canteen ')
        pyautogui.typewrite(receipt_date)
        pyautogui.typewrite(' EFT2 ')
        pyautogui.press("TAB")
        pyautogui.typewrite("EF")
        pyautogui.press("TAB")
        eft2_gl = functionsadv.reference_report()
        click(r'workspace/assets/general/Save.png')
        print_online_print()
        print_bank_deposit()
        print_audit_trail()
        eftdone = 1
        
    guis.Canteen_Overview(cash_total, cash_gl, eft1_total, eft1_gl, eft2_total, eft2_gl, receipt_date).mainloop()

def CSEF():
    print("CSEF Code goes here")
    cases_find('DF21310', 1)
    sleep(3)
    pyautogui.press("Enter")
    pyautogui.moveTo(10,10)
    sleep(3)
    
    if str(pyautogui.locateOnScreen(r'workspace/assets/financial/Errors/There are no records to generate the batch with.png')) != "None":
        print("No BPAY!")
        pyautogui.press("Enter")
        pyautogui.hotkey('alt','f4')
        return()
    print_bank_deposit()
    sleep(4)
    print_audit_trail()
    
def Vehigle_GL():
    filename = customtkinter.filedialog.askopenfilename(title="Dialog box")
    print(filename)
    df = pd.read_excel(filename)

    # Clean up the data
    df = df.drop(labels=range(0,3), axis=0)
    df = df.rename(columns={'Allocation of Motor Vehicle Costs': 'Date', 'Unnamed: 1': 'Department', 'Unnamed: 2': 'KMs', 'Unnamed: 3': 'Cost', 'Unnamed: 4': 'Total', 'Unnamed: 5': 'Sub Prog', 'Unnamed: 6': 'Driver'})
    df = df.dropna(how='any')
    print(df)

    # Initialise
    size = df.shape[0] + 3
    i = 3
    total = 0
    vehicle = None

    months = {'01': 'JAN','02': 'FEB','03': 'MAR','04': 'APR','05': 'MAY','06': 'JUN','07': 'JUL','08': 'AUG','09': 'SEP','10': 'OCT','11': 'NOV','12': 'DEC'}
    lowmonth = '12'
    highmonth = '0'


    # Determine mode of transport
    if str(df.loc[3]['Cost']) == '0.66':
        vehicle = 'Bus'
        print("Bus")
    elif str(df.loc[3]['Cost']) == '0.25':
        vehicle = 'Car'
        print("Car")
    else:
        print("Error: Cost not associated with Vehicle")
        exit()

    #return
    cases_find('GL31081S', 1)
    sleep(3)
    pyautogui.press("Enter")
    sleep(3)
    pyautogui.press("TAB")

    # Loop to run for all entries
    while i < size:
        date = str(df.loc[i]['Date'])
        monthword = date[5:7]
        month = date[5:7]
        day = date[8:10]
        
        
        pyautogui.typewrite(str(df.loc[i]['Sub Prog']))
        pyautogui.press("TAB")
        if vehicle == 'Bus':
            pyautogui.typewrite('89302')
        else:
            pyautogui.typewrite('86701')
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        #pyautogui.typewrite(vehicle + ' Usage ' + months.get(monthword) + ' ' + str(df.loc[i]['Driver']))
        pyautogui.typewrite(vehicle + ' Usage ' + day + '-' + month + ' ' + str(df.loc[i]['Driver']))
        pyautogui.press("TAB")
        pyautogui.typewrite(str(df.loc[i]['Total']))
        pyautogui.press("TAB")
        pyautogui.press("TAB")
        
        if month > highmonth:
            highmonth = month
        if month < lowmonth:
            lowmonth = month
        total = total + df.loc[i]['Total']
        i = i + 1
        
    total = str(total)
    year = date[2:4]
    if vehicle == 'Bus':
        pyautogui.typewrite('9360')
    else:
        pyautogui.typewrite('9361')
    pyautogui.press("TAB")
    pyautogui.typewrite('86701')
    pyautogui.press("TAB")
    pyautogui.press("TAB")
    if vehicle == 'Bus':
        pyautogui.typewrite('Bus Usage ' + months.get(lowmonth) + '-' + months.get(highmonth) + ' ' + year) 
    else:
        pyautogui.typewrite('Car Usage ' + months.get(lowmonth) + '-' + months.get(highmonth) + ' ' + year)   
    pyautogui.press("TAB")
    pyautogui.press("TAB")
    pyautogui.typewrite(total[0:6])
    pyautogui.press("TAB")
    
    print_bank_deposit()
    print_audit_trail()
    
# Accounts Payable

# Student Records

# Business Manager