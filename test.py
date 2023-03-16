import pyautogui
import os
import Functions as function
import time
os.chdir(os.path.dirname(os.path.abspath(__file__)))

function.click('Assets/General/Print.png')
function.click('Assets/General/Bank Deposit Slip.png')
time.sleep(8)
if str(pyautogui.locateOnScreen('Assets/General/Print Job Notification.png')) != "None":
    print("Found")
    while str(pyautogui.locateOnScreen('Assets/General/Print Job Notification zAdmininstration.png')) == "None":
        print("Inorrect Account")
        pyautogui.typewrite("z")
    else:
        print("Correct Account")
        pyautogui.press("Enter")
        
else:
    print("Can't find Papercut notification")