import pytesseract
import pyautogui
import cv2
import os
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'

os.chdir(os.path.dirname(os.path.abspath(__file__)))

def batch_report():
    Batch_Location = pyautogui.locateCenterOnScreen('Batch_Number.png')
    Batch_Location_X = Batch_Location[0]
    Batch_Location_Y = Batch_Location[1]

    im = pyautogui.screenshot(region=((Batch_Location_X+45), (Batch_Location_Y-10), 40, 15),)
    im.save("batch.png")

    img = cv2.imread('batch.png')
    img = cv2.resize(img, None, fx=2, fy=2)
    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        
    config = '--oem 3 --psm 6' 
    batch = pytesseract.image_to_string(img, config = config, lang='eng')
    batch = batch[:5]
    return batch