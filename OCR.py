import pytesseract
import cv2
import numpy as np
import pyautogui

Batch_Location = pyautogui.locateOnScreen('Batch_Number.png')
Batch_Location = pyautogui.center(Batch_Location)
Batch_Location_X = Batch_Location[0]
Batch_Location_Y = Batch_Location[1]

im = pyautogui.screenshot(region=((Batch_Location_X+45), (Batch_Location_Y-10), 40, 20),)
im.save("batch.png")

img = cv2.imread('batch.png')
#  img = cv2.resize(img, None, fx=1.2, fy=1.2, interpolation=cv2.INTER_CUBIC)
img = cv2.resize(img, None, fx=2, fy=2)

img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'
    
config = '--oem 3 --psm 6' 
batch = pytesseract.image_to_string(img, config = config, lang='eng')
print(batch)