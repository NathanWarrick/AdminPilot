import src.functions as fnc

# import src.plugins.seqta as sq

import src.plugins.cases21 as cases
import os
import win32api
import random
from time import sleep
import keyboard
import pynput
from pynput.keyboard import Key, Controller
import csv, os, mss, cv2, pynput, time, win32api, win32con
import numpy as np
from PIL import Image
import pytesseract

import src.plugins.outlook as otl

curr_working_dir = os.getcwd()
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract"

keyboard = Controller()

# # Find batch number image
# x, y = fnc.imagesearch(r"src/assets/cases21/general/Batch_Number.png")
# # Reverse coordinate corrections
# x = x - win32api.GetSystemMetrics(76)
# y = y - win32api.GetSystemMetrics(77)
# # Take screenshot and crop
# with mss.mss() as sct:
#     sct_img = sct.grab(sct.monitors[0])
#     img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
#     im = img.crop((x, y, x + 40, y + 15))
#     output = "test.png"
#     im.save(r"src/assets/temp/batch.png")
#     print(output)


def reference_report():
    try:
        x, y = fnc.imagesearch(r"src/assets/cases21/general/Reference.png")
        # Reverse coordinate corrections
        x = x - win32api.GetSystemMetrics(76)
        y = y - win32api.GetSystemMetrics(77)
        # Take screenshot and crop
        with mss.mss() as sct:
            sct_img = sct.grab(sct.monitors[0])
            img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
            im = img.crop((x + 35, y - 10, x + 105, y + 10))
            im.save(r"src/assets/temp/reference.png")
    except:
        return "Can't find batch number"

    img = cv2.imread(r"src/assets/temp/reference.png")
    img = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    config = "--oem 3 --psm 6 -c load_system_dawg=false -c load_freq_dawg=false -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"  # Removed O from whitelist
    reference = pytesseract.image_to_string(
        img,
        config=config,
        lang=None,
    )

    reference = reference[:3] + reference[3:10].replace("O", "0")
    return reference


def batch_report():
    try:
        x, y = fnc.imagesearch(r"src/assets/cases21/general/Batch_Number.png")
        # Reverse coordinate corrections
        x = x - win32api.GetSystemMetrics(76)
        y = y - win32api.GetSystemMetrics(77)
        # Take screenshot and crop
        with mss.mss() as sct:
            sct_img = sct.grab(sct.monitors[0])
            img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
            im = img.crop((x + 40, y - 10, x + 85, y + 10))
            im.save(r"src/assets/temp/batch.png")
    except:
        return "Can't find batch number"

    img = cv2.imread(r"src/assets/temp/batch.png")
    img = cv2.resize(img, None, fx=2, fy=2)
    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    config = "--oem 3 --psm 6 -c load_system_dawg=false -c load_freq_dawg=false -c tessedit_char_whitelist=0123456789"
    batch = pytesseract.image_to_string(img, config=config, lang="eng")
    batch = batch[:5]
    return batch


def family_report():
    try:
        x, y = fnc.imagesearch(r"src/assets/cases21/general/Family.png")
        # Reverse coordinate corrections
        x = x - win32api.GetSystemMetrics(76)
        y = y - win32api.GetSystemMetrics(77)
        # Take screenshot and crop
        with mss.mss() as sct:
            sct_img = sct.grab(sct.monitors[0])
            img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
            im = img.crop((x + 25, y - 10, x + 85, y + 10))
            im.save(r"src/assets/temp/family.png")
    except:
        return "Can't find family code"

    img = cv2.imread(r"src/assets/temp/family.png")
    img = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    config = "--oem 3 --psm 6 -c load_system_dawg=false -c load_freq_dawg=false -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"  # Removed O from whitelist
    family = pytesseract.image_to_string(
        img,
        config=config,
        lang=None,
    )

    family = family[:3] + family[3:7].replace("O", "0")
    return family


# print(family_report())

# otl.student_ID("Nathan Warrick")
# otl.attendance_update("Nathan Warrick", "1/1/01", "", "", "", "")

keyboard.press(Key.tab.value)
