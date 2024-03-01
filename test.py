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
import ctypes

import src.plugins.outlook as otl
import src.plugins.cases21 as css

curr_working_dir = os.getcwd()
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract"

keyboard = Controller()


x, y = fnc.check(r"src/assets/general/Print Job Notification.png")
if x != -1:
    print("Found")
    ctypes.windll.user32.SetCursorPos(int(x), int(y))
    sleep(1)
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
    sleep(0.1)
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)

# css.print_bank_deposit_fake()

# MOUSE_LEFTDOWN = 0x0002  # left button down
# MOUSE_LEFTUP = 0x0004  # left button up
# MOUSE_RIGHTDOWN = 0x0008  # right button down
# MOUSE_RIGHTUP = 0x0010  # right button up
# MOUSE_MIDDLEDOWN = 0x0020  # middle button down
# MOUSE_MIDDLEUP = 0x0040  # middle button up
