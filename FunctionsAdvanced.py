import os
import re

import cv2
import pyautogui
import pytesseract
import win32gui

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'

os.chdir(os.path.dirname(os.path.abspath(__file__)))



class WindowMgr:
    """Encapsulates some calls to the winapi for window management"""

    def __init__ (self):
        """Constructor"""
        self._handle = None

    def find_window(self, class_name, window_name=None):
        """find a window by its class_name"""
        self._handle = win32gui.FindWindow(class_name, window_name)

    def _window_enum_callback(self, hwnd, wildcard):
        """Pass to win32gui.EnumWindows() to check all the opened windows"""
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) is not None:
            self._handle = hwnd

    def find_window_wildcard(self, wildcard):
        """find a window whose title matches the wildcard regex"""
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)

    def set_foreground(self):
        """put the window in the foreground"""
        win32gui.SetForegroundWindow(self._handle)
        
def batch_report():
    Batch_Location = pyautogui.locateCenterOnScreen('Assets/General/Batch_Number.png')
    Batch_Location_X = Batch_Location[0]
    Batch_Location_Y = Batch_Location[1]
    im = pyautogui.screenshot(region=((Batch_Location_X+45), (Batch_Location_Y-10), 40, 15),)
    im.save('Assets/Temp/batch.png')

    img = cv2.imread('Assets/Temp/batch.png')
    img = cv2.resize(img, None, fx=2, fy=2)
    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        
    config = '--oem 3 --psm 6' 
    batch = pytesseract.image_to_string(img, config = config, lang='eng')
    batch = batch[:5]
    os.remove('Assets/Temp/batch.png')
    return batch

def focus(focus_on):
    focus_program = focus_on
    print("Bringing %s to front" % focus_on)
    focus = WindowMgr()
    focus.find_window_wildcard(".*%s.*" % focus_program)
    focus.set_foreground()