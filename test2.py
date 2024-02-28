import src.functions as fnc
from pynput.keyboard import Key, Controller
from time import sleep
import csv, os, mss, cv2, pynput, time, win32api, win32con

keyboard = Controller()


def click_left(x, y):
    if x is not None and y is not None:
        x = int(x)
        y = int(y)
        win32api.SetCursorPos((x, y))
        sleep(0.1)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
    else:
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)


def print_bank_deposit():
    print("Print Bank Deposit")
    fnc.clickon(r"src/assets/cases21/general/Print.png")
    sleep(0.2)
    fnc.clickon(r"src/assets/cases21/general/Bank Deposit Slip.png")
    sleep(10)

    x, y = fnc.imagesearch(r"src/assets/general/Print Job Notification.png")
    if x != -1:
        print("Found")
        fnc.clickon(r"src/assets/general/Print Job Notification.png")
        sleep(1)

        while fnc.imagesearch(
            r"src/assets/general/Print Job Notification zAdmininstration.png"
        ) == [-1, -1]:
            print("Inorrect Account")
            keyboard.type("z")
        else:
            print("Correct Account")
            keyboard.press(Key.enter.value)

    else:
        print("Can't find Papercut notification")
    sleep(3)


print_bank_deposit()

# x, y = fnc.imagesearch(r"src/assets/general/Print Job Notification.png")
# click_left(x, y)
