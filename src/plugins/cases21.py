import src.functions as fnc
from time import sleep
import keyboard
from pynput.keyboard import Key, Controller
import version
import win32api
import mss
from PIL import Image
import cv2
import pytesseract
import pandas as pd
import customtkinter
from datetime import date, datetime

__version__ = version.version

keyboard = Controller()

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract"


def cases_check():
    try:
        fnc.focus("CASES21")
        sleep(1)
    except:
        print("Can't find the application")


def cases_find(name, tries):
    cases_check()
    sleep(0.3)
    keyboard.press(Key.ctrl.value)
    keyboard.type("f")
    keyboard.release(Key.ctrl.value)
    sleep(0.3)
    keyboard.type(name)
    sleep(0.3)
    i = 0
    while i < tries:
        keyboard.press(Key.enter.value)
        sleep(0.3)
        i = i + 1
    sleep(0.3)
    keyboard.press(Key.esc.value)
    sleep(0.3)
    keyboard.press(Key.enter.value)


def print_bank_deposit():
    print("Print Bank Deposit")
    fnc.clickon(r"src/assets/cases21/general/Print.png")
    sleep(0.2)
    fnc.clickon(r"src/assets/cases21/general/Bank Deposit Slip.png")
    sleep(10)

    x, y = fnc.imagesearch(r"src/assets/general/Print Job Notification.png")
    if x != -1:
        print("Found")
        fnc.click_left(x, y)
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


def print_bank_deposit_fake():
    print("Fake Bank Deposit")
    fnc.clickon(r"src/assets/cases21/general/Print.png")
    sleep(0.2)
    fnc.clickon(r"src/assets/cases21/general/Bank Deposit Slip.png")
    sleep(10)
    x, y = fnc.imagesearch(r"src/assets/general/Print Job Notification.png")
    if x != -1:
        print("Found")
        fnc.click_left(x, y)
        sleep(1)
        keyboard.press(Key.alt.value)
        keyboard.type(Key.f4.value)
        keyboard.release(Key.alt.value)

    else:
        print("Can't find Papercut notification")


def print_online_print():
    print("Online Print")
    fnc.clickon(r"src/assets/cases21/general/Print.png")
    sleep(0.2)
    fnc.clickon(r"src/assets/cases21/general/Online Print.png")
    sleep(10)


def print_audit_trail():
    print("Print Audit Trail")
    batch = batch_report()
    print("Batch Number")
    print(batch)
    fnc.clickon(r"src/assets/cases21/general/Print.png")
    sleep(0.2)
    fnc.clickon(r"src/assets/cases21/general/Audit Trail.png")
    sleep(10)
    if fnc.imagesearch(r"src/assets/general/Filename_Dark.png") != [-1, -1]:
        fnc.clickon(r"src/assets/general/Filename_Dark.png")
    else:
        if fnc.imagesearch(r"src/assets/general/Filename_Light.png") != [-1, -1]:
            fnc.clickon(r"src/assets/general/Filename_Light.png")
    sleep(1)
    keyboard.type(batch + " by AdminPilot v" + __version__)
    keyboard.press(Key.enter.value)
    sleep(3)
    fnc.clickon(r"src/assets/cases21/general/Batch Print Yes.png")
    sleep(2)
    keyboard.press(Key.alt.value)
    keyboard.type(Key.f4.value)
    keyboard.release(Key.alt.value)
    return batch


##TODO: Condense Batch, reference and family report to use a function in functions.py for expandability later on


def batch_report():
    """Returns the batch number"""
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


def reference_report():
    """Returns the batch reference number"""
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
        return "Can't find reference number"

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


def family_report():
    """Returns the family code"""
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


##TODO: Develop Centerpay
def Centerpay(student_code, receipt_date, payment_total, fee_total):
    print("Centerpay Code Here")
    # Centerpay is going to be a bit of a challenge to do well.
    # Leave Centerpay for last


def BPAY():
    cases_find("DF31062", 1)
    keyboard.press(Key.enter.value)
    sleep(4)
    keyboard.press(Key.enter.value)
    sleep(2)
    keyboard.press(Key.enter.value)
    # pyautogui.moveTo(10, 10)
    sleep(3)

    # If there are no records close the BPAY menu
    if fnc.imagesearch(
        r"src/assets/cases21/financial/Errors/There are no records to generate the batch with.png"
    ) != [-1, -1]:
        print("No BPAY!")
        keyboard.press(Key.enter.value)
        keyboard.press(Key.alt.value)
        keyboard.type(Key.f4.value)
        keyboard.release(Key.alt.value)
        return ()
    print_bank_deposit()
    sleep(4)
    print_audit_trail()


def QKR_Canteen(total, receipt_date):
    print("Processing QKR Canteen")
    cases_find("GL31061", 1)
    sleep(4)
    keyboard.press(Key.enter.value)
    sleep(6)
    keyboard.press(Key.tab.value)
    fnc.kbd_type("CANTEEN")
    sleep(0.1)
    keyboard.press(Key.tab.value)
    sleep(0.1)
    keyboard.press(Key.tab.value)
    sleep(0.1)
    keyboard.press(Key.tab.value)
    sleep(0.1)
    keyboard.press(Key.tab.value)
    sleep(0.1)
    keyboard.press(Key.tab.value)
    sleep(0.1)
    keyboard.press(Key.tab.value)
    sleep(0.1)
    keyboard.press(Key.tab.value)
    sleep(0.1)
    fnc.kbd_type(total)
    sleep(0.1)
    keyboard.press(Key.tab.value)
    sleep(0.1)
    fnc.kbd_type("QKR Canteen " + receipt_date)
    sleep(0.1)
    keyboard.press(Key.tab.value)
    sleep(0.1)
    fnc.kbd_type("EF")
    sleep(0.1)
    keyboard.press(Key.tab.value)
    sleep(0.1)
    fnc.clickon(r"src/assets/cases21/general/Save.png")
    print_bank_deposit_fake()
    sleep(3)
    print_audit_trail()


def Canteen(
    cash_total, eft1_total, eft2_total, receipt_date
):  # Issues negotiating papercut print window
    print("Processing Canteen Payments")

    if receipt_date == "":
        receipt_date = date.today().strftime("%d/%m/%Y")
    cashdone = 0
    eft1done = 0
    eft2done = 0

    # Canteen Cash
    if cash_total != "":
        cases_find("GL31061", 1)
        sleep(4)
        keyboard.press(Key.enter.value)
        sleep(6)
        keyboard.press(Key.tab.value)
        sleep(1)
        fnc.kbd_type("CANTEEN")
        sleep(1)
        keyboard.press(Key.tab.value)
        sleep(1)
        keyboard.press(Key.tab.value)
        sleep(1)
        keyboard.press(Key.tab.value)
        sleep(1)
        keyboard.press(Key.tab.value)
        sleep(1)
        keyboard.press(Key.tab.value)
        sleep(1)
        keyboard.press(Key.tab.value)
        sleep(1)
        keyboard.press(Key.tab.value)
        sleep(1)
        fnc.kbd_type(cash_total)
        sleep(1)
        keyboard.press(Key.tab.value)
        sleep(1)
        fnc.kbd_type("Canteen " + receipt_date + " CSH ")
        sleep(1)
        keyboard.press(Key.tab.value)
        sleep(1)
        fnc.kbd_type("CA")
        sleep(1)
        keyboard.press(Key.tab.value)
        sleep(1)
        cash_gl = reference_report()
        fnc.clickon(r"src/assets/cases21/general/Save.png")
        # print_online_print()
        print_bank_deposit()
        # cash_batch = print_audit_trail()
        cashdone = 1

        sleep(5)
        return

    # Canteen Eft 1
    if eft1_total != "":
        if cashdone != 1:
            cases_find("GL31061", 1)
        else:
            keyboard.press(Key.enter.value)
        sleep(4)
        keyboard.press(Key.enter.value)
        sleep(6)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        fnc.kbd_type("CANTEEN")
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        fnc.kbd_type(eft1_total)
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        fnc.kbd_type("Canteen " + receipt_date + " EFT 1")
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        fnc.kbd_type("EF")
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        eft1_gl = reference_report()
        fnc.clickon(r"src/assets/cases21/general/Save.png")
        print_online_print()
        print_bank_deposit()
        eft1_batch = print_audit_trail()
        eft1done = 1

        sleep(5)

    # Canteen Eft 2
    if eft2_total != "":
        if cashdone != 1:
            if eft1done != 1:
                cases_find("GL31061", 1)
            else:
                keyboard.press(Key.enter.value)
        else:
            keyboard.press(Key.enter.value)
        sleep(4)  # TODO: Change these waits to object recognition logic
        keyboard.press(Key.enter.value)
        # pyautogui.moveTo(10, 10)
        sleep(3)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        fnc.kbd_type("CANTEEN")
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        fnc.kbd_type(eft2_total)
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        fnc.kbd_type("Canteen " + receipt_date + " EFT2 ")
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        fnc.kbd_type("EF")
        sleep(0.1)
        keyboard.press(Key.tab.value)
        sleep(0.1)
        eft2_gl = reference_report()
        fnc.clickon(r"src/assets/cases21/general/Save.png")
        print_online_print()
        print_bank_deposit()
        eft2_batch = print_audit_trail()
        eft2done = 1

    return (
        cash_total,
        # cash_batch,
        cash_total,
        eft1_total,
        eft1_batch,
        eft1_gl,
        eft2_total,
        eft2_batch,
        eft2_gl,
    )

    ##TODO: Reimplement GUI
    # guis.Canteen_Overview(
    #     cash_total, cash_gl, eft1_total, eft1_gl, eft2_total, eft2_gl, receipt_date
    # ).mainloop()


def CSEF():
    print("CSEF Code goes here")
    cases_find("DF21310", 1)
    sleep(3)
    keyboard.press(Key.enter.value)
    # pyautogui.moveTo(10, 10)
    sleep(3)

    if fnc.imagesearch(
        r"src/assets/cases21/financial/Errors/There are no records to generate the batch with.png"
    ) != [-1, -1]:
        print("No CSEF!")
        keyboard.press(Key.enter.value)
        keyboard.press(Key.alt.value)
        keyboard.type(Key.f4.value)
        keyboard.release(Key.alt.value)
        return ()
    print_bank_deposit()
    sleep(4)
    print_audit_trail()


def Vehigle_GL():
    filename = customtkinter.filedialog.askopenfilename(title="Dialog box")
    print(filename)
    df = pd.read_excel(filename)

    # Clean up the data
    df = df.drop(labels=range(0, 3), axis=0)
    df = df.rename(
        columns={
            "Allocation of Motor Vehicle Costs": "Date",
            "Unnamed: 1": "Department",
            "Unnamed: 2": "KMs",
            "Unnamed: 3": "Cost",
            "Unnamed: 4": "Total",
            "Unnamed: 5": "Sub Prog",
            "Unnamed: 6": "Driver",
        }
    )
    df = df.dropna(how="any")
    print(df)

    # Initialise
    size = df.shape[0] + 3
    i = 3
    total = 0
    vehicle = None

    months = {
        "01": "JAN",
        "02": "FEB",
        "03": "MAR",
        "04": "APR",
        "05": "MAY",
        "06": "JUN",
        "07": "JUL",
        "08": "AUG",
        "09": "SEP",
        "10": "OCT",
        "11": "NOV",
        "12": "DEC",
    }
    lowmonth = "12"
    highmonth = "0"

    # Determine mode of transport
    if str(df.loc[3]["Cost"]) == "0.66":
        vehicle = "Bus"
        print("Bus")
    elif str(df.loc[3]["Cost"]) == "0.25":
        vehicle = "Car"
        print("Car")
    else:
        print("Error: Cost not associated with Vehicle")
        exit()

    # return
    cases_find("GL31081S", 1)
    sleep(3)
    keyboard.press(Key.enter.value)
    sleep(3)
    keyboard.press(Key.tab.value)

    # Loop to run for all entries
    while i < size:
        date = str(df.loc[i]["Date"])
        monthword = date[5:7]
        month = date[5:7]
        day = date[8:10]

        fnc.kbd_type(str(df.loc[i]["Sub Prog"]))
        keyboard.press(Key.tab.value)
        if vehicle == "Bus":

            fnc.kbd_type("89302")
        else:
            fnc.kbd_type("86701")
        keyboard.press(Key.tab.value)
        keyboard.press(Key.tab.value)
        # fnc.kbd_type(vehicle + ' Usage ' + months.get(monthword) + ' ' + str(df.loc[i]['Driver']))
        fnc.kbd_type(
            vehicle + " Usage " + day + "-" + month + " " + str(df.loc[i]["Driver"])
        )
        keyboard.press(Key.tab.value)
        fnc.kbd_type(str(df.loc[i]["Total"]))
        keyboard.press(Key.tab.value)
        keyboard.press(Key.tab.value)

        if month > highmonth:
            highmonth = month
        if month < lowmonth:
            lowmonth = month
        total = total + df.loc[i]["Total"]
        i = i + 1

    total = str(total)
    year = date[2:4]
    if vehicle == "Bus":
        fnc.kbd_type("9360")
    else:
        fnc.kbd_type("9361")
    keyboard.press(Key.tab.value)
    fnc.kbd_type("86701")
    keyboard.press(Key.tab.value)
    keyboard.press(Key.tab.value)
    if vehicle == "Bus":
        fnc.kbd_type(
            "Bus Usage "
            + months.get(lowmonth)
            + "-"
            + months.get(highmonth)
            + " "
            + year
        )
    else:
        fnc.kbd_type(
            "Car Usage "
            + months.get(lowmonth)
            + "-"
            + months.get(highmonth)
            + " "
            + year
        )
    keyboard.press(Key.tab.value)
    keyboard.press(Key.tab.value)
    fnc.kbd_type(total[0:6])
    keyboard.press(Key.tab.value)

    print_bank_deposit()
    print_audit_trail()
