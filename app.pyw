import os
import sys
import time

import customtkinter
import keyboard

from workspace.bots.papercut_bot import papercut_bot
from workspace.bots.qkr_bot import qkr_bot
from workspace.data import (
    functions as functions,
)  # Import Functions.py file as functions

customtkinter.set_appearance_mode("System")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme(
    "blue"
)  # Themes: blue (default), dark-blue, green

import version

__version__ = version.version


class GUI(customtkinter.CTk):  # Main GUI Config
    def __init__(self):
        super().__init__()

        self.title("AdminPilot")
        self.geometry("450x450")
        self.geometry("+1400+50")
        self.lift()

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # General Windows
        self.absence_window = None
        self.student_ID_window = None
        self.papercut_deposit_window = None
        self.wwcc_window = None

        # Acc Receivable windows
        self.centerpay_window = None
        self.canteen_window = None
        self.qkr_canteen_window = None

        # Acc Payable Windows

        # Student Records Windows

        # Business Manager Windows

        # Navigation Frame
        self.navigation_frame = customtkinter.CTkFrame(self, corner_radius=0)
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(
            6, weight=1
        )  # Change number of rows in the naviation frame

        self.navigation_frame_label = customtkinter.CTkLabel(
            self.navigation_frame,
            text=" AdminPilot v" + __version__,
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.navigation_frame_label.grid(row=0, column=0, padx=20, pady=20)

        # General Button Navigation Frame
        self.General_button = customtkinter.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="General",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.General_button_event,
        )
        self.General_button.grid(row=1, column=0, sticky="ew")

        # Accounts Receivable Button Navigation Frame
        self.Accounts_Receivable_button = customtkinter.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="Accounts Receivable",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.Accounts_Receivable_button_event,
        )
        self.Accounts_Receivable_button.grid(row=2, column=0, sticky="ew")

        # Accounts Payable Button Navigation Frame
        self.Accounts_Payable_button = customtkinter.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="Accounts Payable",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.Accounts_Payable_button_event,
        )
        self.Accounts_Payable_button.grid(row=3, column=0, sticky="ew")

        # Student Records Button Navigation Frame
        self.Student_Records_button = customtkinter.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="Student Records",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.Student_Records_button_event,
        )
        self.Student_Records_button.grid(row=4, column=0, sticky="ew")

        # Business Manager Button Navigation Frame
        self.Business_Manager_button = customtkinter.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="Business Manager",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.Business_Manager_button_event,
        )
        self.Business_Manager_button.grid(row=5, column=0, sticky="ew")

        # Reset Button
        self.Reset_button = customtkinter.CTkButton(
            self.navigation_frame,
            corner_radius=0,
            height=40,
            border_spacing=10,
            text="Restart",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray70", "gray30"),
            anchor="w",
            command=self.Adminpilot_Reset,
        )
        self.Reset_button.grid(row=6, column=0, sticky="sew")

        # Theme button

        self.appearance_mode_menu = customtkinter.CTkOptionMenu(
            self.navigation_frame,
            values=["Light", "Dark", "System"],
            command=self.change_appearance_mode_event,
        )
        self.appearance_mode_menu.grid(row=7, column=0, padx=20, pady=20, sticky="s")

        # Create General Frame
        self.General_frame = customtkinter.CTkFrame(
            self, corner_radius=0, fg_color="transparent"
        )
        self.General_frame.grid_columnconfigure(0, weight=1)

        self.General_frame_large_image_label = customtkinter.CTkLabel(
            self.General_frame, text=""
        )
        self.General_frame_large_image_label.grid(row=0, column=0, padx=20, pady=10)

        self.General_frame_button_1 = customtkinter.CTkButton(
            self.General_frame,
            text="Student Absence",
            command=self.Student_Absence_Button_Event,
        )
        self.General_frame_button_1.grid(row=1, column=0, padx=20, pady=10)
        self.General_frame_button_2 = customtkinter.CTkButton(
            self.General_frame, text="Student ID", command=self.Student_ID_Button_Event
        )
        self.General_frame_button_2.grid(row=2, column=0, padx=20, pady=10)
        self.General_frame_button_3 = customtkinter.CTkButton(
            self.General_frame,
            text="Papercut Deposit",
            compound="top",
            command=self.Papercut_Deposit_Button_Event,
        )
        self.General_frame_button_3.grid(row=3, column=0, padx=20, pady=10)
        self.General_frame_button_4 = customtkinter.CTkButton(
            self.General_frame,
            text="WWCC Check",
            compound="top",
            command=self.WWCC_Check_Button_Event,
        )
        self.General_frame_button_4.grid(row=4, column=0, padx=20, pady=10)

        # Create Acc Rec frame
        self.Accounts_Receivable_frame = customtkinter.CTkFrame(
            self, corner_radius=0, fg_color="transparent"
        )
        self.Accounts_Receivable_frame.grid_columnconfigure(0, weight=1)

        self.Accounts_Receivable_large_image_label = customtkinter.CTkLabel(
            self.Accounts_Receivable_frame, text=""
        )
        self.Accounts_Receivable_large_image_label.grid(
            row=0, column=0, padx=20, pady=10
        )

        self.Accounts_Receivable_button_1 = customtkinter.CTkButton(
            self.Accounts_Receivable_frame,
            text="Centerpay",
            command=self.Centerpay_Button_Event,
            fg_color="dark red",
            hover_color="dark red",
        )
        self.Accounts_Receivable_button_1.grid(row=1, column=0, padx=20, pady=10)
        self.Accounts_Receivable_button_2 = customtkinter.CTkButton(
            self.Accounts_Receivable_frame, text="BPAY", command=self.BPAY_Button_Event
        )
        self.Accounts_Receivable_button_2.grid(row=2, column=0, padx=20, pady=10)
        self.Accounts_Receivable_button_3 = customtkinter.CTkButton(
            self.Accounts_Receivable_frame,
            text="QKR Canteen",
            command=self.QKR_Canteen_Button_Event,
        )
        self.Accounts_Receivable_button_3.grid(row=3, column=0, padx=20, pady=10)
        self.Accounts_Receivable_button_4 = customtkinter.CTkButton(
            self.Accounts_Receivable_frame,
            text="Canteen",
            command=self.Canteen_Button_Event,
        )
        self.Accounts_Receivable_button_4.grid(row=4, column=0, padx=20, pady=10)
        self.Accounts_Receivable_button_5 = customtkinter.CTkButton(
            self.Accounts_Receivable_frame,
            text="CSEF Receipt",
            command=self.CSEF_Button_Event,
        )
        self.Accounts_Receivable_button_5.grid(row=5, column=0, padx=20, pady=10)
        self.Accounts_Receivable_button_5 = customtkinter.CTkButton(
            self.Accounts_Receivable_frame,
            text="QKR Update",
            command=self.QKR_Update_Button_Event,
            fg_color="dark red",
            hover_color="dark red",
        )
        self.Accounts_Receivable_button_5.grid(row=6, column=0, padx=20, pady=10)
        self.Accounts_Receivable_button_5 = customtkinter.CTkButton(
            self.Accounts_Receivable_frame,
            text="Vehicle Expense",
            command=self.Vehicle_Expense_Button_Event,
        )
        self.Accounts_Receivable_button_5.grid(row=7, column=0, padx=20, pady=10)

        # Create Acc Pay frame
        self.Accounts_Payable_frame = customtkinter.CTkFrame(
            self, corner_radius=0, fg_color="transparent"
        )

        # Create Student Rec frame
        self.Student_Records_frame = customtkinter.CTkFrame(
            self, corner_radius=0, fg_color="transparent"
        )

        # Create BM frame
        self.Business_Manager_frame = customtkinter.CTkFrame(
            self, corner_radius=0, fg_color="transparent"
        )

        # select default frame
        self.select_frame_by_name("General")

    def select_frame_by_name(self, name):
        # set button color for selected button
        self.General_button.configure(
            fg_color=("gray75", "gray25") if name == "General" else "transparent"
        )
        self.Accounts_Receivable_button.configure(
            fg_color=("gray75", "gray25")
            if name == "Accounts_Receivable"
            else "transparent"
        )
        self.Accounts_Payable_button.configure(
            fg_color=("gray75", "gray25")
            if name == "Accounts_Payable"
            else "transparent"
        )
        self.Student_Records_button.configure(
            fg_color=("gray75", "gray25")
            if name == "Student_Records"
            else "transparent"
        )
        self.Business_Manager_button.configure(
            fg_color=("gray75", "gray25")
            if name == "Business_Manager"
            else "transparent"
        )

        # show selected frame
        if name == "General":
            self.General_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.General_frame.grid_forget()
        if name == "Accounts_Receivable":
            self.Accounts_Receivable_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.Accounts_Receivable_frame.grid_forget()
        if name == "Accounts_Payable":
            self.Accounts_Payable_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.Accounts_Payable_frame.grid_forget()
        if name == "Student_Records":
            self.Student_Records_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.Student_Records_frame.grid_forget()
        if name == "Business_Manager":
            self.Business_Manager_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.Business_Manager_frame.grid_forget()

    # Goto selected frame
    def General_button_event(self):
        self.select_frame_by_name("General")

    def Accounts_Receivable_button_event(self):
        self.select_frame_by_name("Accounts_Receivable")

    def Accounts_Payable_button_event(self):
        self.select_frame_by_name("Accounts_Payable")

    def Student_Records_button_event(self):
        self.select_frame_by_name("Student_Records")

    def Business_Manager_button_event(self):
        self.select_frame_by_name("Business Manager")

    def Adminpilot_Reset(self):
        print("Restarting")
        os.execv(sys.executable, ["python"] + sys.argv)

    def change_appearance_mode_event(self, new_appearance_mode):
        customtkinter.set_appearance_mode(new_appearance_mode)

    # Scripts for buttons go here
    def Student_Absence_Button_Event(self):
        if self.absence_window is None or not self.absence_window.winfo_exists():
            self.absence_window = Absence()
            return ()
        else:
            self.absence_window.focus()

    def Student_ID_Button_Event(self):
        if self.student_ID_window is None or not self.student_ID_window.winfo_exists():
            self.student_ID_window = Student_ID()
            return ()
        else:
            self.student_ID_window.focus()

    def Papercut_Deposit_Button_Event(self):
        if (
            self.papercut_deposit_window is None
            or not self.papercut_deposit_window.winfo_exists()
        ):
            self.papercut_deposit_window = Papercut_Deposit()
            return ()
        else:
            self.papercut_deposit_window.focus()

    def WWCC_Check_Button_Event(self):
        if self.wwcc_window is None or not self.wwcc_window.winfo_exists():
            self.wwcc_window = wwcc_check()
            return ()
        else:
            self.wwcc_window.focus()

    # Accounts Receivable
    def Centerpay_Button_Event(self):
        if self.centerpay_window is None or not self.centerpay_window.winfo_exists():
            self.centerpay_window = Centerpay()
        else:
            self.centerpay_window.focus()

    def BPAY_Button_Event(self):
        self.bpay_window = BPAY()

    def QKR_Canteen_Button_Event(self):
        if (
            self.qkr_canteen_window is None
            or not self.qkr_canteen_window.winfo_exists()
        ):
            self.qkr_canteen_window = QKR_Canteen()
        else:
            self.qkr_canteen_window.focus()

    def Canteen_Button_Event(self):
        if self.canteen_window is None or not self.canteen_window.winfo_exists():
            self.canteen_window = Canteen()
        else:
            self.canteen_window.focus()

    def CSEF_Button_Event(self):
        self.csef_window = CSEF()

    def QKR_Update_Button_Event(self):
        self.qkr_update_window = QKR_Update()

    def Vehicle_Expense_Button_Event(self):
        self.vehicle_expense_window = Vehicle_Expense()


class Absence(customtkinter.CTkToplevel):
    def __init__(self):
        super().__init__()
        self.geometry("450x400")
        self.lift()
        self.title("Student Absence")

        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.absence_frame = customtkinter.CTkFrame(
            self, corner_radius=0, width=400, height=300
        )
        self.absence_frame.grid(row=0, column=0, sticky="nsew")
        self.absence_frame.grid_rowconfigure(
            6, weight=1
        )  # Change number of rows in the naviation frame

        self.buttons_frame = customtkinter.CTkFrame(
            self, corner_radius=0, width=400, height=100
        )
        self.buttons_frame.grid(row=1, column=0, sticky="nsew")
        self.buttons_frame.grid_rowconfigure(
            1, weight=1
        )  # Change number of rows in the naviation frame
        self.buttons_frame.grid_columnconfigure(3, weight=1)

        # Name Field
        self.absence_frame_label = customtkinter.CTkLabel(
            self.absence_frame,
            text="Name *",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.absence_frame_label.grid(row=0, column=0, padx=10, pady=8)
        self.Absence_Name = customtkinter.CTkEntry(
            self.absence_frame,
            placeholder_text="Name",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.Absence_Name.grid(row=0, column=1, padx=10, pady=8)
        # Absence_Name_Var = Absence_Name.get()

        # Time Field
        self.absence_frame_label = customtkinter.CTkLabel(
            self.absence_frame,
            text="Time",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.absence_frame_label.grid(row=1, column=0, padx=10, pady=8)
        self.Absence_Time = customtkinter.CTkEntry(
            self.absence_frame,
            placeholder_text="Time",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.Absence_Time.grid(row=1, column=1, padx=10, pady=8)

        # Date Field
        self.absence_frame_label = customtkinter.CTkLabel(
            self.absence_frame,
            text="Date",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.absence_frame_label.grid(row=2, column=0, padx=10, pady=8)
        self.Absence_Date = customtkinter.CTkEntry(
            self.absence_frame,
            placeholder_text="Date",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.Absence_Date.grid(row=2, column=1, padx=10, pady=8)

        # Returning Field
        self.absence_frame_label = customtkinter.CTkLabel(
            self.absence_frame,
            text="Returning? Time?",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.absence_frame_label.grid(row=3, column=0, padx=10, pady=8)
        self.Absence_Returning = customtkinter.CTkEntry(
            self.absence_frame,
            placeholder_text="Are they returning? When?",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.Absence_Returning.grid(row=3, column=1, padx=10, pady=8)

        # Reason Field
        self.absence_frame_label = customtkinter.CTkLabel(
            self.absence_frame,
            text="Reason *",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.absence_frame_label.grid(row=4, column=0, padx=10, pady=8)
        self.Absence_Reason = customtkinter.CTkEntry(
            self.absence_frame,
            placeholder_text="Why did they leave?",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.Absence_Reason.grid(row=4, column=1, padx=10, pady=8)

        # Collected Field
        self.absence_frame_label = customtkinter.CTkLabel(
            self.absence_frame,
            text="Collected By",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.absence_frame_label.grid(row=5, column=0, padx=10, pady=8)
        self.Absence_Collector = customtkinter.CTkEntry(
            self.absence_frame,
            placeholder_text="Who has collected the student?",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.Absence_Collector.grid(row=5, column=1, padx=10, pady=8)

        # Buttons (Seperate frame)
        self.cancel = customtkinter.CTkButton(
            self.buttons_frame,
            command=self.Cancel_button_event,
            text="Cancel",
            fg_color="Red",
            hover_color="Dark Red",
        )
        self.cancel.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        self.submit = customtkinter.CTkButton(
            self.buttons_frame,
            command=self.Submit_button_event,
            text="Finish",
            fg_color="Green",
            hover_color="Dark Green",
        )
        self.submit.grid(row=0, column=3, padx=20, pady=20, sticky="ew")

        # Cancel and submit buttons

    def Cancel_button_event(self):
        print("Cancel Attendance Submision")
        self.destroy()

    def Submit_button_event(self):
        print("Submit Attendance")
        functions.attendance_update(
            self.Absence_Name.get(),
            self.Absence_Date.get(),
            self.Absence_Time.get(),
            self.Absence_Returning.get(),
            self.Absence_Reason.get(),
            self.Absence_Collector.get(),
        )
        self.destroy()


class Student_ID(customtkinter.CTkToplevel):
    def __init__(self):
        super().__init__()
        self.geometry("375x125")
        self.lift()
        self.title("Student ID")

        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.student_ID_frame = customtkinter.CTkFrame(
            self, corner_radius=0, width=400, height=300
        )
        self.student_ID_frame.grid(row=0, column=0, sticky="nsew")
        self.student_ID_frame.grid_rowconfigure(
            1, weight=1
        )  # Change number of rows in the naviation frame

        self.buttons_frame = customtkinter.CTkFrame(
            self, corner_radius=0, width=400, height=100
        )
        self.buttons_frame.grid(row=1, column=0, sticky="nsew")
        self.buttons_frame.grid_rowconfigure(
            1, weight=1
        )  # Change number of rows in the naviation frame
        self.buttons_frame.grid_columnconfigure(3, weight=1)

        # Name Field
        self.student_ID_frame_label = customtkinter.CTkLabel(
            self.student_ID_frame,
            text="Name *",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.student_ID_frame_label.grid(row=0, column=0, padx=10, pady=8)
        self.Student_ID_Name = customtkinter.CTkEntry(
            self.student_ID_frame,
            placeholder_text="Name",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.Student_ID_Name.grid(row=0, column=1, padx=10, pady=8)

        # Buttons (Seperate frame)
        self.cancel = customtkinter.CTkButton(
            self.buttons_frame,
            command=self.Cancel_button_event,
            text="Cancel",
            fg_color="Red",
            hover_color="Dark Red",
        )
        self.cancel.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        self.submit = customtkinter.CTkButton(
            self.buttons_frame,
            command=self.Submit_button_event,
            text="Finish",
            fg_color="Green",
            hover_color="Dark Green",
        )
        self.submit.grid(row=0, column=3, padx=20, pady=20, sticky="ew")

        # Cancel and submit buttons

    def Cancel_button_event(self):
        print("Cancel Student ID Submission")
        self.destroy()

    def Submit_button_event(self):
        print("Submit Student ID Request")
        functions.student_ID(self.Student_ID_Name.get())
        self.destroy()


class Papercut_Deposit(customtkinter.CTkToplevel):
    def __init__(self):
        super().__init__()
        self.geometry("450x292")
        self.lift()
        self.title("Papercut Deposit")

        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.papercut_deposit_frame = customtkinter.CTkFrame(
            self, corner_radius=0, width=450, height=300
        )
        self.papercut_deposit_frame.grid(row=0, column=0, sticky="nsew")
        self.papercut_deposit_frame.grid_rowconfigure(
            6, weight=1
        )  # Change number of rows in the naviation frame

        self.buttons_frame = customtkinter.CTkFrame(
            self, corner_radius=0, width=400, height=100
        )
        self.buttons_frame.grid(row=1, column=0, sticky="nsew")
        self.buttons_frame.grid_rowconfigure(
            1, weight=1
        )  # Change number of rows in the naviation frame
        self.buttons_frame.grid_columnconfigure(3, weight=1)

        # Name Field
        self.papercut_deposit_frame_label = customtkinter.CTkLabel(
            self.papercut_deposit_frame,
            text="Name *",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
            width=110,
        )
        self.papercut_deposit_frame_label.grid(row=0, column=0, padx=10, pady=8)
        self.papercut_deposit_name = customtkinter.CTkEntry(
            self.papercut_deposit_frame,
            placeholder_text="Name",
            width=300,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.papercut_deposit_name.grid(row=0, column=1, padx=10, pady=8)

        # Amount Field
        self.papercut_deposit_frame_label = customtkinter.CTkLabel(
            self.papercut_deposit_frame,
            text="Amount *",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
            width=110,
        )
        self.papercut_deposit_frame_label.grid(row=1, column=0, padx=10, pady=8)
        self.papercut_deposit_amount = customtkinter.CTkEntry(
            self.papercut_deposit_frame,
            placeholder_text="Amount",
            width=300,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.papercut_deposit_amount.grid(row=1, column=1, padx=10, pady=8)

        # Method Field
        self.papercut_deposit_frame_label = customtkinter.CTkLabel(
            self.papercut_deposit_frame,
            text="Method *",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
            width=110,
        )
        self.papercut_deposit_frame_label.grid(row=2, column=0, padx=10, pady=8)
        self.papercut_deposit_method = customtkinter.CTkComboBox(
            self.papercut_deposit_frame,
            values=["Cash", "EFT", "Other"],
            width=300,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.papercut_deposit_method.grid(row=2, column=1, padx=10, pady=8)

        # Comment Field
        self.papercut_deposit_frame_label = customtkinter.CTkLabel(
            self.papercut_deposit_frame,
            text="Comment ",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
            width=110,
        )
        self.papercut_deposit_frame_label.grid(row=3, column=0, padx=10, pady=8)
        self.papercut_deposit_comment = customtkinter.CTkEntry(
            self.papercut_deposit_frame,
            placeholder_text="",
            width=300,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.papercut_deposit_comment.grid(row=3, column=1, padx=10, pady=8)

        # Buttons (Seperate frame)
        self.cancel = customtkinter.CTkButton(
            self.buttons_frame,
            command=self.Cancel_button_event,
            text="Cancel",
            fg_color="Red",
            hover_color="Dark Red",
        )
        self.cancel.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        self.submit = customtkinter.CTkButton(
            self.buttons_frame,
            command=self.Submit_button_event,
            text="Finish",
            fg_color="Green",
            hover_color="Dark Green",
        )
        self.submit.grid(row=0, column=3, padx=20, pady=20, sticky="ew")

        # Cancel and submit buttons

    def Cancel_button_event(self):
        print("Cancel Student ID Submission")
        self.destroy()

    def Submit_button_event(self):
        print("Submit Student ID Request")
        papercut_bot.deposit(
            self.papercut_deposit_name.get(),
            self.papercut_deposit_amount.get(),
            self.papercut_deposit_method.get(),
            self.papercut_deposit_comment.get(),
        )
        self.destroy()


class wwcc_check(customtkinter.CTkToplevel):
    def __init__(self):
        super().__init__()
        self.geometry("390x180")
        self.lift()
        self.title("Working with Children Check Search")

        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.wwcc_check_frame = customtkinter.CTkFrame(
            self, corner_radius=0, width=400, height=300
        )
        self.wwcc_check_frame.grid(row=0, column=0, sticky="nsew")
        self.wwcc_check_frame.grid_rowconfigure(
            1, weight=1
        )  # Change number of rows in the naviation frame

        self.buttons_frame = customtkinter.CTkFrame(
            self, corner_radius=0, width=400, height=100
        )
        self.buttons_frame.grid(row=1, column=0, sticky="nsew")
        self.buttons_frame.grid_rowconfigure(
            1, weight=1
        )  # Change number of rows in the naviation frame
        self.buttons_frame.grid_columnconfigure(3, weight=1)

        # Name Field
        self.wwcc_check_frame_label = customtkinter.CTkLabel(
            self.wwcc_check_frame,
            text="Name *",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.wwcc_check_frame_label.grid(row=0, column=0, padx=10, pady=8)
        self.wwcc_check_name = customtkinter.CTkEntry(
            self.wwcc_check_frame,
            placeholder_text="Name",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.wwcc_check_name.grid(row=0, column=1, padx=10, pady=8)

        # Name Field
        self.wwcc_check_frame_label = customtkinter.CTkLabel(
            self.wwcc_check_frame,
            text="Number *",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.wwcc_check_frame_label.grid(row=1, column=0, padx=10, pady=8)
        self.wwcc_check_num = customtkinter.CTkEntry(
            self.wwcc_check_frame,
            placeholder_text="00000000",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.wwcc_check_num.grid(row=1, column=1, padx=10, pady=8)

        # Buttons (Seperate frame)
        self.cancel = customtkinter.CTkButton(
            self.buttons_frame,
            command=self.Cancel_button_event,
            text="Close",
            fg_color="Red",
            hover_color="Dark Red",
        )
        self.cancel.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        self.submit = customtkinter.CTkButton(
            self.buttons_frame,
            command=self.Submit_button_event,
            text="Finish",
            fg_color="Green",
            hover_color="Dark Green",
        )
        self.submit.grid(row=0, column=3, padx=20, pady=20, sticky="ew")

        # Cancel and submit buttons

    def Cancel_button_event(self):
        print("Cancel WWCC Search Request")
        self.destroy()

    def Submit_button_event(self):
        print("Submit WWCC Search Request")
        print(
            functions.wwcc_check(self.wwcc_check_num.get(), self.wwcc_check_name.get())
        )
        self.destroy()


class Centerpay(customtkinter.CTkToplevel):
    def __init__(self):
        super().__init__()
        self.geometry("430x290")
        self.title("Centerpay")
        self.lift()

        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.centerpay_frame = customtkinter.CTkFrame(
            self, corner_radius=0, width=400, height=300
        )
        self.centerpay_frame.grid(row=0, column=0, sticky="nsew")
        self.centerpay_frame.grid_rowconfigure(
            5, weight=1
        )  # Change number of rows in the frame

        self.buttons_frame = customtkinter.CTkFrame(
            self, corner_radius=0, width=400, height=100
        )
        self.buttons_frame.grid(row=1, column=0, sticky="nsew")
        self.buttons_frame.grid_rowconfigure(
            1, weight=1
        )  # Change number of rows in the frame
        self.buttons_frame.grid_columnconfigure(3, weight=1)

        # Student Code Field
        self.centerpay_frame_label = customtkinter.CTkLabel(
            self.centerpay_frame,
            text="Student Code",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.centerpay_frame_label.grid(row=0, column=0, padx=10, pady=8)
        self.Centerpay_Student_Code = customtkinter.CTkEntry(
            self.centerpay_frame,
            placeholder_text="eg. WAR0076",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.Centerpay_Student_Code.grid(row=0, column=1, padx=10, pady=8)

        # Payment Date Field
        self.centerpay_frame_label = customtkinter.CTkLabel(
            self.centerpay_frame,
            text="Payment Date",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.centerpay_frame_label.grid(row=1, column=0, padx=10, pady=8)
        self.Centerpay_Date = customtkinter.CTkEntry(
            self.centerpay_frame,
            placeholder_text="DD/MM/YY",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.Centerpay_Date.grid(row=1, column=1, padx=10, pady=8)

        # Payment Total Field
        self.centerpay_frame_label = customtkinter.CTkLabel(
            self.centerpay_frame,
            text="Payment Total",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.centerpay_frame_label.grid(row=2, column=0, padx=10, pady=8)
        self.Centerpay_Payment_Total = customtkinter.CTkEntry(
            self.centerpay_frame,
            placeholder_text="20",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.Centerpay_Payment_Total.grid(row=2, column=1, padx=10, pady=8)

        # Fee Total Field
        self.centerpay_frame_label = customtkinter.CTkLabel(
            self.centerpay_frame,
            text="Fee Total",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.centerpay_frame_label.grid(row=3, column=0, padx=10, pady=8)
        self.Centerpay_Fee_Total = customtkinter.CTkEntry(
            self.centerpay_frame,
            placeholder_text=".99",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.Centerpay_Fee_Total.grid(row=3, column=1, padx=10, pady=8)

        # Buttons (Seperate frame)
        self.cancel = customtkinter.CTkButton(
            self.buttons_frame,
            command=self.Cancel_button_event,
            text="Cancel",
            fg_color="Red",
            hover_color="Dark Red",
        )
        self.cancel.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        self.submit = customtkinter.CTkButton(
            self.buttons_frame,
            command=self.Submit_button_event,
            text="Finish",
            fg_color="Green",
            hover_color="Dark Green",
        )
        self.submit.grid(row=0, column=3, padx=20, pady=20, sticky="ew")

        # Cancel and submit buttons

    def Cancel_button_event(self):
        print("Cancel Centerpay")
        self.destroy()

    def Submit_button_event(self):
        print("Submit Centerpay")
        functions.Centerpay(
            self.Centerpay_Student_Code.get(),
            self.Centerpay_Date.get(),
            self.Centerpay_Payment_Total.get(),
            self.Centerpay_Fee_Total.get(),
        )
        self.destroy()


class BPAY:
    def __init__(self):
        super().__init__()
        functions.BPAY()


class Canteen(customtkinter.CTkToplevel):
    def __init__(self):
        super().__init__()
        self.geometry("415x290")
        self.title("Canteen")
        self.lift()

        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.canteen_frame = customtkinter.CTkFrame(
            self, corner_radius=0, width=400, height=300
        )
        self.canteen_frame.grid(row=0, column=0, sticky="nsew")
        self.canteen_frame.grid_rowconfigure(
            5, weight=1
        )  # Change number of rows in the frame

        self.buttons_frame = customtkinter.CTkFrame(
            self, corner_radius=0, width=400, height=100
        )
        self.buttons_frame.grid(row=1, column=0, sticky="nsew")
        self.buttons_frame.grid_rowconfigure(
            1, weight=1
        )  # Change number of rows in the frame
        self.buttons_frame.grid_columnconfigure(3, weight=1)

        # Cash Total Field
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_frame,
            text="Cash Total",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.canteen_frame_label.grid(row=0, column=0, padx=10, pady=8)
        self.Canteen_Cash_Total = customtkinter.CTkEntry(
            self.canteen_frame,
            placeholder_text="",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.Canteen_Cash_Total.grid(row=0, column=1, padx=10, pady=8)

        # EFT 1 Field Field
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_frame,
            text="EFT 1 Total",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.canteen_frame_label.grid(row=1, column=0, padx=10, pady=8)
        self.Canteen_EFT1_Total = customtkinter.CTkEntry(
            self.canteen_frame,
            placeholder_text="",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.Canteen_EFT1_Total.grid(row=1, column=1, padx=10, pady=8)

        # EFT 2 Field Field
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_frame,
            text="EFT 2 Total",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.canteen_frame_label.grid(row=2, column=0, padx=10, pady=8)
        self.Canteen_EFT2_Total = customtkinter.CTkEntry(
            self.canteen_frame,
            placeholder_text="",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.Canteen_EFT2_Total.grid(row=2, column=1, padx=10, pady=8)

        # Receipt Date Field
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_frame,
            text="Receipt Date",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.canteen_frame_label.grid(row=3, column=0, padx=10, pady=8)
        self.Canteen_Receipt_Date = customtkinter.CTkEntry(
            self.canteen_frame,
            placeholder_text="DD/MM/YY",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.Canteen_Receipt_Date.grid(row=3, column=1, padx=10, pady=8)

        # Buttons (Seperate frame)
        self.cancel = customtkinter.CTkButton(
            self.buttons_frame,
            command=self.Cancel_button_event,
            text="Cancel",
            fg_color="Red",
            hover_color="Dark Red",
        )
        self.cancel.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        self.submit = customtkinter.CTkButton(
            self.buttons_frame,
            command=self.Submit_button_event,
            text="Finish",
            fg_color="Green",
            hover_color="Dark Green",
        )
        self.submit.grid(row=0, column=3, padx=20, pady=20, sticky="ew")

        # Cancel and submit buttons

    def Cancel_button_event(self):
        print("Cancel Canteen")
        self.destroy()

    def Submit_button_event(self):
        print("Submit Canteen")
        functions.Canteen(
            self.Canteen_Cash_Total.get(),
            self.Canteen_EFT1_Total.get(),
            self.Canteen_EFT2_Total.get(),
            self.Canteen_Receipt_Date.get(),
        )
        self.destroy()


class QKR_Canteen(customtkinter.CTkToplevel):
    def __init__(self):
        super().__init__()
        self.geometry("415x180")
        self.title("QKR Canteen")
        self.lift()

        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.qkr_canteen_frame = customtkinter.CTkFrame(
            self, corner_radius=0, width=400, height=300
        )
        self.qkr_canteen_frame.grid(row=0, column=0, sticky="nsew")
        self.qkr_canteen_frame.grid_rowconfigure(
            2, weight=1
        )  # Change number of rows in the frame

        self.buttons_frame = customtkinter.CTkFrame(
            self, corner_radius=0, width=400, height=100
        )
        self.buttons_frame.grid(row=1, column=0, sticky="nsew")
        self.buttons_frame.grid_rowconfigure(
            1, weight=1
        )  # Change number of rows in the frame
        self.buttons_frame.grid_columnconfigure(3, weight=1)

        # Receipt Total Field
        self.qkr_canteen_frame_label = customtkinter.CTkLabel(
            self.qkr_canteen_frame,
            text="Receipt Total",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.qkr_canteen_frame_label.grid(row=0, column=0, padx=10, pady=8)
        self.QKR_Canteen_Total = customtkinter.CTkEntry(
            self.qkr_canteen_frame,
            placeholder_text="",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.QKR_Canteen_Total.grid(row=0, column=1, padx=10, pady=8)

        # Receipt Date Field
        self.qkr_canteen_frame_label = customtkinter.CTkLabel(
            self.qkr_canteen_frame,
            text="Receipt Date",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.qkr_canteen_frame_label.grid(row=1, column=0, padx=10, pady=8)
        self.Receipt_Date = customtkinter.CTkEntry(
            self.qkr_canteen_frame,
            placeholder_text="DD/MM/YY",
            width=280,
            height=40,
            border_width=1,
            corner_radius=10,
        )
        self.Receipt_Date.grid(row=1, column=1, padx=10, pady=8)

        # Buttons (Seperate frame)
        self.cancel = customtkinter.CTkButton(
            self.buttons_frame,
            command=self.Cancel_button_event,
            text="Cancel",
            fg_color="Red",
            hover_color="Dark Red",
        )
        self.cancel.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        self.submit = customtkinter.CTkButton(
            self.buttons_frame,
            command=self.Submit_button_event,
            text="Finish",
            fg_color="Green",
            hover_color="Dark Green",
        )
        self.submit.grid(row=0, column=3, padx=20, pady=20, sticky="ew")

        # Cancel and submit buttons

    def Cancel_button_event(self):
        print("Cancel QKR Canteen")
        self.destroy()

    def Submit_button_event(self):
        print("Submit Canteen")
        functions.QKR_Canteen(self.QKR_Canteen_Total.get(), self.Receipt_Date.get())
        self.destroy()


class CSEF:
    def __init__(self):
        super().__init__()
        functions.CSEF()


class QKR_Update:
    def __init__(self):
        super().__init__()
        qkr_bot.main()


class Vehicle_Expense:
    def __init__(self):
        super().__init__()
        functions.Vehigle_GL()


if __name__ == "__main__":
    print("Initialised")
    while True:
        time.sleep(0.1)
        if keyboard.is_pressed("alt+x"):
            GUI().mainloop()
