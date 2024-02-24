import customtkinter


class Canteen_Overview(
    customtkinter.CTk
):  # This isnt finished. Still needs lots of work and possibly needs to be moved
    def __init__(
        self,
        cash_total,
        cash_gl,
        eft1_total,
        eft1_gl,
        eft2_total,
        eft2_gl,
        receipt_date,
    ):
        super().__init__()
        self.geometry("620x280")
        self.title("Canteen Overview")
        self.lift()

        total = float(cash_total) + float(eft1_total) + float(eft2_total)

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.canteen_overview_frame = customtkinter.CTkFrame(
            self, corner_radius=0, width=400, height=300
        )
        self.canteen_overview_frame.grid(row=0, column=0, sticky="nsew")
        self.canteen_overview_frame.grid_rowconfigure(
            5, weight=1
        )  # Change number of rows in the frame
        self.canteen_overview_frame.grid_columnconfigure(
            6, weight=1
        )  # Change number of rows in the frame

        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text="Canteen Overview",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.canteen_frame_label.grid(row=0, column=0, padx=20, pady=10, columnspan=6)

        # Cash
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text="Cash",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.canteen_frame_label.grid(row=1, column=0, padx=10, pady=8, columnspan=2)

        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text="Cash Total:",
            compound="left",
            font=customtkinter.CTkFont(size=15),
        )
        self.canteen_frame_label.grid(row=2, column=0, padx=10, pady=3)
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text="$" + str(cash_total),
            compound="left",
            font=customtkinter.CTkFont(size=15),
        )
        self.canteen_frame_label.grid(row=2, column=1, padx=10, pady=3)
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text="Receipt:",
            compound="left",
            font=customtkinter.CTkFont(size=15),
        )
        self.canteen_frame_label.grid(row=3, column=0, padx=10, pady=3)
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text=cash_gl,
            compound="left",
            font=customtkinter.CTkFont(size=15),
        )
        self.canteen_frame_label.grid(row=3, column=1, padx=10, pady=3)

        # EFT 1
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text="EFT 1",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.canteen_frame_label.grid(row=1, column=2, padx=10, pady=8, columnspan=2)

        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text="EFT1 Total:",
            compound="left",
            font=customtkinter.CTkFont(size=15),
        )
        self.canteen_frame_label.grid(row=2, column=2, padx=10, pady=3)
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text="$" + str(eft1_total),
            compound="left",
            font=customtkinter.CTkFont(size=15),
        )
        self.canteen_frame_label.grid(row=2, column=3, padx=10, pady=3)
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text="Receipt:",
            compound="left",
            font=customtkinter.CTkFont(size=15),
        )
        self.canteen_frame_label.grid(row=3, column=2, padx=10, pady=3)
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text=eft1_gl,
            compound="left",
            font=customtkinter.CTkFont(size=15),
        )
        self.canteen_frame_label.grid(row=3, column=3, padx=10, pady=3)

        # EFT 2
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text="EFT 2",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.canteen_frame_label.grid(row=1, column=4, padx=10, pady=8, columnspan=2)

        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text="EFT2 Total:",
            compound="left",
            font=customtkinter.CTkFont(size=15),
        )
        self.canteen_frame_label.grid(row=2, column=4, padx=10, pady=3)
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text="$" + str(eft2_total),
            compound="left",
            font=customtkinter.CTkFont(size=15),
        )
        self.canteen_frame_label.grid(row=2, column=5, padx=10, pady=3)
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text="Receipt:",
            compound="left",
            font=customtkinter.CTkFont(size=15),
        )
        self.canteen_frame_label.grid(row=3, column=4, padx=10, pady=3)
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text=eft2_gl,
            compound="left",
            font=customtkinter.CTkFont(size=15),
        )
        self.canteen_frame_label.grid(row=3, column=5, padx=10, pady=3)

        # Totals

        # Receipt Date Field
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text="Total:",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.canteen_frame_label.grid(row=8, column=1, padx=0, pady=0, columnspan=2)
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text="Date:",
            compound="left",
            font=customtkinter.CTkFont(size=15, weight="bold"),
        )
        self.canteen_frame_label.grid(row=8, column=3, padx=0, pady=0, columnspan=2)

        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text="$" + str(total),
            compound="left",
            font=customtkinter.CTkFont(size=15),
        )
        self.canteen_frame_label.grid(row=9, column=1, padx=0, pady=0, columnspan=2)
        self.canteen_frame_label = customtkinter.CTkLabel(
            self.canteen_overview_frame,
            text=receipt_date,
            compound="left",
            font=customtkinter.CTkFont(size=15),
        )
        self.canteen_frame_label.grid(row=9, column=3, padx=0, pady=0, columnspan=2)

        self.buttons_frame = customtkinter.CTkFrame(
            self, corner_radius=0, width=400, height=100
        )
        self.buttons_frame.grid(row=1, column=0, sticky="nsew")
        self.buttons_frame.grid_rowconfigure(
            1, weight=1
        )  # Change number of rows in the frame
        self.buttons_frame.grid_columnconfigure(1, weight=1)
        # Buttons (Seperate frame)
        self.submit = customtkinter.CTkButton(
            self.canteen_overview_frame,
            command=self.Done_button_event,
            text="Finish",
            fg_color="Green",
            hover_color="Dark Green",
        )
        self.submit.grid(row=10, column=2, padx=20, pady=20, sticky="ew", columnspan=2)

    def Done_button_event(self):
        print("Close Canteen Overview")
        self.destroy()


# Canteen_Overview(20,'GL00453456',20,'GL00453457',20,'GL00453458','22/3/23').mainloop()
