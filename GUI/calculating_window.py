import tkinter as tk
from tkinter import messagebox
from tkinter.ttk import Treeview


GREEN = "#1CAC78"
BLUE = "#6495ED"
ORANGE = "#FF8243"
GREY = "#434B4D"
WHITE = "#FFFAFA"
PINGY_WHITE = "#FFF0F5"
MAIN_WINDOW_EDGE_INTERVAL = 28
MAIN_WINDOW_IPADX = 5
MAIN_WINDOW_IPADY = 5
FONT_STYLE = "Calibri 10"
VALUES_FOR_HEADINGS_TABLE = ("Year", "Week", "Month", "Day")
TABLE_COLUMNS_WIDTH = 110


class CalculatingWindow:

    def __init__(self,):
        self.main_window_link = None  # Attach for link a main window in future in init a main_window object

        # GUI1 init

        self.calculating_window = tk.Tk()
        self.calculating_window.geometry('520x365')
        self.calculating_window.resizable(False, False)
        self.calculating_window.title("Расчёт")
        self.calculating_window["bg"] = GREY

        # Target file button

        self.target_file_button = tk.Button(
            self.calculating_window,
            text="Выбрать целевой файл",
            font=FONT_STYLE,
            command=lambda: print(1),
            width=20
        )
        self.target_file_button.grid(
            column=0,
            row=0,
            pady=(MAIN_WINDOW_EDGE_INTERVAL - 9, 15),
            padx=(MAIN_WINDOW_EDGE_INTERVAL - 18, 10)

        )

        # Target file path text field

        self.target_file_path_field = tk.Entry(
            self.calculating_window,
            font=FONT_STYLE,
            width=45,
            state="readonly",
        )
        self.target_file_path_field.grid(
            column=1,
            row=0,
            ipadx=10,
            ipady=2,
            pady=(MAIN_WINDOW_EDGE_INTERVAL - 9, 15),
            padx=(MAIN_WINDOW_EDGE_INTERVAL - 28, 90),
        )

        # Source file button

        self.source_file_button = tk.Button(
            self.calculating_window,
            text="Выбрать файл",
            font=FONT_STYLE,
            command=lambda: print(1),
            width=20
        )
        self.source_file_button.grid(
            column=0,
            row=1,
            pady=(0, 5),
            padx=(MAIN_WINDOW_EDGE_INTERVAL - 18, 10)
        )

        # Source file path text field

        self.source_file_path_field = tk.Entry(
            self.calculating_window,
            font=FONT_STYLE,
            width=45,
            state="readonly"
        )
        self.source_file_path_field.grid(
            column=1,
            row=1,
            ipadx=10,
            ipady=2,
            pady=(0, 5),
            padx=(MAIN_WINDOW_EDGE_INTERVAL - 28, 90),

        )

        # Operation table

        self.operation_table = Treeview(
            self.calculating_window,
            columns=VALUES_FOR_HEADINGS_TABLE,
            show='headings',
            selectmode="extended",
        )

        self.operation_table.heading("Year", text="Год")
        self.operation_table.column("Year", width=TABLE_COLUMNS_WIDTH)
        self.operation_table.heading("Month", text="Месяц")
        self.operation_table.column("Month", width=TABLE_COLUMNS_WIDTH)
        self.operation_table.heading("Week", text="Неделя")
        self.operation_table.column("Week", width=TABLE_COLUMNS_WIDTH)
        self.operation_table.heading("Day", text="День")
        self.operation_table.column("Day", width=TABLE_COLUMNS_WIDTH)

        self.operation_table.grid(
            sticky="w",
            column=0,
            row=2,
            columnspan=2,
            ipady=0,
            ipadx=29,
            padx=(MAIN_WINDOW_EDGE_INTERVAL - 18, 5),
            pady=0,
        )

        # Back button

        self.back_button = tk.Button(
            self.calculating_window,
            text="Главное меню",
            font=FONT_STYLE,
            command=lambda: self.show_main_window(),
            width=13
        )
        self.back_button.grid(
            column=1,
            row=3,
            pady=(7, 15),
            padx=(120, 0),
            sticky="w"

        )

        # Execute button

        self.execute_button = tk.Button(
            self.calculating_window,
            text="Посчитать",
            font=FONT_STYLE,
            command=lambda: print(1),
            width=13,
            background=GREEN,
        )
        self.execute_button.grid(
            column=1,
            row=3,
            pady=(7, 15),
            padx=(MAIN_WINDOW_EDGE_INTERVAL - 18, 90),
            sticky="e"

        )

        # method for closing handling
        def on_closing():
            if messagebox.askokcancel("Выход", "Закрыть программу? Данные не сохранятся."):
                self.main_window_link.main_window.destroy()
                self.calculating_window.destroy()

        self.calculating_window.protocol("WM_DELETE_WINDOW", on_closing)


    def show_main_window(self):
        self.calculating_window.withdraw()
        self.main_window_link.main_window.deiconify()
