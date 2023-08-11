import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter.ttk import Treeview, Scrollbar, Progressbar, Label
from constants import *

FONT_STYLE = "Calibri 10"
MAIN_WINDOW_EDGE_INTERVAL = 28


class CalculatingWindow:

    def __init__(self, ):
        self.main_window_link = None  # Attach for link a main window in future in init a main_window object

        # GUI init

        self.current_window_mode = -1
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
            command=lambda: self.select_target_file(),
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
            text="Выбрать источник",
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

        self.operation_table.heading("Column", text="ИД", )
        self.operation_table.column("Column", width=TABLE_COLUMNS_WIDTH, anchor="center")
        self.operation_table.heading("Year", text="Год")
        self.operation_table.column("Year", width=TABLE_COLUMNS_WIDTH, anchor="center")
        self.operation_table.heading("Month", text="Месяц")
        self.operation_table.column("Month", width=TABLE_COLUMNS_WIDTH, anchor="center")
        self.operation_table.heading("Week", text="Неделя")
        self.operation_table.column("Week", width=TABLE_COLUMNS_WIDTH, anchor="center")
        self.operation_table.heading("Day", text="День")
        self.operation_table.column("Day", width=TABLE_COLUMNS_WIDTH, anchor="center")

        # set tag

        self.operation_table.tag_configure('gray', background='#cccccc')
        self.operation_table.tag_configure('white', background='#ffffff')

        # Hide the Column column

        self.operation_table["displaycolumns"] = ("Year", "Week", "Month", "Day")

        self.operation_table.grid(
            sticky="w",
            column=0,
            row=2,
            columnspan=2,
            ipady=0,
            ipadx=29,
            padx=(MAIN_WINDOW_EDGE_INTERVAL - 18, 5),
            pady=0
        )

        # Scrollbar set

        self.vsb = Scrollbar(
            self.calculating_window,
            orient="vertical",
            command=self.operation_table.yview
        )
        self.vsb.grid(column=1,
                      row=2,
                      sticky='nse',
                      padx=(0, 90),
                      pady=(1, 1)
                      )

        self.operation_table.configure(yscrollcommand=self.vsb.set)

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

        # Progress bar

        self.progressbar = Progressbar(
            self.calculating_window,
            orient='horizontal',
            mode='determinate',
            value=0,
            maximum=100,
            length=110
        )
        self.start_progressbar = lambda: self.progressbar.grid(
            column=0,
            row=3,
            pady=(7, 15),
            padx=MAIN_WINDOW_EDGE_INTERVAL - 18,
            sticky='w',
            ipadx=20,
        )

        # Label percent of progress bar

        self.label_pb = Label(
            self.calculating_window,
            font="Calibri 12",
            foreground=WHITE,
            background=GREY,
            text=f"{self.progressbar['value']}%"
        )
        self.show_label_pb = lambda: self.label_pb.grid(
            column=1,
            row=3,
            columnspan=1,
            pady=(7, 15),
            padx=0,
            sticky='w',
            ipadx=0
        )

        # method for closing handling
        def on_closing():
            if messagebox.askokcancel("Выход", "Закрыть программу?"):
                self.main_window_link.main_window.destroy()
                self.calculating_window.destroy()

        self.calculating_window.protocol("WM_DELETE_WINDOW", on_closing)

    def show_main_window(self):
        self.calculating_window.withdraw()
        self.main_window_link.main_window.deiconify()

    def select_target_file(self):
        self.target_file_path_field.config(state="normal")
        self.target_file_path_field.delete(0, "end")
        self.target_file_path_field.insert(0, filedialog.askopenfilename(
            initialdir="..",
            title="Выберите целевой файл",
            filetypes=EXCEL_FILE_EXTENSIONS,
        )
                                           )
        self.target_file_path_field.config(state="readonly")

    def empty_window(self):
        self.source_file_path_field.config(state="normal")
        self.target_file_path_field.config(state="normal")
        self.target_file_path_field.delete(0, "end")
        self.source_file_path_field.delete(0, "end")
        self.source_file_path_field.config(state="readonly")
        self.target_file_path_field.config(state="readonly")
        for i in self.operation_table.get_children():
            self.operation_table.delete(i)

    def update_pb(self, value):
        if value == 0:
            percent = value
        else:
            percent = self.progressbar['value'] + value
        if percent > 100:
            percent = 100
        self.progressbar['value'] = percent
        self.label_pb['text'] = f"{self.progressbar['value']}%"

    def disable_all_elements(self):
        self.operation_table.configure(selectmode="none")
        self.target_file_button.config(state="disabled")
        self.source_file_button.config(state="disabled")
        self.back_button.config(state="disabled")
        self.execute_button.config(state="disabled")

    def enable_all_elements(self):
        self.operation_table.configure(selectmode="browse")
        self.target_file_button.config(state="normal")
        self.source_file_button.config(state="normal")
        self.back_button.config(state="normal")
        self.execute_button.config(state="normal")

