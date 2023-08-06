import tkinter as tk
from datetime import datetime
from tkinter.ttk import Combobox, Treeview
from

GREEN = "#1CAC78"
BLUE = "#6495ED"
ORANGE = "#FF8243"
GREY = "#434B4D"
WHITE = "#FFFAFA"
PINGY_WHITE = "#FFF0F5"
MAIN_WINDOW_EDGE_INTERVAL = 18
MAIN_WINDOW_IPADX = 5
MAIN_WINDOW_IPADY = 5
CURRENT_YEAR = int(datetime.today().strftime("%Y"))
FONT_STYLE = "Calibri 12"
VALUES_FOR_CATEGORY_TABLE = (("Упаковка",), ("Плёнка",), ("Сырьё",))


class MainWindow:

    def __init__(self,
                 calculating_window_link):
                 # calculate_fact_intake,
                 # calculate_intake_corrug_by_demand_plan,
                 # calculate_intake_corrug_by_detail_plan,
                 # calculate_intake_film_raw_by_plan_from_sap):
        # Methods init

        self.calculating_window_link = calculating_window_link  # link to calculate window
        # self.calculate_fact_intake = calculate_fact_intake
        # self.calculate_intake_corrug_by_demand_plan = calculate_intake_corrug_by_demand_plan
        # self.calculate_intake_corrug_by_detail_plan = calculate_intake_corrug_by_detail_plan
        # self.calculate_intake_film_raw_by_plan_from_sap = calculate_intake_film_raw_by_plan_from_sap

        # GUI1 init

        self.main_window = tk.Tk()
        self.main_window.geometry('400x250')
        self.main_window.resizable(False, False)
        self.main_window.title("Выбор категории")
        self.main_window["bg"] = GREY

        # Category table

        self.category_table = Treeview(
            self.main_window,
            columns="Category",
            show='headings',
            selectmode="browse"
        )

        def change_category(event):
            selected_category = self.category_table.item(self.category_table.selection()[0])["values"][0]
            if selected_category == "Упаковка":
                self.change_category_menu("Расчёт упаковки", "По плану производства", "По прогнозу",
                                          "Фактические остатки", lambda: print(1), lambda: print(2), lambda: print(3))
            elif selected_category == "Сырьё":
                self.change_category_menu("Расчёт сырья", "По деманду", "Фактические остатки",
                                          None, lambda: print(1), lambda: print(2), None)
            elif selected_category == "Плёнка":
                self.change_category_menu("Расчёт плёнки", "По деманду", "Фактические остатки",
                                          None, lambda: print(1), lambda: print(2), None)

        self.category_table.heading('# 1', text="Категория")
        self.category_table.column('# 1', width=80)
        self.category_table.bind("<<TreeviewSelect>>", change_category)

        for category in VALUES_FOR_CATEGORY_TABLE:
            self.category_table.insert("", "end", values=tuple(category))

        self.category_table.pack(
            side="left",
            ipady=12,
        )

        # Label for year combobox

        self.year_text = tk.Label(
            self.main_window,
            text="Год:",
            background=GREY,
            foreground=WHITE,
            font="Calibri 10"
        )
        self.year_text.pack(
            anchor="e",
            pady=1,
            padx=(0.00001, 0.0001),
            ipadx=0.001,
            ipady=0.1,
            expand=False,
            side="top"
        )

        # Year combobox

        self.selected_year = tk.StringVar()
        year_box = Combobox(
            self.main_window,
            textvariable=self.selected_year,
            width=4,
        )
        year_box['values'] = (CURRENT_YEAR - 1, CURRENT_YEAR, CURRENT_YEAR + 1)
        year_box.current(1)
        year_box.pack(
            pady=1,
            ipadx=0.0001,
            ipady=0.1,
            expand=False,
            side="right",
            anchor='n'
        )

        # Up button

        self.up_menu_btn = tk.Button(
            self.main_window,
            text="Упаковка",
            width=MAIN_WINDOW_EDGE_INTERVAL,
            background=PINGY_WHITE,
            font=FONT_STYLE,
        )

        # Middle button

        self.mid_menu_btn = tk.Button(
            self.main_window,
            text="Сырьё",
            width=MAIN_WINDOW_EDGE_INTERVAL,
            background=PINGY_WHITE,
            font=FONT_STYLE
        )

        # Film button

        self.down_menu_btn = tk.Button(
            self.main_window,
            text="Плёнка",
            width=MAIN_WINDOW_EDGE_INTERVAL,
            background=PINGY_WHITE,
            font=FONT_STYLE
        )

        # Handling links between main window and calculating window

        self.calculating_window_link.calculating_window.withdraw()
        self.calculating_window_link.main_window_link = self

        # method for closing handling
        def on_closing():
            self.main_window.destroy()
            self.calculating_window_link.calculating_window.destroy()

        self.main_window.protocol("WM_DELETE_WINDOW", on_closing)

        self.main_window.mainloop()

    def change_category_menu(self, menu_tile, up_button_text, mid_button_text,
                             down_button_text,
                             up_button_command, mid_button_command, down_button_command):

        self.main_window.title(menu_tile)

        # Up button

        self.up_menu_btn.pack(
            ipadx=MAIN_WINDOW_IPADX,
            ipady=MAIN_WINDOW_IPADY,
            padx=35,
            pady=10,
            expand=True,
            side="top",
            anchor='nw',
            fill="x"
        )
        self.up_menu_btn["text"] = up_button_text
        self.up_menu_btn["command"] = lambda: self.show_calculating_window()

        # Middle button

        self.mid_menu_btn.pack(
            ipadx=MAIN_WINDOW_IPADX,
            ipady=MAIN_WINDOW_IPADY,
            padx=35,
            pady=10,
            expand=True,
            side="top",
            anchor='nw',
            fill="x"
        )
        self.mid_menu_btn["text"] = mid_button_text
        self.mid_menu_btn["command"] = mid_button_command

        # Down button

        if down_button_text is None:
            self.down_menu_btn.pack_forget()
        else:
            self.down_menu_btn.pack(
                ipadx=MAIN_WINDOW_IPADX,
                ipady=MAIN_WINDOW_IPADY,
                padx=35,
                pady=10,
                expand=True,
                side="top",
                anchor='nw',
                fill="x"
            )
            self.down_menu_btn["text"] = down_button_text
            self.down_menu_btn["command"] = down_button_command

    def show_calculating_window(self):
        self.main_window.withdraw()
        self.calculating_window_link.calculating_window.deiconify()
