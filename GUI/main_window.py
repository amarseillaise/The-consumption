import inspect
import queue
from threading import Thread
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox, Treeview
import _queue
from data_io import calculate_and_echo_to_target_file
from constants import *
from exceptions import *


class MainWindow:

    def __init__(self, calculating_window_link, get_source_and_target_data):
        # Methods init

        self.calculating_window_link = calculating_window_link  # link to calculate window
        self.get_source_and_target_data = get_source_and_target_data

        # GUI init

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
                                          "Фактические остатки",
                                          lambda: self.show_calculating_window(0, clear_table=True), lambda: print(2),
                                          lambda: print(3))
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

        self.selected_year = tk.StringVar(self.main_window)
        self.year_box = Combobox(
            self.main_window,
            width=4,
            textvariable=self.selected_year
        )
        self.year_box['values'] = (CURRENT_YEAR - 1, CURRENT_YEAR, CURRENT_YEAR + 1)
        self.year_box.current(1)
        self.year_box.pack(
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

        self.main_window.title(MODE.get(menu_tile))

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
        self.up_menu_btn["command"] = up_button_command

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

    def show_calculating_window(self, mode, manual_path_to_source_file="", clear_table=True):

        # first clean the table if it almost has info
        if self.calculating_window_link.current_window_mode != mode:
            self.calculating_window_link.empty_window()
        if clear_table:
            for i in self.calculating_window_link.operation_table.get_children():
                self.calculating_window_link.operation_table.delete(i)

        #  getting an info from source file and handling an exceptions

        try:
            data = self.get_source_and_target_data(mode, manual_path_to_source_file)
        except TodayNotFindInSourceFileException:
            messagebox.showwarning("Внимание!", "В файле-источнике не найден текущий день. Возможно вы используете "
                                                "старый файл или в нём съехали таблицы.")
            return

        except UnableToFindExcelSheet as e:
            messagebox.showwarning('Внимание!',
                                   f"Не найдена вкладка '{str(e)}' в файле-источнике, из которой подгружаются данные. "
                                   f"Возможно, она была переименована.")
            return

        # insert values in table

        self.calculating_window_link.calculating_window.title(MODE.get(mode))
        for day_info in data["day_info"]:
            tag = "white" if day_info.id % 2 == 0 else "gray"
            self.calculating_window_link.operation_table.insert('', 'end', values=
            (day_info.column, day_info.year, day_info.week, day_info.month, day_info.day), tag=tag)

        # insert values in entry fields

        if data.get("path_to_target_file"):
            self.calculating_window_link.target_file_path_field.config(state="normal")
            self.calculating_window_link.target_file_path_field.delete(0, "end")
            self.calculating_window_link.target_file_path_field.insert(0, data.get("path_to_target_file"))
            self.calculating_window_link.target_file_path_field.config(state="readonly")

        if data.get("path_to_source_file"):
            self.calculating_window_link.source_file_path_field.config(state="normal")
            self.calculating_window_link.source_file_path_field.delete(0, "end")
            self.calculating_window_link.source_file_path_field.insert(0, data.get("path_to_source_file"))
            self.calculating_window_link.source_file_path_field.config(state="readonly")

        # Focus all

        child_id = self.calculating_window_link.operation_table.get_children()
        if child_id:
            last_child_id = self.calculating_window_link.operation_table.get_children()[-1]
            self.calculating_window_link.operation_table.focus(last_child_id)
            self.calculating_window_link.operation_table.selection_set(child_id)

        #  Select source file and reinit window

        self.calculating_window_link.source_file_button.config(command=lambda: select_and_reinit_source_file())

        def select_and_reinit_source_file():
            self.calculating_window_link.source_file_path_field.config(state="normal")
            self.calculating_window_link.source_file_path_field.delete(0, "end")
            self.calculating_window_link.source_file_path_field.insert(0, filedialog.askopenfilename(
                initialdir="..",
                title="Выберите источник",
                filetypes=EXCEL_FILE_EXTENSIONS,
            )
                                                                       )
            self.calculating_window_link.source_file_path_field.config(state="readonly")

            self.show_calculating_window(mode, self.calculating_window_link.source_file_path_field.get(), False)

            #  Execute button

        self.calculating_window_link.execute_button.config(command=lambda: execute())

        def execute():
            #  Check fields first

            if not self.calculating_window_link.target_file_path_field.get() \
                    or not self.calculating_window_link.source_file_path_field.get():
                messagebox.showwarning("Внимание!", "Выберите источник и целевой файл.")
                return

            if not self.calculating_window_link.operation_table.get_children():
                messagebox.showwarning("Внимание!", "Не удалось подгрузить даты для расчёта. Выберите источник заново")
                return

            # Display progress info
            self.calculating_window_link.start_progressbar()
            self.calculating_window_link.show_label_pb()

            # Collecting selected days

            day_array = [[], [], [], [], []]
            for selected_item in self.calculating_window_link.operation_table.selection():
                item = self.calculating_window_link.operation_table.item(selected_item)
                day = item["values"]
                day_array[0].append(day[1])
                day_array[1].append(day[3])
                day_array[2].append(day[4])
                day_array[3].append(day[0])
                day_array[4].append(day[2])

            # execute calculating

            queue_var = queue.Queue()  # var for exchange between threads
            executing_thread = Thread(target=calculate_and_echo_to_target_file,
                                      args=(day_array,
                                            self.calculating_window_link.target_file_path_field.get(),
                                            self.calculating_window_link.source_file_path_field.get(),
                                            self.selected_year.get(),
                                            queue_var))
            executing_thread.start()

            def check_exec_thread():
                def end_checking():
                    self.calculating_window_link.update_pb(100)
                    self.calculating_window_link.enable_all_elements()
                    messagebox.showinfo("Успех!",
                                        f"Рассчёт успешно завершён"
                                        f" и занесён в файл"
                                        f" {self.calculating_window_link.target_file_path_field.get()}")

                    self.calculating_window_link.update_pb(0)
                    self.calculating_window_link.progressbar.grid_forget()
                    self.calculating_window_link.label_pb.grid_forget()

                if executing_thread.is_alive() or not queue_var.empty():
                    try:
                        progress_value = queue_var.get(timeout=180)
                    except _queue.Empty:
                        messagebox.showerror("Ошибка!", "Возникла непредвиденная ошибка.")
                        exit(-1)
                    if inspect.isfunction(progress_value):
                        progress_value()
                        self.calculating_window_link.enable_all_elements()
                        self.calculating_window_link.update_pb(0)
                        self.calculating_window_link.progressbar.grid_forget()
                        self.calculating_window_link.label_pb.grid_forget()
                        return
                    elif progress_value == 100:
                        end_checking()
                        return
                    else:
                        self.calculating_window_link.update_pb(progress_value)
                        self.calculating_window_link.calculating_window.after(100, check_exec_thread)
                else:
                    end_checking()
                    return

            self.calculating_window_link.disable_all_elements()
            self.calculating_window_link.calculating_window.after(100, check_exec_thread)

        # Show window. Hide this windows

        if not manual_path_to_source_file:
            self.main_window.withdraw()
            self.calculating_window_link.calculating_window.deiconify()

        self.calculating_window_link.current_window_mode = mode
