from GUI.main_window import MainWindow
from GUI.calculating_window import CalculatingWindow


def init_gui():
    cw = CalculatingWindow()  # Create an object with calculating window
    mw = MainWindow(cw)  # Create an object with a main window and link it with a calc window
    return mw


def init_files_to_gui(mw, mode: int):
    if mode == 0:
        mw.main_window.calculating_window_link.calculating_window.operation_table.destroy()
