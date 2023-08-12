from GUI.main_window import MainWindow
from GUI.calculating_window import CalculatingWindow
from data_io import *
from objects import ExcelDayInfo


def init_gui():
    cw = CalculatingWindow()  # Create an object with calculating window
    mw = MainWindow(cw, init_files_to_gui)  # Create an object with a main window and link it with a calc window
    return mw


def init_files_to_gui(mode: int, manual_path_to_source_file=""):
    result = {
        "day_info": [],
        "path_to_target_file": "",
        "path_to_source_file": "",
        "window_title": ""
    }

    # CORRUG by production plan

    if mode == 0:
        path_to_target_file = ""
        if not manual_path_to_source_file:
            paths_to_files = get_paths_to_files()
            path_to_target_file = paths_to_files.get("картон.xl")
            path_to_source_file = paths_to_files.get("план")
        else:
            path_to_source_file = manual_path_to_source_file

        if path_to_source_file:
            got_data = get_corrug_calculating_days_from_plan(path_to_source_file)

            pre_result = []
            for i in range(len(got_data[0])):
                pre_result.append(ExcelDayInfo(i + 1, got_data[3][i], got_data[0][i], "", got_data[1][i], got_data[2][i]))
            result["day_info"] = pre_result
            result["path_to_source_file"] = path_to_source_file
            result["path_to_target_file"] = path_to_target_file
            result["window_title"] = "Расчёт картона по плану производства"

    # CORRUG by forecast SAP

    elif mode == 1:
        path_to_target_file = ""
        if not manual_path_to_source_file:
            paths_to_files = get_paths_to_files()
            path_to_target_file = paths_to_files.get("картон.xl")
            path_to_source_file = paths_to_files.get("production")
        else:
            path_to_source_file = manual_path_to_source_file

        if path_to_source_file:
            got_data = get_corrug_calculating_weeks_from_forecast(path_to_source_file)

            pre_result = []
            for i in range(len(got_data[0])):
                pre_result.append((ExcelDayInfo(i + 1, got_data[1][i], CURRENT_YEAR, got_data[0][i], "", "")))
            result["day_info"] = pre_result
            result["path_to_source_file"] = path_to_source_file
            result["path_to_target_file"] = path_to_target_file
            result["window_title"] = "Расчёт картона по прогнозу"

        # path_to_target_file = ""
        # if not manual_path_to_source_file:
        #     paths_to_files = get_paths_to_files()
        #     path_to_target_file = paths_to_files.get("картон.xl")
        #     path_to_source_dir = "./" + NAME_OF_DIR_SAP_DEMAND
        # else:
        #     path_to_source_dir = manual_path_to_source_file
        #     !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        #     pre_result = []
        #     for i in os.listdir(path=path_to_source_dir):  # collecting week and day from folder with SAP_load
        #         pre_result.append(ExcelDayInfo(i + 1, got_data[3][i], got_data[0][i], "", got_data[1][i], got_data[2][i]))
        #     weeks_for_raw_demand[0].append(str(j[0:4]))
        #     weeks_for_raw_demand[1].append(str(j[5:7]))

    return result
