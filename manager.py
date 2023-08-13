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

    def collect_data(target_name_sheet_key, source_name_sheet_key, executing_func,
                     window_title_name, indexes):

        path_to_target_file = ""
        if not manual_path_to_source_file:
            paths_to_files = get_paths_to_files()
            path_to_target_file = paths_to_files.get(target_name_sheet_key)
            path_to_source_file = paths_to_files.get(source_name_sheet_key)
        else:
            path_to_source_file = manual_path_to_source_file

        if path_to_source_file:
            got_data = executing_func(path_to_source_file)

            pre_result = []
            for i in range(len(got_data[0])):
                pre_result.append((ExcelDayInfo(i + 1, got_data[indexes[0]][i], got_data[indexes[1]][i],
                                                got_data[indexes[2]][i], got_data[indexes[3]][i],
                                                got_data[indexes[4]][i])))
            result["day_info"] = pre_result
            result["path_to_source_file"] = path_to_source_file
            result["path_to_target_file"] = path_to_target_file
            result["window_title"] = window_title_name

        return result

    # CORRUG by production plan
    if mode == 0:
        result = collect_data("картон.xl", "план", get_corrug_calculating_days_from_plan,
                     "Расчёт расхода картона по плану производства", (3, 0, 4, 1, 2))

    # CORRUG by forecast plan
    elif mode == 1:
        result = collect_data("картон.xl", "production", get_corrug_calculating_weeks_from_forecast,
                     "Расчёт расхода картона по прогнозу", (1, 2, 0, 4, 4))

    # FILM by SAP forecast
    elif mode == 3:
        result = collect_data("пленка.xl", "weeklyintakebysap", get_film_raw_week_year_from_sap_demand,
                     "Расчёт расхода плёнки по деманду", (4, 0, 1, 4, 4))

    # RAW by SAP forecast
    elif mode == 5:
        result = collect_data("сырье.xl", "weeklyintakebysap", get_film_raw_week_year_from_sap_demand,
                     "Расчёт расхода сырья по деманду", (4, 0, 1, 4, 4))

    elif mode in SIMPLE_MODS:
        pass

    return result
