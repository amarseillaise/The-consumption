import os
import re
import shutil
from pathlib import Path
import openpyxl
from constants import *
from exceptions import *
from intake_by_plan_corrug_calculate import get_intake_by_plan_corrug_calculate


def get_paths_to_files():
    # get paths to main files

    paths_to_files = {
        "картон.xl": "",
        "пленка.xl": "",
        "сырье.xl": "",
        "production": "",
        "план": "",
        "WeeklyIntakeBySAP": None
    }
    for pattern in paths_to_files.keys():
        for file in os.listdir():
            if re.match(pattern, file.lower()):
                paths_to_files[pattern] = os.getcwd() + r"\\"[0] + file

    # get paths to files in WeeklyIntakeBySAP dir

    try:
        files = []
        for file in os.listdir(path="./" + NAME_OF_DIR_SAP_DEMAND):
            if re.match(r"20\d{2}-\d{2}.(xl|XL)", file):
                files.append(os.getcwd() + r"\\"[0] + NAME_OF_DIR_SAP_DEMAND + r"\\"[0] + file)
                paths_to_files[NAME_OF_DIR_SAP_DEMAND] = files
    except IOError:
        pass

    return paths_to_files


def get_corrug_calculating_days_from_plan(path_to_file):
    plan_wb = openpyxl.load_workbook(filename=path_to_file, data_only=True)
    try:
        plan_wb_s = plan_wb["Plan"]
    except KeyError:
        raise UnableToFindExcelSheet("Plan")
    day_array = [[], [], [], []]
    today_value = ""

    for i in range(1, plan_wb_s.max_column):
        today_value = plan_wb_s.cell(row=3, column=i).value
        if today_value and str(today_value) == str(TODAY):
            break
    if str(today_value) != str(TODAY):
        plan_wb.close()
        raise TodayNotFindInSourceFileException("Не удалось найти ячейку текущего дня в файле плана.")

    for j in range(i, plan_wb_s.max_column):
        val = plan_wb_s.cell(row=3, column=j).value
        if val is not None:
            day_array[0].append(int(val.strftime("%Y")))
            day_array[1].append(int(val.strftime("%m")))
            day_array[2].append(int(val.strftime("%d")))
            day_array[3].append(int(j))

    plan_wb.close()
    return day_array


def calculate_and_echo_to_target_file(selected_days_arr, path_to_target_file, path_to_source_file, selected_year, pb):
    # reserve copy first

    Path("./ReserveCopy").mkdir(parents=True, exist_ok=True)

    shutil.copyfile(path_to_target_file, "./ReserveCopy/" + "reserve_copy_" + os.path.basename(path_to_target_file))
    shutil.copyfile(path_to_source_file, "./ReserveCopy/" + "reserve_copy_" + os.path.basename(path_to_source_file))

    return get_intake_by_plan_corrug_calculate(selected_days_arr[0:-1], path_to_target_file, path_to_source_file,
                                               selected_year)
