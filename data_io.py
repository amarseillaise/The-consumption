import datetime
import os
import re
import openpyxl
from exceptions import TodayNotFindInPlanFileException

NAME_OF_DIR_SAP_DEMAND = "WeeklyIntakeBySAP"
TODAY = datetime.datetime.today().strftime("%Y-%m-%d") + " 00:00:00"


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
    plan_wb_s = plan_wb["Plan"]
    day_array = [[], [], [], []]
    today_value = ""

    for i in range(1, plan_wb_s.max_column):
        today_value = plan_wb_s.cell(row=3, column=i).value
        if today_value and str(today_value) == str(TODAY):
            break
    if str(today_value) != str(TODAY):
        plan_wb.close()
        return TodayNotFindInPlanFileException("Не удалось найти ячейку текущего дня в файле плана.")

    for j in range(i, plan_wb_s.max_column):
        val = plan_wb_s.cell(row=3, column=j).value
        if val is not None:
            day_array[0].append(int(val.strftime("%Y")))
            day_array[1].append(int(val.strftime("%m")))
            day_array[2].append(int(val.strftime("%d")))
            day_array[3].append(int(j))

    plan_wb.close()
    return day_array

