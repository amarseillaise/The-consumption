from datetime import datetime

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
CURRENT_WEEK = int(datetime.today().strftime("%W"))
FONT_STYLE = "Calibri 12"
TODAY = datetime.today().strftime("%Y-%m-%d") + " 00:00:00"
VALUES_FOR_HEADINGS_TABLE = ("Column", "Year", "Week", "Month", "Day")
TABLE_COLUMNS_WIDTH = 110
VALUES_FOR_CATEGORY_TABLE = (("Упаковка",), ("Плёнка",), ("Сырьё",))
EXCEL_FILE_EXTENSIONS = (("Excel files", "*.xlsx"), ("Excel files", "*.xls"),
                         ("Excel files", "*.xlsm"), ("Excel files", "*.xlsb"))

NAME_OF_DIR_SAP_DEMAND = "WeeklyIntakeBySAP"
MODE = {
    0: "Расчёт расхода картона по плану производства",
    1: "Расчёт расхода картона по прогнозу",
    2: "Расчёт фактического расхода картона на сегодня",
    3: "Расчёт расхода плёнки по деманду",
    4: "Расчёт фактического расхода плёнки на сегодня ",
    5: "Расчёт расхода сырья по деманду",
    6: "Расчёт фактического расхода сырья на сегодня ",
}
SIMPLE_MODS = (2, 4, 6)
SAP_DEMAND_MODS = {
    3: "ПЛЕНКА",
    5: "Сырье"
}
NAMES_SHEETS_IN_SIMPLE_MOD = {
    2: "КАРТОН",
    4: "ПЛЕНКА",
    6: "Сырье"
}

