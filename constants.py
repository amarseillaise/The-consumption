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
FONT_STYLE = "Calibri 12"
NAME_OF_DIR_SAP_DEMAND = "WeeklyIntakeBySAP"
TODAY = datetime.today().strftime("%Y-%m-%d") + " 00:00:00"
VALUES_FOR_HEADINGS_TABLE = ("Column", "Year", "Week", "Month", "Day")
TABLE_COLUMNS_WIDTH = 110
VALUES_FOR_CATEGORY_TABLE = (("Упаковка",), ("Плёнка",), ("Сырьё",))
EXCEL_FILE_EXTENSIONS = (("Excel files", "*.xlsx"), ("Excel files", "*.xls"),
                         ("Excel files", "*.xlsm"), ("Excel files", "*.xlsb"))
MODE = {
    0: "Расчёт картона по плану производства"
}

