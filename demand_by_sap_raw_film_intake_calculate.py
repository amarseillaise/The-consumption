import os
from openpyxl.styles import Font
import openpyxl
import datetime


def getRawConsumption(raw_file):
    now = datetime.datetime.today().strftime("%d.%m")
    raw = openpyxl.load_workbook(raw_file, data_only=False)
    raw_s = raw['Сырье']
    day_array = [[], []]
    required_dates = [[], []]

    for j in os.listdir(path='WeeklyIntakeBySAP'):  # collecting week and day from folder with SAP_load
        day_array[0].append(str(j[0:4]))
        day_array[1].append(str(j[5:7]))

    for i in range(3, raw_s.max_column):  # detecting coordinates in raw file (year)
        val = raw_s.cell(row=2, column=i).value
        if val is not None:
            if str(val) in day_array[0]:
                break

    for j in range(i, raw_s.max_column):  # detecting coordinates in raw file (weak)
        val = str(raw_s.cell(row=2, column=j).value)
        if val in day_array[1] and val not in required_dates[1]:
            required_dates[0].append(j)
            required_dates[1].append(val)

    for raw_demand_path in os.listdir(path='WeeklyIntakeBySAP'):  # Here begin the main cycle
        demand = openpyxl.load_workbook(os.getcwd() + '/WeeklyIntakeBySAP/' + raw_demand_path, data_only=True)
        demand_s = demand.worksheets[0]
        collected_data = [[], [], []]

        for i in range(2, demand_s.max_row):  # collecting SKU-consumption from SAP load
            curr_sku = demand_s.cell(row=i, column=1).value
            curr_quant = int(demand_s.cell(row=i, column=4).value)
            if curr_quant >= 0:
                collected_data[0].append(raw_demand_path[5:7])
                collected_data[1].append(curr_sku)
                collected_data[2].append(curr_quant)

        for i in range(2, raw_s.max_row):  # 5

            val = raw_s.cell(row=i, column=2).value

            if str(val) in str(collected_data[1]):
                # print(i)

                for j in range(len(collected_data[1])):

                    act_sku = collected_data[1][j]
                    act_date = collected_data[0][j]
                    act_qnt = collected_data[2][j]

                    if str(val) == str(act_sku):
                        for d in range(7):
                            if raw_s.cell(row=i - 2,
                                          column=required_dates[0][
                                                     required_dates[1].index(
                                                         act_date)] + d).value is None and "TOTAL" not in str(
                                raw_s.cell(row=1, column=required_dates[0][required_dates[1].index(
                                    act_date)] + d).value):
                                raw_s.cell(row=i - 2,
                                           column=required_dates[0][
                                                      required_dates[1].index(act_date)] + d).value = str(
                                    act_qnt) + " |" + str(now)
                                break

                        raw_s.cell(row=i - 2,
                                   column=required_dates[0][required_dates[1].index(act_date)] + d).font = Font(
                            color='00008B', bold=False, size=11, name="Arial")

                        # Тут начинается точечный расход по дням
                        additional_row = False  # булеан для разделителя месяцев
                        if act_qnt > 0:
                            for h in range(3, 7):
                                if "TOTAL" in str(raw_s.cell(row=1, column=required_dates[0][required_dates[1].index(
                                        act_date)] + h).value):
                                    additional_row = True
                                    continue
                                raw_s.cell(row=i + 1,
                                           column=required_dates[0][required_dates[1].index(act_date)] + h).value = int(
                                    act_qnt) / 4
                                raw_s.cell(row=i + 1,
                                           column=required_dates[0][required_dates[1].index(act_date)] + h).font = Font(
                                    color='00008B', bold=False, size=11, name="Arial")
                            if additional_row:
                                raw_s.cell(row=i + 1,
                                           column=required_dates[0][
                                                      required_dates[1].index(act_date)] + h + 1).value = int(
                                    act_qnt) / 4
                                raw_s.cell(row=i + 1,
                                           column=required_dates[0][
                                                      required_dates[1].index(act_date)] + h + 1).font = Font(
                                    color='00008B', bold=False, size=11, name="Arial")

    raw.save(raw_file)
