import openpyxl
from datetime import time, timedelta, datetime
import csv

wb_obj = openpyxl.load_workbook('input.xlsx')

sheet = wb_obj.active
a = []
for row in sheet.iter_rows(min_row=3, max_row=16, max_col=26):
    n = 0
    dt = datetime.combine(row[0].value.date(), time(0, 0))
    date_time = dt.strftime("%m/%d/%Y %H:%M:%S")
    for cell in row:
        if cell.value is not None:
            try:
                cell.value.date()
            except AttributeError:
                a.append((date_time, "XYZ", cell.value))
                n += 1
                dt = datetime.combine(row[0].value.date(), time(0, 0)) + timedelta(hours=n)
                date_time = dt.strftime("%m/%d/%Y %H:%M:%S")
for i in a:
    print(i)

with open('output.csv', 'w') as f:
    # create the csv writer
    writer = csv.writer(f)

    # write a row to the csv file
    writer.writerow(('dateTime', 'deviceId', 'value'))
    for i in a:
        writer.writerow(i)
    f.close()
