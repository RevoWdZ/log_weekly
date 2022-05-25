from openpyxl import Workbook
from datetime import date, timedelta

wb = Workbook()
ws = wb.active
ws.title = "Data"

ws.append(["ID", "Log Description/Name", "Review Data", "Person", "Remarks"])
sh = wb.active

if date.today().weekday() in range(0, 5):
    now = date.today()

log_files = ["/xyz/xyz1/xyz2/xyz12.log",
             "/xyz/xyz1/xyz2/xyz123.log", "/xyz/xyz1/xyz2/xyz12." + str(now) + ".log"]
id = 0
iteration = -1
days_before = 5
for row in sh.iter_rows(min_row=1, min_col=1, max_row=1, max_col=5):
    for cell in row:
        for x in log_files:

            id += 1
            iteration += 1

            if (iteration % 3 == 0):
                days_before -= 1
                ws.append([id, x, now - timedelta(days_before),
                          "NAME", "OK"])
                log_files[2] = "/xyz/xyz1/xyz2/xyz12." + \
                    str(now - timedelta(days_before)) + ".log"
                continue
            else:
                ws.append([id, x, now - timedelta(days_before),
                          "NAME", "OK"])

wb.save(r"C:\Users\Public\Log Weekly " + str(now.isocalendar().week) + "_" +
        str(now.isocalendar().year) + ".xlsx")
