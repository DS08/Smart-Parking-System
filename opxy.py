from openpyxl import load_workbook

wb = load_workbook(filename='Vehicle_Registration.xlsx')

ws = wb.worksheets[0]

rows = ws.rows

next(rows)

def parking_allowed(max_limit_of_parking_slots):
    count = 0
    for row in ws.rows:
        if (row[4].value==True):
            count +=1
    if count < max_limit_of_parking_slots:
         return True
    return False
