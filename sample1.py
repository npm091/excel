import openpyxl

wb = openpyxl.load_workbook("sample.xlsx", data_only=True)
ws = wb["Sheet1"]

for num in range(1, 20):
    c1 = ws.cell(row=num, column=6).value
    print(c1)

