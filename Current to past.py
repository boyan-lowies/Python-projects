import openpyxl as op

wb = op.load_workbook('C:\\Users\\lowie\\Desktop\\python test.xlsx', data_only=True)
sheetTP = wb['TP current']
sheetP = wb['TP Past']

N = 2


while sheetTP['D' + str(N)].value is not None:
    if sheetTP['B' + str(N)].value == -1:
        print(N)
    N += 1