import xlrd
loc = 'C:\\Users\\Wolverine\\PycharmProjects\\Automation\\dataFiles\\OFC_NB.xlsx'
wb = xlrd.open_workbook(loc)
data = {}
innerDataMap = {}
innerSheets = []
outData = []
for sheet in wb.sheets():
    for i in range(sheet.ncols):
        if sheet.cell_value(0, i) == 'Automation Key':
            autoIDcol = i
        if sheet.cell_value(0, i) == 'OFC_NB_03':
            for j in range(1, sheet.nrows):
                if ";" in str(sheet.cell_value(j, i)):
                    outData.append(sheet.cell_value(j, i))
                    innerSheets.append(((sheet.cell_value(j, i)).split(';')[0]).split('_')[0])
                data[sheet.cell_value(j, autoIDcol)] = sheet.cell_value(j, i)
print(data.get('NB_Status'))
print(innerSheets)
print(outData)

