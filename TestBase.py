import xlrd
from selenium import webdriver
def readexcel(wb):
    for sheet in wb.sheets():
        for i in range(sheet.ncols):
            if sheet.cell_value(0, i) == 'Automation Key':
                autoIDcol = i
            if sheet.cell_value(0, i) == 'OFC_NB_03':
                for j in range(1, sheet.nrows):
                    data[sheet.cell_value(j, autoIDcol)] = sheet.cell_value(j, i)
                break
    print(data)
loc = 'C:\\Users\\Wolverine\\PycharmProjects\\Automation\\dataFiles\\OFC_NB.xlsx'
wb = xlrd.open_workbook(loc)
data = {}
readexcel(wb)
