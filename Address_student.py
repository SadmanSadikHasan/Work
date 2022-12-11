# Python3 code to select
# data from excel
import xlwings as xw
import openpyxl as xl
from openpyxl import load_workbook
import pandas as pd

# Specifying a sheet
ws = xw.Book("F:\Work\DATASET\Details Data Final Fall-2007.xls").sheets['CSE']

# Selecting data from
# a single cell
v1 = ws.range("B8:B64").value
# v2 = ws.range("F5").value
#v1 = v1.replace('.', "")
for i in range(len(v1)):
    v1[i] = v1[i].replace('.', "")
print("Result:", v1)

v2 = len(v1)

XL = xl.load_workbook("F:\Work\TABLE TEMPLATE\STUDENT_ADDRESS.xlsx")
sheet1 = XL["Result 1"]
for i in range(3, v2):
    sheet1.cell(row=i, column=2).value = v1[i-3]

XL.save('F:\Work\TABLE TEMPLATE\STUDENT_ADDRESS.xlsx')