#coding=gbk
import os
from openpyxl import load_workbook
readFilepath = "C:/Users/admin/Desktop/test.xlsx"
wb = load_workbook(readFilepath)
newSheet = wb.active
