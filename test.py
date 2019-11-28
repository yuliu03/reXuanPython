#coding=gbk
# import pyautogui
#
# screenWidth, screenHeight = pyautogui.size()
# print(screenWidth, screenHeight)
#
# currentMouseX, currentMouseY = pyautogui.position()
# print(currentMouseX, currentMouseY )
#
# pyautogui.moveTo(100, 150)
import os

from openpyxl import load_workbook
from openpyxl.styles import Side, Border


def file_name(file_dir):
    filesList = []
    for root, dirs, files in os.walk(file_dir):
        # print(root)  # 当前目录路径
        # print(dirs)  # 当前路径下所有子目录
        # print(files)  # 当前路径下所有非目录子文件
        filesList = files
    return filesList


#增加边框
readFilepath = "C:/Users/admin/Desktop/test.xlsx"
wb = load_workbook(readFilepath)
side = Side(border_style='thin',color='000000')
newSheet = wb.active
border = Border(left=side,
right=side,
top=side,
bottom=side)

newSheet.cell(row=1, column=1).border = border
wb.save(readFilepath)



