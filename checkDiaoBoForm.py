#coding=gbk
import os

import pymysql
from openpyxl import load_workbook

#只是检查各个字段的抬头
#######################

#获取个性化最大列数，只读取第一行，最大行数为第一个字段为 string 所属行的位置
def personal_max_row(sheet,string):
    maxRowToCheck = 5000
    count = 1
    value = sheet.cell(row=count, column=1).value
    if value is None:
        value = ""
    while count <= maxRowToCheck and string not in value:
        count = count + 1
        value = sheet.cell(row=count, column=1).value
        if value is None:
            value = ""
    return count

#获取个性化最大列数，只读取第一列，最大列数为第一个空白cell
def personal_max_col(sheet):
    count = 1
    beginRow = 1
    value = sheet.cell(row=1, column=count).value
    if value is None or "条码" not in value:
        beginRow = 2

    while sheet.cell(row=beginRow, column=count).value is not None:
        count = count + 1
    return count-1

#此方法查询路径下所有文件，包括子目录文件
def file_name(file_dir):
    toReturn = None
    for root, dirs, files in os.walk(file_dir):
        # print("=========当前目录路径=================")
        # print(root)  # 当前目录路径
        # print("=============前路径下所有子目录===================")
        # print(dirs)  # 当前路径下所有子目录
        # print("===============前路径下所有非目录子文件=================")
        toReturn = (files) # 当前路径下所有非目录子文件
        # toReturn = (files)
    return toReturn

#检查是否item
def checkInfoFromSheetsAux(sheet,item):
    beginRow = personal_max_row(sheet, "国际条码")
    maxCol = personal_max_col(sheet)
    count = 1
    while count <= maxCol:
        if item in sheet.cell(row=beginRow, column=count).value:
            return True
        count = count + 1

    return  False

#检查是否有以下几个字段： 国际条码",	"商品名称",	"单位",	"规格",	"调拨门店",	"调拨数量",	"实际可调拨数量（仓库填写）" 的字段
def checkInfoFromSheets(sheet):
    toCheckNames = ["国际条码",	"商品名称",	"单位",	"规格",	"调拨门店",	"调拨数量",	"实际可调拨数量（仓库填写）"]

    for item in toCheckNames:
        if not checkInfoFromSheetsAux(sheet,item):
            print(item + "未出现")
            return False

    return True



root = "C:/Users/admin/Desktop/reXuan/晕晕/佳成出库"
readFilepathList = file_name(root)
index = 0
size = len(readFilepathList)

badWorkBookList = list()

#检查所有excel
while index < size:
    try:
        wb = load_workbook( root + "/" + readFilepathList[index])
        sheetNamesList = wb.sheetnames #获取所有sheet名称
        if "仓库填写" in sheetNamesList:
            sheetNamesList.remove("仓库填写")
        for sheetName in sheetNamesList:
            if not checkInfoFromSheets(wb[sheetName]):
                badWorkBookList.append(root + "/" + readFilepathList[index]) #记录为符合标准excel
        wb.close()
        index = index + 1
    except:
        print("error +++")
        badWorkBookList.append(root + "/" + readFilepathList[index])
        index = index + 1
        pass

print("============")
print(badWorkBookList)
# print(readFilepathList)
i = 0
# 默认可读写，若有需要可以指定write_only和read_only为True

# sheet = wb.active