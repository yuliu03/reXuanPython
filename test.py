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


def file_name(file_dir):
    filesList = []
    for root, dirs, files in os.walk(file_dir):
        # print(root)  # 当前目录路径
        # print(dirs)  # 当前路径下所有子目录
        # print(files)  # 当前路径下所有非目录子文件
        filesList = files
    return filesList

print(file_name("C:/Users/admin/Desktop/reXuan/e电宝入库/11.5义乌补货清单"))


