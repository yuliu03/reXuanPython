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
        # print(root)  # ��ǰĿ¼·��
        # print(dirs)  # ��ǰ·����������Ŀ¼
        # print(files)  # ��ǰ·�������з�Ŀ¼���ļ�
        filesList = files
    return filesList

print(file_name("C:/Users/admin/Desktop/reXuan/e�籦���/11.5���ڲ����嵥"))


