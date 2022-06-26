# 将条形码生成图片

import os, datetime, time
import tkinter as tk
from tkinter import filedialog

# import barcode
# from barcode.writer import ImageWriter

# 导入本地文件
EAN_list = []
root = tk.Tk()
root.withdraw()

print("请选择（输入序号）： \n"
      "1.打开单个文件\n"
      "2.打开多个文件\n"
      "3.打开文件夹下所有文件")
choose = input("您的选择：")
chosen = int(choose)
if chosen == 1:
    # 选择单个文件
    # print("您的选择是：1.打开单个文件")
    file_path = filedialog.askopenfile(title="Secret txt or xlsx file", filetypes=[("文本文档或表格", ".txt .xls .xlsx")])
    print(file_path)
elif chosen == 2:
    # 限制格式选择多个文件
    print("您的选择是：2.打开多个文件")
    file_paths = filedialog.askopenfiles(title="Select TXT & XLSX files",
                                         filetypes=[("文本文档或表格", ".txt .xls .xlsx")])  # 字典[(A,B),(A,B),(A,B),]
    print(file_paths)
elif chosen == 3:
    # 打开文件夹
    print("您的选择是：3.打开文件夹下所有文件")
    folder_path = filedialog.askdirectory()
    print(folder_path)
else:
    print("选择错误，请重新选择！")
    choose = input("重新选择：")

"""
# for file in file_paths:
#     f = open(file, "r+")
#     while True:
#         line = f.readline()  # 读文件
#         if line == '':
#             f.close()
#             break
#         else:
#             EAN_list.append(line)
# print(EAN_list)

# currentTime = time.time()
# os.mkdir("./images/" + str(currentTime) + "输出")
#
# x = 1
# for i in EAN_list:
#   EAN = barcode.get_barcode_class("code128")
#   ean = EAN(i, writer=ImageWriter)
#   #去除特殊字符
#   j = i.replace("\n", "")
#   k = j.replace(":", "")
#   #输出文件夹
#   route = ean.save("./images/" + str(currentTime) + "输出/" + str(k) + "barcode")
#   x = x + 1
#   print("第" + str(x) + "个已成功，输出文件：" + route)
"""
