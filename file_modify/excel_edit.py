# -*- coding:utf-8 -*-
# authored by rainrelaxme
# 操作excel文件

# 导入需要使用的包
import os
import xlrd  # version：1.2.0，读取Excel文件的包，不能用2.0以上版本
import tkinter as tk  # 导入打开文件浏览器的包
from tkinter import filedialog

# 导入自己编写的模块
from file_edit import add_x


def choose_file(is_single: int = 0, filetypes=None):
    """用文件浏览器选择文件"""
    if filetypes is None:  # 形参调整，不用写在括号里
        # 只打开一种格式文件
        # inputPath=askopenfilename(title="Select PDF file", filetypes=(("pdf files", "*.pdf"),))
        filetypes = [("All files", "*.*")]
    root = tk.Tk()  # 实例化
    root.withdraw()  # 销毁窗口
    if is_single == 0:  # is_single=0 ：选择单个文件
        file_path = tk.filedialog.askopenfilename(filetypes=filetypes)
    else:
        file_path = tk.filedialog.askopenfilenames(filetypes=filetypes)
    return file_path


def save_file(initialfile="Untitled", filetypes=None, defaultextension=None):
    """打开文件浏览器，存储文件"""
    """initialfile:默认文件名，filetypes：文件后缀名选择，"""
    if filetypes is None:  # 形参调整，不用写在括号里
        filetypes = [("All files", "*.*")]
    root = tk.Tk()  # 实例化
    root.withdraw()  # 销毁窗口
    file = tk.filedialog.asksaveasfile(filetypes=filetypes, initialfile=initialfile, defaultextension=defaultextension)
    # 妄图修改使其添加默认的后缀名，fail
    # file.close()
    # file_new_name = add_x(file.name, ".xlsx")
    # os.rename(file.name, file_new_name)
    # return open(file_new_name, "w")
    return file


def open_xls(file):
    """打开一个excel文件"""
    file_opened = xlrd.open_workbook(file)
    return file_opened


def get_sheet(file):
    """获取excel中所有的sheet表"""
    return file.sheets()


def get_sheet_num(file):
    """获取sheet表的个数"""
    num = 0
    sh = get_sheet(file)
    for sheet in sh:
        num += 1
    return num


def get_all_rows(file, sheet):
    """获取sheet表的行数"""
    table = file.sheets()[sheet]
    return table.nrows


def get_file(file, sheet_num):
    """读取文件内容并返回行内容"""
    data_value = []
    file_opened = open_xls(file)
    table = file_opened.sheets()[sheet_num]
    num = table.nrows
    for row in range(num):
        rdata = table.row_values(row)
        data_value.append(rdata)
    return data_value


# print(get_file('E:/project/pythonProject/Little_tools/src/3.xlsx', 0))
