# -*- coding:utf-8 -*-
# authored by rainrelaxme
# 操作excel文件

# 导入需要使用的包
import xlrd  # 读取Excel文件的包
import xlsxwriter  # 将文件写入Excel的包
import tkinter as tk  # 导入打开文件浏览器的包
from tkinter import filedialog


def choose_file(is_single: int = 0, filetypes=None):
    """用文件浏览器选择文件"""
    if filetypes is None:   # 形参调整，不用写在括号里
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
    file_opened = open_xls(file)
    table = file_opened.sheets()[sheet_num]
    num = table.nrows
    for row in range(num):
        rdata = table.row_values(row)
        datavalue.append(rdata)
    return datavalue
