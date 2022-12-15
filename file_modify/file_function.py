# -*- coding:utf-8 -*-
# authored by rainrelaxme
# 说明：文件工具的中间过程及最终函数实现

import os
import xlsxwriter

import file_edit
import excel_edit


def rename_add_space(path):
    """文件名称增加空格"""
    file_list = os.listdir(path)
    for file in file_list:
        absolute_file = path + "/" + file
        if os.path.isfile(absolute_file):  # 不处理文件夹
            if file.endswith(".py") or file.endswith(".exe") or file.endswith(".txt") or file.endswith(".ipynb"):
                continue
            else:
                old_name = absolute_file
                forward_name = file_edit.split_name(file)[0]
                forward_name = file_edit.add_space(forward_name)
                back_name = file_edit.split_name(file)[1]
                new_name = path + "/" + forward_name + back_name
                os.rename(old_name, new_name)
        else:
            # 递归文件夹修改文件名
            rename_add_space(absolute_file)


def rename_sub_space(path):
    """文件名称去除空格"""
    file_list = os.listdir(path)
    for file in file_list:
        absolute_file = path + "/" + file
        if os.path.isfile(absolute_file):  # 不处理文件夹
            if file.endswith(".py") or file.endswith(".exe") or file.endswith(".txt") or file.endswith(".ipynb"):
                continue
            else:
                old_name = absolute_file
                forward_name = file_edit.split_name(file)[0]
                forward_name = file_edit.sub_space1(forward_name)
                back_name = file_edit.split_name(file)[1]
                new_name = path + "/" + forward_name + back_name
                os.rename(old_name, new_name)
        else:
            # 递归文件夹修改文件名
            rename_sub_space(absolute_file)


def rename_add_x(path, x):
    """给当前路径下的所有文件名后边都加“x”"""
    file_list = os.listdir(path)
    for file in file_list:
        absolute_file = path + "/" + file
        if os.path.isfile(absolute_file):  # 不处理文件夹
            if file.endswith(".py") or file.endswith(".exe") or file.endswith(".txt") or file.endswith(".ipynb"):
                continue
            else:
                old_name = absolute_file
                forward_name = file_edit.split_name(file)[0]
                forward_name = file_edit.add_x(forward_name, x)
                back_name = file_edit.split_name(file)[1]
                new_name = path + "/" + forward_name + back_name
                os.rename(old_name, new_name)
        else:
            # 递归文件夹修改文件名
            rename_add_x(absolute_file, x)


def rename_sub_x(path, x):
    """给当前路径下的所有文件名后边都去除“x”"""
    file_list = os.listdir(path)
    for file in file_list:
        absolute_file = path + "/" + file
        if os.path.isfile(absolute_file):  # 不处理文件夹
            if file.endswith(".py") or file.endswith(".exe") or file.endswith(".txt") or file.endswith(".ipynb"):
                continue
            else:
                old_name = absolute_file
                forward_name = file_edit.split_name(file)[0]
                forward_name = file_edit.sub_x(forward_name, x)
                back_name = file_edit.split_name(file)[1]
                new_name = path + "/" + forward_name + back_name
                os.rename(old_name, new_name)
        else:
            # 递归文件夹修改文件名
            rename_sub_x(absolute_file, x)


def format_name1(path, girl):
    """rename1(name, girl)"""
    pass


def format_name2(path, forward_name, start_num: int):
    """rename2(name, forward_name, start_num: int)"""
    pass


def excel_merge(files: tuple):
    """利用Python将多个excel文件合并为一个文件"""
    # 存储所有读取的结果
    value = []
    for file in files:
        f_open = excel_edit.open_xls(file)
        sh_num = excel_edit.get_sheet_num(f_open)
        for sheet in range(sh_num):
            print("正在读取文件：" + str(file) + "的第" + str(sheet+1) + "个sheet表的内容...")
            sheet_value = excel_edit.get_file(file, sheet)
            if sheet_value:
                value.append(sheet_value)
    print(value)
    # 定义最终合并后生成的新文件
    filetypes = [("Microsoft Excel文件", "*.xlsx"),
                 ("Microsoft Excel文件", "*.xls"),
                 ("CSV文件", "*.csv")]
    result_file = excel_edit.save_file(filetypes=filetypes, defaultextension=filetypes[0][1])
    work_file = xlsxwriter.Workbook(result_file.name)  # 创建工作表
    work_sheet = work_file.add_worksheet()  # 默认创建sheet1
    for sheet in range(len(value)):
        for row in range(len(value[sheet])):
            for col in range(len(value[sheet][row])):
                cell = value[sheet][row][col]
                work_sheet.write(row, col, cell)
    work_file.close()


all_file = ('E:/project/pythonProject/Little_tools/src/1.xlsx', 'E:/project/pythonProject/Little_tools/src/2.xlsx')
# all_file = ('E:/project/pythonProject/Little_tools/src/3.xlsx')
excel_merge(all_file)
