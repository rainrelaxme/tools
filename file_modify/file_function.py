# -*- coding:utf-8 -*-
# authored by rainrelaxme
# 说明：文件工具的中间过程及最终函数实现

import os

from . import file_edit as fe


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
                forward_name = fe.split_name(file)[0]
                forward_name = fe.add_space(forward_name)
                back_name = fe.split_name(file)[1]
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
                forward_name = fe.split_name(file)[0]
                forward_name = fe.sub_space1(forward_name)
                back_name = fe.split_name(file)[1]
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
                forward_name = fe.split_name(file)[0]
                forward_name = fe.add_x(forward_name, x)
                back_name = fe.split_name(file)[1]
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
                forward_name = fe.split_name(file)[0]
                forward_name = fe.sub_x(forward_name, x)
                back_name = fe.split_name(file)[1]
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


def excel_merge(files: tuple = ()):
    """利用Python将多个excel文件合并为一个文件"""
    for file in files:
        f_open = open


# 函数入口
if __name__ == '__main__':
    # 定义要合并的excel文件列表
    allxls = ['F:/test/excel1.xlsx', 'F:/test/excel2.xlsx']  # 列表中的为要读取文件的路径
    # 存储所有读取的结果
    datavalue = []
    for fl in allxls:
        f = open_xls(fl)
        x = getshnum(f)
        for shnum in range(x):
            print("正在读取文件：" + str(fl) + "的第" + str(shnum) + "个sheet表的内容...")
            rvalue = getFilect(fl, shnum)
    # 定义最终合并后生成的新文件
    endfile = 'F:/test/excel3.xlsx'
    wb = xlsxwriter.Workbook(endfile)
    # 创建一个sheet工作对象
    ws = wb.add_worksheet()
    for a in range(len(rvalue)):
        for b in range(len(rvalue[a])):
            c = rvalue[a][b]
            ws.write(a, b, c)
    wb1.close()

    print("文件合并完成")
