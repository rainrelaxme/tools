# -*- coding:utf-8 -*-
# authored by rainrelaxme
# 对系统文件的操作

import os
import re  # 引入正则


def get_filepath():
    """获取当前程序所在的文件路径"""
    path = os.getcwd()
    return path


def get_filelist(path):
    """获取文件夹下的文件列表"""
    file_list = os.listdir(path)  # 获取该目录下所有文件，存入列表中
    return file_list


def get_absolute_filelist(path):
    """获取文件夹下的文件列表的绝对路径-单层"""
    files = os.listdir(path)
    files_list = []
    for file in files:
        file_path = os.path.join(path, file)
        if os.path.isfile(file_path):  # 判断是否为文件
            files_list.append(file_path)
        else:
            pass
    return files_list


"""
无法一直循环下去
def get_absolute_filelist_alllevel(path):
    获取文件夹下的文件列表的绝对路径-多层
    files = os.listdir(path)
    files_list = []
    for file in files:
        file_path = os.path.join(path, file)
        if os.path.isdir(file_path):  # 判断是否为目录
            current_files_list = get_absolute_filelist(file_path)
            files_list.append(current_files_list)
        elif os.path.isfile(file_path):  # 判断是否为文件
            files_list.append(file_path)
    return files_list

path = input("请输入文件夹路径：")
print(get_absolute_filelist_alllevel(path))
"""


def split_name(name):
    """分离文件名与扩展名；默认返回(fname,fextension)元组，可做分片操作，后缀名和“.”在一起"""
    name_tuple = os.path.splitext(name)
    forward_name = name_tuple[0]
    back_name = name_tuple[1]
    return forward_name, back_name


def sub_space1(name):
    """去除文件名称内的空格   str.replace"""
    new_mame = name.replace(" ", "")
    return new_mame


def sub_space2(name):
    """去除文件名称内的空格   str.split & join"""
    new_name = "".join(name.split())
    return new_name


def sub_space3(name):
    """去除文件名称内的空格   正则"""
    pattern = re.compile(r'\s+')
    new_name = re.sub(pattern, '', name)
    return new_name


def add_space(name):
    """文件名称加上空格"""
    new_name = ' '.join(name)
    return new_name


def add_x(name, x):
    """文件名称结尾加上自定义内容"""
    new_name = name + x
    return new_name


def sub_x(name, x):
    """文件名称头尾去除自定义内容。"""
    """Python中有三个去除头尾字符、空白符的函数，它们依次为:
        strip： 用来去除头尾字符、空白符(包括\n、\r、\t、' '，即：换行、回车、制表符、空格)
        lstrip：用来去除开头字符、空白符(包括\n、\r、\t、' '，即：换行、回车、制表符、空格)
        rstrip：用来去除结尾字符、空白符(包括\n、\r、\t、' '，即：换行、回车、制表符、空格)"""
    if not name.endswith(x):
        pass
    else:
        new_name = name.strip(x)  # 同样的内容会全部去除，而非一遍
    return new_name


def rename1(name, girl):
    """将文件名字按照特定要求调整"""
    """此例为调整文件名为ABP-925-RION"""
    if re.match(r'^([a-zA-Z]+)(\d+)(.*)', name) or re.match(r'^([a-zA-Z]+)(\-{2,})(\d+)(.*)', name):  # 检测是否有分割线-
        m = re.match(r'^([a-zA-Z]*)(\-*)(\d+)(.*)', name)
        # 区分四部分，一个括号为一部分
        group1 = m.group(1)
        group2 = m.group(2)
        group3 = m.group(3)
        group4 = m.group(4)
        new_name = group1 + '-' + group3 + girl
        return new_name


def rename2(name, forward_name, start_num: int):
    """按特定序号修改文件名称"""
    new_name = forward_name + str(start_num + 1)
    return new_name


