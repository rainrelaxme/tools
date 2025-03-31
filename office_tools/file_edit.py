# -*- coding:utf-8 -*-
# authored by RainRelaxMe
# 对文件的操作

import os
import re  # 引入正则


def get_filepath():
    """获取当前程序所在的文件路径"""
    path = os.getcwd()
    return path


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


def split_name(name):
    """分离文件名与扩展名；默认返回(fname,fextension)元组，可做分片操作，后缀名和“.”在一起"""
    name_tuple = os.path.splitext(name)
    forward_name = name_tuple[0]
    back_name = name_tuple[1]
    return forward_name, back_name


def sub_space(name):
    """去除文件名称内的空格   str.split & join"""
    new_name = "".join(name.split())
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


