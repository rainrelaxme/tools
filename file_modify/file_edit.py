# -*- coding:utf-8 -*-
# authored by rainrelaxme
# 对系统文件的操作

import os
import re   # 引入正则


def get_filepath():
    """获取当前程序所在的文件路径"""
    path = os.getcwd()
    return path


def get_filelist(path):
    """获取文件夹下的文件列表"""
    file_list = os.listdir(path)  # 获取该目录下所有文件，存入列表中
    return file_list


def split_name(name):
    """分离文件名与扩展名；默认返回(fname,fextension)元组，可做分片操作，后缀名和“.”在一起"""
    name_tuple = os.path.splitext(name)
    forward_name = name_tuple[0]
    back_name = name_tuple[1]
    return forward_name, back_name


def sub_space1(name):
    """去除文件名称内的空格   str.replace"""
    newName = name.replace(" ", "")
    return newName


def sub_space2(name):
    """去除文件名称内的空格   str.split & join"""
    newName = "".join(name.split())
    return newName


def sub_space3(name):
    """去除文件名称内的空格   正则"""
    pattern = re.compile(r'\s+')
    newName = re.sub(pattern, '', name)
    return newName


def add_space(name):
    """文件名称加上空格"""
    newName = ' '.join(name)
    return newName


def add_x(name, x):
    """文件名称结尾加上自定义内容"""
    newName = name + x
    return newName


def sub_x(name, x):
    """文件名称头尾去除自定义内容。"""
    """Python中有三个去除头尾字符、空白符的函数，它们依次为:
        strip： 用来去除头尾字符、空白符(包括\n、\r、\t、' '，即：换行、回车、制表符、空格)
        lstrip：用来去除开头字符、空白符(包括\n、\r、\t、' '，即：换行、回车、制表符、空格)
        rstrip：用来去除结尾字符、空白符(包括\n、\r、\t、' '，即：换行、回车、制表符、空格)"""
    if not name.endswith(x):
        pass
    else:
        newName = name.strip(x)     # 同样的内容会全部去除，而非一遍
    return newName




