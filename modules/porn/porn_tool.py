#!/usr/bin/env python3
# -*- coding:utf-8 -*-
"""
@Project: tools
@File   : porn_tool
@Version:
@Author : RainRelaxMe
@Date   : 2025/10/6 1:49
@Info   : 
"""

import os


def get_all_filename(filepath):
    """
    获取文件夹下所有的文件名称。
    """
    all_files = []
    for root, dirs, files in os.walk(filepath):
        for file in files:
            file_path = os.path.join(root, file)
            all_files.append(file_path)
    return all_files


def get_keyname():
    """去除前面的内容，获取关键的文件名称，即番号名称"""
    pass


if __name__ == '__main__':
    # 使用示例
    directory_path = r"G:\迅雷下载\学习资料\移动"
    files = get_all_filename(directory_path)
    for file in files:
        print(file)
