# -*- coding:utf-8 -*-
import os 

def find_prefix_of_path(path, file_list):
    files = os.listdir(path)
    for file_name in files:
        file_path = os.path.join(path, file_name)

        if os.path.isdir(file_path):    #判断是否为目录
            find_prefix_of_path(file_path, file_list)
        elif os.path.isfile(file_path):    #判断是否为文件
            file_list.append(file_path)

paths = u'F:\TEST'
file_list = []
find_prefix_of_path(paths, file_list)

print(file_list)