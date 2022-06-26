#获取文件名-切分后缀名-添加空格-拼接-修改名称

import os

def rename_add_1(path):
    file_list = os.listdir(path)
    for file in file_list:
        absolute_file = path + "/" + file
        qiekaihouzhui = os.path.splitext(file)[0]
        print(qiekaihouzhui)
        old_name = absolute_file

"""
 if os.path.isfile(absolute_file):
    if file.endswith(".py") or file.endswith(".exe") or file.endswith(".txt"):
        continue
    old_name = absolute_file
    new_name = old_name + "1"
    os.rename(old_name, new_name)
    print("新文件名：", new_name)
else:

# 递归文件夹修改文件名
rename_add_1(absolute_file)
"""


def rename_sub_1(path):
    file_list = os.listdir(path)
    for file in file_list:
        absolute_file = path + "/" + file
        if os.path.isfile(absolute_file):
            if not file.endswith("1"):
                continue
            old_name = absolute_file
            new_name = old_name.strip('1')
            os.rename(old_name, new_name)
            print("新文件名：", new_name)
        else:
            # 递归文件夹修改文件名
            rename_sub_1(absolute_file)

if __name__ == '__main__':
    # 获取当前程序所在的文件路径
    path = os.getcwd()

    # 给当前路径下的所有文件名后边都+1
    rename_add_1(path)

    # 把当前路径下的所有文件名后边的1去掉
    rename_sub_1(path)