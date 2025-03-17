# -*- coding:utf-8 -*-
# authored by rainrelaxme
# 说明：工具的调用

from office_tools import file_edit as fe
from office_tools import file_function as ff
from office_tools import excel_edit as ee

if __name__ == '__main__':
    print('请选择工具（输入序号，Enter确认，输入q退出）：\n'
          '1.当前文件的路径\n'
          '2.文件夹下的文件清单\n'
          '3.文件夹下的文件增加空格\n'
          '4.文件夹下的文件去除空格\n'
          '5.文件夹下的文件增加自定义字符\n'
          '6.文件夹下的文件去除自定义字符\n'
          '7.文件夹下的Excel合并\n'
          )

    while True:
        tool_choose = input('>>')
        if tool_choose == "q":
            break
        elif tool_choose == "1":
            result = fe.get_filepath()
            print(result)
        elif tool_choose == "2":
            path = input("请输入文件夹路径：")
            if path == "q":
                break
            else:
                result = fe.get_filelist(path)
            print(result)
        elif tool_choose == "3":
            path = input("请输入文件夹路径：")
            if path == "q":
                break
            else:
                ff.rename_add_space(path)
            print("已完成，请前往查看结果。")
        elif tool_choose == "4":
            path = input("请输入文件夹路径：")
            if path == "q":
                break
            else:
                ff.rename_sub_space(path)
            print("已完成，请前往查看结果。")
        elif tool_choose == "5":
            path = input("请输入文件夹路径：")
            x = input("请输入要增加的字符（禁止输入q）：")
            if path == "q" or x == "q":
                break
            else:
                ff.rename_add_x(path, x)
            print("已完成，请前往查看结果。")
        elif tool_choose == "6":
            path = input("请输入文件夹路径：")
            x = input("请输入要去除的字符（禁止输入q）：")
            if path == "q" or x == "q":
                break
            else:
                ff.rename_sub_x(path, x)
            print("已完成，请前往查看结果。")
        elif tool_choose == "7":
            print("请选择文件：")
            files = ee.choose_file(is_single=1)
            ff.excel_merge(files)
            print("已完成，请前往查看结果。")
        else:
            print('选择错误，请重新选择！')
