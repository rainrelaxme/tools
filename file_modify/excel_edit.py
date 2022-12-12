# -*- coding:utf-8 -*-
# authored by rainrelaxme
# 操作excel文件

# 导入需要使用的包
import xlrd  # 读取Excel文件的包
import xlsxwriter  # 将文件写入Excel的包


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
