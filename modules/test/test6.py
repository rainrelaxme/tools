# import openpyxl
#
# # 创建一个新的Excel工作簿
# workbook = openpyxl.Workbook()
#
# # 获取默认的工作表
# sheet = workbook.active
#
# # 写入数据
# sheet['A1'] = '姓名'
# sheet['B1'] = '年龄'
#
# # 添加一行数据
# sheet.append(['Alice', 25])
#
# # 保存工作簿
# workbook.save('example.xlsx')
import copy

a = [0, [1, 2, {"text": 'origin'}]]
# b = a.copy()
b = copy.deepcopy(a)
# a[1][2]["text"] = 'new_origin'
b[1][2]["text"] = 'new'

print(a)
print(b)
