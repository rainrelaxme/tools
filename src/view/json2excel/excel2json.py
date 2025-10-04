import pandas as pd
import json


def set_nested_value(d, keys, value):
    """
    将嵌套的键值对插入字典
    :param d: 当前字典
    :param keys: 键的层级列表
    :param value: 最终要插入的值
    """
    key = keys.pop(0)

    # 如果是最后一层，直接赋值
    if len(keys) == 0:
        d[key] = value
    else:
        # 如果字典中没有这个键，则创建一个空字典
        if key not in d:
            d[key] = {}
        # 递归设置下一层级
        set_nested_value(d[key], keys, value)


def excel_to_json(excel_file, output_file):
    # 读取 Excel 文件
    df = pd.read_excel(excel_file)

    # 假设第一列是键，第二列是值
    json_data = {}

    for _, row in df.iterrows():
        key = row[0]  # 第一列为键
        value = row[1]  # 第二列为值

        # 将键按 "." 分割，形成层级结构
        keys = key.split(".")

        # 将值插入到正确的层级
        set_nested_value(json_data, keys, value)

    # 保存为 JSON 文件
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, ensure_ascii=False, indent=4)


# 示例输入和输出文件
excel_file = 'output - yinni.xlsx'  # 你的 Excel 文件路径
output_file = 'out/output - yinni.json'  # 输出的 JSON 文件路径

# 调用函数
excel_to_json(excel_file, output_file)
print(f"数据已经成功保存到 {output_file}")
