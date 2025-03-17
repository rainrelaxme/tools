import json
import pandas as pd


def extract_nested_json(json_data, parent_key=''):
    """
    递归函数，提取最底层的键值对
    :param json_data: JSON 数据
    :param parent_key: 当前层级的父键
    :return: 扁平化的键值对列表
    """
    items = []

    # 如果json_data是字典类型，递归遍历每个键
    if isinstance(json_data, dict):
        for key, value in json_data.items():
            new_key = f"{parent_key}.{key}" if parent_key else key
            items.extend(extract_nested_json(value, new_key))

    # 如果json_data是列表类型，遍历列表中的每个元素
    elif isinstance(json_data, list):
        for i, item in enumerate(json_data):
            new_key = f"{parent_key}[{i}]"
            items.extend(extract_nested_json(item, new_key))

    # 如果是基础数据类型，返回当前键值对
    else:
        items.append((parent_key, json_data))

    return items


def json_file_to_excel(json_file, output_file):
    # 读取JSON文件
    with open(json_file, 'r', encoding='utf-8') as file:
        json_data = json.load(file)

    # 扁平化 JSON 数据，提取最底层的键值对
    flattened_data = extract_nested_json(json_data)

    # 使用Pandas创建DataFrame
    df = pd.DataFrame(flattened_data, columns=['Key', 'Value'])

    # 将DataFrame导出到Excel文件
    df.to_excel(output_file, index=False)


# 示例输入和输出文件
json_file = 'zh_cn.json'  # 你的JSON文件路径
output_file = 'output.xlsx'  # 输出的Excel文件路径

# 调用函数
json_file_to_excel(json_file, output_file)
print(f"数据已经成功保存到 {output_file}")
