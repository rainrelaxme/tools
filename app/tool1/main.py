# 提取json文件的key和value

import json
import pandas as pd


# 递归提取嵌套的键和值
def flatten_json(y):
    out = {}

    def flatten(x, name=''):
        if type(x) is dict:
            for a in x:
                flatten(x[a], name + a + '.')
        elif type(x) is list:
            i = 0
            for a in x:
                flatten(a, name + str(i) + '.')
                i += 1
        else:
            out[name[:-1]] = x

    flatten(y)
    return out


# 读取JSON文件
def read_json(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        data = json.load(file)
    return data


# 将JSON数据转换为DataFrame
def json_to_dataframe(data):
    flattened_data = flatten_json(data)
    df = pd.DataFrame(list(flattened_data.items()), columns=['Key', 'Value'])
    return df


# 将DataFrame写入Excel文件
def write_to_excel(df, excel_file):
    df.to_excel(excel_file, index=False)


# 主函数
def main(json_file, excel_file):
    # 读取JSON文件
    data = read_json(json_file)

    # 将JSON数据转换为DataFrame
    df = json_to_dataframe(data)

    # 将DataFrame写入Excel文件
    write_to_excel(df, excel_file)

    print(f"数据已成功写入 {excel_file}")


# 示例用法
if __name__ == "__main__":
    json_file = 'data.json'  # 替换为你的JSON文件路径
    excel_file = 'output.xlsx'  # 替换为你想保存的Excel文件路径
    main(json_file, excel_file)