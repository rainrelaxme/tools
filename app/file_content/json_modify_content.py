import json


# 将扁平化的键值对合并到嵌套的字典结构中
def merge_flatten_into_nested(flat_data, nested_data):
    for key, value in flat_data.items():
        keys = key.split('.')  # 按点号分割键
        current_level = nested_data
        for k in keys[:-1]:  # 遍历除最后一个键之外的所有键
            if k not in current_level:
                current_level[k] = {}  # 如果键不存在，创建一个空字典
            elif not isinstance(current_level[k], dict):
                # 如果当前键的值不是字典，则覆盖为一个空字典
                current_level[k] = {}
            current_level = current_level[k]  # 进入下一层
        current_level[keys[-1]] = value  # 最后一个键赋值
    return nested_data


# 读取JSON文件
def read_json(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        data = json.load(file)
    return data


# 写入JSON文件
def write_json(file_path, data):
    with open(file_path, 'w', encoding='utf-8') as file:
        json.dump(data, file, ensure_ascii=False, indent=4)


# 主函数
def main(input_file, output_file):
    # 读取JSON文件
    data = read_json(input_file)

    # 分离嵌套数据和扁平化数据
    nested_data = {}
    flat_data = {}
    for key, value in data.items():
        if '.' in key:
            flat_data[key] = value  # 扁平化数据
        else:
            nested_data[key] = value  # 嵌套数据

    # 将扁平化数据合并到嵌套数据中
    merged_data = merge_flatten_into_nested(flat_data, nested_data)

    # 将合并后的JSON写入输出文件
    write_json(output_file, merged_data)

    print(f"JSON 文件已合并并写入 {output_file}")


# 示例用法
if __name__ == "__main__":
    input_file = 'input_data.json'  # 输入的JSON文件
    output_file = 'merged_data.json'  # 输出的合并后JSON文件
    main(input_file, output_file)