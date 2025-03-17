import pandas as pd


def replace_chinese_with_english(txt_file, excel_file, output_excel, output_txt):
    # 读取 Excel 文件
    df = pd.read_excel(excel_file)

    # 将 Excel 的行号、中文内容、英文内容存储为字典，方便查找
    translation_dict = {row[0]: (row[1], row[2]) for row in df.itertuples(index=False, name=None)}

    # 存储替换结果的列表
    results = []

    # 打开 TXT 文件并读取内容
    with open(txt_file, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    # 用于保存新的 TXT 文件内容
    new_lines = []

    # 遍历 TXT 文件中的每一行
    for line_num, line in enumerate(lines, start=1):
        original_line = line.strip()  # 去掉多余的空白字符
        if line_num in translation_dict:
            chinese_content, english_content = translation_dict[line_num]
            # 替换中文内容为英文
            replaced_line = line.replace(chinese_content, english_content)
            results.append([line_num, chinese_content, english_content, replaced_line.strip()])
            new_lines.append(replaced_line)  # 将替换后的行添加到新的文件内容
        else:
            # 没有匹配的行号时，保留原始行
            results.append([line_num, "", "", original_line.strip()])
            new_lines.append(line)  # 保持原始行内容

    # 将替换结果保存到新的 Excel 文件
    result_df = pd.DataFrame(results, columns=["Line Number", "Original Chinese", "English Content", "Replaced Line"])
    result_df.to_excel(output_excel, index=False, engine='openpyxl')
    print(f"替换结果已保存到 {output_excel}")

    # 将新的 TXT 文件保存
    with open(output_txt, 'w', encoding='utf-8') as file:
        file.writelines(new_lines)  # 将所有行写入新的 TXT 文件
    print(f"新的 TXT 文件已保存到 {output_txt}")


# 示例文件路径
txt_file = 'input.txt'  # 你的 TXT 文件路径
excel_file = 'output.xlsx'  # 包含行号、中文和英文的 Excel 文件路径
output_excel = 'replaced_results.xlsx'  # 输出的 Excel 文件路径
output_txt = 'output.txt'  # 输出的新的 TXT 文件路径

# 调用函数
replace_chinese_with_english(txt_file, excel_file, output_excel, output_txt)
