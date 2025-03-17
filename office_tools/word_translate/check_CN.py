import re
import pandas as pd


def extract_chinese_from_file(input_file, output_file):
    # 正则表达式用于匹配中文字符
    chinese_pattern = re.compile(r'[\u4e00-\u9fff]+')

    # 用于存储行号和中文内容
    data = []

    with open(input_file, 'r', encoding='utf-8') as file:
        # 逐行读取文件
        for line_num, line in enumerate(file, start=1):
            # 提取每行中的所有中文字符
            chinese_content = chinese_pattern.findall(line)
            if chinese_content:
                # 将行号和中文内容加入数据列表
                data.append([line_num, ' '.join(chinese_content)])

    # 将数据存储到 DataFrame 中
    df = pd.DataFrame(data, columns=['Line Number', 'Chinese Content'])

    # 保存到 Excel 文件
    df.to_excel(output_file, index=False, engine='openpyxl')


# 示例输入和输出文件
input_file = 'output.txt'  # 你的文本文件路径
output_file = 'check-output.xlsx'  # 输出的 Excel 文件路径

# 调用函数
extract_chinese_from_file(input_file, output_file)
print(f"中文内容已经成功提取并保存到 {output_file}")
