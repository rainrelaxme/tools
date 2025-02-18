import re
import pandas as pd


def extract_chinese_with_symbols(txt_file, output_excel):
    # 扩展正则表达式，匹配中文字符和常见中文标点符号
    chinese_pattern = re.compile(r'[\u4e00-\u9fff，。！？、；：“”‘’（）【】]')

    # 用于存储行号和中文内容
    data = []

    with open(txt_file, 'r', encoding='utf-8') as file:
        # 逐行读取文件
        for line_num, line in enumerate(file, start=1):
            # 提取每行中的所有中文字符和符号
            chinese_content = ''.join(chinese_pattern.findall(line))
            if chinese_content:  # 如果找到中文内容，添加到数据列表
                data.append([line_num, chinese_content])

    # 将数据存储到 DataFrame 中
    df = pd.DataFrame(data, columns=['Line Number', 'Chinese Content'])

    # 保存到 Excel 文件
    df.to_excel(output_excel, index=False, engine='openpyxl')
    print(f"中文内容已经成功提取并保存到 {output_excel}")


# 示例输入和输出文件
txt_file = 'input.txt'  # 你的文本文件路径
output_excel = 'output.xlsx'  # 输出的 Excel 文件路径

# 调用函数
extract_chinese_with_symbols(txt_file, output_excel)
