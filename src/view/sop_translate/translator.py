#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : translator.py 
@Author  : Shawn
@Date    : 2025/9/25 11:13 
@Info    : Description of this file
"""

import datetime
import os

from sop_translate import login, check_license, get_content, add_paragraph_translation, add_table_translation, \
    create_new_document, doc_to_docx
from translate_by_deepseek import Translator


def text_translate(language):
    original_text = input("请输入待翻译内容：\n").strip()
    translator = Translator()
    for lang in language:
        print(f"{lang}: {translator.translate(original_text, language=lang)}")


def docx_translate(language):
    # 设置输入文件夹路径
    input_folder = input("请输入待翻译的文件目录：\n")

    # 设置输出文件夹路径（二级文件夹）
    output_folder = os.path.join(input_folder, "translate_output")
    os.makedirs(output_folder, exist_ok=True)

    # 初始化翻译器
    translator = Translator()

    # 遍历文件夹中的所有文件
    for filename in os.listdir(input_folder):
        print(f"正在处理文件: {filename}")

        # 检查文件是否为Word文档（.docx格式）,若是.doc文件则转为.docx
        if filename.endswith('.doc') and not filename.startswith('~$'):  # 忽略临时文件
            doc_path = os.path.join(input_folder, filename)
            filename = os.path.basename(doc_to_docx(doc_path))

        if filename.endswith('.docx') and not filename.startswith('~$'):
            input_file = os.path.join(input_folder, filename)

            # 生成输出文件名（原文件名+时间）
            current_time = datetime.datetime.now().strftime('%y%m%d%H%M%S')
            file_base_name = os.path.splitext(filename)[0]  # 去掉扩展名
            output_filename = f"{file_base_name}_translate_{current_time}.docx"
            output_file = os.path.join(output_folder, output_filename)

            # 1. 读取原文档内容
            content_data = get_content(input_file)

            # 2. 翻译
            after_para = add_paragraph_translation(content_data, translator, language)
            after_table = add_table_translation(after_para, translator, language)

            # 3. 创建新文档
            create_new_document(after_table, output_file)
            print(f"  已完成翻译: {output_filename}")

    print(f"所有文件处理完成！输出目录: {output_folder}")


if __name__ == '__main__':

    print("\n" + "=" * 100)
    print("\x20" * 45 + f"文档翻译系统")
    print("=" * 100)

    # 首先进行登录验证
    # if not login():
    #     exit()

    # 可选：进行许可证检查
    if not check_license():
        exit()

    try:
        print("注意：仅支持运行于windows系统！！！")
        print("请选择目标语言（可多选，用\",\"分隔）：\n"
              "1. 英语\n"
              "2. 越南语")
        lang_input = input().strip()
        lang_list = lang_input.split(',')
        language = []
        for lang in lang_list:
            if lang.strip() == '1':
                language.append('英语')
            if lang.strip() == '2':
                language.append('越南语')
        print(f"您选择的目标语言为{language}")

        print("请选择使用方式：\n"
              "1. 文本翻译\n"
              "2. 文档翻译")
        option = input().strip()
        if option == '1':
            text_translate(language)
        elif option == '2':
            docx_translate(language)
        else:
            print("输入错误，请重新输入")

    except Exception as e:
        print(f"处理过程中出现错误: {e}")

    print(f"\n********************end**********************")
