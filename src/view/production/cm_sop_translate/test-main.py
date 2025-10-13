#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : test-main.py 
@Author  : Shawn
@Date    : 2025/10/11 9:31 
@Info    : Description of this file

"""
import datetime
import os

from docx import Document

from src.view.production.cm_sop_translate.sop_translate import DocContent, add_cover_translation, \
    add_paragraph_translation, \
    add_table_translation, create_new_document, add_header_translation, add_footer_translation
from src.view.production.cm_sop_translate.template import apply_cover_template
from src.view.production.cm_sop_translate.translate_by_deepseek import Translator

if __name__ == "__main__":
    # 获取当前时间戳
    current_time = datetime.datetime.now().strftime('%y%m%d%H%M%S')
    language = ['英语', '越南语']

    input_file = r"D:\Code\Project\tools\data\test\test.docx"
    output_folder = r"D:\Code\Project\tools\data\temp"
    file_base_name = os.path.basename(input_file)
    output_file = output_folder + "/" + file_base_name.replace(".docx", f"_translate_{current_time}.docx")

    translator = Translator()
    print(f"********************start at {current_time}********************")
    # 1. 读取原文档内容
    doc = Document(input_file)
    new_doc = DocContent()
    # 正文主题
    content_data = new_doc.get_content(doc)
    # 页眉、页脚
    header_data = new_doc.get_header_content(doc)
    footer_data = new_doc.get_footer_content(doc)

    # 2. 标注关键信息：大标题、头信息，同时修改content_data内容
    new_doc.flag_title(content_data)
    new_doc.flag_preamble(content_data)
    new_doc.flag_approveTable(content_data)

    # 3. 获取封面信息
    cover_data = new_doc.get_cover_content(content_data)

    # 4. 翻译
    # ① 翻译封面
    translated_cover_data = add_cover_translation(cover_data, translator, language)
    # ② 翻译正文内容
    body_data = content_data[len(cover_data):]
    after_para = add_paragraph_translation(body_data, translator, language)
    after_table = add_table_translation(after_para, translator, language)
    # ③ 翻译页眉、页脚
    after_header = add_header_translation(header_data, translator, language)
    after_footer = add_footer_translation(footer_data, translator, language)

    # ④ 合并
    translated_content_data = []
    for item in translated_cover_data:
        translated_content_data.append(item)
    for item in after_table:
        translated_content_data.append(item)
    for item in after_header:
        translated_content_data.append(item)
    for item in after_footer:
        translated_content_data.append(item)

    # 5. 处理内容
    formatted_content = apply_cover_template(translated_content_data, translated_cover_data)

    # 6. 创建新文档
    create_new_document(formatted_content, output_file)

    print(f"********************end********************")