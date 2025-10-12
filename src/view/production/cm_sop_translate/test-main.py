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

from src.view.professional_project.sop_translate.sop_translate import DocContent, add_cover_translation, add_paragraph_translation, \
    add_table_translation, create_new_document
from src.view.professional_project.sop_translate.template import apply_cover_template
from src.view.professional_project.sop_translate.translate_by_deepseek import Translator

if __name__ == "__main__":
    # 获取当前时间戳
    current_time = datetime.datetime.now().strftime('%y%m%d%H%M%S')
    language = ['英语', '越南语']

    input_file = r"/data/test.docx"
    # input_file = r"D:\Code\Project\tools\data\1.C2LG-001-000-A08 供应商管理程序.docx"
    # input_file = r"D:\Code\Project\tools\data\13.C2GM-Z13-000-A00 管理评审程序.docx"
    # input_file = r"F:\Code\Project\tools\data\13.C2GM-Z13-000-A00 管理评审程序.docx"
    # input_file = r"F:\Code\Project\tools\data\1.C2LG-001-000-A08 供应商管理程序.docx"
    # input_file = r"D:\Code\Project\tools\data\13. 封面模板.docx"

    output_folder = r"D:\Code\Project\tools\data\temp"
    # output_folder = r"F:\Code\Project\tools\data\temp"

    file_base_name = os.path.basename(input_file)
    output_file = output_folder + "/" + file_base_name.replace(".docx", f"_translate_{current_time}.docx")

    translator = Translator()
    print(f"********************start at {current_time}********************")
    # 1. 读取原文档内容
    doc = Document(input_file)
    new_doc = DocContent()
    content_data = new_doc.get_content(doc)

    header_content = new_doc.get_header_content(doc)
    footer_content = new_doc.get_footer_content(doc)

    # 2. 标注关键信息：大标题、头信息，同时修改content_data内容
    new_doc.flag_title(content_data)
    new_doc.flag_preamble(content_data)
    new_doc.flag_approveTable(content_data)

    # 3. 获取封面信息
    cover_data = new_doc.get_cover_content(content_data)

    # 4. 翻译
    # ① 翻译封面
    translated_cover_data = add_cover_translation(cover_data, translator, language)
    # ② 翻译其它内容
    body_data = content_data[len(cover_data):]
    after_para = add_paragraph_translation(body_data, translator, language)
    after_table = add_table_translation(after_para, translator, language)
    # ③ 合并
    translated_content_data = []
    for item in translated_cover_data:
        translated_content_data.append(item)
    for item in after_table:
        translated_content_data.append(item)

    # 5. 处理内容
    formatted_content = apply_cover_template(translated_content_data, translated_cover_data)

    # 6. 创建新文档
    create_new_document(formatted_content, output_file)

    print(f"********************end********************")