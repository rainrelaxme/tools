#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : template.py 
@Author  : Shawn
@Date    : 2025/9/30 14:15 
@Info    : Description of this file
"""

import datetime
import json
import os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_TAB_ALIGNMENT, WD_LINE_SPACING, WD_PARAGRAPH_ALIGNMENT, WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, Cm, RGBColor


def apply_template(body_data, header_data=None, footer_data=None, cover_data=None):
    """应用模板内容，修改data"""
    after_header = apply_header_template(header_data)
    after_footer = apply_footer_template(footer_data)
    after_cover = apply_cover_template(cover_data)
    new_data = {
        "header": after_header,
        "footer": after_footer,
        "cover": after_cover,
        "body": body_data,
    }
    return new_data


def apply_header_template(header_data=None):
    """应用页眉模板，修改data"""
    # 如果没有此内容，则不应用模板。
    if header_data is None:
        return None

    # start
    return header_data


def apply_footer_template(footer_data=None):
    """应用页眉模板，修改data"""
    # 如果没有此内容，则不应用模板。
    if footer_data is None:
        return None
    # start
    return footer_data


def apply_cover_template(cover_data=None):
    """应用封面模板，修改data"""
    # 如果没有此内容，则不应用模板。
    if cover_data is None:
        return None

    new_cover_data = []
    para_data = {
        'type': 'paragraph',
        'index': None,
        'element_index': None,
        'flag': '',
        'text': '',
        'para_format': {
            'style': 'Normal',
            'alignment': None,
            'space_before': None,
            'space_after': None,
            'line_spacing': 2.0,
            'first_line_indent': None,
            'left_indent': None
        },
        'runs': []
    }
    table_data = {
        'type': 'table',
        'index': None,
        'element_index': None,
        'flag': None,
        'table_alignment': WD_TABLE_ALIGNMENT.CENTER,  # 表格居中，非内容居中
        'rows': None,
        'cols': None,
    }

    # 1. 第一行空格
    para_1 = para_data.copy()
    para_1['index'] = 0
    para_1['element_index'] = 0
    new_cover_data.append(para_1)

    # 2. 大标题
    table_1 = table_data.copy()
    table_1['index'] = 1
    table_1['element_index'] = 0
    table_1['flag'] = 'top_title'
    table_1['cols'] = 1
    rows = []
    for item in cover_data:
        if item['flag'] == 'top_title':
            row = {
                'cells': [
                    {
                        'row': 0,
                        'col': 0,
                        'grid_span': 1,
                        'is_merge_start': False,
                        'text': item['text'],
                        'paragraphs': [
                            {
                                'para_index': 0,
                                'text': item['text'],
                                'para_format': item['para_format'],
                                'runs': item['runs'] if 'runs' in item else []
                            }
                        ]
                    }
                ]
            }
            rows.append(row)
    new_rows = [
        {
            'cells': [
                {
                    'row': rows[0]['cells'][0]['row'],
                    'col': rows[0]['cells'][0]['col'],
                    'grid_span': rows[0]['cells'][0]['grid_span'],
                    'is_merge_start': rows[0]['cells'][0]['is_merge_start'],
                    'text': rows[0]['cells'][0]['text'],
                    'paragraphs': []
                }

            ]
        }
    ]
    for item in rows:
        new_rows[0]['cells'][0]['paragraphs'].append(item['cells'][0]['paragraphs'][0])
    table_1['rows'] = new_rows
    new_cover_data.append(table_1)

    # 3. 接空行
    para_2 = para_data.copy()
    para_2['index'] = 2
    para_2['element_index'] = 1
    new_cover_data.append(para_2)

    # 4. 文件头信息：文档编号等
    index = len(new_cover_data)
    para_index = 2
    i = 0
    for item in cover_data:
        if item['flag'] == 'preamble':
            item['index'] = index
            item['element_index'] = para_index
            new_cover_data.append(item)
            index += 1
            para_index += 1
            i += 1
            if i < 3:  # 翻译语言的种类
                continue

            # 加一个空行
            new_blank = para_data.copy()
            new_blank['index'] = index + 1
            new_blank['element_index'] = para_index + 1
            new_cover_data.append(new_blank)
            index += 1
            para_index += 1
            i = 0

    # 5. 审批表格
    for item in cover_data:
        if item['type'] == 'table':
            item['index'] = len(new_cover_data) + 1
            item['element_index'] = 1
            for row in item['rows']:
                for cell in row['cells']:
                    for para in cell['paragraphs']:
                        para['para_format']['line_spacing'] = Pt(20)
            new_cover_data.append(item)

    return new_cover_data


def apply_preamble_format(paragraph, preamble_data):
    """
    应用文件头信息格式，通过制表符对齐，生成的格式
    """
    # 清除默认制表位
    paragraph.paragraph_format.tab_stops.clear_all()

    # 固定行距20磅
    paragraph_format = paragraph.paragraph_format
    paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY  # 固定行距
    paragraph_format.line_spacing = Pt(20)
    paragraph_format.space_after = 0

    # 添加自定义制表位
    tab_stop1 = paragraph.paragraph_format.tab_stops.add_tab_stop(
        Inches(1.7),
        WD_TAB_ALIGNMENT.LEFT
    )
    tab_stop2 = paragraph.paragraph_format.tab_stops.add_tab_stop(
        Inches(3.5),
        WD_TAB_ALIGNMENT.LEFT
    )

    # 使用制表位
    split_text = preamble_data['text'].split('：')
    paragraph.add_run("\t")  # 制表符1
    run1 = paragraph.add_run(split_text[0] + '：')
    run1.font.bold = True
    run1.font.size = Pt(16.0)
    run1.font.name = u'宋体'
    run1.element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')  # 设置中文字体

    paragraph.add_run("\t")  # 制表符2
    run2 = paragraph.add_run(split_text[1].strip())
    run2.font.size = Pt(16.0)
    run2.font.name = 'Times New Roman'  # 设置西文字体
    run2.element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')  # 设置中文字体


def apply_approveTable_format(table):
    """
    应用审批表格格式，生成的格式
    """
    for i in range(len(table.rows)):
        table.rows[i].height = Cm(2.5)
    for i in range(len(table.columns)):
        table.columns[i].width = Cm(2.5)


def apply_header_format(doc, header_data):
    """应用页眉，生成的格式"""
    # 设置奇偶页是否相同
    doc.settings.odd_and_even_pages_header_footer = False

    for index, item in enumerate(header_data):
        # 创建节
        section = doc.sections[item['section_index']]

        # 设置首页是否相同
        section.different_first_page_header_footer = True
        # 填充内容
        if item['first_page_header_content']:
            add_content(section.first_page_header, item['first_page_header_content'])
        if item['odd_page_header']:
            add_content(section.header, item['odd_page_header'])


def apply_footer_format(doc, footer_data):
    """应用页脚，生成的格式"""
    # 设置奇偶页是否相同
    doc.settings.odd_and_even_pages_header_footer = False

    for index, item in enumerate(footer_data):
        # 创建节
        section = doc.sections[item['section_index']]

        # 设置首页页脚不同
        section.different_first_page_header_footer = True
        # 填充内容
        if item['first_page_footer_content']:
            add_content(section.first_page_footer, item['first_page_footer_content'])
        if item['odd_page_footer']:
            add_content(section.footer, item['odd_page_footer'])


def apply_paragraph_format(paragraph, format_info):
    """应用段落格式"""
    # 使用安全的get方法访问字典，避免KeyError
    if format_info.get('alignment'):
        alignment_map = {
            'CENTER': WD_ALIGN_PARAGRAPH.CENTER,
            'CENTER (1)': WD_ALIGN_PARAGRAPH.CENTER,
            'RIGHT': WD_ALIGN_PARAGRAPH.RIGHT,
            'JUSTIFY': WD_ALIGN_PARAGRAPH.JUSTIFY,
            'LEFT': WD_ALIGN_PARAGRAPH.LEFT
        }
        alignment = format_info['alignment'].split('.')[-1] if '.' in format_info['alignment'] else format_info[
            'alignment']
        paragraph.alignment = alignment_map.get(alignment, WD_ALIGN_PARAGRAPH.LEFT)

    # 使用get方法安全访问，如果键不存在则返回None
    space_before = format_info.get('space_before')
    space_after = format_info.get('space_after')
    line_spacing = format_info.get('line_spacing')
    first_line_indent = format_info.get('first_line_indent')
    left_indent = format_info.get('left_indent')

    # if space_before is not None:
    #     paragraph.paragraph_format.space_before = space_before
    if space_before:
        paragraph.paragraph_format.space_before = space_before

    if space_after:
        paragraph.paragraph_format.space_after = space_after
    else:
        paragraph.paragraph_format.space_after = 0  # 段落后间距None会默认设置为10，需设置为0

    if line_spacing:
        paragraph.paragraph_format.line_spacing = line_spacing
    if first_line_indent:
        paragraph.paragraph_format.first_line_indent = first_line_indent
    if left_indent:
        paragraph.paragraph_format.left_indent = left_indent


def apply_table_format(table, table_format_info):
    """
    处理表格样式
    """
    # 自动调整列宽：禁用
    table.autofit = False

    # 表格对齐方式，非内容对齐方式
    table.alignment = table_format_info['table_alignment']

    # 单元格样式
    for row in table_format_info['rows']:
        for cell in row['cells']:
            row = cell['row']
            col = cell['col']
            # 设置表格宽度
            if hasattr(cell, 'width'):
                table.cell(row, col).width = Inches(cell['width'])

            # 合并单元格
            if cell['is_merge_start']:
                row = cell['row']
                col = cell['col']
                # 处理水平合并（gridSpan）
                if cell['grid_span'] > 1:
                    # 合并水平单元格
                    start_cell = table.rows[row].cells[col]
                    end_cell = table.rows[row].cells[col + cell['grid_span'] - 1]
                    start_cell.merge(end_cell)
    # if table_format_info['flag'] == 'top_title':
    #     table.autofit = True


def apply_run_format(run, run_data):
    """应用运行文本格式"""
    # 使用get方法安全访问
    run.bold = run_data.get('bold') if run_data.get('bold') else False
    run.italic = run_data.get('italic') if run_data.get('italic') else False
    run.underline = run_data.get('underline') if run_data.get('underline') else False

    font_size = run_data.get('font_size')
    if font_size:
        run.font.size = Pt(font_size)

    font_name = run_data.get('font_name')
    if font_name:
        run.font.name = font_name
        run.element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    else:
        # 无font_name时采用默认字体
        if run.text.strip():
            run.font.name = 'Times New Roman'  # 设置西文字体
            run.element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')  # 设置中文字体
        else:
            run.font.name = u'宋体'
            run.element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')  # 设置中文字体

    font_color = run_data.get('font_color')
    if font_color:
        try:
            run.font.color.rgb = RGBColor(*font_color)
        except:
            pass


def add_content(block, data):
    """对块（doc、section、header、footer）中添加内容"""
    for index, item in enumerate(data):
        if item['type'] == 'paragraph':
            # 创建新段落
            paragraph = block.add_paragraph()
            # 应用段落格式
            apply_paragraph_format(paragraph, item['para_format'])
            # 应用运行文本、格式
            for run_data in item['runs']:
                run = paragraph.add_run(run_data['text'])
                apply_run_format(run, run_data)
        # if item['type'] == 'table':
        #     # 创建表格
        #     table = block.add_table(rows=len(item['rows']), cols=item['cols'])
        #     table.style = 'Table Grid'
        #
        #     # 设置表格样式
        #     apply_table_format(table, item)
        #
        #     # 应用表格内容
        #     for row_idx, row_data in enumerate(item['rows']):
        #         for cell_idx, cell_data in enumerate(row_data['cells']):
        #             if (cell_data['grid_span'] > 1 and cell_data['is_merge_start']) or cell_data['grid_span'] == 1:
        #                 cell = table.rows[row_idx].cells[cell_idx]
        #
        #                 # 清空默认段落
        #                 for paragraph in cell.paragraphs:
        #                     p = paragraph._element
        #                     p.getparent().remove(p)
        #
        #                 # 添加内容到单元格
        #                 if cell_data['paragraphs']:
        #                     for para_data in cell_data['paragraphs']:
        #                         cell_para = cell.add_paragraph()
        #                         apply_paragraph_format(cell_para, para_data['para_format'])
        #
        #                         for run_data in para_data['runs']:
        #                             run = cell_para.add_run(run_data['text'])
        #                             apply_run_format(run, run_data)
        #                 else:
        #                     # 如果没有详细的段落信息，只添加文本
        #                     cell.text = cell_data['text']

def main(data):
    current_time = datetime.datetime.now().strftime('%y%m%d%H%M%S')
    file_path = r"D:\Code\Project\tools\data\temp"
    file = os.path.join(file_path, f'test_{current_time}.docx')

    doc = Document()
    apply_footer_format(doc, data)
    # apply_header_format(doc, header_data)

    i = 1
    while i < 30:
        doc.add_paragraph().add_run(f"这是第{i}段")
        i += 1

    doc.save(file)
    print(f"{file} was saved successfully.")


if __name__ == '__main__':
    # doc.settings.odd_and_even_pages_header_footer = True
    #
    # for section in doc.sections:
    #     section.different_first_page_header_footer = True
    #     first_footer = section.first_page_footer
    #     first_footer.paragraphs[0].add_run("这是首页！")
    #     # first_footer.paragraphs[0].font.bold = True
    #     # first_footer.paragraphs[0].add_run().font.size = Pt(36.0)
    #     first_footer.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    file = r"D:\Code\Project\tools\data\data.json"
    with open(file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    main(data)
