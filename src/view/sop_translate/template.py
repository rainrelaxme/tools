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

from docx.enum.table import WD_TABLE_ALIGNMENT


def apply_cover_template(content_data, cover_data):
    """应用封面模板"""
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
    para_2['index'] = 0
    para_2['element_index'] = 0
    new_cover_data.append(para_2)

    # 4. 文件编号
    for index, item in enumerate(cover_data):
        if item['flag'] == 'top_title':
            title_end_index = index
    process_cover_data = cover_data[title_end_index:]
    for index, item in enumerate(process_cover_data):
        pass


    return new_cover_data

