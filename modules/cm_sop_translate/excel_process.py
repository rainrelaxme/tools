#!/usr/bin/env python3
# -*- coding:utf-8 -*-
"""
@Project: translate_tools.py
@File   : excel_process.py
@Version:
@Author : shawn
@Date   : 2025/12/1 13:00
@Info   : 实现excel文档的翻译，并用翻译后的文本替换原文本后，且保持原文档格式。
"""
import copy
import datetime
import io
import os
import time
from pathlib import Path
from typing import Dict, List, Optional

import openpyxl
from PIL import Image
from openpyxl import load_workbook
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Protection

from modules.cm_sop_translate.translator import Translator


def get_content(file_path: str) -> List[Dict[str, List[List[Optional[str]]]]]:
    """
    读取excel文件的内容，返回每个sheet的二维数组，保留原始顺序
    """
    workbook = load_workbook(file_path, data_only=False)
    data: List[Dict[str, List[List[Optional[str]]]]] = []

    # 设定最大的取值范围，防止空的单元格，但是被误编辑的被统计在内
    max_row = ''
    max_column = ''

    for sheet in workbook.worksheets:
        sheet_matrix: List[List[Optional[str]]] = []
        cells_matrix = []
        for row in sheet.iter_rows():
            row_values = []
            for cell in row:
                # row_values.append(cell.value)
                # 获取单元格格式-文字
                cell_info = {
                    "type": '',
                    "flag": '',
                    "is_merged": False,
                    "is_merge_start": False,
                    "coordinate": cell.coordinate,
                    "row": cell.row,
                    "col": cell.column,
                    "value": cell.value if cell.value else "",
                    "format": get_cell_format(cell)
                }
                # print(cell_info)
                cells_matrix.append(cell_info)
            # sheet_matrix.append(row_values)
        # 合并单元格信息
        merge_info = get_merge_info(sheet)
        sheet_data = {
            "sheet": sheet.title,
            "sheet_color": sheet.sheet_properties.tabColor.rgb if sheet.sheet_properties.tabColor else None,
            # "rows": sheet_matrix,
            # "cells": cells_matrix,
            "cells": update_merge_info(cells_matrix, merge_info),
            "merge_info": merge_info
        }
        data.append(sheet_data)

    return data


def get_row_format(file_path, row_number, sheet_name=None):
    """
    获取整行的格式信息
    """
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active

    row_format = {
        "row_height": ws.row_dimensions[row_number].height if row_number in ws.row_dimensions else None,
        "cells": {}
    }

    # 获取每列的格式
    max_column = ws.max_column
    for col in range(1, max_column + 1):
        cell = ws.cell(row=row_number, column=col)
        row_format["cells"][cell.column_letter] = get_cell_format(cell)

    return row_format


def get_column_format(file_path, column_letter, sheet_name=None):
    """
    获取整列的格式信息
    """
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active

    col_idx = ws[column_letter + "1"].column

    column_format = {
        "column_letter": column_letter,
        "column_width": ws.column_dimensions[column_letter].width if column_letter in ws.column_dimensions else None,
        "cells": {}
    }

    # 获取每行的格式
    max_row = ws.max_row
    for row in range(1, max_row + 1):
        cell = ws.cell(row=row, column=col_idx)
        column_format["cells"][row] = get_cell_format(cell)

    return column_format


def get_cell_format(cell):
    """
    获取单元格的格式内容
    """
    cell_format = {
        "value": cell.value,
        "data_type": type(cell.value).__name__,
        "number_format": cell.number_format,
        "font": {},
        "fill": {},
        "border": {},
        "alignment": {},
        "protection": {}
    }

    # 1. 获取字体格式
    if cell.font:
        font: Font = cell.font
        # 如果有颜色的属性，则可以获取到
        color = None
        color_type = None
        if font.color:
            color = font.color.rgb if font.color.type == "rgb" else font.color.theme
            color_type = font.color.type

        cell_format["font"] = {
            "name": font.name,
            "size": font.sz,
            "bold": font.b,
            "italic": font.i,
            "underline": font.u,
            "strike": font.strike,
            "color": color,
            "color_type": color_type
        }

    # 2. 获取填充格式（背景色）
    if cell.fill:
        fill: PatternFill = cell.fill
        cell_format["fill"] = {
            "fill_type": fill.fill_type,
            "start_color": fill.start_color.rgb if fill.start_color.type == "rgb" else fill.start_color.theme,
            "end_color": fill.end_color.rgb if fill.end_color.type == "rgb" else fill.end_color.theme,
            "fgColor": fill.fgColor.rgb if hasattr(fill.fgColor, 'rgb') else str(fill.fgColor),
            "bgColor": fill.bgColor.rgb if hasattr(fill.bgColor, 'rgb') else str(fill.bgColor)
        }

    # 3. 获取边框格式
    if cell.border:
        border: Border = cell.border
        cell_format["border"] = {
            "left": get_side_info(border.left),
            "right": get_side_info(border.right),
            "top": get_side_info(border.top),
            "bottom": get_side_info(border.bottom),
            "diagonal": get_side_info(border.diagonal),
            "diagonal_direction": border.diagonal_direction,
            "outline": border.outline  # 外边框
        }

    # 4. 获取对齐格式
    if cell.alignment:
        alignment: Alignment = cell.alignment
        cell_format["alignment"] = {
            "horizontal": alignment.horizontal,
            "vertical": alignment.vertical,
            "text_rotation": alignment.textRotation,
            "wrap_text": alignment.wrapText,
            "shrink_to_fit": alignment.shrinkToFit,
            "indent": alignment.indent
        }

    # 5. 获取保护格式
    if cell.protection:
        protection: Protection = cell.protection
        cell_format["protection"] = {
            "locked": protection.locked,
            "hidden": protection.hidden
        }

    return cell_format


def get_side_info(side: Side):
    """获取边框边的信息"""
    if side and side.style:
        return {
            "style": side.style,
            "color": side.color.rgb if side.color and side.color.type == "rgb" else side.color.theme if side.color else None
        }
    return None


def get_merge_info(sheet):
    """
    获取合并单元格信息
    """
    merge_lists = sheet.merged_cells
    # print("merge_list*******", merge_lists)

    merge_infos = []
    for merge_list in merge_lists:
        # 一种方法
        merge_coords = merge_list.coord

        # 另一种方法可以把所有合并的单元格返回
        # 获取合并单元格的起始行列，和终止行列
        row_min, row_max, col_min, col_max = merge_list.min_row, merge_list.max_row, merge_list.min_col, merge_list.max_col
        merge_start_end = [(row_min, col_min), (row_max, col_max)]
        # 定义合并单元格的类型：合并行列/合并行/合并列
        merge_type = ''
        row_col = []
        if row_min != row_max and col_min != col_max:  # 合并行列
            row_col = [(x, y) for x in range(row_min, row_max + 1) for y in range(col_min, col_max + 1)]
            merge_type = 'row_col_merge'
        elif row_min == row_max and col_min != col_max:  # 合并行
            row_col = [(row_min, y) for y in range(col_min, col_max + 1)]
            merge_type = 'row_merge'
        elif row_min != row_max and col_max == col_min:  # 合并行
            row_col = [(x, col_min) for x in range(row_min, row_max + 1)]
            merge_type = 'col_merge'

        merge_info = {
            "merge_coords": merge_coords,
            "merge_start_end": merge_start_end,  # 合并单元格的起止位置
            "merge_cells": row_col,  # 合并的单元格序号
            "merge_type": merge_type,  # 合并单元格的类型
            "merge_data": sheet.cell(row=merge_list.min_row, column=merge_list.min_col).value,  # 合并单元格的内容
        }
        merge_infos.append(merge_info)

    return merge_infos


def update_merge_info(data, merge_info):
    """
    标记单元格是否是合并单元格，合并单元格除左上角外禁止写入
    """
    new_data = copy.deepcopy(data)

    # 记录所有合并单元格的左上角的单元格list
    merge_start_list = []
    # 记录所有合并单元格的单元格list
    merge_cell_list = []
    for merge_block in merge_info:
        start_cell = merge_block["merge_coords"].split(":", 1)[0]
        merge_start_list.append(start_cell)
        merge_cell_list.extend(merge_block["merge_cells"])

    # 更改is_merged和is_merge_start的值
    for cell in new_data:
        # cell_pos = (cell['row'], cell['col'])
        if cell['coordinate'] in merge_start_list:
            cell['is_merge_start'] = True
            cell['is_merged'] = True
        elif (cell['row'], cell['col']) in merge_cell_list:
            cell['is_merge_start'] = False
            cell['is_merged'] = True

    return new_data


def extract_images_from_excel(file_path, output_dir='extracted_images'):
    """
    从Excel文件中提取图片
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    workbook = load_workbook(file_path)
    image_data = {}

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        image_data[sheet_name] = []

        # 获取工作表的所有图片
        for img in sheet._images:
            # 获取图片数据
            img_data = img._data()

            # 创建图片对象
            image = Image.open(io.BytesIO(img_data))

            # 生成唯一文件名
            import uuid
            img_filename = f"{sheet_name}_{uuid.uuid4().hex[:8]}.png"
            img_path = os.path.join(output_dir, img_filename)

            # 保存图片
            image.save(img_path)

            # 记录图片位置信息
            image_info = {
                'path': img_path,
                'anchor': img.anchor,
                'filename': img_filename
            }
            image_data[sheet_name].append(image_info)

    return image_data


def add_translation(data, translator, langs: list, method: str):
    """
    添加翻译段落
    method-replace: 替换原文字，只保留1种语言
    method-add: 增加新的内容，保留3种语言
    """
    new_data = []
    if method == 'replace':  # 替换成一种语言
        for lang in langs:
            lang_data = copy.deepcopy(data)  # 使用copy()会影响原始数据，deepcopy不会影响
            for sheet_data in lang_data:
                for cell_data in sheet_data['cells']:
                    original_text = str(cell_data['value'])
                    if original_text.strip():
                        translated_text = translator.translate(original_text, lang, display=True)
                        cell_data['value'] = translated_text
            new_data.append({
                "language": lang,
                "data": lang_data
            })

    elif method == 'replace_multi':  # 替换成多种语言
        lang_data = copy.deepcopy(data)
        for sheet_data in lang_data:
            for cell_data in sheet_data['cells']:
                original_text = str(cell_data['value'])
                if original_text.strip():
                    all_translated_text = original_text
                    for lang in langs:
                        translated_text = translator.translate(original_text, lang, display=True)
                        all_translated_text = all_translated_text + '\n' + translated_text
                    cell_data['value'] = all_translated_text
        new_data.append({
            "language": langs,
            "data": lang_data
        })

    elif method == 'add':
        pass

    return new_data


def apply_cell_format(cell, data):
    """
    应用单元格的样式
    """
    #  字体
    cell.font = Font(
        name=data['format']['font']['name'],
        sz=data['format']['font']['size'],
        b=data['format']['font']['bold'],
        i=data['format']['font']['italic'],
        u=data['format']['font']['underline'],
        strike=data['format']['font']['strike'],
        # color=data['format']['font']['color']
    )

    #  填充
    cell.fill = PatternFill(
        fill_type=data['format']['fill']['fill_type'],
        start_color=data['format']['fill']['start_color'],
        end_color=data['format']['fill']['end_color'],
        fgColor=data['format']['fill']['fgColor'],
        bgColor=data['format']['fill']['bgColor']
    )

    #  边框
    # 暂时未包含颜色，如color=data['format']['border']['left']['color']
    cell.border = Border(
        left=Side(
            style=data['format']['border']['left'].get('style') if data['format']['border'].get('left') else None),
        right=Side(
            style=data['format']['border']['right'].get('style') if data['format']['border'].get('right') else None),
        top=Side(style=data['format']['border']['top'].get('style') if data['format']['border'].get('top') else None),
        bottom=Side(
            style=data['format']['border']['bottom'].get('style') if data['format']['border'].get('bottom') else None),
        diagonal=Side(
            style=data['format']['border'].get('diagonal') if data['format']['border'].get('diagonal') else None),
        diagonal_direction=Side(
            style=data['format']['border'].get('diagonal_direction') if data['format']['border'].get(
                'diagonal_direction') else None),
        # outline=Side(style=data['format']['border']['outline'])
    )

    #  对齐
    cell.alignment = Alignment(
        horizontal=data['format']['alignment']['horizontal'],
        vertical=data['format']['alignment']['vertical'],
        text_rotation=data['format']['alignment']['text_rotation'],
        wrap_text=data['format']['alignment']['wrap_text'],
        shrink_to_fit=data['format']['alignment']['shrink_to_fit'],
        indent=data['format']['alignment']['indent']
    )

    #  数字格式
    cell.number_format = data['format']['number_format']


def insert_images_to_excel(workbook, image_data):
    """
    将图片插入到Excel工作簿中
    """
    for sheet_name, images in image_data.items():
        if sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]

            for img_info in images:
                try:
                    img = XLImage(img_info['path'])
                    # 设置图片位置（简化处理，实际可能需要更精确的位置计算）
                    sheet.add_image(img, img_info['anchor']._from._col)
                except Exception as e:
                    print(f"插入图片失败: {e}")


def create_new_excel(output_file: str, data, method, input_file=None):
    """
    根据内容创建新的excel文件
    方法-simple：直接替换value
    方法-complex：获取数据+格式等，再生成新的文件
    """
    if method == 'simple':
        """
        if len(data) > 1:
            for lang_data in data:
                lang = lang_data['language']

                # 读取原文件
                excel = openpyxl.load_workbook(input_file)

                # 方法1：合并单元格，先取消再合并，以便写入内容 - 会丢失合并单元格中的图片信息，主要是形状信息
                for sheet_data in lang_data['data']:
                    sheet = excel[sheet_data['sheet']]
                    # 获取合并单元格信息
                    merged_ranges = list(sheet.merged_cells.ranges)
                    # 取消合并
                    for merged_range in merged_ranges:
                        sheet.unmerge_cells(str(merged_range))
                    # 更换单元格内容
                    for cell in sheet_data['cells']:
                        sheet[cell['coordinate']] = cell['value']
                    # 合并单元格
                    for merged_range in merged_ranges:
                        sheet.merge_cells(str(merged_range))

                # 存储文件
                new_output_file = output_file.replace(".xlsx", f"_{lang}.xlsx")
                time.sleep(1)
                excel.save(new_output_file)
                print(f"输出文件：{new_output_file}")

        elif len(data) == 1:
            # 方法2： 找到合并单元格的位置信息，只更新允许写入的内容
            for lang_data in data:
                lang = lang_data['language']
                # 读取原文件
                excel = openpyxl.load_workbook(input_file)

                for sheet_data in lang_data['data']:
                    sheet = excel[sheet_data['sheet']]
                    # images = sheet._images
                    # 更换单元格内容
                    for cell in sheet_data['cells']:
                        if cell['is_merged'] and not cell['is_merge_start']:
                            continue
                        elif cell['value'] == '':
                            continue
                        else:
                            sheet[cell['coordinate']] = cell['value']
                # 存储文件
                new_output_file = output_file.replace(".xlsx", f"_{lang}.xlsx")
                time.sleep(1)
                excel.save(new_output_file)
                print(f"输出文件：{new_output_file}")
        """
        # 方法2： 找到合并单元格的位置信息，只更新允许写入的内容
        for lang_data in data:
            lang = lang_data['language']
            # 读取原文件
            excel = openpyxl.load_workbook(input_file)

            for sheet_data in lang_data['data']:
                sheet = excel[sheet_data['sheet']]
                # images = sheet._images
                # 更换单元格内容
                for cell in sheet_data['cells']:
                    if cell['is_merged'] and not cell['is_merge_start']:
                        continue
                    elif cell['value'] == '':
                        continue
                    else:
                        sheet[cell['coordinate']] = cell['value']
            # 存储文件
            new_output_file = output_file.replace(".xlsx", f"_{lang}.xlsx")
            time.sleep(1)
            excel.save(new_output_file)
            print(f"输出文件：{new_output_file}")

    elif method == 'complex':
        excel = openpyxl.Workbook()
        # 清除已经存在的sheet
        default_sheet = excel.active
        excel.remove(default_sheet)

        # 创建新sheet并填充内容
        for sheet_data in data:
            sheet = excel.create_sheet(sheet_data['sheet'])  # sheet名称
            sheet.sheet_properties.tabColor = sheet_data['sheet_color']  # sheet颜色
            # 填充内容
            for cell in sheet_data['cells']:
                sheet[cell['coordinate']] = cell['value']
                apply_cell_format(sheet[cell['coordinate']], cell)

            # 合并单元格
            for merge_info in sheet_data['merge_info']:
                sheet.merge_cells(str(merge_info))

        # 存储文件
        excel.save(output_file)
        print(f"输出文件：{output_file}")


if __name__ == "__main__":
    # 获取当前时间戳
    current_time = datetime.datetime.now().strftime('%y%m%d%H%M%S')
    language = ['英语', '越南语']
    # language = ['英语']

    input_file = r"D:\Code\Project\tools\data\test\1.xlsx"

    output_folder = r"D:\Code\Project\tools\data\temp"
    file_base_name = os.path.basename(input_file)
    output_file = output_folder + "/" + file_base_name.replace(".xlsx", f"_translate_{current_time}.xlsx")

    translator = Translator()
    print(f"********************start at {current_time}********************")

    # 1. 读取原始内容
    # 文字
    content_data = get_content(input_file)
    # 图片
    # image_data = extract_images_from_excel(input_file, output_folder)

    # 2. 翻译
    # translated_data = add_translation(content_data, translator, language, 'replace')
    translated_data = add_translation(content_data, translator, language, 'replace_multi')

    # 3. 处理数据内容，使其符合创建新文档格式
    # formatted_data = apply_template(origin_data, template)
    formatted_data = translated_data

    # 4. 创建新文档
    create_new_excel(output_file, formatted_data, 'simple', input_file)
