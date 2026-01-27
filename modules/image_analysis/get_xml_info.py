#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : get_xml_info.py 
@Author  : Shawn
@Date    : 2026/1/23 9:31 
@Info    : Description of this file
"""

import os
import shutil
from pathlib import Path
import os
import csv
import pandas as pd
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Optional


def read_xml_result_status(xml_path: str) -> Optional[str]:
    """
    读取XML文件，提取所有resultStatus标签内容
    所有resultStatus均为0时返回"OK"，否则返回"NG"
    """
    try:
        # 解析XML文件
        tree = ET.parse(xml_path)
        root = tree.getroot()

        # 查找所有的resultStatus标签
        result_status_elements = root.findall(".//resultStatus")

        if not result_status_elements:
            print(f"警告: {xml_path} 中未找到resultStatus标签")
            return "NG"

        # 检查所有的resultStatus值
        all_zero = True
        for element in result_status_elements:
            status_text = element.text
            if status_text and status_text.strip() != "0":
                all_zero = False
                break

        return "OK" if all_zero else "NG"

    except ET.ParseError as e:
        print(f"XML解析错误: {xml_path} - {e}")
        return "NG"
    except FileNotFoundError:
        print(f"文件不存在: {xml_path}")
        return "NG"
    except Exception as e:
        print(f"读取XML时发生错误: {xml_path} - {e}")
        return "NG"


def extract_xml_filename_from_image_path(image_path: str, xml_folder: str) -> str:
    """
    从图片路径提取文件名，并构造XML文件路径
    """
    # 获取文件名（不含扩展名）
    file_name = Path(image_path).stem  # 获取不带扩展名的文件名
    xml_filename = f"{file_name}.xml"

    # 构造XML文件完整路径
    xml_path = os.path.join(xml_folder, xml_filename)
    return xml_path


def process_csv_to_xlsx(csv_file_path: str, xml_folder: str, output_xlsx_path: str):
    """
    处理CSV文件，读取XML结果，生成XLSX文件
    """
    # 读取CSV文件
    try:
        # 使用pandas读取CSV，确保处理各种编码
        df = pd.read_csv(csv_file_path, encoding='utf-8', engine='python')
        print(f"成功读取CSV文件: {csv_file_path}")
        print(f"CSV文件列名: {list(df.columns)}")
        print(f"CSV文件形状: {df.shape}")
    except Exception as e:
        print(f"读取CSV文件失败: {e}")
        # 尝试其他编码
        try:
            df = pd.read_csv(csv_file_path, encoding='gbk', engine='python')
            print(f"使用GBK编码成功读取CSV文件")
        except Exception as e2:
            print(f"使用GBK编码也失败: {e2}")
            return

    # 检查第六列是否存在
    if len(df.columns) < 6:
        print(f"错误: CSV文件只有{len(df.columns)}列，需要至少6列")
        return

    # 获取第六列的列名
    sixth_col_name = df.columns[5]
    print(f"第六列列名: {sixth_col_name}")

    # 创建新的列用于存储XML结果
    xml_results = []

    # 处理每一行
    for idx, row in df.iterrows():
        image_path = str(row[sixth_col_name])

        # 提取XML文件路径
        xml_path = extract_xml_filename_from_image_path(image_path, xml_folder)

        # 读取XML文件结果
        result = read_xml_result_status(xml_path)
        xml_results.append(result)

        # 显示进度
        if (idx + 1) % 10 == 0:
            print(f"已处理 {idx + 1}/{len(df)} 行")

    # 将结果添加到DataFrame的第十列
    # 如果列数不足10列，先添加空列
    while len(df.columns) < 10:
        df[f'Column_{len(df.columns) + 1}'] = ''

    # 将第十列（索引9）的列名设置为"XML结果"
    df.columns = list(df.columns[:9]) + ["XML结果"] + list(df.columns[10:])
    # 将结果放入第十列（索引9）
    df.iloc[:, 9] = xml_results

    # 保存为XLSX文件
    try:
        df.to_excel(output_xlsx_path, index=False)
        print(f"成功保存XLSX文件: {output_xlsx_path}")
        print(f"处理完成！共处理 {len(df)} 行数据")
    except Exception as e:
        print(f"保存XLSX文件失败: {e}")


def get_xml_recognition_results():
    # 配置参数
    csv_file_path = input("请输入csv文件:\n")  # 替换为你的CSV文件路径
    xml_folder = input("请输入xml文件所在文件夹:\n")  # 替换为XML文件所在的文件夹路径
    output_xlsx_path = r"D:\Code\Project\tools\data\output\VM_result.xlsx"  # 输出XLSX文件路径

    # 检查输入文件是否存在
    if not os.path.exists(csv_file_path):
        print(f"错误: CSV文件不存在: {csv_file_path}")
        return

    if not os.path.exists(xml_folder):
        print(f"错误: XML文件夹不存在: {xml_folder}")
        return

    # 处理CSV文件
    process_csv_to_xlsx(csv_file_path, xml_folder, output_xlsx_path)
