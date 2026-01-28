#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@Project : tools
@File    : ogp_processor.py
@Author  : Shawn
@Date    : 2026/1/8 9:45
@Info    : OGP数据汇总程序
"""

import re
import os
from datetime import datetime
import chardet

class OGPProcessor:
    def __init__(self):
        pass

    def process_single_file(self, file_path, index, process_mode, file_format, header_lines, output_dir, create_subfolder, overwrite_existing):
        """处理单个文件"""
        try:
            # 检测编码
            encoding = detect_encoding(file_path)

            # 读取文件
            with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                content = f.read()

            # 确定文件格式
            detected_format = self.detect_file_format(content, file_format)

            # 根据处理模式选择处理方法
            if process_mode == "summary":
                # 汇总模式
                if detected_format == "format2":
                    result, block_count = self.summarize_format2_data(content, header_lines)
                else:
                    result, block_count = self.summarize_data(content, header_lines)
                mode_text = "汇总"
            else:
                # 仅排序模式
                if detected_format == "format2":
                    result, block_count = self.sort_format2_data(content, header_lines)
                else:
                    result, block_count = self.sort_data(content, header_lines)
                mode_text = "排序"

            # 确定输出路径
            input_filename = os.path.basename(file_path)
            base_name = os.path.splitext(input_filename)[0]
            suffix = os.path.splitext(input_filename)[1]
            if not suffix:
                suffix = ".txt"

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            # mode_suffix = "_sorted" if process_mode == "sort" else "_summarized"

            # 如果需要创建子文件夹
            if create_subfolder:
                subfolder_name = "OGP数据汇总"
                output_folder = os.path.join(output_dir, subfolder_name)
                os.makedirs(output_folder, exist_ok=True)
                output_filename = f"{base_name}_{timestamp}{suffix}"
                output_file = os.path.join(output_folder, output_filename)
            else:
                output_folder = output_dir
                # 检查是否覆盖
                output_filename = f"{base_name}_{timestamp}{suffix}"
                output_file = os.path.join(output_folder, output_filename)

                # 如果文件已存在且不覆盖，添加序号
                if os.path.exists(output_file) and not overwrite_existing:
                    counter = 1
                    while os.path.exists(output_file):
                        output_filename = f"{base_name}_{timestamp}_{counter}{suffix}"
                        output_file = os.path.join(output_folder, output_filename)
                        counter += 1

            # 保存文件
            with open(output_file, 'w', encoding=encoding, errors='ignore') as f:
                f.write(result)

            return True, f"[{index}] {input_filename}: {mode_text}处理完成，识别到 {block_count} 个区块 -> {output_filename}"

        except Exception as e:
            return False, f"[{index}] {os.path.basename(file_path)}: 处理失败 - {str(e)}"

    def detect_file_format(self, content, user_format):
        """检测文件格式"""
        if user_format != "auto":
            return user_format

        # 自动检测格式
        lines = content.strip().split('\n')

        # 检查是否有空行
        has_empty_lines = any(line.strip() == '' for line in lines)

        if has_empty_lines:
            return "format1"  # 带空行分隔的格式

        # 检查是否是格式2（单表头，通过标签重复区分区块）
        # 统计标签出现的次数
        label_counts = {}
        data_line_pattern = re.compile(r'^\s*(\d+[\*\d]*\s+|[\d\.]+\s+)')

        for line in lines:
            if data_line_pattern.match(line.strip()):
                parts = line.strip().split('\t')
                if parts:
                    label = parts[0].strip()
                    label_counts[label] = label_counts.get(label, 0) + 1

        # 如果有标签出现多次，可能是格式2
        max_count = max(label_counts.values()) if label_counts else 0

        if max_count > 1:
            # 进一步确认：检查是否有明显的表头行
            # 格式2通常有特定的表头关键词
            header_keywords = ['±êÇ©', '³ß´çÀàÐÍ', '±ê×¼Öµ', 'Êµ²âÖµ', 'ÉÏ¹«²î', 'ÏÂ¹«²î', 'Æ«²îÖµ']
            has_cn_header = any(keyword in content for keyword in header_keywords)

            if has_cn_header:
                return "format2"

        return "format1"  # 默认为格式1

    def sort_data(self, content, header_lines):
        """仅排序模式：处理文件内容，按区块排序第一列数据（格式1）"""
        lines = content.strip().split('\n')
        blocks = []
        current_block = []
        data_line_pattern = re.compile(r'^\s*(\d+[\*\d]*)\s+')

        # 识别区块（通过空行分隔）
        for line in lines:
            if line.strip() == '':
                if current_block:
                    blocks.append(current_block)
                    current_block = []
            else:
                current_block.append(line)

        if current_block:
            blocks.append(current_block)

        block_count = len(blocks)
        processed_blocks = []

        # 处理每个区块
        for block in blocks:
            processed_block = self.process_block(block, data_line_pattern, header_lines)
            processed_blocks.append(processed_block)

        # 合并所有区块（用空行分隔）
        result_lines = []
        for i, block in enumerate(processed_blocks):
            if i > 0:
                result_lines.append('')
            result_lines.extend(block)

        return '\n'.join(result_lines), block_count

    def sort_format2_data(self, content, header_lines):
        """仅排序模式：处理格式2的数据（单表头无空行）"""
        lines = content.strip().split('\n')
        block_count = 1  # 格式2只有一个表头，数据是连续的

        # 分离表头和数据
        header_lines_list = []
        data_lines = []
        header_end_index = -1

        # 找到表头结束的位置
        data_line_pattern = re.compile(r'^\s*(\d+[\*\d]*\s+|[\d\.]+\s+)')
        for i, line in enumerate(lines):
            if data_line_pattern.match(line.strip()):
                header_end_index = i
                break
            else:
                header_lines_list.append(line)

        # 如果没找到数据行，直接返回原始内容
        if header_end_index == -1:
            return content, 0

        # 提取所有数据行
        data_lines = lines[header_end_index:]

        # 按标签排序数据行
        sorted_data_lines = self.sort_data_lines(data_lines, data_line_pattern)

        # 重新组合
        result_lines = header_lines_list + sorted_data_lines

        return '\n'.join(result_lines), block_count

    def summarize_data(self, content, header_lines):
        """
        汇总模式：处理格式1的数据
        1. 先排序数据
        2. 调整列顺序并合并实测值
        3. 只保留第一个区块
        """
        lines = content.strip().split('\n')
        blocks = []
        current_block = []
        data_line_pattern = re.compile(r'^\s*(\d+[\*\d]*)\s+')

        # 识别区块（通过空行分隔）
        for line in lines:
            if line.strip() == '':
                if current_block:
                    blocks.append(current_block)
                    current_block = []
            else:
                current_block.append(line)

        if current_block:
            blocks.append(current_block)

        block_count = len(blocks)

        if block_count == 0:
            return content, 0

        # 处理每个区块（先排序）
        processed_blocks = []
        for block in blocks:
            # 先排序数据行
            sorted_block = self.process_block(block, data_line_pattern, header_lines)
            processed_blocks.append(sorted_block)

        if block_count == 1:
            # 只有一个区块，只需调整列顺序
            result = self.reorder_and_format_block(processed_blocks[0], header_lines)
            return '\n'.join(result), block_count

        # 多个区块的情况
        # 1. 从所有区块提取数据
        all_data = self.extract_data_from_blocks(processed_blocks, header_lines)

        # 2. 按照标签排序数据
        sorted_data = self.sort_extracted_data(all_data)

        # 3. 重新构建输出
        result_block = self.rebuild_output_block(processed_blocks[0][:header_lines], sorted_data)

        return '\n'.join(result_block), block_count

    def summarize_format2_data(self, content, header_lines):
        """
        汇总模式：处理格式2的数据（单表头无空行）
        通过标签重复来识别不同的数据块
        """
        lines = content.strip().split('\n')

        # 分离表头和数据
        header_lines_list = []
        data_lines = []
        data_line_pattern = re.compile(r'^[A-Za-z0-9_\-+=*.@#\$%^&()\[\]{}\|;:,.<>?/ \t].*')

        for i, line in enumerate(lines):
            if data_line_pattern.match(line.strip()):
                # 从第一个数据行开始
                data_lines = lines[i:]
                header_lines_list = lines[:i]
                break

        if not data_lines:
            return content, 0

        # 解析数据行，通过标签重复来识别区块
        blocks = self.split_blocks_by_label_repeat(data_lines)
        block_count = len(blocks)

        if block_count == 0:
            return content, 0

        if block_count == 1:
            # 只有一个区块，只需调整列顺序
            result_block = header_lines_list + self.reorder_and_format_block_format2(blocks[0])
            return '\n'.join(result_block), block_count

        # 多个区块的情况
        # 1. 从所有区块提取数据
        all_data = self.extract_data_from_blocks_format2(blocks)

        # 2. 按照标签排序数据
        sorted_data = self.sort_extracted_data(all_data)

        # 3. 重新构建输出
        result_block = self.rebuild_output_block_format2(header_lines_list, sorted_data)

        return '\n'.join(result_block), block_count

    def split_blocks_by_label_repeat(self, data_lines):
        """
        通过标签重复来拆分数据块
        返回列表，每个元素是一个区块的数据行
        """
        blocks = []
        current_block = []
        seen_labels = set()
        block_started = False

        for line in data_lines:
            # 解析标签
            columns = line.strip().split('\t')
            if not columns:
                continue

            label = columns[0].strip()

            # 如果是第一次看到这个标签，或者标签重复出现
            if label in seen_labels and block_started:
                # 标签重复，开始新的区块
                blocks.append(current_block)
                current_block = []
                seen_labels = {label}
                current_block.append(line)
            else:
                # 继续当前区块
                current_block.append(line)
                seen_labels.add(label)
                block_started = True

        # 添加最后一个区块
        if current_block:
            blocks.append(current_block)

        return blocks

    def extract_data_from_blocks_format2(self, blocks):
        """
        从格式2的所有区块中提取数据
        返回字典：{标签: {区块索引: 数据行}}
        """
        data_dict = {}

        for block_idx, block in enumerate(blocks):
            for line in block:
                columns = self.parse_data_line_format2(line)
                if columns and len(columns) >= 6:
                    label = columns[0]  # 第一列是标签
                    if label not in data_dict:
                        data_dict[label] = {}
                    data_dict[label][block_idx] = columns

        return data_dict

    def parse_data_line_format2(self, line):
        """解析格式2的数据行"""
        # 格式2的数据行通常是制表符分隔的
        columns = line.strip().split('\t')
        if len(columns) >= 6:
            return columns
        return None

    def reorder_and_format_block_format2(self, block):
        """调整格式2单个区块的列顺序"""
        formatted_lines = []

        for line in block:
            columns = self.parse_data_line_format2(line)
            if columns and len(columns) >= 6:
                # 格式2的列顺序：标签, 标准值, 上公差, 下公差, 偏差值, 实测值
                # 我们需要重新排序为：标签, 类型(?), 标准值, 上公差, 下公差, 实测值
                label = columns[0]
                nominal = columns[1]
                upper_tol = columns[2]
                lower_tol = columns[3]
                deviation = columns[4]  # 偏差值，可能不需要
                measured = columns[5]

                # 格式化数字
                nominal_formatted = self.format_number(nominal)
                upper_tol_formatted = self.format_number(upper_tol)
                lower_tol_formatted = self.format_number(lower_tol)
                measured_formatted = self.format_number(measured)

                # 构建格式化行
                # 注意：格式2没有明确的"类型"列，我们可以用默认值或留空
                dim_type = "D"  # 默认类型
                formatted_line = f"{label}\t{dim_type}\t{nominal_formatted}\t{upper_tol_formatted}\t{lower_tol_formatted}\t{measured_formatted}"
                formatted_lines.append(formatted_line)
            else:
                formatted_lines.append(line)

        return formatted_lines

    def rebuild_output_block_format2(self, header, sorted_data):
        """重新构建格式2的输出区块"""
        result = header.copy()

        # 确定最大实测值数量（即最大区块数）
        max_block_count = 0
        for sort_key, label, block_data in sorted_data:
            block_count = len(block_data)
            if block_count > max_block_count:
                max_block_count = block_count

        # 修改表头行
        if result:
            # 修改最后一行表头
            last_header = result[-1]
            header_parts = last_header.strip().split('\t')
            if len(header_parts) >= 6:
                # 构建新的表头，添加实测值#1, 实测值#2等
                formatted_title = "标签\t尺寸类型\t标准值\t上公差\t下公差"
                for i in range(1, max_block_count + 1):
                    formatted_title += f"\t实测值#{i}"
                result[-1] = formatted_title

        for sort_key, label, block_data in sorted_data:
            # 获取第一个区块的数据作为基础
            if 0 in block_data:
                base_columns = block_data[0]

                # 提取基础信息
                dim_type = base_columns[1]
                nominal = self.format_number(base_columns[2])
                upper_tol = self.format_number(base_columns[4])
                lower_tol = self.format_number(base_columns[5])

                # 收集所有区块的实测值
                measurements = []
                for block_idx in sorted(block_data.keys()):
                    block_columns = block_data[block_idx]
                    if len(block_columns) >= 6:
                        measurements.append(self.format_number(block_columns[3]))  # 实测值在第6列

                # 构建输出行
                formatted_line = f"{label}\t{dim_type}\t{nominal}\t{upper_tol}\t{lower_tol}"

                # 添加所有实测值
                for measurement in measurements:
                    formatted_line += f"\t{measurement}"

                result.append(formatted_line)

        return result

    def process_block(self, block, data_line_pattern, header_lines):
        """处理单个区块（排序）"""
        data_start = -1

        if len(block) > header_lines:
            for i in range(header_lines, len(block)):
                if data_line_pattern.match(block[i]):
                    data_start = i
                    break

        if data_start == -1:
            for i, line in enumerate(block):
                if data_line_pattern.match(line):
                    data_start = i
                    break

        if data_start == -1:
            return block

        header = block[:data_start]
        data_lines = block[data_start:]

        sorted_data_lines = self.sort_data_lines(data_lines, data_line_pattern)

        return header + sorted_data_lines

    def sort_data_lines(self, data_lines, data_line_pattern):
        """对数据行进行排序"""
        data_with_keys = []

        for i, line in enumerate(data_lines):
            match = data_line_pattern.match(line)
            if match:
                first_col = match.group(1).strip()
                
                # 初始化排序键
                sort_key = None
                
                try:
                    # 检查是否以数字开头
                    if first_col[0].isdigit():
                        # 数字开头的标签
                        if '*' in first_col:
                            # 处理带星号的标签，如 "15*1", "15*2"
                            num_parts = first_col.split('*')
                            main_num = int(num_parts[0])
                            sub_num = int(num_parts[1]) if len(num_parts) > 1 else 0
                            sort_key = (0, '', main_num, sub_num, i)
                        else:
                            # 处理纯数字标签，如 "1", "2", "11"
                            main_num = int(first_col)
                            sort_key = (0, '', main_num, 0, i)
                    else:
                        # 字母开头的标签
                        # 提取字母部分和数字部分
                        alpha_part = ''
                        num_part = ''
                        star_part = ''
                        
                        # 找到第一个数字的位置
                        for j, char in enumerate(first_col):
                            if char.isdigit():
                                alpha_part = first_col[:j]
                                remaining = first_col[j:]
                                # 处理可能的星号
                                if '*' in remaining:
                                    num_star_parts = remaining.split('*')
                                    num_part = num_star_parts[0]
                                    star_part = num_star_parts[1] if len(num_star_parts) > 1 else '0'
                                else:
                                    num_part = remaining
                                    star_part = '0'
                                break
                        
                        # 转换数字部分为整数
                        try:
                            num_val = int(num_part)
                            star_val = int(star_part)
                        except ValueError:
                            num_val = 0
                            star_val = 0
                        
                        # 字母开头的排序键
                        sort_key = (1, alpha_part, num_val, star_val, i)
                except (ValueError, IndexError):
                    # 如果解析失败，放到最后
                    sort_key = (2, '', 0, 0, i)
                
                data_with_keys.append((sort_key, line))
            else:
                data_with_keys.append(((2, '', 0, 0, i), line))

        data_with_keys.sort(key=lambda x: x[0])

        return [line for _, line in data_with_keys]

    def extract_data_from_blocks(self, blocks, header_lines):
        """
        从所有区块中提取数据
        返回字典：{标签: {区块索引: 数据行}}
        """
        data_dict = {}

        for block_idx, block in enumerate(blocks):
            if len(block) > header_lines:
                for line in block[header_lines:]:
                    # 解析数据行
                    columns = self.parse_data_line(line)
                    if columns:
                        label = columns[0]  # 第一列是标签
                        if label not in data_dict:
                            data_dict[label] = {}
                        data_dict[label][block_idx] = columns

        return data_dict

    def parse_data_line(self, line):
        """解析数据行，返回列列表"""
        # 使用正则表达式匹配行中的列
        # 匹配模式：标签 类型 值1 值2 值3 值4 值5 值6 值7
        pattern1 = r'^\s*([^\s]+)\s+([^\s]+)\s+([-\d\.]+)\s+([-\d\.]+)\s+([-\d\.]+)\s+([-\d\.]+)\s+([-\d\.]+)\s+([-\d\.]+)\s+([-\d\.]+)'
        match = re.match(pattern1, line.strip())

        if match:
            return list(match.groups())

        # 尝试用制表符分割
        columns = line.strip().split('\t')
        if len(columns) >= 6:
            return columns

        return None

    def sort_extracted_data(self, data_dict):
        """对提取的数据进行排序"""
        # 将数据转换为列表以便排序
        data_list = []
        for label, block_data in data_dict.items():
            # 去除首尾空格
            label_stripped = label.strip()

            # 初始化排序键
            sort_key = None

            # 尝试解析标签
            try:
                # 分割主要部分和后缀
                parts = label_stripped.split(None, 1)
                main_part = parts[0]
                suffix = parts[1] if len(parts) > 1 else ""

                # 检查是否以数字开头
                if main_part[0].isdigit():
                    # 数字开头的标签
                    if '*' in main_part:
                        # 处理带星号的标签，如 "15*1", "15*2"
                        num_parts = main_part.split('*')
                        main_num = int(num_parts[0])
                        sub_num = int(num_parts[1]) if len(num_parts) > 1 else 0
                        sort_key = (0, '', main_num, sub_num, suffix, label_stripped)
                    else:
                        # 处理纯数字标签，如 "1", "2", "11"
                        main_num = int(main_part)
                        sort_key = (0, '', main_num, 0, suffix, label_stripped)
                else:
                    # 字母开头的标签
                    # 提取字母部分和数字部分
                    alpha_part = ''
                    num_part = ''
                    star_part = ''
                    
                    # 找到第一个数字的位置
                    for i, char in enumerate(main_part):
                        if char.isdigit():
                            alpha_part = main_part[:i]
                            remaining = main_part[i:]
                            # 处理可能的星号
                            if '*' in remaining:
                                num_star_parts = remaining.split('*')
                                num_part = num_star_parts[0]
                                star_part = num_star_parts[1] if len(num_star_parts) > 1 else '0'
                            else:
                                num_part = remaining
                                star_part = '0'
                            break
                    
                    # 转换数字部分为整数
                    try:
                        num_val = int(num_part)
                        star_val = int(star_part)
                    except ValueError:
                        num_val = 0
                        star_val = 0
                    
                    # 字母开头的排序键：(1, 字母部分, 数字部分, 星号部分, 后缀, 原始标签)
                    sort_key = (1, alpha_part, num_val, star_val, suffix, label_stripped)

            except (ValueError, IndexError):
                # 如果解析失败，放到最后
                sort_key = (2, '', 0, 0, '', label_stripped)

            data_list.append((sort_key, label, block_data))

        # 按标签排序
        data_list.sort(key=lambda x: x[0])

        return data_list

    def reorder_and_format_block(self, block, header_lines):
        """调整单个区块的列顺序"""
        if len(block) <= header_lines:
            return block

        header = block[:header_lines]
        data_lines = block[header_lines:]

        formatted_lines = []

        for line in data_lines:
            columns = self.parse_data_line(line)
            if columns and len(columns) >= 9:
                # 根据示例输出文件，列顺序应该是：
                # 标签, 类型, 标准值, 上公差, 下公差, 偏差?, 0?, 百分比?, 实测值1
                label = columns[0]
                dim_type = columns[1]
                nominal = columns[2]
                measured = columns[3]
                upper_tol = columns[4]
                lower_tol = columns[5]
                # 后面几列可能需要调整
                other_cols = columns[6]  # 第6,7,8列

                # 重新格式化行
                # 注意：根据示例，标准值可能需要去除多余的0
                nominal_formatted = self.format_number(nominal)

                formatted_line = f"{label}\t{dim_type}\t{nominal_formatted}\t{upper_tol}\t{lower_tol}\t{measured}"
                formatted_lines.append(formatted_line)
            else:
                formatted_lines.append(line)

        return header + formatted_lines

    def rebuild_output_block(self, header, sorted_data):
        """重新构建输出区块"""
        result = header.copy()
        
        # 确定最大实测值数量（即最大区块数）
        max_block_count = 0
        for sort_key, label, block_data in sorted_data:
            block_count = len(block_data)
            if block_count > max_block_count:
                max_block_count = block_count
        
        # 排序标题行
        title = result[-1].split('\t')
        label = title[0]
        dim_type = title[1]
        nominal = title[2]
        upper_tol = title[4]
        lower_tol = title[5]
        
        # 构建新的表头，添加实测值#1, 实测值#2等
        formatted_title = f"{label}\t{dim_type}\t{nominal}\t{upper_tol}\t{lower_tol}"
        for i in range(1, max_block_count + 1):
            formatted_title += f"\t实测值#{i}"
        result[-1] = formatted_title

        for sort_key, label, block_data in sorted_data:
            # 获取第一个区块的数据作为基础
            if 0 in block_data:
                base_columns = block_data[0]

                # 提取基础信息
                dim_type = base_columns[1]
                nominal = self.format_number(base_columns[2])
                upper_tol = base_columns[4]
                lower_tol = base_columns[5]

                # 收集所有区块的实测值
                measurements = []
                for block_idx in sorted(block_data.keys()):
                    block_columns = block_data[block_idx]
                    if len(block_columns) > 3:
                        measurements.append(block_columns[3])

                # 构建输出行
                formatted_line = f"{label}\t{dim_type}\t{nominal}\t{upper_tol}\t{lower_tol}"

                # 添加所有实测值
                for measurement in measurements:
                    formatted_line += f"\t{measurement}"

                result.append(formatted_line)

        return result

    def format_number(self, num_str):
        """格式化数字，去除多余的0"""
        try:
            # 尝试转换为浮点数
            num = float(num_str)
            # 如果是整数，显示为整数形式
            if num.is_integer():
                return str(int(num))
            else:
                # 去除末尾的0
                formatted = str(num).rstrip('0').rstrip('.')
                return formatted
        except ValueError:
            return num_str


def detect_encoding(file_path):
    """检测文件编码"""
    try:
        with open(file_path, 'rb') as f:
            raw_data = f.read(4096)
            result = chardet.detect(raw_data)

            encoding = result['encoding']
            confidence = result['confidence']

            if not encoding or confidence < 0.7:
                encoding = 'utf-8'
            print("****************", encoding, confidence)
            return encoding
    except Exception as e:
        print(f"编码检测失败: {e}")
        return 'utf-8'
