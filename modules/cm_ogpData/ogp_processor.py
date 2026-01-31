#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@Project : tools
@File    : ogp_processor.py
@Author  : Shawn
@Date    : 2026/1/8 9:45
@Info    : OGP数据汇总程序
"""

import os
from datetime import datetime
import pandas as pd
import re


class OGPProcessor:
    def __init__(self):
        pass

    def process_single_file(self, file_path, index, process_mode, file_format, header_lines, output_dir,
                            create_subfolder, overwrite_existing):
        """处理单个文件"""
        try:
            # 读取文件
            df = self.read_file(file_path)
            if df is None:
                return False, f"[{index}] {os.path.basename(file_path)}: 文件读取失败"

            # 查找数据开始行
            data_start_row = self.find_data_start_row(df)
            if data_start_row == -1:
                return False, f"[{index}] {os.path.basename(file_path)}: 未找到数据行"

            # 检测列索引
            col_indices = self.detect_column_indices(df, data_start_row)
            if not col_indices:
                return False, f"[{index}] {os.path.basename(file_path)}: 列索引检测失败"

            # 从文件内容中提取元数据（模数、穴数范围等）
            metadata = self.extract_metadata(df, data_start_row)
            mold_num = metadata.get('mold_num', 1)  # 默认模数为1
            cavity_range = metadata.get('cavity_range', (1, 1))  # 默认穴数范围为(1, 1)

            # 提取数据
            all_data = self.extract_data(df, data_start_row, col_indices, mold_num, cavity_range)
            if not all_data:
                return False, f"[{index}] {os.path.basename(file_path)}: 数据提取失败"

            # 根据处理模式选择处理方法
            if process_mode == "summary":
                # 汇总模式
                result, block_count = self.summarize_data(all_data, df, data_start_row)
                mode_text = "汇总"
            else:
                # 仅排序模式
                result, block_count = self.sort_data(all_data, df, data_start_row)
                mode_text = "排序"

            # 确定输出路径
            input_filename = os.path.basename(file_path)
            base_name = os.path.splitext(input_filename)[0]
            suffix = os.path.splitext(input_filename)[1]
            if not suffix:
                suffix = ".xlsx"

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

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
            self.save_output(result, output_file)

            return True, f"[{index}] {input_filename}: {mode_text}处理完成，识别到 {block_count} 个区块 -> {output_filename}"

        except Exception as e:
            return False, f"[{index}] {os.path.basename(file_path)}: 处理失败 - {str(e)}"

    def extract_metadata(self, df, data_start_row):
        """从文件内容中提取元数据（模数、穴数范围等）"""
        metadata = {
            'mold_num': 1,  # 默认模数为1
            'cavity_range': (1, 1)  # 默认穴数范围为(1, 1)
        }
        
        # 尝试从文件内容中提取模数和穴数范围
        # 1. 检查表头行是否包含相关信息
        for i in range(min(data_start_row, 10)):
            row = df.iloc[i]
            row_str = ' '.join([str(cell) for cell in row if pd.notna(cell)])
            
            # 尝试提取模数
            mold_match = re.search(r'第(\d+)模', row_str)
            if mold_match:
                try:
                    metadata['mold_num'] = int(mold_match.group(1))
                except ValueError:
                    pass
            
            # 尝试提取穴数范围
            cavity_match = re.search(r'(\d+)穴', row_str)
            if cavity_match:
                try:
                    cavity_num = int(cavity_match.group(1))
                    metadata['cavity_range'] = (cavity_num, cavity_num)
                except ValueError:
                    pass
            
            # 尝试提取穴数范围（带~符号）
            cavity_range_match = re.search(r'(\d+)穴~(\d+)穴', row_str)
            if cavity_range_match:
                try:
                    start_cavity = int(cavity_range_match.group(1))
                    end_cavity = int(cavity_range_match.group(2))
                    metadata['cavity_range'] = (start_cavity, end_cavity)
                except ValueError:
                    pass
        
        return metadata

    def read_file(self, file_path):
        """读取文件，仅支持.xls, .xlsx"""
        try:
            # 根据文件扩展名选择读取方法
            ext = os.path.splitext(file_path)[1].lower()
            if ext in ['.xls', '.xlsx']:
                try:
                    # 尝试使用xlrd引擎读取
                    df = pd.read_excel(file_path, header=None, engine='xlrd')
                    return df
                except Exception as e1:
                    print(f"使用xlrd引擎读取失败: {e1}")
                    # 尝试使用openpyxl引擎读取
                    try:
                        df = pd.read_excel(file_path, header=None, engine='openpyxl')
                        return df
                    except Exception as e2:
                        print(f"使用openpyxl引擎读取失败: {e2}")
                        # 尝试以制表符分隔的文本文件读取
                        try:
                            df = pd.read_csv(file_path, header=None, sep='\t', encoding='gbk')
                            return df
                        except Exception as e3:
                            print(f"以文本文件读取失败: {e3}")
                            return None
            else:
                # 不支持的文件格式
                print(f"不支持的文件格式: {ext}")
                return None
        except Exception as e:
            print(f"文件读取失败: {e}")
            return None

    def find_data_start_row(self, df):
        """查找数据开始行"""
        for i in range(min(10, len(df))):
            row = df.iloc[i]
            # 检查是否包含标签列（通常第一列是标签）
            row_str = ' '.join([str(cell) for cell in row if pd.notna(cell)])
            if any(keyword in row_str for keyword in ['标签', '尺寸', '序号']):
                # 下一行开始是数据
                if i + 1 < len(df):
                    return i + 1
        # 默认从第二行开始
        return 1 if len(df) > 1 else 0

    def detect_column_indices(self, df, data_start_row):
        """检测列索引"""
        # 尝试获取数据行
        if data_start_row >= len(df):
            return None
        
        data_row = df.iloc[data_start_row]
        
        # 根据用户提供的文件结构，列顺序为：标签, 尺寸类型, 标准值, 实测值, 上公差, 下公差
        col_indices = {
            'label': 0,        # 标签
            'size_type': 1,    # 尺寸类型
            'nominal': 2,      # 标准值
            'measured': 3,     # 实测值
            'upper_tol': 4,    # 上公差
            'lower_tol': 5,    # 下公差
            'deviation': 6     # 偏差值
        }
        
        # 检查列数是否足够
        if len(data_row) < 6:
            return None
        
        return col_indices

    def extract_data(self, df, data_start_row, col_indices, mold_num, cavity_range):
        """提取数据"""
        all_data = {}
        
        # 用于识别模数的变量
        current_mold = 1
        seen_labels = set()
        labels_in_current_mold = set()
        
        for row_idx in range(data_start_row, len(df)):
            row = df.iloc[row_idx]
            
            # 提取标签
            label = self.get_cell_value(row, col_indices['label'])
            if label is None or pd.isna(label) or str(label).strip() == '':
                continue
            
            label_str = str(label).strip()
            
            # 提取其他数据
            size_type = self.get_cell_value(row, col_indices['size_type'])
            nominal = self.get_cell_value(row, col_indices['nominal'])
            upper_tol = self.get_cell_value(row, col_indices['upper_tol'])
            lower_tol = self.get_cell_value(row, col_indices['lower_tol'])
            measured = self.get_cell_value(row, col_indices['measured'])
            
            # 检查是否开始新的一模
            if label_str in seen_labels and label_str in labels_in_current_mold:
                # 遇到重复的标签，说明开始新的一模
                current_mold += 1
                labels_in_current_mold = set()
            
            # 初始化数据结构
            if label_str not in all_data:
                all_data[label_str] = {
                    'size_type': size_type,
                    'nominal': nominal,
                    'upper_tol': upper_tol,
                    'lower_tol': lower_tol,
                    'measurements': {}
                }
            
            # 初始化模序号的测量值字典
            if current_mold not in all_data[label_str]['measurements']:
                all_data[label_str]['measurements'][current_mold] = measured
            
            # 更新已见标签
            seen_labels.add(label_str)
            labels_in_current_mold.add(label_str)
        
        return all_data

    def get_cell_value(self, row, col_idx):
        """获取单元格值"""
        if col_idx >= len(row):
            return None
        value = row.iloc[col_idx]
        return value if pd.notna(value) else None

    def summarize_data(self, all_data, df, data_start_row):
        """汇总数据"""
        # 准备输出数据
        output_rows = []

        # 保留文件的前5行信息
        for i in range(min(5, len(df))):
            row = df.iloc[i]
            # 转换为列表并添加到输出行
            output_rows.append(list(row))

        # 收集所有的模序号
        all_mold_nums = set()
        for label, data in all_data.items():
            for mold_num in data['measurements']:
                all_mold_nums.add(mold_num)
        
        # 排序模序号
        sorted_mold_nums = sorted(all_mold_nums)

        # 生成表头
        header = ['标签', '尺寸类型', '标准值', '上公差', '下公差']
        for mold_num in sorted_mold_nums:
            header.append(f'#{mold_num}')
        output_rows.append(header)

        # 生成数据行
        sorted_labels = sorted(all_data.keys(), key=self.get_sort_key)
        for label in sorted_labels:
            data = all_data[label]
            row = [
                label,
                data['size_type'] if pd.notna(data['size_type']) else '',
                data['nominal'] if pd.notna(data['nominal']) else '',
                data['upper_tol'] if pd.notna(data['upper_tol']) else '',
                data['lower_tol'] if pd.notna(data['lower_tol']) else ''
            ]
            
            # 添加各模次的实测值
            for mold_num in sorted_mold_nums:
                if mold_num in data['measurements']:
                    row.append(data['measurements'][mold_num])
                else:
                    row.append('')
            
            output_rows.append(row)

        # 创建输出DataFrame
        output_df = pd.DataFrame(output_rows)
        return output_df, len(all_data)

    def sort_data(self, all_data, df, data_start_row):
        """排序数据"""
        # 准备输出数据
        output_rows = []

        # 保留文件的前5行信息
        for i in range(min(5, len(df))):
            row = df.iloc[i]
            # 转换为列表并添加到输出行
            output_rows.append(list(row))

        # 生成表头
        header = ['标签', '尺寸类型', '标准值', '上公差', '下公差', '实测值']
        output_rows.append(header)

        # 生成数据行
        sorted_labels = sorted(all_data.keys(), key=self.get_sort_key)
        for label in sorted_labels:
            data = all_data[label]
            # 获取第一个模序号的实测值
            mold_num = next(iter(data['measurements'].keys()), None)
            measured = data['measurements'][mold_num] if mold_num else ''
            
            row = [
                label,
                data['size_type'] if pd.notna(data['size_type']) else '',
                data['nominal'] if pd.notna(data['nominal']) else '',
                data['upper_tol'] if pd.notna(data['upper_tol']) else '',
                data['lower_tol'] if pd.notna(data['lower_tol']) else '',
                measured
            ]
            output_rows.append(row)

        # 创建输出DataFrame
        output_df = pd.DataFrame(output_rows)
        return output_df, len(all_data)

    def get_sort_key(self, label):
        """生成排序键"""
        try:
            label_str = str(label).strip()
            # 检查是否以数字开头
            if label_str and label_str[0].isdigit():
                # 数字开头的标签
                if '*' in label_str:
                    # 处理带星号的标签，如 "15*1", "15*2"
                    num_parts = label_str.split('*')
                    main_num = int(num_parts[0])
                    sub_num = int(num_parts[1]) if len(num_parts) > 1 else 0
                    return (0, '', main_num, sub_num)
                else:
                    # 处理纯数字标签，如 "1", "2", "11"
                    main_num = int(label_str)
                    return (0, '', main_num, 0)
            else:
                # 字母开头的标签
                # 提取字母部分和数字部分
                alpha_part = ''
                num_part = ''
                star_part = ''
                
                # 找到第一个数字的位置
                for i, char in enumerate(label_str):
                    if char.isdigit():
                        alpha_part = label_str[:i]
                        remaining = label_str[i:]
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
                return (1, alpha_part.upper(), num_val, star_val)
        except (ValueError, IndexError):
            # 如果解析失败，放到最后
            return (2, '', 0, 0)

    def save_output(self, result, output_file):
        """保存输出文件"""
        try:
            # 确保保存为.xlsx格式
            output_file_xlsx = os.path.splitext(output_file)[0] + '.xlsx'
            # 保存为Excel文件
            result.to_excel(output_file_xlsx, index=False, header=False)
        except Exception as e:
            print(f"文件保存失败: {e}")
            # 尝试保存为文本文件
            try:
                output_file_txt = os.path.splitext(output_file)[0] + '.txt'
                result.to_csv(output_file_txt, index=False, header=False, sep='\t')
            except:
                pass

    def merge_ogp_files(self, file_paths, output_dir, processing, create_subfolder):
        """合并多个OGP文件到一个汇总文件（OGP专用）"""
        try:
            # 收集所有文件的数据
            all_files_data = []
            all_labels = set()
            all_mold_nums = set()

            for i, file_path in enumerate(file_paths):
                if not processing:
                    return False, "处理已停止", None

                # 读取文件
                df = self.read_file(file_path)
                if df is None:
                    return False, f"文件读取失败: {os.path.basename(file_path)}", None

                # 查找数据开始行
                data_start_row = self.find_data_start_row(df)
                if data_start_row == -1:
                    return False, f"未找到数据行: {os.path.basename(file_path)}", None

                # 检测列索引
                col_indices = self.detect_column_indices(df, data_start_row)
                if not col_indices:
                    return False, f"列索引检测失败: {os.path.basename(file_path)}", None

                # 提取数据
                # 这里我们需要修改extract_data方法，使其返回所有数据，包括标签和模次
                # 为了简化，我们先调用现有的方法，然后处理结果
                metadata = self.extract_metadata(df, data_start_row)
                mold_num = metadata.get('mold_num', 1)
                cavity_range = metadata.get('cavity_range', (1, 1))

                file_data = self.extract_data(df, data_start_row, col_indices, mold_num, cavity_range)
                if not file_data:
                    return False, f"数据提取失败: {os.path.basename(file_path)}", None

                # 收集标签和模次
                for label, data in file_data.items():
                    all_labels.add(label)
                    for mold in data['measurements']:
                        all_mold_nums.add(mold)

                # 保存文件数据
                all_files_data.append((os.path.basename(file_path), file_data))

            # 如果没有数据
            if not all_files_data:
                return False, "没有可处理的数据", None

            # 准备汇总数据
            output_rows = []

            # 生成表头：标签列 + 尺寸类型列 + 标准值列 + 公差列 + 各模次列（去掉文件名）
            header = ['标签', '尺寸类型', '标准值', '上公差', '下公差']
            sorted_mold_nums = sorted(all_mold_nums)
            for mold in sorted_mold_nums:
                header.append(f'#{mold}')
            output_rows.append(header)

            # 生成数据行
            for file_name, file_data in all_files_data:
                sorted_labels = sorted(file_data.keys(), key=self.get_sort_key)
                for label in sorted_labels:
                    data = file_data[label]
                    row = [
                        label,
                        data['size_type'] if pd.notna(data['size_type']) else '',
                        data['nominal'] if pd.notna(data['nominal']) else '',
                        data['upper_tol'] if pd.notna(data['upper_tol']) else '',
                        data['lower_tol'] if pd.notna(data['lower_tol']) else ''
                    ]

                    # 添加各模次的实测值
                    for mold in sorted_mold_nums:
                        if mold in data['measurements']:
                            row.append(data['measurements'][mold])
                        else:
                            row.append('')

                    output_rows.append(row)

            # 创建输出DataFrame
            output_df = pd.DataFrame(output_rows)

            # 从第一个文件中提取产品名（第一个"-"之前的内容）
            product_name = "OGP"
            if all_files_data:
                first_file_name = all_files_data[0][0]
                # 提取第一个"-"之前的内容
                if "-" in first_file_name:
                    product_name = first_file_name.split("-")[0].strip()
                else:
                    # 如果没有"-"，使用文件名（不含扩展名）
                    product_name = os.path.splitext(first_file_name)[0].strip()

            # 确定输出路径
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"{product_name}_汇总_{timestamp}.xlsx"

            # 如果需要创建子文件夹
            if create_subfolder:
                subfolder_name = "OGP数据汇总"
                output_folder = os.path.join(output_dir, subfolder_name)
                os.makedirs(output_folder, exist_ok=True)
                output_file = os.path.join(output_folder, output_filename)
            else:
                output_file = os.path.join(output_dir, output_filename)

            # 保存文件
            self.save_output(output_df, output_file)

            return True, f"汇总完成，共处理 {len(file_paths)} 个文件 -> {os.path.basename(output_file)}", output_file

        except Exception as e:
            return False, f"汇总失败: {str(e)}", None
