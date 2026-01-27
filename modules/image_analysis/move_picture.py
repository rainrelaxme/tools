#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : move_picture.py 
@Author  : Shawn
@Date    : 2026/1/16 13:22 
@Info    : 将指定文件夹下的第三层文件夹下图片，全部移动到指定文件夹下
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


def move_images_from_nested_folders(source_root, target_dir, image_extensions=None):
    """
    将三层嵌套文件夹中的所有图片剪切到指定目录

    参数:
        source_root: 源文件夹根目录
        target_dir: 目标目录
        image_extensions: 图片扩展名列表，默认为常见图片格式
    """
    # 默认支持的图片格式
    if image_extensions is None:
        image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp',
                            '.tiff', '.webp', '.svg', '.ico', '.jfif']

    # 确保目标目录存在
    target_path = Path(target_dir)
    target_path.mkdir(parents=True, exist_ok=True)

    # 计数器
    moved_count = 0
    skipped_count = 0
    error_count = 0

    # 遍历三层文件夹结构
    source_path = Path(source_root)

    # 第一层文件夹
    for level1_dir in source_path.iterdir():
        if not level1_dir.is_dir():
            continue

        # 第二层文件夹
        for level2_dir in level1_dir.iterdir():
            if not level2_dir.is_dir():
                continue

            # 遍历第三层文件夹中的所有文件
            for file_path in level2_dir.iterdir():
                if file_path.is_file():
                    # 检查是否为图片文件
                    if file_path.suffix.lower() in image_extensions:
                        try:
                            # 生成目标文件名（可以添加原始路径信息避免重名）
                            # 方法1：使用原始路径结构作为前缀
                            relative_parts = file_path.relative_to(source_path).parts[:-1]
                            prefix = "_".join(relative_parts)
                            # new_filename = f"{prefix}_{file_path.name}"
                            new_filename = f"{file_path.name}"

                            # 方法2：使用数字序号（简单重命名）
                            # new_filename = f"image_{moved_count:04d}{file_path.suffix}"

                            target_file = target_path / new_filename

                            # 如果目标文件已存在，添加序号
                            counter = 1
                            original_target = target_file
                            while target_file.exists():
                                name_without_ext = original_target.stem
                                target_file = target_path / f"{name_without_ext}_{counter}{original_target.suffix}"
                                counter += 1

                            # 剪切文件（移动）
                            shutil.move(str(file_path), str(target_file))
                            moved_count += 1
                            print(f"✓ 已移动: {file_path} -> {target_file}")

                        except Exception as e:
                            error_count += 1
                            print(f"✗ 移动失败 {file_path}: {e}")
                    else:
                        skipped_count += 1

    # 输出统计信息
    print("\n" + "=" * 50)
    print("移动完成！统计信息:")
    print(f"成功移动: {moved_count} 个文件")
    print(f"跳过文件: {skipped_count} 个（非图片文件）")
    print(f"错误文件: {error_count} 个")
    print(f"目标目录: {target_path}")
    print("=" * 50)


def interactive_mode():
    """交互式模式"""
    print("图片批量剪切工具")
    print("-" * 30)

    # 获取源目录
    while True:
        source = input("请输入源文件夹路径（三层嵌套结构根目录）: ").strip()
        if os.path.exists(source):
            break
        print("目录不存在，请重新输入！")

    # 获取目标目录
    target = input("请输入目标文件夹路径: ").strip()

    # 是否使用自定义扩展名
    use_custom = input("是否自定义图片格式？(y/N): ").strip().lower()

    if use_custom == 'y':
        extensions_input = input("请输入图片扩展名（用逗号分隔，例如 .jpg,.png,.gif）: ")
        extensions = [ext.strip().lower() for ext in extensions_input.split(',')]
    else:
        extensions = None

    # 确认操作
    print("\n即将执行以下操作:")
    print(f"源目录: {source}")
    print(f"目标目录: {target}")
    if extensions:
        print(f"图片格式: {extensions}")
    print("-" * 30)

    confirm = input("是否继续？(y/N): ").strip().lower()
    if confirm == 'y':
        move_images_from_nested_folders(source, target, extensions)
    else:
        print("操作已取消")




