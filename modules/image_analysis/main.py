#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : main.py 
@Author  : Shawn
@Date    : 2026/1/23 9:32 
@Info    : Description of this file
"""

from modules.image_analysis.get_xml_info import get_xml_recognition_results
from modules.image_analysis.image_compare import start_image_analysis_app
from modules.image_analysis.move_picture import interactive_mode


def main():
    """主函数"""
    chosen = input("请选择功能\n"
                   "1. 移动图片或文件\n"
                   "2. 读取xml中的识别结果\n"
                   "3. 启动图片对比标注工具\n")
    if chosen == '1':
        interactive_mode()
    elif chosen == '2':
        get_xml_recognition_results()
    elif chosen == '3':
        start_image_analysis_app()


if __name__ == "__main__":
    main()
