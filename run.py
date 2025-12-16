#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@Project : tools
@File    : run.py
@Author  : Shawn
@Date    : 2025/9/24 11:53
@Info    : Description of this file
"""
import modules.cm_sop_translate.main as cm_sop_translate


def main():
    """功能选择器"""
    print("\n" + "=" * 100)
    print("\x20" * 45 + "功能选择器")
    print("=" * 100)

    while True:
        print("请选择使用方式(请输入序号)：\n"
              "1. 诚模sop翻译工具\n"
              "2. 文档翻译\n"
              "3. excel翻译\n"
              "4. 退出")

        option = input().strip()
        if option == '1':
            cm_sop_translate.main()
        else:
            print("输入错误，请重新输入")


if __name__ == '__main__':
    main()
