#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : main.py 
@Author  : Shawn
@Date    : 2026/1/29 1:02 
@Info    : Description of this file
"""

import tkinter as tk
from tkinter import ttk

from modules.cm_ogpData.app_ui import DataSorterApp


def main():
    root = tk.Tk()

    # 配置进度条样式
    style = ttk.Style()
    style.theme_use('clam')
    style.configure("red.Horizontal.TProgressbar",
                    troughcolor='white',
                    background='red',
                    bordercolor='gray',
                    lightcolor='red',
                    darkcolor='red')

    app = DataSorterApp(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except ImportError as e:
        print(f"错误: 缺少必要的库 - {e}")
        print("请运行: pip install chardet")
        input("按Enter键退出...")