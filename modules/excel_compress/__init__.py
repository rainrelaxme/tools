#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : __init__.py.py 
@Author  : Shawn
@Date    : 2025/10/12 17:14 
@Info    : Description of this file
"""

import datetime


def get_datetime():
    time = datetime.datetime.now().strftime('%y%m%d%H%M%S')
    return time


if __name__ == '__main__':
    print(f"********************start at {get_datetime()}********************")
    print("主程序执行中...")

    print(f"********************end at {get_datetime()}**********************")
