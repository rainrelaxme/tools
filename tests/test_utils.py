#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : test_utils.py 
@Author  : Shawn
@Date    : 2025/10/12 15:17 
@Info    : 单元测试等测试用例，pytest 或 unittest
"""

from src.test.test import get_footer_content


def test_add():
    assert get_footer_content(1) == 3

# 运行
# pytest tests/
