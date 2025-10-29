#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : config.py
@Author  : Shawn
@Date    : 2025/9/24 11:53 
@Info    : Description of this file
"""
import sys

sys.path.append(r"D:\Code\Config\Private")
import sean

# DeepSeek KEY
DS_KEY = sean.API_KEY

# 预设的账号密码（在实际应用中应该使用更安全的方式存储）
VALID_ACCOUNTS = {
    "admin": "admin123",
    "user": "user123",
    "translator": "cmtech.20250924"
}

# 词库位置
GLOSSARY = {
    # 'dir': './config',
    'dir': 'D:/Code/Project/tools/config',
    'languages': {
        '英语': 'glossary_en.json',
        '越南语': 'glossary_vi.json',
    }
}
