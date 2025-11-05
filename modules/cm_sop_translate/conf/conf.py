#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : conf.py
@Author  : Shawn
@Date    : 2025/9/24 11:53 
@Info    : Description of this file
"""

import os

from config.config import DS_KEY, ROOT_PATH

# DeepSeek KEY
DS_KEY = DS_KEY

# 预设的账号密码（在实际应用中应该使用更安全的方式存储）
VALID_ACCOUNTS = {
    "admin": "admin",
    "user": "user123",
}

# .env.dev
# 词库位置
GLOSSARY = {
    'dir': os.path.join(ROOT_PATH, 'modules/cm_sop_translate/conf'),
    'languages': {
        '英语': 'glossary_en.json',
        '越南语': 'glossary_vi.json',
    }
}
# 日志位置
LOG_PATH = os.path.join(ROOT_PATH, "logs")
# 模板内容
TEMPLATE = os.path.join(ROOT_PATH, 'modules/cm_sop_translate/conf/template.json')


# .env.prd
# # 词库位置
# GLOSSARY = {
#     'dir': './_internal/config',
#     'languages': {
#         '英语': 'glossary_en.json',
#         '越南语': 'glossary_vi.json',
#     }
# }
# # 日志位置
# LOG_PATH = './logs'
# # 模板内容
# TEMPLATE = './_internal/config/template.json'
