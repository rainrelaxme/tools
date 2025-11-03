#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@Project : tools
@File    : config.py
@Author  : Shawn
@Date    : 2025/9/24 11:53
@Info    : Description of this file
"""
import os
import sys

sys.path.append(r"D:\Code\Config\Private")
sys.path.append(r"F:\Code\Config\Private")
import sean

# 项目根目录
ROOT_PATH = r"F:\Code\Project\tools"

# DeepSeek KEY
DS_KEY = sean.API_KEY

# 日志位置
LOG_PATH = {
    # 'path': './logs',
    'path': os.path.join(ROOT_PATH, "logs"),
}

