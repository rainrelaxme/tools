#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : config.py
@Author  : Shawn
@Date    : 2025/9/24 11:53 
@Info    : 配置文件，快速切换环境
"""

import os

from shawn.kawang import config as conf


class Config(object):  # 默认配置
    def __getitem__(self, key):
        # get attribute
        return self.__getattribute__(key)

    DEBUG = False

    # DeepSeek KEY
    DS_KEY = conf.DS_KEY

    # 预设的账号密码（在实际应用中应该使用更安全的方式存储）
    VALID_ACCOUNTS = {
        "admin": "admin",
        "user": "user123",
    }


class DevConfig(Config):  # 开发环境
    # 项目根目录
    ROOT_PATH = r'D:\Code\Project\tools\modules\cm_sop_translate'
    # 词库位置
    GLOSSARY = {
        'dir': os.path.join(ROOT_PATH, 'config'),
        'languages': {
            '英语': 'dev_glossary_en.json',
            '越南语': 'dev_glossary_vi.json',
        }
    }
    # 日志位置
    LOG_PATH = os.path.join(ROOT_PATH, "logs")
    # 模板内容
    TEMPLATE = os.path.join(ROOT_PATH, 'config/dev_template.json')
    # 缓存文件
    TEMP_PATH = os.path.join(ROOT_PATH, 'temp')
    # database
    DATABASE = {
        'host': '106.54.47.212',
        'port': 3306,
        'user': 'root',
        'password': 'MCQM#Yx1nghe!',
        'database': 'tool'
    }


class TestConfig(Config):  # 测试环境
    pass


class ProdConfig(Config):  # 生产环境
    # 词库位置
    GLOSSARY = {
        'dir': './_internal/config',
        'languages': {
            '英语': 'glossary_en.json',
            '越南语': 'glossary_vi.json',
        }
    }
    # 日志位置
    LOG_PATH = './logs'
    # 模板内容
    TEMPLATE = './_internal/config/template.json'
    # 缓存文件
    TEMP_PATH = './temp'
    # database
    DATABASE = {
        'host': '106.54.47.212',
        'port': 3306,
        'user': 'root',
        'password': 'MCQM#Yx1nghe!',
        'database': 'tool'
    }


# 环境映射关系
mapping = {
    'dev': DevConfig,
    'test': TestConfig,
    'prod': ProdConfig,
    'default': DevConfig
}

# 一键切换环境
# APP_ENV = os.environ.get('APP_ENV', 'dev').lower()  # 设置环境变量
APP_ENV = os.environ.get('APP_ENV', 'prod').lower()  # 设置环境变量

config = mapping[APP_ENV]()
