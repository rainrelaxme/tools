#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@Project : tools
@File    : kawang.py
@Author  : Shawn
@Date    : 2025/9/24 11:53
@Info    : 配置文件，快速切换环境（示例）
"""
import os


class Config(object):  # 默认配置
    DEBUG = False
    DS_KEY = "sk-a950c7b878a940be93136750a9a47860"  # DeepSeek KEY

    # get attribute
    def __getitem__(self, key):
        return self.__getattribute__(key)


class DevConfig(Config):  # 开发环境
    HOST = '127.0.0.1'
    ROOT_PATH = r"D:\Code\Project\tools"    # 项目根目录


class ProdConfig(Config):  # 生产环境
    HOST = '172.16.17.33'
    PORT = 3306
    ROOT_PATH = r"D:\Code\Project\tools"    # 项目根目录


# 环境映射关系
mapping = {
    'dev': DevConfig,
    'prod': ProdConfig,
    'default': DevConfig
}

# 一键切换环境
APP_ENV = os.environ.get('APP_ENV', 'dev').lower()  # 设置环境变量
config = mapping[APP_ENV]()


# 根据脚本参数，来决定用那个环境配置
# import sys
# # print(sys.argv)
# num = len(sys.argv) - 1  #参数个数
# if num < 1 or num > 1:
#     exit("参数错误,必须传环境变量!比如: python xx.py dev|test|prod|default")
#
# env = sys.argv[1]  # 环境
# # print(env)
# APP_ENV = os.environ.get('APP_ENV', env).lower()
# config = mapping[APP_ENV]()  # 实例化对应的环境
