#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : auth.py 
@Author  : Shawn
@Date    : 2025/10/14 18:16 
@Info    : Description of this file
"""

import getpass

from modules.user.user import User
from modules.cm_sop_translate.config.config import config

DATABASE = config.DATABASE


def login():
    """登录验证函数"""
    user = User(DATABASE)

    username = input("请输入用户名: ").strip()
    password = getpass.getpass("请输入密码: ").strip()
    # username = 'admin'
    # password = 'admin'

    # 验证账号密码
    login_result = user.verify_login(username, password)
    print(login_result['message'])
    return login_result.get('success')


def check_license():
    """简单的许可证检查（可选功能）"""
    import datetime as dt
    expiry_date = dt.datetime(2026, 2, 10)  # 设置过期时间

    if dt.datetime.now() > expiry_date:
        print("🚫 软件许可证已过期，请联系管理员！")
        return False
    return True
