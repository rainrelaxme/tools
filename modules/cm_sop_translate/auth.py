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

from conf.conf import VALID_ACCOUNTS


def login():
    """登录验证函数"""
    max_attempts = 3
    attempts = 0

    while attempts < max_attempts:
        username = input("请输入用户名: ").strip()
        password = getpass.getpass("请输入密码: ").strip()

        # 验证账号密码
        if username in VALID_ACCOUNTS and VALID_ACCOUNTS[username] == password:
            print(f"\n✅ 登录成功！欢迎 {username}！")
            return True
        else:
            attempts += 1
            remaining_attempts = max_attempts - attempts
            print(f"\n❌ 用户名或密码错误！剩余尝试次数: {remaining_attempts}")

            if remaining_attempts > 0:
                print("请重新输入...")
            else:
                print("\n🚫 登录失败次数过多，程序退出！")
                return False

    return False


def check_license():
    """简单的许可证检查（可选功能）"""
    import datetime as dt
    expiry_date = dt.datetime(2025, 10, 31)  # 设置过期时间

    if dt.datetime.now() > expiry_date:
        print("🚫 软件许可证已过期，请联系管理员！")
        return False
    return True