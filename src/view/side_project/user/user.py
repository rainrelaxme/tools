#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : user.py 
@Author  : Shawn
@Date    : 2025/9/24 15:33 
@Info    : Description of this file
"""

import datetime

import mysql.connector

from config.config import VALID_ACCOUNTS, TEST_DB


def get_datetime():
    time = datetime.datetime.now().strftime('%y%m%d%H%M%S')
    return time


class User:
    def __init__(self):
        pass

    def create_user(self, account, password):
        # 连接数据库
        connection = mysql.connector.connect(TEST_DB)
        if connection is None:
            return False

        cursor = connection.cursor()
        check_query = (""
                       "SELECT id FROM platform_user WHERE account = %s")


if __name__ == '__main__':
    print(f"********************start at {get_datetime()}********************")
    print("主程序执行中...")

    print(f"********************end at {get_datetime()}**********************")
