#!/usr/bin/env python3
# -*- coding:utf-8 -*-
"""
@Project: tools
@File   : db_import.py
@Version:
@Author : RainRelaxMe
@Date   : 2025/9/2 0:13
@Info   : 欧普模具系统导入模具清单
"""

import pandas as pd
import pymysql
from datetime import datetime
import os

# 数据库配置
DB_CONFIG = {
    'host': '10.10.57.182',
    'user': 'opple_mplm_dev',
    'password': '2N8hJAQS?cJJ',
    'database': 'test',
    'port': 3306,
    'charset': 'utf8mb4'
}

# Excel文件路径
EXCEL_FILE = 'your_data.xlsx'

# 备份表后缀格式
BACKUP_SUFFIX = "_backup_%Y%m%d_%H%M%S"


def get_db_connection():
    """创建数据库连接"""
    try:
        conn = pymysql.connect(**DB_CONFIG)
        return conn
    except Exception as e:
        print(f"数据库连接失败: {e}")
        return None


def create_backup_table(conn, table_name):
    """创建备份表"""
    try:
        cursor = conn.cursor()

        # 生成带时间戳的备份表名
        backup_table = table_name + datetime.now().strftime(BACKUP_SUFFIX)

        # 创建备份表（结构和数据）
        cursor.execute(f"CREATE TABLE {backup_table} LIKE {table_name}")
        cursor.execute(f"INSERT INTO {backup_table} SELECT * FROM {table_name}")

        conn.commit()
        print(f"已创建备份表: {backup_table}")
        return backup_table

    except Exception as e:
        conn.rollback()
        print(f"创建备份表失败: {e}")
        return None


def read_excel_data(file_path):
    """读取Excel数据"""
    try:
        # 检查文件是否存在
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Excel文件不存在: {file_path}")

        # 假设Excel有两个sheet，分别对应两个表
        dc_mould_df = pd.read_excel(file_path, sheet_name='dc_mould')
        asset_card_df = pd.read_excel(file_path, sheet_name='asset_card_info')

        # 简单数据校验
        if 'MOULD_CODE' not in dc_mould_df.columns:
            raise ValueError("dc_mould sheet中缺少MOULD_CODE列")

        if 'ASSET_CODE' not in asset_card_df.columns:
            raise ValueError("asset_card_info sheet中缺少ASSET_CODE列")

        return dc_mould_df, asset_card_df

    except Exception as e:
        print(f"读取Excel文件失败: {e}")
        return None, None


def update_dc_mould(conn, df):
    """更新dc_mould表"""
    try:
        cursor = conn.cursor()

        # 获取Excel中的MOULD_CODE列表
        mould_codes = df['MOULD_CODE'].unique().tolist()

        # 先查询数据库中已存在的记录
        placeholders = ','.join(['%s'] * len(mould_codes))
        query = f"SELECT MOULD_CODE FROM dc_mould WHERE MOULD_CODE IN ({placeholders})"
        cursor.execute(query, mould_codes)
        existing_codes = [row[0] for row in cursor.fetchall()]

        # 准备统计数据
        stats = {'insert': 0, 'update': 0, 'skip': 0}

        for _, row in df.iterrows():
            # 检查必要字段是否存在
            if pd.isna(row['MOULD_CODE']):
                stats['skip'] += 1
                continue

            if row['MOULD_CODE'] in existing_codes:
                # 更新操作
                sql = """
                UPDATE dc_mould 
                SET MOULD_NAME = %s, MOULD_TYPE = %s, MOULD_STATUS = %s, 
                    LAST_UPDATE_TIME = %s, UPDATE_BY = %s
                WHERE MOULD_CODE = %s
                """
                cursor.execute(sql, (
                    row.get('MOULD_NAME', ''),
                    row.get('MOULD_TYPE', ''),
                    row.get('MOULD_STATUS', ''),
                    datetime.now(),
                    'system',  # 或从Excel/配置中获取
                    row['MOULD_CODE']
                ))
                stats['update'] += 1
            else:
                # 插入操作
                sql = """
                INSERT INTO dc_mould 
                (MOULD_CODE, MOULD_NAME, MOULD_TYPE, MOULD_STATUS, 
                 CREATE_TIME, LAST_UPDATE_TIME, CREATE_BY, UPDATE_BY)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                """
                cursor.execute(sql, (
                    row['MOULD_CODE'],
                    row.get('MOULD_NAME', ''),
                    row.get('MOULD_TYPE', ''),
                    row.get('MOULD_STATUS', ''),
                    datetime.now(),
                    datetime.now(),
                    'system',
                    'system'
                ))
                stats['insert'] += 1

        conn.commit()
        print(f"dc_mould表更新完成: 新增{stats['insert']}条, 更新{stats['update']}条, 跳过{stats['skip']}条")

    except Exception as e:
        conn.rollback()
        print(f"更新dc_mould表失败: {e}")
        raise


def update_asset_card_info(conn, df):
    """更新asset_card_info表"""
    try:
        cursor = conn.cursor()

        # 获取Excel中的资产编号列表(假设主键是ASSET_CODE)
        asset_codes = df['ASSET_CODE'].unique().tolist()

        # 先查询数据库中已存在的记录
        placeholders = ','.join(['%s'] * len(asset_codes))
        query = f"SELECT ASSET_CODE FROM asset_card_info WHERE ASSET_CODE IN ({placeholders})"
        cursor.execute(query, asset_codes)
        existing_codes = [row[0] for row in cursor.fetchall()]

        # 准备统计数据
        stats = {'insert': 0, 'update': 0, 'skip': 0}

        for _, row in df.iterrows():
            # 检查必要字段是否存在
            if pd.isna(row['ASSET_CODE']):
                stats['skip'] += 1
                continue

            if row['ASSET_CODE'] in existing_codes:
                # 更新操作
                sql = """
                UPDATE asset_card_info 
                SET ASSET_NAME = %s, ASSET_TYPE = %s, MOULD_CODE = %s, 
                    STATUS = %s, LAST_UPDATE_TIME = %s, UPDATE_BY = %s
                WHERE ASSET_CODE = %s
                """
                cursor.execute(sql, (
                    row.get('ASSET_NAME', ''),
                    row.get('ASSET_TYPE', ''),
                    row.get('MOULD_CODE', ''),
                    row.get('STATUS', ''),
                    datetime.now(),
                    'system',
                    row['ASSET_CODE']
                ))
                stats['update'] += 1
            else:
                # 插入操作
                sql = """
                INSERT INTO asset_card_info 
                (ASSET_CODE, ASSET_NAME, ASSET_TYPE, MOULD_CODE, 
                 STATUS, CREATE_TIME, LAST_UPDATE_TIME, CREATE_BY, UPDATE_BY)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                """
                cursor.execute(sql, (
                    row['ASSET_CODE'],
                    row.get('ASSET_NAME', ''),
                    row.get('ASSET_TYPE', ''),
                    row.get('MOULD_CODE', ''),
                    row.get('STATUS', ''),
                    datetime.now(),
                    datetime.now(),
                    'system',
                    'system'
                ))
                stats['insert'] += 1

        conn.commit()
        print(f"asset_card_info表更新完成: 新增{stats['insert']}条, 更新{stats['update']}条, 跳过{stats['skip']}条")

    except Exception as e:
        conn.rollback()
        print(f"更新asset_card_info表失败: {e}")
        raise

def test_sql(conn, df):
    try:
        cursor = conn.cursor()
        sql = """
        select * from dc_mould 
        where mould_type = "OPB"
        and deleted = 0
        """
        message = cursor.execute(sql)
        print(f"表中的数据", message)

    except Exception as e:
        conn.rollback()
        print(f"更新dc_mould表失败: {e}")
        raise


def main():
    print("=== 开始执行数据库更新程序 ===")
    print(f"当前时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # 读取Excel数据
    # print("\n正在读取Excel数据...")
    # dc_mould_df, asset_card_df = read_excel_data(EXCEL_FILE)
    # if dc_mould_df is None or asset_card_df is None:
    #     return

    # 连接数据库
    print("\n正在连接数据库...")
    conn = get_db_connection()
    if conn is None:
        return

    try:
        # 备份表
        print("\n开始备份数据库表...")
        backup_tables = []
        for table in ['dc_mould', 'asset_card_info']:
            backup_name = create_backup_table(conn, table)
            if backup_name:
                backup_tables.append(backup_name)

        if len(backup_tables) != 2:
            print("备份表创建不完整，终止操作")
            return

        # 更新表
        print("\n开始更新数据库表...")
        # if not dc_mould_df.empty:
        #     print("正在更新dc_mould表...")
        #     update_dc_mould(conn, dc_mould_df)
        #
        # if not asset_card_df.empty:
        #     print("正在更新asset_card_info表...")
        #     update_asset_card_info(conn, asset_card_df)

        test_sql(conn,0)

        print("\n所有操作已完成!")
        print(f"备份表: {', '.join(backup_tables)}")

    except Exception as e:
        print(f"\n程序执行出错: {e}")
        print("已回滚所有未提交的更改")
    finally:
        conn.close()
        print("数据库连接已关闭")


if __name__ == "__main__":
    main()