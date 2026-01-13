#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : add_data_to_csv.py 
@Author  : Shawn
@Date    : 2025/12/23 13:06 
@Info    : Description of this file
"""

import csv
import os
import random
from datetime import datetime, timedelta
import time


def generate_data(start_time, rows_per_batch=1, total_rows=10):
    """
    生成模拟数据

    参数:
    start_time: 开始时间
    rows_per_batch: 每次写入的行数
    total_rows: 总共要生成的行数
    """

    # 定义表头
    headers = [
        "条码", "PointCount", "Endtime", "焊烙铁使用总次数", "焊点总数量",
        "焊锡1温度", "焊锡1时间", "焊锡1长度", "焊锡1速度",
        "焊锡2温度", "焊锡2时间", "焊锡2长度", "焊锡2速度",
        "焊锡3温度", "焊锡3时间", "焊锡3长度", "焊锡3速度",
        "焊锡4温度", "焊锡4时间", "焊锡4长度", "焊锡4速度"
    ]

    # 基础路径
    base_dir = r"D:\Code\data\MESDATA"

    # 创建基础目录
    if not os.path.exists(base_dir):
        os.makedirs(base_dir)

    current_time = start_time
    barcode = 1

    # 计算总共需要写入的批次
    total_batches = (total_rows + rows_per_batch - 1) // rows_per_batch

    print(f"开始生成数据，总共需要 {total_batches} 批次，每批 {rows_per_batch} 行数据...")

    for batch in range(total_batches):
        # 确定当前批次实际要写入的行数
        current_batch_rows = min(rows_per_batch, total_rows - batch * rows_per_batch)

        # 根据当前时间创建子文件夹和文件路径
        date_str = current_time.strftime("%y-%m-%d")
        folder_path = os.path.join(base_dir, current_time.strftime("%Y-%m"))
        file_name = f"{date_str}.csv"
        file_path = os.path.join(folder_path, file_name)

        # 创建子文件夹
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        data_to_write = []

        # 如果是文件第一次写入，先写入表头
        file_exists = os.path.exists(file_path)

        for i in range(current_batch_rows):
            # 随机生成数据
            row_data = {
                "条码": barcode,
                "PointCount": random.randint(1, 100),
                "Endtime": current_time.strftime("%Y/%m/%d %H:%M:%S"),
                "焊烙铁使用总次数": random.randint(100, 5000),
                "焊点总数量": random.randint(50, 3000),
            }

            # 生成4组焊锡数据
            for j in range(1, 5):
                row_data[f"焊锡{j}温度"] = round(random.uniform(250, 350), 1)
                row_data[f"焊锡{j}时间"] = round(random.uniform(0.5, 3.0), 2)
                row_data[f"焊锡{j}长度"] = round(random.uniform(1.0, 10.0), 2)
                row_data[f"焊锡{j}速度"] = round(random.uniform(5.0, 20.0), 2)

            # 按表头顺序整理数据
            ordered_row = [row_data[header] for header in headers]
            data_to_write.append(ordered_row)

            # 递增条码和时间
            barcode += 1
            # 每次增加随机的时间间隔（1-60秒）
            current_time += timedelta(minutes=random.randint(1, 10))

        # 写入CSV文件
        mode = 'a' if file_exists else 'w'
        with open(file_path, mode, newline='', encoding='gb2312') as file:
            writer = csv.writer(file)
            if not file_exists:
                writer.writerow(headers)
            writer.writerows(data_to_write)

        print(f"批次 {batch + 1}/{total_batches}: 已写入 {len(data_to_write)} 行数据到 {file_path}")
        print(f"最后一条数据时间: {data_to_write[-1][2]}")

        # 如果不是最后一批，等待一下（模拟实时数据生成）
        if batch < total_batches - 1:
            time.sleep(1)  # 等待1秒

    print(f"\n数据生成完成！总共生成了 {barcode - 1} 条数据")
    print(f"数据保存在 {base_dir} 目录中")


def main():
    """主函数"""
    print("CSV数据生成程序")
    print("=" * 50)

    # 获取用户输入
    try:
        n = int(input("请输入每次写入的行数 (默认1): ") or "1")
        total_rows = int(input("请输入总共要生成的行数 (默认10): ") or "10")

        # 设置开始时间
        start_time_str = input("请输入开始时间 (格式: 2025/12/18 11:21:00，直接回车使用当前时间): ")

        if start_time_str:
            start_time = datetime.strptime(start_time_str, "%Y/%m/%d %H:%M:%S")
        else:
            start_time = datetime.now()

        print(f"\n开始时间: {start_time.strftime('%Y/%m/%d %H:%M:%S')}")
        print(f"每次写入: {n} 行")
        print(f"总共生成: {total_rows} 行")

        confirm = input("\n确认开始生成数据？(y/n): ")

        if confirm.lower() == 'y':
            generate_data(start_time, n, total_rows)
        else:
            print("程序已取消")

    except ValueError as e:
        print(f"输入错误: {e}")
    except Exception as e:
        print(f"发生错误: {e}")


def simple_demo():
    """简单演示函数，直接运行生成10行数据"""
    print("运行简单演示...")
    # start_time = datetime(2025, 12, 18, 11, 21, 0)
    start_time = datetime.now()
    generate_data(start_time, rows_per_batch=1, total_rows=1)


if __name__ == "__main__":
    # 直接运行演示
    simple_demo()

    # 如果想要交互式输入，取消下面的注释
    # main()