import pandas as pd
import os
from datetime import datetime


def extract_timestamp(filename):
    """从文件名中提取时间戳部分并转换为datetime对象"""
    # 假设文件名格式为：YYYYMMDDHHMMSSmmm.jpg
    # 例如：20250618000007286.jpg
    timestamp_str = filename.split('.')[0]  # 去掉扩展名
    try:
        return datetime.strptime(timestamp_str, "%Y%m%d%H%M%S%f")
    except ValueError:
        # 如果毫秒部分不足3位，可能需要特殊处理
        # 这里假设所有文件名都有完整的17位数字（包括毫秒）
        return None


def calculate_time_differences(filenames):
    """计算相邻文件之间的时间差（毫秒）"""
    results = []
    for i in range(1, len(filenames)):
        file1 = filenames[i - 1]
        file2 = filenames[i]

        dt1 = extract_timestamp(file1)
        dt2 = extract_timestamp(file2)

        if dt1 and dt2:
            time_diff = (dt2 - dt1).total_seconds() * 1000  # 转换为毫秒
            results.append({
                "前一个文件": file1,
                "后一个文件": file2,
                "时间间隔(ms)": time_diff
            })

    return results


def main(folder_path, output_excel):
    # 获取文件夹中的所有jpg文件并按文件名排序
    filenames = [f for f in os.listdir(folder_path) if f.lower().endswith('.jpg')]
    filenames.sort()  # 按文件名排序

    if len(filenames) < 2:
        print("文件夹中需要至少2个文件才能计算时间间隔")
        return

    # 计算时间差
    time_diffs = calculate_time_differences(filenames)

    # 创建DataFrame并导出Excel
    df = pd.DataFrame(time_diffs)
    df.to_excel(output_excel, index=False)
    print(f"结果已导出到: {output_excel}")


if __name__ == "__main__":
    folder_path = r"D:\Project\2025.04.14-AI应用\2025.04.14 模内监控\程序图片\2025-06-26\1"  # 当前文件夹，可以修改为目标文件夹路径
    output_excel = "文件时间间隔统计.xlsx"
    main(folder_path, output_excel)