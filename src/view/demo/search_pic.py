import os
import shutil
from datetime import datetime


def copy_whole_hour_images(source_folder, target_folder):
    """
    将创建时间是整点的图片复制到目标文件夹

    Args:
        source_folder (str): 源文件夹路径
        target_folder (str): 目标文件夹路径
    """
    # 确保目标文件夹存在
    os.makedirs(target_folder, exist_ok=True)

    # 支持的图片扩展名
    image_extensions = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff')

    # 遍历源文件夹
    for filename in os.listdir(source_folder):
        # 检查是否是图片文件
        if filename.lower().endswith(image_extensions):
            filepath = os.path.join(source_folder, filename)

            # 获取文件创建时间（Windows）或最后修改时间（其他系统）
            # if os.name == 'nt':
            #     timestamp = os.path.getctime(filepath)
            # else:
            timestamp = os.path.getmtime(filepath)

            # 转换为datetime对象
            create_time = datetime.fromtimestamp(timestamp)

            # 检查是否是整点（分钟和秒都为0）
            if create_time.minute == 0:
                # 构建目标路径
                target_path = os.path.join(target_folder, filename)

                # 复制文件
                shutil.copy2(filepath, target_path)
                print(f"Copied: {filename} (Created at: {create_time})")


if __name__ == "__main__":
    # 设置源文件夹和目标文件夹路径
    source_dir = f"D:/Project/2025.04.14-AI应用/2025.04.14 模内监控/程序图片/2025-06-20/0/good"

    target_dir = f"C:/Users/CM001620/Desktop/图像对齐测试图片/新建文件夹"

    # 执行复制操作
    copy_whole_hour_images(source_dir, target_dir)
    print("Operation completed!")