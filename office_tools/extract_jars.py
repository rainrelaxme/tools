# -*- coding:utf-8 -*-
# authored by rainrelaxme
# 解压目录下的所有压缩文件，以jar结尾的

import os
import zipfile
import shutil


def extract_and_remove_jar_files(root_dir):
    """
    递归解压所有子目录中的JAR文件到原目录，并删除原始JAR文件
    :param root_dir: 要搜索的根目录
    """
    for root, dirs, files in os.walk(root_dir):
        for file in files:
            if file.lower().endswith('.jar'):
                jar_path = os.path.join(root, file)
                extract_dir = os.path.splitext(jar_path)[0]  # 去掉.jar后缀作为解压目录

                print(f"正在处理: {jar_path}")

                try:
                    # 创建解压目录
                    os.makedirs(extract_dir, exist_ok=True)

                    # 解压JAR文件
                    with zipfile.ZipFile(jar_path, 'r') as zip_ref:
                        zip_ref.extractall(extract_dir)

                    # 验证解压是否成功（至少有一个文件被解压）
                    if os.listdir(extract_dir):
                        # 删除原始JAR文件
                        os.remove(jar_path)
                        print(f"成功解压并删除: {jar_path}")
                    else:
                        # 解压失败，删除空目录
                        shutil.rmtree(extract_dir)
                        print(f"警告: {jar_path} 解压后为空，已保留原始文件")

                except Exception as e:
                    print(f"处理 {jar_path} 失败: {str(e)}")
                    # 清理可能已创建的空目录
                    if os.path.exists(extract_dir) and not os.listdir(extract_dir):
                        shutil.rmtree(extract_dir)


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 2:
        print("用法: python extract_and_remove_jars.py <目录路径>")
        sys.exit(1)

    target_dir = sys.argv[1]

    if not os.path.isdir(target_dir):
        print(f"错误: {target_dir} 不是有效目录")
        sys.exit(1)

    print(f"开始处理目录: {target_dir}")
    extract_and_remove_jar_files(target_dir)
    print("处理完成")