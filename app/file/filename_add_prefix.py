# AI训练的数据集，要按照格式命名
"""
6.图片命名规范示例如下（部分示例）：
a.样本ID_相机ID_点位号_光源号_产品型号_相机组号_机台号_颜色号_时间戳_二维码.jpg
S00099_C01_P001_L3_PI127_G1_M1_Y01_20220318112324_F77207602RW1FGG3Y+LVJ7B1.jpg

b.样本ID_点位号.jpg
S00001_P001.jpg

c.样本ID_点位号_颜色号.jpg
S00001_P001_Y01.jpg

d.样本ID_相机ID_点位号_光源号_产品型号_相机组号_机台号_时间戳.jpg
S00099_C01_P001_L3_PI127_G1_M1_20220318112324.jpg
"""

import os
import glob
from app.file_content.json_edit import add_image_data_to_json, update_json_image_path


def rename_file(directory_path, filetype, is_index):
    """
    重命名目录中的图片文件

    参数:
    directory_path: 包含图片文件的目录路径
    filetype: 文件类型
    is_index： 是否增加index：0-增加，1-不增加
    """
    image_extensions = filetype
    image_files = []

    # 获取所有图片文件
    for extension in image_extensions:
        image_files.extend(glob.glob(os.path.join(directory_path, extension)))

    # 按文件名排序以确保顺序一致
    image_files.sort()

    # 开始重命名
    for index, old_path in enumerate(image_files, start=1):
        # 获取文件扩展名
        file_extension = os.path.splitext(old_path)[1]

        # 获取原文件名（不含扩展名）
        old_filename = os.path.splitext(os.path.basename(old_path))[0]

        # 生成新文件名
        if is_index == 0:
            new_filename = f"S{index:05d}-P000-{old_filename}{file_extension}"
        else:
            new_filename = f"P000_{old_filename}{file_extension}"

        new_path = os.path.join(directory_path, new_filename)

        # 重命名文件
        try:
            os.rename(old_path, new_path)
            print(f"重命名成功: {os.path.basename(old_path)} -> {new_filename}")
            
            # 如果重命名的是图片文件，查找对应的JSON文件并更新
            if file_extension.lower() in ['.bmp', '.jpg', '.jpeg', '.png', '.gif', '.tiff']:
                # 查找对应的JSON文件
                json_filename = os.path.splitext(os.path.basename(old_path))[0] + '.json'
                json_path = os.path.join(directory_path, json_filename)
                
                if os.path.exists(json_path):
                    # 更新JSON文件中的imagePath
                    update_json_image_path(json_path, os.path.basename(old_path), new_filename)
                else:
                    print(f"未找到对应的JSON文件: {json_filename}")
                    
        except Exception as e:
            print(f"重命名失败 {old_path}: {e}")


if __name__ == "__main__":
    # 指定要处理的目录路径
    target_directory = input("请输入图片所在目录路径:\n ").strip()
    is_index_input = input("请选择是否增加文件序列号前缀(S000X),输入y/n,默认是y:\n").strip()
    operate_way = input("请选择操作方式：\n"
                        "1. 仅修改图片文件\n"
                        "2. 分别修改图片+json文件，同时修改json中对应的文件名，并补全imageData: null\n"
                        "3. 仅补全json文件中imageData: null\n").strip()

    if is_index_input == "y" or is_index_input == "Y" or not is_index_input:
        # 增加序列号前缀
        is_index = 0
    else:
        is_index = 1

    # 检查目录是否存在
    if os.path.exists(target_directory) and os.path.isdir(target_directory):
        if operate_way == "1":
            rename_file(target_directory, filetype=['*.bmp', '*.jpg', '*.jpeg', '*.png', '*.gif', '*.tiff'], is_index=is_index)
        elif operate_way == "2":
            rename_file(target_directory, filetype=['*.bmp', '*.jpg', '*.jpeg', '*.png', '*.gif', '*.tiff'], is_index=is_index)
            # 处理剩余的JSON文件（如果有的话）
            add_image_data_to_json(target_directory)
            rename_file(target_directory, filetype=['*.json'], is_index=is_index)
        elif operate_way == "3":
            add_image_data_to_json(target_directory)
        print("重命名完成！")
    else:
        print("指定的目录不存在或不是有效目录！")






