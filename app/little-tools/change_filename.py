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
from app.text.json_edit import add_image_data_to_json


def rename_images(directory_path, filetype, is_index):
    """
    重命名目录中的图片文件

    参数:
    directory_path: 包含图片文件的目录路径
    filetype: 文件类型：0-图片，1-json
    is_index： 是否增加index：0-增加，1-不增加
    """
    # 支持的图片格式
    if filetype == "0":
        image_extensions = ['*.bmp', '*.jpg', '*.jpeg', '*.png', '*.gif', '*.tiff']
    elif filetype == "1":
        image_extensions = ['*.json']
        is_editJson = input("请选择是否修改json文件，补全imageData: null,输入y/n,默认是n:\n")
        if is_editJson == "y" or is_editJson == "Y":
            add_image_data_to_json(directory_path)
    else:
        image_extensions = [filetype]
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
        if is_index == "Y" or is_index == "y":
            new_filename = f"S{index:05d}-P000-{old_filename}{file_extension}"
        else:
            new_filename = f"P000_{old_filename}{file_extension}"

        new_path = os.path.join(directory_path, new_filename)

        # 重命名文件
        try:
            os.rename(old_path, new_path)
            print(f"重命名成功: {os.path.basename(old_path)} -> {new_filename}")
        except Exception as e:
            print(f"重命名失败 {old_path}: {e}")


if __name__ == "__main__":
    # 指定要处理的目录路径
    target_directory = input("请输入图片所在目录路径:\n ").strip()
    filetype_input = input("请输入要重命名的文件类型,输入[0]-图片,[1]-json,其他请输入(如 *.jpg),默认是图片:\n").strip()
    is_index_input = input("请选择是否增加文件序列号前缀(S000X),输入y/n,默认是y:\n").strip()

    if not filetype_input:  # 回车默认
        filetype = "0"
        print("使用默认值: 图片文件")
    else:
        filetype = filetype_input

    if not is_index_input:  # 回车默认
        is_index = "y"
        print("使用默认值: 添加序列号前缀")
    else:
        is_index = is_index_input

    # 检查目录是否存在
    if os.path.exists(target_directory) and os.path.isdir(target_directory):
        rename_images(target_directory, filetype=filetype, is_index=is_index)
        print("重命名完成！")
    else:
        print("指定的目录不存在或不是有效目录！")
