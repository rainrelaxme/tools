# json增加字符内容
# 最顶层增加  "imageData": null,
import json
import os
import glob


def add_image_data_to_json(folder_path):
    """
    批量修改文件夹下的所有JSON文件，在imagePath后添加imageData: null
    """
    # 查找文件夹中的所有JSON文件
    json_files = glob.glob(os.path.join(folder_path, "*.json"))

    for json_file in json_files:
        try:
            # 读取JSON文件
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)

            # 检查是否已经存在imageData字段
            if "imageData" not in data:
                # 创建一个新的有序字典（或直接修改）
                # 由于JSON对象是无序的，我们需要重新构建
                new_data = {}

                # 复制所有字段，在imagePath后插入imageData
                for key in data:
                    new_data[key] = data[key]
                    if key == "imagePath":
                        new_data["imageData"] = None

                # 写回文件
                with open(json_file, 'w', encoding='utf-8') as f:
                    json.dump(new_data, f, indent=2, ensure_ascii=False)

                print(f"已更新: {json_file}")
            else:
                print(f"已包含imageData字段，跳过: {json_file}")

        except Exception as e:
            print(f"处理文件 {json_file} 时出错: {str(e)}")


# 使用示例
if __name__ == "__main__":
    folder_path = input("请输入包含JSON文件的文件夹路径: ")

    if os.path.exists(folder_path):
        add_image_data_to_json(folder_path)
        print("处理完成！")
    else:
        print("指定的文件夹路径不存在！")