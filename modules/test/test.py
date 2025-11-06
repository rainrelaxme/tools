from modules.user.user import User
from config.kawang import DATABASE


def main():
    user = User(DATABASE)

    print("功能菜单如下：\n"
          "1. 创建用户\n"
          "2. 获取用户信息\n"
          "3. 更新用户信息\n"
          "4. \n"
          "5. \n")
