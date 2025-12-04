#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : user.py 
@Author  : Shawn
@Date    : 2025/9/24 15:33 
@Info    : Description of this file
"""
import pymysql
import bcrypt
import logging
from datetime import datetime
from typing import Optional, Dict, Any
from modules.common.result import res_format

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class User:
    def __init__(self, db_config: Dict[str, Any]):
        self.db_config = db_config
        # self._create_table_if_not_exists()

    def _get_connection(self):
        """获取数据库连接"""
        return pymysql.connect(**self.db_config)

    def _hash_password(self, password: str) -> str:
        """加密密码"""
        salt = bcrypt.gensalt()
        hashed = bcrypt.hashpw(password.encode('utf-8'), salt)
        return hashed.decode('utf-8')

    def _verify_password(self, password: str, hashed_password: str) -> bool:
        """验证密码"""
        return bcrypt.checkpw(password.encode('utf-8'), hashed_password.encode('utf-8'))

    def create_user(self, username: str, password: str, **kwargs) -> Dict[str, Any]:
        """创建用户"""
        try:
            password = self._hash_password(password)
            display_name = kwargs.get('display_name') or username
            email = kwargs.get('email') or None
            avatar_url = kwargs.get('avatar_url') or None
            sex = kwargs.get('sex') or None
            status = kwargs.get('status') or 0
            is_disable = kwargs.get('is_disable')
            is_admin = kwargs.get('is_admin')
            is_delete = 0
            create_by = kwargs.get('create_by') or 1
            create_time = kwargs.get('create_time') or datetime.now()

            with (self._get_connection() as conn):
                with conn.cursor() as cursor:
                    sql = """
                    INSERT INTO user (username, password, display_name, email, avatar_url, sex, status, is_disable, is_admin, is_delete, create_by, create_time)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """
                    cursor.execute(sql, (
                        username, password, display_name, email, avatar_url, sex, status, is_disable, is_admin,
                        is_delete,
                        create_by, create_time))
                    user_id = cursor.lastrowid
                    conn.commit()

                    logger.info(f"用户创建成功: {username} (ID: {user_id})")

                    return res_format('success', 200, '用户创建成功', {'user_id': user_id})

        except pymysql.IntegrityError as e:
            error_msg = "用户名或邮箱已存在" if "Duplicate entry" in str(e) else "数据完整性错误"
            logger.error(f"创建用户失败: {error_msg}")
            return res_format('success', 200, error_msg)
        except Exception as e:
            logger.error(f"创建用户失败: {e}")
            return res_format('success', 200, e)

    def get_user(self, user_id: int = None, username: str = None, email: str = None) -> Optional[Dict[str, Any]]:
        """获取用户信息（排除密码字段）"""
        try:
            with self._get_connection() as conn:
                with conn.cursor(pymysql.cursors.DictCursor) as cursor:
                    if user_id:
                        sql = "SELECT id, username, email, display_name, avatar_url, sex, status, is_disable, is_admin, is_delete, create_by, create_time FROM user WHERE id = %s"
                        cursor.execute(sql, (user_id,))
                    elif username:
                        sql = "SELECT id, username, email, display_name, avatar_url, sex, status, is_disable, is_admin, is_delete, create_by, create_time FROM user WHERE username = %s"
                        cursor.execute(sql, (username,))
                    elif email:
                        sql = "SELECT id, username, email, display_name, avatar_url, sex, status, is_disable, is_admin, is_delete, create_by, create_time FROM user WHERE email = %s"
                        cursor.execute(sql, (email,))
                    else:
                        return None

                    user = cursor.fetchone()
                    return user

        except Exception as e:
            logger.error(f"获取用户信息失败: {e}")
            return None

    def update_user(self, user_id: int, **kwargs) -> Dict[str, Any]:
        """更新用户信息"""
        allowed_fields = ['display_name', 'avatar_url', 'status', 'email', 'sex']
        update_data = {k: v for k, v in kwargs.items() if k in allowed_fields and v is not None}

        if not update_data:
            return {'success': False, 'message': '没有有效的更新字段'}

        try:
            with self._get_connection() as conn:
                with conn.cursor() as cursor:
                    set_clause = ", ".join([f"{field} = %s" for field in update_data.keys()])
                    sql = f"UPDATE user SET {set_clause} WHERE id = %s"
                    values = list(update_data.values()) + [user_id]

                    cursor.execute(sql, values)
                    conn.commit()

                    if cursor.rowcount > 0:
                        logger.info(f"用户更新成功: ID {user_id}")
                        return {'success': True, 'message': '用户更新成功'}
                    else:
                        return {'success': False, 'message': '用户不存在或数据未变化'}

        except Exception as e:
            logger.error(f"更新用户失败: {e}")
            return {'success': False, 'message': str(e)}

    def change_password(self, user_id: int, old_password: str, new_password: str) -> Dict[str, Any]:
        """修改密码"""
        try:
            with self._get_connection() as conn:
                with conn.cursor() as cursor:
                    # 先验证旧密码
                    cursor.execute("SELECT password FROM user WHERE id = %s", (user_id,))
                    result = cursor.fetchone()

                    if not result:
                        return {'success': False, 'message': '用户不存在'}

                    if not self._verify_password(old_password, result[0]):
                        return {'success': False, 'message': '原密码错误'}

                    # 更新密码
                    new_password = self._hash_password(new_password)
                    cursor.execute("UPDATE user SET password = %s WHERE id = %s",
                                   (new_password, user_id))
                    conn.commit()

                    logger.info(f"用户密码修改成功: ID {user_id}")
                    return {'success': True, 'message': '密码修改成功'}

        except Exception as e:
            logger.error(f"修改密码失败: {e}")
            return {'success': False, 'message': str(e)}

    def delete_user(self, user_id: int) -> Dict[str, Any]:
        """删除用户"""
        try:
            with self._get_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute("UPDATE user SET is_delete = 1 WHERE id = %s", (user_id,))

                    conn.commit()

                    if cursor.rowcount > 0:
                        logger.info(f"用户删除成功: ID {user_id}")
                        return {'success': True, 'message': '用户删除成功'}
                    else:
                        return {'success': False, 'message': '用户不存在'}

        except Exception as e:
            logger.error(f"删除用户失败: {e}")
            return {'success': False, 'message': str(e)}

    def verify_login(self, username: str, password: str) -> Dict[str, Any]:
        """验证用户登录"""
        try:
            with self._get_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute(
                        "SELECT id, password, status, is_disable FROM user WHERE username = %s",
                        (username,))
                    result = cursor.fetchone()

                    if not result:
                        return {'success': False, 'message': '用户不存在'}

                    user_id, password_hash, status, is_disable = result

                    if status != 0 or is_disable == 1:
                        return {'success': False, 'message': '用户账户已被禁用'}

                    if not self._verify_password(password, password_hash):
                        # 记录登录失败次数
                        cursor.execute("UPDATE user SET login_attempts = login_attempts + 1 WHERE id = %s", (user_id,))
                        conn.commit()
                        return {'success': False, 'message': '密码错误'}

                    # 登录成功，更新最后登录时间并重置登录尝试次数
                    cursor.execute("UPDATE user SET last_login = %s, login_attempts = 0 WHERE id = %s",
                                   (datetime.now(), user_id))
                    conn.commit()

                    user_info = self.get_user(user_id=user_id)
                    return {
                        'success': True,
                        'message': '登录成功',
                        'user': user_info
                    }

        except Exception as e:
            logger.error(f"登录验证失败: {e}")
            return {'success': False, 'message': str(e)}


from config.kawang import DATABASE


def main():
    # 初始化用户管理器
    user_manager = User(DATABASE)

    chosen = input("功能菜单如下，请选择：\n"
                   "1. 创建用户\n"
                   "2. 获取用户信息\n"
                   "3. 更新用户信息\n"
                   "4. 验证登录\n"
                   "5. 修改密码\n"
                   "6. 删除用户\n")

    if chosen == '1':
        # 1. 创建用户
        print("=== 创建用户 ===")
        user_info = {
            "email": "shawn.zzz@foxmail.com",
            "display_name": "shawn",
            "avatar_url": "https://example.com/avatar.jpg"
        }
        result = user_manager.create_user(
            username="shawn",
            password="meiyoumima",
            **user_info,
        )
        # print(result)

    if chosen == '2':
        # 2. 获取用户信息
        print("\n=== 获取用户信息 ===")
        user_id = int(input("\n请输入用户id"))
        user = user_manager.get_user(user_id=user_id)
        print(user)

    if chosen == '3':
        # 3. 更新用户信息
        print("\n=== 更新用户信息 ===")
        user_id = int(input("\n请输入用户id"))
        update_result = user_manager.update_user(
            user_id=user_id,
            display_name="John Smith",
            avatar_url="https://example.com/new_avatar.jpg"
        )
        print(update_result)

    if chosen == "4":
        # 4. 验证登录
        print("\n=== 验证登录 ===")
        username = input("请输入用户名\n").strip()
        password = input("请输入密码：\n").strip()
        login_result = user_manager.verify_login(username, password)
        print(login_result)

    if chosen == "5":
        # 5. 修改密码
        print("\n=== 修改密码 ===")
        user_id = int(input("\n请输入用户id"))
        change_pw_result = user_manager.change_password(
            user_id=user_id,
            old_password="password123",
            new_password="admin"
        )
        print(change_pw_result)

    if chosen == "6":
        # 6. 删除用户
        print("\n=== 删除用户 ===")
        user_id = int(input("\n请输入用户id"))
        delete_result = user_manager.delete_user(user_id)
        print(delete_result)


if __name__ == "__main__":
    main()
