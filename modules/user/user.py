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
            display_name = display_name or username

            with self._get_connection() as conn:
                with conn.cursor() as cursor:
                    sql = """
                    INSERT INTO user (username, password, email, display_name, sex, status, is_able, is_delete, create_by, create_time)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """
                    cursor.execute(sql, (username, email, password, display_name, avatar_url))
                    user_id = cursor.lastrowid
                    conn.commit()

                    logger.info(f"用户创建成功: {username} (ID: {user_id})")

                    return {
                        'success': True,
                        'user_id': user_id,
                        'message': '用户创建成功'
                    }

        except pymysql.IntegrityError as e:
            error_msg = "用户名或邮箱已存在" if "Duplicate entry" in str(e) else "数据完整性错误"
            logger.error(f"创建用户失败: {error_msg}")
            return {'success': False, 'message': error_msg}
        except Exception as e:
            logger.error(f"创建用户失败: {e}")
            return {'success': False, 'message': str(e)}

    def get_user(self, user_id: int = None, username: str = None, email: str = None) -> Optional[Dict[str, Any]]:
        """获取用户信息（排除密码字段）"""
        try:
            with self._get_connection() as conn:
                with conn.cursor(pymysql.cursors.DictCursor) as cursor:
                    if user_id:
                        sql = "SELECT id, username, email, display_name, avatar_url, status, email_verified, last_login_at, created_at, updated_at FROM user WHERE id = %s"
                        cursor.execute(sql, (user_id,))
                    elif username:
                        sql = "SELECT id, username, email, display_name, avatar_url, status, email_verified, last_login_at, created_at, updated_at FROM user WHERE username = %s"
                        cursor.execute(sql, (username,))
                    elif email:
                        sql = "SELECT id, username, email, display_name, avatar_url, status, email_verified, last_login_at, created_at, updated_at FROM user WHERE email = %s"
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
        allowed_fields = ['display_name', 'avatar_url', 'status', 'email_verified']
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
                    cursor.execute("DELETE FROM user WHERE id = %s", (user_id,))
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
                    cursor.execute("SELECT id, password, status FROM user WHERE username = %s OR email = %s",
                                   (username, username))
                    result = cursor.fetchone()

                    if not result:
                        return {'success': False, 'message': '用户不存在'}

                    user_id, password, status = result

                    if status != 'active':
                        return {'success': False, 'message': '用户账户已被禁用'}

                    if not self._verify_password(password, password):
                        # 记录登录失败次数
                        cursor.execute("UPDATE user SET login_attempts = login_attempts + 1 WHERE id = %s", (user_id,))
                        conn.commit()
                        return {'success': False, 'message': '密码错误'}

                    # 登录成功，更新最后登录时间并重置登录尝试次数
                    cursor.execute("UPDATE user SET last_login_at = %s, login_attempts = 0 WHERE id = %s",
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

    choosen = input("功能菜单如下，请选择：\n"
                    "1. 创建用户\n"
                    "2. 获取用户信息\n"
                    "3. 更新用户信息\n"
                    "4. 验证登录\n"
                    "5. 修改密码\n"
                    "6. 删除用户\n")

    if choosen == '1':
        # 1. 创建用户
        print("=== 创建用户 ===")
        result = user_manager.create_user(
            username="john_doe",
            email="john@example.com",
            password="securepassword123",
            display_name="John Doe",
            avatar_url="https://example.com/avatar.jpg"
        )
        print(result)

    if choosen == '2':
        # 2. 获取用户信息
        print("\n=== 获取用户信息 ===")
        user_id = int(input("\n请输入用户id"))
        user = user_manager.get_user(user_id=user_id)
        print(user)

    if choosen == '3':
        # 3. 更新用户信息
        print("\n=== 更新用户信息 ===")
        user_id = int(input("\n请输入用户id"))
        update_result = user_manager.update_user(
            user_id=user_id,
            display_name="John Smith",
            avatar_url="https://example.com/new_avatar.jpg"
        )
        print(update_result)

    if choosen == "4":
        # 4. 验证登录
        print("\n=== 验证登录 ===")
        user_id = int(input("\n请输入用户id"))
        login_result = user_manager.verify_login("john_doe", "securepassword123")
        print(login_result)

    if choosen == "5":
        # 5. 修改密码
        print("\n=== 修改密码 ===")
        user_id = int(input("\n请输入用户id"))
        change_pw_result = user_manager.change_password(
            user_id=user_id,
            old_password="securepassword123",
            new_password="newsecurepassword456"
        )
        print(change_pw_result)

    if choosen == "6":
        # 6. 删除用户
        print("\n=== 删除用户 ===")
        user_id = int(input("\n请输入用户id"))
        delete_result = user_manager.delete_user(user_id)
        print(delete_result)


if __name__ == "__main__":
    main()
