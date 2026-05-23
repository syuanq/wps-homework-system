# -*- coding: utf-8 -*-
"""
用户模型
"""
import hashlib
from app.models.database import db, get_db_connection


class User:
    """用户模型类"""

    @staticmethod
    def _hash_password(password):
        """密码哈希"""
        return hashlib.sha256(password.encode('utf-8')).hexdigest()

    @classmethod
    def get_by_username(cls, username):
        """根据用户名获取用户"""
        row = db.fetchone("SELECT * FROM users WHERE username = ?", (username,))
        if row:
            return dict(row)
        return None

    @classmethod
    def get_all(cls):
        """获取所有用户"""
        rows = db.fetchall("SELECT * FROM users ORDER BY created_at DESC")
        return [dict(row) for row in rows]

    @classmethod
    def get_all_students(cls):
        """获取所有学生"""
        rows = db.fetchall("SELECT * FROM users WHERE role = 'student' ORDER BY created_at DESC")
        return [dict(row) for row in rows]

    @classmethod
    def get_all_teachers(cls):
        """获取所有教师"""
        rows = db.fetchall("SELECT * FROM users WHERE role IN ('teacher', 'admin') ORDER BY created_at DESC")
        return [dict(row) for row in rows]

    @classmethod
    def authenticate(cls, username, password):
        """验证用户登录"""
        user = cls.get_by_username(username)
        if not user:
            return None
        if user['password_hash'] == cls._hash_password(password):
            return {
                "username": username,
                "name": user['name'],
                "role": user['role'],
                "title": user.get('title', ''),
                "class_name": user.get('class_name', '')
            }
        return None

    @classmethod
    def create(cls, username, password, name, role, title=None, class_name=None):
        """创建用户"""
        # 检查用户名是否已存在
        if cls.get_by_username(username):
            return False, "用户名已存在"

        password_hash = cls._hash_password(password)

        if role == 'student':
            title = title or '学生'
        elif role == 'teacher':
            title = title or '教师'
        elif role == 'admin':
            title = title or '管理员'

        db.execute(
            """INSERT INTO users (username, password_hash, name, role, title, class_name)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (username, password_hash, name, role, title, class_name)
        )
        return True, "添加成功"

    @classmethod
    def update(cls, username, **kwargs):
        """更新用户信息"""
        allowed_fields = ['name', 'role', 'title', 'class_name']
        updates = []
        params = []

        for field in allowed_fields:
            if field in kwargs:
                updates.append(f"{field} = ?")
                params.append(kwargs[field])

        if not updates:
            return False

        params.append(username)
        sql = f"UPDATE users SET {', '.join(updates)}, updated_at = CURRENT_TIMESTAMP WHERE username = ?"
        db.execute(sql, params)
        return True

    @classmethod
    def change_password(cls, username, old_password, new_password):
        """修改密码"""
        user = cls.get_by_username(username)
        if not user:
            return False, "用户不存在"
        if user['password_hash'] != cls._hash_password(old_password):
            return False, "原密码错误"

        new_hash = cls._hash_password(new_password)
        db.execute(
            "UPDATE users SET password_hash = ?, updated_at = CURRENT_TIMESTAMP WHERE username = ?",
            (new_hash, username)
        )
        return True, "密码修改成功"

    @classmethod
    def delete(cls, username):
        """删除用户"""
        db.execute("DELETE FROM users WHERE username = ?", (username,))
        return True

    @classmethod
    def init_default_users(cls):
        """初始化默认用户"""
        defaults = [
            ("admin", "123456", "管理员", "admin", "管理员", None),
            ("teacher", "123456", "教师", "teacher", "教师", None),
            ("student", "123456", "学生", "student", "学生", "计算机2401班"),
        ]

        for username, password, name, role, title, class_name in defaults:
            if not cls.get_by_username(username):
                cls.create(username, password, name, role, title, class_name)
