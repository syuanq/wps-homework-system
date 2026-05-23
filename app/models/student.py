# -*- coding: utf-8 -*-
"""
学生模型
"""
import uuid
from datetime import datetime
from app.models.database import db
from app.models.user import User


class Student:
    """学生模型类"""

    @classmethod
    def get_all(cls):
        """获取所有学生"""
        rows = db.fetchall("SELECT * FROM students ORDER BY created_at DESC")
        return [dict(row) for row in rows]

    @classmethod
    def get_by_id(cls, student_id):
        """根据ID获取学生"""
        row = db.fetchone("SELECT * FROM students WHERE id = ?", (student_id,))
        if row:
            return dict(row)
        return None

    @classmethod
    def get_by_student_no(cls, student_no):
        """根据学号获取学生"""
        row = db.fetchone("SELECT * FROM students WHERE student_no = ?", (student_no,))
        if row:
            return dict(row)
        return None

    @classmethod
    def create(cls, name, student_no=None, class_name=None, **kwargs):
        """创建学生"""
        student_id = str(uuid.uuid4())[:8]

        # 检查学号是否重复
        if student_no:
            existing = cls.get_by_student_no(student_no)
            if existing:
                return None

        db.execute(
            """INSERT INTO students (id, name, student_no, class_name, gender, phone, email, remark)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (student_id, name.strip(), student_no.strip() if student_no else None,
             class_name.strip() if class_name else None,
             kwargs.get('gender'), kwargs.get('phone'), kwargs.get('email'), kwargs.get('remark'))
        )

        # 自动创建登录账号
        if student_no:
            cls._create_user_account(student_no, name)

        return cls.get_by_id(student_id)

    @classmethod
    def update(cls, student_id, **kwargs):
        """更新学生信息"""
        allowed_fields = ['name', 'student_no', 'class_name', 'gender', 'phone', 'email', 'remark']
        updates = []
        params = []

        for field in allowed_fields:
            if field in kwargs:
                updates.append(f"{field} = ?")
                value = kwargs[field]
                if value and isinstance(value, str):
                    value = value.strip()
                params.append(value)

        if not updates:
            return None

        params.append(student_id)
        sql = f"UPDATE students SET {', '.join(updates)}, updated_at = CURRENT_TIMESTAMP WHERE id = ?"
        db.execute(sql, params)
        return cls.get_by_id(student_id)

    @classmethod
    def delete(cls, student_id):
        """删除学生"""
        student = cls.get_by_id(student_id)
        if student:
            student_no = student.get('student_no')
            if student_no:
                cls._delete_user_account(student_no)
            db.execute("DELETE FROM students WHERE id = ?", (student_id,))
            return True
        return False

    @classmethod
    def batch_delete(cls, student_ids):
        """批量删除学生"""
        deleted_count = 0
        for sid in student_ids:
            if cls.delete(sid):
                deleted_count += 1
        return deleted_count

    @classmethod
    def get_class_list(cls):
        """获取所有班级列表"""
        rows = db.fetchall("SELECT DISTINCT class_name FROM students WHERE class_name IS NOT NULL AND class_name != '' ORDER BY class_name")
        return [row['class_name'] for row in rows if row['class_name']]

    @classmethod
    def get_by_class(cls, class_name):
        """获取指定班级的学生"""
        rows = db.fetchall("SELECT * FROM students WHERE class_name = ? ORDER BY name", (class_name,))
        return [dict(row) for row in rows]

    @classmethod
    def search(cls, keyword=''):
        """搜索学生"""
        if not keyword:
            return cls.get_all()

        keyword = f"%{keyword.lower()}%"
        rows = db.fetchall(
            """SELECT * FROM students WHERE
               LOWER(name) LIKE ? OR
               LOWER(COALESCE(student_no, '')) LIKE ? OR
               LOWER(COALESCE(class_name, '')) LIKE ?
               ORDER BY name""",
            (keyword, keyword, keyword)
        )
        return [dict(row) for row in rows]

    @classmethod
    def get_stats(cls):
        """获取学生统计信息"""
        total = db.fetchone("SELECT COUNT(*) as count FROM students")['count']
        classes = cls.get_class_list()

        class_distribution = {}
        for c in classes:
            count = db.fetchone("SELECT COUNT(*) as count FROM students WHERE class_name = ?", (c,))['count']
            class_distribution[c] = count

        return {
            'total': total,
            'class_count': len(classes),
            'classes': classes,
            'class_distribution': class_distribution
        }

    @classmethod
    def _create_user_account(cls, student_no, name):
        """自动创建登录账号"""
        if not student_no:
            return

        password = student_no[-6:] if len(student_no) >= 6 else student_no
        User.create(student_no, password, name, 'student', '学生', None)

    @classmethod
    def _delete_user_account(cls, student_no):
        """删除登录账号"""
        if student_no:
            User.delete(student_no)

    @classmethod
    def sync_user_accounts(cls):
        """为所有学生同步登录账号"""
        students = cls.get_all()
        created = 0
        for s in students:
            student_no = s.get('student_no')
            name = s.get('name')
            if student_no and name:
                user = User.get_by_username(student_no)
                if not user:
                    cls._create_user_account(student_no, name)
                    created += 1
        return created
