# -*- coding: utf-8 -*-
"""
答案文件模型
"""
from app.models.database import db


class Answer:
    """答案文件模型类"""

    @classmethod
    def get_all(cls):
        """获取所有答案"""
        rows = db.fetchall("SELECT * FROM answers ORDER BY uploaded_at DESC")
        return [dict(row) for row in rows]

    @classmethod
    def get_by_id(cls, answer_id):
        """根据ID获取答案"""
        row = db.fetchone("SELECT * FROM answers WHERE answer_id = ?", (answer_id,))
        if row:
            return dict(row)
        return None

    @classmethod
    def get_by_task(cls, task_id):
        """获取指定任务的答案"""
        rows = db.fetchall("SELECT * FROM answers WHERE task_id = ? ORDER BY uploaded_at DESC", (task_id,))
        return [dict(row) for row in rows]

    @classmethod
    def create(cls, answer_id, task_id, filename, filepath):
        """创建答案记录"""
        db.execute(
            """INSERT INTO answers (answer_id, task_id, filename, filepath)
               VALUES (?, ?, ?, ?)""",
            (answer_id, task_id, filename, filepath)
        )
        return cls.get_by_id(answer_id)

    @classmethod
    def delete(cls, answer_id):
        """删除答案记录"""
        db.execute("DELETE FROM answers WHERE answer_id = ?", (answer_id,))
        return True

    @classmethod
    def delete_by_task(cls, task_id):
        """删除指定任务的所有答案"""
        db.execute("DELETE FROM answers WHERE task_id = ?", (task_id,))
        return True

    @classmethod
    def get_task_answers_map(cls):
        """获取任务到答案的映射（用于兼容旧代码）"""
        rows = db.fetchall("SELECT * FROM answers ORDER BY task_id, uploaded_at")
        result = {}
        for row in rows:
            task_id = row['task_id']
            if task_id not in result:
                result[task_id] = []

            result[task_id].append({
                'answer_id': row['answer_id'],
                'filename': row['filename'],
                'filepath': row['filepath'],
                'uploaded_at': row['uploaded_at']
            })
        return result

    @classmethod
    def init_from_json(cls, json_data):
        """从JSON数据初始化"""
        for task_id, answers in json_data.items():
            for a in answers:
                answer_id = a.get('answer_id')
                if answer_id and not cls.get_by_id(answer_id):
                    cls.create(
                        answer_id=answer_id,
                        task_id=task_id,
                        filename=a.get('filename', ''),
                        filepath=a.get('filepath', '')
                    )
        return True
