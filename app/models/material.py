# -*- coding: utf-8 -*-
"""
素材文件模型
"""
import json
from app.models.database import db


class Material:
    """素材文件模型类"""

    @classmethod
    def get_all(cls):
        """获取所有素材"""
        rows = db.fetchall("SELECT * FROM materials ORDER BY uploaded_at DESC")
        result = []
        for row in rows:
            data = dict(row)
            try:
                data['keywords'] = json.loads(data.get('keywords', '[]'))
            except:
                data['keywords'] = []
            result.append(data)
        return result

    @classmethod
    def get_by_id(cls, material_id):
        """根据ID获取素材"""
        row = db.fetchone("SELECT * FROM materials WHERE material_id = ?", (material_id,))
        if row:
            data = dict(row)
            try:
                data['keywords'] = json.loads(data.get('keywords', '[]'))
            except:
                data['keywords'] = []
            return data
        return None

    @classmethod
    def get_by_task(cls, task_id):
        """获取指定任务的素材"""
        rows = db.fetchall("SELECT * FROM materials WHERE task_id = ? ORDER BY uploaded_at DESC", (task_id,))
        result = []
        for row in rows:
            data = dict(row)
            try:
                data['keywords'] = json.loads(data.get('keywords', '[]'))
            except:
                data['keywords'] = []
            result.append(data)
        return result

    @classmethod
    def create(cls, material_id, task_id, filename, filepath, keywords=None):
        """创建素材记录"""
        keywords_json = json.dumps(keywords, ensure_ascii=False) if keywords else '[]'

        db.execute(
            """INSERT INTO materials (material_id, task_id, filename, filepath, keywords)
               VALUES (?, ?, ?, ?, ?)""",
            (material_id, task_id, filename, filepath, keywords_json)
        )
        return cls.get_by_id(material_id)

    @classmethod
    def delete(cls, material_id):
        """删除素材记录"""
        db.execute("DELETE FROM materials WHERE material_id = ?", (material_id,))
        return True

    @classmethod
    def delete_by_task(cls, task_id):
        """删除指定任务的所有素材"""
        db.execute("DELETE FROM materials WHERE task_id = ?", (task_id,))
        return True

    @classmethod
    def get_task_materials_map(cls):
        """获取任务到素材的映射（用于兼容旧代码）"""
        rows = db.fetchall("SELECT * FROM materials ORDER BY task_id, uploaded_at")
        result = {}
        for row in rows:
            task_id = row['task_id']
            if task_id not in result:
                result[task_id] = []

            data = {
                'material_id': row['material_id'],
                'filename': row['filename'],
                'filepath': row['filepath'],
                'uploaded_at': row['uploaded_at']
            }
            try:
                data['keywords'] = json.loads(row.get('keywords', '[]'))
            except:
                data['keywords'] = []
            result[task_id].append(data)
        return result

    @classmethod
    def init_from_json(cls, json_data):
        """从JSON数据初始化"""
        for task_id, materials in json_data.items():
            for m in materials:
                material_id = m.get('material_id')
                if material_id and not cls.get_by_id(material_id):
                    cls.create(
                        material_id=material_id,
                        task_id=task_id,
                        filename=m.get('filename', ''),
                        filepath=m.get('filepath', ''),
                        keywords=m.get('keywords', [])
                    )
        return True
