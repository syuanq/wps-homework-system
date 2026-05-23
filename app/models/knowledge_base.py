# -*- coding: utf-8 -*-
"""
知识库模型
"""
import json
from app.models.database import db


class KnowledgeBase:
    """知识库模型类"""

    @classmethod
    def get_all_by_module(cls, module_group):
        """获取指定模块的所有知识"""
        rows = db.fetchall(
            "SELECT * FROM knowledge_base WHERE module_group = ? ORDER BY content_type, task_id",
            (module_group,)
        )
        return [dict(row) for row in rows]

    @classmethod
    def get_by_task(cls, task_id):
        """获取指定任务的知识"""
        rows = db.fetchall(
            "SELECT * FROM knowledge_base WHERE task_id = ? ORDER BY content_type",
            (task_id,)
        )
        return [dict(row) for row in rows]

    @classmethod
    def get_concepts(cls, module_group):
        """获取核心概念"""
        rows = db.fetchall(
            "SELECT content FROM knowledge_base WHERE module_group = ? AND content_type = 'concept'",
            (module_group,)
        )
        concepts = []
        for row in rows:
            try:
                content = json.loads(row['content'])
                if isinstance(content, list):
                    concepts.extend(content)
                else:
                    concepts.append(content)
            except:
                concepts.append(row['content'])
        return concepts

    @classmethod
    def get_operations(cls, module_group):
        """获取操作步骤"""
        rows = db.fetchall(
            "SELECT content FROM knowledge_base WHERE module_group = ? AND content_type = 'operation'",
            (module_group,)
        )
        operations = []
        for row in rows:
            try:
                content = json.loads(row['content'])
                operations.append(content)
            except:
                operations.append({'name': '操作', 'steps': [row['content']]})
        return operations

    @classmethod
    def get_mistakes(cls, module_group):
        """获取常见错误"""
        rows = db.fetchall(
            "SELECT content FROM knowledge_base WHERE module_group = ? AND content_type = 'mistake'",
            (module_group,)
        )
        mistakes = []
        for row in rows:
            try:
                content = json.loads(row['content'])
                if isinstance(content, list):
                    mistakes.extend(content)
                else:
                    mistakes.append(content)
            except:
                mistakes.append(row['content'])
        return mistakes

    @classmethod
    def get_tips(cls, module_group):
        """获取技巧提示"""
        rows = db.fetchall(
            "SELECT content FROM knowledge_base WHERE module_group = ? AND content_type = 'tip'",
            (module_group,)
        )
        tips = []
        for row in rows:
            try:
                content = json.loads(row['content'])
                if isinstance(content, list):
                    tips.extend(content)
                else:
                    tips.append(content)
            except:
                tips.append(row['content'])
        return tips

    @classmethod
    def get_task_knowledge(cls, task_id):
        """获取任务相关知识"""
        row = db.fetchone(
            "SELECT content FROM knowledge_base WHERE task_id = ? AND content_type = 'task_knowledge'",
            (task_id,)
        )
        if row:
            try:
                return json.loads(row['content'])
            except:
                return {'chapter': '', 'key_points': [], 'skills': []}
        return None

    @classmethod
    def add(cls, module_group, module_name, content_type, content, task_id=None):
        """添加知识条目"""
        if isinstance(content, (list, dict)):
            content = json.dumps(content, ensure_ascii=False)

        db.execute(
            """INSERT INTO knowledge_base (module_group, module_name, content_type, task_id, content)
               VALUES (?, ?, ?, ?, ?)""",
            (module_group, module_name, content_type, task_id, content)
        )
        return True

    @classmethod
    def update(cls, id, **kwargs):
        """更新知识条目"""
        allowed_fields = ['module_group', 'module_name', 'content_type', 'task_id', 'content']
        updates = []
        params = []

        for field in allowed_fields:
            if field in kwargs:
                updates.append(f"{field} = ?")
                value = kwargs[field]
                if field == 'content' and isinstance(value, (list, dict)):
                    value = json.dumps(value, ensure_ascii=False)
                params.append(value)

        if not updates:
            return False

        params.append(id)
        sql = f"UPDATE knowledge_base SET {', '.join(updates)}, updated_at = CURRENT_TIMESTAMP WHERE id = ?"
        db.execute(sql, params)
        return True

    @classmethod
    def delete(cls, id):
        """删除知识条目"""
        db.execute("DELETE FROM knowledge_base WHERE id = ?", (id,))
        return True

    @classmethod
    def clear_module(cls, module_group):
        """清空指定模块的知识"""
        db.execute("DELETE FROM knowledge_base WHERE module_group = ?", (module_group,))
        return True

    @classmethod
    def get_full_knowledge_base(cls):
        """获取完整知识库（用于API返回）"""
        modules = {}

        # 获取所有模块
        rows = db.fetchall("SELECT DISTINCT module_group, module_name FROM knowledge_base")
        for row in rows:
            module_group = row['module_group']
            module_name = row['module_name']

            modules[module_group] = {
                'module_name': module_name,
                'concepts': cls.get_concepts(module_group),
                'operations': cls.get_operations(module_group),
                'common_mistakes': cls.get_mistakes(module_group),
                'tips': cls.get_tips(module_group),
                'task_knowledge': {}
            }

            # 获取任务知识
            task_rows = db.fetchall(
                "SELECT task_id, content FROM knowledge_base WHERE module_group = ? AND content_type = 'task_knowledge'",
                (module_group,)
            )
            for tr in task_rows:
                try:
                    task_data = json.loads(tr['content'])
                    modules[module_group]['task_knowledge'][tr['task_id']] = task_data
                except:
                    pass

        return modules

    @classmethod
    def init_from_json(cls, kb_data):
        """从JSON数据初始化知识库"""
        for module_group, module_data in kb_data.items():
            module_name = module_data.get('module_name', '')

            # 清空旧数据
            cls.clear_module(module_group)

            # 添加核心概念
            concepts = module_data.get('concepts', [])
            if concepts:
                cls.add(module_group, module_name, 'concept', concepts)

            # 添加操作步骤
            operations = module_data.get('operations', [])
            for op in operations:
                cls.add(module_group, module_name, 'operation', op)

            # 添加常见错误
            mistakes = module_data.get('common_mistakes', [])
            if mistakes:
                cls.add(module_group, module_name, 'mistake', mistakes)

            # 添加技巧提示
            tips = module_data.get('tips', [])
            if tips:
                cls.add(module_group, module_name, 'tip', tips)

            # 添加任务知识
            task_knowledge = module_data.get('task_knowledge', {})
            for task_id, task_data in task_knowledge.items():
                cls.add(module_group, module_name, 'task_knowledge', task_data, task_id)

        return True
