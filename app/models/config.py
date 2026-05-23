# -*- coding: utf-8 -*-
"""
系统配置模型
"""
import json
from app.models.database import db


class Config:
    """配置模型类"""

    @classmethod
    def get(cls, key, default=None):
        """获取配置项"""
        row = db.fetchone("SELECT value FROM configs WHERE key = ?", (key,))
        if row:
            return row['value']
        return default

    @classmethod
    def get_json(cls, key, default=None):
        """获取JSON格式的配置"""
        value = cls.get(key)
        if value:
            try:
                return json.loads(value)
            except json.JSONDecodeError:
                return default
        return default

    @classmethod
    def set(cls, key, value, description=None):
        """设置配置项"""
        existing = db.fetchone("SELECT 1 FROM configs WHERE key = ?", (key,))
        if existing:
            db.execute(
                "UPDATE configs SET value = ?, description = ?, updated_at = CURRENT_TIMESTAMP WHERE key = ?",
                (value, description, key)
            )
        else:
            db.execute(
                "INSERT INTO configs (key, value, description) VALUES (?, ?, ?)",
                (key, value, description)
            )
        return True

    @classmethod
    def set_json(cls, key, value, description=None):
        """设置JSON格式的配置"""
        return cls.set(key, json.dumps(value, ensure_ascii=False), description)

    @classmethod
    def delete(cls, key):
        """删除配置项"""
        db.execute("DELETE FROM configs WHERE key = ?", (key,))
        return True

    @classmethod
    def get_all(cls):
        """获取所有配置"""
        rows = db.fetchall("SELECT * FROM configs ORDER BY key")
        return [dict(row) for row in rows]

    @classmethod
    def init_default_configs(cls):
        """初始化默认配置"""
        defaults = {
            'ai_api_key': ('', 'AI API密钥'),
            'ai_api_url': ('https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions', 'AI API地址'),
            'ai_model': ('deepseek-v3.2', 'AI模型名称'),
        }

        for key, (value, desc) in defaults.items():
            if cls.get(key) is None:
                cls.set(key, value, desc)

    @classmethod
    def get_ai_config(cls):
        """获取AI配置"""
        return {
            'api_key': cls.get('ai_api_key', ''),
            'api_url': cls.get('ai_api_url', ''),
            'model': cls.get('ai_model', 'deepseek-v3.2')
        }
