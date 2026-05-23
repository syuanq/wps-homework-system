# -*- coding: utf-8 -*-
"""
数据库连接管理
"""
import os
import sqlite3
from flask import g, current_app
from contextlib import contextmanager

# 数据库文件路径
DATABASE_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'data', 'app.db')


def get_db_path():
    """获取数据库文件路径"""
    return DATABASE_PATH


@contextmanager
def get_db_connection():
    """获取数据库连接的上下文管理器"""
    conn = sqlite3.connect(DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
    finally:
        conn.close()


class Database:
    """数据库操作类"""

    def __init__(self):
        self.db_path = DATABASE_PATH
        # 确保数据目录存在
        os.makedirs(os.path.dirname(DATABASE_PATH), exist_ok=True)

    def get_connection(self):
        """获取数据库连接"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    def execute(self, sql, params=()):
        """执行SQL语句"""
        with get_db_connection() as conn:
            cursor = conn.execute(sql, params)
            conn.commit()
            return cursor

    def fetchone(self, sql, params=()):
        """查询单条记录"""
        with get_db_connection() as conn:
            cursor = conn.execute(sql, params)
            return cursor.fetchone()

    def fetchall(self, sql, params=()):
        """查询多条记录"""
        with get_db_connection() as conn:
            cursor = conn.execute(sql, params)
            return cursor.fetchall()

    def init_tables(self):
        """初始化数据库表"""
        schema_sql = '''
        -- 用户表
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password_hash TEXT NOT NULL,
            name TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin', 'teacher', 'student')),
            title TEXT,
            class_name TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        -- 学生表
        CREATE TABLE IF NOT EXISTS students (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            student_no TEXT UNIQUE,
            class_name TEXT,
            gender TEXT,
            phone TEXT,
            email TEXT,
            remark TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        -- 系统配置表
        CREATE TABLE IF NOT EXISTS configs (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL,
            description TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        -- 知识库表
        CREATE TABLE IF NOT EXISTS knowledge_base (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            module_group TEXT NOT NULL,
            module_name TEXT NOT NULL,
            content_type TEXT NOT NULL CHECK(content_type IN ('concept', 'operation', 'mistake', 'tip', 'task_knowledge')),
            task_id TEXT,
            content TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        -- 素材文件表
        CREATE TABLE IF NOT EXISTS materials (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            material_id TEXT UNIQUE NOT NULL,
            task_id TEXT NOT NULL,
            filename TEXT NOT NULL,
            filepath TEXT NOT NULL,
            keywords TEXT,  -- JSON格式存储
            uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        -- 答案文件表
        CREATE TABLE IF NOT EXISTS answers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            answer_id TEXT UNIQUE NOT NULL,
            task_id TEXT NOT NULL,
            filename TEXT NOT NULL,
            filepath TEXT NOT NULL,
            uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        -- 创建索引
        CREATE INDEX IF NOT EXISTS idx_users_role ON users(role);
        CREATE INDEX IF NOT EXISTS idx_students_class ON students(class_name);
        CREATE INDEX IF NOT EXISTS idx_students_no ON students(student_no);
        CREATE INDEX IF NOT EXISTS idx_kb_module ON knowledge_base(module_group);
        CREATE INDEX IF NOT EXISTS idx_kb_task ON knowledge_base(task_id);
        CREATE INDEX IF NOT EXISTS idx_materials_task ON materials(task_id);
        CREATE INDEX IF NOT EXISTS idx_answers_task ON answers(task_id);
        '''

        with get_db_connection() as conn:
            conn.executescript(schema_sql)
            conn.commit()


# 全局数据库实例
db = Database()


def init_db(app=None):
    """初始化数据库"""
    db.init_tables()
