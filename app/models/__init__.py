# -*- coding: utf-8 -*-
"""
数据库模型模块
"""
from app.models.database import db, init_db
from app.models.user import User
from app.models.student import Student
from app.models.config import Config
from app.models.knowledge_base import KnowledgeBase
from app.models.material import Material
from app.models.answer import Answer

__all__ = ['db', 'init_db', 'User', 'Student', 'Config', 'KnowledgeBase', 'Material', 'Answer']
