# -*- coding: utf-8 -*-
"""
WSGI入口文件 - 用于Gunicorn/Apache部署
"""
import sys
import os

# 确保项目目录在Python路径中
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 直接导入app.py模块（避免与app包冲突）
import importlib.util
spec = importlib.util.spec_from_file_location("app_module", os.path.join(os.path.dirname(__file__), "app.py"))
app_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app_module)

application = app_module.create_app()
