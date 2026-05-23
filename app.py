# -*- coding: utf-8 -*-
"""
基于DeepSeek的WPS作业智能批阅与分层推送决策系统
Flask 主应用
"""
import os
import json
import uuid
from datetime import datetime
from flask import Flask, render_template, request, jsonify, redirect, url_for, send_from_directory
from config.settings import Config


def create_app():
    app = Flask(__name__,
                template_folder=os.path.join(os.path.dirname(__file__), 'app', 'templates'),
                static_folder=os.path.join(os.path.dirname(__file__), 'app', 'static'))
    app.config.from_object(Config)
    app.config['SECRET_KEY'] = Config.SECRET_KEY
    Config.init_app(app)

    # 初始化数据库
    from app.models import init_db, User, Config as ConfigModel
    init_db()

    # 初始化默认数据（仅在数据库为空时）
    if not User.get_all():
        User.init_default_users()
    if not ConfigModel.get('ai_api_key'):
        ConfigModel.init_default_configs()

    # 注册路由
    from app.routers.main_routes import main_bp
    from app.routers.api_routes import api_bp
    app.register_blueprint(main_bp)
    app.register_blueprint(api_bp, url_prefix='/api')

    return app


if __name__ == '__main__':
    app = create_app()
    app.run(host='0.0.0.0', port=8080, debug=True)
