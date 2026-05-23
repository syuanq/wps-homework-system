# -*- coding: utf-8 -*-
"""
信息技术基础作业智能评价与分层学习系统
配置文件
"""
import os
import json

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data')
CONFIG_FILE = os.path.join(DATA_DIR, 'config.json')


def _load_config():
    """从JSON文件加载配置"""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def _save_config(cfg):
    """保存配置到JSON文件"""
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def get_ai_config():
    """获取AI配置"""
    cfg = _load_config()
    return {
        'api_key': cfg.get('ai_api_key', os.environ.get('DEEPSEEK_API_KEY', '')),
        'api_url': cfg.get('ai_api_url', '') or 'https://api.deepseek.com/chat/completions',
        'model': cfg.get('ai_model', '') or 'deepseek-chat',
    }


def save_ai_config(api_key, api_url, model):
    """保存AI配置"""
    cfg = _load_config()
    cfg['ai_api_key'] = api_key
    cfg['ai_api_url'] = api_url or 'https://api.deepseek.com/chat/completions'
    cfg['ai_model'] = model or 'deepseek-chat'
    _save_config(cfg)


class Config:
    """基础配置"""
    SECRET_KEY = os.environ.get('SECRET_KEY', 'it-homework-system-2026')

    # 上传文件配置
    UPLOAD_FOLDER = os.path.join(DATA_DIR, 'uploads')
    RESULT_FOLDER = os.path.join(DATA_DIR, 'results')
    GENERATED_FOLDER = os.path.join(DATA_DIR, 'generated')
    MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB
    ALLOWED_EXTENSIONS = {'docx', 'xlsx', 'pptx'}

    # 数据库配置
    DATABASE = os.path.join(DATA_DIR, 'system.db')

    # DeepSeek API 配置（从配置文件读取，在init_app中加载）
    DEEPSEEK_API_KEY = ''
    DEEPSEEK_API_URL = ''
    DEEPSEEK_MODEL = ''

    @staticmethod
    def init_app(app):
        """初始化应用配置"""
        # 从配置文件加载AI配置
        ai_cfg = get_ai_config()
        Config.DEEPSEEK_API_KEY = ai_cfg['api_key']
        Config.DEEPSEEK_API_URL = ai_cfg['api_url']
        Config.DEEPSEEK_MODEL = ai_cfg['model']
        app.config['DEEPSEEK_API_KEY'] = ai_cfg['api_key']
        app.config['DEEPSEEK_API_URL'] = ai_cfg['api_url']
        app.config['DEEPSEEK_MODEL'] = ai_cfg['model']
        
        os.makedirs(Config.UPLOAD_FOLDER, exist_ok=True)
        os.makedirs(Config.RESULT_FOLDER, exist_ok=True)
        os.makedirs(Config.GENERATED_FOLDER, exist_ok=True)
        os.makedirs(DATA_DIR, exist_ok=True)
