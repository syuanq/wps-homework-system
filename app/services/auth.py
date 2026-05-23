# -*- coding: utf-8 -*-
"""
用户认证模块
"""
import os
import json
import hashlib
from functools import wraps
from flask import session, redirect, url_for, jsonify

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'data')
USERS_FILE = os.path.join(DATA_DIR, 'users.json')


def _load_users():
    """加载用户数据"""
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    # 默认用户
    default_users = {
        "admin": {
            "password_hash": _hash_password("123456"),
            "name": "管理员",
            "role": "admin",
            "title": "管理员"
        },
        "teacher": {
            "password_hash": _hash_password("123456"),
            "name": "教师",
            "role": "teacher",
            "title": "教师"
        },
        "student": {
            "password_hash": _hash_password("123456"),
            "name": "学生",
            "role": "student",
            "title": "学生",
            "class_name": "计算机2401班"
        }
    }
    _save_users(default_users)
    return default_users


def _save_users(users):
    """保存用户数据"""
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(users, f, ensure_ascii=False, indent=2)


def _hash_password(password):
    """密码哈希"""
    return hashlib.sha256(password.encode('utf-8')).hexdigest()


def authenticate(username, password):
    """验证用户登录"""
    users = _load_users()
    user = users.get(username)
    if not user:
        return None
    if user['password_hash'] == _hash_password(password):
        return {
            "username": username,
            "name": user['name'],
            "role": user['role'],
            "title": user.get('title', ''),
            "class_name": user.get('class_name', '')
        }
    return None


def login_user(user_info):
    """登录用户（写入session）"""
    session['user'] = user_info
    session['logged_in'] = True


def logout_user():
    """登出用户"""
    session.clear()


def get_current_user():
    """获取当前登录用户"""
    return session.get('user')


def is_logged_in():
    """是否已登录"""
    return session.get('logged_in', False)


def is_teacher():
    """是否是教师（包括管理员）"""
    user = get_current_user()
    return user and user.get('role') in ('teacher', 'admin')


def is_student():
    """是否是学生"""
    user = get_current_user()
    return user and user.get('role') == 'student'


def teacher_required(f):
    """教师权限装饰器"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not is_teacher():
            if request.is_json or request.path.startswith('/api/'):
                return jsonify({"success": False, "message": "需要教师权限"}), 403
            return redirect(url_for('main.login', next=request.url))
        return f(*args, **kwargs)
    return decorated_function


def login_required(f):
    """登录要求装饰器"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not is_logged_in():
            if request.is_json or request.path.startswith('/api/'):
                return jsonify({"success": False, "message": "请先登录"}), 401
            return redirect(url_for('main.login', next=request.url))
        return f(*args, **kwargs)
    return decorated_function


def get_all_students():
    """获取所有学生列表"""
    users = _load_users()
    return [
        {"username": uid, **info}
        for uid, info in users.items()
        if info.get('role') == 'student'
    ]


def get_all_teachers():
    """获取所有教师列表"""
    users = _load_users()
    return [
        {"username": uid, **info}
        for uid, info in users.items()
        if info.get('role') == 'teacher'
    ]


def add_user(username, password, name, role, class_name=''):
    """添加用户"""
    users = _load_users()
    if username in users:
        return False, "用户名已存在"
    users[username] = {
        "password_hash": _hash_password(password),
        "name": name,
        "role": role,
        "title": "教师" if role == "teacher" else "学生",
        "class_name": class_name
    }
    _save_users(users)
    return True, "添加成功"


def change_password(username, old_password, new_password):
    """修改密码"""
    users = _load_users()
    user = users.get(username)
    if not user:
        return False, "用户不存在"
    if user['password_hash'] != _hash_password(old_password):
        return False, "原密码错误"
    user['password_hash'] = _hash_password(new_password)
    _save_users(users)
    return True, "密码修改成功"
