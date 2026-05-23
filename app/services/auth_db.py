# -*- coding: utf-8 -*-
"""
用户认证模块 - SQLite版本
"""
from functools import wraps
from flask import session, redirect, url_for, jsonify, request
from app.models.user import User


def authenticate(username, password):
    """验证用户登录"""
    return User.authenticate(username, password)


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
    users = User.get_all_students()
    return [{"username": u['username'], **u} for u in users]


def get_all_teachers():
    """获取所有教师列表"""
    users = User.get_all_teachers()
    return [{"username": u['username'], **u} for u in users]


def add_user(username, password, name, role, class_name=''):
    """添加用户"""
    return User.create(username, password, name, role, None, class_name)


def change_password(username, old_password, new_password):
    """修改密码"""
    return User.change_password(username, old_password, new_password)
