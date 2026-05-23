# -*- coding: utf-8 -*-
"""
页面路由
"""
from flask import Blueprint, render_template, request, redirect, url_for, session
from app.services.auth_db import get_current_user, is_logged_in, is_teacher, is_student, logout_user

main_bp = Blueprint('main', __name__)


def login_required(view_func):
    """登录检查装饰器（页面路由用，未登录重定向到登录页）"""
    from functools import wraps
    @wraps(view_func)
    def wrapper(*args, **kwargs):
        if not is_logged_in():
            return redirect(url_for('main.login', next=request.url))
        return view_func(*args, **kwargs)
    return wrapper


def teacher_required(view_func):
    """教师权限装饰器（页面路由用）"""
    from functools import wraps
    @wraps(view_func)
    def wrapper(*args, **kwargs):
        if not is_logged_in():
            return redirect(url_for('main.login', next=request.url))
        if not is_teacher():
            return redirect(url_for('main.index'))
        return view_func(*args, **kwargs)
    return wrapper


@main_bp.route('/login')
def login():
    """登录页面"""
    if is_logged_in():
        user = get_current_user()
        if user.get('role') == 'teacher':
            return redirect(url_for('main.admin'))
        return redirect(url_for('main.index'))
    return render_template('login.html')


@main_bp.route('/logout')
def logout():
    """登出"""
    logout_user()
    return redirect(url_for('main.login'))


@main_bp.route('/')
@login_required
def index():
    """首页 - 学生上传作业"""
    return render_template('index.html')


@main_bp.route('/result/<submission_id>')
@login_required
def result(submission_id):
    """评分结果页"""
    return render_template('result.html', submission_id=submission_id)


@main_bp.route('/tasks')
@login_required
def tasks():
    """分层学习任务页"""
    return render_template('tasks.html')


@main_bp.route('/my-history')
@login_required
def my_history():
    """学生作业提交记录"""
    return render_template('my_history.html')


@main_bp.route('/admin')
@teacher_required
def admin():
    """教师管理后台"""
    return render_template('admin.html')


@main_bp.route('/admin/knowledge-base')
@teacher_required
def knowledge_base():
    """知识库管理"""
    return render_template('knowledge_base.html')


@main_bp.route('/admin/materials')
@teacher_required
def materials():
    """素材文件管理"""
    return render_template('materials.html')
