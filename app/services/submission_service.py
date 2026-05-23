# -*- coding: utf-8 -*-
"""
提交记录服务 - SQLite 版本
"""
import os
import json
import sqlite3
from datetime import datetime
from contextlib import contextmanager

# 数据库路径
DATABASE_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'data', 'app.db')

@contextmanager
def get_db():
    """获取数据库连接"""
    conn = sqlite3.connect(DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
    finally:
        conn.close()

def row_to_dict(row):
    """将 sqlite3.Row 转换为字典"""
    if row is None:
        return None
    return dict(row)

def get_all_submissions():
    """获取所有提交记录"""
    with get_db() as conn:
        cursor = conn.execute('''
            SELECT * FROM submissions ORDER BY submitted_at DESC
        ''')
        rows = cursor.fetchall()
        return [row_to_dict(row) for row in rows]

def get_submission_by_id(submission_id):
    """根据ID获取提交记录"""
    with get_db() as conn:
        cursor = conn.execute('''
            SELECT * FROM submissions WHERE submission_id = ?
        ''', (submission_id,))
        row = cursor.fetchone()
        return row_to_dict(row)

def get_submissions_by_task(task_id):
    """根据任务ID获取提交记录"""
    with get_db() as conn:
        cursor = conn.execute('''
            SELECT * FROM submissions WHERE task_id = ? ORDER BY submitted_at DESC
        ''', (task_id,))
        rows = cursor.fetchall()
        return [row_to_dict(row) for row in rows]

def get_submissions_by_user(user_id):
    """根据用户ID获取提交记录"""
    with get_db() as conn:
        cursor = conn.execute('''
            SELECT * FROM submissions WHERE user_id = ? ORDER BY submitted_at DESC
        ''', (user_id,))
        rows = cursor.fetchall()
        return [row_to_dict(row) for row in rows]

def get_submissions_by_class(class_name):
    """根据班级获取提交记录"""
    with get_db() as conn:
        cursor = conn.execute('''
            SELECT * FROM submissions WHERE student_class = ? ORDER BY submitted_at DESC
        ''', (class_name,))
        rows = cursor.fetchall()
        return [row_to_dict(row) for row in rows]

def save_submission(submission_data):
    """保存提交记录"""
    score_result = submission_data.get('score_result', {})
    
    with get_db() as conn:
        conn.execute('''
            INSERT OR REPLACE INTO submissions 
            (submission_id, task_id, task_name, module, student_name, student_class, user_id, 
             filename, filepath, total_score, max_score, percentage, level, level_name,
             score_result, learning_tasks, study_advice, submitted_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        ''', (
            submission_data.get('submission_id', ''),
            submission_data.get('task_id', ''),
            score_result.get('task_name', ''),
            score_result.get('module', ''),
            submission_data.get('student_name', ''),
            submission_data.get('student_class', ''),
            submission_data.get('user_id', ''),
            submission_data.get('filename', ''),
            submission_data.get('filepath', ''),
            score_result.get('total_score', 0),
            score_result.get('max_score', 100),
            score_result.get('percentage', 0),
            score_result.get('level', ''),
            score_result.get('level_name', ''),
            json.dumps(score_result, ensure_ascii=False),
            json.dumps(submission_data.get('learning_tasks'), ensure_ascii=False) if submission_data.get('learning_tasks') else None,
            submission_data.get('study_advice', ''),
            submission_data.get('submitted_at', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        ))
        conn.commit()

def update_submission(submission_id, update_data):
    """更新提交记录"""
    # 构建动态更新语句
    set_clauses = []
    params = []
    
    for key, value in update_data.items():
        if key in ['score_result', 'learning_tasks']:
            set_clauses.append(f"{key} = ?")
            params.append(json.dumps(value, ensure_ascii=False) if value else None)
        elif key in ['total_score', 'max_score', 'percentage']:
            # 这些字段在 score_result 中，需要特殊处理
            continue
        else:
            set_clauses.append(f"{key} = ?")
            params.append(value)
    
    if not set_clauses:
        return
    
    # 处理 score_result 中的字段
    if 'score_result' in update_data:
        score_result = update_data['score_result']
        for field in ['total_score', 'max_score', 'percentage', 'level', 'level_name', 'task_name', 'module']:
            if field in score_result:
                set_clauses.append(f"{field} = ?")
                params.append(score_result[field])
    
    params.append(submission_id)
    
    with get_db() as conn:
        conn.execute(f'''
            UPDATE submissions SET {', '.join(set_clauses)}, updated_at = CURRENT_TIMESTAMP
            WHERE submission_id = ?
        ''', params)
        conn.commit()

def delete_submission(submission_id):
    """删除提交记录"""
    with get_db() as conn:
        conn.execute('DELETE FROM submissions WHERE submission_id = ?', (submission_id,))
        conn.commit()

def delete_all_submissions():
    """删除所有提交记录"""
    with get_db() as conn:
        conn.execute('DELETE FROM submissions')
        conn.commit()

def get_submission_count():
    """获取提交记录总数"""
    with get_db() as conn:
        cursor = conn.execute('SELECT COUNT(*) as cnt FROM submissions')
        return cursor.fetchone()['cnt']

def get_submission_full_data(submission_id):
    """获取完整的提交记录数据（包含解析后的 score_result 和 learning_tasks）"""
    submission = get_submission_by_id(submission_id)
    if submission:
        # 解析 JSON 字段
        if submission.get('score_result'):
            submission['score_result'] = json.loads(submission['score_result'])
        if submission.get('learning_tasks'):
            submission['learning_tasks'] = json.loads(submission['learning_tasks'])
    return submission

def get_all_submissions_full():
    """获取所有提交记录的完整数据"""
    submissions = get_all_submissions()
    for sub in submissions:
        if sub.get('score_result'):
            sub['score_result'] = json.loads(sub['score_result'])
        if sub.get('learning_tasks'):
            sub['learning_tasks'] = json.loads(sub['learning_tasks'])
    return submissions
