# -*- coding: utf-8 -*-
"""
DeepSeek AI 服务
用于根据学生成绩生成分层学习任务
"""
import json
import requests
from .scoring_rules import ALL_TASKS, LEVEL_RULES, get_level


class DeepSeekService:
    """DeepSeek API 服务"""
    
    def __init__(self, api_key=None, api_url=None, model=None):
        self.api_key = api_key
        self.api_url = api_url or 'https://api.deepseek.com/chat/completions'
        self.model = model or 'deepseek-chat'
    
    def _call_api(self, messages, temperature=0.7, max_tokens=2000):
        """调用DeepSeek API"""
        if not self.api_key or not self.api_url:
            # 无API Key或URL时返回模拟数据
            return self._mock_response(messages)
        
        headers = {
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {self.api_key}'
        }
        payload = {
            'model': self.model,
            'messages': messages,
            'temperature': temperature,
            'max_tokens': max_tokens,
        }
        
        # 重试机制（最多3次，超时递增）
        last_error = None
        for attempt in range(3):
            try:
                timeout = 60 + attempt * 30  # 60s, 90s, 120s
                response = requests.post(self.api_url, headers=headers, json=payload, timeout=timeout)
                
                # 检查HTTP状态码
                if response.status_code != 200:
                    error_text = response.text[:200]
                    last_error = Exception(f"API返回HTTP {response.status_code}: {error_text}")
                    if attempt < 2:
                        import time
                        time.sleep(2)
                        continue
                    return None  # 返回None表示调用失败
                
                result = response.json()
                return result['choices'][0]['message']['content']
            except requests.exceptions.Timeout:
                last_error = Exception("API请求超时")
                if attempt < 2:
                    import time
                    time.sleep(2)
                    continue
            except requests.exceptions.ConnectionError:
                last_error = Exception("无法连接到AI服务")
                if attempt < 2:
                    import time
                    time.sleep(3)
                    continue
            except Exception as e:
                last_error = e
                if attempt < 2:
                    import time
                    time.sleep(2)
                    continue
        return None  # 所有重试都失败
    
    def _mock_response(self, messages):
        """无API Key时的模拟响应"""
        # 从消息中提取关键信息
        last_msg = messages[-1]["content"] if messages else ""
        
        if "分层学习任务" in last_msg or "学习任务" in last_msg:
            return json.dumps({
                "excellent": [
                    {"name": "综合排版挑战", "description": "制作一份包含所有已学排版技巧的精美文档", "difficulty": "拓展", "steps": [
                        "新建WPS文档，设置页面为A4纸，上下页边距2.5cm，左右页边距3cm",
                        "输入标题文字，设置为宋体、小三号、加粗、居中对齐",
                        "输入正文内容（至少3段），设置正文为宋体、小四号",
                        "选中正文所有段落，设置首行缩进2字符，行距固定值36磅",
                        "在第一段中间插入一张图片，设置为\"浮于文字上方\"，调整大小为高5.5cm",
                        "将图片拖动到页面右侧合适位置",
                        "选中第二段文字，添加1.5磅的段落边框，底纹颜色设为浅蓝色",
                        "插入艺术字作为副标题，设置为华文彩云字体，大小1.37×2.86cm",
                        "设置页面艺术型边框（选择任意一种样式）",
                        "保存文档，检查整体排版效果"
                    ]},
                    {"name": "创意海报设计", "description": "运用所学Word排版技能，自主设计一张A4大小的活动宣传海报", "difficulty": "拓展", "steps": [
                        "新建WPS文档，页面方向设为纵向A4",
                        "插入艺术字作为海报主标题，调整大小和颜色使其醒目",
                        "插入一张与主题相关的图片，设置为\"浮于文字上方\"",
                        "调整图片大小使其占据页面约1/3空间",
                        "添加活动时间、地点等文字信息，设置合适的字体和字号",
                        "使用文本框添加装饰性文字或说明",
                        "为整个页面添加艺术型边框",
                        "调整各元素位置，确保排版美观协调",
                        "保存文档"
                    ]},
                ],
                "good": [
                    {"name": "图文混排强化练习", "description": "制作一份图文并茂的人物介绍文档，重点练习图片环绕方式、位置调整和大小缩放", "difficulty": "巩固", "steps": [
                        "新建WPS文档，输入标题\"我最敬佩的人\"，设置为宋体、小三号、加粗、居中",
                        "输入3段人物介绍文字",
                        "选中所有正文段落，设置首行缩进2字符，行距固定值36磅",
                        "在第一段后插入一张人物照片",
                        "右键点击图片 → \"大小和位置\" → 环绕方式选\"四周型\"",
                        "拖动图片到页面右侧，调整大小为高5cm",
                        "为最后一段添加浅灰色底纹",
                        "保存文档"
                    ]},
                    {"name": "边框底纹美化练习", "description": "为文档中的重点段落添加精美的边框和底纹效果", "difficulty": "巩固", "steps": [
                        "打开已有文档或新建文档输入内容",
                        "选中需要添加边框的段落",
                        "点击\"开始\"选项卡 → \"段落\"组中的\"边框\"下拉按钮",
                        "选择\"边框和底纹\"打开设置对话框",
                        "在\"边框\"选项卡中，选择\"方框\"样式，颜色选蓝色，宽度选1.5磅",
                        "切换到\"底纹\"选项卡，填充颜色选浅黄色",
                        "点击\"确定\"应用设置",
                        "为页面添加艺术型边框：\"页面布局\" → \"页面边框\" → 选择艺术型",
                        "保存文档"
                    ]},
                ],
                "pass": [
                    {"name": "字体段落格式复习", "description": "重新完成教材任务的操作步骤，重点练习字体、字号、加粗、对齐方式等基本格式设置", "difficulty": "基础", "steps": [
                        "打开WPS文字，新建空白文档",
                        "输入标题文字，如\"练习文档\"",
                        "选中标题文字，在\"开始\"选项卡中设置字体为\"宋体\"",
                        "继续在\"字号\"下拉框中选择\"小三\"",
                        "点击\"加粗\"按钮（B图标）使标题加粗",
                        "点击\"居中对齐\"按钮使标题居中",
                        "按回车键换行，输入一段正文文字",
                        "选中正文，设置字体为\"宋体\"，字号为\"小四\"",
                        "保存文档到桌面"
                    ]},
                    {"name": "行距与缩进设置练习", "description": "对一段纯文本进行行距、段间距、首行缩进等段落格式设置", "difficulty": "基础", "steps": [
                        "新建WPS文档，输入至少3段文字",
                        "选中所有段落（Ctrl+A全选）",
                        "点击\"开始\"选项卡 → \"段落\"组右下角的小箭头",
                        "在弹出的对话框中，找到\"特殊格式\"下拉框",
                        "选择\"首行缩进\"，磅值设为\"2字符\"",
                        "找到\"行距\"下拉框，选择\"固定值\"",
                        "将\"设置值\"改为\"36磅\"",
                        "点击\"确定\"按钮",
                        "观察段落格式变化，保存文档"
                    ]},
                ],
                "fail": [
                    {"name": "文档基本操作入门", "description": "从创建新文档开始，学习输入文字、选中文字、修改字体和字号等最基本的操作", "difficulty": "入门", "steps": [
                        "双击桌面上的WPS文字图标，启动软件",
                        "点击\"新建空白文档\"",
                        "在光标闪烁处输入\"信息技术基础练习\"",
                        "用鼠标拖动选中刚才输入的文字（按住左键从第一个字拖到最后一个字）",
                        "在上方\"开始\"选项卡中找到\"字体\"下拉框，点击选择\"宋体\"",
                        "在\"字号\"下拉框中点击选择\"四号\"",
                        "观察文字的变化",
                        "按Ctrl+S保存文档，选择保存位置，输入文件名，点击\"保存\""
                    ]},
                    {"name": "简单排版跟练", "description": "按照教材步骤，从零开始制作一份简单文档，掌握最基本的文档编辑操作", "difficulty": "入门", "steps": [
                        "打开WPS文字，新建空白文档",
                        "输入标题\"自我介绍\"，按回车键换行",
                        "输入3-5句介绍自己的文字，每段结束后按回车键",
                        "选中标题\"自我介绍\"三个字",
                        "设置字体为\"黑体\"，字号为\"小二\"，点击\"加粗\"按钮",
                        "点击\"居中对齐\"按钮让标题居中",
                        "选中正文所有段落",
                        "设置字体为\"宋体\"，字号为\"小四\"",
                        "点击\"段落\"组右下角箭头，设置首行缩进2字符",
                        "保存文档"
                    ]},
                ],
            }, ensure_ascii=False)
        
        return "AI服务未配置API Key，当前为模拟响应模式。请在系统设置中配置DeepSeek API Key以获取真实的AI生成内容。"
    
    def generate_learning_tasks(self, task_id, score, scoring_details):
        """根据成绩生成分层学习任务"""
        task_rules = ALL_TASKS.get(task_id)
        if not task_rules:
            return {"error": f"未找到任务 {task_id}"}
        
        level_key, level_info = get_level(score)
        
        # 构建提示词——只生成当前等级的任务，减少AI响应时间和JSON解析失败率
        level_name_map = {
            "excellent": "优秀（拓展挑战）",
            "good": "良好（巩固提升）",
            "pass": "及格（基础补强）",
            "fail": "需努力（入门基础）"
        }
        level_desc_map = {
            "excellent": "要有拓展性和挑战性，综合运用多种高级技巧",
            "good": "要巩固已掌握的技能，适当增加难度",
            "pass": "要针对薄弱环节加强练习，步骤要详细",
            "fail": "要从最基础的操作开始，每一步都要非常详细，适合零基础学生"
        }
        
        system_prompt = f"""你是一位信息技术基础课程教师。任务：{task_rules['task_name']}（{task_rules['module']}）。
请为「{level_name_map[level_key]}」等级的学生设计1个具体的WPS操作练习任务。

要求：
1. 任务包含name（名称）、description（描述）、steps（操作步骤列表）、material（素材数据描述）
2. steps是字符串数组，每一步要具体到操作位置和操作目的
3. **material是素材数据描述**，用于生成练习素材文件。格式为对象：
   - type: "xlsx" 或 "docx" 或 "pptx"
   - sheets: 工作表数组，每个工作表包含 name（名称）、headers（表头数组）、data（行数据二维数组）
   - 如果是docx，用 paragraphs（段落文本数组）代替 sheets
   - 如果是pptx，用 slides（幻灯片数组），每个幻灯片包含 title 和 content
4. **关键要求**：material中的数据必须与steps中的操作步骤完全对应！
   - 如果步骤说"在G1输入'销售额'"，material的headers必须包含"销售额"在G列位置
   - 如果步骤说"选中'原始数据'工作表"，material的sheets必须包含名为"原始数据"的工作表
   - 如果步骤说"使用公式计算销售额"，material的data中必须有用于计算的原数据列（如单价列和数量列）
   - 先设计material数据，再根据material编写steps
5. **操作步骤设计原则**：
   - 如果需要在数据区域上方添加标题行，步骤应先说"在第1行上方插入一行"或"右键点击行号1选择插入行"，然后再说"合并A1:G1并输入标题"
   - **不要直接写出公式的完整内容**（如"=IF(H3>=100000,"是","否")"），而是描述函数的用途（如"使用IF函数判断H3单元格的值是否大于等于100000，满足条件显示'是'，否则显示'否'"），让学生自己思考公式写法
   - 同理，VLOOKUP、SUMIF、COUNTIF等函数也只描述用途和参数含义，不直接写出完整公式
   - 条件格式只描述规则（如"设置条件格式：当销售额大于10000时字体为红色"），不直接写出具体操作路径
6. 任务难度：{level_desc_map[level_key]}
7. 步骤数量：8-12步，每一步都要详细具体
8. 练习任务必须基于以下教材知识设计，操作步骤要与教材内容一致

重要：只返回JSON数组（1个元素），不要返回其他任何内容。格式示例（xlsx）：
[{{"name":"任务名","description":"描述","steps":["步骤1","步骤2"],"material":{{"type":"xlsx","sheets":[{{"name":"原始数据","headers":["编号","名称","单价","数量"],"data":[["001","商品A",10,5],["002","商品B",20,3]]}}]}}}}]"""

        # RAG：检索教材知识库
        from app.services.knowledge_base import retrieve_knowledge
        failed_skills = [d['name'] for d in scoring_details if not d['passed']]
        knowledge_text = retrieve_knowledge(task_id, level_key, failed_skills)

        user_prompt = f"学生得分：{score}/{task_rules['max_score']}（{level_info['name']}等级）\n未掌握的技能：\n"
        failed = [d for d in scoring_details if not d["passed"]]
        if failed:
            for d in failed[:8]:
                user_prompt += f"- {d['name']}\n"
        else:
            user_prompt += "- 无（全部掌握）\n"
        user_prompt += f"\n请设计1个{level_name_map[level_key]}等级的练习任务。只返回JSON数组。"
        user_prompt += f"\n\n另外，请用一句话（50字以内）给出学习建议，放在JSON数组后面，用换行分隔。"

        # 将知识库内容作为额外上下文注入
        if knowledge_text:
            messages = [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"以下是教材知识库参考内容，请基于这些知识设计练习任务：\n\n{knowledge_text}\n\n---\n\n{user_prompt}"}
            ]
        else:
            messages = [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ]
        
        response = self._call_api(messages, temperature=0.7, max_tokens=2000)
        
        # AI调用失败时直接使用兜底数据
        if response is None:
            # 生成兜底建议
            failed_names = '、'.join([d['name'] for d in scoring_details if not d['passed']][:5])
            fallback_advice = (f"同学你好！得分{score}分（{level_info['name']}等级）。"
                    f"{'建议重点练习：' + failed_names + '。' if failed_names else '继续保持！'}")
            return {
                "current_level": level_key,
                "current_level_name": level_info["name"],
                "current_level_desc": level_info["description"],
                "tasks": {level_key: self._get_fallback_tasks(level_key)},
                "used_fallback": True,
                "advice": fallback_advice,
                "error": "AI服务调用失败，已使用预设练习任务"
            }
        
        # 提取建议（JSON数组后面的文本）
        advice = ""
        if isinstance(response, str):
            # 找到JSON数组结束位置之后的内容
            json_end = response.rfind(']')
            if json_end != -1 and json_end < len(response) - 1:
                after_json = response[json_end + 1:].strip()
                if after_json:
                    advice = after_json[:200]  # 取前200字作为建议
                    # 去掉markdown代码块标记（处理 ``` 和 ```json 等情况）
                    advice = advice.strip()
                    if advice.startswith('```'):
                        first_nl = advice.find('\n')
                        if first_nl != -1:
                            advice = advice[first_nl + 1:]
                        else:
                            advice = advice[3:]
                    if advice.rstrip().endswith('```'):
                        advice = advice.rstrip()[:-3].rstrip()
                    advice = advice.strip()
        
        try:
            if isinstance(response, str):
                cleaned = self._extract_json(response)
                tasks = json.loads(cleaned)
            else:
                tasks = response
            
            if not isinstance(tasks, list):
                raise ValueError("返回的不是JSON数组")
            
            # 包装成四等级格式（兼容下游）
            all_tasks = {
                "excellent": [], "good": [], "pass": [], "fail": []
            }
            all_tasks[level_key] = tasks
            
            return {
                "current_level": level_key,
                "current_level_name": level_info["name"],
                "current_level_desc": level_info["description"],
                "tasks": all_tasks,
                "advice": advice,
            }
        except (json.JSONDecodeError, ValueError) as e:
            failed_names = '、'.join([d['name'] for d in scoring_details if not d['passed']][:5])
            fallback_advice = (f"得分{score}分（{level_info['name']}等级）。"
                    f"{'建议重点练习：' + failed_names + '。' if failed_names else '继续保持！'}")
            return {
                "current_level": level_key,
                "current_level_name": level_info["name"],
                "current_level_desc": level_info["description"],
                "tasks": {level_key: self._get_fallback_tasks(level_key)},
                "used_fallback": True,
                "advice": fallback_advice,
                "raw_response": response[:500] if isinstance(response, str) else str(response),
                "error": f"AI返回的内容无法解析为JSON格式: {str(e)}"
            }
    
    def _extract_json(self, text):
        """从AI返回的文本中提取JSON内容"""
        cleaned = text.strip()
        
        # 去掉markdown代码块标记（```json ... ``` 或 ``` ... ```）
        if cleaned.startswith("```"):
            # 去掉开头的 ```json 或 ```
            first_newline = cleaned.find("\n")
            if first_newline != -1:
                cleaned = cleaned[first_newline + 1:]
            else:
                cleaned = cleaned[3:]
            # 去掉结尾的```（可能有多个）
            while cleaned.rstrip().endswith("```"):
                cleaned = cleaned.rstrip()[:-3].rstrip()
            cleaned = cleaned.strip()
        
        # 尝试找到JSON内容（以[或{开头，以]或}结尾）
        start = -1
        for i, c in enumerate(cleaned):
            if c in '[{':
                start = i
                break
        
        if start == -1:
            return cleaned
        
        # 找到匹配的结束符
        end = -1
        open_char = cleaned[start]
        close_char = ']' if open_char == '[' else '}'
        # 从后往前找
        for i in range(len(cleaned) - 1, start, -1):
            if cleaned[i] == close_char:
                end = i
                break
        
        if end != -1:
            cleaned = cleaned[start:end + 1]
        
        return cleaned
    
    def _get_fallback_tasks(self, level_key):
        """JSON解析失败时的兜底任务数据"""
        fallback = {
            "excellent": [
                {"name": "综合排版挑战", "description": "综合运用所有排版技巧制作精美文档", "difficulty": "拓展", "steps": [
                    "新建WPS文档，设置页面为A4纸，上下页边距2.5cm，左右页边距3cm",
                    "输入标题文字，设置为宋体、小三号、加粗、居中对齐",
                    "输入正文内容（至少3段），设置正文为宋体、小四号",
                    "选中正文所有段落，设置首行缩进2字符，行距固定值36磅",
                    "在第一段中间插入一张图片，设置为\"浮于文字上方\"，调整大小为高5.5cm",
                    "将图片拖动到页面右侧合适位置",
                    "选中第二段文字，添加1.5磅的段落边框，底纹颜色设为浅蓝色",
                    "插入艺术字作为副标题，设置为华文彩云字体",
                    "设置页面艺术型边框（选择任意一种样式）",
                    "保存文档，检查整体排版效果"
                ]},
            ],
            "good": [
                {"name": "图文混排强化练习", "description": "制作一份图文并茂的人物介绍文档", "difficulty": "巩固", "steps": [
                    "新建WPS文档，输入标题\"我最敬佩的人\"，设置为宋体、小三号、加粗、居中",
                    "输入3段人物介绍文字",
                    "选中所有正文段落，设置首行缩进2字符，行距固定值36磅",
                    "在第一段后插入一张人物照片",
                    "右键点击图片 → \"大小和位置\" → 环绕方式选\"四周型\"",
                    "拖动图片到页面右侧，调整大小为高5cm",
                    "为最后一段添加浅灰色底纹",
                    "保存文档"
                ]},
            ],
            "pass": [
                {"name": "字体段落格式复习", "description": "重新完成教材任务的操作步骤", "difficulty": "基础", "steps": [
                    "打开WPS文字，新建空白文档",
                    "输入标题文字，如\"练习文档\"",
                    "选中标题文字，在\"开始\"选项卡中设置字体为\"宋体\"",
                    "继续在\"字号\"下拉框中选择\"小三\"",
                    "点击\"加粗\"按钮（B图标）使标题加粗",
                    "点击\"居中对齐\"按钮使标题居中",
                    "按回车键换行，输入一段正文文字",
                    "选中正文，设置字体为\"宋体\"，字号为\"小四\"",
                    "保存文档到桌面"
                ]},
            ],
            "fail": [
                {"name": "文档基本操作入门", "description": "从创建新文档开始，学习最基本的操作", "difficulty": "入门", "steps": [
                    "双击桌面上的WPS文字图标，启动软件",
                    "点击\"新建空白文档\"",
                    "在光标闪烁处输入\"信息技术基础练习\"",
                    "用鼠标拖动选中刚才输入的文字",
                    "在上方\"开始\"选项卡中找到\"字体\"下拉框，点击选择\"宋体\"",
                    "在\"字号\"下拉框中点击选择\"四号\"",
                    "观察文字的变化",
                    "按Ctrl+S保存文档，选择保存位置，输入文件名，点击\"保存\""
                ]},
                {"name": "简单排版跟练", "description": "按照教材步骤，从零开始制作一份简单文档", "difficulty": "入门", "steps": [
                    "打开WPS文字，新建空白文档",
                    "输入标题\"自我介绍\"，按回车键换行",
                    "输入3-5句介绍自己的文字，每段结束后按回车键",
                    "选中标题\"自我介绍\"三个字",
                    "设置字体为\"黑体\"，字号为\"小二\"，点击\"加粗\"按钮",
                    "点击\"居中对齐\"按钮让标题居中",
                    "选中正文所有段落",
                    "设置字体为\"宋体\"，字号为\"小四\"",
                    "点击\"段落\"组右下角箭头，设置首行缩进2字符",
                    "保存文档"
                ]},
            ],
        }
        return fallback.get(level_key, fallback["fail"])
    
    def generate_study_advice(self, task_id, score, scoring_details):
        """生成个性化学习建议"""
        task_rules = ALL_TASKS.get(task_id)
        if not task_rules:
            return "未找到任务信息"
        
        level_key, level_info = get_level(score)
        
        failed_items = [d for d in scoring_details if not d["passed"]]
        passed_items = [d for d in scoring_details if d["passed"]]
        
        system_prompt = f"""你是一位信息技术基础课程的教师，请根据学生的作业评分结果，给出个性化的学习建议。
要求：语气亲切鼓励，建议具体可操作，字数200-300字。"""
        
        user_prompt = f"""任务：{task_rules['task_name']}
得分：{score}/{task_rules['max_score']}（{level_info['name']}等级）

已掌握的技能：
{chr(10).join([f'  ✓ {d["name"]}' for d in passed_items]) if passed_items else '  无'}

未掌握的技能：
{chr(10).join([f'  ✗ {d["name"]}：{d["message"]}' for d in failed_items]) if failed_items else '  无'}

请给出针对性的学习建议。"""
        
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]
        
        response = self._call_api(messages, temperature=0.8, max_tokens=500)
        
        # AI调用失败时返回预设建议
        if response is None:
            failed_names = '、'.join([d['name'] for d in failed_items[:5]])
            passed_count = len(passed_items)
            total_count = len(scoring_details)
            return (f"同学你好！你本次作业得分{score}分（{level_info['name']}等级），"
                    f"共{total_count}个评分项中通过了{passed_count}项。"
                    f"{'未掌握的技能包括：' + failed_names + '。' if failed_items else ''}"
                    f"建议你重点练习未掌握的操作，参照教材步骤反复练习，完成后可重新提交作业进行评分。加油！")
        
        # 去掉markdown代码块标记（处理 ``` 和 ```json 等情况）
        if isinstance(response, str):
            response = response.strip()
            if response.startswith('```'):
                first_nl = response.find('\n')
                if first_nl != -1:
                    response = response[first_nl + 1:]
                else:
                    response = response[3:]
            if response.rstrip().endswith('```'):
                response = response.rstrip()[:-3].rstrip()
            response = response.strip()
        
        return response
