# -*- coding: utf-8 -*-
"""
评分规则配置
基于教材《信息技术基础（WPS Office）》各任务的操作步骤定义评分标准
每项评分规则严格对应教材操作步骤，步骤中没有的不评分
"""

# ============================================================
# 模块四：WPS文字
# ============================================================

# 任务4.1：制作"致大一新生的一封信"（5个评分项，满分100）
TASK_4_1_RULES = {
    "task_id": "task_4_1",
    "task_name": '4.1 制作\u201c致大一新生的一封信\u201d',
    "module": "WPS文字",
    "module_group": "wps_word",
    "file_type": "docx",
    "description": '根据教材任务4.1要求，制作\u201c致大一新生的一封信\u201d文档。',
    "max_score": 100,
    "content_keywords": {
        "required": ["大一新生", "新生", "信"],
        "any_of": ["大一", "新生", "同学", "校园", "大学", "信", "寄语", "祝福"],
        "min_match": 2,
        "description": "文档应包含与'致大一新生的一封信'相关的内容关键词"
    },
    "check_items": [
        {"id": "4_1_01", "name": "文档内容完整(标题+正文+落款)", "category": "内容", "score": 20,
         "check_type": "doc_min_paragraphs", "params": {"min_paragraphs": 3},
         "description": "步骤2：输入标题、正文和落款"},
        {"id": "4_1_02", "name": "插入日期", "category": "内容", "score": 20,
         "check_type": "doc_has_date", "params": {},
         "description": "步骤3：在文档末尾插入当前日期"},
        {"id": "4_1_03", "name": "标题居中对齐", "category": "排版", "score": 20,
         "check_type": "title_alignment", "params": {},
         "description": "步骤4：标题居中对齐"},
        {"id": "4_1_04", "name": "落款右对齐", "category": "排版", "score": 20,
         "check_type": "doc_has_right_align", "params": {},
         "description": "步骤4：落款靠右对齐"},
        {"id": "4_1_05", "name": "文档包含关键词", "category": "内容", "score": 20,
         "check_type": "doc_contains_keywords", "params": {"keywords": ["新生", "大学", "信"]},
         "description": "步骤2：文档内容包含与'致大一新生的一封信'相关的关键词"},
    ]
}

# 任务4.2：制作"感动中国人物"宣传文档（18个评分项，满分100）
TASK_4_2_RULES = {
    "task_id": "task_4_2",
    "task_name": '4.2 制作\u201c感动中国人物\u201d宣传文档',
    "module": "WPS文字",
    "module_group": "wps_word",
    "file_type": "docx",
    "description": '根据教材任务4.2要求，制作一份\u201c感动中国人物\u201d宣传文档。',
    "max_score": 100,
    "content_keywords": {
        "required": ["感动中国", "人物"],
        "any_of": ["感动中国", "人物", "感动", "年度", "颁奖", "事迹", "评选", "颁奖典礼"],
        "min_match": 2,
        "description": "文档应包含与'感动中国人物'相关的内容关键词"
    },
    "check_items": [
        # 步骤4.2.1 字体和段落格式
        {"id": "4_2_01", "name": "标题宋体小三加粗", "category": "字符格式", "score": 5,
         "check_type": "title_font_name", "params": {"font_name": "宋体"},
         "description": "步骤2-1：标题字体设置为宋体"},
        {"id": "4_2_02", "name": "标题加粗", "category": "字符格式", "score": 5,
         "check_type": "title_bold", "params": {},
         "description": "步骤2-1：标题加粗"},
        {"id": "4_2_03", "name": "标题字号小三(15磅)", "category": "字符格式", "score": 5,
         "check_type": "title_font_size_exact", "params": {"exact_size_pt": 15},
         "description": "步骤2-1：标题字号设置为小三（15磅）"},
        {"id": "4_2_04", "name": "标题艺术字效果", "category": "字符格式", "score": 5,
         "check_type": "doc_has_wordart_or_text_effect", "params": {},
         "description": "步骤2-1：标题添加艺术字效果（填充-橙色，着色4，软边缘）"},
        {"id": "4_2_05", "name": "正文字体宋体", "category": "字符格式", "score": 5,
         "check_type": "doc_body_font_name", "params": {"font_name": "宋体"},
         "description": "步骤2-3：其余正文设置为宋体"},
        {"id": "4_2_06", "name": "正文字号16磅", "category": "字符格式", "score": 5,
         "check_type": "body_font_size_exact", "params": {"exact_size_pt": 16},
         "description": "步骤2-3：其余正文设置为16磅"},
        {"id": "4_2_07", "name": "标题和部分正文居中", "category": "段落格式", "score": 10,
         "check_type": "title_alignment", "params": {"alignment": "center"},
         "description": "步骤3-1：标题和部分正文居中对齐"},
        {"id": "4_2_08", "name": "正文首行缩进2字符", "category": "段落格式", "score": 5,
         "check_type": "body_first_indent", "params": {"chars": 2},
         "description": "步骤3-2：正文首行缩进2字符"},
        {"id": "4_2_09", "name": "全文行距固定值36磅", "category": "段落格式", "score": 5,
         "check_type": "body_line_spacing_exact", "params": {"exact_spacing": 36},
         "description": "步骤3-3：全文行距设置为固定值36磅"},
        # 步骤4.2.2 边框和底纹
        {"id": "4_2_10", "name": "页面边框(方框/1.5磅)", "category": "边框底纹", "score": 5,
         "check_type": "doc_page_border_exact", "params": {"border_sz_pt": 1.5},
         "description": "步骤4.2.2-2：添加页面边框（方框、1.5磅）"},
        {"id": "4_2_11", "name": "文字和段落底纹设置", "category": "边框底纹", "score": 10,
         "check_type": "paragraph_border_or_shading", "params": {},
         "description": "步骤4.2.2-3：设置文字和段落底纹"},
        # 步骤4.2.3 图片和艺术字
        {"id": "4_2_13", "name": "插入图片", "category": "图文混排", "score": 5,
         "check_type": "has_image", "params": {},
         "description": "步骤4.2.3-1：在文档中插入人物图片"},
        {"id": "4_2_14", "name": "图片高度5.5厘米", "category": "图文混排", "score": 5,
         "check_type": "doc_has_picture_with_size", "params": {"height_cm": 5.5},
         "description": "步骤4.2.3-1-2：图片高度设置为5.5厘米"},
        {"id": "4_2_16", "name": "插入艺术字(\u201c颁奖词\u201d)", "category": "图文混排", "score": 5,
         "check_type": "has_wordart", "params": {},
         "description": "步骤4.2.3-2：插入\u201c颁奖词\u201d艺术字"},
        {"id": "4_2_17", "name": "艺术字大小设置", "category": "图文混排", "score": 5,
         "check_type": "doc_has_wordart_with_size", "params": {"height_cm": 1.37, "width_cm": 2.86},
         "description": "步骤4.2.3-2-5：修改艺术字大小（高度1.37厘米、宽度2.86厘米）"},
        {"id": "4_2_18", "name": "艺术字位置设置", "category": "图文混排", "score": 5,
         "check_type": "doc_has_wordart_with_position", "params": {},
         "description": "步骤4.2.3-2-6：艺术字绝对定位"},
        {"id": "4_2_19", "name": "背景图片浮于文字上方", "category": "图文混排", "score": 5,
         "check_type": "check_image_position", "params": {"position_type": "non_inline"},
         "description": "步骤4.2.3-3-2：背景图片浮于文字上方"},
        {"id": "4_2_20", "name": "包含多张图片", "category": "图文混排", "score": 5,
         "check_type": "doc_has_multiple_images", "params": {"min_count": 2},
         "description": "步骤4.2.3-1/3：插入人物图片和背景图片"},
    ]
}

# 任务4.3："大学生职业生涯规划"页面排版（16个评分项，满分100）
TASK_4_3_RULES = {
    "task_id": "task_4_3",
    "task_name": "4.3 \u201c大学生职业生涯规划\u201d页面排版",
    "module": "WPS文字",
    "module_group": "wps_word",
    "file_type": "docx",
    "description": '根据教材任务4.3要求，对\u201c大学生职业生涯规划\u201d文档进行页面排版。',
    "max_score": 100,
    "content_keywords": {
        "required": ["职业", "规划"],
        "any_of": ["职业", "规划", "生涯", "就业", "发展", "目标", "大学生", "人生", "未来", "梦想"],
        "min_match": 2,
        "description": "文档应包含与'大学生职业生涯规划'相关的内容关键词"
    },
    "check_items": [
        # 步骤4.3.1 页面设置
        {"id": "4_3_01", "name": "插入分页符(至少2个)", "category": "页面设置", "score": 5,
         "check_type": "doc_has_page_breaks", "params": {"min_breaks": 2},
         "description": "步骤4.3.1-1：在各章节前插入分页符"},
        {"id": "4_3_02", "name": "标题字体为宋体", "category": "字符格式", "score": 4,
         "check_type": "title_font", "params": {"font_name": "宋体"},
         "description": "步骤4.3.1-2：标题字体设置为宋体"},
        {"id": "4_3_03", "name": "标题字号45磅", "category": "字符格式", "score": 4,
         "check_type": "title_font_size_exact", "params": {"exact_size_pt": 45},
         "description": "步骤4.3.1-2：标题字号设置为45磅"},
        {"id": "4_3_04", "name": "标题加粗", "category": "字符格式", "score": 3,
         "check_type": "title_bold", "params": {},
         "description": "步骤4.3.1-2：标题加粗"},
        {"id": "4_3_05", "name": "标题水平居中", "category": "字符格式", "score": 3,
         "check_type": "title_alignment", "params": {"alignment": "center"},
         "description": "步骤4.3.1-2：标题水平居中"},
        {"id": "4_3_06", "name": "副标题行距50磅", "category": "段落格式", "score": 5,
         "check_type": "doc_paragraph_line_spacing", "params": {"spacing_pt": 50},
         "description": "步骤4.3.1-2：副标题行距设置为50磅"},
        {"id": "4_3_07", "name": "正文字体为宋体", "category": "字符格式", "score": 5,
         "check_type": "doc_body_font_name", "params": {"font_name": "宋体"},
         "description": "步骤4.3.1-2：正文字体设置为宋体"},
        {"id": "4_3_08", "name": "正文字号三号(16磅)", "category": "字符格式", "score": 5,
         "check_type": "body_font_size_exact", "params": {"exact_size_pt": 16},
         "description": "步骤4.3.1-2：正文字号设置为三号（16磅）"},
        {"id": "4_3_09", "name": "正文首行缩进2字符", "category": "段落格式", "score": 5,
         "check_type": "body_first_indent", "params": {"chars": 2},
         "description": "步骤4.3.1-2：正文首行缩进2字符"},
        {"id": "4_3_10", "name": "左右页边距2.5厘米", "category": "页面设置", "score": 5,
         "check_type": "doc_page_margins", "params": {"left_cm": 2.5, "right_cm": 2.5},
         "description": "步骤4.3.1-3：左右页边距设置为2.5厘米"},
        {"id": "4_3_11", "name": "奇偶页不同", "category": "页面设置", "score": 5,
         "check_type": "doc_has_even_odd_header", "params": {},
         "description": "步骤4.3.1-3：勾选奇偶页不同"},
        {"id": "4_3_12", "name": "首页不同", "category": "页面设置", "score": 5,
         "check_type": "doc_has_different_first_page", "params": {},
         "description": "步骤4.3.1-3：勾选首页不同"},
        {"id": "4_3_13", "name": "添加页眉", "category": "页眉页脚", "score": 5,
         "check_type": "doc_has_header", "params": {},
         "description": "步骤4.3.1-4：插入页眉内容"},
        {"id": "4_3_14", "name": "插入页码", "category": "页眉页脚", "score": 6,
         "check_type": "doc_has_page_number", "params": {},
         "description": "步骤4.3.1-4：奇偶页均插入页码"},
        # 步骤4.3.2 添加目录
        {"id": "4_3_15", "name": "设置标题1样式", "category": "目录", "score": 5,
         "check_type": "doc_has_heading_style", "params": {},
         "description": "步骤4.3.2-2：为章节标题应用标题1样式并居中"},
        {"id": "4_3_16", "name": "标题1居中", "category": "目录", "score": 4,
         "check_type": "doc_heading1_centered", "params": {},
         "description": "步骤4.3.2-2：标题1居中对齐"},
        {"id": "4_3_17", "name": "设置二级标题", "category": "目录", "score": 6,
         "check_type": "doc_has_heading2_style", "params": {},
         "description": "步骤4.3.2-3：设置二级标题"},
        {"id": "4_3_18", "name": "生成目录", "category": "目录", "score": 6,
         "check_type": "doc_has_toc", "params": {},
         "description": "步骤4.3.2-4：使用自动目录功能生成文档目录"},
        {"id": "4_3_19", "name": "目录标题居中", "category": "目录", "score": 4,
         "check_type": "doc_toc_centered", "params": {},
         "description": "步骤4.3.2-6：目录标题居中"},
        {"id": "4_3_20", "name": "文档结构完整(至少8段)", "category": "内容完整性", "score": 10,
         "check_type": "doc_min_paragraphs", "params": {"min_paragraphs": 8},
         "description": "文档包含前言、自我分析、职业分析等完整内容"},
    ]
}

# 任务4.4：个人简历制作（12个评分项，满分100）
TASK_4_4_RULES = {
    "task_id": "task_4_4",
    "task_name": "4.4 个人简历制作",
    "module": "WPS文字",
    "module_group": "wps_word",
    "file_type": "docx",
    "description": "根据教材任务4.4要求，制作个人求职简历。",
    "max_score": 100,
    "content_keywords": {
        "required": ["简历"],
        "any_of": ["简历", "个人", "求职", "教育", "经历", "技能", "姓名", "联系方式", "自我评价", "学历"],
        "min_match": 2,
        "description": "文档应包含与'个人简历'相关的内容关键词"
    },
    "check_items": [
        # 步骤4.4.1 制作个人求职简历
        {"id": "4_4_01", "name": "标题宋体小二加粗居中", "category": "字符格式", "score": 7,
         "check_type": "title_bold_center", "params": {},
         "description": "步骤4.4.1-3：标题宋体小二加粗居中"},
        {"id": "4_4_02", "name": "标题字号小二(18磅)", "category": "字符格式", "score": 7,
         "check_type": "title_font_size_exact", "params": {"exact_size_pt": 18},
         "description": "步骤4.4.1-3：标题字号设置为小二（18磅）"},
        {"id": "4_4_03", "name": "标题字符间距加宽0.1厘米", "category": "字符格式", "score": 7,
         "check_type": "doc_char_spacing", "params": {"spacing_cm": 0.1},
         "description": "步骤4.4.1-3：标题字符间距加宽0.1厘米"},
        {"id": "4_4_04", "name": "上下页边距2.54厘米", "category": "页面设置", "score": 7,
         "check_type": "doc_page_margins_tb", "params": {"top_cm": 2.54, "bottom_cm": 2.54},
         "description": "步骤4.4.1-2：上下页边距设置为2.54厘米"},
        {"id": "4_4_05", "name": "插入表格(至少2列10行)", "category": "表格操作", "score": 7,
         "check_type": "doc_has_table", "params": {"min_rows": 10, "min_cols": 2},
         "description": "步骤4.4.1-4：插入表格20行2列"},
        {"id": "4_4_06", "name": "合并单元格", "category": "表格操作", "score": 8,
         "check_type": "doc_has_merged_cells", "params": {},
         "description": "步骤4.4.1-7：对表格中的单元格进行合并操作"},
        # 步骤4.4.2 格式设置
        {"id": "4_4_07", "name": "表格水平居中", "category": "表格美化", "score": 7,
         "check_type": "doc_table_alignment", "params": {},
         "description": "步骤4.4.2-1：表格水平居中"},
        {"id": "4_4_08", "name": "表格底纹设置", "category": "表格美化", "score": 7,
         "check_type": "doc_table_has_shading", "params": {},
         "description": "步骤4.4.2-2：第1列设置底纹"},
        {"id": "4_4_09", "name": "表格边框(深蓝1.5磅)", "category": "表格美化", "score": 8,
         "check_type": "doc_table_border_exact", "params": {"border_sz_pt": 1.5, "border_color": "000080"},
         "description": "步骤4.4.2-3：设置表格边框（深蓝1.5磅）"},
        # 步骤4.4.3 数据统计
        {"id": "4_4_10", "name": "表格公式计算(合计/利润列)", "category": "表格操作", "score": 15,
         "check_type": "doc_has_table_formula", "params": {},
         "description": "步骤4.4.3-2/3：使用公式计算合计列和利润列"},
        {"id": "4_4_12", "name": "表格文字宋体11号", "category": "表格美化", "score": 7,
         "check_type": "doc_body_font_name", "params": {"font_name": "宋体"},
         "description": "步骤4.4.1-6：表格文字设置为宋体11号"},
        {"id": "4_4_13", "name": "简历内容完整(至少8段)", "category": "内容完整性", "score": 13,
         "check_type": "doc_min_paragraphs", "params": {"min_paragraphs": 8},
         "description": "简历包含基本信息、教育经历、技能特长等内容"},
    ]
}

# ============================================================
# 模块五：WPS表格
# ============================================================

# 任务5.1：制作二手房房源信息汇总表（10个评分项，满分100）
TASK_5_1_RULES = {
    "task_id": "task_5_1",
    "task_name": "5.1 制作二手房房源信息汇总表",
    "module": "WPS表格",
    "module_group": "wps_excel",
    "file_type": "xlsx",
    "description": "根据教材任务5.1要求，创建二手房房源信息汇总表。",
    "max_score": 100,
    "content_keywords": {
        "required": [],
        "any_of": ["二手房", "房源", "房价", "面积", "单价", "总价", "户型", "楼层", "小区", "朝向", "城区", "格局", "建造时间"],
        "min_match": 3,
        "description": "表格应包含与二手房数据管理相关的数据列"
    },
    "check_items": [
        {"id": "5_1_01", "name": "工作表命名(房源信息汇总表)", "category": "基本操作", "score": 10,
         "check_type": "excel_sheet_named", "params": {"sheet_names": ["房源信息汇总表"]},
         "description": "步骤2：重命名工作表为'房源信息汇总表'"},
        {"id": "5_1_02", "name": "输入表标题", "category": "数据输入", "score": 10,
         "check_type": "excel_has_keywords_in_sheet", "params": {"keywords": ["二手房房源信息汇总表"]},
         "description": "步骤3：输入表标题'二手房房源信息汇总表'"},
        {"id": "5_1_03", "name": "输入列标题(11列)", "category": "数据输入", "score": 10,
         "check_type": "excel_min_columns", "params": {"min_cols": 11},
         "description": "步骤4：输入11个列标题"},
        {"id": "5_1_04", "name": "数据行至少25行", "category": "数据输入", "score": 10,
         "check_type": "excel_min_data_rows", "params": {"min_rows": 25},
         "description": "步骤5~7：输入和复制房源数据"},
        {"id": "5_1_05", "name": "序号列填充(1~25)", "category": "数据输入", "score": 10,
         "check_type": "excel_has_keywords_in_sheet", "params": {"keywords": ["序号"]},
         "description": "步骤6：使用填充方式填充序号"},
        {"id": "5_1_06", "name": "A1:K1合并后居中", "category": "格式设置", "score": 10,
         "check_type": "excel_has_merged_cells", "params": {},
         "description": "步骤8：A1:K1合并后居中"},
        {"id": "5_1_07", "name": "第2行合并居中(记录人/审核人)", "category": "格式设置", "score": 10,
         "check_type": "excel_has_keywords_in_sheet", "params": {"keywords": ["记录人", "审核人"]},
         "description": "步骤9：第2行输入'记录人：审核人：'"},
        {"id": "5_1_08", "name": "数据区居中对齐", "category": "格式设置", "score": 15,
         "check_type": "excel_data_centered", "params": {},
         "description": "步骤10：A3:K28水平居中、垂直居中对齐"},
        {"id": "5_1_10", "name": "行高设置(第3行22/第4~28行18)", "category": "格式设置", "score": 15,
         "check_type": "excel_row_height_set", "params": {},
         "description": "步骤11：第3行行高22，第4~28行行高18"},
    ]
}

# 任务5.2：美化二手房房源信息汇总表（12个评分项，满分100）
# 任务5.2.1：字体美化与边框设置（7个评分项，满分100）
TASK_5_2_1_RULES = {
    "task_id": "task_5_2_1",
    "task_name": "5.2.1 字体美化与边框设置",
    "module": "WPS表格",
    "module_group": "wps_excel",
    "file_type": "xlsx",
    "description": "根据教材任务5.2.1要求，设置标题、表头、数据区的字体格式和边框样式。",
    "max_score": 100,
    "content_keywords": {
        "required": [],
        "any_of": ["二手房", "房源", "房价", "面积", "单价", "总价", "户型", "楼层", "小区", "朝向", "城区", "格局", "建造时间"],
        "min_match": 3,
        "description": "表格应包含与二手房数据管理相关的数据列"
    },
    "check_items": [
        {"id": "5_2_1_01", "name": "A1标题黑体18号加粗深蓝", "category": "字体格式", "score": 14,
         "check_type": "excel_cell_font_check", "params": {"row": 1, "col": 1, "font_name": "黑体", "font_size": 18, "bold": True},
         "description": "步骤1：A1黑体18号加粗深蓝色"},
        {"id": "5_2_1_02", "name": "A2:G2仿宋13号加粗", "category": "字体格式", "score": 14,
         "check_type": "excel_cell_font_check", "params": {"row": 2, "col": 1, "font_name": "仿宋", "font_size": 13, "bold": True},
         "description": "步骤2：A2:G2仿宋13号加粗"},
        {"id": "5_2_1_03", "name": "A3:K28楷体12号蓝色", "category": "字体格式", "score": 14,
         "check_type": "excel_cell_font_check", "params": {"row": 3, "col": 1, "font_name": "楷体", "font_size": 12},
         "description": "步骤3：A3:K28楷体12号蓝色"},
        {"id": "5_2_1_04", "name": "外边框深蓝粗线+内边框蓝色细线", "category": "边框设置", "score": 15,
         "check_type": "excel_has_border", "params": {},
         "description": "步骤4：设置外边框深蓝粗线、内边框蓝色细线"},
        {"id": "5_2_1_05", "name": "第2行下框线双线红色", "category": "边框设置", "score": 14,
         "check_type": "excel_has_border", "params": {},
         "description": "步骤5：第2行下框线双线红色"},
        {"id": "5_2_1_06", "name": "D4:D28水平左对齐", "category": "对齐设置", "score": 15,
         "check_type": "excel_cell_alignment_check", "params": {"row_start": 4, "row_end": 28, "col": 4, "align": "left"},
         "description": "步骤6：D4:D28水平左对齐"},
        {"id": "5_2_1_07", "name": "K列日期格式yyyy-mm-dd", "category": "数值格式", "score": 14,
         "check_type": "excel_cell_number_format", "params": {"col_start": 11, "col_end": 11, "row_start": 3, "format_type": "date"},
         "description": "步骤7：K列日期格式为yyyy-mm-dd"},
    ]
}

# 任务5.2.2：数据有效性（3个评分项，满分100）
TASK_5_2_2_RULES = {
    "task_id": "task_5_2_2",
    "task_name": "5.2.2 数据有效性",
    "module": "WPS表格",
    "module_group": "wps_excel",
    "file_type": "xlsx",
    "description": "根据教材任务5.2.2要求，为单元格设置数据有效性验证规则。",
    "max_score": 100,
    "content_keywords": {
        "required": ["员工"],
        "any_of": ["员工", "工号", "入职", "姓名", "性别", "部门", "岗位", "身份证", "学历", "联系电话", "工资"],
        "min_match": 2,
        "description": "表格应包含与'员工信息表'相关的数据列"
    },
    "check_items": [
        {"id": "5_2_2_01", "name": "身份证18位文本长度验证", "category": "数据有效性", "score": 34,
         "check_type": "excel_has_validation", "params": {"col": 8, "validation_type": "textLength", "formula1": "18"},
         "description": "步骤1：H列身份证号码18位文本长度验证"},
        {"id": "5_2_2_02", "name": "工资0-30000整数验证", "category": "数据有效性", "score": 33,
         "check_type": "excel_has_validation", "params": {"col": 11, "validation_type": "whole", "min_val": 0, "max_val": 30000},
         "description": "步骤2：K列工资0-30000整数验证"},
        {"id": "5_2_2_03", "name": "性别/状态序列验证", "category": "数据有效性", "score": 33,
         "check_type": "excel_has_validation", "params": {"col": 5, "validation_type": "list"},
         "description": "步骤3：性别和当前状态序列验证"},
    ]
}

# 任务5.2.3：条件格式（5个评分项，满分100）
TASK_5_2_3_RULES = {
    "task_id": "task_5_2_3",
    "task_name": "5.2.3 条件格式",
    "module": "WPS表格",
    "module_group": "wps_excel",
    "file_type": "xlsx",
    "description": "根据教材任务5.2.3要求，使用条件格式突出显示指定数据。",
    "max_score": 100,
    "content_keywords": {
        "required": [],
        "any_of": ["二手房", "房源", "房价", "面积", "单价", "总价", "户型", "楼层", "小区", "朝向", "城区", "格局", "建造时间"],
        "min_match": 3,
        "description": "表格应包含与二手房数据管理相关的数据列"
    },
    "check_items": [
        {"id": "5_2_3_01", "name": "单价列突出显示规则(>10000)", "category": "条件格式", "score": 20,
         "check_type": "excel_has_conditional_format", "params": {},
         "description": "步骤1：单价>10000设为浅红填充深红文本"},
        {"id": "5_2_3_02", "name": "单价列突出显示规则(<5000)", "category": "条件格式", "score": 20,
         "check_type": "excel_has_conditional_format", "params": {},
         "description": "步骤1：单价<5000设为巧克力黄着色6"},
        {"id": "5_2_3_03", "name": "面积列红色数据条", "category": "条件格式", "score": 20,
         "check_type": "excel_has_data_bars", "params": {},
         "description": "步骤2：面积列设置红色数据条"},
        {"id": "5_2_3_04", "name": "总价列高于/低于平均值规则", "category": "条件格式", "score": 20,
         "check_type": "excel_has_conditional_format", "params": {},
         "description": "步骤3：总价高于平均值浅红填充，低于平均值绿填充"},
        {"id": "5_2_3_05", "name": "小区名称包含'东江湾'规则", "category": "条件格式", "score": 20,
         "check_type": "excel_has_conditional_format", "params": {},
         "description": "步骤4：小区名称包含'东江湾'设为浅红色填充"},
    ]
}

# 任务5.2.4：套用表格样式（3个评分项，满分100）
TASK_5_2_4_RULES = {
    "task_id": "task_5_2_4",
    "task_name": "5.2.4 套用表格样式",
    "module": "WPS表格",
    "module_group": "wps_excel",
    "file_type": "xlsx",
    "description": "根据教材任务5.2.4要求，套用表格样式并设置数值格式。",
    "max_score": 100,
    "content_keywords": {
        "required": [],
        "any_of": ["二手房", "房源", "房价", "面积", "单价", "总价", "户型", "楼层", "小区", "朝向", "城区", "格局", "建造时间"],
        "min_match": 3,
        "description": "表格应包含与二手房数据管理相关的数据列"
    },
    "check_items": [
        {"id": "5_2_4_01", "name": "套用表格样式", "category": "表格样式", "score": 34,
         "check_type": "excel_has_table_style", "params": {},
         "description": "步骤1：A4:K28应用表格样式"},
        {"id": "5_2_4_02", "name": "C列数值保留一位小数", "category": "数值格式", "score": 33,
         "check_type": "excel_cell_number_format", "params": {"col_start": 3, "col_end": 3, "row_start": 4, "format_type": "decimal"},
         "description": "步骤2：C4:C28数值保留一位小数"},
        {"id": "5_2_4_03", "name": "C列设置千位分隔符", "category": "数值格式", "score": 33,
         "check_type": "excel_cell_number_format", "params": {"col_start": 3, "col_end": 3, "row_start": 4, "format_type": "thousands"},
         "description": "步骤2：C4:C28设置千位分隔符"},
    ]
}

# 任务5.3.1：计算总价与数值格式（2个评分项，满分100）
TASK_5_3_1_RULES = {
    "task_id": "task_5_3_1",
    "task_name": "5.3.1 计算总价与数值格式",
    "module": "WPS表格",
    "module_group": "wps_excel",
    "file_type": "xlsx",
    "description": "根据教材任务5.3.1要求，使用公式计算总价列并设置数值格式。",
    "max_score": 100,
    "content_keywords": {
        "required": [],
        "any_of": ["二手房", "房源", "房价", "面积", "单价", "总价", "户型", "楼层", "小区", "朝向", "城区", "格局", "建造时间"],
        "min_match": 3,
        "description": "表格应包含与二手房数据管理相关的数据列"
    },
    "check_items": [
        {"id": "5_3_1_01", "name": "使用公式计算总价列", "category": "公式计算", "score": 50,
         "check_type": "excel_has_formula", "params": {"col": 9},
         "description": "步骤1：使用公式计算总价（=单价*面积/10000）"},
        {"id": "5_3_1_02", "name": "总价列数值保留一位小数", "category": "数值格式", "score": 50,
         "check_type": "excel_cell_number_format", "params": {"col_start": 9, "col_end": 9, "row_start": 4, "format_type": "decimal"},
         "description": "步骤1：总价列数值格式保留一位小数"},
    ]
}

# 任务5.3.2：常用函数统计（10个评分项，满分100）
TASK_5_3_2_RULES = {
    "task_id": "task_5_3_2",
    "task_name": "5.3.2 常用函数统计",
    "module": "WPS表格",
    "module_group": "wps_excel",
    "file_type": "xlsx",
    "description": "根据教材任务5.3.2要求，使用COUNTIF、AVERAGEIF、SUMIF、IF、RANK等函数统计二手房数据。",
    "max_score": 100,
    "content_keywords": {
        "required": [],
        "any_of": ["二手房", "房源", "房价", "面积", "单价", "总价", "户型", "楼层", "小区", "朝向", "城区", "格局", "建造时间"],
        "min_match": 3,
        "description": "表格应包含与二手房数据管理相关的数据列"
    },
    "check_items": [
        {"id": "5_3_2_01", "name": "COUNTIF函数", "category": "统计函数", "score": 10,
         "check_type": "excel_has_formula", "params": {"formula_keyword": "COUNTIF"},
         "description": "步骤5.3.2-1：统计各城区房源数量"},
        {"id": "5_3_2_02", "name": "AVERAGEIF函数", "category": "统计函数", "score": 10,
         "check_type": "excel_has_formula", "params": {"formula_keyword": "AVERAGEIF"},
         "description": "步骤5.3.2-2：计算各城区单价均值"},
        {"id": "5_3_2_03", "name": "SUMIF函数", "category": "统计函数", "score": 10,
         "check_type": "excel_has_formula", "params": {"formula_keyword": "SUMIF"},
         "description": "步骤5.3.2-3：计算各城区总价和"},
        {"id": "5_3_2_04", "name": "IF函数(面积分类)", "category": "逻辑函数", "score": 10,
         "check_type": "excel_has_formula", "params": {"formula_keyword": "IF("},
         "description": "步骤5.3.2-4：面积分类（小/中/大/超大户型）"},
        {"id": "5_3_2_05", "name": "IF+AND组合", "category": "逻辑函数", "score": 10,
         "check_type": "excel_has_formula", "params": {"formula_keyword": "AND"},
         "description": "步骤5.3.2-5：多条件判断"},
        {"id": "5_3_2_06", "name": "RANK.EQ函数", "category": "排名函数", "score": 10,
         "check_type": "excel_has_formula", "params": {"formula_keyword": "RANK"},
         "description": "步骤5.3.2-6：按单价排名"},
        {"id": "5_3_2_07", "name": "MIN/MAX函数", "category": "统计函数", "score": 10,
         "check_type": "excel_has_formula", "params": {"formula_keyword": ["MIN", "MAX"]},
         "description": "步骤5.3.2-7：统计单价最小值和面积最大值"},
        {"id": "5_3_2_08", "name": "SUBTOTAL函数", "category": "统计函数", "score": 10,
         "check_type": "excel_has_formula", "params": {"formula_keyword": "SUBTOTAL"},
         "description": "步骤5.3.2-8：SUBTOTAL统计"},
        {"id": "5_3_2_09", "name": "VLOOKUP函数", "category": "查找函数", "score": 10,
         "check_type": "excel_has_formula", "params": {"formula_keyword": "VLOOKUP"},
         "description": "步骤5.3.2-9：匹配销售经理和联系方式"},
        {"id": "5_3_2_10", "name": "AVERAGEIF千位分隔符", "category": "格式设置", "score": 10,
         "check_type": "excel_has_number_format_custom", "params": {},
         "description": "步骤5.3.2-2：AVERAGEIF结果设置千位分隔符"},
    ]
}

# 任务5.4.1：排序与分类汇总（5个评分项，满分100）
TASK_5_4_1_RULES = {
    "task_id": "task_5_4_1",
    "task_name": "5.4.1 排序与分类汇总",
    "module": "WPS表格",
    "module_group": "wps_excel",
    "file_type": "xlsx",
    "description": "根据教材任务5.4.1要求，对二手房数据进行多关键字排序和分类汇总。",
    "max_score": 100,
    "content_keywords": {
        "required": [],
        "any_of": ["二手房", "房源", "房价", "面积", "单价", "总价", "户型", "楼层", "小区", "朝向", "城区", "格局", "建造时间"],
        "min_match": 3,
        "description": "表格应包含与二手房数据管理相关的数据列"
    },
    "check_items": [
        {"id": "5_4_1_01", "name": "排序表1-城区升序+建造时间降序", "category": "排序", "score": 25,
         "check_type": "excel_data_sorted", "params": {"sheet_name": "排序表1", "sort_col": 4},
         "description": "步骤1：排序表1按城区升序、建造时间降序排列"},
        {"id": "5_4_1_02", "name": "排序表2-建造时间降序+单价升序", "category": "排序", "score": 25,
         "check_type": "excel_data_sorted", "params": {"sheet_name": "排序表2"},
         "description": "步骤2：排序表2按建造时间降序、单价升序排列"},
        {"id": "5_4_1_03", "name": "分类汇总表1-按城区汇总单价平均值", "category": "分类汇总", "score": 25,
         "check_type": "excel_has_subtotals", "params": {"sheet_name": "分类汇总表1"},
         "description": "步骤3：分类汇总表1按城区统计单价平均值"},
        {"id": "5_4_1_04", "name": "分类汇总表2-按朝向汇总单价/总价/面积", "category": "分类汇总", "score": 25,
         "check_type": "excel_has_subtotals", "params": {"sheet_name": "分类汇总表2"},
         "description": "步骤4：分类汇总表2按朝向统计单价、总价、面积平均值"},
    ]
}

# 任务5.4.2：数据筛选（5个评分项，满分100）
TASK_5_4_2_RULES = {
    "task_id": "task_5_4_2",
    "task_name": "5.4.2 数据筛选",
    "module": "WPS表格",
    "module_group": "wps_excel",
    "file_type": "xlsx",
    "description": "根据教材任务5.4.2要求，使用自动筛选和高级筛选分析二手房数据。",
    "max_score": 100,
    "content_keywords": {
        "required": [],
        "any_of": ["二手房", "房源", "房价", "面积", "单价", "总价", "户型", "楼层", "小区", "朝向", "城区", "格局", "建造时间"],
        "min_match": 3,
        "description": "表格应包含与二手房数据管理相关的数据列"
    },
    "check_items": [
        {"id": "5_4_2_01", "name": "自动筛选1-筛选建造时间", "category": "筛选", "score": 25,
         "check_type": "excel_has_autofilter", "params": {"sheet_name": "自动筛选1"},
         "description": "步骤1：自动筛选建造时间为2020年和2021年"},
        {"id": "5_4_2_02", "name": "自动筛选2-筛选总价低于平均值", "category": "筛选", "score": 25,
         "check_type": "excel_has_autofilter", "params": {"sheet_name": "自动筛选2"},
         "description": "步骤2：筛选总价低于平均值的房源"},
        {"id": "5_4_2_03", "name": "高级筛选1-AND条件(朝向+格局)", "category": "筛选", "score": 25,
         "check_type": "excel_has_advanced_filter", "params": {"sheet_name": "高级筛选1", "original_row_count": 31},
         "description": "步骤3：高级筛选AND条件（朝向南北且格局4室2厅2卫）"},
        {"id": "5_4_2_04", "name": "高级筛选2-OR条件(单价+总价)", "category": "筛选", "score": 25,
         "check_type": "excel_has_advanced_filter", "params": {"sheet_name": "高级筛选2", "original_row_count": 31},
         "description": "步骤4：高级筛选OR条件（单价<5000或总价<50万）"},
    ]
}

# 任务5.4.3：数据透视表（4个评分项，满分100）
TASK_5_4_3_RULES = {
    "task_id": "task_5_4_3",
    "task_name": "5.4.3 数据透视表",
    "module": "WPS表格",
    "module_group": "wps_excel",
    "file_type": "xlsx",
    "description": "根据教材任务5.4.3要求，创建并配置数据透视表分析二手房数据。",
    "max_score": 100,
    "content_keywords": {
        "required": [],
        "any_of": ["二手房", "房源", "房价", "面积", "单价", "总价", "户型", "楼层", "小区", "朝向", "城区", "格局", "建造时间"],
        "min_match": 3,
        "description": "表格应包含与二手房数据管理相关的数据列"
    },
    "check_items": [
        {"id": "5_4_3_01", "name": "创建数据透视表", "category": "数据透视表", "score": 25,
         "check_type": "excel_has_pivot_table", "params": {},
         "description": "步骤1~2：创建数据透视表"},
        {"id": "5_4_3_02", "name": "拖动字段到行/列/值区域", "category": "数据透视表", "score": 25,
         "check_type": "excel_pivot_configured", "params": {},
         "description": "步骤3：拖动字段到行区域、列区域、值区域"},
        {"id": "5_4_3_03", "name": "设置值字段汇总方式", "category": "数据透视表", "score": 25,
         "check_type": "excel_pivot_configured", "params": {},
         "description": "步骤4：值字段设置汇总方式（求和/计数/平均值）"},
        {"id": "5_4_3_04", "name": "透视表放在新工作表", "category": "数据透视表", "score": 25,
         "check_type": "excel_sheet_named", "params": {"sheet_names": ["透视表"]},
         "description": "步骤2：将数据透视表放在新工作表'透视表'中"},
    ]
}

# 任务5.4.4：创建图表（6个评分项，满分100）
TASK_5_4_4_RULES = {
    "task_id": "task_5_4_4",
    "task_name": "5.4.4 创建图表",
    "module": "WPS表格",
    "module_group": "wps_excel",
    "file_type": "xlsx",
    "description": "根据教材任务5.4.4要求，创建簇状柱形图分析不同城区房价均值。",
    "max_score": 100,
    "content_keywords": {
        "required": [],
        "any_of": ["城区名称", "低层单价", "高层单价", "中层单价", "总均值", "单价", "房价"],
        "min_match": 3,
        "description": "表格应包含与二手房数据分析相关的数据列"
    },
    "check_items": [
        {"id": "5_4_4_01", "name": "创建簇状柱形图", "category": "创建图表", "score": 15,
         "check_type": "excel_chart_type_bar", "params": {"sheet_name": "图表分析表"},
         "description": "步骤1：创建簇状柱形图"},
        {"id": "5_4_4_02", "name": "图表标题包含'城区'和'房价'", "category": "图表美化", "score": 25,
         "check_type": "excel_chart_title_contains", "params": {"sheet_name": "图表分析表", "keywords": ["城区", "房价", "均值"]},
         "description": "步骤1~2：设置图表标题为'不同城区房价均值柱形图'"},
        {"id": "5_4_4_03", "name": "图表数据系列包含'总均值'", "category": "图表数据", "score": 25,
         "check_type": "excel_chart_series_contains", "params": {"sheet_name": "图表分析表", "keywords": ["总均值"]},
         "description": "步骤2：修改数据源为各城区总均值数据"},
        {"id": "5_4_4_04", "name": "图表有数据标签", "category": "图表美化", "score": 15,
         "check_type": "excel_has_data_labels", "params": {"sheet_name": "图表分析表"},
         "description": "步骤3：添加数据标签"},
        {"id": "5_4_4_05", "name": "图表有图例", "category": "图表美化", "score": 20,
         "check_type": "excel_chart_has_legend", "params": {"sheet_name": "图表分析表"},
         "description": "步骤2：添加图例"},
    ]
}

# 任务5.4.5：页面设置与打印（5个评分项，满分100）
TASK_5_4_5_RULES = {
    "task_id": "task_5_4_5",
    "task_name": "5.4.5 页面设置与打印",
    "module": "WPS表格",
    "module_group": "wps_excel",
    "file_type": "xlsx",
    "description": "根据教材任务5.4.5要求，设置页面布局和打印选项。",
    "max_score": 100,
    "content_keywords": {
        "required": [],
        "any_of": ["二手房", "房源", "房价", "面积", "单价", "总价", "户型", "楼层", "小区", "朝向", "城区", "格局", "建造时间"],
        "min_match": 3,
        "description": "表格应包含与二手房数据管理相关的数据列"
    },
    "check_items": [
        {"id": "5_4_5_01", "name": "设置打印区域A1:P28", "category": "页面设置", "score": 20,
         "check_type": "excel_has_print_area_exact", "params": {"expected_area": "A1:P28"},
         "description": "步骤1：设置打印区域为A1:P28"},
        {"id": "5_4_5_02", "name": "设置页边距(水平居中)", "category": "页面设置", "score": 20,
         "check_type": "excel_page_horizontal_center", "params": {},
         "description": "步骤2：页边距居中方式设为水平，上边距2厘米，下边距3厘米"},
        {"id": "5_4_5_03", "name": "设置顶端标题行", "category": "页面设置", "score": 20,
         "check_type": "excel_has_header_rows", "params": {},
         "description": "步骤3：设置顶端标题行为第1、2、3行"},
        {"id": "5_4_5_04", "name": "页面方向设为横向", "category": "页面设置", "score": 20,
         "check_type": "excel_page_landscape", "params": {},
         "description": "步骤4：设置页面方向为横向"},
        {"id": "5_4_5_05", "name": "将所有列打印在一页", "category": "页面设置", "score": 20,
         "check_type": "excel_fit_to_page", "params": {},
         "description": "步骤4：将所有列打印在一页"},
    ]
}

# ============================================================
# 模块六：WPS演示
# ============================================================

# 任务6.1：制作静态红色旅游地宣传演示文稿（16个评分项，满分100）
TASK_6_1_RULES = {
    "task_id": "task_6_1",
    "task_name": "6.1 制作静态红色旅游地宣传演示文稿",
    "module": "WPS演示",
    "module_group": "wps_ppt",
    "file_type": "pptx",
    "description": "根据教材任务6.1要求，制作红色旅游地宣传演示文稿。",
    "max_score": 100,
    "content_keywords": {
        "required": ["红色"],
        "any_of": ["红色", "旅游", "革命", "景点", "景区", "纪念", "历史", "圣地", "井冈山", "延安", "遵义"],
        "min_match": 2,
        "description": "演示文稿应包含与'红色旅游地'相关的内容"
    },
    "check_items": [
        {"id": "6_1_01", "name": "导入设计模板+编辑幻灯片母版", "category": "设计美化", "score": 12,
         "check_type": "ppt_has_slide_layout", "params": {},
         "description": "步骤3/13：导入\u201c红色.dpt\u201d模板，编辑母版标题样式、复制背景框"},
        {"id": "6_1_02", "name": "标题幻灯片(主标题+副标题)", "category": "内容结构", "score": 5,
         "check_type": "ppt_has_title_subtitle", "params": {},
         "description": "步骤4：主标题+副标题"},
        {"id": "6_1_03", "name": "主标题方正舒体48号", "category": "字符格式", "score": 7,
         "check_type": "ppt_has_wordart_with_font", "params": {"font_name": "方正舒体"},
         "description": "步骤4：主标题方正舒体48号"},
        {"id": "6_1_04", "name": "副标题预设样式", "category": "字符格式", "score": 5,
         "check_type": "ppt_has_wordart_or_text_effect", "params": {},
         "description": "步骤4：副标题预设样式"},
        {"id": "6_1_05", "name": "幻灯片至少7页", "category": "基本操作", "score": 5,
         "check_type": "ppt_min_slides", "params": {"min_slides": 7},
         "description": "步骤4~12：至少7张幻灯片"},
        {"id": "6_1_06", "name": "使用多种版式(\u22653种)", "category": "设计美化", "score": 6,
         "check_type": "ppt_multiple_layouts", "params": {"min_layouts": 3},
         "description": "步骤5~12：使用左右/导航/通用等版式"},
        {"id": "6_1_07", "name": "插入图片(\u22653张)", "category": "图文混排", "score": 5,
         "check_type": "ppt_has_images", "params": {"min_images": 3},
         "description": "步骤7/8/9：插入多张图片"},
        {"id": "6_1_08", "name": "图片大小设置+裁剪为椭圆", "category": "图文混排", "score": 11,
         "check_type": "ppt_has_image_with_effect", "params": {},
         "description": "步骤9：图片高度5厘米宽度6.5厘米，裁剪为椭圆"},
        {"id": "6_1_10", "name": "使用文本框", "category": "图文混排", "score": 5,
         "check_type": "ppt_has_textbox", "params": {},
         "description": "步骤7：插入文本框"},
        {"id": "6_1_11", "name": "使用形状", "category": "图文混排", "score": 5,
         "check_type": "ppt_has_shapes", "params": {},
         "description": "步骤8：插入圆角矩形形状"},
        {"id": "6_1_12", "name": "插入SmartArt(V型列表)", "category": "高级功能", "score": 6,
         "check_type": "ppt_has_smartart", "params": {},
         "description": "步骤11：插入V型列表SmartArt"},
        {"id": "6_1_13", "name": "SmartArt已配置内容", "category": "高级功能", "score": 6,
         "check_type": "ppt_smartart_configured", "params": {},
         "description": "步骤11：SmartArt添加项目内容"},
        {"id": "6_1_14", "name": "插入艺术字", "category": "设计美化", "score": 6,
         "check_type": "ppt_has_wordart", "params": {},
         "description": "步骤12：插入艺术字"},
        {"id": "6_1_16", "name": "创建超链接", "category": "高级功能", "score": 6,
         "check_type": "ppt_has_hyperlinks", "params": {},
         "description": "步骤14：为目录项创建超链接"},
        {"id": "6_1_17", "name": "插入幻灯片编号", "category": "高级功能", "score": 5,
         "check_type": "ppt_has_slide_numbers", "params": {},
         "description": "步骤15：插入幻灯片编号"},
        {"id": "6_1_18", "name": "结尾幻灯片", "category": "内容结构", "score": 5,
         "check_type": "ppt_has_end_slide", "params": {},
         "description": "步骤12：最后一张为结尾页"},
    ]
}

# 任务6.2：制作动态红色旅游地宣传演示文稿（8个评分项，满分100）
TASK_6_2_RULES = {
    "task_id": "task_6_2",
    "task_name": "6.2 制作动态红色旅游地宣传演示文稿",
    "module": "WPS演示",
    "module_group": "wps_ppt",
    "file_type": "pptx",
    "description": "根据教材任务6.2要求，为静态演示文稿添加动态效果。",
    "max_score": 100,
    "content_keywords": {
        "required": ["红色"],
        "any_of": ["红色", "旅游", "革命", "景点", "景区", "纪念", "历史", "圣地"],
        "min_match": 2,
        "description": "演示文稿应包含与'红色旅游地'相关的内容"
    },
    "check_items": [
        {"id": "6_2_01", "name": "幻灯片切换效果", "category": "切换效果", "score": 9,
         "check_type": "ppt_has_transitions", "params": {},
         "description": "步骤2：设置切换效果（切出-全黑）"},
        {"id": "6_2_02", "name": "自动换片时间(5秒)", "category": "切换效果", "score": 8,
         "check_type": "ppt_auto_advance", "params": {},
         "description": "步骤2：第一张设置自动换片5秒"},
        {"id": "6_2_03", "name": "自定义动画效果", "category": "进入动画", "score": 41,
         "check_type": "ppt_has_animations", "params": {},
         "description": "步骤3/5/6：设置多种自定义动画效果（飞入、擦除、阶梯状、缓慢进入、轰然下落等）"},
        {"id": "6_2_06", "name": "多张幻灯片设置动画(\u22653页)", "category": "进入动画", "score": 9,
         "check_type": "ppt_multiple_animated_slides", "params": {"min_slides": 3},
         "description": "步骤5~7：至少3张幻灯片设置动画"},
        {"id": "6_2_07", "name": "强调动画(添加下划线)", "category": "强调动画", "score": 9,
         "check_type": "ppt_has_emphasis_animation", "params": {},
         "description": "步骤6：强调动画-添加下划线"},
        {"id": "6_2_08", "name": "动画顺序调整", "category": "动画设置", "score": 8,
         "check_type": "ppt_animation_order_set", "params": {},
         "description": "步骤6：调整动画播放顺序"},
        {"id": "6_2_11", "name": "放映方式(展台自动循环)", "category": "放映设置", "score": 8,
         "check_type": "ppt_show_type_set", "params": {},
         "description": "步骤8：展台自动循环放映"},
        {"id": "6_2_12", "name": "多种动画类型(\u22653种)", "category": "进入动画", "score": 8,
         "check_type": "ppt_has_multiple_animations_types", "params": {"min_types": 3},
         "description": "步骤3~7：使用飞入、擦除、阶梯状等多种动画"},
    ]
}

# 任务6.3：分享与使用红色旅游地宣传演示文稿（5个评分项，满分100）
TASK_6_3_RULES = {
    "task_id": "task_6_3",
    "task_name": "6.3 分享与使用红色旅游地宣传演示文稿",
    "module": "WPS演示",
    "module_group": "wps_ppt",
    "file_type": "pptx",
    "description": "根据教材任务6.3要求，将演示文稿上传到WPS云文档分享。云分享操作无法通过文件检测，仅检测文件内容质量。",
    "max_score": 100,
    "content_keywords": {
        "required": ["红色"],
        "any_of": ["红色", "旅游", "革命", "景点", "景区", "纪念", "历史", "圣地"],
        "min_match": 2,
        "description": "演示文稿应包含与'红色旅游地'相关的内容"
    },
    "check_items": [
        {"id": "6_3_01", "name": "演示文稿至少5页", "category": "内容完整性", "score": 20,
         "check_type": "ppt_min_slides", "params": {"min_slides": 5},
         "description": "步骤1：演示文稿至少5页"},
        {"id": "6_3_02", "name": "包含标题和结尾幻灯片", "category": "内容结构", "score": 20,
         "check_type": "ppt_has_title_and_end", "params": {},
         "description": "步骤1：包含标题和结尾幻灯片"},
        {"id": "6_3_03", "name": "包含图片和文本框", "category": "图文混排", "score": 20,
         "check_type": "ppt_has_images_and_textbox", "params": {},
         "description": "幻灯片中包含图片和文本框"},
        {"id": "6_3_04", "name": "设置动画或切换效果", "category": "动态效果", "score": 20,
         "check_type": "ppt_has_transitions_or_animations", "params": {},
         "description": "设置了切换效果或对象动画"},
        {"id": "6_3_05", "name": "使用多种版式(\u22653种)", "category": "设计美化", "score": 20,
         "check_type": "ppt_multiple_layouts", "params": {"min_layouts": 3},
         "description": "使用至少3种不同的幻灯片版式"},
    ]
}

# ============================================================
# 所有可用任务（按模块分组）
# ============================================================
ALL_TASKS = {
    "task_4_1": TASK_4_1_RULES,
    "task_4_2": TASK_4_2_RULES,
    "task_4_3": TASK_4_3_RULES,
    "task_4_4": TASK_4_4_RULES,
    "task_5_1": TASK_5_1_RULES,
    "task_5_2_1": TASK_5_2_1_RULES,
    "task_5_2_2": TASK_5_2_2_RULES,
    "task_5_2_3": TASK_5_2_3_RULES,
    "task_5_2_4": TASK_5_2_4_RULES,
    "task_5_3_1": TASK_5_3_1_RULES,
    "task_5_3_2": TASK_5_3_2_RULES,
    "task_5_4_1": TASK_5_4_1_RULES,
    "task_5_4_2": TASK_5_4_2_RULES,
    "task_5_4_3": TASK_5_4_3_RULES,
    "task_5_4_4": TASK_5_4_4_RULES,
    "task_5_4_5": TASK_5_4_5_RULES,
    "task_6_1": TASK_6_1_RULES,
    "task_6_2": TASK_6_2_RULES,
    "task_6_3": TASK_6_3_RULES,
}

MODULE_GROUPS = {
    "wps_word": {
        "name": "WPS文字",
        "icon": "bi-file-earmark-word",
        "color": "#2b5797",
        "tasks": ["task_4_1", "task_4_2", "task_4_3", "task_4_4"]
    },
    "wps_excel": {
        "name": "WPS表格",
        "icon": "bi-file-earmark-excel",
        "color": "#217346",
        "tasks": ["task_5_1", "task_5_2_1", "task_5_2_2", "task_5_2_3", "task_5_2_4", "task_5_3_1", "task_5_3_2", "task_5_4_1", "task_5_4_2", "task_5_4_3", "task_5_4_4", "task_5_4_5"]
    },
    "wps_ppt": {
        "name": "WPS演示",
        "icon": "bi-file-earmark-ppt",
        "color": "#d24726",
        "tasks": ["task_6_1", "task_6_2", "task_6_3"]
    },
}

LEVEL_RULES = {
    "excellent": {"name": "优秀", "min_score": 85, "max_score": 100, "color": "#28a745",
                  "description": "掌握扎实，可挑战拓展任务"},
    "good": {"name": "良好", "min_score": 70, "max_score": 84, "color": "#17a2b8",
             "description": "基本掌握，可进行巩固提升"},
    "pass": {"name": "及格", "min_score": 60, "max_score": 69, "color": "#ffc107",
             "description": "初步掌握，需加强基础练习"},
    "fail": {"name": "需努力", "min_score": 0, "max_score": 59, "color": "#dc3545",
             "description": "基础薄弱，需从基础操作开始学习"},
}

def get_level(score):
    for key, rule in LEVEL_RULES.items():
        if rule["min_score"] <= score <= rule["max_score"]:
            return key, rule
    return "fail", LEVEL_RULES["fail"]
