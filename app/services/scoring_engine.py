# -*- coding: utf-8 -*-
"""
评分引擎
根据评分规则对解析后的文件进行评分
"""
import os
import json

from .file_parser import DocxParser, XlsxParser, PptxParser
from .scoring_rules import ALL_TASKS, get_level


class ScoringEngine:
    """评分引擎"""
    
    def __init__(self, task_id, filepath):
        self.task_id = task_id
        self.filepath = filepath
        self.rules = ALL_TASKS.get(task_id)
        if not self.rules:
            raise ValueError(f"未找到任务 {task_id} 的评分规则")
        self.parser = None
        self.results = []
        self.total_score = 0
        self.max_score = self.rules["max_score"]
    
    def score(self):
        """执行评分"""
        if self.rules["file_type"] == "docx":
            self.parser = DocxParser(self.filepath)
        elif self.rules["file_type"] == "xlsx":
            self.parser = XlsxParser(self.filepath)
        elif self.rules["file_type"] == "pptx":
            self.parser = PptxParser(self.filepath)
        else:
            raise ValueError(f"不支持的文件类型: {self.rules['file_type']}")
        
        # ===== 前置校验：文档内容是否匹配任务 =====
        # 注释掉：去掉文档内容匹配检测
        # content_check = self._check_content_match()
        # if not content_check["matched"]:
        #     # 内容不匹配，直接返回0分
        #     return {
        #         "task_id": self.task_id,
        #         "task_name": self.rules["task_name"],
        #         "module": self.rules["module"],
        #         "total_score": 0,
        #         "max_score": self.max_score,
        #         "percentage": 0,
        #         "level": "fail",
        #         "level_name": "需努力",
        #         "level_color": "#dc3545",
        #         "level_desc": "文档内容与任务不匹配",
        #         "details": [content_check["detail"]],
        #         "category_summary": {"内容校验": {"passed": 0, "total": 1, "score": 0, "max_score": self.max_score}},
        #         "content_warning": content_check["message"],
        #     }
        
        for item in self.rules["check_items"]:
            # 如果指定了工作表名称，切换到对应工作表
            params = item.get("params", {})
            sheet_name = params.get("sheet_name")
            if sheet_name and hasattr(self.parser, 'switch_sheet'):
                self.parser.switch_sheet(sheet_name)

            result = self._check_item(item)
            self.results.append(result)
            if result["passed"]:
                self.total_score += item["score"]
        
        level_key, level_info = get_level(self.total_score)
        
        return {
            "task_id": self.task_id,
            "task_name": self.rules["task_name"],
            "module": self.rules["module"],
            "total_score": self.total_score,
            "max_score": self.max_score,
            "percentage": round(self.total_score / self.max_score * 100, 1),
            "level": level_key,
            "level_name": level_info["name"],
            "level_color": level_info["color"],
            "level_desc": level_info["description"],
            "details": self.results,
            "category_summary": self._get_category_summary(),
        }
    
    def _get_material_keywords(self):
        """从素材文件索引中获取关键词（只保留表头关键词，过滤数据值）"""
        try:
            project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            materials_dir = os.path.join(project_root, 'data', 'materials')
            index_path = os.path.join(materials_dir, 'index.json')
            if not os.path.exists(index_path):
                return None
            with open(index_path, 'r', encoding='utf-8') as f:
                index = json.load(f)
            materials = index.get(self.task_id, [])
            if not materials:
                return None
            raw_keywords = materials[-1].get('keywords', [])
            # 过滤：保留长度>=2、数字占比不超过一半的（排除纯数据值）
            filtered = []
            for kw in raw_keywords:
                if len(kw) < 2:
                    continue
                digit_count = sum(1 for c in kw if c.isdigit())
                if digit_count > len(kw) * 0.5:
                    continue
                filtered.append(kw)
            return filtered if filtered else raw_keywords
        except Exception as e:
            import traceback
            traceback.print_exc()
            return None

    def _check_content_match(self):
        """校验文档内容是否匹配当前任务"""
        keywords_config = self.rules.get("content_keywords")
        if not keywords_config:
            return {"matched": True, "message": "", "detail": None}
        
        # 优先使用素材文件关键词
        material_keywords = self._get_material_keywords()
        if material_keywords:
            # 用素材关键词替换any_of
            keywords_config = dict(keywords_config)
            keywords_config["any_of"] = material_keywords
            keywords_config["required"] = []  # 素材匹配不要求必选关键词
            keywords_config["min_match"] = max(2, len(material_keywords) // 3)
        else:
            # 没有素材文件索引时，放宽匹配条件：不要求必选关键词，降低最低匹配数
            keywords_config = dict(keywords_config)
            keywords_config["required"] = []
            keywords_config["min_match"] = 1
        
        # 提取文档全部文本
        doc_text = ""
        try:
            if self.rules["file_type"] == "docx":
                doc_text = self.parser.get_all_text()
            elif self.rules["file_type"] == "xlsx":
                doc_text = self.parser.get_all_text()
            elif self.rules["file_type"] == "pptx":
                doc_text = self.parser.get_all_text()
        except Exception:
            # 如果无法提取文本，跳过校验
            return {"matched": True, "message": "", "detail": None}
        
        if not doc_text or len(doc_text.strip()) < 10:
            return {"matched": True, "message": "", "detail": None}
        
        doc_text_lower = doc_text.lower()
        
        # 检查必选关键词
        required = keywords_config.get("required", [])
        any_of = keywords_config.get("any_of", [])
        min_match = keywords_config.get("min_match", 2)
        description = keywords_config.get("description", "")
        
        # 统计匹配的关键词
        matched_keywords = []
        for kw in any_of:
            if kw.lower() in doc_text_lower:
                matched_keywords.append(kw)
        
        # 检查必选关键词
        missing_required = []
        for kw in required:
            if kw.lower() not in doc_text_lower:
                missing_required.append(kw)
        
        # 判断是否匹配
        is_matched = len(matched_keywords) >= min_match and len(missing_required) == 0
        
        if is_matched:
            return {"matched": True, "message": "", "detail": None}
        
        # 构建不匹配的详细信息
        if missing_required:
            msg = f"文档缺少必要内容关键词：{', '.join(missing_required)}。{description}"
        else:
            found = len(matched_keywords)
            msg = f"文档内容与任务不匹配（匹配关键词{found}/{min_match}）。{description}"
        
        return {
            "matched": False,
            "message": msg,
            "detail": {
                "id": "content_check",
                "name": "文档内容校验",
                "category": "内容校验",
                "score": 0,
                "max_score": self.max_score,
                "passed": False,
                "message": msg,
            }
        }
    
    def _check_item(self, item):
        """检查单个评分项"""
        check_type = item["check_type"]
        params = item.get("params", {})
        passed = False
        message = ""
        
        try:
            if check_type == "title_font":
                passed = self.parser.check_title_font(params["font_name"])
                message = f"标题字体{'包含' if passed else '未检测到'}\"{params['font_name']}\""
            
            elif check_type == "title_font_size_exact":
                target_pt = params["exact_size_pt"]
                passed = self.parser.check_title_font_size_exact(target_pt)
                message = f"标题字号{'等于' if passed else '不等于'}{target_pt}磅"

            elif check_type == "title_font_size":
                if "exact_size_pt" in params:
                    target_pt = params["exact_size_pt"]
                    passed = self.parser.check_title_font_size_exact(target_pt)
                    message = f"标题字号{'等于' if passed else '不等于'}{target_pt}磅"
                else:
                    min_pt = params.get("min_size_pt", 0)
                    passed = self.parser.check_title_font_size(min_pt)
                    message = f"标题字号{'≥' if passed else '<'}{min_pt}磅"
            
            elif check_type == "title_bold":
                passed = self.parser.check_title_bold()
                message = f"标题{'已' if passed else '未'}设置加粗"
            
            elif check_type == "title_alignment":
                align_names = {"center": "居中", "left": "左对齐", "right": "右对齐", "justify": "两端对齐"}
                passed = self.parser.check_title_alignment(params["alignment"])
                message = f"标题{'已' if passed else '未'}设置为{align_names.get(params['alignment'], params['alignment'])}"
            
            elif check_type == "body_font":
                passed = self.parser.check_body_font(params["font_name"])
                message = f"正文字体{'包含' if passed else '未检测到'}\"{params['font_name']}\""
            
            elif check_type == "body_first_indent":
                passed = self.parser.check_body_first_indent(params["chars"])
                message = f"正文{'已' if passed else '未'}设置首行缩进{params['chars']}字符"
            
            elif check_type == "body_line_spacing_exact":
                target_sp = params["exact_spacing"]
                passed = self.parser.check_body_line_spacing_exact(target_sp)
                message = f"正文行距{'等于' if passed else '不等于'}固定值{target_sp}磅"

            elif check_type == "body_line_spacing":
                if "exact_spacing" in params:
                    target_sp = params["exact_spacing"]
                    passed = self.parser.check_body_line_spacing_exact(target_sp)
                    message = f"正文行距{'等于' if passed else '不等于'}{target_sp}"
                else:
                    min_sp = params.get("min_spacing", 0)
                    passed = self.parser.check_body_line_spacing(min_sp)
                    message = f"正文行距{'≥' if passed else '<'}{min_sp}倍"
            
            elif check_type == "body_para_spacing":
                passed = self.parser.check_body_para_spacing(params["min_before"], params["min_after"])
                message = f"正文{'已' if passed else '未'}设置段间距"
            
            elif check_type == "body_alignment":
                align_names = {"justify": "两端对齐", "left": "左对齐", "center": "居中"}
                passed = self.parser.check_body_alignment(params["alignment"])
                message = f"正文{'已' if passed else '未'}设置为{align_names.get(params['alignment'], '')}"
            
            elif check_type == "section_columns":
                passed = self.parser.check_section_columns(params["min_columns"])
                message = f"{'已' if passed else '未'}设置{params['min_columns']}栏及以上分栏"
            
            elif check_type == "drop_cap":
                passed = self.parser.check_drop_cap()
                message = f"{'已' if passed else '未'}设置首字下沉"
            
            elif check_type == "paragraph_border_or_shading":
                passed = self.parser.check_paragraph_border_or_shading()
                message = f"{'已' if passed else '未'}设置段落边框或底纹"
            
            elif check_type == "has_image":
                passed = self.parser.check_has_image()
                count = len(self.parser.inline_shapes) if hasattr(self.parser, 'inline_shapes') else 0
                message = f"文档{'包含' if passed else '未包含'}图片（检测到{count}张）"
            
            elif check_type == "has_wordart":
                passed = self.parser.check_has_wordart()
                message = f"{'已' if passed else '未'}插入艺术字"
            
            elif check_type == "image_wrap_type":
                passed = self.parser.check_image_wrap_type(params.get("non_inline", False))
                message = f"图片环绕方式{'为非嵌入式' if passed else '为默认嵌入式'}"
            
            # Excel 检查项
            elif check_type == "excel_merged_title":
                passed = self.parser.check_merged_title(params["row"], params["min_cols"])
                message = f"标题行{'已' if passed else '未'}合并居中"
            
            elif check_type == "excel_header_bold":
                passed = self.parser.check_header_bold(params["header_row"])
                message = f"表头{'已' if passed else '未'}设置加粗"
            
            elif check_type == "excel_min_columns":
                actual = self.parser.ws.max_column
                passed = self.parser.check_min_columns(params["min_cols"])
                message = f"数据列数{actual}，{'≥' if passed else '<'}要求{params['min_cols']}列"
            
            elif check_type == "excel_min_data_rows":
                header_rows = params.get("header_rows", 1)
                actual = self.parser.ws.max_row - header_rows
                passed = self.parser.check_min_data_rows(params["min_rows"], header_rows)
                message = f"数据行数{actual}，{'≥' if passed else '<'}要求{params['min_rows']}行"
            
            elif check_type == "excel_has_borders":
                passed = self.parser.check_has_borders()
                message = f"数据区域{'已' if passed else '未'}设置边框"
            
            elif check_type == "excel_column_width":
                passed = self.parser.check_column_width(params["min_width"])
                message = f"列宽{'已' if passed else '未'}合理设置"
            
            elif check_type == "excel_number_format":
                passed = self.parser.check_number_format()
                message = f"{'已' if passed else '未'}设置数值格式"
            
            elif check_type == "excel_sheet_named":
                sheet_names = params.get("sheet_names")
                if sheet_names:
                    name = self.parser.ws.title
                    passed = name in sheet_names
                    message = f"工作表名称为\"{name}\"，{'匹配' if passed else '不匹配'}要求名称"
                else:
                    passed = self.parser.check_sheet_named()
                    name = self.parser.ws.title
                    message = f"工作表名称为\"{name}\"，{'已' if passed else '未'}自定义命名"
            
            # PPT 检查项
            elif check_type == "ppt_min_slides":
                actual = len(self.parser.slides)
                passed = self.parser.check_min_slides(params["min_slides"])
                message = f"幻灯片{actual}页，{'≥' if passed else '<'}要求{params['min_slides']}页"
            
            elif check_type == "ppt_has_title_slide":
                passed = self.parser.check_has_title_slide()
                message = f"{'有' if passed else '无'}标题幻灯片"
            
            elif check_type == "ppt_has_slide_layout":
                passed = self.parser.check_has_slide_layout()
                message = f"{'已' if passed else '未'}使用幻灯片布局/模板"
            
            elif check_type == "ppt_has_images":
                passed = self.parser.check_has_images(params["min_images"])
                message = f"{'包含' if passed else '未包含足够'}图片（要求≥{params['min_images']}张）"
            
            elif check_type == "ppt_has_textbox":
                passed = self.parser.check_has_textbox()
                message = f"{'已' if passed else '未'}使用文本框"
            
            elif check_type == "ppt_has_transitions":
                passed = self.parser.check_has_transitions()
                message = f"{'已' if passed else '未'}设置幻灯片切换效果"
            
            elif check_type == "ppt_has_animations":
                passed = self.parser.check_has_animations()
                message = f"{'已' if passed else '未'}设置自定义动画"
            
            elif check_type == "ppt_has_end_slide":
                passed = self.parser.check_has_end_slide()
                message = f"{'有' if passed else '无'}结尾幻灯片"

            # Docx 新增检查项
            elif check_type == "doc_has_right_align":
                passed = self.parser.doc_has_right_align()
                message = f"文档{'有' if passed else '无'}右对齐段落"

            elif check_type == "doc_has_date":
                passed = self.parser.doc_has_date()
                message = f"文档{'包含' if passed else '未包含'}日期文本"

            elif check_type == "doc_has_page_breaks":
                passed = self.parser.doc_has_page_breaks(params["min_breaks"])
                message = f"文档{'包含' if passed else '未包含足够'}分页符（要求≥{params['min_breaks']}个）"

            elif check_type == "doc_page_margins":
                passed = self.parser.doc_page_margins(params["left_cm"], params["right_cm"])
                message = f"页边距{'符合' if passed else '不符合'}要求（左{params['left_cm']}cm，右{params['right_cm']}cm）"

            elif check_type == "doc_page_margins_tb":
                passed = self.parser.doc_page_margins_tb(params["top_cm"], params["bottom_cm"])
                message = f"上下页边距{'符合' if passed else '不符合'}要求（上{params['top_cm']}cm，下{params['bottom_cm']}cm）"

            elif check_type == "doc_has_header":
                passed = self.parser.doc_has_header()
                message = f"文档{'有' if passed else '无'}页眉"

            elif check_type == "doc_has_page_number":
                passed = self.parser.doc_has_page_number()
                message = f"文档{'有' if passed else '无'}页码"

            elif check_type == "doc_has_heading_style":
                passed = self.parser.doc_has_heading_style()
                message = f"文档{'已' if passed else '未'}使用标题样式"

            elif check_type == "doc_heading1_centered":
                passed = self.parser.doc_heading1_centered()
                message = f"标题1{'已' if passed else '未'}居中对齐"

            elif check_type == "doc_has_toc":
                passed = self.parser.doc_has_toc()
                message = f"文档{'有' if passed else '无'}目录"

            elif check_type == "doc_toc_centered":
                passed = self.parser.doc_toc_centered()
                message = f"目录标题{'已' if passed else '未'}居中"

            elif check_type == "doc_has_table":
                passed = self.parser.doc_has_table(params["min_rows"], params["min_cols"])
                message = f"文档{'包含' if passed else '未包含'}满足要求的表格（≥{params['min_rows']}行×{params['min_cols']}列）"

            elif check_type == "doc_has_merged_cells":
                passed = self.parser.doc_has_merged_cells()
                message = f"表格{'有' if passed else '无'}合并单元格"

            elif check_type == "doc_table_has_borders":
                passed = self.parser.doc_table_has_borders()
                message = f"表格{'有' if passed else '无'}边框"

            elif check_type == "doc_table_has_shading":
                passed = self.parser.doc_table_has_shading()
                message = f"表格{'有' if passed else '无'}底纹"

            elif check_type == "doc_table_alignment":
                passed = self.parser.doc_table_alignment()
                message = f"表格内容{'已' if passed else '未'}设置对齐方式"

            elif check_type == "title_bold_center":
                passed = self.parser.title_bold_center()
                message = f"标题{'已' if passed else '未'}设置加粗居中"

            # Excel 新增检查项
            elif check_type == "excel_title_font":
                passed = self.parser.excel_title_font()
                message = f"标题行{'已' if passed else '未'}设置字体"

            elif check_type == "excel_data_font_set":
                passed = self.parser.excel_data_font_set()
                message = f"数据区{'已' if passed else '未'}设置字体"

            elif check_type == "excel_alignment_set":
                passed = self.parser.excel_alignment_set()
                message = f"单元格{'已' if passed else '未'}设置对齐方式"

            elif check_type == "excel_has_conditional_format":
                passed = self.parser.excel_has_conditional_format()
                message = f"{'已' if passed else '未'}设置条件格式"

            elif check_type == "excel_has_data_bars":
                passed = self.parser.excel_has_data_bars()
                message = f"{'已' if passed else '未'}设置数据条"

            elif check_type == "excel_has_data_validation":
                passed = self.parser.excel_has_data_validation()
                message = f"{'已' if passed else '未'}设置数据有效性"

            elif check_type == "excel_has_validation_type":
                vtype = params.get("validation_type", "whole")
                type_names = {"whole": "整数", "decimal": "小数", "list": "序列", "textLength": "文本长度", "date": "日期", "custom": "自定义"}
                passed = self.parser.excel_has_validation_type(vtype)
                message = f"{'已' if passed else '未'}设置{type_names.get(vtype, vtype)}验证"

            elif check_type == "excel_has_validation_type_with_formula":
                vtype = params.get("validation_type", "list")
                keyword = params.get("formula_keyword", "")
                type_names = {"whole": "整数", "decimal": "小数", "list": "序列", "textLength": "文本长度", "date": "日期", "custom": "自定义"}
                passed = self.parser.excel_has_validation_type_with_formula(vtype, keyword)
                message = f"{'已' if passed else '未'}设置{type_names.get(vtype, vtype)}验证(含'{keyword}')"

            elif check_type == "excel_has_table_style":
                passed = self.parser.excel_has_table_style()
                message = f"{'已' if passed else '未'}应用表格样式"

            elif check_type == "excel_date_format":
                passed = self.parser.excel_date_format()
                message = f"{'已' if passed else '未'}设置日期格式"

            elif check_type == "excel_has_formula":
                col = params.get("col")
                formula_kw = params.get("formula_keyword")
                if col is not None:
                    # 按列号检测公式
                    col_letter = chr(64 + col) if col <= 26 else str(col)
                    passed = self.parser.excel_has_formula_in_col(col)
                    message = f"{'包含' if passed else '未包含'}{col_letter}列公式"
                elif formula_kw:
                    if isinstance(formula_kw, list):
                        kw_display = '、'.join(formula_kw)
                    else:
                        kw_display = formula_kw
                    passed = self.parser.excel_has_formula(formula_kw)
                    message = f"{'包含' if passed else '未包含'}公式\"{kw_display}\""
                else:
                    passed = False
                    message = "未指定公式检测条件"

            elif check_type == "excel_decimal_format":
                passed = self.parser.excel_decimal_format()
                message = f"{'已' if passed else '未'}设置小数位格式"

            elif check_type == "excel_result_area_formatted":
                passed = self.parser.excel_result_area_formatted()
                message = f"结果区域{'已' if passed else '未'}设置格式"

            elif check_type == "excel_data_sorted":
                sort_col = params.get("sort_col", 0)
                passed = self.parser.excel_data_sorted(sort_col=sort_col)
                if sort_col > 0:
                    col_letter = chr(64 + sort_col) if sort_col <= 26 else str(sort_col)
                    message = f"{'已' if passed else '未'}对{col_letter}列数据进行排序"
                else:
                    message = f"数据{'已' if passed else '未'}排序"

            elif check_type == "excel_has_subtotals":
                passed = self.parser.excel_has_subtotals()
                message = f"{'已' if passed else '未'}设置分类汇总"

            elif check_type == "excel_has_autofilter":
                passed = self.parser.excel_has_autofilter()
                message = f"{'已' if passed else '未'}设置自动筛选"

            elif check_type == "excel_has_advanced_filter":
                original_rows = params.get("original_row_count", 31)
                passed = self.parser.excel_has_advanced_filter(original_row_count=original_rows)
                message = f"{'已' if passed else '未'}设置高级筛选"

            elif check_type == "excel_has_pivot_table":
                passed = self.parser.excel_has_pivot_table()
                message = f"{'包含' if passed else '未包含'}数据透视表"

            elif check_type == "excel_pivot_configured":
                passed = self.parser.excel_pivot_configured()
                message = f"数据透视表{'已' if passed else '未'}正确配置"

            elif check_type == "excel_has_chart":
                passed = self.parser.excel_has_chart()
                message = f"{'包含' if passed else '未包含'}图表"

            elif check_type == "excel_chart_type_bar":
                passed = self.parser.excel_chart_type_bar()
                message = f"{'已' if passed else '未'}创建柱形图(BarChart)"

            elif check_type == "excel_chart_title_contains":
                keywords = params.get("keywords", [])
                passed = self.parser.excel_chart_title_contains(keywords)
                kw_display = "、".join(keywords)
                message = f"图表标题{'包含' if passed else '不包含'}关键词（{kw_display}）"

            elif check_type == "excel_chart_has_legend":
                passed = self.parser.excel_chart_has_legend()
                message = f"图表{'已' if passed else '未'}添加图例"

            elif check_type == "excel_chart_axis_formatted":
                passed = self.parser.excel_chart_axis_formatted()
                message = f"坐标轴{'已' if passed else '未'}设置填充颜色"

            elif check_type == "excel_chart_plot_area_filled":
                passed = self.parser.excel_chart_plot_area_filled()
                message = f"绘图区{'已' if passed else '未'}设置填充颜色"

            elif check_type == "excel_chart_area_filled":
                passed = self.parser.excel_chart_area_filled()
                message = f"图表区{'已' if passed else '未'}设置填充颜色"

            elif check_type == "excel_chart_series_count":
                min_series = params.get("min_series", 2)
                passed = self.parser.excel_chart_series_count(min_series)
                message = f"图表数据系列{'≥' if passed else '<'}{min_series}个"

            elif check_type == "excel_chart_series_contains":
                keywords = params.get("keywords", [])
                passed = self.parser.excel_chart_series_contains(keywords)
                kw_display = "、".join(keywords)
                message = f"图表数据系列{'包含' if passed else '不包含'}关键词（{kw_display}）"

            elif check_type == "excel_has_data_labels_show_val":
                passed = self.parser.excel_has_data_labels_show_val()
                message = f"数据标签{'已' if passed else '未'}设置为显示数值"

            elif check_type == "excel_chart_has_title":
                passed = self.parser.excel_chart_has_title()
                message = f"图表{'有' if passed else '无'}标题"

            elif check_type == "excel_has_data_labels":
                passed = self.parser.excel_has_data_labels()
                message = f"{'已' if passed else '未'}添加数据标签"

            elif check_type == "excel_no_gridlines":
                passed = self.parser.excel_no_gridlines()
                message = f"{'已' if passed else '未'}删除网格线"

            # PPT 新增检查项
            elif check_type == "ppt_has_title_subtitle":
                passed = self.parser.ppt_has_title_subtitle()
                message = f"标题幻灯片{'有' if passed else '无'}主副标题"

            elif check_type == "ppt_multiple_layouts":
                passed = self.parser.ppt_multiple_layouts(params["min_layouts"])
                message = f"使用了{len(set(s.slide_layout.name for s in self.parser.slides if s.slide_layout))}种版式，{'≥' if passed else '<'}要求{params['min_layouts']}种"

            elif check_type == "ppt_has_wordart":
                passed = self.parser.ppt_has_wordart()
                message = f"{'包含' if passed else '未包含'}艺术字"

            elif check_type == "ppt_has_shapes":
                passed = self.parser.ppt_has_shapes()
                message = f"{'包含' if passed else '未包含'}形状"

            elif check_type == "ppt_has_hyperlinks":
                passed = self.parser.ppt_has_hyperlinks()
                message = f"{'包含' if passed else '未包含'}超链接"

            elif check_type == "ppt_has_slide_numbers":
                passed = self.parser.ppt_has_slide_numbers()
                message = f"{'有' if passed else '无'}幻灯片编号"

            elif check_type == "ppt_has_smartart":
                passed = self.parser.ppt_has_smartart()
                message = f"{'包含' if passed else '未包含'}智能图形"

            elif check_type == "ppt_multiple_animated_slides":
                passed = self.parser.ppt_multiple_animated_slides(params["min_slides"])
                message = f"{'≥' if passed else '<'}{params['min_slides']}页幻灯片设置了动画"

            elif check_type == "ppt_has_emphasis_animation":
                passed = self.parser.ppt_has_emphasis_animation()
                message = f"{'包含' if passed else '未包含'}强调动画"

            elif check_type == "ppt_has_title_and_end":
                passed = self.parser.ppt_has_title_and_end()
                message = f"{'有' if passed else '无'}标题和结尾幻灯片"

            elif check_type == "ppt_has_images_and_textbox":
                passed = self.parser.ppt_has_images_and_textbox()
                message = f"{'同时包含' if passed else '未同时包含'}图片和文本框"

            elif check_type == "ppt_has_transitions_or_animations":
                passed = self.parser.ppt_has_transitions_or_animations()
                message = f"{'已' if passed else '未'}设置切换效果或动画"

            # 页面边框自动检测
            elif check_type == "doc_has_page_border":
                passed = self.parser.doc_has_page_border()
                message = f"{'已' if passed else '未'}设置页面边框"

            # 页面边框精确检测（粗细+可选样式+可选颜色）
            elif check_type == "doc_page_border_exact":
                sz_pt = params.get("border_sz_pt", 1.5)
                style = params.get("border_style")  # 可选
                color = params.get("border_color")   # 可选
                passed = self.parser.doc_page_border_exact(sz_pt, style, color)
                style_msg = f"、样式{style}" if style else ""
                color_msg = f"、颜色{color}" if color else ""
                message = f"页面边框{'符合' if passed else '不符合'}要求（{sz_pt}磅{style_msg}{color_msg}）"

            # 表格公式自动检测
            elif check_type == "doc_has_table_formula":
                passed = self.parser.doc_has_table_formula()
                message = f"表格{'包含' if passed else '未包含'}公式计算"

            # 表格边框精确检测（粗细+颜色）
            elif check_type == "doc_table_border_exact":
                sz_pt = params.get("border_sz_pt", 1.5)
                color = params.get("border_color")
                passed = self.parser.doc_table_border_exact(sz_pt, color)
                color_msg = f"、颜色{color}" if color else ""
                message = f"表格边框{'符合' if passed else '不符合'}要求（{sz_pt}磅{color_msg}）"

            # 标题字体颜色
            elif check_type == "title_font_color":
                color = params.get("color_hex", "000000")
                passed = self.parser.title_font_color(color)
                message = f"标题字体颜色{'为' if passed else '不为'}#{color}"

            # 正文字体颜色
            elif check_type == "body_font_color":
                color = params.get("color_hex", "000000")
                passed = self.parser.body_font_color(color)
                message = f"正文字体颜色{'为' if passed else '不为'}#{color}"

            # 正文字号精确检测
            elif check_type == "body_font_size_exact":
                size_pt = params.get("exact_size_pt", 12)
                passed = self.parser.body_font_size_exact(size_pt)
                message = f"正文字号{'等于' if passed else '不等于'}{size_pt}磅"

            # 图片尺寸检测
            elif check_type == "check_image_size":
                passed = self.parser.check_image_size(
                    min_width_cm=params.get("min_width_cm"),
                    min_height_cm=params.get("min_height_cm"),
                    max_width_cm=params.get("max_width_cm"),
                    max_height_cm=params.get("max_height_cm")
                )
                size_info = []
                if params.get("min_width_cm"): size_info.append(f"宽≥{params['min_width_cm']}cm")
                if params.get("min_height_cm"): size_info.append(f"高≥{params['min_height_cm']}cm")
                if params.get("max_width_cm"): size_info.append(f"宽≤{params['max_width_cm']}cm")
                if params.get("max_height_cm"): size_info.append(f"高≤{params['max_height_cm']}cm")
                message = f"图片尺寸{'符合' if passed else '不符合'}要求（{'、'.join(size_info)}）"

            # 图片位置/环绕方式检测
            elif check_type == "check_image_position":
                pos = params.get("position_type", "inline")
                passed = self.parser.check_image_position(pos)
                pos_names = {"inline": "嵌入式", "behind_text": "衬于文字下方", "in_front_of_text": "浮于文字上方", "non_inline": "非嵌入式（浮于/衬于文字）"}
                message = f"图片环绕方式{'为' if passed else '不为'}{pos_names.get(pos, pos)}"

            # 段落底纹颜色精确检测
            elif check_type == "paragraph_shading_color":
                color = params.get("color_hex", "auto")
                passed = self.parser.paragraph_shading_color(color)
                message = f"段落底纹颜色{'为' if passed else '不为'}#{color}"

            # 标题字体精确匹配
            elif check_type == "title_font_name_exact":
                font = params.get("font_name", "")
                passed = self.parser.title_font_name_exact(font)
                message = f"标题字体{'精确匹配' if passed else '不匹配'}\"{font}\""

            elif check_type == "title_font_name":
                font = params.get("font_name", "")
                passed = self.parser.doc_title_font_name(font)
                message = f"标题字体{'为' if passed else '不为'}\"{font}\""

            elif check_type == "body_font_name_exact":
                font = params.get("font_name", "")
                passed = self.parser.body_font_name_exact(font)
                message = f"正文字体{'精确匹配' if passed else '不匹配'}\"{font}\""

            elif check_type == "ppt_show_type_set":
                passed = True
                message = "需人工复核：演示文稿放映类型设置"

            elif check_type == "ppt_animation_order_set":
                passed = True
                message = "需人工复核：动画顺序设置"

            elif check_type == "ppt_transition_speed_set":
                passed = True
                message = "需人工复核：切换速度设置"

            elif check_type == "ppt_auto_advance":
                passed = True
                message = "需人工复核：自动换片设置"

            elif check_type == "doc_min_paragraphs":
                min_p = params.get("min_paragraphs", 3)
                actual = len([p for p in self.parser.paragraphs if p.text.strip()])
                passed = actual >= min_p
                message = f"文档共{actual}个非空段落，{'≥' if passed else '<'}要求{min_p}段"

            # ===== 新增检测类型 =====
            elif check_type == "file_type_valid":
                passed = True
                message = "文件格式正确"

            elif check_type == "doc_content_complete":
                min_chars = params.get("min_chars", 100)
                total_chars = sum(len(p.text) for p in self.parser.paragraphs)
                for t in self.parser.tables:
                    for row in t.rows:
                        for cell in row.cells:
                            total_chars += len(cell.text)
                passed = total_chars >= min_chars
                message = f"文档共{total_chars}字，{'≥' if passed else '<'}要求{min_chars}字"

            elif check_type == "doc_has_multiple_images":
                min_count = params.get("min_count", 2)
                passed = self.parser.doc_has_multiple_images(min_count)
                message = f"图片数量{'≥' if passed else '<'}{min_count}张"

            elif check_type == "doc_has_wordart_or_text_effect":
                passed = self.parser.doc_has_wordart_or_text_effect()
                message = f"{'已' if passed else '未'}设置艺术字或文字效果"

            elif check_type == "excel_has_multiple_validations":
                min_count = params.get("min_count", 2)
                passed = self.parser.excel_has_multiple_validations(min_count)
                message = f"数据有效性规则{'≥' if passed else '<'}{min_count}条"

            elif check_type == "excel_has_multiple_conditional_formats":
                min_count = params.get("min_count", 2)
                passed = self.parser.excel_has_multiple_conditional_formats(min_count)
                message = f"条件格式规则{'≥' if passed else '<'}{min_count}条"

            elif check_type == "excel_has_number_format_custom":
                passed = self.parser.excel_has_number_format_custom()
                message = f"{'已' if passed else '未'}设置自定义数值格式"

            elif check_type == "excel_cell_number_format":
                col_start = params.get("col_start", 1)
                col_end = params.get("col_end", 1)
                row_start = params.get("row_start", 2)
                row_end = params.get("row_end")
                format_type = params.get("format_type", "number")
                format_keyword = params.get("format_keyword")
                passed = self.parser.excel_cell_number_format(
                    col_start=col_start, col_end=col_end,
                    row_start=row_start, row_end=row_end,
                    format_type=format_type, format_keyword=format_keyword
                )
                type_names = {"decimal": "小数位", "thousands": "千位分隔符", "number": "数值", "date": "日期", "custom": "自定义"}
                col_range = f"{chr(64+col_start) if col_start<=26 else col_start}{row_start}:{chr(64+col_end) if col_end<=26 else col_end}"
                desc = type_names.get(format_type, format_type)
                if format_keyword:
                    desc += f"(含'{format_keyword}')"
                message = f"{col_range}区域{'已' if passed else '未'}设置{desc}格式"

            elif check_type == "excel_title_font_color":
                color = params.get("color_hex", "000000")
                passed = self.parser.excel_title_font_color(color)
                message = f"标题字体颜色{'为' if passed else '不为'}#{color}"

            elif check_type == "excel_cell_font_exact":
                row = params.get("row", 1)
                col_start = params.get("col_start", 1)
                col_end = params.get("col_end", 7)
                font_name = params.get("font_name")
                font_size = params.get("font_size")
                bold = params.get("bold")
                color_hex = params.get("color_hex")
                color_tolerance = params.get("color_tolerance", 60)
                passed = self.parser.excel_cell_font_exact(
                    row=row, col_start=col_start, col_end=col_end,
                    font_name=font_name, font_size=font_size, bold=bold,
                    color_hex=color_hex, color_tolerance=color_tolerance
                )
                # 生成描述信息
                desc_parts = []
                if font_name: desc_parts.append(f"字体={font_name}")
                if font_size: desc_parts.append(f"字号={font_size}")
                if bold is not None: desc_parts.append(f"加粗={bold}")
                if color_hex: desc_parts.append(f"颜色=#{color_hex}")
                desc_str = "、".join(desc_parts)
                col_range = f"{chr(64+col_start) if col_start<=26 else ''}{row}:{chr(64+col_end) if col_end<=26 else ''}{row}"
                message = f"{'已' if passed else '未'}正确设置{col_range}格式（{desc_str}）"

            elif check_type == "excel_has_print_area":
                passed = self.parser.excel_has_print_area()
                message = f"{'已' if passed else '未'}设置打印区域"

            elif check_type == "excel_has_print_area_exact":
                expected_area = params.get("expected_area", "A1:P28")
                passed = self.parser.excel_has_print_area_exact(expected_area)
                actual_pa = str(self.parser.ws.print_area) if self.parser.ws.print_area else "无"
                message = f"打印区域为\"{actual_pa}\"，{'匹配' if passed else '不匹配'}要求\"{expected_area}\""

            elif check_type == "excel_has_header_rows":
                passed = self.parser.excel_has_header_rows()
                message = f"{'已' if passed else '未'}设置顶端标题行"

            elif check_type == "excel_page_landscape":
                passed = self.parser.excel_page_landscape()
                message = f"页面方向{'已' if passed else '未'}设置为横向"

            elif check_type == "excel_fit_to_page":
                passed = self.parser.excel_fit_to_page()
                message = f"{'已' if passed else '未'}调整为一页宽打印"

            elif check_type == "excel_page_horizontal_center":
                passed = self.parser.excel_page_horizontal_center()
                message = f"页边距{'已' if passed else '未'}设置水平居中"

            elif check_type == "excel_has_keywords_in_sheet":
                keywords = params.get("keywords", [])
                passed = self.parser.excel_has_keywords_in_sheet(keywords)
                matched = sum(1 for kw in keywords if kw in self.parser.get_all_text())
                message = f"工作表中包含{matched}/{len(keywords)}个关键词，{'全部匹配' if passed else '未全部匹配'}"

            elif check_type == "ppt_smartart_configured":
                passed = self.parser.ppt_smartart_configured()
                message = f"SmartArt{'已' if passed else '未'}配置内容"

            elif check_type == "ppt_has_qr_code":
                passed = self.parser.ppt_has_qr_code()
                message = f"{'已' if passed else '未'}插入二维码"

            elif check_type == "ppt_has_image_with_effect":
                passed = self.parser.ppt_has_image_with_effect()
                message = f"图片{'已' if passed else '未'}设置效果（阴影/裁剪等）"

            elif check_type == "doc_title_font_name":
                font = params.get("font_name", "")
                passed = self.parser.doc_title_font_name(font)
                message = f"标题字体{'为' if passed else '不为'}\"{font}\""

            elif check_type == "doc_body_font_name":
                font = params.get("font_name", "")
                passed = self.parser.doc_body_font_name(font)
                message = f"正文字体{'为' if passed else '不为'}\"{font}\""

            elif check_type == "doc_char_spacing":
                spacing = params.get("spacing_cm", 0.1)
                passed = self.parser.doc_char_spacing(spacing)
                message = f"字符间距{'已' if passed else '未'}加宽{spacing}厘米"

            elif check_type == "doc_has_picture_with_size":
                h = params.get("height_cm")
                w = params.get("width_cm")
                passed = self.parser.doc_has_picture_with_size(height_cm=h, width_cm=w)
                size_info = []
                if h: size_info.append(f"高{h}cm")
                if w: size_info.append(f"宽{w}cm")
                message = f"图片大小{'符合' if passed else '不符合'}要求（{'、'.join(size_info)}）"

            elif check_type == "doc_has_wordart_with_size":
                h = params.get("height_cm")
                w = params.get("width_cm")
                passed = self.parser.doc_has_wordart_with_size(height_cm=h, width_cm=w)
                if h and w:
                    message = f"艺术字大小{'符合' if passed else '不符合'}要求（高{h}cm、宽{w}cm）"
                else:
                    message = f"艺术字{'已' if passed else '未'}设置自定义大小"

            elif check_type == "doc_has_wordart_with_position":
                passed = self.parser.doc_has_wordart_with_position()
                message = f"艺术字{'已' if passed else '未'}设置绝对定位"

            elif check_type == "doc_paragraph_line_spacing":
                spacing = params.get("spacing_pt", 50)
                passed = self.parser.doc_paragraph_line_spacing(spacing)
                message = f"段落行距{'为' if passed else '不为'}{spacing}磅"

            elif check_type == "doc_has_even_odd_header":
                passed = self.parser.doc_has_even_odd_header()
                message = f"{'已' if passed else '未'}设置奇偶页不同"

            elif check_type == "doc_has_different_first_page":
                passed = self.parser.doc_has_different_first_page()
                message = f"{'已' if passed else '未'}设置首页不同"

            elif check_type == "doc_has_heading2_style":
                passed = self.parser.doc_has_heading2_style()
                message = f"{'已' if passed else '未'}设置二级标题样式"

            elif check_type == "ppt_has_wordart_with_font":
                font = params.get("font_name", "")
                passed = self.parser.ppt_has_wordart_with_font(font)
                message = f"艺术字字体{'为' if passed else '不为'}\"{font}\""

            elif check_type == "ppt_has_wordart_with_size":
                size = params.get("font_size", 48)
                passed = self.parser.ppt_has_wordart_with_size(size)
                message = f"艺术字字号{'为' if passed else '不为'}{size}磅"

            elif check_type == "ppt_has_multiple_animations_types":
                min_types = params.get("min_types", 3)
                passed = self.parser.ppt_has_multiple_animations_types(min_types)
                message = f"动画类型{'≥' if passed else '<'}{min_types}种"

            elif check_type == "excel_data_complete":
                min_rows = params.get("min_rows", 1)
                min_cols = params.get("min_cols", 1)
                actual_rows = self.parser.ws.max_row - params.get("header_rows", 1)
                actual_cols = self.parser.ws.max_column
                passed = actual_rows >= min_rows and actual_cols >= min_cols
                message = f"数据{actual_rows}行{actual_cols}列，{'≥' if passed else '<'}要求{min_rows}行{min_cols}列"

            elif check_type == "doc_contains_keywords":
                keywords = params.get("keywords", [])
                doc_text = self.parser.get_all_text().lower()
                matched = [kw for kw in keywords if kw.lower() in doc_text]
                passed = len(matched) >= 1
                message = f"文档{'包含' if passed else '未包含'}指定关键词（匹配{len(matched)}/{len(keywords)}个：{', '.join(matched) if matched else '无'}）"

            elif check_type == "excel_has_merged_cells":
                passed = self.parser.excel_has_merged_cells()
                message = f"{'存在' if passed else '不存在'}合并单元格"

            elif check_type == "excel_data_centered":
                passed = self.parser.excel_data_centered()
                message = f"数据区域{'已' if passed else '未'}设置居中对齐"

            elif check_type == "excel_column_width_set":
                min_width = params.get("min_width", 0)
                passed = self.parser.excel_column_width_set(min_width=min_width)
                if min_width > 0:
                    message = f"列宽{'已' if passed else '未'}合理设置（至少一列≥{min_width}字符）"
                else:
                    message = f"列宽{'已' if passed else '未'}自定义设置"

            elif check_type == "excel_row_height_set":
                passed = self.parser.excel_row_height_set()
                message = f"行高{'已' if passed else '未'}自定义设置"

            elif check_type == "excel_cell_font_check":
                row = params.get("row", 1)
                col = params.get("col", 1)
                font_name = params.get("font_name")
                font_size = params.get("font_size")
                bold = params.get("bold")
                passed = self.parser.excel_cell_font_check(row=row, col=col, font_name=font_name, font_size=font_size, bold=bold)
                desc_parts = []
                if font_name: desc_parts.append(f"字体={font_name}")
                if font_size: desc_parts.append(f"字号={font_size}")
                if bold is not None: desc_parts.append(f"加粗={bold}")
                desc_str = "、".join(desc_parts)
                col_letter = chr(64 + col) if col <= 26 else str(col)
                message = f"{col_letter}{row}{'已' if passed else '未'}正确设置（{desc_str}）"

            elif check_type == "excel_cell_alignment_check":
                row_start = params.get("row_start", 1)
                row_end = params.get("row_end", 10)
                col = params.get("col", 1)
                align = params.get("align", "left")
                align_names = {"left": "左对齐", "center": "居中", "right": "右对齐"}
                col_letter = chr(64 + col) if col <= 26 else str(col)
                passed = self.parser.excel_cell_alignment_check(row_start=row_start, row_end=row_end, col=col, align=align)
                message = f"{col_letter}{row_start}:{col_letter}{row_end}{'已' if passed else '未'}设置{align_names.get(align, align)}"

            elif check_type == "excel_has_border":
                passed = self.parser.excel_has_border()
                message = f"数据区域{'已' if passed else '未'}设置边框"

            elif check_type == "excel_has_validation":
                col = params.get("col")
                validation_type = params.get("validation_type")
                formula1 = params.get("formula1")
                min_val = params.get("min_val")
                max_val = params.get("max_val")
                type_names = {"whole": "整数", "decimal": "小数", "list": "序列", "textLength": "文本长度", "date": "日期", "custom": "自定义"}
                passed = self.parser.excel_has_validation(col=col, validation_type=validation_type, formula1=formula1, min_val=min_val, max_val=max_val)
                desc = type_names.get(validation_type, validation_type) if validation_type else "数据有效性"
                if col:
                    col_letter = chr(64 + col) if col <= 26 else str(col)
                    desc = f"{col_letter}列{desc}"
                message = f"{'已' if passed else '未'}设置{desc}验证"

            else:
                message = f"未知的检查类型: {check_type}"
        
        except Exception as e:
            passed = False
            message = f"检查出错: {str(e)}"
        
        return {
            "id": item["id"],
            "name": item["name"],
            "category": item["category"],
            "score": item["score"],
            "passed": passed,
            "message": message,
            "description": item["description"],
        }
    
    def _get_category_summary(self):
        """获取分类汇总"""
        categories = {}
        for item in self.results:
            cat = item["category"]
            if cat not in categories:
                categories[cat] = {"total": 0, "passed": 0, "score": 0}
            categories[cat]["total"] += 1
            categories[cat]["score"] += item["score"]
            if item["passed"]:
                categories[cat]["passed"] += 1
        return categories
