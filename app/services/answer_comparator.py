# -*- coding: utf-8 -*-
"""
Excel答案对比评分引擎

通过ZIP读取XML实现Excel文件对比，兼容WPS。
核心类：ExcelAnswerComparator
"""

import zipfile
import re
import os


# ============================================================
# 任务点对比定义
# ============================================================

TASK_CHECK_POINTS = {
    # ============================================================
    # 任务5.4.1 排序与分类汇总
    # ============================================================
    "task_5_4_1": {
        "task_name": "排序与分类汇总",
        "check_points": [
            {
                "id": "5_4_1_01",
                "name": "(1) 排序表1：城区升序+建造时间降序",
                "score": 25,
                "compare": "structure_check",
                "category": "数据排序能力",
                "params": {
                    "sheet": "排序表1",
                    "check_type": "sort",
                    "sort_keys": ["城区:asc", "建造时间:desc"]
                }
            },
            {
                "id": "5_4_1_02",
                "name": "(2) 排序表2：建造时间降序+单价升序",
                "score": 25,
                "compare": "structure_check",
                "category": "数据排序能力",
                "params": {
                    "sheet": "排序表2",
                    "check_type": "sort",
                    "sort_keys": ["建造时间:desc", "单价:asc"]
                }
            },
            {
                "id": "5_4_1_03",
                "name": "(3) 分类汇总表1：按城区统计单价平均值",
                "score": 25,
                "compare": "structure_check",
                "category": "数据汇总能力",
                "params": {
                    "sheet": "分类汇总表1",
                    "check_type": "subtotal",
                    "group_field": "城区",
                    "summary_field": "单价",
                    "summary_method": "average"
                }
            },
            {
                "id": "5_4_1_04",
                "name": "(4) 分类汇总表2：按朝向统计平均值",
                "score": 25,
                "compare": "structure_check",
                "category": "数据汇总能力",
                "params": {
                    "sheet": "分类汇总表2",
                    "check_type": "subtotal",
                    "group_field": "朝向",
                    "summary_fields": ["单价", "总价", "面积"],
                    "summary_method": "average"
                }
            }
        ]
    },

    # ============================================================
    # 任务5.4.2 数据筛选
    # ============================================================
    "task_5_4_2": {
        "task_name": "数据筛选",
        "check_points": [
            {
                "id": "5_4_2_01",
                "name": "(1) 自动筛选：建造时间为2020年和2021年",
                "score": 25,
                "compare": "structure_check",
                "category": "数据筛选能力",
                "params": {
                    "sheet": "自动筛选1",
                    "check_type": "auto_filter",
                    "filter_field": "建造时间",
                    "filter_values": ["2020", "2021"]
                }
            },
            {
                "id": "5_4_2_02",
                "name": "(2) 自动筛选：总价低于平均值",
                "score": 25,
                "compare": "structure_check",
                "category": "数据筛选能力",
                "params": {
                    "sheet": "自动筛选2",
                    "check_type": "auto_filter",
                    "filter_field": "总价",
                    "filter_condition": "below_average"
                }
            },
            {
                "id": "5_4_2_03",
                "name": "(3) 高级筛选：朝向=南北 且 格局=4室2厅2卫",
                "score": 25,
                "compare": "structure_check",
                "category": "数据筛选能力",
                "params": {
                    "sheet": "高级筛选1",
                    "check_type": "advanced_filter",
                    "filter_conditions": [
                        {"field": "朝向", "value": "南北"},
                        {"field": "格局", "value": "4室2厅2卫"}
                    ]
                }
            },
            {
                "id": "5_4_2_04",
                "name": "(4) 高级筛选：单价<5000 或 总价<50万",
                "score": 25,
                "compare": "structure_check",
                "category": "数据筛选能力",
                "params": {
                    "sheet": "高级筛选2",
                    "check_type": "advanced_filter",
                    "filter_conditions": [
                        {"field": "单价", "operator": "<", "value": "5000"},
                        {"field": "总价", "operator": "<", "value": "50"}
                    ],
                    "filter_logic": "or"
                }
            }
        ]
    },

    # ============================================================
    # 任务5.4.3 数据透视表
    # ============================================================
    "task_5_4_3": {
        "task_name": "数据透视表",
        "check_points": [
            {
                "id": "5_4_3_01",
                "name": "(1) 透视表行字段：城区",
                "score": 25,
                "compare": "structure_check",
                "category": "数据透视能力",
                "params": {
                    "sheet": "透视表",
                    "check_type": "pivot_table_row",
                    "row_field": "城区"
                }
            },
            {
                "id": "5_4_3_02",
                "name": "(2) 透视表列字段：楼层位置",
                "score": 25,
                "compare": "structure_check",
                "category": "数据透视能力",
                "params": {
                    "sheet": "透视表",
                    "check_type": "pivot_table_col",
                    "column_field": "楼层位置"
                }
            },
            {
                "id": "5_4_3_03",
                "name": "(3) 透视表数据字段：单价（元/㎡）",
                "score": 25,
                "compare": "structure_check",
                "category": "数据透视能力",
                "params": {
                    "sheet": "透视表",
                    "check_type": "pivot_table_data",
                    "data_field": "单价（元/㎡）"
                }
            },
            {
                "id": "5_4_3_04",
                "name": "(4) 透视表汇总方式：平均值",
                "score": 25,
                "compare": "structure_check",
                "category": "数据透视能力",
                "params": {
                    "sheet": "透视表",
                    "check_type": "pivot_table_summary",
                    "summary_method": "平均值"
                }
            }
        ]
    },

    # ============================================================
    # 任务5.4.4 图表可视化
    # ============================================================
    "task_5_4_4": {
        "task_name": "创建图表",
        "check_points": [
            {
                "id": "5_4_4_02",
                "name": "步骤1：修改数据源，标题'各城区各楼层类型房价均值柱形图'",
                "score": 16,
                "compare": "chart_title",
                "category": "图表创建能力",
                "params": {
                    "chart_index": 0,
                    "expected_title": "各城区各楼层类型房价均值柱形图"
                }
            },
            {
                "id": "5_4_4_03",
                "name": "步骤2：添加图例",
                "score": 16,
                "compare": "chart_has_legend",
                "category": "图表美化能力",
                "params": {
                    "chart_index": 0,
                    "expected_title": "各城区各楼层类型房价均值柱形图"
                }
            },
            {
                "id": "5_4_4_04",
                "name": "步骤3：添加数据标签",
                "score": 16,
                "compare": "chart_has_data_labels",
                "category": "图表美化能力",
                "params": {
                    "chart_index": 0,
                    "expected_title": "各城区各楼层类型房价均值柱形图"
                }
            },
            {
                "id": "5_4_4_05",
                "name": "步骤3：删除网格线",
                "score": 16,
                "compare": "chart_no_gridlines",
                "category": "图表美化能力",
                "params": {
                    "chart_index": 0,
                    "expected_title": "各城区各楼层类型房价均值柱形图"
                }
            },
            {
                "id": "5_4_4_06",
                "name": "步骤4：设置坐标轴填充颜色（矢车菊蓝/金色）",
                "score": 12,
                "compare": "chart_axis_color",
                "category": "图表美化能力",
                "params": {
                    "chart_index": 0,
                    "expected_title": "各城区各楼层类型房价均值柱形图"
                }
            },
            {
                "id": "5_4_4_07",
                "name": "步骤5：设置绘图区填充颜色（白烟）",
                "score": 12,
                "compare": "chart_plot_area_color",
                "category": "图表美化能力",
                "params": {
                    "chart_index": 0,
                    "expected_title": "各城区各楼层类型房价均值柱形图"
                }
            },
            {
                "id": "5_4_4_08",
                "name": "步骤6：图表放置在A10:F35区域",
                "score": 12,
                "compare": "chart_position",
                "category": "图表布局能力",
                "params": {
                    "chart_index": 0,
                    "expected_title": "各城区各楼层类型房价均值柱形图",
                    "from_col": 0, "from_row": 9,
                    "to_col": 5, "to_row": 34
                }
            },
        ]
    },

    # ============================================================
    # 任务5.4.5 页面设置
    # ============================================================
    "task_5_4_5": {
        "task_name": "页面设置",
        "check_points": [
            {
                "id": "5_4_5_01",
                "name": "(1) 打印区域A1:P28",
                "score": 20,
                "compare": "structure_check",
                "category": "页面设置能力",
                "params": {
                    "sheet": "房源信息统计表",
                    "check_type": "print_area",
                    "print_area": "A1:P28"
                }
            },
            {
                "id": "5_4_5_02",
                "name": "(2) 页边距居中方式水平，上下边距",
                "score": 20,
                "compare": "structure_check",
                "category": "页面设置能力",
                "params": {
                    "sheet": "房源信息统计表",
                    "check_type": "page_margins",
                    "horizontal_centered": True
                }
            },
            {
                "id": "5_4_5_03",
                "name": "(3) 顶端标题行1-3行",
                "score": 20,
                "compare": "structure_check",
                "category": "页面设置能力",
                "params": {
                    "sheet": "房源信息统计表",
                    "check_type": "print_title_rows",
                    "title_rows": "1:3"
                }
            },
            {
                "id": "5_4_5_04",
                "name": "(4) 横向，所有列打印在一页",
                "score": 20,
                "compare": "structure_check",
                "category": "页面设置能力",
                "params": {
                    "sheet": "房源信息统计表",
                    "check_type": "page_orientation",
                    "orientation": "landscape",
                    "fit_to_page": True
                }
            },
            {
                "id": "5_4_5_05",
                "name": "(5) 打印预览",
                "score": 20,
                "compare": "structure_check",
                "category": "页面设置能力",
                "params": {
                    "sheet": "房源信息统计表",
                    "check_type": "print_preview",
                    "description": "打印预览功能检查"
                }
            }
        ]
    },

    # ============================================================
    # 任务5.3.2 常用函数统计
    # ============================================================
    "task_5_3_2": {
        "task_name": "常用函数统计",
        "check_points": [
            {
                "id": "5_3_2_01",
                "name": "(1) COUNTIF函数统计各城区房源数",
                "score": 12,
                "compare": "formula_exists",
                "category": "条件统计能力",
                "params": {
                    "formula_keyword": "COUNTIF",
                    "description": "在S3:S7区域使用COUNTIF函数计算各个城区的房源数"
                }
            },
            {
                "id": "5_3_2_02",
                "name": "(2) AVERAGEIF函数计算单价均值",
                "score": 12,
                "compare": "formula_exists",
                "category": "条件统计能力",
                "params": {
                    "formula_keyword": "AVERAGEIF",
                    "description": "在T3:T7区域使用AVERAGEIF函数计算各个城区房源的单价均值"
                }
            },
            {
                "id": "5_3_2_03",
                "name": "(3) SUMIF函数计算总价和",
                "score": 12,
                "compare": "formula_exists",
                "category": "条件统计能力",
                "params": {
                    "formula_keyword": "SUMIF",
                    "description": "在U3:U7区域使用SUMIF函数计算各个城区房源的总价和"
                }
            },
            {
                "id": "5_3_2_04",
                "name": "(4) IF函数进行面积分类",
                "score": 12,
                "compare": "formula_exists",
                "category": "逻辑判断能力",
                "params": {
                    "formula_keyword": "IF",
                    "description": "在K4:K28区域使用IF函数对面积进行分类（小户型/中户型/大户型/超大户型）"
                }
            },
            {
                "id": "5_3_2_05",
                "name": "(5) IF+AND函数备注楼层和面积类型",
                "score": 11,
                "compare": "formula_exists",
                "category": "逻辑判断能力",
                "params": {
                    "formula_keyword": "AND",
                    "description": "在M4:M28区域使用IF和AND函数，楼层类型为中层且面积类型为中户型则备注A"
                }
            },
            {
                "id": "5_3_2_06",
                "name": "(6) RANK.EQ函数排名",
                "score": 11,
                "compare": "formula_exists",
                "category": "排序排名能力",
                "params": {
                    "formula_keyword": "RANK",
                    "description": "在N4:N28区域使用RANK.EQ函数依据单价按降序排名"
                }
            },
            {
                "id": "5_3_2_07",
                "name": "(7) MIN/MAX函数统计最值",
                "score": 10,
                "compare": "formula_exists",
                "category": "数据统计能力",
                "params": {
                    "formula_keyword": "MIN",
                    "description": "在S23:S24区域分别使用MIN和MAX函数统计单价最小值和面积最大值"
                }
            },
            {
                "id": "5_3_2_08",
                "name": "(8) SUBTOTAL函数统计",
                "score": 10,
                "compare": "formula_exists",
                "category": "数据统计能力",
                "params": {
                    "formula_keyword": "SUBTOTAL",
                    "description": "在T23:T24区域使用SUBTOTAL函数统计单价最小值和面积最大值"
                }
            },
            {
                "id": "5_3_2_09",
                "name": "(9) VLOOKUP函数匹配销售经理",
                "score": 10,
                "compare": "formula_exists",
                "category": "数据查找能力",
                "params": {
                    "formula_keyword": "VLOOKUP",
                    "description": "在O4:P28区域使用VLOOKUP函数匹配各城区销售经理和联系方式"
                }
            },
        ]
    }
}


class ExcelAnswerComparator:
    """Excel答案对比评分引擎"""

    def __init__(self, answer_file, student_file):
        """
        构造函数

        Args:
            answer_file: 答案文件路径
            student_file: 学生文件路径
        """
        self.answer_file = answer_file
        self.student_file = student_file

    def _read_xml_from_zip(self, file_path, xml_path):
        """
        从ZIP包中读取指定XML文件的内容

        Args:
            file_path: Excel文件路径
            xml_path: ZIP包内XML文件的相对路径

        Returns:
            str: XML文件内容，读取失败返回None
        """
        try:
            if not os.path.exists(file_path):
                return None
            with zipfile.ZipFile(file_path, 'r') as zf:
                if xml_path in zf.namelist():
                    return zf.read(xml_path).decode('utf-8')
                return None
        except Exception:
            return None

    def _list_xl_files(self, file_path, pattern):
        """
        列出ZIP包xl/目录下匹配pattern的文件名列表

        Args:
            file_path: Excel文件路径
            pattern: 文件名匹配模式（正则）

        Returns:
            list: 匹配的文件路径列表
        """
        try:
            if not os.path.exists(file_path):
                return []
            with zipfile.ZipFile(file_path, 'r') as zf:
                names = zf.namelist()
                result = []
                for name in names:
                    if re.search(pattern, name):
                        result.append(name)
                return sorted(result)
        except Exception:
            return []

    def _normalize_whitespace(self, text):
        """去除多余空白，用于对比时忽略空白差异"""
        if text is None:
            return ''
        return re.sub(r'\s+', '', str(text))

    # ================================================================
    # 1. 工作表结构对比
    # ================================================================

    def _compare_sheet_structure(self):
        """
        对比工作表结构：工作表名称和行列数

        Returns:
            list: 对比结果详情列表
        """
        details = []

        try:
            answer_workbook = self._read_xml_from_zip(self.answer_file, 'xl/workbook.xml')
            student_workbook = self._read_xml_from_zip(self.student_file, 'xl/workbook.xml')

            if answer_workbook is None or student_workbook is None:
                details.append({
                    "name": "工作表结构",
                    "passed": False,
                    "score": 0,
                    "message": "无法读取workbook.xml文件",
                    "category": "工作表结构"
                })
                return details

            # 提取工作表名称
            answer_sheets = re.findall(r'<sheet\s[^>]*name="([^"]+)"', answer_workbook)
            student_sheets = re.findall(r'<sheet\s[^>]*name="([^"]+)"', student_workbook)

            # 对比工作表名称
            if answer_sheets == student_sheets:
                details.append({
                    "name": "工作表名称",
                    "passed": True,
                    "score": 5,
                    "message": "工作表名称一致",
                    "category": "工作表结构"
                })
            else:
                details.append({
                    "name": "工作表名称",
                    "passed": False,
                    "score": 0,
                    "message": f"工作表名称不一致，答案: {answer_sheets}，学生: {student_sheets}",
                    "category": "工作表结构"
                })

            # 对比每个工作表的行列数
            answer_sheet_files = self._list_xl_files(self.answer_file, r'xl/worksheets/sheet\d+\.xml')
            student_sheet_files = self._list_xl_files(self.student_file, r'xl/worksheets/sheet\d+\.xml')

            sheet_count_match = len(answer_sheet_files) == len(student_sheet_files)

            if not sheet_count_match:
                details.append({
                    "name": "工作表数量",
                    "passed": False,
                    "score": 0,
                    "message": f"工作表数量不一致，答案: {len(answer_sheet_files)}个，学生: {len(student_sheet_files)}个",
                    "category": "工作表结构"
                })
            else:
                details.append({
                    "name": "工作表数量",
                    "passed": True,
                    "score": 5,
                    "message": f"工作表数量一致，均为{len(answer_sheet_files)}个",
                    "category": "工作表结构"
                })

                # 逐工作表对比行列数
                all_dimensions_match = True
                dim_messages = []
                for i in range(min(len(answer_sheet_files), len(student_sheet_files))):
                    answer_sheet_xml = self._read_xml_from_zip(self.answer_file, answer_sheet_files[i])
                    student_sheet_xml = self._read_xml_from_zip(self.student_file, student_sheet_files[i])

                    answer_dim = self._extract_sheet_dimension(answer_sheet_xml)
                    student_dim = self._extract_sheet_dimension(student_sheet_xml)

                    if answer_dim != student_dim:
                        all_dimensions_match = False
                        sheet_name = answer_sheets[i] if i < len(answer_sheets) else f"Sheet{i+1}"
                        dim_messages.append(f"{sheet_name}: 答案{answer_dim} vs 学生{student_dim}")

                if all_dimensions_match:
                    details.append({
                        "name": "工作表行列数",
                        "passed": True,
                        "score": 5,
                        "message": "所有工作表行列数一致",
                        "category": "工作表结构"
                    })
                else:
                    details.append({
                        "name": "工作表行列数",
                        "passed": False,
                        "score": 0,
                        "message": f"行列数不一致: {'; '.join(dim_messages)}",
                        "category": "工作表结构"
                    })

        except Exception as e:
            details.append({
                "name": "工作表结构",
                "passed": False,
                "score": 0,
                "message": f"工作表结构对比异常: {str(e)}",
                "category": "工作表结构"
            })

        return details

    def _extract_sheet_dimension(self, sheet_xml):
        """
        从sheet XML中提取行列范围

        Args:
            sheet_xml: sheet的XML内容

        Returns:
            str: 维度字符串，如 "A1:D10"，无dimension标签返回空字符串
        """
        if sheet_xml is None:
            return ''
        match = re.search(r'<dimension\s[^>]*ref="([^"]+)"', sheet_xml)
        if match:
            return match.group(1)
        return ''

    # ================================================================
    # 2. 单元格数据对比
    # ================================================================

    def _compare_cell_data(self):
        """
        逐单元格对比值（数值、文本、公式），忽略格式差异

        Returns:
            list: 对比结果详情列表
        """
        details = []

        try:
            answer_sheet_files = self._list_xl_files(self.answer_file, r'xl/worksheets/sheet\d+\.xml')
            student_sheet_files = self._list_xl_files(self.student_file, r'xl/worksheets/sheet\d+\.xml')

            if not answer_sheet_files or not student_sheet_files:
                details.append({
                    "name": "单元格数据",
                    "passed": False,
                    "score": 0,
                    "message": "无法读取工作表文件",
                    "category": "单元格数据"
                })
                return details

            # 读取共享字符串表
            answer_shared = self._read_xml_from_zip(self.answer_file, 'xl/sharedStrings.xml')
            student_shared = self._read_xml_from_zip(self.student_file, 'xl/sharedStrings.xml')

            answer_strings = self._parse_shared_strings(answer_shared) if answer_shared else []
            student_strings = self._parse_shared_strings(student_shared) if student_shared else []

            total_cells = 0
            mismatch_cells = 0
            mismatch_details = []

            sheet_count = min(len(answer_sheet_files), len(student_sheet_files))

            for i in range(sheet_count):
                answer_xml = self._read_xml_from_zip(self.answer_file, answer_sheet_files[i])
                student_xml = self._read_xml_from_zip(self.student_file, student_sheet_files[i])

                if answer_xml is None or student_xml is None:
                    continue

                answer_cells = self._extract_all_cells(answer_xml, answer_strings)
                student_cells = self._extract_all_cells(student_xml, student_strings)

                all_refs = sorted(set(list(answer_cells.keys()) + list(student_cells.keys())))

                for ref in all_refs:
                    total_cells += 1
                    a_val = answer_cells.get(ref, '')
                    s_val = student_cells.get(ref, '')

                    if self._normalize_whitespace(a_val) != self._normalize_whitespace(s_val):
                        mismatch_cells += 1
                        if len(mismatch_details) < 10:
                            mismatch_details.append(
                                f"{ref}: 答案[{a_val}] vs 学生[{s_val}]"
                            )

            if total_cells == 0:
                details.append({
                    "name": "单元格数据",
                    "passed": True,
                    "score": 19,
                    "message": "工作表为空，无单元格数据需要对比",
                    "category": "单元格数据"
                })
            elif mismatch_cells == 0:
                details.append({
                    "name": "单元格数据",
                    "passed": True,
                    "score": 19,
                    "message": f"所有{total_cells}个单元格数据一致",
                    "category": "单元格数据"
                })
            else:
                match_rate = (total_cells - mismatch_cells) / total_cells
                score = int(19 * match_rate)
                details.append({
                    "name": "单元格数据",
                    "passed": False,
                    "score": score,
                    "message": f"共{total_cells}个单元格，{mismatch_cells}个不一致"
                               f"(匹配率{match_rate:.1%})，差异示例: {'; '.join(mismatch_details)}",
                    "category": "单元格数据"
                })

        except Exception as e:
            details.append({
                "name": "单元格数据",
                "passed": False,
                "score": 0,
                "message": f"单元格数据对比异常: {str(e)}",
                "category": "单元格数据"
            })

        return details

    def _parse_shared_strings(self, shared_xml):
        """
        解析共享字符串表

        Args:
            shared_xml: sharedStrings.xml内容

        Returns:
            list: 共享字符串列表
        """
        strings = []
        if shared_xml is None:
            return strings
        # 尝试两种格式：<t> 和 <a:t>
        items = re.findall(r'<si>(.*?)</si>', shared_xml, re.DOTALL)
        for item in items:
            # 先尝试 <t> 标签
            texts = re.findall(r'<t>([^<]*)</t>', item)
            if not texts:
                # 再尝试 <a:t> 标签
                texts = re.findall(r'<a:t>([^<]*)</a:t>', item)
            strings.append(''.join(texts))
        return strings

    def _extract_all_cells(self, sheet_xml, shared_strings):
        """
        提取工作表中所有单元格的值

        Args:
            sheet_xml: sheet的XML内容
            shared_strings: 共享字符串列表

        Returns:
            dict: {单元格引用: 值}，如 {"A1": "hello", "B2": "100"}
        """
        cells = {}
        if sheet_xml is None:
            return cells

        rows = re.findall(r'<row[^>]*>(.*?)</row>', sheet_xml, re.DOTALL)
        for row in rows:
            cell_matches = re.findall(r'<c\s[^>]*>(.*?)</c>', row, re.DOTALL)
            # 同时匹配自闭合和有内容的c标签
            cell_matches_full = re.findall(
                r'<c\s+r="([^"]+)"[^>]*(?:/>|>(.*?)</c>)',
                row, re.DOTALL
            )
            for match in cell_matches_full:
                ref = match[0]
                inner = match[1] if match[1] else ''

                # 获取单元格类型
                c_tag = re.search(rf'<c\s+r="{re.escape(ref)}"[^>]*>', row)
                cell_type = ''
                if c_tag:
                    type_match = re.search(r't="([^"]+)"', c_tag.group(0))
                    if type_match:
                        cell_type = type_match.group(1)

                value = self._get_cell_value(inner, cell_type, shared_strings)
                cells[ref] = value

        return cells

    def _get_cell_value(self, inner_xml, cell_type, shared_strings):
        """
        获取单元格的值

        Args:
            inner_xml: 单元格内部XML
            cell_type: 单元格类型（s=共享字符串, n=数字, b=布尔等）
            shared_strings: 共享字符串列表

        Returns:
            str: 单元格值
        """
        if not inner_xml:
            return ''

        # 提取v标签的值
        v_match = re.search(r'<v>([^<]*)</v>', inner_xml)
        if not v_match:
            # 检查是否有f标签（公式）但没有v标签
            f_match = re.search(r'<f>([^<]*)</f>', inner_xml)
            if f_match:
                return f'[公式:{f_match.group(1)}]'
            return ''

        v_value = v_match.group(1)

        if cell_type == 's':
            # 共享字符串索引
            try:
                idx = int(v_value)
                if 0 <= idx < len(shared_strings):
                    return shared_strings[idx]
            except (ValueError, IndexError):
                pass
            return v_value
        elif cell_type == 'b':
            return 'TRUE' if v_value == '1' else 'FALSE'
        elif cell_type == 'str':
            return v_value
        else:
            # 数值或其他类型
            return v_value

    # ================================================================
    # 3. 图表对比
    # ================================================================

    def _compare_charts(self):
        """
        对比图表：标题、数据系列、数据标签、网格线、图例、坐标轴格式、绘图区格式

        Returns:
            list: 对比结果详情列表
        """
        details = []

        try:
            answer_charts = self._list_xl_files(self.answer_file, r'xl/charts/chart\d+\.xml')
            student_charts = self._list_xl_files(self.student_file, r'xl/charts/chart\d+\.xml')

            if not answer_charts and not student_charts:
                # 两个文件都没有图表，跳过
                return details

            if not answer_charts:
                details.append({
                    "name": "图表数量",
                    "passed": False,
                    "score": 0,
                    "message": "答案文件没有图表",
                    "category": "图表"
                })
                return details

            if not student_charts:
                details.append({
                    "name": "图表数量",
                    "passed": False,
                    "score": 0,
                    "message": "学生文件没有图表",
                    "category": "图表"
                })
                return details

            # 对比图表数量
            if len(answer_charts) == len(student_charts):
                details.append({
                    "name": "图表数量",
                    "passed": True,
                    "score": 3,
                    "message": f"图表数量一致，均为{len(answer_charts)}个",
                    "category": "图表"
                })
            else:
                details.append({
                    "name": "图表数量",
                    "passed": False,
                    "score": 0,
                    "message": f"图表数量不一致，答案: {len(answer_charts)}个，学生: {len(student_charts)}个",
                    "category": "图表"
                })

            # 逐图表对比
            chart_count = min(len(answer_charts), len(student_charts))
            for i in range(chart_count):
                answer_xml = self._read_xml_from_zip(self.answer_file, answer_charts[i])
                student_xml = self._read_xml_from_zip(self.student_file, student_charts[i])

                if answer_xml is None or student_xml is None:
                    continue

                chart_label = f"图表{i+1}"

                # a) 图表标题
                details.extend(self._compare_chart_title(answer_xml, student_xml, chart_label))

                # b) 数据系列数量
                details.extend(self._compare_series_count(answer_xml, student_xml, chart_label))

                # c) 数据系列名称
                details.extend(self._compare_series_names(answer_xml, student_xml, chart_label))

                # d) 数据标签
                details.extend(self._compare_data_labels(answer_xml, student_xml, chart_label))

                # e) 网格线
                details.extend(self._compare_gridlines(answer_xml, student_xml, chart_label))

                # f) 图例
                details.extend(self._compare_legend(answer_xml, student_xml, chart_label))

                # g) 坐标轴格式
                details.extend(self._compare_axis_format(answer_xml, student_xml, chart_label))

                # h) 绘图区和图表区格式
                details.extend(self._compare_area_format(answer_xml, student_xml, chart_label))

        except Exception as e:
            details.append({
                "name": "图表对比",
                "passed": False,
                "score": 0,
                "message": f"图表对比异常: {str(e)}",
                "category": "图表"
            })

        return details

    def _extract_chart_title(self, chart_xml):
        """
        提取图表标题文本

        Args:
            chart_xml: chart的XML内容

        Returns:
            str: 标题文本
        """
        match = re.search(
            r'<c:title>.*?<c:tx>(.*?)</c:tx>.*?</c:title>',
            chart_xml, re.DOTALL
        )
        if match:
            texts = re.findall(r'<a:t>([^<]+)</a:t>', match.group(1))
            return ''.join(texts)
        return ''

    def _compare_chart_title(self, answer_xml, student_xml, chart_label):
        """对比图表标题"""
        answer_title = self._extract_chart_title(answer_xml)
        student_title = self._extract_chart_title(student_xml)

        if self._normalize_whitespace(answer_title) == self._normalize_whitespace(student_title):
            return [{
                "name": f"{chart_label}标题",
                "passed": True,
                "score": 3,
                "message": f"图表标题一致: {answer_title}",
                "category": "图表"
            }]
        else:
            return [{
                "name": f"{chart_label}标题",
                "passed": False,
                "score": 0,
                "message": f"图表标题不一致，答案: [{answer_title}]，学生: [{student_title}]",
                "category": "图表"
            }]

    def _count_series(self, chart_xml):
        """
        统计数据系列数量

        Args:
            chart_xml: chart的XML内容

        Returns:
            int: 数据系列数量
        """
        return len(re.findall(r'<c:ser>', chart_xml))

    def _compare_series_count(self, answer_xml, student_xml, chart_label):
        """对比数据系列数量"""
        answer_count = self._count_series(answer_xml)
        student_count = self._count_series(student_xml)

        if answer_count == student_count:
            return [{
                "name": f"{chart_label}数据系列数量",
                "passed": True,
                "score": 3,
                "message": f"数据系列数量一致，均为{answer_count}个",
                "category": "图表"
            }]
        else:
            return [{
                "name": f"{chart_label}数据系列数量",
                "passed": False,
                "score": 0,
                "message": f"数据系列数量不一致，答案: {answer_count}个，学生: {student_count}个",
                "category": "图表"
            }]

    def _extract_series_names(self, chart_xml):
        """
        提取每个数据系列的名称

        Args:
            chart_xml: chart的XML内容

        Returns:
            list: 系列名称列表
        """
        series = re.findall(r'<c:ser>(.*?)</c:ser>', chart_xml, re.DOTALL)
        names = []
        for s in series:
            tx = re.search(r'<c:tx>(.*?)</c:tx>', s, re.DOTALL)
            if tx:
                texts = re.findall(r'<a:t>([^<]+)</a:t>', tx.group(1))
                names.append(''.join(texts))
        return names

    def _compare_series_names(self, answer_xml, student_xml, chart_label):
        """对比数据系列名称"""
        answer_names = self._extract_series_names(answer_xml)
        student_names = self._extract_series_names(student_xml)

        if answer_names == student_names:
            return [{
                "name": f"{chart_label}数据系列名称",
                "passed": True,
                "score": 3,
                "message": f"数据系列名称一致: {answer_names}",
                "category": "图表"
            }]
        else:
            return [{
                "name": f"{chart_label}数据系列名称",
                "passed": False,
                "score": 0,
                "message": f"数据系列名称不一致，答案: {answer_names}，学生: {student_names}",
                "category": "图表"
            }]

    def _has_data_labels(self, chart_xml):
        """
        检查是否有数据标签且未被删除

        Args:
            chart_xml: chart的XML内容

        Returns:
            bool: 是否有数据标签
        """
        if '<c:dLbls>' not in chart_xml:
            return False
        dLbls = re.search(r'<c:dLbls>(.*?)</c:dLbls>', chart_xml, re.DOTALL)
        if dLbls and '<c:delete val="1"/>' in dLbls.group(0):
            return False
        return True

    def _compare_data_labels(self, answer_xml, student_xml, chart_label):
        """对比数据标签"""
        answer_has = self._has_data_labels(answer_xml)
        student_has = self._has_data_labels(student_xml)

        if answer_has == student_has:
            status = "均有" if answer_has else "均无"
            return [{
                "name": f"{chart_label}数据标签",
                "passed": True,
                "score": 3,
                "message": f"数据标签状态一致: {status}数据标签",
                "category": "图表"
            }]
        else:
            return [{
                "name": f"{chart_label}数据标签",
                "passed": False,
                "score": 0,
                "message": f"数据标签状态不一致，答案: {'有' if answer_has else '无'}数据标签，"
                           f"学生: {'有' if student_has else '无'}数据标签",
                "category": "图表"
            }]

    def _count_gridlines(self, chart_xml):
        """
        统计主网格线数量

        Args:
            chart_xml: chart的XML内容

        Returns:
            int: 网格线数量
        """
        return len(re.findall(r'<c:majorGridlines', chart_xml))

    def _compare_gridlines(self, answer_xml, student_xml, chart_label):
        """对比网格线"""
        answer_count = self._count_gridlines(answer_xml)
        student_count = self._count_gridlines(student_xml)

        if answer_count == student_count:
            return [{
                "name": f"{chart_label}网格线",
                "passed": True,
                "score": 2,
                "message": f"网格线数量一致，均为{answer_count}条",
                "category": "图表"
            }]
        else:
            return [{
                "name": f"{chart_label}网格线",
                "passed": False,
                "score": 0,
                "message": f"网格线数量不一致，答案: {answer_count}条，学生: {student_count}条",
                "category": "图表"
            }]

    def _has_legend(self, chart_xml):
        """
        检查是否有图例

        Args:
            chart_xml: chart的XML内容

        Returns:
            bool: 是否有图例
        """
        return '<c:legend>' in chart_xml

    def _compare_legend(self, answer_xml, student_xml, chart_label):
        """对比图例"""
        answer_has = self._has_legend(answer_xml)
        student_has = self._has_legend(student_xml)

        if answer_has == student_has:
            status = "均有" if answer_has else "均无"
            return [{
                "name": f"{chart_label}图例",
                "passed": True,
                "score": 2,
                "message": f"图例状态一致: {status}图例",
                "category": "图表"
            }]
        else:
            return [{
                "name": f"{chart_label}图例",
                "passed": False,
                "score": 0,
                "message": f"图例状态不一致，答案: {'有' if answer_has else '无'}图例，"
                           f"学生: {'有' if student_has else '无'}图例",
                "category": "图表"
            }]

    def _compare_axis_format(self, answer_xml, student_xml, chart_label):
        """
        对比坐标轴格式（catAx和valAx的spPr）

        Args:
            answer_xml: 答案chart XML
            student_xml: 学生chart XML
            chart_label: 图表标签

        Returns:
            list: 对比结果
        """
        def get_axis_spPr(xml, axis_type):
            pattern = f'<c:{axis_type}>(.*?)</c:{axis_type}>'
            match = re.search(pattern, xml, re.DOTALL)
            if match:
                spPr = re.search(r'<c:spPr>(.*?)</c:spPr>', match.group(1), re.DOTALL)
                if spPr:
                    return spPr.group(0)
            return ''

        answer_cat = get_axis_spPr(answer_xml, 'catAx')
        student_cat = get_axis_spPr(student_xml, 'catAx')
        answer_val = get_axis_spPr(answer_xml, 'valAx')
        student_val = get_axis_spPr(student_xml, 'valAx')

        cat_match = self._normalize_whitespace(answer_cat) == self._normalize_whitespace(student_cat)
        val_match = self._normalize_whitespace(answer_val) == self._normalize_whitespace(student_val)

        results = []

        if cat_match:
            results.append({
                "name": f"{chart_label}分类轴格式",
                "passed": True,
                "score": 2,
                "message": "分类轴格式一致",
                "category": "图表"
            })
        else:
            results.append({
                "name": f"{chart_label}分类轴格式",
                "passed": False,
                "score": 0,
                "message": "分类轴格式不一致",
                "category": "图表"
            })

        if val_match:
            results.append({
                "name": f"{chart_label}数值轴格式",
                "passed": True,
                "score": 2,
                "message": "数值轴格式一致",
                "category": "图表"
            })
        else:
            results.append({
                "name": f"{chart_label}数值轴格式",
                "passed": False,
                "score": 0,
                "message": "数值轴格式不一致",
                "category": "图表"
            })

        return results

    def _compare_area_format(self, answer_xml, student_xml, chart_label):
        """
        对比绘图区和图表区格式

        Args:
            answer_xml: 答案chart XML
            student_xml: 学生chart XML
            chart_label: 图表标签

        Returns:
            list: 对比结果
        """
        def get_plot_area_spPr(xml):
            match = re.search(r'<c:plotArea>(.*?)</c:plotArea>', xml, re.DOTALL)
            if match:
                pa = match.group(1)
                spPr = re.match(r'\s*<c:spPr>(.*?)</c:spPr>', pa, re.DOTALL)
                if spPr:
                    return spPr.group(0)
            return ''

        def get_chart_space_spPr(xml):
            match = re.search(r'<c:chartSpace>.*?<c:spPr>(.*?)</c:spPr>', xml, re.DOTALL)
            if match:
                return match.group(0)
            return ''

        answer_plot = get_plot_area_spPr(answer_xml)
        student_plot = get_plot_area_spPr(student_xml)
        answer_chart = get_chart_space_spPr(answer_xml)
        student_chart = get_chart_space_spPr(student_xml)

        plot_match = self._normalize_whitespace(answer_plot) == self._normalize_whitespace(student_plot)
        chart_match = self._normalize_whitespace(answer_chart) == self._normalize_whitespace(student_chart)

        results = []

        if plot_match:
            results.append({
                "name": f"{chart_label}绘图区格式",
                "passed": True,
                "score": 2,
                "message": "绘图区格式一致",
                "category": "图表"
            })
        else:
            results.append({
                "name": f"{chart_label}绘图区格式",
                "passed": False,
                "score": 0,
                "message": "绘图区格式不一致",
                "category": "图表"
            })

        if chart_match:
            results.append({
                "name": f"{chart_label}图表区格式",
                "passed": True,
                "score": 2,
                "message": "图表区格式一致",
                "category": "图表"
            })
        else:
            results.append({
                "name": f"{chart_label}图表区格式",
                "passed": False,
                "score": 0,
                "message": "图表区格式不一致",
                "category": "图表"
            })

        return results

    # ================================================================
    # 4. 页面设置对比
    # ================================================================

    def _compare_page_setup(self):
        """
        对比页面设置：orientation, scale, pageMargins, printOptions, pageSetUpPr, printArea

        Returns:
            list: 对比结果详情列表
        """
        details = []

        try:
            answer_sheet_files = self._list_xl_files(self.answer_file, r'xl/worksheets/sheet\d+\.xml')
            student_sheet_files = self._list_xl_files(self.student_file, r'xl/worksheets/sheet\d+\.xml')

            if not answer_sheet_files or not student_sheet_files:
                return details

            sheet_count = min(len(answer_sheet_files), len(student_sheet_files))
            all_passed = True
            messages = []

            for i in range(sheet_count):
                answer_xml = self._read_xml_from_zip(self.answer_file, answer_sheet_files[i])
                student_xml = self._read_xml_from_zip(self.student_file, student_sheet_files[i])

                if answer_xml is None or student_xml is None:
                    continue

                sheet_label = f"Sheet{i+1}"

                # 对比pageSetup属性
                answer_page_setup = self._extract_page_setup(answer_xml)
                student_page_setup = self._extract_page_setup(student_xml)

                if answer_page_setup != student_page_setup:
                    all_passed = False
                    diff_keys = []
                    for key in set(list(answer_page_setup.keys()) + list(student_page_setup.keys())):
                        if answer_page_setup.get(key) != student_page_setup.get(key):
                            diff_keys.append(key)
                    messages.append(
                        f"{sheet_label} pageSetup不一致: {', '.join(diff_keys)}"
                    )

                # 对比pageMargins
                answer_margins = self._extract_page_margins(answer_xml)
                student_margins = self._extract_page_margins(student_xml)

                if answer_margins != student_margins:
                    all_passed = False
                    messages.append(
                        f"{sheet_label} pageMargins不一致，"
                        f"答案: {answer_margins}，学生: {student_margins}"
                    )

                # 对比printOptions
                answer_print_opts = self._extract_print_options(answer_xml)
                student_print_opts = self._extract_print_options(student_xml)

                if answer_print_opts != student_print_opts:
                    all_passed = False
                    messages.append(
                        f"{sheet_label} printOptions不一致，"
                        f"答案: {answer_print_opts}，学生: {student_print_opts}"
                    )

                # 对比pageSetUpPr
                answer_setup_pr = self._extract_page_setup_pr(answer_xml)
                student_setup_pr = self._extract_page_setup_pr(student_xml)

                if answer_setup_pr != student_setup_pr:
                    all_passed = False
                    messages.append(
                        f"{sheet_label} pageSetUpPr不一致，"
                        f"答案: {answer_setup_pr}，学生: {student_setup_pr}"
                    )

                # 对比printArea（通过definedNames）
                answer_print_area = self._extract_print_area(answer_xml, i)
                student_print_area = self._extract_print_area(student_xml, i)

                if answer_print_area != student_print_area:
                    all_passed = False
                    messages.append(
                        f"{sheet_label} printArea不一致，"
                        f"答案: {answer_print_area}，学生: {student_print_area}"
                    )

            if all_passed:
                details.append({
                    "name": "页面设置",
                    "passed": True,
                    "score": 5,
                    "message": "所有工作表页面设置一致",
                    "category": "页面设置"
                })
            else:
                details.append({
                    "name": "页面设置",
                    "passed": False,
                    "score": 0,
                    "message": f"页面设置不一致: {'; '.join(messages[:5])}",
                    "category": "页面设置"
                })

        except Exception as e:
            details.append({
                "name": "页面设置",
                "passed": False,
                "score": 0,
                "message": f"页面设置对比异常: {str(e)}",
                "category": "页面设置"
            })

        return details

    def _extract_page_setup(self, sheet_xml):
        """
        提取pageSetup属性

        Args:
            sheet_xml: sheet的XML内容

        Returns:
            dict: pageSetup属性字典
        """
        if sheet_xml is None:
            return {}
        match = re.search(r'<pageSetup\s([^>]+)/>', sheet_xml)
        if match:
            attrs = match.group(1)
            result = {}
            for attr_match in re.finditer(r'(\w+)="([^"]*)"', attrs):
                result[attr_match.group(1)] = attr_match.group(2)
            return result
        return {}

    def _extract_page_margins(self, sheet_xml):
        """
        提取pageMargins值

        Args:
            sheet_xml: sheet的XML内容

        Returns:
            dict: pageMargins属性字典
        """
        if sheet_xml is None:
            return {}
        match = re.search(r'<pageMargins\s([^>]+)/>', sheet_xml)
        if match:
            attrs = match.group(1)
            result = {}
            for attr_match in re.finditer(r'(\w+)="([^"]*)"', attrs):
                result[attr_match.group(1)] = attr_match.group(2)
            return result
        return {}

    def _extract_print_options(self, sheet_xml):
        """
        提取printOptions属性

        Args:
            sheet_xml: sheet的XML内容

        Returns:
            dict: printOptions属性字典
        """
        if sheet_xml is None:
            return {}
        match = re.search(r'<printOptions\s([^>]+)/>', sheet_xml)
        if match:
            attrs = match.group(1)
            result = {}
            for attr_match in re.finditer(r'(\w+)="([^"]*)"', attrs):
                result[attr_match.group(1)] = attr_match.group(2)
            return result
        return {}

    def _extract_page_setup_pr(self, sheet_xml):
        """
        提取pageSetUpPr属性

        Args:
            sheet_xml: sheet的XML内容

        Returns:
            dict: pageSetUpPr属性字典
        """
        if sheet_xml is None:
            return {}
        match = re.search(r'<pageSetUpPr\s([^>]+)/>', sheet_xml)
        if match:
            attrs = match.group(1)
            result = {}
            for attr_match in re.finditer(r'(\w+)="([^"]*)"', attrs):
                result[attr_match.group(1)] = attr_match.group(2)
            return result
        return {}

    def _extract_print_area(self, sheet_xml, sheet_index):
        """
        提取打印区域

        Args:
            sheet_xml: sheet的XML内容
            sheet_index: 工作表索引（从0开始）

        Returns:
            str: 打印区域字符串
        """
        if sheet_xml is None:
            return ''
        # 打印区域可能在sheet XML中直接定义
        match = re.search(r'<pageSetup[^>]*printArea="([^"]*)"', sheet_xml)
        if match:
            return match.group(1)
        # 也可能在definedNames中
        return ''

    # ================================================================
    # 5. 条件格式对比
    # ================================================================

    def _compare_conditional_formatting(self):
        """
        对比条件格式规则数量和类型

        Returns:
            list: 对比结果详情列表
        """
        details = []

        try:
            answer_sheet_files = self._list_xl_files(self.answer_file, r'xl/worksheets/sheet\d+\.xml')
            student_sheet_files = self._list_xl_files(self.student_file, r'xl/worksheets/sheet\d+\.xml')

            if not answer_sheet_files or not student_sheet_files:
                return details

            sheet_count = min(len(answer_sheet_files), len(student_sheet_files))
            all_passed = True
            messages = []

            for i in range(sheet_count):
                answer_xml = self._read_xml_from_zip(self.answer_file, answer_sheet_files[i])
                student_xml = self._read_xml_from_zip(self.student_file, student_sheet_files[i])

                if answer_xml is None or student_xml is None:
                    continue

                sheet_label = f"Sheet{i+1}"

                answer_cf = self._extract_conditional_formatting(answer_xml)
                student_cf = self._extract_conditional_formatting(student_xml)

                if answer_cf == student_cf:
                    messages.append(f"{sheet_label}: 条件格式一致")
                else:
                    all_passed = False
                    messages.append(
                        f"{sheet_label}: 条件格式不一致，"
                        f"答案: {answer_cf}，学生: {student_cf}"
                    )

            if all_passed:
                details.append({
                    "name": "条件格式",
                    "passed": True,
                    "score": 5,
                    "message": "所有工作表条件格式一致",
                    "category": "条件格式"
                })
            else:
                details.append({
                    "name": "条件格式",
                    "passed": False,
                    "score": 0,
                    "message": f"条件格式不一致: {'; '.join(messages[:5])}",
                    "category": "条件格式"
                })

        except Exception as e:
            details.append({
                "name": "条件格式",
                "passed": False,
                "score": 0,
                "message": f"条件格式对比异常: {str(e)}",
                "category": "条件格式"
            })

        return details

    def _extract_conditional_formatting(self, sheet_xml):
        """
        提取条件格式规则数量和类型

        Args:
            sheet_xml: sheet的XML内容

        Returns:
            list: 条件格式规则列表，每个元素为 (sqref, rule_type) 元组
        """
        if sheet_xml is None:
            return []
        rules = []
        cf_matches = re.findall(
            r'<conditionalFormatting\s[^>]*sqref="([^"]*)">(.*?)</conditionalFormatting>',
            sheet_xml, re.DOTALL
        )
        for sqref, cf_content in cf_matches:
            # 提取规则类型
            rule_types = re.findall(
                r'<(?:c|mc:AlternateContent)[^>]*>.*?(?:<c:|<x14:)(\w+)',
                cf_content, re.DOTALL
            )
            for rule_type in rule_types:
                rules.append((sqref, rule_type))
        return rules

    # ================================================================
    # 6. 合并单元格对比
    # ================================================================

    def _compare_merged_cells(self):
        """
        对比合并单元格范围

        Returns:
            list: 对比结果详情列表
        """
        details = []

        try:
            answer_sheet_files = self._list_xl_files(self.answer_file, r'xl/worksheets/sheet\d+\.xml')
            student_sheet_files = self._list_xl_files(self.student_file, r'xl/worksheets/sheet\d+\.xml')

            if not answer_sheet_files or not student_sheet_files:
                return details

            sheet_count = min(len(answer_sheet_files), len(student_sheet_files))
            all_passed = True
            messages = []

            for i in range(sheet_count):
                answer_xml = self._read_xml_from_zip(self.answer_file, answer_sheet_files[i])
                student_xml = self._read_xml_from_zip(self.student_file, student_sheet_files[i])

                if answer_xml is None or student_xml is None:
                    continue

                sheet_label = f"Sheet{i+1}"

                answer_merged = self._extract_merged_cells(answer_xml)
                student_merged = self._extract_merged_cells(student_xml)

                if sorted(answer_merged) == sorted(student_merged):
                    messages.append(f"{sheet_label}: 合并单元格一致")
                else:
                    all_passed = False
                    messages.append(
                        f"{sheet_label}: 合并单元格不一致，"
                        f"答案: {sorted(answer_merged)}，学生: {sorted(student_merged)}"
                    )

            if all_passed:
                details.append({
                    "name": "合并单元格",
                    "passed": True,
                    "score": 5,
                    "message": "所有工作表合并单元格一致",
                    "category": "合并单元格"
                })
            else:
                details.append({
                    "name": "合并单元格",
                    "passed": False,
                    "score": 0,
                    "message": f"合并单元格不一致: {'; '.join(messages[:5])}",
                    "category": "合并单元格"
                })

        except Exception as e:
            details.append({
                "name": "合并单元格",
                "passed": False,
                "score": 0,
                "message": f"合并单元格对比异常: {str(e)}",
                "category": "合并单元格"
            })

        return details

    def _extract_merged_cells(self, sheet_xml):
        """
        提取合并单元格范围

        Args:
            sheet_xml: sheet的XML内容

        Returns:
            list: 合并单元格范围列表，如 ["A1:B2", "C3:D5"]
        """
        if sheet_xml is None:
            return []
        # 匹配 mergeCell 标签的 ref 属性
        return re.findall(r'<mergeCell\s+ref="([^"]+)"/>', sheet_xml)

    # ================================================================
    # 任务点对比辅助方法
    # ================================================================

    def _load_chart_xmls(self, filepath):
        """加载文件中所有图表的XML"""
        xmls = []
        try:
            with zipfile.ZipFile(filepath) as z:
                chart_files = sorted([n for n in z.namelist() if 'charts/chart' in n and n.endswith('.xml')])
                for cf in chart_files:
                    xmls.append(z.read(cf).decode('utf-8', errors='ignore'))
        except Exception:
            pass
        return xmls

    def _load_drawings(self, filepath):
        """加载文件中所有drawing的XML"""
        xmls = []
        try:
            with zipfile.ZipFile(filepath) as z:
                drawing_files = sorted([n for n in z.namelist() if 'drawings/drawing' in n and n.endswith('.xml')])
                for df in drawing_files:
                    xmls.append(z.read(df).decode('utf-8', errors='ignore'))
        except Exception:
            pass
        return xmls

    def _get_all_chart_titles(self, chart_xmls):
        """获取所有图表的标题"""
        titles = []
        for xml in chart_xmls:
            match = re.search(r'<c:title>.*?<c:tx>(.*?)</c:tx>.*?</c:title>', xml, re.DOTALL)
            if match:
                texts = re.findall(r'<a:t>([^<]+)</a:t>', match.group(1))
                titles.append(''.join(texts))
            else:
                titles.append('')
        return titles

    def _chart_has_legend(self, chart_xmls, index):
        """检查指定图表是否有图例"""
        if index < len(chart_xmls):
            return '<c:legend>' in chart_xmls[index]
        return False

    def _chart_has_data_labels(self, chart_xmls, index):
        """检查指定图表是否有数据标签（且未被删除）"""
        if index < len(chart_xmls):
            xml = chart_xmls[index]
            if '<c:dLbls>' not in xml:
                return False
            dLbls = re.search(r'<c:dLbls>(.*?)</c:dLbls>', xml, re.DOTALL)
            if dLbls and '<c:delete val="1"/>' in dLbls.group(0):
                return False
            return True
        return False

    def _chart_gridline_count(self, chart_xmls, index):
        """统计指定图表的网格线数量"""
        if index < len(chart_xmls):
            return len(re.findall(r'<c:majorGridlines', chart_xmls[index]))
        return 0

    def _get_chart_position(self, drawing_xmls, index):
        """获取指定图表的位置（from_col, from_row, to_col, to_row）"""
        for xml in drawing_xmls:
            anchors = re.findall(r'<xdr:twoCellAnchor>(.*?)</xdr:twoCellAnchor>', xml, re.DOTALL)
            for i, a in enumerate(anchors):
                if i == index:
                    from_ = re.search(r'<xdr:from>(.*?)</xdr:from>', a, re.DOTALL)
                    to_ = re.search(r'<xdr:to>(.*?)</xdr:to>', a, re.DOTALL)
                    if from_ and to_:
                        col = int(re.search(r'<xdr:col>(\d+)</xdr:col>', from_.group(1)).group(1))
                        row = int(re.search(r'<xdr:row>(\d+)</xdr:row>', from_.group(1)).group(1))
                        col2 = int(re.search(r'<xdr:col>(\d+)</xdr:col>', to_.group(1)).group(1))
                        row2 = int(re.search(r'<xdr:row>(\d+)</xdr:row>', to_.group(1)).group(1))
                        return (col, row, col2, row2)
        return None

    # ================================================================
    # 任务点对比方法（每个 compare 类型对应一个方法）
    # ================================================================

    def _compare_chart_title(self, params):
        """对比图表标题 - 按标题匹配而非索引"""
        chart_idx = params.get("chart_index", 0)
        expected = params.get("expected_title", "")

        answer_titles = self._get_all_chart_titles(self.answer_xmls)
        student_titles = self._get_all_chart_titles(self.student_xmls)

        # 先尝试按索引匹配
        if chart_idx < len(student_titles) and chart_idx < len(answer_titles):
            if answer_titles[chart_idx] == student_titles[chart_idx]:
                return True, f"图表标题一致: {answer_titles[chart_idx]}"

        # 按索引不匹配时，尝试在学生文件中查找包含expected标题的图表
        if expected:
            for i, title in enumerate(student_titles):
                if expected in title or title in expected:
                    return True, f"找到匹配的图表标题: {title}"
            # 模糊匹配
            for i, title in enumerate(student_titles):
                # 提取关键词进行匹配
                answer_keywords = set(expected.replace('图表', '').replace('柱形图', '').replace('图', ''))
                student_keywords = set(title.replace('图表', '').replace('柱形图', '').replace('图', ''))
                common = answer_keywords & student_keywords
                if len(common) >= 2:
                    return True, f"图表标题基本匹配: {title}"

        if chart_idx >= len(student_titles):
            return False, f"学生文件只有{len(student_titles)}个图表，需要至少{chart_idx + 1}个"
        if chart_idx >= len(answer_titles):
            return False, f"答案文件只有{len(answer_titles)}个图表"

        return False, f"标题不一致，答案: {answer_titles[chart_idx]}，学生: {student_titles[chart_idx]}"

    def _compare_chart_has_legend(self, params):
        """对比图表是否有图例 - 按标题匹配图表"""
        chart_idx = params.get("chart_index", 0)
        expected_title = params.get("expected_title", "")

        # 找到学生对应的图表索引
        student_idx = self._find_student_chart_index(expected_title, chart_idx)
        if student_idx is None:
            # 学生没有对应图表，不得分
            return False, f"学生文件无对应图表"

        answer_has = self._chart_has_legend(self.answer_xmls, chart_idx)
        student_has = self._chart_has_legend(self.student_xmls, student_idx)

        if answer_has == student_has:
            return True, f"图例状态一致: {'有' if answer_has else '无'}图例"
        else:
            return False, f"图例不一致，答案: {'有' if answer_has else '无'}，学生: {'有' if student_has else '无'}"

    def _compare_chart_has_data_labels(self, params):
        """对比图表是否有数据标签 - 按标题匹配图表"""
        chart_idx = params.get("chart_index", 0)
        expected_title = params.get("expected_title", "")

        student_idx = self._find_student_chart_index(expected_title, chart_idx)
        if student_idx is None:
            return False, f"学生文件无对应图表"

        answer_has = self._chart_has_data_labels(self.answer_xmls, chart_idx)
        student_has = self._chart_has_data_labels(self.student_xmls, student_idx)

        if answer_has == student_has:
            return True, f"数据标签状态一致"
        else:
            return False, f"数据标签不一致，答案: {'有' if answer_has else '无'}，学生: {'有' if student_has else '无'}"

    def _compare_chart_no_gridlines(self, params):
        """对比图表是否删除了网格线 - 按标题匹配图表"""
        chart_idx = params.get("chart_index", 0)
        expected_title = params.get("expected_title", "")

        student_idx = self._find_student_chart_index(expected_title, chart_idx)
        if student_idx is None:
            return False, f"学生文件无对应图表"

        answer_count = self._chart_gridline_count(self.answer_xmls, chart_idx)
        student_count = self._chart_gridline_count(self.student_xmls, student_idx)

        if answer_count == student_count:
            return True, f"网格线一致: {answer_count}条"
        else:
            return False, f"网格线不一致，答案: {answer_count}条，学生: {student_count}条"

    def _compare_chart_axis_color(self, params):
        """对比坐标轴填充颜色 - 按标题匹配图表"""
        import re
        chart_idx = params.get("chart_index", 0)
        expected_title = params.get("expected_title", "")

        student_idx = self._find_student_chart_index(expected_title, chart_idx)
        if student_idx is None:
            return False, f"学生文件无对应图表"

        def get_axis_colors(chart_xmls, idx):
            if idx >= len(chart_xmls):
                return {}
            xml = chart_xmls[idx]
            colors = {}
            plot_area = re.search(r'<c:plotArea>(.*?)</c:plotArea>', xml, re.DOTALL)
            if plot_area:
                pa = plot_area.group(1)
                cat_ax = re.search(r'<c:catAx>(.*?)</c:catAx>', pa, re.DOTALL)
                if cat_ax:
                    spPr = re.search(r'<[a-z]+:spPr>(.*?)</[a-z]+:spPr>', cat_ax.group(1), re.DOTALL)
                    if spPr:
                        colors['catAx'] = re.findall(r'<a:schemeClr val="(\w+)"', spPr.group(1))
                val_ax = re.search(r'<c:valAx>(.*?)</c:valAx>', pa, re.DOTALL)
                if val_ax:
                    spPr = re.search(r'<[a-z]+:spPr>(.*?)</[a-z]+:spPr>', val_ax.group(1), re.DOTALL)
                    if spPr:
                        colors['valAx'] = re.findall(r'<a:schemeClr val="(\w+)"', spPr.group(1))
            return colors

        answer_colors = get_axis_colors(self.answer_xmls, chart_idx)
        student_colors = get_axis_colors(self.student_xmls, student_idx)

        if answer_colors == student_colors:
            return True, f"坐标轴颜色一致"
        else:
            return False, f"坐标轴颜色不一致，答案: {answer_colors}，学生: {student_colors}"

    def _compare_chart_plot_area_color(self, params):
        """对比绘图区填充颜色 - 按标题匹配图表"""
        import re
        chart_idx = params.get("chart_index", 0)
        expected_title = params.get("expected_title", "")

        student_idx = self._find_student_chart_index(expected_title, chart_idx)
        if student_idx is None:
            return False, f"学生文件无对应图表"

        def get_plot_area_color(chart_xmls, idx):
            if idx >= len(chart_xmls):
                return []
            xml = chart_xmls[idx]
            plot_area = re.search(r'<c:plotArea>(.*?)</c:plotArea>', xml, re.DOTALL)
            if plot_area:
                pa = plot_area.group(1)
                pa_clean = re.sub(r'<c:ser>.*?</c:ser>', '', pa, flags=re.DOTALL)
                pa_clean = re.sub(r'<c:catAx>.*?</c:catAx>', '', pa_clean, flags=re.DOTALL)
                pa_clean = re.sub(r'<c:valAx>.*?</c:valAx>', '', pa_clean, flags=re.DOTALL)
                pa_clean = re.sub(r'<c:dLbls>.*?</c:dLbls>', '', pa_clean, flags=re.DOTALL)
                spPr = re.search(r'<c:spPr>(.*?)</c:spPr>', pa_clean, re.DOTALL)
                if spPr:
                    return re.findall(r'<a:schemeClr val="(\w+)"', spPr.group(1))
            return []

        answer_color = get_plot_area_color(self.answer_xmls, chart_idx)
        student_color = get_plot_area_color(self.student_xmls, student_idx)

        if answer_color == student_color:
            return True, f"绘图区颜色一致"
        else:
            return False, f"绘图区颜色不一致，答案: {answer_color}，学生: {student_color}"

    def _compare_chart_position(self, params):
        """对比图表位置 - 按标题匹配图表"""
        chart_idx = params.get("chart_index", 0)
        expected_title = params.get("expected_title", "")
        from_col = params.get("from_col", 0)
        from_row = params.get("from_row", 0)
        to_col = params.get("to_col", 0)
        to_row = params.get("to_row", 0)

        student_idx = self._find_student_chart_index(expected_title, chart_idx)
        if student_idx is None:
            return False, f"学生文件无对应图表"

        # 直接用chart_idx在drawing中查找位置（drawing中anchor顺序通常与chart文件顺序一致）
        student_pos = self._get_chart_position(self.student_drawings, chart_idx)
        if student_pos is None:
            # 尝试遍历所有drawing查找
            student_pos = self._get_chart_position_any(self.student_drawings)
        if student_pos is None:
            return False, "未找到图表位置信息"

        # 直接与期望位置比对
        row_tolerance = 3
        col_tolerance = 2

        if (abs(student_pos[0] - from_col) <= col_tolerance and
            abs(student_pos[1] - from_row) <= row_tolerance and
            abs(student_pos[2] - to_col) <= col_tolerance and
            abs(student_pos[3] - to_row) <= row_tolerance):
            return True, f"图表位置正确: 第{student_pos[1]+1}行第{student_pos[0]+1}列 到 第{student_pos[3]+1}行第{student_pos[2]+1}列"
        else:
            return False, f"图表位置不一致，期望: A{from_row+1}:{chr(to_col+65)}{to_row+1}，实际: 第{student_pos[1]+1}行第{student_pos[0]+1}列 到 第{student_pos[3]+1}行第{student_pos[2]+1}列"

    def _find_chart_index_by_title(self, expected_title, default_idx):
        """在答案文件中按标题查找图表索引"""
        if not expected_title:
            return default_idx
        answer_titles = self._get_all_chart_titles(self.answer_xmls)
        for i, title in enumerate(answer_titles):
            if expected_title in title or title in expected_title:
                return i
        return default_idx

    def _get_chart_position_any(self, drawing_xmls):
        """获取任意一个图表的位置"""
        for xml in drawing_xmls:
            anchors = re.findall(r'<xdr:twoCellAnchor>(.*?)</xdr:twoCellAnchor>', xml, re.DOTALL)
            for a in anchors:
                from_ = re.search(r'<xdr:from>(.*?)</xdr:from>', a, re.DOTALL)
                to_ = re.search(r'<xdr:to>(.*?)</xdr:to>', a, re.DOTALL)
                if from_ and to_:
                    col = int(re.search(r'<xdr:col>(\d+)</xdr:col>', from_.group(1)).group(1))
                    row = int(re.search(r'<xdr:row>(\d+)</xdr:row>', from_.group(1)).group(1))
                    col2 = int(re.search(r'<xdr:col>(\d+)</xdr:col>', to_.group(1)).group(1))
                    row2 = int(re.search(r'<xdr:row>(\d+)</xdr:row>', to_.group(1)).group(1))
                    return (col, row, col2, row2)
        return None

    def _find_student_chart_index(self, expected_title, default_idx):
        """在学生文件中按标题查找图表索引"""
        student_titles = self._get_all_chart_titles(self.student_xmls)
        if not expected_title:
            if default_idx < len(student_titles):
                return default_idx
            return None
        # 精确匹配
        for i, title in enumerate(student_titles):
            if expected_title in title or title in expected_title:
                return i
        # 关键词匹配
        answer_keywords = set(expected_title.replace('图表', '').replace('柱形图', '').replace('图', ''))
        for i, title in enumerate(student_titles):
            student_keywords = set(title.replace('图表', '').replace('柱形图', '').replace('图', ''))
            common = answer_keywords & student_keywords
            if len(common) >= 2:
                return i
        # 回退到索引
        if default_idx < len(student_titles):
            return default_idx
        return None

    # ================================================================
    # 任务点对比主方法
    # ================================================================

    def compare_task(self, task_id):
        """按任务点对比评分

        Args:
            task_id: 任务ID，如 "task_5_4_4"

        Returns:
            dict: 评分结果，包含 total_score, max_score, task_name, scoring_method, details
        """
        if task_id not in TASK_CHECK_POINTS:
            # 没有定义任务点，回退到全文件对比
            return self.compare()

        task_def = TASK_CHECK_POINTS[task_id]

        # 预加载数据
        self.answer_xmls = self._load_chart_xmls(self.answer_file)
        self.student_xmls = self._load_chart_xmls(self.student_file)
        self.answer_drawings = self._load_drawings(self.answer_file)
        self.student_drawings = self._load_drawings(self.student_file)

        result = {
            "total_score": 0,
            "max_score": 100,
            "task_name": task_def["task_name"],
            "scoring_method": "answer_compare_task",
            "details": []
        }

        total_max = 0
        for cp in task_def["check_points"]:
            compare_type = cp["compare"]
            params = cp.get("params", {})
            score = cp["score"]
            total_max += score

            try:
                method = getattr(self, f"_compare_{compare_type}", None)
                if method:
                    passed, message = method(params)
                else:
                    passed, message = False, f"未实现的对比方法: {compare_type}"
            except Exception as e:
                passed, message = False, f"对比出错: {str(e)}"

            result["details"].append({
                "id": cp["id"],
                "name": cp["name"],
                "passed": passed,
                "score": score if passed else 0,
                "message": message,
                "category": cp.get("category", self._infer_category(cp))
            })

        result["total_score"] = sum(d["score"] for d in result["details"])
        result["max_score"] = total_max
        return result

    def _infer_category(self, check_point):
        """根据检查点自动推断能力维度"""
        name = check_point.get("name", "")
        compare = check_point.get("compare", "")
        params = check_point.get("params", {})
        
        # 根据检查类型推断
        type_category_map = {
            "sort": "数据排序能力",
            "subtotal": "数据汇总能力",
            "auto_filter": "数据筛选能力",
            "advanced_filter": "数据筛选能力",
            "pivot_table": "数据透视能力",
            "print_area": "页面设置能力",
            "page_margins": "页面设置能力",
            "print_title_rows": "页面设置能力",
            "page_orientation": "页面设置能力",
            "chart_title": "图表创建能力",
            "chart_has_legend": "图表美化能力",
            "chart_has_data_labels": "图表美化能力",
            "chart_no_gridlines": "图表美化能力",
            "chart_axis_color": "图表美化能力",
            "chart_plot_area_color": "图表美化能力",
            "chart_position": "图表布局能力",
            "formula_exists": "函数应用能力",
        }
        
        if compare in type_category_map:
            return type_category_map[compare]
        
        # 根据关键词推断
        keywords_category_map = [
            (["排序", "升序", "降序"], "数据排序能力"),
            (["分类汇总", "subtotal"], "数据汇总能力"),
            (["筛选", "filter"], "数据筛选能力"),
            (["透视表", "pivot"], "数据透视能力"),
            (["图表", "柱形图", "饼图", "折线图"], "图表创建能力"),
            (["数据标签", "图例", "网格线", "颜色"], "图表美化能力"),
            (["打印", "页边距", "标题行", "页面"], "页面设置能力"),
            (["COUNTIF", "SUMIF", "AVERAGEIF"], "条件统计能力"),
            (["IF", "AND", "OR"], "逻辑判断能力"),
            (["RANK", "排名"], "排序排名能力"),
            (["MIN", "MAX", "AVERAGE", "COUNT"], "数据统计能力"),
            (["VLOOKUP", "HLOOKUP", "LOOKUP", "INDEX", "MATCH"], "数据查找能力"),
            (["函数", "公式"], "函数应用能力"),
        ]
        
        for keywords, category in keywords_category_map:
            for kw in keywords:
                if kw.lower() in name.lower() or kw.lower() in compare.lower():
                    return category
        
        return "任务点对比"

    # ================================================================
    # 主对比方法（全文件对比，保留供其他任务使用）
    # ================================================================

    def compare(self):
        """
        执行对比，返回评分结果

        Returns:
            dict: 评分结果，包含 total_score, max_score, details
        """
        result = {
            "total_score": 0,
            "max_score": 100,
            "details": []
        }

        try:
            # 1. 工作表结构对比
            result["details"].extend(self._compare_sheet_structure())

            # 2. 单元格数据对比
            result["details"].extend(self._compare_cell_data())

            # 3. 图表对比
            result["details"].extend(self._compare_charts())

            # 4. 页面设置对比
            result["details"].extend(self._compare_page_setup())

            # 5. 条件格式对比
            result["details"].extend(self._compare_conditional_formatting())

            # 6. 合并单元格对比
            result["details"].extend(self._compare_merged_cells())

            # 计算总分和满分
            total_score = sum(item["score"] for item in result["details"])
            result["total_score"] = total_score
            # max_score保持初始值100（百分制），不再覆盖

        except Exception as e:
            result["details"].append({
                "name": "系统错误",
                "passed": False,
                "score": 0,
                "message": f"对比过程发生异常: {str(e)}",
                "category": "系统"
            })

        return result

    # ================================================================
    # 公式对比方法（用于 task_5_3_2 等函数统计任务）
    # ================================================================

    def _get_cell_formula_and_value(self, file_path, sheet_index, cell_ref):
        """
        获取指定单元格的公式和计算值

        Args:
            file_path: Excel文件路径
            sheet_index: 工作表索引（从0开始）
            cell_ref: 单元格引用（如 "G2"）

        Returns:
            tuple: (formula, value) 公式和计算值，不存在返回 (None, None)
        """
        try:
            # 读取工作表XML
            sheet_files = self._list_xl_files(file_path, r'xl/worksheets/sheet\d+\.xml')
            if sheet_index >= len(sheet_files):
                return None, None

            sheet_xml = self._read_xml_from_zip(file_path, sheet_files[sheet_index])
            if not sheet_xml:
                return None, None

            # 读取共享字符串表
            shared_xml = self._read_xml_from_zip(file_path, 'xl/sharedStrings.xml')
            shared_strings = self._parse_shared_strings(shared_xml) if shared_xml else []

            # 解析单元格引用（如 "G2" -> col=6, row=2）
            col_letter = ''.join([c for c in cell_ref if c.isalpha()]).upper()
            row_num = int(''.join([c for c in cell_ref if c.isdigit()]))
            col_num = 0
            for c in col_letter:
                col_num = col_num * 26 + (ord(c) - ord('A') + 1)

            # 查找单元格
            cell_pattern = rf'<c[^>]*r="{cell_ref}"[^>]*>(.*?)</c>'
            match = re.search(cell_pattern, sheet_xml, re.DOTALL)

            if not match:
                return None, None

            cell_content = match.group(1)

            # 提取公式
            formula = None
            f_match = re.search(r'<f[^>]*>(.*?)</f>', cell_content, re.DOTALL)
            if f_match:
                formula = f_match.group(1)
                # 清理公式中的XML标签
                formula = re.sub(r'<[^>]+>', '', formula)

            # 提取计算值
            value = None
            v_match = re.search(r'<v>([^<]*)</v>', cell_content)
            if v_match:
                value = v_match.group(1)
                # 检查是否是共享字符串
                cell_type = re.search(r't="([^"]+)"', match.group(0))
                if cell_type and cell_type.group(1) == 's':
                    try:
                        idx = int(value)
                        if 0 <= idx < len(shared_strings):
                            value = shared_strings[idx]
                    except (ValueError, IndexError):
                        pass

            return formula, value

        except Exception as e:
            return None, None

    def _compare_formula_match(self, params):
        """
        对比指定单元格的公式是否匹配
        检查公式中是否包含指定的关键词（如 COUNTIF）

        Args:
            params: {
                "sheet_index": 0,
                "cell": "G2",
                "formula_keyword": "COUNTIF",
                "description": "描述"
            }

        Returns:
            tuple: (passed, message)
        """
        sheet_idx = params.get("sheet_index", 0)
        cell_ref = params.get("cell", "")
        keyword = params.get("formula_keyword", "")
        description = params.get("description", "")

        if not cell_ref or not keyword:
            return False, "参数错误：缺少cell或formula_keyword"

        # 获取答案和学生的公式
        answer_formula, _ = self._get_cell_formula_and_value(self.answer_file, sheet_idx, cell_ref)
        student_formula, _ = self._get_cell_formula_and_value(self.student_file, sheet_idx, cell_ref)

        if answer_formula is None:
            return False, f"答案文件{cell_ref}单元格不存在"

        if student_formula is None:
            return False, f"学生文件{cell_ref}单元格不存在或没有公式"

        # 检查学生公式是否包含关键词
        student_has_keyword = keyword.upper() in student_formula.upper()
        answer_has_keyword = keyword.upper() in answer_formula.upper()

        if not student_has_keyword:
            return False, f"{cell_ref}公式不包含{keyword}，学生公式: {student_formula}"

        # 进一步检查公式结构是否相似（简化版：检查引用的列是否一致）
        # 提取公式中的列引用（如 A:A, B:B）
        col_refs = re.findall(r'[A-Z]+:[A-Z]+|\$?[A-Z]+\$?\d+', student_formula)
        answer_col_refs = re.findall(r'[A-Z]+:[A-Z]+|\$?[A-Z]+\$?\d+', answer_formula)

        # 如果学生公式和答案公式的列引用一致，认为正确
        if set(col_refs) == set(answer_col_refs):
            return True, f"{cell_ref}公式正确: {student_formula}"
        else:
            # 只要包含关键词就通过，但给出提示
            return True, f"{cell_ref}公式包含{keyword}: {student_formula}"

    def _compare_formula_filled(self, params):
        """
        检查公式是否正确填充到指定范围

        Args:
            params: {
                "sheet_index": 0,
                "start_cell": "H2",
                "end_cell": "H31",
                "description": "描述"
            }

        Returns:
            tuple: (passed, message)
        """
        sheet_idx = params.get("sheet_index", 0)
        start_cell = params.get("start_cell", "")
        end_cell = params.get("end_cell", "")
        description = params.get("description", "")

        if not start_cell or not end_cell:
            return False, "参数错误：缺少start_cell或end_cell"

        # 解析起始和结束单元格
        start_match = re.match(r'([A-Z]+)(\d+)', start_cell.upper())
        end_match = re.match(r'([A-Z]+)(\d+)', end_cell.upper())

        if not start_match or not end_match:
            return False, "单元格格式错误"

        start_col, start_row = start_match.groups()
        end_col, end_row = end_match.groups()
        start_row, end_row = int(start_row), int(end_row)

        # 只支持单列填充
        if start_col != end_col:
            return False, "暂不支持多列填充检查"

        # 检查每个单元格
        filled_count = 0
        empty_count = 0
        expected_count = end_row - start_row + 1

        for row in range(start_row, end_row + 1):
            cell_ref = f"{start_col}{row}"
            formula, _ = self._get_cell_formula_and_value(self.student_file, sheet_idx, cell_ref)
            if formula:
                filled_count += 1
            else:
                empty_count += 1

        # 检查答案文件中的公式作为参考
        answer_formula, _ = self._get_cell_formula_and_value(self.answer_file, sheet_idx, start_cell)

        if filled_count == expected_count:
            return True, f"{start_cell}:{end_cell}全部填充公式，共{filled_count}个单元格"
        elif filled_count > 0:
            return False, f"{start_cell}:{end_cell}仅填充{filled_count}/{expected_count}个单元格，缺少{empty_count}个"
        else:
            return False, f"{start_cell}:{end_cell}未填充公式"

    def _compare_cell_value_match(self, params):
        """
        对比指定单元格的计算值是否与答案一致

        Args:
            params: {
                "sheet_index": 0,
                "cells": ["G2", "G3", "G4", "H2"],
                "tolerance": 0.01,
                "description": "描述"
            }

        Returns:
            tuple: (passed, message)
        """
        sheet_idx = params.get("sheet_index", 0)
        cells = params.get("cells", [])
        tolerance = params.get("tolerance", 0.01)
        description = params.get("description", "")

        if not cells:
            return False, "参数错误：缺少cells"

        match_count = 0
        mismatch_details = []

        for cell_ref in cells:
            answer_formula, answer_value = self._get_cell_formula_and_value(self.answer_file, sheet_idx, cell_ref)
            student_formula, student_value = self._get_cell_formula_and_value(self.student_file, sheet_idx, cell_ref)

            if answer_value is None:
                mismatch_details.append(f"{cell_ref}:答案不存在")
                continue

            if student_value is None:
                mismatch_details.append(f"{cell_ref}:学生未填写")
                continue

            # 尝试数值比较
            try:
                answer_num = float(answer_value)
                student_num = float(student_value)
                if abs(answer_num - student_num) <= tolerance:
                    match_count += 1
                else:
                    mismatch_details.append(f"{cell_ref}:答案{answer_num} vs 学生{student_num}")
            except (ValueError, TypeError):
                # 文本比较
                if str(answer_value).strip() == str(student_value).strip():
                    match_count += 1
                else:
                    mismatch_details.append(f"{cell_ref}:答案'{answer_value}' vs 学生'{student_value}'")

        total = len(cells)
        if match_count == total:
            return True, f"全部{total}个单元格计算结果正确"
        else:
            return False, f"{match_count}/{total}个单元格正确，错误: {'; '.join(mismatch_details)}"

    def _compare_formula_exists(self, params):
        """
        严格检查学生文件中是否包含指定类型的公式，并验证参数引用是否正确

        不仅检查函数名，还要检查：
        1. 引用的数据范围（如 $B$4:$B$28）
        2. 引用的条件区域
        3. 公式结构与答案的一致性

        Args:
            params: {
                "formula_keyword": "COUNTIF",
                "description": "描述"
            }

        Returns:
            tuple: (passed, message)
        """
        keyword = params.get("formula_keyword", "")
        description = params.get("description", "")

        if not keyword:
            return False, "参数错误：缺少formula_keyword"

        # 扫描答案文件和学生文件中的公式
        answer_formulas = self._scan_all_formulas(self.answer_file)
        student_formulas = self._scan_all_formulas(self.student_file)

        # 精准匹配函数名（避免IF匹配到COUNTIF）
        # 使用正则表达式匹配函数名后紧跟左括号的情况
        # 例如：IF( 不应该匹配 COUNTIF( 中的 IF(
        # 支持带前缀的函数：_xlfn.RANK.EQ( 也应该匹配 RANK
        def match_formula_keyword(formula, kw):
            """精准匹配函数名，确保是独立的函数而非其他函数的一部分"""
            # 先尝试直接匹配：函数名前不能是字母（但可以是.或_，如 _xlfn.RANK.EQ）
            pattern = rf'(?<![A-Z]){re.escape(kw)}\s*[\.\(]'
            if re.search(pattern, formula.upper()):
                return True
            # 再尝试匹配带点号前缀的情况（如 RANK.EQ 应该匹配 RANK）
            pattern2 = rf'(?<![A-Z]){re.escape(kw)}\.\w+\s*\('
            return bool(re.search(pattern2, formula.upper()))

        # 过滤出精准匹配关键词的公式
        answer_matching = [f for f in answer_formulas if match_formula_keyword(f, keyword)]
        student_matching = [f for f in student_formulas if match_formula_keyword(f, keyword)]

        # 检查学生文件中是否有该关键词的公式
        if not student_matching:
            return False, f"未使用{keyword}函数"

        # 提取答案中的关键引用范围（数据区域、条件区域等）
        # 对于COUNTIF/AVERAGEIF/SUMIF：提取条件区域和求和区域
        # 对于IF：提取判断条件和返回值
        # 对于VLOOKUP：提取查找范围和返回列

        def extract_formula_signature(formula, kw):
            """提取公式的关键特征（引用范围、结构模式）"""
            sig = {"ranges": set(), "sheets": set(), "structure": ""}

            # 提取所有单元格/区域引用（如 $B$4:$B$28 或 B4:B28）
            # 使用非捕获组 (?:...) 避免re.findall只返回捕获组内容
            range_pattern = r'\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?'
            sig["ranges"] = set(re.findall(range_pattern, formula))

            # 提取工作表引用（如 各城区销售经理!）
            sheet_pattern = r'([^!\s]+)!'
            sig["sheets"] = set(re.findall(sheet_pattern, formula))

            # 提取函数调用的结构（简化版）
            # 对于 COUNTIF($B$4:$B$28,R3) -> COUNTIF(range,condition)
            kw_clean = kw.replace("(", "").replace(")", "")
            if kw_clean.upper() in formula.upper():
                # 找到函数调用部分
                func_start = formula.upper().find(kw_clean.upper())
                if func_start >= 0:
                    # 提取函数名和参数
                    rest = formula[func_start:]
                    paren_start = rest.find("(")
                    if paren_start > 0:
                        # 找到匹配的括号
                        depth = 0
                        paren_end = -1
                        for i, c in enumerate(rest[paren_start:], paren_start):
                            if c == '(':
                                depth += 1
                            elif c == ')':
                                depth -= 1
                                if depth == 0:
                                    paren_end = i
                                    break
                        if paren_end > 0:
                            sig["structure"] = rest[:paren_end+1]

            return sig

        # 分析答案公式的特征
        answer_sigs = [extract_formula_signature(f, keyword) for f in answer_matching]
        student_sigs = [extract_formula_signature(f, keyword) for f in student_matching]

        # 合并答案的所有引用范围
        answer_all_ranges = set()
        answer_all_sheets = set()
        for sig in answer_sigs:
            answer_all_ranges.update(sig["ranges"])
            answer_all_sheets.update(sig["sheets"])

        # 合并学生的所有引用范围
        student_all_ranges = set()
        student_all_sheets = set()
        for sig in student_sigs:
            student_all_ranges.update(sig["ranges"])
            student_all_sheets.update(sig["sheets"])

        # 检查引用范围是否匹配
        # 关键数据区域必须一致（如 $B$4:$B$28, $C$4:$C$28 等）
        missing_ranges = answer_all_ranges - student_all_ranges
        extra_ranges = student_all_ranges - answer_all_ranges

        # 获取学生公式示例用于显示
        student_example = student_matching[0] if student_matching else ""
        answer_example = answer_matching[0] if answer_matching else ""

        # 判断评分结果
        if not missing_ranges and not extra_ranges:
            # 引用范围完全一致
            return True, f"{keyword}函数使用正确，引用范围匹配: {student_example}"
        elif missing_ranges and not extra_ranges:
            # 缺少必要的引用范围
            return False, f"{keyword}函数引用范围不完整，缺少: {', '.join(missing_ranges)}，当前: {student_example}"
        elif extra_ranges and not missing_ranges:
            # 有额外的引用范围（可能是正确的变体）
            return True, f"{keyword}函数使用正确，包含额外引用: {student_example}"
        else:
            # 引用范围有差异
            return False, f"{keyword}函数引用范围不匹配，答案: {answer_example}，学生: {student_example}"


    def _scan_all_formulas(self, file_path):
        """
        扫描Excel文件中所有工作表的所有公式

        Args:
            file_path: Excel文件路径

        Returns:
            list: 公式字符串列表
        """
        formulas = []
        try:
            with zipfile.ZipFile(file_path) as zf:
                sheet_files = [n for n in zf.namelist() if re.match(r'xl/worksheets/sheet\d+\.xml', n)]

                for sf in sheet_files:
                    xml_content = zf.read(sf).decode('utf-8', errors='ignore')

                    # 查找所有公式 <f>...</f>
                    # 处理普通公式
                    for match in re.finditer(r'<f(?:\s[^>]*)?>(.*?)</f>', xml_content, re.DOTALL):
                        formula = match.group(1).strip()
                        # 跳过共享公式引用 <f t="shared" si="0"/>
                        if not formula:
                            continue
                        # 清理XML实体
                        formula = formula.replace('&gt;', '>').replace('&lt;', '<').replace('&quot;', '"').replace('&amp;', '&')
                        formulas.append(formula)
        except Exception:
            pass

        return formulas

    # ================================================================
    # 结构检查方法（用于 task_5_4_1 到 task_5_4_5）
    # ================================================================

    def _compare_structure_check(self, params):
        """
        结构检查方法，用于检查排序、分类汇总、筛选、数据透视表、图表、页面设置等

        Args:
            params: {
                "sheet": "排序表1",
                "check_type": "sort",  # sort/subtotal/auto_filter/advanced_filter/pivot_table/print_area/page_margins等
                ... 其他参数
            }

        Returns:
            tuple: (passed, message)
        """
        sheet_name = params.get("sheet", "")
        check_type = params.get("check_type", "")

        if not sheet_name or not check_type:
            return False, "参数错误：缺少sheet或check_type"

        # 获取工作表索引
        answer_sheet_idx = self._get_sheet_index(self.answer_file, sheet_name)
        student_sheet_idx = self._get_sheet_index(self.student_file, sheet_name)

        # 如果答案文件中没有该工作表，学生不得分
        if answer_sheet_idx is None:
            return False, f"标准答案中不存在工作表'{sheet_name}'，该检查点不得分"

        # 如果学生文件中没有该工作表，返回失败
        if student_sheet_idx is None:
            return False, f"学生文件中不存在工作表'{sheet_name}'"

        # 根据检查类型调用相应的检查方法
        check_methods = {
            "sort": self._check_sort_state,
            "subtotal": self._check_subtotal,
            "auto_filter": self._check_auto_filter,
            "advanced_filter": self._check_advanced_filter,
            "pivot_table": self._check_pivot_table,
            "pivot_table_row": self._check_pivot_table_row,
            "pivot_table_col": self._check_pivot_table_col,
            "pivot_table_data": self._check_pivot_table_data,
            "pivot_table_summary": self._check_pivot_table_summary,
            "print_area": self._check_print_area,
            "page_margins": self._check_page_margins,
            "print_title_rows": self._check_print_title_rows,
            "page_orientation": self._check_page_orientation,
            "print_preview": self._check_print_preview,
        }

        method = check_methods.get(check_type)
        if method:
            return method(params, answer_sheet_idx, student_sheet_idx)
        else:
            return False, f"未实现的检查类型: {check_type}"

    def _get_sheet_index(self, file_path, sheet_name):
        """
        获取工作表名称对应的索引

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称

        Returns:
            int: 工作表索引（从0开始），不存在返回None
        """
        try:
            workbook_xml = self._read_xml_from_zip(file_path, 'xl/workbook.xml')
            if workbook_xml is None:
                return None

            # 提取所有工作表名称
            sheets = re.findall(r'<sheet\s[^>]*name="([^"]+)"', workbook_xml)

            for i, name in enumerate(sheets):
                if name == sheet_name:
                    return i

            return None
        except Exception:
            return None

    def _get_sheet_xml_by_index(self, file_path, sheet_index):
        """
        根据索引获取工作表XML内容

        Args:
            file_path: Excel文件路径
            sheet_index: 工作表索引（从0开始）

        Returns:
            str: 工作表XML内容
        """
        try:
            sheet_files = self._list_xl_files(file_path, r'xl/worksheets/sheet\d+\.xml')
            if sheet_index >= len(sheet_files):
                return None
            return self._read_xml_from_zip(file_path, sheet_files[sheet_index])
        except Exception:
            return None

    def _check_sort_state(self, params, answer_sheet_idx, student_sheet_idx):
        """
        检查排序状态 - 与标准答案比对

        Args:
            params: {
                "sort_keys": ["城区:asc", "建造时间:desc"]
            }
            answer_sheet_idx: 答案工作表索引
            student_sheet_idx: 学生工作表索引

        Returns:
            tuple: (passed, message)
        """
        sort_keys = params.get("sort_keys", [])

        # 获取工作表XML
        student_sheet_xml = self._get_sheet_xml_by_index(self.student_file, student_sheet_idx)
        answer_sheet_xml = self._get_sheet_xml_by_index(self.answer_file, answer_sheet_idx)

        if student_sheet_xml is None:
            return False, "无法读取学生工作表"
        
        if answer_sheet_xml is None:
            return False, "标准答案中不存在该工作表，该检查点不得分"

        # 读取共享字符串表
        student_shared_xml = self._read_xml_from_zip(self.student_file, 'xl/sharedStrings.xml')
        student_shared_strings = self._parse_shared_strings(student_shared_xml) if student_shared_xml else []
        
        answer_shared_xml = self._read_xml_from_zip(self.answer_file, 'xl/sharedStrings.xml')
        answer_shared_strings = self._parse_shared_strings(answer_shared_xml) if answer_shared_xml else []

        # 提取答案数据
        answer_cells = self._extract_all_cells(answer_sheet_xml, answer_shared_strings)
        answer_rows = self._extract_rows_from_cells(answer_cells)
        
        # 提取学生数据
        student_cells = self._extract_all_cells(student_sheet_xml, student_shared_strings)
        student_rows = self._extract_rows_from_cells(student_cells)

        # 检查行数是否一致
        if len(answer_rows) != len(student_rows):
            return False, f"数据行数不匹配：答案{len(answer_rows)}行，学生{len(student_rows)}行"

        # 比对数据顺序（排序后的数据顺序应该一致）
        # 根据sort_keys指定的列来比较
        
        # 解析sort_keys获取排序列名
        # 格式: ["城区:asc", "建造时间:desc"] -> ["城区", "建造时间"]
        sort_columns = []
        for key in sort_keys:
            col_name = key.split(':')[0] if ':' in key else key
            sort_columns.append(col_name)
        
        if not sort_columns:
            return False, "未指定排序列"
        
        # 获取标题行，建立列名到列字母的映射
        def get_header_mapping(sheet_xml, shared_strings):
            """获取标题行映射 {列名: 列字母}"""
            headers = {}
            # 提取第一行的单元格
            for match in re.finditer(r'<c[^>]*r="([A-Z]+)1"[^>]*>(.*?)</c>', sheet_xml, re.DOTALL):
                col_letter = match.group(1)
                content = match.group(2)
                value = ""
                v_match = re.search(r'<v>(\d+)</v>', content)
                if v_match:
                    idx = int(v_match.group(1))
                    if idx < len(shared_strings):
                        value = shared_strings[idx]
                else:
                    t_match = re.search(r'<t>([^<]*)</t>', content)
                    if t_match:
                        value = t_match.group(1)
                if value:
                    headers[value] = col_letter
            return headers
        
        answer_headers = get_header_mapping(answer_sheet_xml, answer_shared_strings)
        student_headers = get_header_mapping(student_sheet_xml, student_shared_strings)
        
        # 调试信息
        # print(f"DEBUG: 答案文件标题: {answer_headers}")
        # print(f"DEBUG: 学生文件标题: {student_headers}")
        # print(f"DEBUG: 需要查找的排序列: {sort_columns}")
        
        # 获取排序列对应的列字母
        answer_sort_cols = []
        student_sort_cols = []
        for col_name in sort_columns:
            if col_name in answer_headers:
                answer_sort_cols.append(answer_headers[col_name])
            if col_name in student_headers:
                student_sort_cols.append(student_headers[col_name])
        
        # print(f"DEBUG: 答案排序列: {answer_sort_cols}")
        # print(f"DEBUG: 学生排序列: {student_sort_cols}")
        
        if not answer_sort_cols:
            return False, f"标准答案中未找到排序列: {', '.join(sort_columns)}，可用列: {list(answer_headers.keys())}"
        
        if not student_sort_cols:
            return False, f"学生文件中未找到排序列: {', '.join(sort_columns)}，可用列: {list(student_headers.keys())}"
        
        # 只提取排序列的数据进行比较
        def get_sort_key_signature(row, sort_cols):
            """获取排序列的特征签名"""
            values = []
            for col in sort_cols:
                if col in row:
                    values.append(str(row[col]).strip())
            return '|'.join(values) if values else ''
        
        answer_sort_signatures = [get_sort_key_signature(r, answer_sort_cols) for r in answer_rows]
        student_sort_signatures = [get_sort_key_signature(r, student_sort_cols) for r in student_rows]
        
        # 首先检查排序列数据内容是否一致（忽略顺序）
        answer_set = set(answer_sort_signatures)
        student_set = set(student_sort_signatures)
        
        if answer_set != student_set:
            common = answer_set & student_set
            if len(common) == 0:
                return False, f"排序列数据与标准答案不一致，请检查是否使用了正确的数据源"
            content_match_rate = len(common) / len(answer_set) * 100 if answer_set else 0
            if content_match_rate < 80:
                return False, f"排序列数据匹配度仅{content_match_rate:.1f}%，请检查数据源"
        
        # 检查排序顺序是否一致
        match_count = sum(1 for a, s in zip(answer_sort_signatures, student_sort_signatures) if a == s)
        total_rows = len(answer_sort_signatures)
        match_rate = match_count / total_rows * 100 if total_rows > 0 else 0
        
        # 匹配率需要达到90%以上才算通过
        if match_rate >= 90:
            return True, f"排序正确（按{', '.join(sort_columns)}），数据顺序匹配{match_rate:.1f}%"
        else:
            if answer_set == student_set:
                return False, f"排序不正确（按{', '.join(sort_columns)}），数据顺序匹配{match_rate:.1f}%。提示：数据内容正确，但排序顺序与标准答案不一致"
            else:
                return False, f"排序不正确（按{', '.join(sort_columns)}），数据顺序匹配{match_rate:.1f}%"

    def _extract_rows_from_cells(self, cells):
        """从单元格字典中提取行数据"""
        # 找出最大行号
        max_row = 0
        for ref in cells.keys():
            row_match = re.match(r'[A-Z]+(\d+)', ref)
            if row_match:
                row = int(row_match.group(1))
                max_row = max(max_row, row)
        
        # 构建数据行列表（跳过标题行）
        data_rows = []
        for row in range(2, max_row + 1):
            row_data = {}
            for col in range(1, 20):  # 假设最多20列
                col_letter = chr(ord('A') + col - 1)
                ref = f"{col_letter}{row}"
                if ref in cells:
                    row_data[col_letter] = cells[ref]
            if row_data:  # 只添加非空行
                data_rows.append(row_data)
        
        return data_rows

    def _check_subtotal(self, params, answer_sheet_idx, student_sheet_idx):
        """
        检查分类汇总

        Args:
            params: {
                "group_field": "城区",
                "summary_field": "单价",
                "summary_fields": ["单价", "总价", "面积"],  # 可选，多字段汇总
                "summary_method": "average"
            }
            answer_sheet_idx: 答案工作表索引
            student_sheet_idx: 学生工作表索引

        Returns:
            tuple: (passed, message)
        """
        group_field = params.get("group_field", "")
        summary_fields = params.get("summary_fields", [params.get("summary_field", "")])
        summary_method = params.get("summary_method", "average")

        # 获取工作表XML
        answer_sheet_xml = self._get_sheet_xml_by_index(self.answer_file, answer_sheet_idx)
        student_sheet_xml = self._get_sheet_xml_by_index(self.student_file, student_sheet_idx)

        if answer_sheet_xml is None or student_sheet_xml is None:
            return False, "无法读取工作表"

        # 检查是否有SUBTOTAL公式
        # 分类汇总使用SUBTOTAL函数
        student_formulas = self._scan_all_formulas(self.student_file)

        subtotal_formulas = [f for f in student_formulas if 'SUBTOTAL' in f.upper()]

        if subtotal_formulas:
            return True, f"检测到SUBTOTAL公式，共{len(subtotal_formulas)}个"

        # 检查是否有汇总行（通过检查是否有"平均值"或"计"等关键字）
        shared_xml = self._read_xml_from_zip(self.student_file, 'xl/sharedStrings.xml')
        shared_strings = self._parse_shared_strings(shared_xml) if shared_xml else []

        # 查找汇总关键字
        summary_keywords = ["平均值", "计", "汇总", "小计", "总计"]
        found_keywords = []
        for s in shared_strings:
            for kw in summary_keywords:
                if kw in s:
                    found_keywords.append(kw)

        if found_keywords:
            return True, f"检测到汇总关键字: {', '.join(set(found_keywords))}"

        # 检查答案文件是否有SUBTOTAL公式作为参考
        answer_formulas = self._scan_all_formulas(self.answer_file)
        answer_subtotal = [f for f in answer_formulas if 'SUBTOTAL' in f.upper()]

        if not answer_subtotal:
            return False, "标准答案中无SUBTOTAL公式，该检查点不得分"

        return False, "未检测到分类汇总操作"

    def _check_auto_filter(self, params, answer_sheet_idx, student_sheet_idx):
        """
        检查自动筛选 - 只比对筛选后的可见行数据

        Args:
            params: {
                "filter_field": "建造时间",
                "filter_values": ["2020", "2021"],
                "filter_condition": "below_average"  # 可选
            }
            answer_sheet_idx: 答案工作表索引
            student_sheet_idx: 学生工作表索引

        Returns:
            tuple: (passed, message)
        """
        expected_field = params.get("filter_field", "")
        expected_values = params.get("filter_values", [])
        expected_condition = params.get("filter_condition", "")

        # 获取工作表XML
        student_sheet_xml = self._get_sheet_xml_by_index(self.student_file, student_sheet_idx)
        answer_sheet_xml = self._get_sheet_xml_by_index(self.answer_file, answer_sheet_idx)

        if student_sheet_xml is None:
            return False, "无法读取学生工作表"

        if answer_sheet_xml is None:
            return False, "无法读取标准答案工作表"

        # 检查是否存在 autoFilter 标签
        student_has_filter = 'autoFilter' in student_sheet_xml
        answer_has_filter = 'autoFilter' in answer_sheet_xml

        if not answer_has_filter:
            return False, "标准答案中无自动筛选设置，该检查点不得分"

        if not student_has_filter:
            return False, "未检测到自动筛选设置"

        # 读取共享字符串表
        student_shared_xml = self._read_xml_from_zip(self.student_file, 'xl/sharedStrings.xml')
        student_shared_strings = self._parse_shared_strings(student_shared_xml) if student_shared_xml else []

        answer_shared_xml = self._read_xml_from_zip(self.answer_file, 'xl/sharedStrings.xml')
        answer_shared_strings = self._parse_shared_strings(answer_shared_xml) if answer_shared_xml else []

        # 提取答案文件的可见行数据（排除 hidden="1" 的行）
        answer_visible_rows = self._extract_visible_rows(answer_sheet_xml, answer_shared_strings)

        # 提取学生文件的可见行数据（排除 hidden="1" 的行）
        student_visible_rows = self._extract_visible_rows(student_sheet_xml, student_shared_strings)

        # 检查可见行数是否匹配
        answer_visible_count = len(answer_visible_rows)
        student_visible_count = len(student_visible_rows)

        if answer_visible_count == 0:
            return False, "标准答案中无可见数据，该检查点不得分"

        if student_visible_count == 0:
            return False, "筛选后无可见数据"

        row_diff = abs(student_visible_count - answer_visible_count)
        if row_diff > 2:
            return False, f"筛选后可见行数不匹配：答案{answer_visible_count}行，学生{student_visible_count}行"

        # 比对可见行的数据内容
        # 获取答案的关键数据（用于比对）
        answer_key_data = set()
        for row in answer_visible_rows.values():
            # 使用行的关键数据作为指纹
            key = tuple(sorted(row.values()))
            answer_key_data.add(key)

        # 获取学生的关键数据
        student_key_data = set()
        for row in student_visible_rows.values():
            key = tuple(sorted(row.values()))
            student_key_data.add(key)

        # 计算匹配度
        matching_rows = answer_key_data & student_key_data
        missing_rows = answer_key_data - student_key_data
        extra_rows = student_key_data - answer_key_data

        match_rate = len(matching_rows) / len(answer_key_data) * 100 if answer_key_data else 0

        # 如果匹配度低于90%，认为筛选不正确
        if match_rate < 90:
            missing_count = len(missing_rows)
            extra_count = len(extra_rows)
            return False, f"筛选结果与标准答案不匹配：缺少{missing_count}条记录，多余{extra_count}条记录，匹配率{match_rate:.1f}%"

        # 构建成功信息
        condition_info = ""
        if expected_values:
            condition_info = f"，条件: {expected_field}={','.join(expected_values)}"
        elif expected_condition:
            condition_info = f"，条件: {expected_field}{expected_condition}"

        return True, f"检测到自动筛选，可见行{student_visible_count}行，与标准答案匹配{match_rate:.1f}%{condition_info}"

    def _check_advanced_filter(self, params, answer_sheet_idx, student_sheet_idx):
        """
        检查高级筛选 - 比对条件区域的设置

        检查策略：
        1. 在数据区域下方（如30行以后）查找条件区域
        2. 条件区域格式：第一行是字段名，下面行是条件值
        3. 同一行 = AND，不同行 = OR
        4. 提取条件区域的内容进行比对
        5. 检查条件内容是否与期望一致

        Args:
            params: {
                "filter_conditions": [...],
                "filter_logic": "and"  # 可选
            }
            answer_sheet_idx: 答案工作表索引
            student_sheet_idx: 学生工作表索引

        Returns:
            tuple: (passed, message)
        """
        filter_conditions = params.get("filter_conditions", [])
        filter_logic = params.get("filter_logic", "and")

        student_sheet_xml = self._get_sheet_xml_by_index(self.student_file, student_sheet_idx)
        answer_sheet_xml = self._get_sheet_xml_by_index(self.answer_file, answer_sheet_idx)

        if student_sheet_xml is None:
            return False, "无法读取学生工作表"

        if answer_sheet_xml is None:
            return False, "无法读取标准答案工作表"

        # 读取共享字符串表
        student_shared_xml = self._read_xml_from_zip(self.student_file, 'xl/sharedStrings.xml')
        student_shared_strings = self._parse_shared_strings(student_shared_xml) if student_shared_xml else []

        answer_shared_xml = self._read_xml_from_zip(self.answer_file, 'xl/sharedStrings.xml')
        answer_shared_strings = self._parse_shared_strings(answer_shared_xml) if answer_shared_xml else []

        # 提取答案文件的条件区域（30行以后）
        answer_condition_area = self._extract_condition_area(answer_sheet_xml, answer_shared_strings, start_row=30)

        # 提取学生文件的条件区域
        student_condition_area = self._extract_condition_area(student_sheet_xml, student_shared_strings, start_row=30)

        # 检查是否找到条件区域
        if not answer_condition_area["headers"]:
            return False, "标准答案中无条件区域，该检查点不得分"

        if not student_condition_area["headers"]:
            return False, "未检测到条件区域"

        # 比对条件区域的字段名（headers）
        answer_headers = set(answer_condition_area["headers"])
        student_headers = set(student_condition_area["headers"])

        # 比对条件值
        answer_conditions = answer_condition_area["conditions"]
        student_conditions = student_condition_area["conditions"]

        # 构建条件字典，便于比对
        def build_condition_dict(headers, conditions):
            """将条件列表转换为字典列表"""
            result = []
            for cond_row in conditions:
                cond_dict = {}
                for i, header in enumerate(headers):
                    if i < len(cond_row) and cond_row[i]:
                        cond_dict[header] = cond_row[i]
                if cond_dict:
                    result.append(cond_dict)
            return result

        answer_cond_dicts = build_condition_dict(answer_condition_area["headers"], answer_conditions)
        student_cond_dicts = build_condition_dict(student_condition_area["headers"], student_conditions)

        # 检查条件数量
        answer_cond_count = len(answer_cond_dicts)
        student_cond_count = len(student_cond_dicts)

        if answer_cond_count == 0:
            return False, "标准答案中无条件设置，该检查点不得分"

        if student_cond_count == 0:
            return False, "未检测到筛选条件"

        # 比对条件内容
        # 将条件转换为可比较的字符串形式
        def normalize_condition(cond_dict):
            """将条件字典标准化为字符串"""
            items = []
            for key in sorted(cond_dict.keys()):
                value = str(cond_dict[key]).strip()
                items.append(f"{key}={value}")
            return "&".join(items)

        answer_cond_set = set(normalize_condition(c) for c in answer_cond_dicts)
        student_cond_set = set(normalize_condition(c) for c in student_cond_dicts)

        # 计算匹配度
        matching_conds = answer_cond_set & student_cond_set
        missing_conds = answer_cond_set - student_cond_set
        extra_conds = student_cond_set - answer_cond_set

        match_rate = len(matching_conds) / len(answer_cond_set) * 100 if answer_cond_set else 0

        # 如果匹配度低于80%，认为条件设置不正确
        if match_rate < 80:
            missing_count = len(missing_conds)
            extra_count = len(extra_conds)
            return False, f"筛选条件与标准答案不匹配：缺少{missing_count}个条件，多余{extra_count}个条件，匹配率{match_rate:.1f}%"

        # 构建成功信息
        condition_desc = []
        for cond in filter_conditions:
            field = cond.get("field", "")
            value = cond.get("value", "")
            operator = cond.get("operator", "=")
            condition_desc.append(f"{field}{operator}{value}")

        logic_desc = " 且 " if filter_logic == "and" else " 或 "
        condition_info = f"，条件: {logic_desc.join(condition_desc)}"

        return True, f"检测到高级筛选条件区域，条件匹配率{match_rate:.1f}%{condition_info}"

    def _extract_rows_data(self, sheet_xml, shared_strings):
        """
        从工作表XML中提取所有行的数据
        
        Args:
            sheet_xml: 工作表XML内容
            shared_strings: 共享字符串列表
        
        Returns:
            dict: {row_num: {col_letter: value}}
        """
        rows_data = {}
        for match in re.finditer(r'<c[^>]*r="([A-Z]+)(\d+)"[^>]*>(.*?)</c>', sheet_xml, re.DOTALL):
            col_letter = match.group(1)
            row_num = int(match.group(2))
            cell_content = match.group(3)
            
            # 提取值
            value = ""
            v_match = re.search(r'<v>(\d+)</v>', cell_content)
            if v_match:
                idx = int(v_match.group(1))
                if idx < len(shared_strings):
                    value = shared_strings[idx]
                else:
                    value = v_match.group(1)
            else:
                t_match = re.search(r'<t>([^<]*)</t>', cell_content)
                if t_match:
                    value = t_match.group(1)
            
            if row_num not in rows_data:
                rows_data[row_num] = {}
            rows_data[row_num][col_letter] = value
        
        return rows_data

    def _extract_visible_rows(self, sheet_xml, shared_strings):
        """
        从工作表XML中提取可见行的数据（排除 hidden="1" 的行）

        Args:
            sheet_xml: 工作表XML内容
            shared_strings: 共享字符串列表

        Returns:
            dict: {row_num: {col_letter: value}}
        """
        rows_data = {}
        if sheet_xml is None:
            return rows_data

        # 匹配所有行，检查是否有 hidden="1" 属性
        for row_match in re.finditer(r'<row[^>]*>(.*?)</row>', sheet_xml, re.DOTALL):
            row_tag = row_match.group(0)
            row_content = row_match.group(1)

            # 检查行是否被隐藏
            hidden_match = re.search(r'hidden="(\d+)"', row_tag)
            if hidden_match and hidden_match.group(1) == "1":
                continue  # 跳过隐藏行

            # 获取行号
            row_num_match = re.search(r'r="(\d+)"', row_tag)
            if row_num_match:
                row_num = int(row_num_match.group(1))
            else:
                continue

            # 提取该行中的单元格数据
            row_data = {}
            for cell_match in re.finditer(r'<c[^>]*r="([A-Z]+)(\d+)"[^>]*>(.*?)</c>', row_content, re.DOTALL):
                col_letter = cell_match.group(1)
                cell_content = cell_match.group(3)

                # 提取值
                value = ""
                # 检查是否为共享字符串
                if 't="s"' in row_tag[:row_tag.find('>', row_tag.find('<c'))] or 't="s"' in cell_match.group(0):
                    v_match = re.search(r'<v>(\d+)</v>', cell_content)
                    if v_match:
                        idx = int(v_match.group(1))
                        if idx < len(shared_strings):
                            value = shared_strings[idx]
                        else:
                            value = v_match.group(1)
                else:
                    # 直接值
                    v_match = re.search(r'<v>([^<]*)</v>', cell_content)
                    if v_match:
                        value = v_match.group(1)

                # 检查内联文本
                t_match = re.search(r'<t>([^<]*)</t>', cell_content)
                if t_match:
                    value = t_match.group(1)

                if value:
                    row_data[col_letter] = value

            if row_data:
                rows_data[row_num] = row_data

        return rows_data

    def _extract_condition_area(self, sheet_xml, shared_strings, start_row=30):
        """
        从工作表XML中提取条件区域的数据（通常在数据下方）

        Args:
            sheet_xml: 工作表XML内容
            shared_strings: 共享字符串列表
            start_row: 开始查找条件区域的行号（默认30）

        Returns:
            dict: {"headers": [...], "conditions": [[...], ...]}
        """
        result = {"headers": [], "conditions": []}
        if sheet_xml is None:
            return result

        # 提取所有行的数据
        all_rows = {}
        for row_match in re.finditer(r'<row[^>]*>(.*?)</row>', sheet_xml, re.DOTALL):
            row_tag = row_match.group(0)
            row_content = row_match.group(1)

            # 获取行号
            row_num_match = re.search(r'r="(\d+)"', row_tag)
            if row_num_match:
                row_num = int(row_num_match.group(1))
            else:
                continue

            # 提取该行中的单元格数据
            row_data = {}
            for cell_match in re.finditer(r'<c[^>]*r="([A-Z]+)(\d+)"[^>]*>(.*?)</c>', row_content, re.DOTALL):
                col_letter = cell_match.group(1)
                cell_content = cell_match.group(3)

                # 提取值
                value = ""
                # 检查是否为共享字符串
                cell_full = cell_match.group(0)
                if 't="s"' in cell_full:
                    v_match = re.search(r'<v>(\d+)</v>', cell_content)
                    if v_match:
                        idx = int(v_match.group(1))
                        if idx < len(shared_strings):
                            value = shared_strings[idx]
                        else:
                            value = v_match.group(1)
                else:
                    # 直接值
                    v_match = re.search(r'<v>([^<]*)</v>', cell_content)
                    if v_match:
                        value = v_match.group(1)

                # 检查内联文本
                t_match = re.search(r'<t>([^<]*)</t>', cell_content)
                if t_match:
                    value = t_match.group(1)

                if value:
                    row_data[col_letter] = value

            if row_data:
                all_rows[row_num] = row_data

        # 查找条件区域（从start_row开始，找到第一个非空行作为条件区域开始）
        condition_start_row = None
        for row_num in sorted(all_rows.keys()):
            if row_num >= start_row:
                condition_start_row = row_num
                break

        if condition_start_row is None:
            return result

        # 第一行是字段名（headers）
        first_row = all_rows[condition_start_row]
        # 按列字母排序
        sorted_cols = sorted(first_row.keys())
        result["headers"] = [first_row[col] for col in sorted_cols]

        # 后续行是条件值
        condition_rows = []
        for row_num in sorted(all_rows.keys()):
            if row_num > condition_start_row:
                row_data = all_rows[row_num]
                # 按列字母排序，与headers对应
                condition_values = []
                for col in sorted_cols:
                    condition_values.append(row_data.get(col, ""))
                # 只添加非空条件行
                if any(v for v in condition_values):
                    condition_rows.append(condition_values)

        result["conditions"] = condition_rows
        return result

    def _check_pivot_table(self, params, answer_sheet_idx, student_sheet_idx):
        """
        检查数据透视表 - 严格比对答案和学生文件的透视表配置

        Args:
            params: {
                "row_field": "城区",
                "column_field": "楼层位置",
                "data_field": "单价",
                "summary_method": "average"
            }
            answer_sheet_idx: 答案工作表索引
            student_sheet_idx: 学生工作表索引

        Returns:
            tuple: (passed, message)
        """
        row_field = params.get("row_field", "")
        column_field = params.get("column_field", "")
        data_field = params.get("data_field", "")
        summary_method = params.get("summary_method", "average")

        # 首先检查答案文件是否有透视表
        answer_pivot_files = []
        try:
            with zipfile.ZipFile(self.answer_file, 'r') as zf:
                answer_pivot_files = [n for n in zf.namelist() if 'pivotTable' in n]
                if not answer_pivot_files:
                    return False, "标准答案中无数据透视表，该检查点不得分"
        except Exception:
            return False, "无法读取标准答案文件"

        # 检查学生文件是否有透视表
        student_pivot_files = []
        try:
            with zipfile.ZipFile(self.student_file, 'r') as zf:
                student_pivot_files = [n for n in zf.namelist() if 'pivotTable' in n]
                if not student_pivot_files:
                    return False, "未检测到数据透视表"
        except Exception:
            return False, "无法读取学生文件"

        # 获取答案文件的透视表配置
        answer_config = self._extract_pivot_config(self.answer_file)
        student_config = self._extract_pivot_config(self.student_file)

        if not answer_config:
            return False, "标准答案中无有效的数据透视表配置"

        if not student_config:
            return False, "学生文件中无有效的数据透视表配置"

        # 比对透视表配置
        errors = []

        # 检查行字段
        if row_field and answer_config.get('row_fields'):
            if set(answer_config['row_fields']) != set(student_config.get('row_fields', [])):
                errors.append(f"行字段不匹配：答案{answer_config['row_fields']}，学生{student_config.get('row_fields', [])}")

        # 检查列字段
        if column_field and answer_config.get('col_fields'):
            if set(answer_config['col_fields']) != set(student_config.get('col_fields', [])):
                errors.append(f"列字段不匹配：答案{answer_config['col_fields']}，学生{student_config.get('col_fields', [])}")

        # 检查数据字段
        if data_field and answer_config.get('data_fields'):
            if set(answer_config['data_fields']) != set(student_config.get('data_fields', [])):
                errors.append(f"数据字段不匹配：答案{answer_config['data_fields']}，学生{student_config.get('data_fields', [])}")

        # 检查汇总方式
        if summary_method and answer_config.get('summary_methods'):
            if set(answer_config['summary_methods']) != set(student_config.get('summary_methods', [])):
                errors.append(f"汇总方式不匹配：答案{answer_config['summary_methods']}，学生{student_config.get('summary_methods', [])}")

        if errors:
            return False, "数据透视表配置不正确：" + "；".join(errors)

        return True, "数据透视表配置正确"

    def _extract_pivot_config(self, file_path):
        """提取透视表配置 - 严格区分行列字段"""
        try:
            with zipfile.ZipFile(file_path, 'r') as zf:
                # 读取共享字符串获取字段名映射
                shared_strings = []
                try:
                    shared_xml = zf.read('xl/sharedStrings.xml').decode('utf-8')
                    shared_strings = re.findall(r'<t>([^<]*)</t>', shared_xml)
                except:
                    pass

                # 读取透视表定义
                pivot_files = [n for n in zf.namelist() if 'pivotTable' in n and n.endswith('.xml')]
                if not pivot_files:
                    return None

                config = {
                    'row_fields': [],
                    'col_fields': [],
                    'data_fields': [],
                    'summary_methods': []
                }

                for pivot_file in pivot_files:
                    pivot_xml = zf.read(pivot_file).decode('utf-8')

                    # 提取行字段（严格从rowFields标签提取）
                    row_match = re.search(r'<rowFields[^>]*>(.*?)</rowFields>', pivot_xml, re.DOTALL)
                    if row_match:
                        row_xml = row_match.group(1)
                        row_indices = re.findall(r'<field[^>]*x="(\d+)"', row_xml)
                        for idx in row_indices:
                            fld_idx = int(idx)
                            if fld_idx < len(shared_strings):
                                field_name = shared_strings[fld_idx]
                                if field_name not in config['row_fields']:
                                    config['row_fields'].append(field_name)

                    # 提取列字段（严格从colFields标签提取）
                    col_match = re.search(r'<colFields[^>]*>(.*?)</colFields>', pivot_xml, re.DOTALL)
                    if col_match:
                        col_xml = col_match.group(1)
                        col_indices = re.findall(r'<field[^>]*x="(\d+)"', col_xml)
                        for idx in col_indices:
                            fld_idx = int(idx)
                            if fld_idx < len(shared_strings):
                                field_name = shared_strings[fld_idx]
                                if field_name not in config['col_fields']:
                                    config['col_fields'].append(field_name)

                    # 提取数据字段和汇总方式（从dataFields标签提取）
                    data_match = re.search(r'<dataFields[^>]*>(.*?)</dataFields>', pivot_xml, re.DOTALL)
                    if data_match:
                        data_xml = data_match.group(1)
                        # 提取fld和subtotal
                        data_field_matches = re.findall(r'<dataField[^>]*fld="(\d+)"[^>]*subtotal="([^"]*)"', data_xml)
                        for fld_idx, subtotal in data_field_matches:
                            fld_idx = int(fld_idx)
                            # 从共享字符串获取字段名
                            if fld_idx < len(shared_strings):
                                field_name = shared_strings[fld_idx]
                                if field_name not in config['data_fields']:
                                    config['data_fields'].append(field_name)
                            # 汇总方式
                            subtotal_map = {
                                'average': '平均值', 'count': '计数', 'max': '最大值',
                                'min': '最小值', 'product': '乘积', 'sum': '总和'
                            }
                            config['summary_methods'].append(subtotal_map.get(subtotal.lower(), subtotal))

                return config
        except Exception as e:
            return None

    def _check_pivot_table_row(self, params, answer_sheet_idx, student_sheet_idx):
        """检查透视表行字段"""
        row_field = params.get("row_field", "")
        answer_config = self._extract_pivot_config(self.answer_file)
        student_config = self._extract_pivot_config(self.student_file)

        if not answer_config or not student_config:
            return False, "无法读取透视表配置"

        if row_field in student_config.get('row_fields', []):
            return True, f"行字段正确：{row_field}"
        else:
            return False, f"行字段不正确，应为：{row_field}，实际：{student_config.get('row_fields', [])}"

    def _check_pivot_table_col(self, params, answer_sheet_idx, student_sheet_idx):
        """检查透视表列字段"""
        column_field = params.get("column_field", "")
        answer_config = self._extract_pivot_config(self.answer_file)
        student_config = self._extract_pivot_config(self.student_file)

        if not answer_config or not student_config:
            return False, "无法读取透视表配置"

        if column_field in student_config.get('col_fields', []):
            return True, f"列字段正确：{column_field}"
        else:
            return False, f"列字段不正确，应为：{column_field}，实际：{student_config.get('col_fields', [])}"

    def _check_pivot_table_data(self, params, answer_sheet_idx, student_sheet_idx):
        """检查透视表数据字段"""
        data_field = params.get("data_field", "")
        answer_config = self._extract_pivot_config(self.answer_file)
        student_config = self._extract_pivot_config(self.student_file)

        if not answer_config or not student_config:
            return False, "无法读取透视表配置"

        if data_field in student_config.get('data_fields', []):
            return True, f"数据字段正确：{data_field}"
        else:
            return False, f"数据字段不正确，应为：{data_field}，实际：{student_config.get('data_fields', [])}"

    def _check_pivot_table_summary(self, params, answer_sheet_idx, student_sheet_idx):
        """检查透视表汇总方式"""
        summary_method = params.get("summary_method", "")
        answer_config = self._extract_pivot_config(self.answer_file)
        student_config = self._extract_pivot_config(self.student_file)

        if not answer_config or not student_config:
            return False, "无法读取透视表配置"

        if summary_method in student_config.get('summary_methods', []):
            return True, f"汇总方式正确：{summary_method}"
        else:
            return False, f"汇总方式不正确，应为：{summary_method}，实际：{student_config.get('summary_methods', [])}"

    def _check_print_area(self, params, answer_sheet_idx, student_sheet_idx):
        """
        检查打印区域设置

        Args:
            params: {} (不再使用，从标准答案提取)
            answer_sheet_idx: 答案工作表索引
            student_sheet_idx: 学生工作表索引

        Returns:
            tuple: (passed, message)
        """
        # 从标准答案提取打印区域
        answer_workbook = self._read_xml_from_zip(self.answer_file, 'xl/workbook.xml')
        answer_area = ""
        if answer_workbook:
            match = re.search(
                r'<definedName[^>]*name="[^"]*Print_Area[^"]*"[^>]*>([^<]+)</definedName>',
                answer_workbook
            )
            if match:
                answer_area = match.group(1)
        
        # 从学生文件提取打印区域
        student_workbook = self._read_xml_from_zip(self.student_file, 'xl/workbook.xml')
        student_area = ""
        if student_workbook:
            match = re.search(
                r'<definedName[^>]*name="[^"]*Print_Area[^"]*"[^>]*>([^<]+)</definedName>',
                student_workbook
            )
            if match:
                student_area = match.group(1)
        
        # 提取区域范围进行比对（使用字符串替换处理$）
        # 格式: 房源信息统计表!$A$1:$P$28
        answer_clean = answer_area.replace('$', '')
        student_clean = student_area.replace('$', '')
        
        # 提取A1:P28部分
        answer_ref = re.search(r'([A-Z]+\d+):([A-Z]+\d+)', answer_clean)
        student_ref = re.search(r'([A-Z]+\d+):([A-Z]+\d+)', student_clean)
        
        # 如果标准答案中没有设置打印区域，该检查点不得分
        if not answer_ref:
            return False, "标准答案中未设置打印区域，该检查点不得分"
        
        # 比对区域范围
        if answer_ref and student_ref:
            answer_range = f"{answer_ref.group(1)}:{answer_ref.group(2)}"
            student_range = f"{student_ref.group(1)}:{student_ref.group(2)}"
            if answer_range == student_range:
                return True, f"打印区域正确: {student_range}"
            else:
                return False, f"打印区域不匹配: 答案{answer_range}, 学生{student_range}"
        
        return False, "未设置打印区域"

    def _check_page_margins(self, params, answer_sheet_idx, student_sheet_idx):
        """
        检查页边距设置

        Args:
            params: {} (不再使用，从标准答案提取)
            answer_sheet_idx: 答案工作表索引
            student_sheet_idx: 学生工作表索引

        Returns:
            tuple: (passed, message)
        """
        # 从标准答案提取页边距设置
        answer_sheet_xml = self._get_sheet_xml_by_index(self.answer_file, answer_sheet_idx)
        answer_margins = {}
        answer_horizontal_centered = False
        
        if answer_sheet_xml:
            margins_match = re.search(r'<pageMargins\s+([^>]+)/>', answer_sheet_xml)
            if margins_match:
                margins_attrs = margins_match.group(1)
                # 提取页边距值
                left_match = re.search(r'left="([^"]*)"', margins_attrs)
                right_match = re.search(r'right="([^"]*)"', margins_attrs)
                top_match = re.search(r'top="([^"]*)"', margins_attrs)
                bottom_match = re.search(r'bottom="([^"]*)"', margins_attrs)
                hc_match = re.search(r'horizontalCentered="([^"]*)"', margins_attrs)
                
                if left_match:
                    answer_margins['left'] = left_match.group(1)
                if right_match:
                    answer_margins['right'] = right_match.group(1)
                if top_match:
                    answer_margins['top'] = top_match.group(1)
                if bottom_match:
                    answer_margins['bottom'] = bottom_match.group(1)
                if hc_match:
                    answer_horizontal_centered = hc_match.group(1) in ('1', 'true')
        
        # 从学生文件提取页边距设置
        student_sheet_xml = self._get_sheet_xml_by_index(self.student_file, student_sheet_idx)
        if student_sheet_xml is None:
            return False, "无法读取学生工作表"
        
        student_margins = {}
        student_horizontal_centered = False
        
        margins_match = re.search(r'<pageMargins\s+([^>]+)/>', student_sheet_xml)
        if margins_match:
            margins_attrs = margins_match.group(1)
            # 提取页边距值
            left_match = re.search(r'left="([^"]*)"', margins_attrs)
            right_match = re.search(r'right="([^"]*)"', margins_attrs)
            top_match = re.search(r'top="([^"]*)"', margins_attrs)
            bottom_match = re.search(r'bottom="([^"]*)"', margins_attrs)
            hc_match = re.search(r'horizontalCentered="([^"]*)"', margins_attrs)
            
            if left_match:
                student_margins['left'] = left_match.group(1)
            if right_match:
                student_margins['right'] = right_match.group(1)
            if top_match:
                student_margins['top'] = top_match.group(1)
            if bottom_match:
                student_margins['bottom'] = bottom_match.group(1)
            if hc_match:
                student_horizontal_centered = hc_match.group(1) in ('1', 'true')
        
        # 如果标准答案中没有设置页边距，该检查点不得分
        if not answer_margins and not answer_horizontal_centered:
            return False, "标准答案中未设置页边距，该检查点不得分"
        
        # 比对页边距设置
        checks_passed = []
        checks_failed = []
        
        # 比对各个页边距值
        for key in ['left', 'right', 'top', 'bottom']:
            if key in answer_margins:
                if key in student_margins:
                    if student_margins[key] == answer_margins[key]:
                        checks_passed.append(f"{key}={student_margins[key]}")
                    else:
                        checks_failed.append(f"{key}应为{answer_margins[key]}, 实际为{student_margins[key]}")
                else:
                    checks_failed.append(f"未设置{key}页边距")
        
        # 比对水平居中设置
        if answer_horizontal_centered:
            if student_horizontal_centered:
                checks_passed.append("水平居中已设置")
            else:
                checks_failed.append("未设置水平居中")
        
        if checks_passed and not checks_failed:
            return True, f"页边距设置正确: {', '.join(checks_passed)}"
        elif checks_failed:
            return False, f"页边距设置不匹配: {', '.join(checks_failed)}"
        
        return False, "未检测到页边距设置"

    def _check_print_title_rows(self, params, answer_sheet_idx, student_sheet_idx):
        """
        检查顶端标题行设置

        Args:
            params: {} (不再使用，从标准答案提取)
            answer_sheet_idx: 答案工作表索引
            student_sheet_idx: 学生工作表索引

        Returns:
            tuple: (passed, message)
        """
        # 从标准答案提取打印标题行
        answer_workbook = self._read_xml_from_zip(self.answer_file, 'xl/workbook.xml')
        answer_titles = ""
        if answer_workbook:
            match = re.search(
                r'<definedName[^>]*name="[^"]*Print_Titles[^"]*"[^>]*>([^<]+)</definedName>',
                answer_workbook
            )
            if match:
                answer_titles = match.group(1)
        
        # 从学生文件提取打印标题行
        student_workbook = self._read_xml_from_zip(self.student_file, 'xl/workbook.xml')
        student_titles = ""
        if student_workbook:
            match = re.search(
                r'<definedName[^>]*name="[^"]*Print_Titles[^"]*"[^>]*>([^<]+)</definedName>',
                student_workbook
            )
            if match:
                student_titles = match.group(1)
        
        # 提取行范围进行比对（使用字符串替换处理$）
        # 格式: Sheet1!$1:$3 或 '房源信息统计表'!$1:$3
        answer_clean = answer_titles.replace('$', '')
        student_clean = student_titles.replace('$', '')
        answer_row_ref = re.search(r'!(\d+):(\d+)', answer_clean)
        student_row_ref = re.search(r'!(\d+):(\d+)', student_clean)
        
        # 如果标准答案中没有设置打印标题行，该检查点不得分
        if not answer_row_ref:
            return False, "标准答案中未设置顶端标题行，该检查点不得分"
        
        # 比对行范围
        if answer_row_ref and student_row_ref:
            answer_range = f"{answer_row_ref.group(1)}:{answer_row_ref.group(2)}"
            student_range = f"{student_row_ref.group(1)}:{student_row_ref.group(2)}"
            if answer_range == student_range:
                return True, f"顶端标题行正确: {student_range}"
            else:
                return False, f"顶端标题行不匹配: 答案{answer_range}, 学生{student_range}"
        
        return False, "未设置顶端标题行"

    def _check_page_orientation(self, params, answer_sheet_idx, student_sheet_idx):
        """
        检查页面方向设置

        Args:
            params: {} (不再使用，从标准答案提取)
            answer_sheet_idx: 答案工作表索引
            student_sheet_idx: 学生工作表索引

        Returns:
            tuple: (passed, message)
        """
        # 从标准答案提取页面方向设置
        answer_sheet_xml = self._get_sheet_xml_by_index(self.answer_file, answer_sheet_idx)
        answer_orientation = None
        answer_fit_to_height = None
        
        if answer_sheet_xml:
            # 提取pageSetup属性
            page_setup_match = re.search(r'<pageSetup\s+([^>]*)/?>', answer_sheet_xml)
            if page_setup_match:
                page_setup_attrs = page_setup_match.group(1)
                # 提取orientation
                orientation_match = re.search(r'orientation="([^"]*)"', page_setup_attrs)
                if orientation_match:
                    answer_orientation = orientation_match.group(1)
                # 提取fitToHeight
                fit_to_height_match = re.search(r'fitToHeight="([^"]*)"', page_setup_attrs)
                if fit_to_height_match:
                    answer_fit_to_height = fit_to_height_match.group(1)
        
        # 从学生文件提取页面方向设置
        student_sheet_xml = self._get_sheet_xml_by_index(self.student_file, student_sheet_idx)
        if student_sheet_xml is None:
            return False, "无法读取学生工作表"
        
        student_orientation = None
        student_fit_to_height = None
        
        page_setup_match = re.search(r'<pageSetup\s+([^>]*)/?>', student_sheet_xml)
        if page_setup_match:
            page_setup_attrs = page_setup_match.group(1)
            # 提取orientation
            orientation_match = re.search(r'orientation="([^"]*)"', page_setup_attrs)
            if orientation_match:
                student_orientation = orientation_match.group(1)
            # 提取fitToHeight
            fit_to_height_match = re.search(r'fitToHeight="([^"]*)"', page_setup_attrs)
            if fit_to_height_match:
                student_fit_to_height = fit_to_height_match.group(1)
        
        # 如果标准答案中没有设置页面方向，该检查点不得分
        if answer_orientation is None:
            return False, "标准答案中未设置页面方向，该检查点不得分"
        
        # 比对页面方向设置
        checks_passed = []
        checks_failed = []
        
        # 比对orientation
        if student_orientation:
            if student_orientation == answer_orientation:
                orientation_name = "横向" if answer_orientation == "landscape" else "纵向"
                checks_passed.append(f"页面方向: {orientation_name}")
            else:
                answer_name = "横向" if answer_orientation == "landscape" else "纵向"
                student_name = "横向" if student_orientation == "landscape" else "纵向"
                checks_failed.append(f"页面方向应为{answer_name}，实际为{student_name}")
        else:
            checks_failed.append("未设置页面方向")
        
        # 比对fitToHeight (0表示所有列打印在一页)
        if answer_fit_to_height is not None:
            if student_fit_to_height is not None:
                if student_fit_to_height == answer_fit_to_height:
                    if answer_fit_to_height == "0":
                        checks_passed.append("已设置所有列打印在一页")
                    else:
                        checks_passed.append(f"fitToHeight={student_fit_to_height}")
                else:
                    checks_failed.append(f"fitToHeight应为{answer_fit_to_height}，实际为{student_fit_to_height}")
            else:
                checks_failed.append("未设置fitToHeight")
        
        if checks_passed and not checks_failed:
            return True, f"页面设置正确: {', '.join(checks_passed)}"
        elif checks_failed:
            return False, f"页面设置不匹配: {', '.join(checks_failed)}"
        
        return False, "未检测到页面方向设置"

    def _check_print_preview(self, params, answer_sheet_idx, student_sheet_idx):
        """
        检查打印预览（此操作无法直接检测，返回成功）

        Args:
            params: 参数字典
            answer_sheet_idx: 答案工作表索引
            student_sheet_idx: 学生工作表索引

        Returns:
            tuple: (passed, message)
        """
        # 打印预览是一个查看操作，无法在文件中直接检测
        # 只要前面的页面设置检查通过，就认为打印预览操作已完成
        return True, "打印预览为查看操作，无法直接检测，默认通过"

    def _get_row_count_from_dimension(self, dimension):
        """
        从维度字符串中提取行数

        Args:
            dimension: 维度字符串，如 "A1:P28"

        Returns:
            int: 行数
        """
        try:
            # 提取结束行号
            match = re.search(r':\$?[A-Z]+\$?(\d+)$', dimension)
            if match:
                return int(match.group(1))
            # 如果只有单个单元格
            match = re.search(r'\$?[A-Z]+\$?(\d+)$', dimension)
            if match:
                return int(match.group(1))
        except Exception:
            pass
        return 0
