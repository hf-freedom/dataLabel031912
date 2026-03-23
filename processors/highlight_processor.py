# -*- coding: utf-8 -*-
"""
条件格式标记超出阈值处理器
将指定范围内大于/小于阈值的数据标记为红色
"""

from typing import Tuple, Dict, Any
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter
from .base import BaseProcessor


class HighlightProcessor(BaseProcessor):
    """
    条件格式标记超出阈值处理器
    
    功能：将指定范围内大于/小于阈值的数据标记为红色
    使用Excel条件格式实现，修改单元格值后会自动变色
    """
    
    def __init__(self):
        """初始化处理器"""
        super().__init__()
        self.name = "条件格式标记超出阈值"
        self.description = "将指定范围内大于/小于阈值的数据标记为红色（使用条件格式，修改值后自动变色）"
        self.params = {
            'start_row': {
                'label': '开始行',
                'type': 'row',
                'required': True,
                'default': 2,
                'hint': '标记起始行（第1行是表头）'
            },
            'end_row': {
                'label': '结束行',
                'type': 'row',
                'required': True,
                'default': '',
                'hint': '标记结束行'
            },
            'condition': {
                'label': '条件',
                'type': 'text',
                'required': True,
                'default': '大于',
                'hint': '大于或小于'
            },
            'threshold': {
                'label': '阈值',
                'type': 'number',
                'required': True,
                'default': 100,
                'hint': '阈值数值'
            },
            'target_col': {
                'label': '目标列',
                'type': 'col',
                'required': False,
                'default': '',
                'hint': '要标记的列（为空则标记所有列）'
            }
        }
    
    def get_display_text(self, param_values: Dict[str, Any] = None) -> str:
        """获取带参数的显示文本"""
        if param_values:
            start_row = param_values.get('start_row', '')
            end_row = param_values.get('end_row', '')
            condition = param_values.get('condition', '')
            threshold = param_values.get('threshold', '')
            target_col = param_values.get('target_col', '')
            
            if start_row and end_row and condition and threshold:
                if target_col:
                    return f"将第{target_col}列从第{start_row}行到第{end_row}行中{condition}{threshold}的数据标记为红色"
                else:
                    return f"将从第{start_row}行到第{end_row}行中{condition}{threshold}的数据标记为红色"
        return self.description
    
    def validate_params(self, param_values: Dict[str, Any]) -> Tuple[bool, str]:
        """验证参数"""
        start_row = param_values.get('start_row', '')
        end_row = param_values.get('end_row', '')
        condition = param_values.get('condition', '')
        threshold = param_values.get('threshold', '')
        
        if not start_row:
            return False, "请输入开始行"
        if not end_row:
            return False, "请输入结束行"
        if not condition:
            return False, "请输入条件（大于或小于）"
        if threshold == '':
            return False, "请输入阈值"
        
        try:
            start_row = int(start_row)
            end_row = int(end_row)
            float(threshold)
        except ValueError:
            return False, "开始行、结束行必须是数字，阈值必须是数值"
        
        if start_row < 2:
            return False, "开始行必须大于等于2（第1行是表头）"
        if end_row < start_row:
            return False, "结束行必须大于等于开始行"
        
        condition_str = str(condition).strip()
        if condition_str not in ['大于', '小于']:
            return False, "条件必须是'大于'或'小于'"
        
        target_col = param_values.get('target_col', '')
        if target_col and str(target_col).strip():
            try:
                int(target_col)
            except ValueError:
                return False, "目标列必须是数字"
        
        return True, ""
    
    def process(self, df: pd.DataFrame, wb: Workbook, sheet_name: str,
                param_values: Dict[str, Any]) -> Tuple[pd.DataFrame, Workbook]:
        """执行条件格式标记 - 使用Excel条件格式，修改值后自动变色"""
        if df.empty:
            return df, wb

        start_row = int(self.get_param_value(param_values, 'start_row', 2))
        end_row = int(self.get_param_value(param_values, 'end_row', 2))
        condition = str(self.get_param_value(param_values, 'condition', '大于')).strip()
        threshold = float(self.get_param_value(param_values, 'threshold', 100))
        target_col_str = param_values.get('target_col', '')

        ws = wb[sheet_name]
        max_col = df.shape[1]

        # 确定要处理的列范围
        if target_col_str and str(target_col_str).strip():
            target_col = int(target_col_str)
            col_range = range(target_col, target_col + 1)
        else:
            col_range = range(1, max_col + 1)

        # 创建红色填充样式
        red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')

        # 为每一列添加条件格式规则
        for col_idx in col_range:
            col_letter = get_column_letter(col_idx)
            
            # 构建单元格范围，例如 "D2:D11"
            cell_range = "{}{}:{}{}".format(col_letter, start_row, col_letter, end_row)
            
            # 创建条件格式规则
            if condition == '大于':
                rule = CellIsRule(operator='greaterThan', formula=[threshold], fill=red_fill)
            else:  # 小于
                rule = CellIsRule(operator='lessThan', formula=[threshold], fill=red_fill)
            
            # 应用条件格式到整个范围
            ws.conditional_formatting.add(cell_range, rule)
            
            print("Applied conditional formatting to {} (condition: {} threshold: {})".format(
                cell_range, condition, threshold
            ))

        return df, wb
