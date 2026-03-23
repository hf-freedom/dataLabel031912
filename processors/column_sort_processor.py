# -*- coding: utf-8 -*-
"""
按指定列排序处理器
将指定列从指定行到指定行进行排序
"""

from typing import Tuple, Dict, Any
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from .base import BaseProcessor


class ColumnSortProcessor(BaseProcessor):
    """
    按指定列排序处理器
    
    功能：将指定列从指定行到指定行进行排序（正序/反序）
    """
    
    def __init__(self):
        super().__init__()
        self.name = "按指定列排序"
        self.description = "将指定列从指定行到指定行进行排序"
        self.params = {
            'sort_col': {
                'label': '排序列',
                'type': 'col',
                'required': True,
                'default': 1,
                'hint': '按该列的值排序'
            },
            'start_row': {
                'label': '开始行',
                'type': 'row',
                'required': True,
                'default': 2,
                'hint': '排序起始行'
            },
            'end_row': {
                'label': '结束行',
                'type': 'row',
                'required': True,
                'default': 10,
                'hint': '排序结束行'
            },
            'sort_order': {
                'label': '排序方式',
                'type': 'text',
                'required': True,
                'default': '正序',
                'hint': '正序或反序'
            }
        }
    
    def get_display_text(self, param_values: Dict[str, Any] = None) -> str:
        if param_values:
            sort_col = param_values.get('sort_col', '')
            start_row = param_values.get('start_row', '')
            end_row = param_values.get('end_row', '')
            sort_order = param_values.get('sort_order', '')
            if sort_col and start_row and end_row and sort_order:
                order_text = "升序" if sort_order == "正序" else "降序"
                return f"将第{sort_col}列从第{start_row}行到第{end_row}行{order_text}排序"
        return self.description
    
    def validate_params(self, param_values: Dict[str, Any]) -> Tuple[bool, str]:
        sort_col = param_values.get('sort_col', '')
        start_row = param_values.get('start_row', '')
        end_row = param_values.get('end_row', '')
        sort_order = param_values.get('sort_order', '')
        
        if not sort_col:
            return False, "请输入排序列"
        if not start_row:
            return False, "请输入开始行"
        if not end_row:
            return False, "请输入结束行"
        if not sort_order:
            return False, "请输入排序方式"
        
        try:
            start = int(start_row)
            end = int(end_row)
            if start > end:
                return False, "开始行不能大于结束行"
            if start < 1:
                return False, "开始行必须大于0"
        except ValueError:
            return False, "行号必须是数字"
        
        if sort_order not in ['正序', '反序']:
            return False, "排序方式必须是'正序'或'反序'"
        
        return True, ""
    
    def process(self, df: pd.DataFrame, wb: Workbook, sheet_name: str,
                param_values: Dict[str, Any]) -> Tuple[pd.DataFrame, Workbook]:
        if df.empty:
            return df, wb
        
        sort_col = self.get_param_value(param_values, 'sort_col', 1)
        start_row = self.get_param_value(param_values, 'start_row', 2)
        end_row = self.get_param_value(param_values, 'end_row', 10)
        sort_order = self.get_param_value(param_values, 'sort_order', '正序')
        
        ascending = (sort_order == '正序')
        
        df_start_idx = start_row - 2
        df_end_idx = end_row - 2
        
        if df_start_idx < 0:
            df_start_idx = 0
        if df_end_idx >= len(df):
            df_end_idx = len(df) - 1
        
        if df_start_idx > df_end_idx:
            return df, wb
        
        sort_col_idx = sort_col - 1
        
        before_rows = df.iloc[:df_start_idx].copy() if df_start_idx > 0 else pd.DataFrame()
        sort_rows = df.iloc[df_start_idx:df_end_idx + 1].copy()
        after_rows = df.iloc[df_end_idx + 1:].copy() if df_end_idx + 1 < len(df) else pd.DataFrame()
        
        if not sort_rows.empty and sort_col_idx < len(sort_rows.columns):
            sort_rows = sort_rows.sort_values(by=sort_rows.columns[sort_col_idx], ascending=ascending)
            sort_rows = sort_rows.reset_index(drop=True)
        
        if not before_rows.empty:
            if not sort_rows.empty:
                if not after_rows.empty:
                    df_result = pd.concat([before_rows, sort_rows, after_rows], ignore_index=True)
                else:
                    df_result = pd.concat([before_rows, sort_rows], ignore_index=True)
            else:
                df_result = before_rows
        else:
            if not sort_rows.empty:
                if not after_rows.empty:
                    df_result = pd.concat([sort_rows, after_rows], ignore_index=True)
                else:
                    df_result = sort_rows
            else:
                df_result = after_rows if not after_rows.empty else df
        
        ws = wb[sheet_name]
        
        for row_idx in range(start_row, end_row + 1):
            for col_idx in range(1, df.shape[1] + 1):
                df_row_idx = row_idx - 2
                if df_row_idx < len(df_result):
                    cell_value = df_result.iloc[df_row_idx, col_idx - 1]
                    ws.cell(row=row_idx, column=col_idx, value=cell_value)
        
        return df_result, wb
