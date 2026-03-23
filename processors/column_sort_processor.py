# -*- coding: utf-8 -*-
"""
按指定列升序/降序排序处理器
将指定列的数据从指定行到指定行按值排序
"""

import pandas as pd
from openpyxl import Workbook
from .base import BaseProcessor


class ColumnSortProcessor(BaseProcessor):
    """
    按指定列排序处理器
    
    功能：将指定列从指定行到指定行按值排序（正序/反序）
    """
    
    def __init__(self):
        """初始化处理器"""
        super(ColumnSortProcessor, self).__init__()
        self.name = "按指定列升序排序"
        self.description = "将指定列从指定行到指定行按值排序（正序/反序）"
        self.params = {
            'sort_col': {
                'label': '排序列号',
                'type': 'col',
                'required': True,
                'default': 1,
                'hint': '按该列的值进行排序'
            },
            'start_row': {
                'label': '开始行',
                'type': 'row',
                'required': True,
                'default': 2,
                'hint': '从该行开始排序（第1行是表头）'
            },
            'end_row': {
                'label': '结束行',
                'type': 'row',
                'required': False,
                'default': '',
                'hint': '排序到该行（留空则到最后一行）'
            },
            'sort_order': {
                'label': '排序方式',
                'type': 'text',
                'required': False,
                'default': '正序',
                'hint': '正序 或 反序（默认正序）'
            }
        }
    
    def get_display_text(self, param_values=None):
        """获取带参数的显示文本"""
        if param_values:
            sort_col = param_values.get('sort_col', '')
            start_row = param_values.get('start_row', '')
            end_row = param_values.get('end_row', '')
            sort_order = param_values.get('sort_order', '正序')
            
            if sort_col and start_row:
                end_text = end_row if end_row else '最后'
                return "将第{}列从第{}行到第{}行按{}排序".format(sort_col, start_row, end_text, sort_order)
        return self.description
    
    def validate_params(self, param_values):
        """验证参数"""
        sort_col = param_values.get('sort_col', '')
        start_row = param_values.get('start_row', '')
        
        if not sort_col or not str(sort_col).strip():
            return False, "请输入排序列号"
        
        if not start_row or not str(start_row).strip():
            return False, "请输入开始行"
        
        try:
            sort_col = int(sort_col)
            start_row = int(start_row)
        except (ValueError, TypeError):
            return False, "排序列号和开始行必须是数字"
        
        if sort_col < 1:
            return False, "排序列号必须大于0"
        
        if start_row < 2:
            return False, "开始行必须大于等于2（第1行是表头）"
        
        end_row = param_values.get('end_row', '')
        if end_row and str(end_row).strip():
            try:
                end_row = int(end_row)
                if end_row < start_row:
                    return False, "结束行必须大于等于开始行"
            except (ValueError, TypeError):
                return False, "结束行必须是数字"
        
        sort_order = param_values.get('sort_order', '正序')
        if sort_order and sort_order not in ['正序', '反序', '升序', '降序']:
            return False, "排序方式请输入：正序 或 反序"
        
        return True, ""
    
    def process(self, df, wb, sheet_name, param_values):
        """执行列排序"""
        if df.empty:
            return df, wb
        
        sort_col = int(self.get_param_value(param_values, 'sort_col', 1))
        start_row = int(self.get_param_value(param_values, 'start_row', 2))
        end_row = self.get_param_value(param_values, 'end_row', '')
        sort_order = self.get_param_value(param_values, 'sort_order', '正序')
        
        sort_col_idx = sort_col - 1
        df_start_idx = start_row - 2
        
        if end_row and str(end_row).strip():
            end_row = int(end_row)
            df_end_idx = end_row - 2
        else:
            df_end_idx = len(df) - 1
        
        if df_start_idx >= len(df):
            return df, wb
        
        df_end_idx = min(df_end_idx, len(df) - 1)
        if df_end_idx < df_start_idx:
            return df, wb
        
        ascending = sort_order in ['正序', '升序']
        
        sort_range = df.iloc[df_start_idx:df_end_idx + 1].copy()
        sort_range = sort_range.sort_values(
            by=df.columns[sort_col_idx],
            ascending=ascending,
            na_position='last'
        )
        sort_range = sort_range.reset_index(drop=True)
        
        df.iloc[df_start_idx:df_end_idx + 1] = sort_range.values
        
        ws = wb[sheet_name]
        
        excel_end_row = end_row if end_row and str(end_row).strip() else (len(df) + 1)
        
        data = []
        for row in range(start_row, excel_end_row + 1):
            row_data = []
            for col in range(1, df.shape[1] + 1):
                cell = ws.cell(row=row, column=col)
                row_data.append((cell.value, cell.style))
            data.append(row_data)
        
        sort_keys = []
        for i, row_data in enumerate(data):
            key_value = row_data[sort_col_idx][0]
            sort_keys.append((key_value, i))
        
        sort_keys.sort(key=lambda x: (x[0] is None, x[0]), reverse=not ascending)
        
        for new_idx, (_, old_idx) in enumerate(sort_keys):
            target_row = start_row + new_idx
            for col in range(1, df.shape[1] + 1):
                cell_value, cell_style = data[old_idx][col - 1]
                target_cell = ws.cell(row=target_row, column=col)
                target_cell.value = cell_value
                target_cell.style = cell_style
        
        return df, wb
