# -*- coding: utf-8 -*-
"""
列重新编号处理器
将指定列从指定行开始按顺序重新编号
"""

import pandas as pd
from openpyxl import Workbook
from .base import BaseProcessor


class SortFirstColumnProcessor(BaseProcessor):
    """
    列重新编号处理器
    
    功能：将指定列从指定行开始按顺序重新编号（1, 2, 3...）
    """
    
    def __init__(self):
        """初始化处理器"""
        super(SortFirstColumnProcessor, self).__init__()
        self.name = "列重新编号"
        self.description = "将指定列从指定行开始按顺序重新编号"
        self.params = {
            'start_row': {
                'label': '起始行',
                'type': 'row',
                'required': True,
                'default': 2,
                'hint': '从该行开始编号（第1行是表头）'
            },
            'target_col': {
                'label': '目标列',
                'type': 'col',
                'required': True,
                'default': 1,
                'hint': '要重新编号的列'
            }
        }
    
    def get_display_text(self, param_values=None):
        """获取带参数的显示文本"""
        if param_values:
            start_row = param_values.get('start_row', '')
            target_col = param_values.get('target_col', '')
            if start_row and target_col:
                return "将第{}列从第{}行开始重新编号".format(target_col, start_row)
        return self.description
    
    def process(self, df, wb, sheet_name, param_values):
        """执行列重新编号"""
        if df.empty:
            return df, wb
        
        start_row = self.get_param_value(param_values, 'start_row', 2)
        target_col = self.get_param_value(param_values, 'target_col', 1)
        
        df_start_idx = start_row - 2
        
        number = 1
        for i in range(df_start_idx, len(df)):
            df.iloc[i, target_col - 1] = number
            number += 1
        
        ws = wb[sheet_name]
        
        number = 1
        for row_idx in range(start_row, len(df) + 2):
            ws.cell(row=row_idx, column=target_col, value=number)
            number += 1
        
        return df, wb
