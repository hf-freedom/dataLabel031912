# -*- coding: utf-8 -*-
"""
按指定列升序排序处理器
将指定列从指定行到指定行进行排序
"""

from typing import Tuple, Dict, Any
import pandas as pd
from openpyxl import Workbook
from .base import BaseProcessor


class ColumnSortProcessor(BaseProcessor):
    """
    按指定列升序排序处理器
    
    功能：将指定列从指定行到指定行进行排序（正序/反序）
    """
    
    def __init__(self):
        """初始化处理器"""
        super().__init__()
        self.name = "按指定列排序"
        self.description = "将指定列从指定行到指定行进行排序"
        self.params = {
            'sort_col': {
                'label': '排序列号',
                'type': 'col',
                'required': True,
                'default': 1,
                'hint': '要排序的列号'
            },
            'start_row': {
                'label': '开始行',
                'type': 'row',
                'required': True,
                'default': 2,
                'hint': '排序起始行（第1行是表头）'
            },
            'end_row': {
                'label': '结束行',
                'type': 'row',
                'required': True,
                'default': '',
                'hint': '排序结束行'
            },
            'sort_order': {
                'label': '排序方式',
                'type': 'text',
                'required': True,
                'default': '升序',
                'hint': '升序或降序'
            }
        }
    
    def get_display_text(self, param_values: Dict[str, Any] = None) -> str:
        """获取带参数的显示文本"""
        if param_values:
            sort_col = param_values.get('sort_col', '')
            start_row = param_values.get('start_row', '')
            end_row = param_values.get('end_row', '')
            sort_order = param_values.get('sort_order', '')
            if sort_col and start_row and end_row and sort_order:
                return f"将第{sort_col}列从第{start_row}行到第{end_row}行按{sort_order}排序"
        return self.description
    
    def validate_params(self, param_values: Dict[str, Any]) -> Tuple[bool, str]:
        """验证参数"""
        sort_col = param_values.get('sort_col', '')
        start_row = param_values.get('start_row', '')
        end_row = param_values.get('end_row', '')
        sort_order = param_values.get('sort_order', '')
        
        if not sort_col:
            return False, "请输入排序列号"
        if not start_row:
            return False, "请输入开始行"
        if not end_row:
            return False, "请输入结束行"
        if not sort_order:
            return False, "请输入排序方式（升序或降序）"
        
        try:
            sort_col = int(sort_col)
            start_row = int(start_row)
            end_row = int(end_row)
        except ValueError:
            return False, "排序列号、开始行和结束行必须是数字"
        
        if start_row < 2:
            return False, "开始行必须大于等于2（第1行是表头）"
        if end_row < start_row:
            return False, "结束行必须大于等于开始行"
        
        sort_order_str = str(sort_order).strip()
        if sort_order_str not in ['升序', '降序']:
            return False, "排序方式必须是'升序'或'降序'"
        
        return True, ""
    
    def process(self, df: pd.DataFrame, wb: Workbook, sheet_name: str,
                param_values: Dict[str, Any]) -> Tuple[pd.DataFrame, Workbook]:
        """执行按指定列排序"""
        if df.empty:
            return df, wb
        
        sort_col = int(self.get_param_value(param_values, 'sort_col', 1))
        start_row = int(self.get_param_value(param_values, 'start_row', 2))
        end_row = int(self.get_param_value(param_values, 'end_row', 2))
        sort_order = str(self.get_param_value(param_values, 'sort_order', '升序')).strip()
        
        ascending = sort_order == '升序'
        
        # DataFrame索引从0开始，Excel行号从1开始，第1行是表头，数据从第2行开始
        df_start_idx = start_row - 2
        df_end_idx = end_row - 1
        
        # 获取排序列的DataFrame索引（从0开始）
        sort_col_idx = sort_col - 1
        
        # 分离表头、排序区域和剩余数据
        header_df = df.iloc[:df_start_idx] if df_start_idx > 0 else pd.DataFrame()
        sort_df = df.iloc[df_start_idx:df_end_idx].copy()
        tail_df = df.iloc[df_end_idx:] if df_end_idx < len(df) else pd.DataFrame()
        
        if not sort_df.empty:
            # 按指定列排序
            sort_df = sort_df.sort_values(by=sort_df.columns[sort_col_idx], ascending=ascending)
            # 重置索引以保持连续性
            sort_df = sort_df.reset_index(drop=True)
        
        # 重新组合数据
        if not header_df.empty and not tail_df.empty:
            df = pd.concat([header_df, sort_df, tail_df], ignore_index=True)
        elif not header_df.empty:
            df = pd.concat([header_df, sort_df], ignore_index=True)
        elif not tail_df.empty:
            df = pd.concat([sort_df, tail_df], ignore_index=True)
        else:
            df = sort_df
        
        # 更新openpyxl工作簿
        ws = wb[sheet_name]
        
        # 获取原始数据（不包括表头）
        original_data = []
        for row_idx in range(start_row, end_row + 1):
            row_data = []
            for col_idx in range(1, df.shape[1] + 1):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                row_data.append(cell_value)
            original_data.append(row_data)
        
        # 按指定列排序
        sort_col_idx_zero_based = sort_col - 1
        
        def sort_key(row):
            value = row[sort_col_idx_zero_based]
            # 尝试转换为数字进行比较
            try:
                return (0, float(value))
            except (ValueError, TypeError):
                return (1, str(value) if value is not None else '')
        
        sorted_data = sorted(original_data, key=sort_key, reverse=not ascending)
        
        # 将排序后的数据写回工作表
        for i, row_data in enumerate(sorted_data):
            excel_row_idx = start_row + i
            for j, value in enumerate(row_data):
                ws.cell(row=excel_row_idx, column=j + 1, value=value)
        
        return df, wb
