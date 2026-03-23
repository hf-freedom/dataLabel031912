# -*- coding: utf-8 -*-
"""
Excel处理器模块包
该模块提供了可扩展的Excel处理功能框架
"""

from .base import BaseProcessor
from .sort_processor import SortFirstColumnProcessor
from .count_char_processor import CountCharProcessor
from .sum_processor import SumColumnProcessor
from .column_sort_processor import ColumnSortProcessor
from .highlight_threshold_processor import HighlightThresholdProcessor

__all__ = [
    'BaseProcessor',
    'SortFirstColumnProcessor',
    'CountCharProcessor',
    'SumColumnProcessor',
    'ColumnSortProcessor',
    'HighlightThresholdProcessor'
]
