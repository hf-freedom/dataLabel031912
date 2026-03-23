# -*- coding: utf-8 -*-
"""
Excel处理器基类模块
定义了所有处理器必须实现的接口和属性
"""

from abc import ABC, abstractmethod
from typing import Tuple, Optional, Dict, Any
import pandas as pd
from openpyxl import Workbook


class BaseProcessor(ABC):
    """
    Excel处理器抽象基类
    
    所有具体的Excel处理功能都需要继承此类。
    每个处理器包含以下核心属性：
    - name: 功能名称（显示在下拉框中）
    - description: 功能描述
    - params: 参数定义（字典，定义需要的参数）
    
    属性:
        name: 功能名称
        description: 功能描述
        params: 参数定义字典
            key: 参数名称
            value: dict包含:
                - label: 参数显示名称
                - type: 参数类型 ('row' 或 'col')
                - required: 是否必填
                - default: 默认值
                - hint: 参数提示
    """
    
    def __init__(self):
        """初始化处理器"""
        self.name = "基础处理器"
        self.description = "基础处理器描述"
        self.params = {}
    
    @abstractmethod
    def process(self, df: pd.DataFrame, wb: Workbook, sheet_name: str,
                param_values: Dict[str, Any]) -> Tuple[pd.DataFrame, Workbook]:
        """
        处理Excel数据的核心方法
        
        参数:
            df: pandas DataFrame对象
            wb: openpyxl Workbook对象
            sheet_name: 工作表名称
            param_values: 参数值字典，key为参数名，value为参数值
        
        返回:
            Tuple[pd.DataFrame, Workbook]: 处理后的DataFrame和Workbook
        """
        pass
    
    def get_display_text(self, param_values: Dict[str, Any] = None) -> str:
        """
        获取带参数的显示文本
        
        参数:
            param_values: 参数值字典
        
        返回:
            str: 显示文本
        """
        return self.description
    
    def validate_params(self, param_values: Dict[str, Any]) -> Tuple[bool, str]:
        """
        验证参数是否有效
        
        参数:
            param_values: 参数值字典
        
        返回:
            Tuple[bool, str]: (是否有效, 错误信息)
        """
        for param_name, param_def in self.params.items():
            if param_def.get('required', False):
                value = param_values.get(param_name)
                if value is None or value == '':
                    return False, f"请输入{param_def.get('label', param_name)}"
                if isinstance(value, str) and not value.strip():
                    return False, f"请输入{param_def.get('label', param_name)}"
        return True, ""
    
    def get_param_value(self, param_values: Dict[str, Any], param_name: str, 
                        default: Any = None) -> Any:
        """
        获取参数值，支持默认值
        
        参数:
            param_values: 参数值字典
            param_name: 参数名称
            default: 默认值
        
        返回:
            Any: 参数值
        """
        value = param_values.get(param_name)
        if value is None or value == '':
            return self.params.get(param_name, {}).get('default', default)
        return value
