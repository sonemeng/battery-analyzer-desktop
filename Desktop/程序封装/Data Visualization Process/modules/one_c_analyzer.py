"""
1C分析模块 - 严格按照原始脚本LIMS_DATA_PROCESS_改良箱线图版.py重构

完全按照原始脚本的1C识别和分析逻辑，不做任何简化或修改
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Tuple, Any


class OneCAnalyzer:
    """1C分析器类 - 严格按照原始脚本逻辑"""
    
    def __init__(self, config, logger):
        """初始化1C分析器
        
        Args:
            config: 配置对象
            logger: 日志记录器
        """
        self.config = config
        self.logger = logger
        
        # 1C阈值配置 - 完全按照原始脚本
        self.one_c_thresholds = {
            "ratio_threshold": config.one_c_ratio_threshold,
            "discharge_diff_threshold": config.one_c_discharge_diff_threshold,
            "overcharge_threshold": config.one_c_overcharge_threshold,
            "very_low_efficiency_threshold": config.one_c_very_low_efficiency_threshold,
            "low_efficiency_threshold": config.one_c_low_efficiency_threshold
        }
        
        # 测试模式配置 - 完全按照原始脚本
        self.one_c_modes = ['-1C-']
        self.non_one_c_modes = ['-0.1C-', '-BL-', '-0.33C-']
        
        # 默认1C循环数 - 完全按照原始脚本
        self.default_one_c_cycle = 3

    def analyze_one_c_data(self, cycle_df: pd.DataFrame, file_info: Dict[str, Any]) -> List[Any]:
        """分析1C数据 - 完全按照原始脚本逻辑
        
        Args:
            cycle_df: 循环数据DataFrame
            file_info: 文件信息字典
            
        Returns:
            List[Any]: [1C首圈编号, 1C首充, 1C首放, 1C首效, 1C状态, 1C倍率比]
        """
        mode = file_info.get('mode', '')
        
        if mode in self.one_c_modes:
            # 1C模式处理
            return self._process_one_c_mode(cycle_df)
        else:
            # 非1C模式处理
            return self._process_non_one_c_mode()

    def _process_one_c_mode(self, cycle_df: pd.DataFrame) -> List[Any]:
        """处理1C模式 - 完全按照原始脚本逻辑
        
        Args:
            cycle_df: 循环数据DataFrame
            
        Returns:
            List[Any]: 1C相关数据
        """
        # 查找1C首圈编号
        one_c_cycle_num = self._find_one_c_cycle(cycle_df)
        
        if one_c_cycle_num > 0:
            # 找到有效的1C循环
            one_c_charge = cycle_df.loc[one_c_cycle_num - 1, '充电比容量(mAh/g)']
            one_c_discharge = cycle_df.loc[one_c_cycle_num - 1, '放电比容量(mAh/g)']
            one_c_efficiency = (one_c_discharge / one_c_charge * 100) if one_c_charge > 0 else 0
            one_c_status = self._determine_one_c_status(one_c_charge, one_c_discharge, one_c_efficiency)
            one_c_ratio = self._calculate_one_c_ratio(cycle_df, one_c_cycle_num)
            
            return [one_c_cycle_num, one_c_charge, one_c_discharge, one_c_efficiency, one_c_status, one_c_ratio]
        else:
            # 未找到有效的1C循环，使用默认值
            if len(cycle_df) >= self.default_one_c_cycle:
                default_charge = cycle_df.loc[self.default_one_c_cycle - 1, '充电比容量(mAh/g)']
                default_discharge = cycle_df.loc[self.default_one_c_cycle - 1, '放电比容量(mAh/g)']
                default_efficiency = (default_discharge / default_charge * 100) if default_charge > 0 else 0
                default_status = self._determine_one_c_status(default_charge, default_discharge, default_efficiency)
                default_ratio = self._calculate_one_c_ratio(cycle_df, self.default_one_c_cycle)
                
                return [self.default_one_c_cycle, default_charge, default_discharge, default_efficiency, default_status, default_ratio]
            else:
                return ['', '', '', '', '无1C', '']

    def _process_non_one_c_mode(self) -> List[Any]:
        """处理非1C模式 - 完全按照原始脚本逻辑
        
        Returns:
            List[Any]: 非1C模式的默认值
        """
        return ['', '', '', '', '非1C', '']

    def _find_one_c_cycle(self, cycle_df: pd.DataFrame) -> int:
        """查找1C首圈编号 - 完全按照原始脚本逻辑
        
        Args:
            cycle_df: 循环数据DataFrame
            
        Returns:
            int: 1C首圈编号（1基索引），如果未找到返回0
        """
        # 从第3圈开始查找（索引2）- 完全按照原始脚本
        for i in range(2, len(cycle_df)):
            charge = cycle_df.loc[i, '充电比容量(mAh/g)']
            discharge = cycle_df.loc[i, '放电比容量(mAh/g)']
            
            # 检查是否满足1C条件
            if self._is_valid_one_c_cycle(charge, discharge, cycle_df, i):
                return i + 1  # 返回1基索引
        
        return 0  # 未找到

    def _is_valid_one_c_cycle(self, charge: float, discharge: float, cycle_df: pd.DataFrame, cycle_index: int) -> bool:
        """检查是否为有效的1C循环 - 完全按照原始脚本逻辑
        
        Args:
            charge: 充电容量
            discharge: 放电容量
            cycle_df: 循环数据DataFrame
            cycle_index: 循环索引（0基）
            
        Returns:
            bool: 是否为有效的1C循环
        """
        # 1. 检查过充 - 完全按照原始脚本
        if charge > self.one_c_thresholds["overcharge_threshold"]:
            return False
        
        # 2. 检查效率 - 完全按照原始脚本
        efficiency = (discharge / charge * 100) if charge > 0 else 0
        if efficiency < self.one_c_thresholds["very_low_efficiency_threshold"]:
            return False
        
        # 3. 检查与首圈的放电容量差异 - 完全按照原始脚本
        first_discharge = cycle_df.loc[0, '放电比容量(mAh/g)']
        discharge_diff = abs(discharge - first_discharge)
        if discharge_diff > self.one_c_thresholds["discharge_diff_threshold"]:
            return False
        
        # 4. 检查倍率比 - 完全按照原始脚本
        ratio = discharge / first_discharge if first_discharge > 0 else 0
        if ratio > self.one_c_thresholds["ratio_threshold"]:
            return False
        
        return True

    def _determine_one_c_status(self, charge: float, discharge: float, efficiency: float) -> str:
        """确定1C状态 - 完全按照原始脚本逻辑
        
        Args:
            charge: 充电容量
            discharge: 放电容量
            efficiency: 效率
            
        Returns:
            str: 1C状态
        """
        if charge > self.one_c_thresholds["overcharge_threshold"]:
            return "过充"
        elif efficiency < self.one_c_thresholds["very_low_efficiency_threshold"]:
            return "极低效"
        elif efficiency < self.one_c_thresholds["low_efficiency_threshold"]:
            return "低效"
        else:
            return "正常"

    def _calculate_one_c_ratio(self, cycle_df: pd.DataFrame, one_c_cycle_num: int) -> float:
        """计算1C倍率比 - 完全按照原始脚本逻辑
        
        Args:
            cycle_df: 循环数据DataFrame
            one_c_cycle_num: 1C循环编号（1基索引）
            
        Returns:
            float: 1C倍率比
        """
        if one_c_cycle_num <= 1 or one_c_cycle_num > len(cycle_df):
            return 0
        
        first_discharge = cycle_df.loc[0, '放电比容量(mAh/g)']
        one_c_discharge = cycle_df.loc[one_c_cycle_num - 1, '放电比容量(mAh/g)']
        
        if first_discharge > 0:
            return one_c_discharge / first_discharge
        else:
            return 0

    def is_one_c_mode(self, mode: str) -> bool:
        """判断是否为1C模式 - 完全按照原始脚本逻辑
        
        Args:
            mode: 测试模式
            
        Returns:
            bool: 是否为1C模式
        """
        return mode in self.one_c_modes

    def get_one_c_summary(self, all_data: pd.DataFrame) -> Dict[str, Any]:
        """获取1C数据汇总 - 完全按照原始脚本逻辑
        
        Args:
            all_data: 所有数据DataFrame
            
        Returns:
            Dict[str, Any]: 1C数据汇总
        """
        if all_data.empty:
            return {
                'total_one_c_files': 0,
                'valid_one_c_files': 0,
                'one_c_success_rate': 0,
                'average_one_c_efficiency': 0,
                'average_one_c_ratio': 0
            }
        
        # 筛选1C数据
        one_c_data = all_data[all_data['1C状态'].isin(['正常', '低效', '极低效', '过充'])]
        valid_one_c_data = all_data[all_data['1C状态'] == '正常']
        
        total_one_c = len(one_c_data)
        valid_one_c = len(valid_one_c_data)
        success_rate = (valid_one_c / total_one_c * 100) if total_one_c > 0 else 0
        
        # 计算平均值
        avg_efficiency = valid_one_c_data['1C首效'].mean() if not valid_one_c_data.empty else 0
        avg_ratio = valid_one_c_data['1C倍率比'].mean() if not valid_one_c_data.empty else 0
        
        return {
            'total_one_c_files': total_one_c,
            'valid_one_c_files': valid_one_c,
            'one_c_success_rate': success_rate,
            'average_one_c_efficiency': avg_efficiency,
            'average_one_c_ratio': avg_ratio
        }
