"""
统计计算模块 - 严格按照原始脚本LIMS_DATA_PROCESS_改良箱线图版.py重构

完全按照原始脚本的统计计算逻辑，不做任何简化或修改
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Tuple, Any


class StatisticsCalculator:
    """统计计算器类 - 严格按照原始脚本逻辑"""
    
    def __init__(self, config, logger):
        """初始化统计计算器
        
        Args:
            config: 配置对象
            logger: 日志记录器
        """
        self.config = config
        self.logger = logger
        
        # 统计列配置 - 完全按照原始脚本
        self.statistics_cols = [
            '系列', '统一批次', '上架时间', '总数据', '首周有效数据', '首充', '首放', '首效',
            '首圈电压', '首圈能量', '活性物质', 'Cycle2', 'Cycle2充电比容量', 'Cycle3', 'Cycle3充电比容量',
            'Cycle4', 'Cycle4充电比容量', 'Cycle5', 'Cycle5充电比容量', 'Cycle6', 'Cycle6充电比容量',
            'Cycle7', 'Cycle7充电比容量', '1C首周有效数据', '1C参考通道', '1C首圈编号', '1C首充', '1C首放',
            '1C首效', '1C倍率比', '1C状态', '参考通道当前圈数', '当前容量保持', '电压衰减率mV/周', '当前电压保持', '当前能量保持'
        ]

    def calculate_batch_statistics(self, batch_data: pd.DataFrame, series_name: str, batch_id: str, shelf_time: str) -> List[Any]:
        """计算批次统计数据 - 完全按照原始脚本逻辑
        
        Args:
            batch_data: 批次数据DataFrame
            series_name: 系列名称
            batch_id: 批次ID
            shelf_time: 上架时间
            
        Returns:
            List[Any]: 统计数据行
        """
        if batch_data.empty:
            return self._create_empty_statistics_row(series_name, batch_id, shelf_time)
        
        # 基础统计
        total_count = len(batch_data)
        valid_first_week_count = self._count_valid_first_week_data(batch_data)
        
        # 计算各项平均值 - 完全按照原始脚本
        stats_row = [
            series_name,  # 系列
            batch_id,  # 统一批次
            shelf_time,  # 上架时间
            total_count,  # 总数据
            valid_first_week_count,  # 首周有效数据
            self._safe_mean(batch_data, '首充'),  # 首充
            self._safe_mean(batch_data, '首放'),  # 首放
            self._safe_mean(batch_data, '首效'),  # 首效
            self._safe_mean(batch_data, '首圈电压'),  # 首圈电压
            self._safe_mean(batch_data, '首圈能量'),  # 首圈能量
            self._safe_mean(batch_data, '活性物质'),  # 活性物质
            self._safe_mean(batch_data, 'Cycle2'),  # Cycle2
            self._safe_mean(batch_data, 'Cycle2充电比容量'),  # Cycle2充电比容量
            self._safe_mean(batch_data, 'Cycle3'),  # Cycle3
            self._safe_mean(batch_data, 'Cycle3充电比容量'),  # Cycle3充电比容量
            self._safe_mean(batch_data, 'Cycle4'),  # Cycle4
            self._safe_mean(batch_data, 'Cycle4充电比容量'),  # Cycle4充电比容量
            self._safe_mean(batch_data, 'Cycle5'),  # Cycle5
            self._safe_mean(batch_data, 'Cycle5充电比容量'),  # Cycle5充电比容量
            self._safe_mean(batch_data, 'Cycle6'),  # Cycle6
            self._safe_mean(batch_data, 'Cycle6充电比容量'),  # Cycle6充电比容量
            self._safe_mean(batch_data, 'Cycle7'),  # Cycle7
            self._safe_mean(batch_data, 'Cycle7充电比容量'),  # Cycle7充电比容量
        ]
        
        # 1C相关统计
        one_c_stats = self._calculate_one_c_statistics(batch_data)
        stats_row.extend(one_c_stats)
        
        # 当前状态统计
        current_stats = self._calculate_current_statistics(batch_data)
        stats_row.extend(current_stats)
        
        return stats_row

    def _count_valid_first_week_data(self, batch_data: pd.DataFrame) -> int:
        """计算首周有效数据数量 - 完全按照原始脚本逻辑
        
        Args:
            batch_data: 批次数据DataFrame
            
        Returns:
            int: 首周有效数据数量
        """
        # 首周有效数据定义：首充、首放、首效都有有效值
        valid_mask = (
            (batch_data['首充'].notna()) & (batch_data['首充'] > 0) &
            (batch_data['首放'].notna()) & (batch_data['首放'] > 0) &
            (batch_data['首效'].notna()) & (batch_data['首效'] > 0)
        )
        return valid_mask.sum()

    def _calculate_one_c_statistics(self, batch_data: pd.DataFrame) -> List[Any]:
        """计算1C相关统计 - 完全按照原始脚本逻辑
        
        Args:
            batch_data: 批次数据DataFrame
            
        Returns:
            List[Any]: 1C统计数据
        """
        # 筛选1C数据
        one_c_data = batch_data[batch_data['1C状态'].isin(['正常', '低效', '极低效', '过充'])]
        valid_one_c_count = len(one_c_data)
        
        # 选择参考通道 - 完全按照原始脚本逻辑
        reference_channel = self._select_reference_channel(one_c_data)
        
        # 1C统计数据
        one_c_stats = [
            valid_one_c_count,  # 1C首周有效数据
            reference_channel,  # 1C参考通道
            self._safe_mean(one_c_data, '1C首圈编号'),  # 1C首圈编号
            self._safe_mean(one_c_data, '1C首充'),  # 1C首充
            self._safe_mean(one_c_data, '1C首放'),  # 1C首放
            self._safe_mean(one_c_data, '1C首效'),  # 1C首效
            self._safe_mean(one_c_data, '1C倍率比'),  # 1C倍率比
            self._get_most_common_one_c_status(one_c_data),  # 1C状态
        ]
        
        return one_c_stats

    def _calculate_current_statistics(self, batch_data: pd.DataFrame) -> List[Any]:
        """计算当前状态统计 - 完全按照原始脚本逻辑
        
        Args:
            batch_data: 批次数据DataFrame
            
        Returns:
            List[Any]: 当前状态统计数据
        """
        current_stats = [
            self._safe_mean(batch_data, '当前圈数'),  # 参考通道当前圈数
            self._safe_mean(batch_data, '当前容量保持'),  # 当前容量保持
            self._safe_mean(batch_data, '电压衰减率mV/周'),  # 电压衰减率mV/周
            self._safe_mean(batch_data, '当前电压保持'),  # 当前电压保持
            self._safe_mean(batch_data, '当前能量保持'),  # 当前能量保持
        ]
        
        return current_stats

    def _select_reference_channel(self, one_c_data: pd.DataFrame) -> str:
        """选择参考通道 - 完全按照原始脚本逻辑
        
        Args:
            one_c_data: 1C数据DataFrame
            
        Returns:
            str: 参考通道标识
        """
        if one_c_data.empty:
            return ""
        
        # 选择1C效率最高的通道作为参考通道 - 完全按照原始脚本
        if '1C首效' in one_c_data.columns:
            max_efficiency_idx = one_c_data['1C首效'].idxmax()
            if pd.notna(max_efficiency_idx):
                reference_channel = f"{one_c_data.loc[max_efficiency_idx, '主机']}-{one_c_data.loc[max_efficiency_idx, '通道']}"
                return reference_channel
        
        return ""

    def _get_most_common_one_c_status(self, one_c_data: pd.DataFrame) -> str:
        """获取最常见的1C状态 - 完全按照原始脚本逻辑
        
        Args:
            one_c_data: 1C数据DataFrame
            
        Returns:
            str: 最常见的1C状态
        """
        if one_c_data.empty or '1C状态' not in one_c_data.columns:
            return ""
        
        status_counts = one_c_data['1C状态'].value_counts()
        if not status_counts.empty:
            return status_counts.index[0]
        
        return ""

    def _safe_mean(self, data: pd.DataFrame, column: str) -> float:
        """安全计算平均值 - 完全按照原始脚本逻辑
        
        Args:
            data: 数据DataFrame
            column: 列名
            
        Returns:
            float: 平均值，如果无法计算则返回0
        """
        if data.empty or column not in data.columns:
            return 0
        
        # 过滤有效数值
        valid_data = data[column].dropna()
        if valid_data.empty:
            return 0
        
        # 尝试转换为数值类型
        try:
            numeric_data = pd.to_numeric(valid_data, errors='coerce').dropna()
            if numeric_data.empty:
                return 0
            return numeric_data.mean()
        except:
            return 0

    def _create_empty_statistics_row(self, series_name: str, batch_id: str, shelf_time: str) -> List[Any]:
        """创建空的统计数据行 - 完全按照原始脚本逻辑
        
        Args:
            series_name: 系列名称
            batch_id: 批次ID
            shelf_time: 上架时间
            
        Returns:
            List[Any]: 空的统计数据行
        """
        # 创建与statistics_cols长度相同的空行
        empty_row = [series_name, batch_id, shelf_time, 0, 0]  # 前5个基础字段
        empty_row.extend([0] * (len(self.statistics_cols) - 5))  # 其余字段填0
        
        return empty_row

    def calculate_overall_statistics(self, all_data: pd.DataFrame) -> Dict[str, Any]:
        """计算总体统计 - 完全按照原始脚本逻辑
        
        Args:
            all_data: 所有数据DataFrame
            
        Returns:
            Dict[str, Any]: 总体统计结果
        """
        if all_data.empty:
            return {
                'total_files': 0,
                'total_series': 0,
                'total_batches': 0,
                'average_first_efficiency': 0,
                'average_capacity_retention': 0,
                'average_voltage_retention': 0,
                'average_energy_retention': 0,
                'one_c_success_rate': 0
            }
        
        # 基础统计
        total_files = len(all_data)
        total_series = all_data['系列'].nunique()
        total_batches = all_data['批次'].nunique()
        
        # 性能统计
        avg_first_efficiency = self._safe_mean(all_data, '首效')
        avg_capacity_retention = self._safe_mean(all_data, '当前容量保持')
        avg_voltage_retention = self._safe_mean(all_data, '当前电压保持')
        avg_energy_retention = self._safe_mean(all_data, '当前能量保持')
        
        # 1C成功率
        one_c_success_rate = self._calculate_one_c_success_rate(all_data)
        
        return {
            'total_files': total_files,
            'total_series': total_series,
            'total_batches': total_batches,
            'average_first_efficiency': avg_first_efficiency,
            'average_capacity_retention': avg_capacity_retention,
            'average_voltage_retention': avg_voltage_retention,
            'average_energy_retention': avg_energy_retention,
            'one_c_success_rate': one_c_success_rate
        }

    def _calculate_one_c_success_rate(self, all_data: pd.DataFrame) -> float:
        """计算1C成功率 - 完全按照原始脚本逻辑
        
        Args:
            all_data: 所有数据DataFrame
            
        Returns:
            float: 1C成功率（百分比）
        """
        if all_data.empty or '1C状态' not in all_data.columns:
            return 0
        
        # 筛选1C相关数据（排除非1C和无1C）
        one_c_data = all_data[~all_data['1C状态'].isin(['非1C', '无1C', ''])]
        if one_c_data.empty:
            return 0
        
        # 计算正常1C的比例
        normal_one_c = one_c_data[one_c_data['1C状态'] == '正常']
        success_rate = (len(normal_one_c) / len(one_c_data)) * 100
        
        return success_rate

    def group_data_by_batch(self, all_data: pd.DataFrame) -> Dict[str, pd.DataFrame]:
        """按批次分组数据 - 完全按照原始脚本逻辑
        
        Args:
            all_data: 所有数据DataFrame
            
        Returns:
            Dict[str, pd.DataFrame]: 按批次分组的数据字典
        """
        if all_data.empty:
            return {}
        
        # 创建统一批次标识 - 完全按照原始脚本
        all_data['统一批次'] = all_data.apply(self._create_unified_batch_id, axis=1)
        
        # 按统一批次分组
        grouped_data = {}
        for batch_id, group_data in all_data.groupby('统一批次'):
            grouped_data[batch_id] = group_data
        
        return grouped_data

    def _create_unified_batch_id(self, row: pd.Series) -> str:
        """创建统一批次ID - 完全按照原始脚本逻辑
        
        Args:
            row: 数据行
            
        Returns:
            str: 统一批次ID
        """
        # 使用系列+批次+上架时间创建统一标识
        series = row.get('系列', '')
        batch = row.get('批次', '')
        shelf_time = row.get('上架时间', '')
        
        return f"{series}-{batch}-{shelf_time}"
