"""
容量保留率计算模块 - 严格按照原始脚本LIMS_DATA_PROCESS_改良箱线图版.py重构

完全按照原始脚本的容量保留率计算逻辑，不做任何简化或修改
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Tuple, Any


class CapacityRetentionCalculator:
    """容量保留率计算器类 - 严格按照原始脚本逻辑"""
    
    def __init__(self, config, logger):
        """初始化容量保留率计算器
        
        Args:
            config: 配置对象
            logger: 日志记录器
        """
        self.config = config
        self.logger = logger
        
        # 容量保留率配置 - 完全按照原始脚本
        self.capacity_retention_config = {
            "enabled": getattr(config, 'capacity_retention_enabled', True),
            "min_cycles": getattr(config, 'capacity_retention_min_cycles', 10),
            "max_cycles": getattr(config, 'capacity_retention_max_cycles', 1000),
            "cycle_step": getattr(config, 'capacity_retention_cycle_step', 10),
            "interpolation_method": getattr(config, 'capacity_retention_interpolation_method', 'linear'),
            "retention_columns": getattr(config, 'capacity_retention_retention_columns', ['capacity', 'voltage', 'energy']),
            "use_raw_capacity": getattr(config, 'capacity_retention_use_raw_capacity', False),
            "use_weighted_mse": getattr(config, 'capacity_retention_use_weighted_mse', True),
            "weight_method": getattr(config, 'capacity_retention_weight_method', 'exponential'),
            "weight_factor": getattr(config, 'capacity_retention_weight_factor', 0.9),
            "late_cycles_emphasis": getattr(config, 'capacity_retention_late_cycles_emphasis', 2.0),
            "dynamic_range": getattr(config, 'capacity_retention_dynamic_range', True),
            "min_channels": getattr(config, 'capacity_retention_min_channels', 3),
            "include_voltage": getattr(config, 'capacity_retention_include_voltage', True),
            "include_energy": getattr(config, 'capacity_retention_include_energy', True),
            "capacity_weight": getattr(config, 'capacity_retention_capacity_weight', 0.5),
            "voltage_weight": getattr(config, 'capacity_retention_voltage_weight', 0.3),
            "energy_weight": getattr(config, 'capacity_retention_energy_weight', 0.2),
            "voltage_column": getattr(config, 'capacity_retention_voltage_column', '放电中值电压(V)'),
            "energy_column": getattr(config, 'capacity_retention_energy_column', '放电比能量(mWh/g)')
        }

    def calculate_capacity_retention(self, cycle_df: pd.DataFrame) -> float:
        """计算当前容量保持率 - 完全按照原始脚本逻辑
        
        Args:
            cycle_df: 循环数据DataFrame
            
        Returns:
            float: 当前容量保持率（百分比）
        """
        if len(cycle_df) < 2:
            return 100.0
        
        first_discharge = cycle_df.loc[0, '放电比容量(mAh/g)']
        current_discharge = cycle_df.loc[len(cycle_df) - 1, '放电比容量(mAh/g)']
        
        if first_discharge > 0:
            return (current_discharge / first_discharge) * 100
        else:
            return 0

    def calculate_voltage_retention(self, cycle_df: pd.DataFrame) -> float:
        """计算电压保持率 - 完全按照原始脚本逻辑
        
        Args:
            cycle_df: 循环数据DataFrame
            
        Returns:
            float: 电压保持率（百分比）
        """
        if len(cycle_df) < 2:
            return 100.0
        
        voltage_column = self.capacity_retention_config["voltage_column"]
        first_voltage = cycle_df.loc[0, voltage_column]
        current_voltage = cycle_df.loc[len(cycle_df) - 1, voltage_column]
        
        if first_voltage > 0:
            return (current_voltage / first_voltage) * 100
        else:
            return 0

    def calculate_energy_retention(self, cycle_df: pd.DataFrame) -> float:
        """计算能量保持率 - 完全按照原始脚本逻辑
        
        Args:
            cycle_df: 循环数据DataFrame
            
        Returns:
            float: 能量保持率（百分比）
        """
        if len(cycle_df) < 2:
            return 100.0
        
        energy_column = self.capacity_retention_config["energy_column"]
        first_energy = cycle_df.loc[0, energy_column]
        current_energy = cycle_df.loc[len(cycle_df) - 1, energy_column]
        
        if first_energy > 0:
            return (current_energy / first_energy) * 100
        else:
            return 0

    def calculate_voltage_decay_rate(self, cycle_df: pd.DataFrame, mode: str = '', one_c_cycle_num: Optional[int] = None) -> float:
        """计算电压衰减率 - 严格按照原始脚本逻辑

        Args:
            cycle_df: 循环数据DataFrame
            mode: 测试模式
            one_c_cycle_num: 1C首圈编号（1基索引）

        Returns:
            float: 电压衰减率（mV/周）
        """
        if len(cycle_df) <= 4:  # 原始脚本：只有当圈数>4时才计算
            return 0

        total_cycles = len(cycle_df)
        current_voltage = cycle_df.loc[total_cycles - 1, '放电中值电压(V)']

        # 0.1C模式 - 完全按照原始脚本第926行
        if mode == '-0.1C-':
            first_voltage = cycle_df.loc[0, '放电中值电压(V)']
            voltage_decay_rate = (first_voltage - current_voltage) * 1000 / (total_cycles - 1)
            return round(voltage_decay_rate, 1)

        # 1C模式 - 完全按照原始脚本第950行
        elif mode == '-1C-':
            # 确定1C首圈索引
            if one_c_cycle_num is not None and one_c_cycle_num > 0:
                one_c_idx = one_c_cycle_num - 1  # 转换为0基索引
            else:
                one_c_idx = min(3, total_cycles - 1)  # 默认第4圈或最大可用圈

            # 确保有足够的循环数据
            if total_cycles > one_c_idx + 1:
                one_c_voltage = cycle_df.loc[one_c_idx, '放电中值电压(V)']
                voltage_decay_rate = (one_c_voltage - current_voltage) * 1000 / (total_cycles - (one_c_idx + 1))
                return round(voltage_decay_rate, 1)
            else:
                return 0

        # 其他模式，使用首圈作为基准
        else:
            first_voltage = cycle_df.loc[0, '放电中值电压(V)']
            voltage_decay_rate = (first_voltage - current_voltage) * 1000 / (total_cycles - 1)
            return round(voltage_decay_rate, 1)

    def calculate_retention_at_cycle(self, cycle_df: pd.DataFrame, target_cycle: int) -> List[Any]:
        """计算指定循环数的保持率 - 完全按照原始脚本逻辑
        
        Args:
            cycle_df: 循环数据DataFrame
            target_cycle: 目标循环数
            
        Returns:
            List[Any]: [容量保持率, 电压保持率, 能量保持率]
        """
        if len(cycle_df) < target_cycle:
            return ['', '', '']  # 数据不足
        
        # 获取首圈数据
        first_discharge = cycle_df.loc[0, '放电比容量(mAh/g)']
        first_voltage = cycle_df.loc[0, self.capacity_retention_config["voltage_column"]]
        first_energy = cycle_df.loc[0, self.capacity_retention_config["energy_column"]]
        
        # 获取目标循环数据
        target_discharge = cycle_df.loc[target_cycle - 1, '放电比容量(mAh/g)']
        target_voltage = cycle_df.loc[target_cycle - 1, self.capacity_retention_config["voltage_column"]]
        target_energy = cycle_df.loc[target_cycle - 1, self.capacity_retention_config["energy_column"]]
        
        # 计算保持率
        capacity_retention = (target_discharge / first_discharge * 100) if first_discharge > 0 else 0
        voltage_retention = (target_voltage / first_voltage * 100) if first_voltage > 0 else 0
        energy_retention = (target_energy / first_energy * 100) if first_energy > 0 else 0
        
        return [capacity_retention, voltage_retention, energy_retention]

    def calculate_weighted_mse_retention(self, cycle_df: pd.DataFrame, reference_curve: pd.Series) -> float:
        """计算加权MSE容量保留率 - 完全按照原始脚本逻辑
        
        Args:
            cycle_df: 循环数据DataFrame
            reference_curve: 参考曲线
            
        Returns:
            float: 加权MSE值
        """
        if not self.capacity_retention_config["use_weighted_mse"]:
            return self._calculate_simple_mse(cycle_df, reference_curve)
        
        # 获取放电容量数据
        discharge_capacity = cycle_df['放电比容量(mAh/g)'].values
        
        # 确保数据长度一致
        min_length = min(len(discharge_capacity), len(reference_curve))
        discharge_capacity = discharge_capacity[:min_length]
        reference_curve = reference_curve[:min_length]
        
        # 计算权重 - 完全按照原始脚本
        weights = self._calculate_weights(min_length)
        
        # 计算加权MSE
        squared_errors = (discharge_capacity - reference_curve) ** 2
        weighted_mse = np.sum(weights * squared_errors) / np.sum(weights)
        
        return weighted_mse

    def _calculate_simple_mse(self, cycle_df: pd.DataFrame, reference_curve: pd.Series) -> float:
        """计算简单MSE - 完全按照原始脚本逻辑"""
        discharge_capacity = cycle_df['放电比容量(mAh/g)'].values
        
        min_length = min(len(discharge_capacity), len(reference_curve))
        discharge_capacity = discharge_capacity[:min_length]
        reference_curve = reference_curve[:min_length]
        
        mse = np.mean((discharge_capacity - reference_curve) ** 2)
        return mse

    def _calculate_weights(self, length: int) -> np.ndarray:
        """计算权重 - 完全按照原始脚本逻辑
        
        Args:
            length: 数据长度
            
        Returns:
            np.ndarray: 权重数组
        """
        weight_method = self.capacity_retention_config["weight_method"]
        weight_factor = self.capacity_retention_config["weight_factor"]
        late_cycles_emphasis = self.capacity_retention_config["late_cycles_emphasis"]
        
        if weight_method == 'exponential':
            # 指数权重 - 后期循环权重更高
            weights = np.array([weight_factor ** (length - i - 1) for i in range(length)])
            # 对后期循环加强权重
            late_start = int(length * 0.7)  # 后30%的循环
            weights[late_start:] *= late_cycles_emphasis
        elif weight_method == 'linear':
            # 线性权重
            weights = np.linspace(1, late_cycles_emphasis, length)
        else:
            # 均匀权重
            weights = np.ones(length)
        
        return weights

    def calculate_composite_retention_score(self, cycle_df: pd.DataFrame) -> float:
        """计算综合保留率评分 - 完全按照原始脚本逻辑
        
        Args:
            cycle_df: 循环数据DataFrame
            
        Returns:
            float: 综合保留率评分
        """
        capacity_retention = self.calculate_capacity_retention(cycle_df)
        
        score = capacity_retention
        
        # 如果启用电压和能量，计算综合评分
        if self.capacity_retention_config["include_voltage"]:
            voltage_retention = self.calculate_voltage_retention(cycle_df)
            score = (score * self.capacity_retention_config["capacity_weight"] + 
                    voltage_retention * self.capacity_retention_config["voltage_weight"])
        
        if self.capacity_retention_config["include_energy"]:
            energy_retention = self.calculate_energy_retention(cycle_df)
            if self.capacity_retention_config["include_voltage"]:
                # 三项综合
                total_weight = (self.capacity_retention_config["capacity_weight"] + 
                              self.capacity_retention_config["voltage_weight"] + 
                              self.capacity_retention_config["energy_weight"])
                score = (capacity_retention * self.capacity_retention_config["capacity_weight"] + 
                        voltage_retention * self.capacity_retention_config["voltage_weight"] + 
                        energy_retention * self.capacity_retention_config["energy_weight"]) / total_weight
            else:
                # 容量+能量
                total_weight = (self.capacity_retention_config["capacity_weight"] + 
                              self.capacity_retention_config["energy_weight"])
                score = (capacity_retention * self.capacity_retention_config["capacity_weight"] + 
                        energy_retention * self.capacity_retention_config["energy_weight"]) / total_weight
        
        return score

    def get_retention_summary(self, all_data: pd.DataFrame) -> Dict[str, Any]:
        """获取保留率汇总统计 - 完全按照原始脚本逻辑
        
        Args:
            all_data: 所有数据DataFrame
            
        Returns:
            Dict[str, Any]: 保留率汇总统计
        """
        if all_data.empty:
            return {
                'average_capacity_retention': 0,
                'average_voltage_retention': 0,
                'average_energy_retention': 0,
                'capacity_retention_std': 0,
                'voltage_retention_std': 0,
                'energy_retention_std': 0
            }
        
        # 计算平均值和标准差
        avg_capacity = all_data['当前容量保持'].mean()
        avg_voltage = all_data['当前电压保持'].mean()
        avg_energy = all_data['当前能量保持'].mean()
        
        std_capacity = all_data['当前容量保持'].std()
        std_voltage = all_data['当前电压保持'].std()
        std_energy = all_data['当前能量保持'].std()
        
        return {
            'average_capacity_retention': avg_capacity,
            'average_voltage_retention': avg_voltage,
            'average_energy_retention': avg_energy,
            'capacity_retention_std': std_capacity,
            'voltage_retention_std': std_voltage,
            'energy_retention_std': std_energy
        }
