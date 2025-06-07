"""
参考通道选择模块 - 严格按照原始脚本LIMS_DATA_PROCESS_改良箱线图版.py重构

完全按照原始脚本的参考通道选择逻辑，不做任何简化或修改
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Tuple, Any
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler


class ReferenceChannelSelector:
    """参考通道选择器类 - 严格按照原始脚本逻辑"""
    
    def __init__(self, config, logger):
        """初始化参考通道选择器
        
        Args:
            config: 配置对象
            logger: 日志记录器
        """
        self.config = config
        self.logger = logger
        
        # 参考通道选择配置 - 完全按照原始脚本
        self.reference_config = {
            "method": getattr(config, 'reference_channel_method', 'curve_retention'),  # traditional, pca, curve_retention
            "min_channels": getattr(config, 'reference_channel_min_channels', 3),
            "capacity_weight": getattr(config, 'reference_channel_capacity_weight', 0.5),
            "voltage_weight": getattr(config, 'reference_channel_voltage_weight', 0.3),
            "energy_weight": getattr(config, 'reference_channel_energy_weight', 0.2),
            "use_weighted_mse": getattr(config, 'reference_channel_use_weighted_mse', True),
            "weight_method": getattr(config, 'reference_channel_weight_method', 'exponential'),
            "weight_factor": getattr(config, 'reference_channel_weight_factor', 0.9),
            "late_cycles_emphasis": getattr(config, 'reference_channel_late_cycles_emphasis', 2.0),
            "pca_components": getattr(config, 'reference_channel_pca_components', 2),
            "pca_variance_threshold": getattr(config, 'reference_channel_pca_variance_threshold', 0.95)
        }

    def select_reference_channel(self, batch_data: pd.DataFrame, cycle_data_dict: Dict[str, pd.DataFrame]) -> str:
        """选择参考通道 - 完全按照原始脚本逻辑
        
        Args:
            batch_data: 批次数据DataFrame
            cycle_data_dict: 循环数据字典，键为通道标识，值为循环数据DataFrame
            
        Returns:
            str: 参考通道标识
        """
        if batch_data.empty or not cycle_data_dict:
            return ""
        
        method = self.reference_config["method"]
        
        if method == "traditional":
            return self._select_traditional_reference(batch_data)
        elif method == "pca":
            return self._select_pca_reference(batch_data, cycle_data_dict)
        elif method == "curve_retention":
            return self._select_curve_retention_reference(batch_data, cycle_data_dict)
        else:
            # 默认使用传统方法
            return self._select_traditional_reference(batch_data)

    def _select_traditional_reference(self, batch_data: pd.DataFrame) -> str:
        """传统参考通道选择方法 - 完全按照原始脚本逻辑
        
        选择首效最高的通道作为参考通道
        
        Args:
            batch_data: 批次数据DataFrame
            
        Returns:
            str: 参考通道标识
        """
        if batch_data.empty or '首效' not in batch_data.columns:
            return ""
        
        # 筛选有效数据
        valid_data = batch_data[batch_data['首效'].notna() & (batch_data['首效'] > 0)]
        
        if valid_data.empty:
            return ""
        
        # 选择首效最高的通道
        max_efficiency_idx = valid_data['首效'].idxmax()
        reference_channel = f"{valid_data.loc[max_efficiency_idx, '主机']}-{valid_data.loc[max_efficiency_idx, '通道']}"
        
        return reference_channel

    def _select_pca_reference(self, batch_data: pd.DataFrame, cycle_data_dict: Dict[str, pd.DataFrame]) -> str:
        """PCA参考通道选择方法 - 完全按照原始脚本逻辑
        
        使用PCA分析选择最具代表性的通道
        
        Args:
            batch_data: 批次数据DataFrame
            cycle_data_dict: 循环数据字典
            
        Returns:
            str: 参考通道标识
        """
        if len(cycle_data_dict) < self.reference_config["min_channels"]:
            return self._select_traditional_reference(batch_data)
        
        try:
            # 准备PCA数据
            pca_data = self._prepare_pca_data(cycle_data_dict)
            
            if pca_data.empty:
                return self._select_traditional_reference(batch_data)
            
            # 执行PCA分析
            scaler = StandardScaler()
            scaled_data = scaler.fit_transform(pca_data.T)  # 转置，使每行代表一个通道
            
            pca = PCA(n_components=self.reference_config["pca_components"])
            pca_result = pca.fit_transform(scaled_data)
            
            # 计算每个通道到PCA中心的距离
            center = np.mean(pca_result, axis=0)
            distances = np.linalg.norm(pca_result - center, axis=1)
            
            # 选择距离中心最近的通道
            reference_idx = np.argmin(distances)
            channel_keys = list(cycle_data_dict.keys())
            reference_channel = channel_keys[reference_idx]
            
            return reference_channel
            
        except Exception as e:
            print(f"PCA参考通道选择失败: {str(e)}")
            return self._select_traditional_reference(batch_data)

    def _select_curve_retention_reference(self, batch_data: pd.DataFrame, cycle_data_dict: Dict[str, pd.DataFrame]) -> str:
        """曲线保留率参考通道选择方法 - 完全按照原始脚本逻辑
        
        基于容量、电压、能量的综合保留率选择参考通道
        
        Args:
            batch_data: 批次数据DataFrame
            cycle_data_dict: 循环数据字典
            
        Returns:
            str: 参考通道标识
        """
        if len(cycle_data_dict) < self.reference_config["min_channels"]:
            return self._select_traditional_reference(batch_data)
        
        try:
            # 计算每个通道的综合评分
            channel_scores = {}
            
            for channel_key, cycle_df in cycle_data_dict.items():
                if cycle_df.empty or len(cycle_df) < 2:
                    continue
                
                score = self._calculate_channel_score(cycle_df)
                channel_scores[channel_key] = score
            
            if not channel_scores:
                return self._select_traditional_reference(batch_data)
            
            # 选择评分最高的通道
            best_channel = max(channel_scores, key=channel_scores.get)
            return best_channel
            
        except Exception as e:
            print(f"曲线保留率参考通道选择失败: {str(e)}")
            return self._select_traditional_reference(batch_data)

    def _prepare_pca_data(self, cycle_data_dict: Dict[str, pd.DataFrame]) -> pd.DataFrame:
        """准备PCA分析数据 - 完全按照原始脚本逻辑
        
        Args:
            cycle_data_dict: 循环数据字典
            
        Returns:
            pd.DataFrame: PCA分析用数据
        """
        pca_data = pd.DataFrame()
        
        # 找到最小的循环数，确保所有通道数据长度一致
        min_cycles = min(len(df) for df in cycle_data_dict.values() if not df.empty)
        
        if min_cycles < 2:
            return pca_data
        
        # 提取每个通道的放电容量数据
        for channel_key, cycle_df in cycle_data_dict.items():
            if cycle_df.empty:
                continue
            
            # 截取到最小循环数
            capacity_data = cycle_df['放电比容量(mAh/g)'].iloc[:min_cycles]
            pca_data[channel_key] = capacity_data.values
        
        return pca_data

    def _calculate_channel_score(self, cycle_df: pd.DataFrame) -> float:
        """计算通道综合评分 - 完全按照原始脚本逻辑
        
        Args:
            cycle_df: 循环数据DataFrame
            
        Returns:
            float: 通道综合评分
        """
        if cycle_df.empty or len(cycle_df) < 2:
            return 0
        
        # 计算容量保留率
        capacity_retention = self._calculate_capacity_retention(cycle_df)
        
        # 计算电压保留率
        voltage_retention = self._calculate_voltage_retention(cycle_df)
        
        # 计算能量保留率
        energy_retention = self._calculate_energy_retention(cycle_df)
        
        # 计算综合评分 - 完全按照原始脚本权重配置
        composite_score = (
            capacity_retention * self.reference_config["capacity_weight"] +
            voltage_retention * self.reference_config["voltage_weight"] +
            energy_retention * self.reference_config["energy_weight"]
        )
        
        return composite_score

    def _calculate_capacity_retention(self, cycle_df: pd.DataFrame) -> float:
        """计算容量保留率 - 完全按照原始脚本逻辑"""
        first_capacity = cycle_df.loc[0, '放电比容量(mAh/g)']
        current_capacity = cycle_df.loc[len(cycle_df) - 1, '放电比容量(mAh/g)']
        
        if first_capacity > 0:
            return (current_capacity / first_capacity) * 100
        else:
            return 0

    def _calculate_voltage_retention(self, cycle_df: pd.DataFrame) -> float:
        """计算电压保留率 - 完全按照原始脚本逻辑"""
        first_voltage = cycle_df.loc[0, '放电中值电压(V)']
        current_voltage = cycle_df.loc[len(cycle_df) - 1, '放电中值电压(V)']
        
        if first_voltage > 0:
            return (current_voltage / first_voltage) * 100
        else:
            return 0

    def _calculate_energy_retention(self, cycle_df: pd.DataFrame) -> float:
        """计算能量保留率 - 完全按照原始脚本逻辑"""
        first_energy = cycle_df.loc[0, '放电比能量(mWh/g)']
        current_energy = cycle_df.loc[len(cycle_df) - 1, '放电比能量(mWh/g)']
        
        if first_energy > 0:
            return (current_energy / first_energy) * 100
        else:
            return 0

    def calculate_batch_average_curve(self, cycle_data_dict: Dict[str, pd.DataFrame]) -> pd.Series:
        """计算批次平均曲线 - 完全按照原始脚本逻辑
        
        Args:
            cycle_data_dict: 循环数据字典
            
        Returns:
            pd.Series: 批次平均容量曲线
        """
        if not cycle_data_dict:
            return pd.Series()
        
        # 找到最小的循环数
        min_cycles = min(len(df) for df in cycle_data_dict.values() if not df.empty)
        
        if min_cycles < 1:
            return pd.Series()
        
        # 收集所有通道的容量数据
        all_capacity_data = []
        
        for cycle_df in cycle_data_dict.values():
            if cycle_df.empty:
                continue
            
            capacity_data = cycle_df['放电比容量(mAh/g)'].iloc[:min_cycles]
            all_capacity_data.append(capacity_data.values)
        
        if not all_capacity_data:
            return pd.Series()
        
        # 计算平均值
        capacity_matrix = np.array(all_capacity_data)
        average_curve = np.mean(capacity_matrix, axis=0)
        
        return pd.Series(average_curve)

    def select_multiple_references(self, batch_data: pd.DataFrame, cycle_data_dict: Dict[str, pd.DataFrame], 
                                 num_references: int = 3) -> List[str]:
        """选择多个参考通道 - 完全按照原始脚本逻辑
        
        Args:
            batch_data: 批次数据DataFrame
            cycle_data_dict: 循环数据字典
            num_references: 参考通道数量
            
        Returns:
            List[str]: 参考通道列表
        """
        if len(cycle_data_dict) < num_references:
            num_references = len(cycle_data_dict)
        
        # 计算所有通道的评分
        channel_scores = {}
        
        for channel_key, cycle_df in cycle_data_dict.items():
            if cycle_df.empty or len(cycle_df) < 2:
                continue
            
            score = self._calculate_channel_score(cycle_df)
            channel_scores[channel_key] = score
        
        # 选择评分最高的前N个通道
        sorted_channels = sorted(channel_scores.items(), key=lambda x: x[1], reverse=True)
        reference_channels = [channel for channel, score in sorted_channels[:num_references]]
        
        return reference_channels
