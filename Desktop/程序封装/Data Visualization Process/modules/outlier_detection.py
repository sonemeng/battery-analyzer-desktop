#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
异常检测器模块
实现箱线图和Z-score+MAD异常检测方法
"""

import pandas as pd
import numpy as np
from typing import Optional, List, Dict, Any
from scipy import stats

from .config_parser import Config
from utils.logger import ProcessingLogger


class OutlierDetector:
    """异常检测器类 - 实现多种异常检测方法"""

    def __init__(self, config: Config, logger: ProcessingLogger):
        """初始化异常检测器

        Args:
            config: 配置对象
            logger: 日志记录器
        """
        self.config = config
        self.logger = logger

        self.logger.log_debug("OutlierDetector初始化完成")
    
    def detect_and_remove_outliers(self, data: pd.DataFrame) -> Optional[pd.DataFrame]:
        """检测并移除异常值
        
        Args:
            data: 输入数据DataFrame
            
        Returns:
            Optional[pd.DataFrame]: 清理后的数据，如果失败则返回None
        """
        if data.empty:
            self.logger.log_warning("输入数据为空，跳过异常检测")
            return data
        
        method = self.config.outlier_method
        self.logger.log_info(f"开始异常检测，使用方法: {method}")

        try:
            if method == 'boxplot':
                return self._boxplot_outlier_detection(data)
            elif method == 'zscore_mad':
                return self._zscore_mad_outlier_detection(data)
            else:
                self.logger.log_error(f"未知的异常检测方法: {method}")
                return data
        except Exception as e:
            self.logger.log_error(f"异常检测过程中发生错误: {str(e)}")
            return data
    
    def detect_outliers_with_method(self, data: pd.DataFrame, method: str) -> Optional[pd.DataFrame]:
        """使用指定方法检测异常值
        
        Args:
            data: 输入数据DataFrame
            method: 异常检测方法
            
        Returns:
            Optional[pd.DataFrame]: 清理后的数据
        """
        original_method = self.config.outlier_method
        self.config.outlier_method = method
        
        try:
            result = self.detect_and_remove_outliers(data)
            return result
        finally:
            self.config.outlier_method = original_method
    
    def _boxplot_outlier_detection(self, data: pd.DataFrame) -> pd.DataFrame:
        """箱线图异常检测方法
        
        Args:
            data: 输入数据DataFrame
            
        Returns:
            pd.DataFrame: 清理后的数据
        """
        self.logger.log_info("使用改良箱线图方法进行异常检测")

        cleaned_data = data.copy()

        # 按批次分组处理
        for batch in data['批次'].unique():
            batch_data = data[data['批次'] == batch].copy()

            if len(batch_data) < 3:  # 样本数太少，跳过
                continue

            self.logger.log_info(f"处理批次 {batch}，样本数: {len(batch_data)}")

            # 对首放进行异常检测
            batch_data = self._boxplot_detect_column(
                batch_data, '首放',
                self.config.boxplot_threshold_discharge
            )

            # 对首效进行异常检测
            batch_data = self._boxplot_detect_column(
                batch_data, '首效',
                self.config.boxplot_threshold_efficiency
            )
            
            # 更新清理后的数据
            cleaned_data = cleaned_data[~cleaned_data.index.isin(batch_data.index)]
            cleaned_data = pd.concat([cleaned_data, batch_data], ignore_index=True)
        
        removed_count = len(data) - len(cleaned_data)
        self.logger.log_info(f"箱线图异常检测完成，移除 {removed_count} 个异常值")
        
        return cleaned_data
    
    def _boxplot_detect_column(self, data: pd.DataFrame, column: str, threshold: float) -> pd.DataFrame:
        """对单列进行箱线图异常检测
        
        Args:
            data: 数据DataFrame
            column: 列名
            threshold: 极差阈值
            
        Returns:
            pd.DataFrame: 清理后的数据
        """
        if column not in data.columns:
            return data
        
        values = data[column].dropna()
        if len(values) < 3:
            return data
        
        filtered_data = data.copy()
        iteration = 0

        while iteration < self.config.max_iterations:
            current_values = filtered_data[column].dropna()
            if len(current_values) < 3:
                break
            
            # 计算四分位数
            q1 = current_values.quantile(0.25)
            q3 = current_values.quantile(0.75)
            iqr = q3 - q1
            current_range = current_values.max() - current_values.min()
            
            if current_range <= threshold:
                break
            
            # 计算动态收缩的IQR倍数
            iqr_multiplier = self.config.boxplot_shrink_factor ** iteration
            
            # 计算异常值边界
            lower_bound = q1 - 1.5 * iqr_multiplier * iqr
            upper_bound = q3 + 1.5 * iqr_multiplier * iqr
            
            # 移除异常值
            outlier_mask = (filtered_data[column] < lower_bound) | (filtered_data[column] > upper_bound)
            outliers = filtered_data[outlier_mask]
            
            if len(outliers) == 0:
                break
            
            self.logger.log_debug(f"第{iteration+1}次迭代，{column}列移除 {len(outliers)} 个异常值")
            filtered_data = filtered_data[~outlier_mask]
            iteration += 1
        
        return filtered_data
    
    def _zscore_mad_outlier_detection(self, data: pd.DataFrame) -> pd.DataFrame:
        """Z-score+MAD异常检测方法
        
        Args:
            data: 输入数据DataFrame
            
        Returns:
            pd.DataFrame: 清理后的数据
        """
        self.logger.log_info("使用Z-score+MAD方法进行异常检测")

        cleaned_data = data.copy()

        # 记录详细信息
        self.logger.log_outlier_detection(f"MAD常数: {self.config.zscore_mad_constant}")
        self.logger.log_outlier_detection(f"首放阈值: {self.config.zscore_threshold_discharge}")
        self.logger.log_outlier_detection(f"首效阈值: {self.config.zscore_threshold_efficiency}")
        
        # 按批次分组处理
        for batch in data['批次'].unique():
            batch_data = data[data['批次'] == batch].copy()
            
            if len(batch_data) < 3:  # 样本数太少，跳过
                continue
            
            self.logger.log_info(f"处理批次 {batch}，样本数: {len(batch_data)}")
            
            # 检测列配置
            detection_columns = {
                '首放': self.config.zscore_threshold_discharge,
                '首效': self.config.zscore_threshold_efficiency
            }
            
            outlier_indices = set()
            
            for column, threshold in detection_columns.items():
                if column in batch_data.columns:
                    column_outliers = self._zscore_mad_detect_column(
                        batch_data, column, threshold
                    )
                    outlier_indices.update(column_outliers)
            
            # 移除异常值
            if outlier_indices:
                self.logger.log_outlier_detection(f"批次 {batch} 移除异常值索引: {sorted(outlier_indices)}")
                batch_data = batch_data.drop(index=outlier_indices)
            
            # 更新清理后的数据
            cleaned_data = cleaned_data[~cleaned_data.index.isin(data[data['批次'] == batch].index)]
            cleaned_data = pd.concat([cleaned_data, batch_data], ignore_index=True)
        
        removed_count = len(data) - len(cleaned_data)
        self.logger.log_info(f"Z-score+MAD异常检测完成，移除 {removed_count} 个异常值")
        
        return cleaned_data
    
    def _zscore_mad_detect_column(self, data: pd.DataFrame, column: str, threshold: float) -> List[int]:
        """对单列进行Z-score+MAD异常检测
        
        Args:
            data: 数据DataFrame
            column: 列名
            threshold: Z-score阈值
            
        Returns:
            List[int]: 异常值索引列表
        """
        values = data[column].dropna()
        if len(values) < 3:
            return []
        
        # 计算中位数和MAD
        median = values.median()
        mad = self._calculate_mad(values, self.config.zscore_mad_constant)

        # 检查MAD是否太小
        min_mad_threshold = median * self.config.zscore_min_mad_ratio
        if mad < min_mad_threshold:
            mad = min_mad_threshold

        # 计算修正Z分数
        modified_z_scores = self.config.zscore_mad_constant * (values - median) / mad
        
        # 识别异常值
        outlier_mask = np.abs(modified_z_scores) > threshold
        outlier_indices = values[outlier_mask].index.tolist()
        
        if outlier_indices:
            self.logger.log_outlier_detection(f"{column}列检测到 {len(outlier_indices)} 个异常值")
        
        return outlier_indices
    
    def _calculate_mad(self, values: pd.Series, mad_constant: float) -> float:
        """计算中位数绝对偏差(MAD)

        Args:
            values: 数值序列
            mad_constant: MAD常数

        Returns:
            float: MAD值
        """
        median = values.median()
        deviations = np.abs(values - median)
        mad = deviations.median()

        # 应用MAD常数进行标准化
        return mad / mad_constant if mad > 0 else 1.0
