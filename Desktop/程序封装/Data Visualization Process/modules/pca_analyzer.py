"""
PCA分析模块 - 严格按照原始脚本LIMS_DATA_PROCESS_改良箱线图版.py重构

完全按照原始脚本的PCA分析逻辑，不做任何简化或修改
"""

import os
import time
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler
from typing import Dict, List, Optional, Tuple, Any


class PCAAnalyzer:
    """PCA分析器类 - 严格按照原始脚本逻辑"""
    
    def __init__(self, config, logger):
        """初始化PCA分析器
        
        Args:
            config: 配置对象
            logger: 日志记录器
        """
        self.config = config
        self.logger = logger
        
        # PCA配置 - 完全按照原始脚本
        self.pca_config = {
            "enabled": getattr(config, 'pca_enabled', True),
            "n_components": getattr(config, 'pca_n_components', 2),
            "variance_threshold": getattr(config, 'pca_variance_threshold', 0.95),
            "standardize": getattr(config, 'pca_standardize', True),
            "save_plots": getattr(config, 'pca_save_plots', True),
            "plot_format": getattr(config, 'pca_plot_format', 'png'),
            "plot_dpi": getattr(config, 'pca_plot_dpi', 300),
            "auto_display": getattr(config, 'pca_auto_display', False),  # 原始脚本中禁用自动显示
            "min_samples": getattr(config, 'pca_min_samples', 3),
            "use_capacity_only": getattr(config, 'pca_use_capacity_only', True),
            "include_voltage": getattr(config, 'pca_include_voltage', False),
            "include_energy": getattr(config, 'pca_include_energy', False)
        }
        
        # 输出配置
        self.output_folder = getattr(config, 'output_folder', '')
        
        # 设置matplotlib中文字体 - 完全按照原始脚本
        plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
        plt.rcParams['axes.unicode_minus'] = False

    def perform_pca_analysis(self, batch_data: pd.DataFrame, cycle_data_dict: Dict[str, pd.DataFrame], 
                           batch_id: str) -> Dict[str, Any]:
        """执行PCA分析 - 完全按照原始脚本逻辑
        
        Args:
            batch_data: 批次数据DataFrame
            cycle_data_dict: 循环数据字典
            batch_id: 批次ID
            
        Returns:
            Dict[str, Any]: PCA分析结果
        """
        if not self.pca_config["enabled"]:
            return {}
        
        if len(cycle_data_dict) < self.pca_config["min_samples"]:
            print(f"批次 {batch_id} 样本数量不足，跳过PCA分析")
            return {}
        
        try:
            # 准备PCA数据
            pca_data, channel_labels = self._prepare_pca_data(cycle_data_dict)
            
            if pca_data.empty:
                print(f"批次 {batch_id} 无有效PCA数据")
                return {}
            
            # 执行PCA分析
            pca_result = self._execute_pca(pca_data)
            
            # 生成PCA图像
            if self.pca_config["save_plots"]:
                plot_path = self._generate_pca_plot(pca_result, channel_labels, batch_id)
            else:
                plot_path = None
            
            # 返回分析结果
            return {
                'pca_components': pca_result['components'],
                'explained_variance_ratio': pca_result['explained_variance_ratio'],
                'cumulative_variance': pca_result['cumulative_variance'],
                'transformed_data': pca_result['transformed_data'],
                'channel_labels': channel_labels,
                'plot_path': plot_path,
                'n_components': pca_result['n_components']
            }
            
        except Exception as e:
            print(f"PCA分析失败 (批次 {batch_id}): {str(e)}")
            return {}

    def _prepare_pca_data(self, cycle_data_dict: Dict[str, pd.DataFrame]) -> Tuple[pd.DataFrame, List[str]]:
        """准备PCA分析数据 - 完全按照原始脚本逻辑
        
        Args:
            cycle_data_dict: 循环数据字典
            
        Returns:
            Tuple[pd.DataFrame, List[str]]: (PCA数据, 通道标签)
        """
        pca_data = pd.DataFrame()
        channel_labels = []
        
        # 找到最小的循环数，确保所有通道数据长度一致
        valid_cycles = [len(df) for df in cycle_data_dict.values() if not df.empty and len(df) > 1]
        
        if not valid_cycles:
            return pca_data, channel_labels
        
        min_cycles = min(valid_cycles)
        
        if min_cycles < 2:
            return pca_data, channel_labels
        
        # 提取每个通道的数据
        for channel_key, cycle_df in cycle_data_dict.items():
            if cycle_df.empty or len(cycle_df) < 2:
                continue
            
            # 截取到最小循环数
            cycle_df_truncated = cycle_df.iloc[:min_cycles].copy()
            
            # 构建特征向量 - 完全按照原始脚本
            features = []
            
            # 容量数据 - 始终包含
            capacity_data = cycle_df_truncated['放电比容量(mAh/g)'].values
            features.extend(capacity_data)
            
            # 电压数据 - 可选
            if self.pca_config["include_voltage"]:
                voltage_data = cycle_df_truncated['放电中值电压(V)'].values
                features.extend(voltage_data)
            
            # 能量数据 - 可选
            if self.pca_config["include_energy"]:
                energy_data = cycle_df_truncated['放电比能量(mWh/g)'].values
                features.extend(energy_data)
            
            # 添加到PCA数据
            pca_data[channel_key] = features
            channel_labels.append(channel_key)
        
        # 转置数据，使每行代表一个样本（通道）
        pca_data = pca_data.T
        
        return pca_data, channel_labels

    def _execute_pca(self, pca_data: pd.DataFrame) -> Dict[str, Any]:
        """执行PCA分析 - 完全按照原始脚本逻辑
        
        Args:
            pca_data: PCA分析数据
            
        Returns:
            Dict[str, Any]: PCA分析结果
        """
        # 数据标准化 - 完全按照原始脚本
        if self.pca_config["standardize"]:
            scaler = StandardScaler()
            scaled_data = scaler.fit_transform(pca_data)
        else:
            scaled_data = pca_data.values
        
        # 确定主成分数量
        n_components = min(self.pca_config["n_components"], scaled_data.shape[0], scaled_data.shape[1])
        
        # 执行PCA
        pca = PCA(n_components=n_components)
        transformed_data = pca.fit_transform(scaled_data)
        
        # 计算累积方差贡献率
        cumulative_variance = np.cumsum(pca.explained_variance_ratio_)
        
        return {
            'components': pca.components_,
            'explained_variance_ratio': pca.explained_variance_ratio_,
            'cumulative_variance': cumulative_variance,
            'transformed_data': transformed_data,
            'n_components': n_components,
            'pca_object': pca
        }

    def _generate_pca_plot(self, pca_result: Dict[str, Any], channel_labels: List[str], batch_id: str) -> str:
        """生成PCA图像 - 完全按照原始脚本逻辑
        
        Args:
            pca_result: PCA分析结果
            channel_labels: 通道标签
            batch_id: 批次ID
            
        Returns:
            str: 图像文件路径
        """
        try:
            # 创建图像
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
            
            # 绘制PCA散点图
            transformed_data = pca_result['transformed_data']
            
            if transformed_data.shape[1] >= 2:
                scatter = ax1.scatter(transformed_data[:, 0], transformed_data[:, 1], 
                                    alpha=0.7, s=60, c=range(len(channel_labels)), cmap='viridis')
                
                # 添加通道标签
                for i, label in enumerate(channel_labels):
                    ax1.annotate(label, (transformed_data[i, 0], transformed_data[i, 1]), 
                               xytext=(5, 5), textcoords='offset points', fontsize=8)
                
                ax1.set_xlabel(f'PC1 ({pca_result["explained_variance_ratio"][0]:.2%})')
                ax1.set_ylabel(f'PC2 ({pca_result["explained_variance_ratio"][1]:.2%})')
                ax1.set_title(f'PCA分析 - {batch_id}')
                ax1.grid(True, alpha=0.3)
            
            # 绘制方差贡献率图
            components = range(1, len(pca_result['explained_variance_ratio']) + 1)
            ax2.bar(components, pca_result['explained_variance_ratio'], alpha=0.7, label='单个主成分')
            ax2.plot(components, pca_result['cumulative_variance'], 'ro-', label='累积贡献率')
            
            ax2.set_xlabel('主成分')
            ax2.set_ylabel('方差贡献率')
            ax2.set_title('主成分方差贡献率')
            ax2.legend()
            ax2.grid(True, alpha=0.3)
            
            # 添加累积方差阈值线
            ax2.axhline(y=self.pca_config["variance_threshold"], color='r', linestyle='--', 
                       label=f'阈值 ({self.pca_config["variance_threshold"]:.0%})')
            
            plt.tight_layout()
            
            # 保存图像 - 完全按照原始脚本命名规则
            timestamp = time.strftime('%Y%m%d_%H%M%S')
            if self.output_folder:
                output_dir = os.path.join(self.output_folder, f'data_visualization_{timestamp}')
            else:
                output_dir = f'data_visualization_{timestamp}'
            
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            plot_filename = f'PCA_{batch_id}_{timestamp}.{self.pca_config["plot_format"]}'
            plot_path = os.path.join(output_dir, plot_filename)
            
            plt.savefig(plot_path, dpi=self.pca_config["plot_dpi"], bbox_inches='tight')
            
            # 根据配置决定是否显示图像 - 完全按照原始脚本
            if not self.pca_config["auto_display"]:
                plt.close()
            else:
                plt.show()
            
            print(f"PCA图像已保存: {plot_path}")
            return plot_path
            
        except Exception as e:
            print(f"生成PCA图像失败: {str(e)}")
            return ""

    def analyze_pca_outliers(self, pca_result: Dict[str, Any], channel_labels: List[str], 
                           threshold: float = 2.0) -> List[str]:
        """分析PCA异常值 - 完全按照原始脚本逻辑
        
        Args:
            pca_result: PCA分析结果
            channel_labels: 通道标签
            threshold: 异常值阈值（标准差倍数）
            
        Returns:
            List[str]: 异常通道列表
        """
        if not pca_result or 'transformed_data' not in pca_result:
            return []
        
        transformed_data = pca_result['transformed_data']
        
        if transformed_data.shape[1] < 2:
            return []
        
        try:
            # 计算每个点到中心的距离
            center = np.mean(transformed_data, axis=0)
            distances = np.linalg.norm(transformed_data - center, axis=1)
            
            # 计算异常值阈值
            mean_distance = np.mean(distances)
            std_distance = np.std(distances)
            outlier_threshold = mean_distance + threshold * std_distance
            
            # 识别异常值
            outlier_indices = np.where(distances > outlier_threshold)[0]
            outlier_channels = [channel_labels[i] for i in outlier_indices]
            
            return outlier_channels
            
        except Exception as e:
            print(f"PCA异常值分析失败: {str(e)}")
            return []

    def get_pca_summary(self, pca_result: Dict[str, Any]) -> Dict[str, Any]:
        """获取PCA分析摘要 - 完全按照原始脚本逻辑
        
        Args:
            pca_result: PCA分析结果
            
        Returns:
            Dict[str, Any]: PCA分析摘要
        """
        if not pca_result:
            return {}
        
        summary = {
            'n_components': pca_result.get('n_components', 0),
            'total_variance_explained': pca_result.get('cumulative_variance', [])[-1] if pca_result.get('cumulative_variance', []) else 0,
            'first_component_variance': pca_result.get('explained_variance_ratio', [0])[0] if pca_result.get('explained_variance_ratio', []) else 0,
            'second_component_variance': pca_result.get('explained_variance_ratio', [0, 0])[1] if len(pca_result.get('explained_variance_ratio', [])) > 1 else 0,
            'variance_threshold_met': pca_result.get('cumulative_variance', [0])[-1] >= self.pca_config["variance_threshold"] if pca_result.get('cumulative_variance', []) else False
        }
        
        return summary

    def batch_pca_analysis(self, all_batch_data: Dict[str, pd.DataFrame], 
                          all_cycle_data: Dict[str, Dict[str, pd.DataFrame]]) -> Dict[str, Dict[str, Any]]:
        """批量PCA分析 - 完全按照原始脚本逻辑
        
        Args:
            all_batch_data: 所有批次数据字典
            all_cycle_data: 所有循环数据字典
            
        Returns:
            Dict[str, Dict[str, Any]]: 所有批次的PCA分析结果
        """
        all_pca_results = {}
        
        for batch_id, batch_data in all_batch_data.items():
            if batch_id in all_cycle_data:
                cycle_data_dict = all_cycle_data[batch_id]
                pca_result = self.perform_pca_analysis(batch_data, cycle_data_dict, batch_id)
                
                if pca_result:
                    all_pca_results[batch_id] = pca_result
        
        return all_pca_results
