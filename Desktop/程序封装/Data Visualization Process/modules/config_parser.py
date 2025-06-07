#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置解析器模块
负责解析前端传递的所有配置参数，并提供配置验证功能
"""

import argparse
import os
from typing import Dict, Any, List, Tuple, Optional


class ConfigParser:
    """配置解析器类"""
    
    def __init__(self):
        """初始化配置解析器"""
        self.parser = self._create_argument_parser()
        
    def _create_argument_parser(self) -> argparse.ArgumentParser:
        """创建命令行参数解析器
        
        Returns:
            配置好的ArgumentParser对象
        """
        parser = argparse.ArgumentParser(
            description='电池数据分析工具 - 模块化版本',
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog="""
示例用法:
  python main.py --input_folder "C:/data" --outlier_method boxplot
  python main.py --input_folder "C:/data" --very_low_efficiency_threshold 60
            """
        )
        
        # ===== 基础运行参数 =====
        basic_group = parser.add_argument_group('基础运行参数')
        basic_group.add_argument(
            '--input_folder', 
            required=True, 
            help='输入文件夹路径'
        )
        basic_group.add_argument(
            '--output_folder', 
            help='输出文件夹路径（可选，默认为输入文件夹）'
        )
        basic_group.add_argument(
            '--outlier_method', 
            choices=['boxplot', 'zscore_mad'], 
            default='boxplot',
            help='异常检测方法 (默认: boxplot)'
        )
        basic_group.add_argument(
            '--reference_channel_method',
            choices=['traditional', 'pca', 'retention_curve_mse'],
            default='retention_curve_mse',
            help='参考通道选择方法 (默认: retention_curve_mse)'
        )
        basic_group.add_argument(
            '--output_format', 
            choices=['xlsx', 'csv'], 
            default='xlsx',
            help='输出格式 (默认: xlsx)'
        )
        
        # ===== Excel读取配置 =====
        excel_group = parser.add_argument_group('Excel读取配置')
        excel_group.add_argument(
            '--excel_engine', 
            choices=['calamine', 'openpyxl'], 
            default='calamine',
            help='Excel读取引擎 (默认: calamine)'
        )
        excel_group.add_argument(
            '--cycle_sheet_name', 
            default='Cycle',
            help='循环数据工作表名称 (默认: Cycle)'
        )
        excel_group.add_argument(
            '--test_sheet_name', 
            default='test',
            help='测试数据工作表名称 (默认: test)'
        )
        
        return parser
    
    def _add_one_c_thresholds(self, parser: argparse.ArgumentParser):
        """添加1C阈值配置参数"""
        one_c_group = parser.add_argument_group('1C阈值配置')
        one_c_group.add_argument(
            '--ratio_threshold', 
            type=float, 
            default=0.85,
            help='放电容量与首圈容量比值阈值 (默认: 0.85)'
        )
        one_c_group.add_argument(
            '--discharge_diff_threshold', 
            type=float, 
            default=15,
            help='放电容量与首圈容量差值阈值 (默认: 15)'
        )
        one_c_group.add_argument(
            '--overcharge_threshold', 
            type=float, 
            default=350,
            help='1C过充阈值(mAh/g) (默认: 350)'
        )
        one_c_group.add_argument(
            '--very_low_efficiency_threshold', 
            type=float, 
            default=80,
            help='1C首效过低阈值(%) (默认: 80)'
        )
        one_c_group.add_argument(
            '--low_efficiency_threshold',
            type=float,
            default=85,
            help='1C首效低阈值(%) (默认: 85)'
        )
        one_c_group.add_argument(
            '--default_1c_cycle',
            type=int,
            default=3,
            help='默认1C首圈编号 (默认: 3)'
        )
    
    def parse_arguments(self, args: Optional[List[str]] = None) -> 'Config':
        """解析命令行参数

        Args:
            args: 命令行参数列表，如果为None则从sys.argv获取

        Returns:
            Config对象，包含所有配置参数
        """
        # 添加所有参数组
        self._add_one_c_thresholds(self.parser)
        self._add_outlier_detection_params(self.parser)
        self._add_reference_channel_params(self.parser)
        self._add_plot_params(self.parser)
        self._add_filename_parse_params(self.parser)
        self._add_runtime_params(self.parser)
        self._add_data_validation_params(self.parser)
        self._add_output_params(self.parser)
        self._add_abnormal_thresholds(self.parser)
        self._add_overcharge_thresholds(self.parser)
        self._add_capacity_decay_thresholds(self.parser)
        self._add_mode_config_params(self.parser)
        self._add_series_config_params(self.parser)

        # 解析参数
        parsed_args = self.parser.parse_args(args)

        # 创建配置对象
        config = Config(parsed_args)

        # 验证配置
        self._validate_config(config)

        return config
    
    def _validate_config(self, config: 'Config'):
        """验证配置参数的有效性
        
        Args:
            config: 配置对象
            
        Raises:
            ValueError: 当配置参数无效时
        """
        # 验证输入文件夹是否存在
        if not os.path.exists(config.input_folder):
            raise ValueError(f"输入文件夹不存在: {config.input_folder}")
        
        # 验证阈值范围
        if not (0 <= config.very_low_efficiency_threshold <= 100):
            raise ValueError("very_low_efficiency_threshold 必须在 0-100 之间")
        
        if not (0 <= config.low_efficiency_threshold <= 100):
            raise ValueError("low_efficiency_threshold 必须在 0-100 之间")
        
        if config.very_low_efficiency_threshold >= config.low_efficiency_threshold:
            raise ValueError("very_low_efficiency_threshold 必须小于 low_efficiency_threshold")
        
        # 验证其他关键参数
        if config.ratio_threshold <= 0 or config.ratio_threshold > 1:
            raise ValueError("ratio_threshold 必须在 0-1 之间")
        
        if config.discharge_diff_threshold < 0:
            raise ValueError("discharge_diff_threshold 必须大于等于 0")

    def _add_reference_channel_params(self, parser: argparse.ArgumentParser):
        """添加参考通道选择配置参数"""
        ref_group = parser.add_argument_group('参考通道选择配置')

        # PCA配置
        ref_group.add_argument(
            '--pca_default_features',
            nargs='+',
            default=['首放', '首圈电压', 'Cycle4'],
            help='PCA分析默认特征 (默认: 首放 首圈电压 Cycle4)'
        )
        ref_group.add_argument(
            '--pca_n_components',
            type=int,
            default=2,
            help='PCA组件数 (默认: 2)'
        )

        # 容量保留率曲线比较配置
        ref_group.add_argument(
            '--capacity_retention_capacity_weight',
            type=float,
            default=0.6,
            help='容量保留率权重 (默认: 0.6)'
        )
        ref_group.add_argument(
            '--capacity_retention_voltage_weight',
            type=float,
            default=0.1,
            help='电压保留率权重 (默认: 0.1)'
        )
        ref_group.add_argument(
            '--capacity_retention_energy_weight',
            type=float,
            default=0.3,
            help='能量保留率权重 (默认: 0.3)'
        )

    def _add_plot_params(self, parser: argparse.ArgumentParser):
        """添加绘图配置参数"""
        plot_group = parser.add_argument_group('绘图配置')
        plot_group.add_argument(
            '--plot_font_family',
            default='SimHei',
            help='字体 (默认: SimHei)'
        )
        plot_group.add_argument(
            '--plot_dpi',
            type=int,
            default=300,
            help='图像DPI (默认: 300)'
        )

    def _add_filename_parse_params(self, parser: argparse.ArgumentParser):
        """添加文件名解析配置参数"""
        filename_group = parser.add_argument_group('文件名解析配置')
        filename_group.add_argument(
            '--device_id_max_length',
            type=int,
            default=20,
            help='设备ID最大长度 (默认: 20)'
        )

    def _add_runtime_params(self, parser: argparse.ArgumentParser):
        """添加程序运行配置参数"""
        runtime_group = parser.add_argument_group('程序运行配置')
        runtime_group.add_argument(
            '--verbose',
            action='store_true',
            default=False,
            help='是否显示详细输出 (默认: False)'
        )

    def _add_data_validation_params(self, parser: argparse.ArgumentParser):
        """添加数据验证配置参数"""
        pass  # 简化实现

    def _add_output_params(self, parser: argparse.ArgumentParser):
        """添加输出配置参数"""
        pass  # 简化实现

    def _add_abnormal_thresholds(self, parser: argparse.ArgumentParser):
        """添加异常数据阈值参数"""
        abnormal_group = parser.add_argument_group('异常数据阈值')
        abnormal_group.add_argument(
            '--abnormal_high_charge',
            type=float,
            default=380,
            help='首充容量上限(mAh/g) (默认: 380)'
        )
        abnormal_group.add_argument(
            '--abnormal_low_charge',
            type=float,
            default=200,
            help='首充容量下限(mAh/g) (默认: 200)'
        )
        abnormal_group.add_argument(
            '--abnormal_low_discharge',
            type=float,
            default=200,
            help='首放容量下限(mAh/g) (默认: 200)'
        )

    def _add_overcharge_thresholds(self, parser: argparse.ArgumentParser):
        """添加过充风险阈值参数"""
        overcharge_group = parser.add_argument_group('过充风险阈值')
        overcharge_group.add_argument(
            '--overcharge_voltage_warning',
            type=float,
            default=4.65,
            help='首充截止电压警告阈值(V) (默认: 4.65)'
        )
        overcharge_group.add_argument(
            '--overcharge_voltage_danger',
            type=float,
            default=4.7,
            help='首充截止电压危险阈值(V) (默认: 4.7)'
        )
        overcharge_group.add_argument(
            '--overcharge_efficiency_low',
            type=float,
            default=80,
            help='首效低值阈值(%) (默认: 80)'
        )
        overcharge_group.add_argument(
            '--overcharge_efficiency_warning',
            type=float,
            default=75,
            help='首效警告值阈值(%) (默认: 75)'
        )

    def _add_capacity_decay_thresholds(self, parser: argparse.ArgumentParser):
        """添加容量衰减预警阈值参数"""
        decay_group = parser.add_argument_group('容量衰减预警阈值')
        decay_group.add_argument(
            '--capacity_decay_cycle4_warning',
            type=float,
            default=85,
            help='Cycle4容量保持率警告阈值(%) (默认: 85)'
        )
        decay_group.add_argument(
            '--capacity_decay_cycle4_danger',
            type=float,
            default=80,
            help='Cycle4容量保持率危险阈值(%) (默认: 80)'
        )
        decay_group.add_argument(
            '--capacity_decay_discharge_diff',
            type=float,
            default=50,
            help='首放循环4差异阈值(mAh/g) (默认: 50)'
        )

    def _add_mode_config_params(self, parser: argparse.ArgumentParser):
        """添加测试模式配置参数"""
        mode_group = parser.add_argument_group('测试模式配置')
        mode_group.add_argument(
            '--mode_patterns',
            nargs='+',
            default=['-0.1C-', '-0.5C-', '-1C-', '-BL-', '-0.33C-'],
            help='测试模式标识列表 (默认: -0.1C- -0.5C- -1C- -BL- -0.33C-)'
        )
        mode_group.add_argument(
            '--mode_one_c_modes',
            nargs='+',
            default=['-1C-'],
            help='需要判断1C首圈的模式 (默认: -1C-)'
        )
        mode_group.add_argument(
            '--mode_non_one_c_modes',
            nargs='+',
            default=['-0.1C-', '-BL-', '-0.33C-'],
            help='不需要判断1C首圈的模式 (默认: -0.1C- -BL- -0.33C-)'
        )

    def _add_series_config_params(self, parser: argparse.ArgumentParser):
        """添加文件系列标识配置参数"""
        series_group = parser.add_argument_group('文件系列标识配置')
        series_group.add_argument(
            '--series_default',
            default='Q3',
            help='默认系列标识 (默认: Q3)'
        )

    def _add_outlier_detection_params(self, parser: argparse.ArgumentParser):
        """添加异常检测配置参数"""
        outlier_group = parser.add_argument_group('异常检测配置')

        # 箱线图法配置
        outlier_group.add_argument(
            '--boxplot_use_method',
            action='store_true',
            default=True,
            help='使用改良箱线图法 (默认: True)'
        )
        outlier_group.add_argument(
            '--boxplot_threshold_discharge',
            type=float,
            default=10,
            help='箱线图首放极差阈值 (默认: 10)'
        )
        outlier_group.add_argument(
            '--boxplot_threshold_efficiency',
            type=float,
            default=3,
            help='箱线图首效极差阈值 (默认: 3)'
        )
        outlier_group.add_argument(
            '--boxplot_shrink_factor',
            type=float,
            default=0.95,
            help='箱线图收缩因子 (默认: 0.95)'
        )

        # Z-score+MAD方法配置
        outlier_group.add_argument(
            '--zscore_mad_constant',
            type=float,
            default=0.6745,
            help='MAD常数 (默认: 0.6745)'
        )
        outlier_group.add_argument(
            '--zscore_min_mad_ratio',
            type=float,
            default=0.01,
            help='MAD最小值比例 (默认: 0.01)'
        )
        outlier_group.add_argument(
            '--zscore_threshold_discharge',
            type=float,
            default=3.0,
            help='Z-score首放阈值 (默认: 3.0)'
        )
        outlier_group.add_argument(
            '--zscore_threshold_efficiency',
            type=float,
            default=2.5,
            help='Z-score首效阈值 (默认: 2.5)'
        )


class Config:
    """配置类，存储所有配置参数"""
    
    def __init__(self, args: argparse.Namespace):
        """从解析的参数初始化配置
        
        Args:
            args: argparse解析的参数对象
        """
        # 基础运行参数
        self.input_folder = args.input_folder
        self.output_folder = args.output_folder or args.input_folder
        self.outlier_method = args.outlier_method
        self.reference_channel_method = args.reference_channel_method
        self.output_format = args.output_format
        
        # Excel读取配置
        self.excel_engine = args.excel_engine
        self.cycle_sheet_name = args.cycle_sheet_name
        self.test_sheet_name = args.test_sheet_name
        
        # 1C阈值配置
        self.ratio_threshold = args.ratio_threshold
        self.discharge_diff_threshold = args.discharge_diff_threshold
        self.overcharge_threshold = args.overcharge_threshold
        self.very_low_efficiency_threshold = args.very_low_efficiency_threshold
        self.low_efficiency_threshold = args.low_efficiency_threshold
        self.default_1c_cycle = getattr(args, 'default_1c_cycle', 3)

        # 异常检测配置
        self.boxplot_use_method = getattr(args, 'boxplot_use_method', True)
        self.boxplot_threshold_discharge = getattr(args, 'boxplot_threshold_discharge', 10)
        self.boxplot_threshold_efficiency = getattr(args, 'boxplot_threshold_efficiency', 3)
        self.boxplot_shrink_factor = getattr(args, 'boxplot_shrink_factor', 0.95)

        self.zscore_mad_constant = getattr(args, 'zscore_mad_constant', 0.6745)
        self.zscore_min_mad_ratio = getattr(args, 'zscore_min_mad_ratio', 0.01)
        self.zscore_threshold_discharge = getattr(args, 'zscore_threshold_discharge', 3.0)
        self.zscore_threshold_efficiency = getattr(args, 'zscore_threshold_efficiency', 2.5)
        self.zscore_threshold_voltage = getattr(args, 'zscore_threshold_voltage', 3.0)
        self.zscore_threshold_energy = getattr(args, 'zscore_threshold_energy', 3.0)
        self.zscore_use_time_series = getattr(args, 'zscore_use_time_series', True)
        self.zscore_min_samples_for_stl = getattr(args, 'zscore_min_samples_for_stl', 10)

        # 运行时配置
        self.max_iterations = getattr(args, 'max_iterations', 10)
        self.zscore_generate_plots = getattr(args, 'zscore_generate_plots', True)

        # 参考通道选择配置
        self.pca_default_features = getattr(args, 'pca_default_features', ['首放', '首圈电压', 'Cycle4'])
        self.pca_n_components = getattr(args, 'pca_n_components', 2)
        self.pca_visualization_enabled = getattr(args, 'pca_visualization_enabled', True)
        self.pca_safe_voltage_threshold = getattr(args, 'pca_safe_voltage_threshold', 4.65)

        self.capacity_retention_enabled = getattr(args, 'capacity_retention_enabled', True)
        self.capacity_retention_min_cycles = getattr(args, 'capacity_retention_min_cycles', 5)
        self.capacity_retention_max_cycles = getattr(args, 'capacity_retention_max_cycles', 800)
        self.capacity_retention_cycle_step = getattr(args, 'capacity_retention_cycle_step', 1)
        self.capacity_retention_interpolation_method = getattr(args, 'capacity_retention_interpolation_method', 'linear')
        self.capacity_retention_use_raw_capacity = getattr(args, 'capacity_retention_use_raw_capacity', True)
        self.capacity_retention_use_weighted_mse = getattr(args, 'capacity_retention_use_weighted_mse', True)
        self.capacity_retention_weight_method = getattr(args, 'capacity_retention_weight_method', 'linear')
        self.capacity_retention_weight_factor = getattr(args, 'capacity_retention_weight_factor', 1.0)
        self.capacity_retention_late_cycles_emphasis = getattr(args, 'capacity_retention_late_cycles_emphasis', 2.0)
        self.capacity_retention_dynamic_range = getattr(args, 'capacity_retention_dynamic_range', True)
        self.capacity_retention_min_channels = getattr(args, 'capacity_retention_min_channels', 2)
        self.capacity_retention_include_voltage = getattr(args, 'capacity_retention_include_voltage', True)
        self.capacity_retention_include_energy = getattr(args, 'capacity_retention_include_energy', True)
        self.capacity_retention_capacity_weight = getattr(args, 'capacity_retention_capacity_weight', 0.6)
        self.capacity_retention_voltage_weight = getattr(args, 'capacity_retention_voltage_weight', 0.1)
        self.capacity_retention_energy_weight = getattr(args, 'capacity_retention_energy_weight', 0.3)
        self.capacity_retention_voltage_column = getattr(args, 'capacity_retention_voltage_column', '当前电压保持')
        self.capacity_retention_energy_column = getattr(args, 'capacity_retention_energy_column', '当前能量保持')

        # 绘图配置
        self.plot_font_family = getattr(args, 'plot_font_family', 'SimHei')
        self.plot_font_size = getattr(args, 'plot_font_size', 10)
        self.plot_backend = getattr(args, 'plot_backend', 'Agg')
        self.plot_dpi = getattr(args, 'plot_dpi', 300)
        self.plot_figsize_width = getattr(args, 'plot_figsize_width', 10)
        self.plot_figsize_height = getattr(args, 'plot_figsize_height', 6)
        self.plot_interactive = getattr(args, 'plot_interactive', True)

        # 文件名解析配置
        self.device_id_max_length = getattr(args, 'device_id_max_length', 20)
        self.default_channel = getattr(args, 'default_channel', 'CH-01')
        self.batch_id_prefix = getattr(args, 'batch_id_prefix', 'BATCH-')
        self.device_id_prefix = getattr(args, 'device_id_prefix', 'DEVICE-')

        # 程序运行配置
        self.verbose = getattr(args, 'verbose', False)
        self.max_iterations = getattr(args, 'max_iterations', 10)
        self.chunk_size = getattr(args, 'chunk_size', 50)
        self.memory_limit_mb = getattr(args, 'memory_limit_mb', 500)
        self.enable_progress_bar = getattr(args, 'enable_progress_bar', True)
        self.auto_open_results = getattr(args, 'auto_open_results', False)
        self.backup_original_data = getattr(args, 'backup_original_data', True)
        self.log_level = getattr(args, 'log_level', 'INFO')

        # 数据验证配置
        self.min_cycles_required = getattr(args, 'min_cycles_required', 2)
        self.max_cycles_limit = getattr(args, 'max_cycles_limit', 1000)
        self.capacity_range_min = getattr(args, 'capacity_range_min', 0)
        self.capacity_range_max = getattr(args, 'capacity_range_max', 500)
        self.voltage_range_min = getattr(args, 'voltage_range_min', 2.0)
        self.voltage_range_max = getattr(args, 'voltage_range_max', 5.0)
        self.efficiency_range_min = getattr(args, 'efficiency_range_min', 0)
        self.efficiency_range_max = getattr(args, 'efficiency_range_max', 120)
        self.energy_range_min = getattr(args, 'energy_range_min', 0)
        self.energy_range_max = getattr(args, 'energy_range_max', 2000)

        # 输出配置
        self.output_excel_engine = getattr(args, 'output_excel_engine', 'openpyxl')
        self.include_charts = getattr(args, 'include_charts', True)
        self.chart_dpi = getattr(args, 'chart_dpi', 300)
        self.save_intermediate_results = getattr(args, 'save_intermediate_results', False)
        self.compress_output = getattr(args, 'compress_output', False)
        self.decimal_places = getattr(args, 'decimal_places', 2)

        # 异常数据阈值
        self.abnormal_high_charge = getattr(args, 'abnormal_high_charge', 380)
        self.abnormal_low_charge = getattr(args, 'abnormal_low_charge', 200)
        self.abnormal_low_discharge = getattr(args, 'abnormal_low_discharge', 200)

        # 过充风险阈值
        self.overcharge_voltage_warning = getattr(args, 'overcharge_voltage_warning', 4.65)
        self.overcharge_voltage_danger = getattr(args, 'overcharge_voltage_danger', 4.7)
        self.overcharge_efficiency_low = getattr(args, 'overcharge_efficiency_low', 80)
        self.overcharge_efficiency_warning = getattr(args, 'overcharge_efficiency_warning', 75)

        # 容量衰减预警阈值
        self.capacity_decay_cycle4_warning = getattr(args, 'capacity_decay_cycle4_warning', 85)
        self.capacity_decay_cycle4_danger = getattr(args, 'capacity_decay_cycle4_danger', 80)
        self.capacity_decay_discharge_diff = getattr(args, 'capacity_decay_discharge_diff', 50)

        # 测试模式配置
        self.mode_patterns = getattr(args, 'mode_patterns', ['-0.1C-', '-0.5C-', '-1C-', '-BL-', '-0.33C-'])
        self.mode_one_c_modes = getattr(args, 'mode_one_c_modes', ['-1C-'])
        self.mode_non_one_c_modes = getattr(args, 'mode_non_one_c_modes', ['-0.1C-', '-BL-', '-0.33C-'])

        # 文件系列标识配置
        self.series_default = getattr(args, 'series_default', 'Q3')

        # 存储原始args以便访问其他参数
        self._args = args
    
    def get(self, key: str, default: Any = None) -> Any:
        """获取配置值
        
        Args:
            key: 配置键名
            default: 默认值
            
        Returns:
            配置值
        """
        return getattr(self._args, key, default)
    
    def to_dict(self) -> Dict[str, Any]:
        """将配置转换为字典

        Returns:
            配置字典
        """
        return vars(self._args)

    def get_frontend_adjustable_params(self) -> Dict[str, List[str]]:
        """获取前端可调节的参数分组

        Returns:
            Dict[str, List[str]]: 按调节频率分组的参数列表
        """
        return {
            # 高频调节参数（用户经常需要调整）
            'high_frequency': [
                'outlier_method',  # 异常检测方法
                'very_low_efficiency_threshold',  # 首效过低阈值
                'low_efficiency_threshold',  # 首效低阈值
                'ratio_threshold',  # 比值阈值
                'discharge_diff_threshold',  # 差值阈值
                'overcharge_threshold',  # 过充阈值
                'default_1c_cycle',  # 默认1C首圈编号
            ],

            # 中频调节参数（偶尔需要调整）
            'medium_frequency': [
                'boxplot_threshold_discharge',  # 箱线图首放极差阈值
                'boxplot_threshold_efficiency',  # 箱线图首效极差阈值
                'boxplot_shrink_factor',  # 箱线图收缩因子
                'zscore_threshold_discharge',  # Z-score首放阈值
                'zscore_threshold_efficiency',  # Z-score首效阈值
                'zscore_mad_constant',  # MAD常数
                'reference_channel_method',  # 参考通道选择方法 (三选一)
                'capacity_retention_capacity_weight',  # 容量保留率权重
                'capacity_retention_voltage_weight',   # 电压保留率权重
                'capacity_retention_energy_weight',    # 能量保留率权重
            ],

            # 低频调节参数（很少需要调整）
            'low_frequency': [
                'excel_engine',  # Excel读取引擎
                'output_format',  # 输出格式
                'plot_font_family',  # 绘图字体
                'plot_dpi',  # 图像DPI
                'verbose',  # 详细输出
                'enable_progress_bar',  # 显示进度条
                'auto_open_results',  # 自动打开结果
            ]
        }

    def _add_reference_channel_params(self, parser: argparse.ArgumentParser):
        """添加参考通道选择配置参数"""
        ref_group = parser.add_argument_group('参考通道选择配置')

        # PCA配置
        ref_group.add_argument(
            '--pca_default_features',
            nargs='+',
            default=['首放', '首圈电压', 'Cycle4'],
            help='PCA分析默认特征 (默认: 首放 首圈电压 Cycle4)'
        )
        ref_group.add_argument(
            '--pca_n_components',
            type=int,
            default=2,
            help='PCA组件数 (默认: 2)'
        )
        ref_group.add_argument(
            '--pca_visualization_enabled',
            action='store_true',
            default=True,
            help='是否启用PCA可视化 (默认: True)'
        )
        ref_group.add_argument(
            '--pca_safe_voltage_threshold',
            type=float,
            default=4.65,
            help='安全电压阈值(V) (默认: 4.65)'
        )

        # 容量保留率曲线比较配置
        ref_group.add_argument(
            '--capacity_retention_enabled',
            action='store_true',
            default=True,
            help='是否启用容量保留率曲线比较 (默认: True)'
        )
        ref_group.add_argument(
            '--capacity_retention_min_cycles',
            type=int,
            default=5,
            help='最小循环次数要求 (默认: 5)'
        )
        ref_group.add_argument(
            '--capacity_retention_max_cycles',
            type=int,
            default=800,
            help='最大循环次数限制 (默认: 800)'
        )
        ref_group.add_argument(
            '--capacity_retention_cycle_step',
            type=int,
            default=1,
            help='循环步长 (默认: 1)'
        )
        ref_group.add_argument(
            '--capacity_retention_interpolation_method',
            choices=['linear', 'cubic', 'nearest'],
            default='linear',
            help='插值方法 (默认: linear)'
        )
        ref_group.add_argument(
            '--capacity_retention_use_raw_capacity',
            action='store_true',
            default=True,
            help='是否使用原始放电容量数据 (默认: True)'
        )
        ref_group.add_argument(
            '--capacity_retention_use_weighted_mse',
            action='store_true',
            default=True,
            help='是否使用加权MSE (默认: True)'
        )
        ref_group.add_argument(
            '--capacity_retention_weight_method',
            choices=['linear', 'exp', 'constant'],
            default='linear',
            help='权重方法 (默认: linear)'
        )
        ref_group.add_argument(
            '--capacity_retention_weight_factor',
            type=float,
            default=1.0,
            help='权重因子 (默认: 1.0)'
        )
        ref_group.add_argument(
            '--capacity_retention_late_cycles_emphasis',
            type=float,
            default=2.0,
            help='后期循环的权重倍数 (默认: 2.0)'
        )
        ref_group.add_argument(
            '--capacity_retention_dynamic_range',
            action='store_true',
            default=True,
            help='是否动态确定循环范围 (默认: True)'
        )
        ref_group.add_argument(
            '--capacity_retention_min_channels',
            type=int,
            default=2,
            help='最小通道数要求 (默认: 2)'
        )
        ref_group.add_argument(
            '--capacity_retention_include_voltage',
            action='store_true',
            default=True,
            help='是否包含电压保持率 (默认: True)'
        )
        ref_group.add_argument(
            '--capacity_retention_include_energy',
            action='store_true',
            default=True,
            help='是否包含能量保持率 (默认: True)'
        )
        ref_group.add_argument(
            '--capacity_retention_capacity_weight',
            type=float,
            default=0.6,
            help='容量保留率权重 (默认: 0.6)'
        )
        ref_group.add_argument(
            '--capacity_retention_voltage_weight',
            type=float,
            default=0.1,
            help='电压保持率权重 (默认: 0.1)'
        )
        ref_group.add_argument(
            '--capacity_retention_energy_weight',
            type=float,
            default=0.3,
            help='能量保持率权重 (默认: 0.3)'
        )
        ref_group.add_argument(
            '--capacity_retention_voltage_column',
            default='当前电压保持',
            help='电压保持率列名 (默认: 当前电压保持)'
        )
        ref_group.add_argument(
            '--capacity_retention_energy_column',
            default='当前能量保持',
            help='能量保持率列名 (默认: 当前能量保持)'
        )

    def _add_plot_params(self, parser: argparse.ArgumentParser):
        """添加绘图配置参数"""
        plot_group = parser.add_argument_group('绘图配置')
        plot_group.add_argument(
            '--plot_font_family',
            default='SimHei',
            help='字体 (默认: SimHei)'
        )
        plot_group.add_argument(
            '--plot_font_size',
            type=int,
            default=10,
            help='字体大小 (默认: 10)'
        )
        plot_group.add_argument(
            '--plot_backend',
            default='Agg',
            help='后端 (默认: Agg)'
        )
        plot_group.add_argument(
            '--plot_dpi',
            type=int,
            default=300,
            help='图像DPI (默认: 300)'
        )
        plot_group.add_argument(
            '--plot_figsize_width',
            type=int,
            default=10,
            help='图像宽度 (默认: 10)'
        )
        plot_group.add_argument(
            '--plot_figsize_height',
            type=int,
            default=6,
            help='图像高度 (默认: 6)'
        )
        plot_group.add_argument(
            '--plot_interactive',
            action='store_true',
            default=True,
            help='是否启用交互模式 (默认: True)'
        )

    def _add_filename_parse_params(self, parser: argparse.ArgumentParser):
        """添加文件名解析配置参数"""
        filename_group = parser.add_argument_group('文件名解析配置')
        filename_group.add_argument(
            '--device_id_max_length',
            type=int,
            default=20,
            help='设备ID最大长度 (默认: 20)'
        )
        filename_group.add_argument(
            '--default_channel',
            default='CH-01',
            help='默认通道ID (默认: CH-01)'
        )
        filename_group.add_argument(
            '--batch_id_prefix',
            default='BATCH-',
            help='批次ID前缀 (默认: BATCH-)'
        )
        filename_group.add_argument(
            '--device_id_prefix',
            default='DEVICE-',
            help='设备ID前缀 (默认: DEVICE-)'
        )

    def _add_runtime_params(self, parser: argparse.ArgumentParser):
        """添加程序运行配置参数"""
        runtime_group = parser.add_argument_group('程序运行配置')
        runtime_group.add_argument(
            '--verbose',
            action='store_true',
            default=False,
            help='是否显示详细输出 (默认: False)'
        )
        runtime_group.add_argument(
            '--max_iterations',
            type=int,
            default=10,
            help='异常值检测最大迭代次数 (默认: 10)'
        )
        runtime_group.add_argument(
            '--chunk_size',
            type=int,
            default=50,
            help='文件处理分块大小 (默认: 50)'
        )
        runtime_group.add_argument(
            '--memory_limit_mb',
            type=int,
            default=500,
            help='内存使用限制(MB) (默认: 500)'
        )
        runtime_group.add_argument(
            '--enable_progress_bar',
            action='store_true',
            default=True,
            help='是否显示进度条 (默认: True)'
        )
        runtime_group.add_argument(
            '--auto_open_results',
            action='store_true',
            default=False,
            help='是否自动打开结果文件 (默认: False)'
        )
        runtime_group.add_argument(
            '--backup_original_data',
            action='store_true',
            default=True,
            help='是否备份原始数据 (默认: True)'
        )
        runtime_group.add_argument(
            '--log_level',
            choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
            default='INFO',
            help='日志级别 (默认: INFO)'
        )

    def _add_data_validation_params(self, parser: argparse.ArgumentParser):
        """添加数据验证配置参数"""
        validation_group = parser.add_argument_group('数据验证配置')
        validation_group.add_argument(
            '--min_cycles_required',
            type=int,
            default=2,
            help='最少循环次数要求 (默认: 2)'
        )
        validation_group.add_argument(
            '--max_cycles_limit',
            type=int,
            default=1000,
            help='最大循环次数限制 (默认: 1000)'
        )
        validation_group.add_argument(
            '--capacity_range_min',
            type=float,
            default=0,
            help='容量合理范围最小值(mAh/g) (默认: 0)'
        )
        validation_group.add_argument(
            '--capacity_range_max',
            type=float,
            default=500,
            help='容量合理范围最大值(mAh/g) (默认: 500)'
        )
        validation_group.add_argument(
            '--voltage_range_min',
            type=float,
            default=2.0,
            help='电压合理范围最小值(V) (默认: 2.0)'
        )
        validation_group.add_argument(
            '--voltage_range_max',
            type=float,
            default=5.0,
            help='电压合理范围最大值(V) (默认: 5.0)'
        )
        validation_group.add_argument(
            '--efficiency_range_min',
            type=float,
            default=0,
            help='效率合理范围最小值(%) (默认: 0)'
        )
        validation_group.add_argument(
            '--efficiency_range_max',
            type=float,
            default=120,
            help='效率合理范围最大值(%) (默认: 120)'
        )
        validation_group.add_argument(
            '--energy_range_min',
            type=float,
            default=0,
            help='能量合理范围最小值(mWh/g) (默认: 0)'
        )
        validation_group.add_argument(
            '--energy_range_max',
            type=float,
            default=2000,
            help='能量合理范围最大值(mWh/g) (默认: 2000)'
        )

    def _add_output_params(self, parser: argparse.ArgumentParser):
        """添加输出配置参数"""
        output_group = parser.add_argument_group('输出配置')
        output_group.add_argument(
            '--output_excel_engine',
            choices=['openpyxl', 'xlsxwriter'],
            default='openpyxl',
            help='Excel写入引擎 (默认: openpyxl)'
        )
        output_group.add_argument(
            '--include_charts',
            action='store_true',
            default=True,
            help='是否包含图表 (默认: True)'
        )
        output_group.add_argument(
            '--chart_dpi',
            type=int,
            default=300,
            help='图表分辨率 (默认: 300)'
        )
        output_group.add_argument(
            '--save_intermediate_results',
            action='store_true',
            default=False,
            help='是否保存中间结果 (默认: False)'
        )
        output_group.add_argument(
            '--compress_output',
            action='store_true',
            default=False,
            help='是否压缩输出文件 (默认: False)'
        )
        output_group.add_argument(
            '--decimal_places',
            type=int,
            default=2,
            help='数值保留小数位数 (默认: 2)'
        )

    def _add_abnormal_thresholds(self, parser: argparse.ArgumentParser):
        """添加异常数据阈值参数"""
        abnormal_group = parser.add_argument_group('异常数据阈值')
        abnormal_group.add_argument(
            '--abnormal_high_charge',
            type=float,
            default=380,
            help='首充容量上限(mAh/g) (默认: 380)'
        )
        abnormal_group.add_argument(
            '--abnormal_low_charge',
            type=float,
            default=200,
            help='首充容量下限(mAh/g) (默认: 200)'
        )
        abnormal_group.add_argument(
            '--abnormal_low_discharge',
            type=float,
            default=200,
            help='首放容量下限(mAh/g) (默认: 200)'
        )

    def _add_overcharge_thresholds(self, parser: argparse.ArgumentParser):
        """添加过充风险阈值参数"""
        overcharge_group = parser.add_argument_group('过充风险阈值')
        overcharge_group.add_argument(
            '--overcharge_voltage_warning',
            type=float,
            default=4.65,
            help='首充截止电压警告阈值(V) (默认: 4.65)'
        )
        overcharge_group.add_argument(
            '--overcharge_voltage_danger',
            type=float,
            default=4.7,
            help='首充截止电压危险阈值(V) (默认: 4.7)'
        )
        overcharge_group.add_argument(
            '--overcharge_efficiency_low',
            type=float,
            default=80,
            help='首效低值阈值(%) (默认: 80)'
        )
        overcharge_group.add_argument(
            '--overcharge_efficiency_warning',
            type=float,
            default=75,
            help='首效警告值阈值(%) (默认: 75)'
        )

    def _add_capacity_decay_thresholds(self, parser: argparse.ArgumentParser):
        """添加容量衰减预警阈值参数"""
        decay_group = parser.add_argument_group('容量衰减预警阈值')
        decay_group.add_argument(
            '--capacity_decay_cycle4_warning',
            type=float,
            default=85,
            help='Cycle4容量保持率警告阈值(%) (默认: 85)'
        )
        decay_group.add_argument(
            '--capacity_decay_cycle4_danger',
            type=float,
            default=80,
            help='Cycle4容量保持率危险阈值(%) (默认: 80)'
        )
        decay_group.add_argument(
            '--capacity_decay_discharge_diff',
            type=float,
            default=50,
            help='首放循环4差异阈值(mAh/g) (默认: 50)'
        )

    def _add_mode_config_params(self, parser: argparse.ArgumentParser):
        """添加测试模式配置参数"""
        mode_group = parser.add_argument_group('测试模式配置')
        mode_group.add_argument(
            '--mode_patterns',
            nargs='+',
            default=['-0.1C-', '-0.5C-', '-1C-', '-BL-', '-0.33C-'],
            help='测试模式标识列表 (默认: -0.1C- -0.5C- -1C- -BL- -0.33C-)'
        )
        mode_group.add_argument(
            '--mode_one_c_modes',
            nargs='+',
            default=['-1C-'],
            help='需要判断1C首圈的模式 (默认: -1C-)'
        )
        mode_group.add_argument(
            '--mode_non_one_c_modes',
            nargs='+',
            default=['-0.1C-', '-BL-', '-0.33C-'],
            help='不需要判断1C首圈的模式 (默认: -0.1C- -BL- -0.33C-)'
        )

    def _add_series_config_params(self, parser: argparse.ArgumentParser):
        """添加文件系列标识配置参数"""
        series_group = parser.add_argument_group('文件系列标识配置')
        series_group.add_argument(
            '--series_default',
            default='Q3',
            help='默认系列标识 (默认: Q3)'
        )
        # 注意：系列匹配规则比较复杂，这里简化处理
        # 实际使用时可能需要通过配置文件或其他方式处理
