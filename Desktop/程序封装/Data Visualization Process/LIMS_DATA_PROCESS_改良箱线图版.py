"""
LIMS数据处理程序 - 电池数据分析工具

主要功能：
1. 批量处理电池测试数据文件
2. 使用改良箱线图法进行异常值检测
3. 多种方法选择1C参考通道（容量保留率曲线比较、PCA、传统方法）
4. 生成统计数据和可视化图表
5. 导出Excel汇总表

版本: v1.0
日期: 2024
"""

import os
import re
import time
import warnings
import sys
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import python_calamine
from openpyxl.utils import get_column_letter
from pandas import ExcelWriter
from tqdm import tqdm

from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler

# 忽略警告信息
warnings.filterwarnings('ignore')

class TeeOutput:
    """同时输出到控制台和文件的类"""
    def __init__(self, *files):
        self.files = files

    def write(self, text):
        for file in self.files:
            file.write(text)
            file.flush()

    def flush(self):
        for file in self.files:
            file.flush()

class ProcessingLogger:
    """处理日志管理器"""
    def __init__(self, output_dir=None):
        self.output_dir = output_dir or os.getcwd()
        self.timestamp = time.strftime('%Y%m%d_%H%M%S')

        # 创建日志文件夹
        self.log_dir = os.path.join(self.output_dir, f"处理日志-{self.timestamp}")
        os.makedirs(self.log_dir, exist_ok=True)

        # 创建不同类型的日志文件（使用无缓冲模式确保实时写入）
        self.main_log_file = open(os.path.join(self.log_dir, "主要处理日志.txt"), "w", encoding='utf-8', buffering=1)
        self.outlier_log_file = open(os.path.join(self.log_dir, "异常检测详细日志.txt"), "w", encoding='utf-8', buffering=1)
        self.debug_log_file = open(os.path.join(self.log_dir, "调试详细日志.txt"), "w", encoding='utf-8', buffering=1)

        # 保存原始stdout
        self.original_stdout = sys.stdout

        # 设置tee输出（同时输出到控制台和主日志）
        sys.stdout = TeeOutput(self.original_stdout, self.main_log_file)

        print(f"日志系统已启动，日志保存到: {self.log_dir}")
        print("=" * 80)

    def log_outlier_detection(self, message):
        """记录异常检测相关信息"""
        timestamp = time.strftime('%H:%M:%S')
        self.outlier_log_file.write(f"[{timestamp}] {message}\n")
        self.outlier_log_file.flush()

    def log_debug(self, message):
        """记录调试信息"""
        timestamp = time.strftime('%H:%M:%S')
        self.debug_log_file.write(f"[{timestamp}] {message}\n")
        self.debug_log_file.flush()

    def close(self):
        """关闭日志系统"""
        print("=" * 80)
        print(f"处理完成！详细日志已保存到: {self.log_dir}")
        print("日志文件说明:")
        print(f"  - 主要处理日志.txt: 主要处理过程和结果")
        print(f"  - 异常检测详细日志.txt: Z-score异常检测的详细信息")
        print(f"  - 调试详细日志.txt: 完整的调试信息")

        # 恢复原始stdout
        sys.stdout = self.original_stdout

        # 关闭文件
        self.main_log_file.close()
        self.outlier_log_file.close()
        self.debug_log_file.close()

# ===== 全局配置参数 =====
CONFIG = {
    # Excel读取配置
    "excel_engine": "calamine",  # Excel读取引擎
    "cycle_sheet_name": "Cycle",  # 循环数据工作表名称
    "test_sheet_name": "test",   # 测试数据工作表名称

    # Excel列配置
    "EXCEL_COLS": {
        'main': ['系列', '主机', '通道', '批次', '上架时间', '模式', '活性物质',
                '首充', '首放', '首效', '首圈电压', '首圈能量', 'Cycle2', 'Cycle2充电比容量',
                'Cycle3', 'Cycle3充电比容量', '1C首圈编号', '1C首充', '1C首放', '1C首效', '1C状态', '1C倍率比',
                'Cycle4', 'Cycle4充电比容量', 'Cycle5', 'Cycle5充电比容量', 'Cycle6', 'Cycle6充电比容量',
                'Cycle7', 'Cycle7充电比容量', '当前圈数', '当前容量保持', '电压衰减率mV/周',
                '当前电压保持', '当前能量保持', '100容量保持', '100电压保持', '100能量保持',
                '200容量保持', '200电压保持', '200能量保持' ],
        'cycle': ['系列', '主机', '通道', '批次', '上架时间', '模式', '活性物质',
                '首充', '首放', '首效', '首圈电压', '首圈能量', 'Cycle2', 'Cycle2充电比容量',
                'Cycle3', 'Cycle3充电比容量', '1C首圈编号', '1C首充', '1C首放', '1C首效', '1C状态', '1C倍率比',
                'Cycle4', 'Cycle4充电比容量', 'Cycle5', 'Cycle5充电比容量', 'Cycle6', 'Cycle6充电比容量',
                'Cycle7', 'Cycle7充电比容量', '当前圈数', '当前容量保持', '电压衰减率mV/周',
                '当前电压保持', '当前能量保持', '100容量保持', '100电压保持', '100能量保持',
                '200容量保持', '200电压保持', '200能量保持' ],
        'first_cycle': ['系列', '主机', '通道', '批次', '上架时间', '首充', '首放'],
        'error': ['系列', '主机', '通道', '批次', '上架时间', '首充', '首放', '当前圈数'],
        'inconsistent': ['系列', '主机', '通道', '批次', '上架时间', '首放差异', '首效差异'],
        'statistics': ['系列', '统一批次', '上架时间', '总数据', '首周有效数据', '首充', '首放', '首效',
                    '首圈电压', '首圈能量', '活性物质', 'Cycle2', 'Cycle2充电比容量', 'Cycle3', 'Cycle3充电比容量',
                    'Cycle4', 'Cycle4充电比容量', 'Cycle5', 'Cycle5充电比容量', 'Cycle6', 'Cycle6充电比容量',
                    'Cycle7', 'Cycle7充电比容量', '1C首周有效数据', '1C参考通道', '1C首圈编号', '1C首充', '1C首放',
                    '1C首效', '1C倍率比', '1C状态', '参考通道当前圈数', '当前容量保持', '电压衰减率mV/周', '当前电压保持', '当前能量保持']
    },

    # 循环数据列配置
    "CYCLE_SHEET_COLS": ['充电比容量(mAh/g)', '放电比容量(mAh/g)', '放电中值电压(V)',
                    '充电比能量(mWh/g)', '放电比能量(mWh/g)'],

    # 测试模式配置
    "MODE_CONFIG": {
        "patterns": ['-0.1C-', '-0.5C-', '-1C-', '-BL-', '-0.33C-'],  # 测试模式标识
        "one_c_modes": ['-1C-'],  # 需要判断1C首圈的模式
        "non_one_c_modes": ['-0.1C-', '-BL-', '-0.33C-']  # 不需要判断1C首圈的模式
    },

    # 文件系列标识配置
    "SERIES_CONFIG": {
        # 系列标识及其对应的匹配规则
        "series": {
            'G': {
                'include': ['-G-'],  # 包含这些字符串的文件属于G系列
                'exclude': ['-M-']   # 排除这些字符串的文件
            },
            'Q3': {
                'include': ['-Q3-']  # 包含这些字符串的文件属于Q3系列
            },
            'M': {
                'include': ['-M-']   # 包含这些字符串的文件属于M系列
            },
            'D': {
                'include': ['-D-']   # 包含这些字符串的文件属于D系列
            },
            'Z': {
                'include': ['-Z-']   # 包含这些字符串的文件属于Z系列
            }
        },
        # 默认系列（当所有匹配都失败时使用）
        "default_series": "Q3"
    },

    # 异常数据阈值
    "ABNORMAL_THRESHOLDS": {
        'high_charge': 380,   # 首充容量上限 (mAh/g)
        'low_charge': 200,     # 首充容量下限 (mAh/g)
        'low_discharge': 200  # 首放容量下限 (mAh/g)
    },

    # 1C首效和首圈阈值
    "ONE_C_THRESHOLDS": {
        "ratio_threshold": 0.85,  # 放电容量与首圈容量比值小于此值视为1C首圈
        "discharge_diff_threshold": 15,  # 放电容量与首圈容量差值大于此值视为1C首圈
        "overcharge_threshold": 350,  # 1C过充阈值(mAh/g)
        "very_low_efficiency_threshold": 80,  # 1C首效过低阈值(%) - 可修改为60%~80%之间
        "low_efficiency_threshold": 85  # 1C首效低阈值(%) - 可修改为80%~85%之间
    },

    # 参考通道选择配置
    "REFERENCE_CHANNEL_CONFIG": {
        # PCA参考通道选择配置
        "pca": {
            # 参数阈值范围定义
            "feature_thresholds": {
                '0.1C首充': (250, 350),      # mAh/g
                '0.1C首放': (200, 300),      # mAh/g
                '首圈电压': (3.4, 3.9),      # V
                '首圈能量': (600, 1300),     # 单位根据数据确定
                '首充截止电压': (None, 4.7), # V，上限
                'Cycle4容量保持率': (80, None) # %，下限
            },
            # PCA分析默认特征
            "default_features": ['首放', '首圈电压', 'Cycle4'],
            # PCA分析配置
            "n_components": 2,          # PCA组件数
            "visualization_enabled": True, # 是否启用可视化
            "safe_voltage_threshold": 4.65  # 安全电压阈值(V)
        },

        # 容量保留率曲线比较配置
        "capacity_retention": {
            "enabled": True,           # 是否启用容量保留率曲线比较
            "min_cycles": 5,           # 最小循环次数要求
            "max_cycles": 800,         # 最大循环次数限制
            "cycle_step": 1,           # 循环步长
            "interpolation_method": "linear",  # 插值方法：'linear', 'cubic', 'nearest'
            "retention_columns": [     # 用于计算保留率的列
                "当前容量保持", "100容量保持", "200容量保持"
            ],
            "use_raw_capacity": True,  # 是否使用原始放电容量数据重新计算保留率
            "use_weighted_mse": True,  # 是否使用加权MSE
            "weight_method": "linear", # 权重方法：'linear'(线性增长), 'exp'(指数增长), 'constant'(恒定权重)
            "weight_factor": 1.0,      # 权重因子，影响权重增长速度
            "late_cycles_emphasis": 2.0, # 后期循环的权重倍数
            "dynamic_range": True,     # 是否动态确定循环范围
            "min_channels": 2,         # 最小通道数要求
            "include_voltage": True,   # 是否包含电压保持率
            "include_energy": True,    # 是否包含能量保持率
            "capacity_weight": 0.6,    # 容量保留率在综合评分中的权重(0-1)
            "voltage_weight": 0.1,     # 电压保持率在综合评分中的权重(0-1)
            "energy_weight": 0.3,      # 能量保持率在综合评分中的权重(0-1)
            "voltage_column": "当前电压保持",  # 电压保持率列名
            "energy_column": "当前能量保持"    # 能量保持率列名
        },

        # 选择方法优先级
        "method_priority": [
            "capacity_retention",      # 容量保留率曲线比较
            "pca",                     # PCA多特征分析
            "traditional"              # 传统方法(首放接近均值)
        ]
    },

    # 异常检测配置
    "OUTLIER_DETECTION": {
        # 方法选择
        'method': 'boxplot',  # 'boxplot' 或 'zscore_mad'

        # 改良箱线图法配置
        'boxplot': {
            'use_boxplot_method': True,   # 使用改良箱线图法
            '首放_极差阈值': 10,           # 首放极差阈值
            '首效_极差阈值': 3,            # 首效极差阈值
            'boxplot_shrink_factor': 0.95  # 箱线图收缩因子
        },

        # Z-score+MAD方法配置
        'zscore_mad': {
            'mad_constant': 0.6745,      # MAD常数（标准值）
            'min_mad_ratio': 0.01,       # MAD最小值比例（中位数的1%）
            'thresholds': {
                '首放': 3.0,             # 首放Z-score阈值（降低，因为MAD已调整）
                '首效': 2.5,             # 首效Z-score阈值
                '首圈电压': 3.0,         # 首圈电压Z-score阈值
                '首圈能量': 3.0          # 首圈能量Z-score阈值
            },
            'use_time_series': True,     # 是否使用时间序列分解
            'min_samples_for_stl': 10,   # STL分解最小样本数
            'generate_plots': True       # 是否生成分布图
        }
    },

    # 过充风险阈值
    "OVERCHARGE_THRESHOLDS": {
        '首充截止电压_警告': 4.65,     # V
        '首充截止电压_危险': 4.7,      # V
        '首效_低值': 80,             # %
        '首效_警告值': 75            # %
    },

    # 容量衰减预警阈值
    "CAPACITY_DECAY_THRESHOLDS": {
        'Cycle4容量保持率_警告': 85,  # %
        'Cycle4容量保持率_危险': 80,  # %
        '首放循环4差异': 50          # mAh/g
    },

    # 绘图配置
    "PLOT_CONFIG": {
        "font_family": "SimHei",  # 字体
        "font_size": 10,          # 字体大小
        "backend": "Agg",         # 后端（非交互式，不显示图片）
        "dpi": 300,               # 图像DPI
        "figsize": (10, 6),       # 默认图像大小
        "interactive": True       # 是否启用交互模式
    },

    # 文件名解析配置
    "FILENAME_PARSE_CONFIG": {
        "device_id_max_length": 20,  # 设备ID最大长度
        "default_channel": "CH-01",  # 默认通道ID
        "batch_id_prefix": "BATCH-", # 批次ID前缀
        "device_id_prefix": "DEVICE-" # 设备ID前缀
    },

    # 程序运行配置
    "RUNTIME_CONFIG": {
        "verbose": False,             # 是否显示详细输出
        "max_iterations": 10,         # 异常值检测最大迭代次数
        "chunk_size": 50,             # 文件处理分块大小
        "memory_limit_mb": 500,       # 内存使用限制(MB)
        "enable_progress_bar": True,  # 是否显示进度条
        "auto_open_results": False,   # 是否自动打开结果文件
        "backup_original_data": True, # 是否备份原始数据
        "log_level": "INFO"           # 日志级别: DEBUG, INFO, WARNING, ERROR
    },

    # 数据验证配置
    "DATA_VALIDATION": {
        "min_cycles_required": 2,     # 最少循环次数要求
        "max_cycles_limit": 1000,     # 最大循环次数限制
        "capacity_range": (0, 500),   # 容量合理范围(mAh/g)
        "voltage_range": (2.0, 5.0),  # 电压合理范围(V)
        "efficiency_range": (0, 120), # 效率合理范围(%)
        "energy_range": (0, 2000),    # 能量合理范围(mWh/g)
        "required_columns": [         # 必需的数据列
            "充电比容量(mAh/g)", "放电比容量(mAh/g)", "放电中值电压(V)"
        ]
    },

    # 输出配置
    "OUTPUT_CONFIG": {
        "excel_engine": "openpyxl",   # Excel写入引擎
        "include_charts": True,       # 是否包含图表
        "chart_dpi": 300,            # 图表分辨率
        "save_intermediate_results": False, # 是否保存中间结果
        "compress_output": False,     # 是否压缩输出文件
        "output_formats": ["xlsx"],   # 输出格式列表
        "decimal_places": 2           # 数值保留小数位数
    }
}

# 删除ConfigManager类，直接使用CONFIG

class BatteryDataProcessor:
    """电池数据处理器"""

    def __init__(self, folder_path=None):
        """初始化电池数据处理器

        Args:
            folder_path: 数据文件夹路径
        """
        # 存储文件夹路径
        self.folder_path = folder_path

        # 初始化日志系统
        self.logger = ProcessingLogger(folder_path)

        # 记录初始化信息到调试日志
        self.logger.log_debug("BatteryDataProcessor初始化开始")
        self.logger.log_debug(f"文件夹路径: {folder_path}")
        self.logger.log_debug(f"配置信息: {CONFIG}")

        # 配置绘图环境
        self._setup_plot_environment()

        # 配置pandas显示选项
        self._setup_pandas_display()

        # 初始化数据容器
        self.first_cycle_files = []  # 仅1圈的文件
        self.error_files = []  # 异常数据文件

        # 结果数据框
        self.all_cycle_data = pd.DataFrame(columns=CONFIG["EXCEL_COLS"]['main'])
        self.all_first_cycle = pd.DataFrame(columns=CONFIG["EXCEL_COLS"]['first_cycle'])
        self.all_error_data = pd.DataFrame(columns=CONFIG["EXCEL_COLS"]['error'])
        self.statistics_data = pd.DataFrame()
        self.inconsistent_data = pd.DataFrame()

        # 输出配置
        self.verbose = CONFIG["RUNTIME_CONFIG"]["verbose"]  # 控制详细输出

    # ===== 环境设置方法 =====
    def _setup_plot_environment(self):
        """配置绘图环境，确保中文正确显示"""
        # 设置后端，确保图形可以显示
        import matplotlib
        matplotlib.use(CONFIG["PLOT_CONFIG"]["backend"])

        plt.rcParams['font.sans-serif'] = [CONFIG["PLOT_CONFIG"]["font_family"]]
        plt.rcParams['axes.unicode_minus'] = False
        plt.rc("font", family=CONFIG["PLOT_CONFIG"]["font_family"], size=str(CONFIG["PLOT_CONFIG"]["font_size"]))

        # 关闭交互模式，避免图片弹窗
        plt.ioff()  # 关闭交互模式

    def _setup_pandas_display(self):
        """配置pandas数据显示选项"""
        pd.set_option('max_colwidth', 300)
        pd.set_option('display.max_columns', None)
        pd.set_option('display.max_rows', None)

    # ===== 核心处理方法 =====
    def process_all_files(self, file_groups):
        """处理所有电池数据文件

        Args:
            file_groups: 按系列分组的文件路径列表字典
        """

        total_processed = 0
        total_successful = 0

        # 直接处理文件组，auto_detect_series已经返回了分组好的文件路径列表
        for series_name, files in file_groups.items():
            if not files:
                print(f'系列 {series_name} 没有文件，跳过处理')
                continue

            print(f'正在处理 {series_name} 组数据，共 {len(files)} 个文件')

            results = []
            successful_files = 0

            for file_path in tqdm(files, desc=f"正在处理{series_name}组数据", ncols=100):
                total_processed += 1
                try:
                    data = self._process_single_file(file_path, series_name)
                    if data:
                        results.append(data)
                        successful_files += 1
                        total_successful += 1
                except Exception as e:
                    if self.verbose:
                        print(f"处理文件失败: {os.path.basename(file_path)}, 错误: {str(e)}")

            print(f"系列 {series_name} 处理完成: {successful_files}/{len(files)} 个文件成功处理")

            if results:
                cycle_df = pd.DataFrame(results, columns=CONFIG["EXCEL_COLS"]['main'])
                self.all_cycle_data = pd.concat([self.all_cycle_data, cycle_df], ignore_index=True)
                print(f"系列 {series_name} 添加了 {len(results)} 条数据记录")
            else:
                print(f"系列 {series_name} 没有有效数据")

        print(f"\n数据处理总结:")
        print(f"总共处理文件: {total_processed} 个")
        print(f"成功处理文件: {total_successful} 个")
        print(f"总有效数据记录: {len(self.all_cycle_data)} 条")

    def _process_single_file(self, file_path, series_name):
        """处理单个电池数据文件

        Args:
            file_path: 文件路径
            series_name: 系列名称

        Returns:
            处理后的数据或None（如文件异常）
        """
        file_name = os.path.basename(file_path)

        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"文件不存在: {file_path}")
            return None

        # 读取循环数据
        try:
            # 尝试读取Excel文件
            cycle_df = pd.read_excel(file_path, sheet_name=CONFIG['cycle_sheet_name'],
                                 usecols=CONFIG["CYCLE_SHEET_COLS"], engine=CONFIG['excel_engine'])

            # 检查是否成功读取数据
            if cycle_df.empty:
                if self.verbose:
                    print(f"文件 {file_name} 的Cycle表为空")
                return None

        except Exception as e:
            if self.verbose:
                print(f"无法读取文件: {file_name}, 错误: {str(e)}")
            return None

        # 提取文件信息
        try:
            file_info = self._extract_file_info(file_path)
            if file_info['device_id'] == 'error' or file_info['channel_id'] == 'error':
                print(f"文件名格式不正确: {file_name}")
                return None
        except Exception as e:
            print(f"提取文件信息失败: {file_name}, 错误: {str(e)}")
            return None

        # 检查数据有效性
        if len(cycle_df) == 1:
            if self.verbose:
                print(f"文件 {file_name} 只有1个循环，添加到first_cycle_files")
            self.first_cycle_files.append((file_path, series_name))
            return None

        # 检查首圈数据是否异常
        try:
            if self._is_abnormal_first_cycle(cycle_df):
                print(f"文件 {file_name} 首圈数据异常，添加到error_files")
                self.error_files.append((file_path, series_name))
                return None
        except Exception as e:
            print(f"检查首圈数据异常失败: {file_name}, 错误: {str(e)}")
            return None

        # 处理循环数据
        try:
            result = self._process_cycle_data(cycle_df, file_info, series_name)
            if result:
                print(f"文件 {file_name} 处理成功")
                return result
            else:
                print(f"文件 {file_name} 处理结果为空")
                return None
        except Exception as e:
            print(f"处理循环数据失败: {file_name}, 错误: {str(e)}")
            return None

    def _extract_file_info(self, file_path):
        """从文件路径中提取电池信息

        Args:
            file_path: 文件路径

        Returns:
            包含电池信息的字典
        """
        try:
            # 首先从完整路径中提取文件名
            file_name = os.path.basename(file_path)
            print(f"正在解析文件: {file_name}")

            # 使用文件名进行解析
            parts = file_name.split(sep='-')
            # 分割文件名，获取下划线分隔的部分
            underscore_parts = file_name.split(sep='_')
            print(f"文件名分割结果: 破折号部分={len(parts)}个, 下划线部分={len(underscore_parts)}个")

            # 使用统一的主机通道解析方法
            device_id, channel_id = self._extract_host_and_channel(file_name)

            # 如果解析失败，使用备用方案
            if not device_id or not channel_id:
                print(f"主机通道解析失败，使用备用方案")
                # 备用方案：尝试从文件名中提取设备ID
                device_id = file_name.split('.')[0]  # 使用文件名作为设备ID
                max_length = CONFIG["FILENAME_PARSE_CONFIG"]["device_id_max_length"]
                if len(device_id) > max_length:  # 如果太长，截断
                    device_id = device_id[:max_length]

                # 备用通道ID
                channel_match = re.search(r'CH[-_]?(\d+)', file_name, re.IGNORECASE)
                if channel_match:
                    channel_id = f"CH-{channel_match.group(1)}"
                else:
                    channel_id = CONFIG["FILENAME_PARSE_CONFIG"]["default_channel"]  # 默认通道

            # 3. 批次ID提取 - 使用原始代码的下划线分割法
            try:
                batch_id = self._extract_batch_like_original(file_name)
                print(f"  提取的批次ID: {batch_id}")
            except Exception as e:
                print(f"  批次提取失败: {str(e)}，使用默认值")
                batch_id = CONFIG["FILENAME_PARSE_CONFIG"]["batch_id_prefix"] + file_name[:5]

            # 4. 上架时间提取 - 修改为优先使用原始代码的逻辑
            try:
                # 优先使用原始代码的逻辑：空格前的最后两个破折号部分
                if ' ' in file_name:
                    # 获取空格前的部分
                    before_space = file_name.split(' ')[0]
                    # 分割破折号
                    dash_parts = before_space.split('-')
                    if len(dash_parts) >= 2:
                        # 使用最后两个破折号部分
                        shelf_time = '-'.join(dash_parts[-2:])
                        print(f"  使用空格前的最后两个破折号部分作为上架时间: {shelf_time}")
                    else:
                        # 只有一个部分，使用该部分
                        shelf_time = dash_parts[-1]
                        print(f"  使用空格前的最后一个破折号部分作为上架时间: {shelf_time}")
                else:
                    # 如果没有空格，尝试使用正则表达式查找日期格式
                    date_match = re.search(r'(\d{4}[-/]?\d{2}[-/]?\d{2}|\d{6}|\d{2}\d{2})', file_name)
                    if date_match:
                        shelf_time = date_match.group(1)
                        print(f"  使用正则表达式找到的日期作为上架时间: {shelf_time}")
                    else:
                        # 如果没有找到日期格式，使用文件名中倒数第二个部分和最后一个部分
                        if len(parts) >= 2:
                            shelf_time = '-'.join(parts[-2:])
                            print(f"  使用文件名中最后两个破折号部分作为上架时间: {shelf_time}")
                        else:
                            # 如果没有足够的部分，使用当前日期
                            shelf_time = time.strftime('%m%d', time.localtime())
                            print(f"  使用当前日期作为上架时间: {shelf_time}")
            except Exception as e:
                print(f"  提取上架时间出错: {str(e)}，使用当前日期")
                shelf_time = time.strftime('%m%d', time.localtime())

            # 5. 测试模式识别
            mode = self._identify_test_mode(file_name)

            # 6. 读取活性物质质量
            try:
                test_df = pd.read_excel(file_path, sheet_name=CONFIG["test_sheet_name"], engine=CONFIG["excel_engine"])
                # 查找包含"活性物质"的行
                active_material_row = test_df[test_df.iloc[:, 0] == "活性物质"]
                if not active_material_row.empty:
                    mass = active_material_row.iloc[0, 1]  # 取对应行的第二列值
                else:
                    # 尝试其他可能的列名
                    for col_name in ["活性物质", "活性物质质量", "质量", "mass"]:
                        if col_name in test_df.columns:
                            mass = test_df[col_name].iloc[0]
                            break
                    else:
                        mass = None
            except Exception as e:
                print(f"读取活性物质失败: {os.path.basename(file_path)}, 错误: {str(e)}")
                mass = None

            result = {
                'device_id': device_id,
                'channel_id': channel_id,
                'batch_id': batch_id,
                'shelf_time': shelf_time,
                'mode': mode,
                'mass': mass
            }

            print(f"文件信息提取结果: 设备={device_id}, 通道={channel_id}, 批次={batch_id}, 上架时间={shelf_time}, 模式={mode}")
            return result

        except Exception as e:
            print(f"文件名解析错误: {os.path.basename(file_path)}, 错误: {str(e)}")
            # 提供默认值
            default_result = {
                'device_id': CONFIG["FILENAME_PARSE_CONFIG"]["device_id_prefix"] + os.path.basename(file_path)[:10],
                'channel_id': CONFIG["FILENAME_PARSE_CONFIG"]["default_channel"],
                'batch_id': CONFIG["FILENAME_PARSE_CONFIG"]["batch_id_prefix"] + time.strftime('%m%d', time.localtime()),
                'shelf_time': time.strftime('%m%d', time.localtime()),
                'mode': '-1C-',
                'mass': None
            }
            print(f"使用默认值: {default_result}")
            return default_result

    def _identify_test_mode(self, file_path):
        """识别测试模式

        Args:
            file_path: 文件路径或文件名

        Returns:
            测试模式标识
        """
        # 确保使用文件名而不是完整路径
        file_name = os.path.basename(file_path)

        for pattern in CONFIG["MODE_CONFIG"]["patterns"]:
            if pattern in file_name:
                return pattern
        return '-1C-'  # 默认模式

    def _is_abnormal_first_cycle(self, df):
        """检查首圈数据是否异常

        Args:
            df: 循环数据DataFrame

        Returns:
            布尔值，True表示异常
        """
        first_charge = df.loc[0, '充电比容量(mAh/g)']
        first_discharge = df.loc[0, '放电比容量(mAh/g)']

        return (first_charge > CONFIG["ABNORMAL_THRESHOLDS"]['high_charge'] or
                first_charge < CONFIG["ABNORMAL_THRESHOLDS"]['low_charge'] or
                first_discharge < CONFIG["ABNORMAL_THRESHOLDS"]['low_discharge'])

    def _process_cycle_data(self, df, file_info, series_name):
        """处理循环数据

        Args:
            df: 循环数据DataFrame
            file_info: 文件信息字典
            series_name: 系列名称

        Returns:
            处理后的数据列表
        """
        # 计算当前循环圈数（扣除最后1圈，从1开始）
        total_cycles = len(df) - 1

        # 首圈数据
        first_cycle = {
            'charge': round(df.loc[0, '充电比容量(mAh/g)'], 1),
            'discharge': round(df.loc[0, '放电比容量(mAh/g)'], 1),
            'voltage': round(df.loc[0, '放电中值电压(V)'], 2),
            'energy': round(df.loc[0, '放电比能量(mWh/g)'], 1)
        }
        first_efficiency = round(first_cycle['discharge'] / first_cycle['charge'] * 100, 1)

        # 初始化循环数据
        cycle_data = self._initialize_cycle_data()

        # 获取第2-4圈数据（如果存在）
        self._extract_early_cycles_data(df, total_cycles, cycle_data)

        # 处理1C循环数据（自动识别1C首圈）
        self._process_1c_data(df, total_cycles, cycle_data, file_info['mode'], first_cycle)

        # 处理当前循环和容量保持率数据
        self._process_retention_data(df, total_cycles, cycle_data, file_info['mode'], first_cycle)

        # 构建结果数据
        result = [
            series_name, file_info['device_id'], file_info['channel_id'],
            file_info['batch_id'], file_info['shelf_time'], file_info['mode'],
            file_info['mass'], first_cycle['charge'], first_cycle['discharge'],
            first_efficiency, first_cycle['voltage'], first_cycle['energy']
        ]

        # 添加其他循环数据
        for key in ['Cycle2', 'Cycle2充电比容量', 'Cycle3', 'Cycle3充电比容量',
                   '1C首圈编号', '1C首充', '1C首放', '1C首效', '1C状态', '1C倍率比', 'Cycle4',
                   'Cycle4充电比容量', 'Cycle5', 'Cycle5充电比容量', 'Cycle6', 'Cycle6充电比容量',
                   'Cycle7', 'Cycle7充电比容量', '当前圈数', '当前容量保持', '电压衰减率mV/周',
                   '当前电压保持', '当前能量保持', '100容量保持', '100电压保持',
                   '100能量保持', '200容量保持', '200电压保持', '200能量保持']:
            result.append(cycle_data.get(key))

        return result

    def _initialize_cycle_data(self):
        """初始化循环数据字典

        Returns:
            包含初始值的循环数据字典
        """
        return {
            'Cycle2充电比容量': None, 'Cycle2': None,
            'Cycle3充电比容量': None, 'Cycle3': None,
            'Cycle4充电比容量': None, 'Cycle4': None,
            'Cycle5充电比容量': None, 'Cycle5': None,
            'Cycle6充电比容量': None, 'Cycle6': None,
            'Cycle7充电比容量': None, 'Cycle7': None,
            '1C首圈编号': None, '1C首充': None, '1C首放': None, '1C首效': None,
            '1C状态': None, '1C倍率比': None,
            '当前圈数': None, '当前容量保持': None, '电压衰减率mV/周': None,
            '当前电压保持': None, '当前能量保持': None,
            '100容量保持': None, '100电压保持': None, '100能量保持': None,
            '200容量保持': None, '200电压保持': None, '200能量保持': None
        }

    def _extract_early_cycles_data(self, df, total_cycles, data):
        """提取前几个循环的数据

        Args:
            df: 循环数据DataFrame
            total_cycles: 总循环数
            data: 循环数据字典（将被修改）
        """
        # 更新当前圈数
        data['当前圈数'] = total_cycles

        # 提取第2圈数据
        if total_cycles >= 2:
            data['Cycle2'] = round(df.loc[1, '放电比容量(mAh/g)'], 1)
            data['Cycle2充电比容量'] = round(df.loc[1, '充电比容量(mAh/g)'], 1)

        # 提取第3圈数据
        if total_cycles >= 3:
            data['Cycle3'] = round(df.loc[2, '放电比容量(mAh/g)'], 1)
            data['Cycle3充电比容量'] = round(df.loc[2, '充电比容量(mAh/g)'], 1)

        # 提取第4圈数据
        if total_cycles >= 4:
            data['Cycle4'] = round(df.loc[3, '放电比容量(mAh/g)'], 1)
            data['Cycle4充电比容量'] = round(df.loc[3, '充电比容量(mAh/g)'], 1)

        # 提取第5圈数据
        if total_cycles >= 5:
            data['Cycle5'] = round(df.loc[4, '放电比容量(mAh/g)'], 1)
            data['Cycle5充电比容量'] = round(df.loc[4, '充电比容量(mAh/g)'], 1)

        # 提取第6圈数据
        if total_cycles >= 6:
            data['Cycle6'] = round(df.loc[5, '放电比容量(mAh/g)'], 1)
            data['Cycle6充电比容量'] = round(df.loc[5, '充电比容量(mAh/g)'], 1)

        # 提取第7圈数据
        if total_cycles >= 7:
            data['Cycle7'] = round(df.loc[6, '放电比容量(mAh/g)'], 1)
            data['Cycle7充电比容量'] = round(df.loc[6, '充电比容量(mAh/g)'], 1)

    def _process_1c_data(self, df, total_cycles, data, mode, first_cycle):
        """处理1C相关的循环数据，使用比值和差值双重验证识别1C首圈

        Args:
            df: 循环数据DataFrame
            total_cycles: 总循环数
            data: 循环数据字典（将被修改）
            mode: 测试模式
            first_cycle: 首圈数据字典
        """
        # 调试输出：开始处理1C数据
        if self.verbose:
            print(f"\n开始处理1C数据:")
            print(f"  测试模式: {mode}")
            print(f"  总循环数: {total_cycles}")
            print(f"  首圈放电容量: {first_cycle['discharge']:.2f}mAh/g")
            print(f"  首圈充电容量: {first_cycle['charge']:.2f}mAh/g")
            print(f"  首圈效率: {(first_cycle['discharge']/first_cycle['charge']*100):.2f}%")

        # 仅在1C相关模式下识别1C首圈
        if mode in CONFIG["MODE_CONFIG"]["one_c_modes"]:
            # 寻找前四圈中符合双重验证标准的循环作为1C首圈
            first_discharge = first_cycle['discharge']
            one_c_start_idx = None

            # 检查第2至第5圈
            for idx in range(1, min(4, total_cycles)):
                cycle_discharge = df.loc[idx, '放电比容量(mAh/g)']
                cycle_charge = df.loc[idx, '充电比容量(mAh/g)']
                cycle_efficiency = cycle_discharge / cycle_charge * 100

                # 双重验证：同时满足比值和差值标准
                discharge_ratio = cycle_discharge / first_discharge
                discharge_diff = first_discharge - cycle_discharge

                # 输出比值和差值，便于调试
                print(f"  第{idx+1}圈: 充电={cycle_charge:.2f}mAh/g, 放电={cycle_discharge:.2f}mAh/g, 效率={cycle_efficiency:.2f}%, 与首圈比值={discharge_ratio:.3f}, 差值={discharge_diff:.2f}mAh/g")

                # 同时满足两个条件判定为1C首圈 - 使用常量替代硬编码值
                if discharge_ratio < CONFIG["ONE_C_THRESHOLDS"]["ratio_threshold"] and discharge_diff > CONFIG["ONE_C_THRESHOLDS"]["discharge_diff_threshold"]:
                    one_c_start_idx = idx
                    print(f"  ** 自动识别1C首圈为第{idx+1}圈，比值={discharge_ratio:.3f}，差值={discharge_diff:.2f}mAh/g **")
                    break

            # 如果找到1C首圈
            if one_c_start_idx is not None:
                # 记录1C首圈编号（从1开始计数）
                data['1C首圈编号'] = one_c_start_idx + 1

                # 设置1C首圈数据
                data['1C首充'] = round(df.loc[one_c_start_idx, '充电比容量(mAh/g)'], 1)
                data['1C首放'] = round(df.loc[one_c_start_idx, '放电比容量(mAh/g)'], 1)
                data['1C首效'] = round(100 * data['1C首放'] / data['1C首充'], 1)

                # 计算1C倍率比
                data['1C倍率比'] = round(100 * data['1C首放'] / first_cycle['discharge'], 2)

                # 默认状态设为"正常"，确保没有None值
                data['1C状态'] = "正常"

                # 明确检查1C状态 - 修改判断顺序，优先判断过充 - 使用常量替代硬编码值
                if data['1C首充'] > CONFIG["ONE_C_THRESHOLDS"]["overcharge_threshold"]:
                    data['1C状态'] = '1C过充'
                # 修改：细分首效状态 - 使用常量替代硬编码值
                elif data['1C首效'] < CONFIG["ONE_C_THRESHOLDS"]["very_low_efficiency_threshold"]:
                    data['1C状态'] = '首效过低'  # <80%: 首效过低
                elif data['1C首效'] < CONFIG["ONE_C_THRESHOLDS"]["low_efficiency_threshold"]:
                    data['1C状态'] = '首效低'     # 80-85%: 首效低

                print(f"识别到1C首圈：第{one_c_start_idx+1}圈，放电比容量 {df.loc[one_c_start_idx, '放电比容量(mAh/g)']}，"
                    f"与首圈比值 {discharge_ratio:.3f}，差值 {discharge_diff:.2f}mAh/g，状态为{data['1C状态']}")

            # 如果未找到1C首圈，默认使用第4圈
            elif total_cycles >= 4:
                data['1C首圈编号'] = 4  # 第4圈
                data['1C首充'] = round(df.loc[3, '充电比容量(mAh/g)'], 1)
                data['1C首放'] = round(df.loc[3, '放电比容量(mAh/g)'], 1)
                data['1C首效'] = round(100 * data['1C首放'] / data['1C首充'], 1)
                data['1C倍率比'] = round(100 * data['1C首放'] / first_cycle['discharge'], 2)

                # 默认状态设为"正常"，确保没有None值
                data['1C状态'] = "正常"

                # 明确检查1C状态 - 修改判断顺序，优先判断过充 - 使用常量替代硬编码值
                if data['1C首充'] > CONFIG["ONE_C_THRESHOLDS"]["overcharge_threshold"]:
                    data['1C状态'] = '1C过充'
                # 修改：细分首效状态 - 使用常量替代硬编码值
                elif data['1C首效'] < CONFIG["ONE_C_THRESHOLDS"]["very_low_efficiency_threshold"]:
                    data['1C状态'] = '首效过低'  # <80%: 首效过低
                elif data['1C首效'] < CONFIG["ONE_C_THRESHOLDS"]["low_efficiency_threshold"]:
                    data['1C状态'] = '首效低'     # 80-85%: 首效低

                print(f"未找到同时满足比值(<{CONFIG['ONE_C_THRESHOLDS']['ratio_threshold']})和差值(>{CONFIG['ONE_C_THRESHOLDS']['discharge_diff_threshold']}mAh/g)条件的1C首圈，使用默认第4圈作为1C首圈，状态为{data['1C状态']}")

        # # 非1C模式但总循环数足够，仍计算倍率比（使用第4圈）
        # elif total_cycles >= 4:
        #     data['1C倍率比'] = round(100 * df.loc[3, '放电比容量(mAh/g)'] / first_cycle['discharge'], 2)

    def _process_retention_data(self, df, total_cycles, data, mode, first_cycle):
        """处理容量保持率相关数据

        Args:
            df: 循环数据DataFrame
            total_cycles: 总循环数
            data: 循环数据字典（将被修改）
            mode: 测试模式
            first_cycle: 首圈数据字典
        """
        # 记录1C状态，但不再基于它决定是否计算容量保持率
        # 在"循环2圈以上"表格中记录容量保持率数据
        if '1C状态' in data and (data['1C状态'] == '1C过充' or data['1C状态'] == '首效过低'):
            print(f"  样品状态为{data['1C状态']}，但仍计算容量保持率以便于展示")
            # 不再提前返回，继续执行计算

        # 只有当圈数>4时才计算容量保持率
        if total_cycles <= 4:
            print(f"  当前圈数({total_cycles})不超过4，不计算容量保持率")
            return

        # 0.1C循环容量保持率计算
        if mode == '-0.1C-':
            data['当前容量保持'] = round(100 * df.loc[total_cycles-1, '放电比容量(mAh/g)'] / first_cycle['discharge'], 1)
            data['电压衰减率mV/周'] = round((first_cycle['voltage'] - df.loc[total_cycles-1, '放电中值电压(V)']) * 1000 / (total_cycles-1), 1)
            data['当前电压保持'] = round(100 * df.loc[total_cycles-1, '放电中值电压(V)'] / first_cycle['voltage'], 1)
            data['当前能量保持'] = round(100 * df.loc[total_cycles-1, '放电比能量(mWh/g)'] / first_cycle['energy'], 1)

        # 1C循环容量保持率计算（使用动态识别的1C首圈）
        elif mode in CONFIG["MODE_CONFIG"]["one_c_modes"]:
            # 确定1C首圈索引
            one_c_idx = None
            if '1C首圈编号' in data and data['1C首圈编号'] is not None:
                one_c_idx = data['1C首圈编号'] - 1  # 转换为索引（从0开始）

            # 如果仍未识别，使用默认值
            if one_c_idx is None:
                one_c_idx = min(3, total_cycles-1)  # 默认第4圈或最大可用圈

            # 确保有足够的循环数据
            if total_cycles > one_c_idx + 1:
                # 获取1C首圈数据
                one_c_discharge = df.loc[one_c_idx, '放电比容量(mAh/g)']
                one_c_voltage = df.loc[one_c_idx, '放电中值电压(V)']
                one_c_energy = df.loc[one_c_idx, '放电比能量(mWh/g)']

                # 计算容量保持率
                data['当前容量保持'] = round(100 * df.loc[total_cycles-1, '放电比容量(mAh/g)'] / one_c_discharge, 1)
                data['电压衰减率mV/周'] = round((one_c_voltage - df.loc[total_cycles-1, '放电中值电压(V)']) * 1000 / (total_cycles-(one_c_idx+1)), 1)
                data['当前电压保持'] = round(100 * df.loc[total_cycles-1, '放电中值电压(V)'] / one_c_voltage, 1)
                data['当前能量保持'] = round(100 * df.loc[total_cycles-1, '放电比能量(mWh/g)'] / one_c_energy, 1)

                # 动态计算100圈位置
                idx_100_cycles = one_c_idx + 100
                if total_cycles > idx_100_cycles:
                    data['100容量保持'] = round(100 * df.loc[idx_100_cycles, '放电比容量(mAh/g)'] / one_c_discharge, 1)
                    data['100电压保持'] = round(100 * df.loc[idx_100_cycles, '放电中值电压(V)'] / one_c_voltage, 1)
                    data['100能量保持'] = round(100 * df.loc[idx_100_cycles, '放电比能量(mWh/g)'] / one_c_energy, 1)

                # 动态计算200圈位置
                idx_200_cycles = one_c_idx + 200
                if total_cycles > idx_200_cycles:
                    data['200容量保持'] = round(100 * df.loc[idx_200_cycles, '放电比容量(mAh/g)'] / one_c_discharge, 1)
                    data['200电压保持'] = round(100 * df.loc[idx_200_cycles, '放电中值电压(V)'] / one_c_voltage, 1)
                    data['200能量保持'] = round(100 * df.loc[idx_200_cycles, '放电比能量(mWh/g)'] / one_c_energy, 1)

    # ===== 特殊文件处理方法 =====
    def process_first_cycle_files(self):
        """处理仅有1个循环的文件"""
        if not self.first_cycle_files:
            return

        results = []
        for file_path, series_name in tqdm(self.first_cycle_files, desc="正在处理仅1圈数据", ncols=100):
            try:
                df = pd.read_excel(file_path, sheet_name='Cycle',
                                 usecols=['充电比容量(mAh/g)', '放电比容量(mAh/g)', '放电中值电压(V)'],
                                 engine='calamine')

                # 使用优化后的文件信息提取
                file_info = self._extract_file_info(file_path)

                first_charge = round(df.loc[0, '充电比容量(mAh/g)'], 2)
                first_discharge = round(df.loc[0, '放电比容量(mAh/g)'], 2)

                results.append([series_name, file_info['device_id'], file_info['channel_id'],
                              file_info['batch_id'], file_info['shelf_time'], first_charge, first_discharge])
            except Exception as e:
                print(f"处理文件失败: {file_path}, 错误: {str(e)}")
                continue

        if results:
            first_cycle_df = pd.DataFrame(results, columns=CONFIG["EXCEL_COLS"]['first_cycle'])
            self.all_first_cycle = pd.concat([self.all_first_cycle, first_cycle_df], ignore_index=True)

    def process_error_files(self):
        """处理异常数据文件"""
        if not self.error_files:
            return

        results = []
        for file_path, series_name in tqdm(self.error_files, desc="正在处理异常数据", ncols=100):
            try:
                df = pd.read_excel(file_path, sheet_name='Cycle',
                                 usecols=['充电比容量(mAh/g)', '放电比容量(mAh/g)', '放电中值电压(V)'],
                                 engine='calamine')

                # 使用优化后的文件信息提取
                file_info = self._extract_file_info(file_path)

                first_charge = round(df.loc[0, '充电比容量(mAh/g)'], 2)
                first_discharge = round(df.loc[0, '放电比容量(mAh/g)'], 2)
                total_cycles = len(df) - 1

                results.append([series_name, file_info['device_id'], file_info['channel_id'],
                              file_info['batch_id'], file_info['shelf_time'], first_charge,
                              first_discharge, total_cycles])
            except Exception as e:
                print(f"处理文件失败: {file_path}, 错误: {str(e)}")
                continue

        if results:
            error_df = pd.DataFrame(results, columns=CONFIG["EXCEL_COLS"]['error'])
            self.all_error_data = pd.concat([self.all_error_data, error_df], ignore_index=True)

    # ===== 数据统计方法 =====
    def calculate_statistics(self, use_multi_feature=True):
        """计算统计数据

        Args:
            use_multi_feature: 是否使用多特征方法选择参考通道
        """
        # 保存参数供内部方法使用
        self._use_multi_feature = use_multi_feature
        if self.all_cycle_data.empty:
            # 即使没有数据，也创建一个空的统计数据DataFrame
            self.statistics_data = pd.DataFrame(columns=CONFIG["EXCEL_COLS"].get('statistics',
                ['系列', '统一批次', '样品数', '首放平均值', '首效平均值', '首放标准差', '首效标准差',
                 '首放极差', '首效极差', '参考通道', '参考通道首放', '参考通道首效']))
            print("没有循环数据，创建空的统计数据DataFrame")
            return

        # 添加统一批次列 - 从批次中提取统一批次
        self.all_cycle_data['统一批次'] = self.all_cycle_data['批次'].apply(self._extract_unified_batch)

        # 打印批次和统一批次的对应关系，帮助调试
        if self.verbose:
            print("\n批次和统一批次的对应关系:")
            batch_mapping = {}
            for idx, row in self.all_cycle_data.iterrows():
                batch = row['批次']
                unified_batch = row['统一批次']
                if batch not in batch_mapping:
                    batch_mapping[batch] = unified_batch
                    print(f"  批次: {batch}")
                    print(f"  统一批次: {unified_batch}")
                    print("")

            # 按统一批次分组，显示每个统一批次下的通道
            print("\n按统一批次分组的通道:")
            for unified_batch, group in self.all_cycle_data.groupby('统一批次'):
                channels = group[['主机', '通道']].drop_duplicates()
                print(f"\n统一批次: {unified_batch}, 包含 {len(channels)} 个通道")
                for i, (_, row) in enumerate(channels.iterrows()):
                    print(f"  {i+1}. 主机={row['主机']}, 通道={row['通道']}")



        # 根据配置选择异常检测方法 - 直接从CONFIG获取最新配置
        outlier_method = CONFIG["OUTLIER_DETECTION"].get('method', 'boxplot')
        print(f"使用异常检测方法: {outlier_method}")

        if outlier_method == 'zscore_mad':
            print("\n使用Z-score+MAD方法剔除异常值...")
            filtered_data, completely_removed_batches = self._remove_outliers_zscore_mad(self.all_cycle_data)

            # 生成Z-score分布图（如果配置允许）
            if (hasattr(self, 'folder_path') and self.folder_path and
                CONFIG["OUTLIER_DETECTION"]['zscore_mad'].get('generate_plots', True)):
                try:
                    self._generate_zscore_distribution_plots(self.all_cycle_data, self.folder_path)
                except Exception as e:
                    print(f"生成Z-score分布图时出错: {e}")
                    print("继续执行程序，跳过分布图生成")
        else:
            # 使用改良箱线图法剔除离散点，同时获取完全被剔除的批次
            print("\n开始使用改良箱线图法剔除首放异常值...")
            filtered_data, removed_batches1 = self._remove_outliers(self.all_cycle_data, '首放')

            print("\n开始使用改良箱线图法剔除首效异常值...")
            filtered_data, removed_batches2 = self._remove_outliers(filtered_data, '首效')

            # 合并所有完全剔除的批次
            completely_removed_batches = list(set(removed_batches1 + removed_batches2))

        # 添加调试信息
        print(f"完全剔除的批次数量: {len(completely_removed_batches)}")
        if completely_removed_batches:
            print(f"被剔除的批次列表: {completely_removed_batches}")
        else:
            print("没有批次被完全剔除，因此不会生成'首放一致性差待复测'表格")

        # 确保即使没有完全剔除的批次，也生成统计数据
        # 排序
        filtered_data = filtered_data.sort_values(['系列', '统一批次'], ascending=True)
        filtered_data.index = range(1, len(filtered_data) + 1)

        # 修改: 分析被完全剔除的批次，更新异常判断标准
        problem_batches = []

        for batch in completely_removed_batches:
            batch_data = self.all_cycle_data[self.all_cycle_data['统一批次'] == batch]

            # 如果批次为空，跳过
            if batch_data.empty:
                continue

            # 修改: 计算各类异常样品比例，考虑三级首效状态 - 使用常量替代硬编码值
            very_low_eff_count = sum(
                (batch_data['1C状态'] == '首效过低') |
                ((batch_data['首效'] < CONFIG["ONE_C_THRESHOLDS"]["very_low_efficiency_threshold"]) & pd.notna(batch_data['首效']))
            )
            # 计算首效低的样品数量
            low_eff_count = sum(
                (batch_data['1C状态'] == '首效低') |
                ((batch_data['首效'] >= CONFIG["ONE_C_THRESHOLDS"]["very_low_efficiency_threshold"]) & (batch_data['首效'] < CONFIG["ONE_C_THRESHOLDS"]["low_efficiency_threshold"]) & pd.notna(batch_data['首效']))
            )
            # 注意：这里只计算数量，而在_calculate_cycle_statistics方法中会创建low_eff_samples DataFrame
            overcharge_count = sum(batch_data['1C状态'] == '1C过充')
            total_count = len(batch_data)

            # 判断是否为真正的问题批次 (超过50%样品异常)
            # 修改: 只计算严重异常(过充或首效过低<80%)的比例
            if total_count > 0 and (very_low_eff_count + overcharge_count) / total_count > 0.5:
                problem_batches.append(batch)
                print(f"批次{batch}有{very_low_eff_count}个首效过低(<{CONFIG['ONE_C_THRESHOLDS']['very_low_efficiency_threshold']}%)样品和{overcharge_count}个过充样品，"
                    f"占总数{total_count}的{(very_low_eff_count + overcharge_count)*100/total_count:.1f}%")

        # 按原始代码的方式处理数据：遍历所有原始批次，检查过滤后是否存在
        result_data = []

        # 获取所有批次
        all_series_batches = set(zip(self.all_cycle_data['系列'], self.all_cycle_data['统一批次']))

        # 对每个批次进行处理
        for series, batch in all_series_batches:
            # 检查过滤后的数据中是否有这个批次
            group_mask = (filtered_data['系列'] == series) & (filtered_data['统一批次'] == batch)

            if not filtered_data[group_mask].empty:
                # 批次数据正常，计算统计值
                group = filtered_data[group_mask]
                stats = self._calculate_basic_statistics(series, batch, group)
                self._calculate_cycle_statistics(stats, group)
                result_data.append(stats)
            else:
                # 批次数据异常，需要复测
                # 区分是真正的问题批次还是数据波动大
                mask = ((self.all_cycle_data['系列'] == series) &
                    (self.all_cycle_data['统一批次'] == batch))
                inconsistent = self.all_cycle_data[mask]

                # 添加到首放一致性差待复测表
                self.inconsistent_data = pd.concat(
                    [self.inconsistent_data, inconsistent], ignore_index=True)

                # 根据批次分析结果输出原因
                if batch in problem_batches:
                    print(f"批次{batch}(系列{series})由于异常高比例(>50%)的严重低首效(<{CONFIG['ONE_C_THRESHOLDS']['very_low_efficiency_threshold']}%)或过充被归类为需要复测")
                else:
                    print(f"批次{batch}(系列{series})数据波动过大被归类为需要复测")

        # 转换为DataFrame并排序
        if result_data:
            self.statistics_data = pd.DataFrame(result_data)
            self.statistics_data = self.statistics_data.sort_values(['系列', '统一批次'], ascending=True)
            self.statistics_data.index = range(1, len(self.statistics_data) + 1)
            print(f"成功生成统计数据，共{len(self.statistics_data)}行")
        else:
            # 即使没有有效的统计数据，也创建一个空的DataFrame
            self.statistics_data = pd.DataFrame(columns=CONFIG["EXCEL_COLS"].get('statistics',
                ['系列', '统一批次', '样品数', '首放平均值', '首效平均值', '首放标准差', '首效标准差',
                 '首放极差', '首效极差', '参考通道', '参考通道首放', '参考通道首效']))
            print("没有有效的统计数据，创建空的统计数据DataFrame")

        # 在方法结束前添加调试信息
        print(f"inconsistent_data的行数: {len(self.inconsistent_data)}")
        if self.inconsistent_data.empty:
            print("inconsistent_data为空，不会生成'首放一致性差待复测'表格")
        else:
            print(f"共有{len(self.inconsistent_data)}行数据将写入'首放一致性差待复测'表格")



    def _remove_outliers(self, df, column):
        """
        使用改良箱线图法去除异常值

        Parameters:
        -----------
        df: DataFrame
            待处理的数据
        column: str
            用于判断异常的列名

        Returns:
        --------
        tuple: (处理后的DataFrame, 完全被剔除的批次列表)
        """
        if df.empty:
            return df, []

        # 获取极差阈值和收缩因子（用于改良箱线图法）
        max_range = None
        if column == '首放' and '首放_极差阈值' in CONFIG["OUTLIER_DETECTION"]:
            max_range = CONFIG["OUTLIER_DETECTION"]['首放_极差阈值']
        elif column == '首效' and '首效_极差阈值' in CONFIG["OUTLIER_DETECTION"]:
            max_range = CONFIG["OUTLIER_DETECTION"]['首效_极差阈值']

        shrink_factor = CONFIG["OUTLIER_DETECTION"].get('boxplot_shrink_factor', 0.95)

        # 调试输出：开始异常值检测
        if self.verbose:
            print(f"\n开始使用改良箱线图法检测{column}列的异常值，极差阈值={max_range}")
            print(f"处理前的数据行数: {len(df)}")
            print(f"处理前的通道数量: {len(df[['主机', '通道']].drop_duplicates())}")

        result_df = pd.DataFrame(columns=df.columns)
        completely_removed_batches = []
        removed_channels = []  # 记录被剔除的通道

        # 按批次分组处理
        for batch, group_df in df.groupby('统一批次'):
            # 调试输出：当前处理的批次
            if self.verbose:
                print(f"\n处理批次: {batch}, 包含 {len(group_df)} 个样品")
                print(f"批次 {batch} 的通道列表:")
                for i, (_, row) in enumerate(group_df[['主机', '通道']].drop_duplicates().iterrows()):
                    print(f"  {i+1}. 主机={row['主机']}, 通道={row['通道']}")

            # 如果只有一个样品，则保留
            if len(group_df) <= 1:
                result_df = pd.concat([result_df, group_df])
                continue

            # 使用改良箱线图法
            if max_range is not None:
                # 初始化
                filtered_df = group_df.copy()
                current_range = filtered_df[column].max() - filtered_df[column].min()
                iteration = 0
                max_iterations = CONFIG["RUNTIME_CONFIG"]["max_iterations"]  # 最大迭代次数

                # 调试输出
                if self.verbose:
                    print(f"  使用改良箱线图法，初始极差={current_range:.2f}, 目标极差={max_range}")

                # 迭代直到满足极差要求或达到最大迭代次数
                while current_range > max_range and iteration < max_iterations and len(filtered_df) > 1:
                    iteration += 1
                    prev_size = len(filtered_df)

                    # 计算四分位数
                    q1, q3 = filtered_df[column].quantile([0.25, 0.75])
                    iqr = q3 - q1

                    # 计算边界
                    lower_bound = q1 - iqr * (shrink_factor ** iteration)
                    upper_bound = q3 + iqr * (shrink_factor ** iteration)

                    # 调试输出
                    if self.verbose:
                        print(f"  迭代 {iteration}: Q1={q1:.2f}, Q3={q3:.2f}, IQR={iqr:.2f}")
                        print(f"  边界: [{lower_bound:.2f}, {upper_bound:.2f}]")

                    # 过滤数据
                    outliers_before = filtered_df[~filtered_df[column].between(lower_bound, upper_bound)]
                    filtered_df = filtered_df[filtered_df[column].between(lower_bound, upper_bound)]

                    # 记录被剔除的通道
                    for _, row in outliers_before.iterrows():
                        removed_channels.append((row['主机'], row['通道']))

                    # 计算新的极差
                    if not filtered_df.empty:
                        current_range = filtered_df[column].max() - filtered_df[column].min()

                    # 调试输出
                    if self.verbose and not outliers_before.empty:
                        print(f"  剔除了 {len(outliers_before)} 个异常值:")
                        for _, row in outliers_before.iterrows():
                            print(f"    异常值: 主机={row['主机']}, 通道={row['通道']}, {column}={row[column]:.2f}")

                    # 如果没有点被剔除，退出循环
                    if len(filtered_df) == prev_size:
                        break

                # 调试输出
                if self.verbose:
                    print(f"  最终极差={current_range:.2f}, 迭代次数={iteration}")

            else:
                # 如果没有设置极差阈值，保留原数据
                filtered_df = group_df.copy()
                if self.verbose:
                    print(f"  未设置{column}的极差阈值，保留所有数据")

            # 如果过滤后为空，记录该批次
            if filtered_df.empty:
                completely_removed_batches.append(batch)
                if self.verbose:
                    print(f"  批次 {batch} 的所有样品都被剔除")
            else:
                result_df = pd.concat([result_df, filtered_df])
                if self.verbose:
                    print(f"  批次 {batch} 保留了 {len(filtered_df)} 个样品")

        # 调试输出：处理结果
        if self.verbose:
            print(f"\n异常值检测结果:")
            print(f"处理后的数据行数: {len(result_df)}")
            print(f"处理后的通道数量: {len(result_df[['主机', '通道']].drop_duplicates())}")
            print(f"完全被剔除的批次数量: {len(completely_removed_batches)}")
            if completely_removed_batches:
                print(f"完全被剔除的批次: {completely_removed_batches}")
            print(f"被剔除的通道数量: {len(set(removed_channels))}")
            if removed_channels:
                print(f"被剔除的通道:")
                for host, channel in set(removed_channels):
                    print(f"  主机={host}, 通道={channel}")

        return result_df, completely_removed_batches

    def _remove_outliers_zscore_mad(self, data):
        """
        使用Z-score+MAD方法去除异常点

        Args:
            data: 数据DataFrame

        Returns:
            tuple: (去除异常点后的DataFrame, 完全被剔除的批次列表)
        """
        print("开始使用Z-score+MAD方法去除异常点")

        # 获取配置 - 直接从CONFIG获取最新配置
        config = CONFIG["OUTLIER_DETECTION"]['zscore_mad']
        mad_constant = config['mad_constant']
        min_mad_ratio = config['min_mad_ratio']
        thresholds = config['thresholds']

        # 调试输出：显示当前使用的配置
        print(f"当前Z-score配置: MAD常数={mad_constant}, MAD最小比例={min_mad_ratio}, 阈值={thresholds}")

        # 记录配置到异常检测日志
        self.logger.log_outlier_detection("Z-score+MAD异常检测开始")
        self.logger.log_outlier_detection(f"配置参数: MAD常数={mad_constant}, MAD最小比例={min_mad_ratio}")
        self.logger.log_outlier_detection(f"阈值设置: {thresholds}")

        # 记录到调试日志
        self.logger.log_debug("开始Z-score+MAD异常检测")
        self.logger.log_debug(f"输入数据行数: {len(data)}")
        self.logger.log_debug(f"数据列: {data.columns.tolist()}")
        self.logger.log_debug(f"配置详情: {config}")

        use_time_series = config['use_time_series']
        min_samples_for_stl = config['min_samples_for_stl']

        # 按统一批次分组处理
        grouped = data.groupby('统一批次')
        result_data = pd.DataFrame()
        completely_removed_batches = []

        for batch_name, batch_data in grouped:
            print(f"处理批次: {batch_name}, 原始数据点: {len(batch_data)}")

            # 记录到异常检测日志
            self.logger.log_outlier_detection(f"开始处理批次: {batch_name}")
            self.logger.log_outlier_detection(f"原始数据点: {len(batch_data)}")

            # 记录到调试日志
            self.logger.log_debug(f"处理批次: {batch_name}")
            self.logger.log_debug(f"批次数据形状: {batch_data.shape}")
            self.logger.log_debug(f"批次数据列: {batch_data.columns.tolist()}")

            # 记录批次中的所有通道
            channels_info = []
            for idx, row in batch_data.iterrows():
                channel_info = f"{row['主机']}-{row['通道']}"
                channels_info.append(channel_info)
            self.logger.log_outlier_detection(f"批次中的通道: {', '.join(channels_info)}")

            if len(batch_data) < 2:
                print(f"批次 {batch_name} 数据点太少，跳过异常检测")
                self.logger.log_outlier_detection(f"批次 {batch_name} 数据点太少，跳过异常检测")
                result_data = pd.concat([result_data, batch_data])
                continue

            current_data = batch_data.copy()
            outliers_mask = pd.Series([False] * len(current_data), index=current_data.index)
            removed_channels = []  # 记录被剔除的通道

            # 对每个指标进行异常检测
            for metric, threshold in thresholds.items():
                if metric not in current_data.columns:
                    print(f"警告: 批次 {batch_name} 中缺少列 {metric}，跳过该指标")
                    continue

                values = current_data[metric].dropna()
                if len(values) < 2:
                    continue

                # 计算中位数和MAD
                median = values.median()
                mad = (values - median).abs().median()

                # 应用MAD最小值限制
                min_mad_threshold = median * min_mad_ratio
                original_mad = mad
                if mad < min_mad_threshold:
                    mad = min_mad_threshold
                    print(f"批次 {batch_name} 指标 {metric}: 原始MAD={original_mad:.4f} 过小，调整为最小值={mad:.4f}")
                    # 记录MAD调整到调试日志
                    self.logger.log_debug(f"批次 {batch_name} 指标 {metric}: MAD调整 {original_mad:.4f} -> {mad:.4f}")
                else:
                    # 记录正常MAD值到调试日志
                    self.logger.log_debug(f"批次 {batch_name} 指标 {metric}: MAD={mad:.4f} (未调整)")

                if mad == 0:
                    print(f"批次 {batch_name} 指标 {metric} 的MAD为0，跳过异常检测")
                    continue

                # 计算修正的Z-score
                modified_z_scores = mad_constant * (values - median) / mad

                # 时间序列分解（可选）
                if use_time_series and len(values) > min_samples_for_stl:
                    try:
                        from scipy import signal
                        # 简单的趋势去除
                        detrended = signal.detrend(values.values)
                        detrended_median = np.median(detrended)
                        detrended_mad = np.median(np.abs(detrended - detrended_median))

                        if detrended_mad > 0:
                            detrended_z_scores = mad_constant * (detrended - detrended_median) / detrended_mad
                            # 结合原始Z-score和去趋势Z-score
                            combined_z_scores = np.maximum(np.abs(modified_z_scores), np.abs(detrended_z_scores))
                        else:
                            combined_z_scores = np.abs(modified_z_scores)
                    except Exception as e:
                        print(f"时间序列分解失败: {e}，使用标准Z-score方法")
                        combined_z_scores = np.abs(modified_z_scores)
                else:
                    combined_z_scores = np.abs(modified_z_scores)

                # 标记异常值
                metric_outliers = combined_z_scores > threshold
                outliers_mask.loc[values.index] |= metric_outliers

                outlier_count = metric_outliers.sum()
                if outlier_count > 0:
                    print(f"批次 {batch_name} 指标 {metric}: 检测到 {outlier_count} 个异常值 (阈值: {threshold})")

                    # 记录详细的异常信息
                    self.logger.log_outlier_detection(f"指标 {metric} 异常检测结果:")
                    self.logger.log_outlier_detection(f"  阈值: {threshold}")
                    self.logger.log_outlier_detection(f"  中位数: {median:.3f}")
                    self.logger.log_outlier_detection(f"  MAD: {mad:.3f}")

                    # 记录每个异常样本的详细信息
                    outlier_indices = values.index[metric_outliers]
                    for i, outlier_idx in enumerate(outlier_indices):
                        outlier_row = current_data.loc[outlier_idx]
                        channel_id = f"{outlier_row['主机']}-{outlier_row['通道']}"
                        outlier_value = outlier_row[metric]

                        # 安全地获取Z-score值
                        try:
                            # 找到outlier_idx在values.index中的位置
                            outlier_position = list(values.index).index(outlier_idx)
                            z_score = combined_z_scores.iloc[outlier_position]
                        except (ValueError, IndexError):
                            # 如果找不到对应位置，直接从metric_outliers中获取
                            outlier_positions = np.where(metric_outliers)[0]
                            if i < len(outlier_positions):
                                z_score = combined_z_scores.iloc[outlier_positions[i]]
                            else:
                                z_score = float('nan')

                        self.logger.log_outlier_detection(f"    异常样本: {channel_id}")
                        self.logger.log_outlier_detection(f"      {metric}值: {outlier_value:.3f}")
                        if not np.isnan(z_score):
                            self.logger.log_outlier_detection(f"      Z-score: {z_score:.3f}")
                        else:
                            self.logger.log_outlier_detection(f"      Z-score: 计算错误")

                        # 记录到被剔除通道列表（避免重复）
                        if channel_id not in removed_channels:
                            removed_channels.append(channel_id)

            # 移除异常值
            clean_data = current_data[~outliers_mask]
            removed_count = len(current_data) - len(clean_data)

            if removed_count > 0:
                print(f"批次 {batch_name} 总共去除 {removed_count} 个异常点")

                # 记录被剔除的通道汇总
                self.logger.log_outlier_detection(f"批次 {batch_name} 异常检测汇总:")
                self.logger.log_outlier_detection(f"  原始样本数: {len(current_data)}")
                self.logger.log_outlier_detection(f"  剔除样本数: {removed_count}")
                self.logger.log_outlier_detection(f"  剩余样本数: {len(clean_data)}")

                if removed_channels:
                    self.logger.log_outlier_detection(f"  被剔除的通道: {', '.join(set(removed_channels))}")

                    # 记录剩余的通道
                    remaining_channels = []
                    for idx, row in clean_data.iterrows():
                        channel_id = f"{row['主机']}-{row['通道']}"
                        remaining_channels.append(channel_id)
                    self.logger.log_outlier_detection(f"  剩余的通道: {', '.join(remaining_channels)}")
            else:
                print(f"批次 {batch_name} 未检测到异常点")
                self.logger.log_outlier_detection(f"批次 {batch_name} 未检测到异常点")

            # 如果批次被完全剔除，记录
            if clean_data.empty:
                completely_removed_batches.append(batch_name)
                print(f"批次 {batch_name} 的所有样品都被剔除")
                self.logger.log_outlier_detection(f"批次 {batch_name} 的所有样品都被剔除")
            else:
                result_data = pd.concat([result_data, clean_data])

            self.logger.log_outlier_detection(f"批次 {batch_name} 处理完成\n" + "="*50)

        # 重置索引
        result_data = result_data.sort_values('统一批次', ascending=True)
        result_data.index = range(1, len(result_data) + 1)

        total_removed = len(data) - len(result_data)
        print(f"Z-score+MAD方法完成: 总共去除 {total_removed} 个异常点")

        # 记录总结信息
        self.logger.log_outlier_detection(f"\nZ-score+MAD异常检测总结:")
        self.logger.log_outlier_detection(f"  处理批次数: {len(grouped)}")
        self.logger.log_outlier_detection(f"  原始样本总数: {len(data)}")
        self.logger.log_outlier_detection(f"  剔除样本总数: {total_removed}")
        self.logger.log_outlier_detection(f"  剩余样本总数: {len(result_data)}")
        self.logger.log_outlier_detection(f"  完全剔除的批次数: {len(completely_removed_batches)}")
        if completely_removed_batches:
            self.logger.log_outlier_detection(f"  完全剔除的批次: {', '.join(completely_removed_batches)}")

        return result_data, completely_removed_batches

    def _generate_zscore_distribution_plots(self, data, output_dir):
        """
        生成Z-score分布图并保存到数据可视化文件夹

        Args:
            data: 原始数据DataFrame
            output_dir: 输出目录
        """
        print("开始生成Z-score分布图...")

        # 创建可视化文件夹
        timestamp = time.strftime('%Y%m%d_%H%M%S', time.localtime())
        viz_dir = os.path.join(output_dir, f"数据可视化-{timestamp}")
        zscore_dir = os.path.join(viz_dir, "Z-score分布图")
        os.makedirs(zscore_dir, exist_ok=True)

        # 获取配置 - 直接从CONFIG获取最新配置
        config = CONFIG["OUTLIER_DETECTION"]['zscore_mad']
        mad_constant = config['mad_constant']
        min_mad_ratio = config['min_mad_ratio']
        thresholds = config['thresholds']

        # 按统一批次分组处理
        grouped = data.groupby('统一批次')

        for batch_name, batch_data in grouped:
            if len(batch_data) < 2:
                continue

            # 创建图表
            fig, axes = plt.subplots(2, 2, figsize=(15, 12))
            fig.suptitle(f'批次 {batch_name} 的修正Z-score分布图', fontsize=16, fontweight='bold')

            axes = axes.flatten()
            plot_idx = 0

            for metric, threshold in thresholds.items():
                if metric not in batch_data.columns or plot_idx >= 4:
                    continue

                values = batch_data[metric].dropna()
                if len(values) < 2:
                    continue

                # 计算修正Z-score
                median = values.median()
                mad = (values - median).abs().median()

                # 应用MAD最小值限制（与异常检测保持一致）
                min_mad_threshold = median * min_mad_ratio
                if mad < min_mad_threshold:
                    mad = min_mad_threshold

                if mad == 0:
                    continue

                modified_z_scores = mad_constant * (values - median) / mad

                # 绘制分布图
                ax = axes[plot_idx]

                # 绘制直方图
                ax.hist(modified_z_scores, bins=20, alpha=0.7,
                       color='skyblue', edgecolor='black', density=True)

                # 标记阈值线
                ax.axvline(threshold, color='red', linestyle='--', linewidth=2,
                          label=f'正阈值: {threshold}')
                ax.axvline(-threshold, color='red', linestyle='--', linewidth=2,
                          label=f'负阈值: -{threshold}')

                # 标记异常点
                outliers = np.abs(modified_z_scores) > threshold
                outlier_count = outliers.sum()

                if outlier_count > 0:
                    outlier_scores = modified_z_scores[outliers]
                    outlier_indices = values.index[outliers]

                    # 绘制异常点
                    ax.scatter(outlier_scores, [0.02] * len(outlier_scores),
                             color='red', s=50, marker='x',
                             label=f'异常点: {outlier_count}个')

                    # 为异常点添加通道标签（只在异常点不太多时显示）
                    if outlier_count <= 5:  # 避免标签过多导致图表混乱
                        for i, (score, idx) in enumerate(zip(outlier_scores, outlier_indices)):
                            try:
                                outlier_row = batch_data.loc[idx]
                                channel_label = f"{outlier_row['主机']}-{outlier_row['通道']}"
                                ax.annotate(channel_label,
                                          (score, 0.02),
                                          xytext=(5, 10),
                                          textcoords='offset points',
                                          fontsize=8,
                                          color='red',
                                          ha='left')
                            except (KeyError, IndexError):
                                # 如果索引有问题，跳过这个标签
                                continue

                # 设置标题和标签
                ax.set_title(f'{metric}\n(中位数: {median:.2f}, MAD: {mad:.2f})',
                           fontsize=12, fontweight='bold')
                ax.set_xlabel('修正Z-score', fontsize=10)
                ax.set_ylabel('密度', fontsize=10)
                ax.legend(fontsize=9)
                ax.grid(True, alpha=0.3)

                # 添加统计信息
                stats_text = f'样本数: {len(values)}\n异常数: {outlier_count}\n异常率: {outlier_count/len(values)*100:.1f}%'

                # 如果有异常点且数量不太多，添加异常通道列表
                if outlier_count > 0 and outlier_count <= 3:
                    outlier_channels = []
                    for idx in outlier_indices:
                        try:
                            outlier_row = batch_data.loc[idx]
                            channel_label = f"{outlier_row['主机']}-{outlier_row['通道']}"
                            outlier_channels.append(channel_label)
                        except (KeyError, IndexError):
                            # 如果索引有问题，跳过这个通道
                            continue
                    if outlier_channels:  # 只有成功获取到通道信息时才显示
                        stats_text += f'\n异常通道:\n' + '\n'.join(outlier_channels)
                elif outlier_count > 3:
                    stats_text += f'\n异常通道过多\n详见日志文件'

                ax.text(0.02, 0.98, stats_text, transform=ax.transAxes,
                       verticalalignment='top', fontsize=8,
                       bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.8))

                plot_idx += 1

            # 隐藏未使用的子图
            for i in range(plot_idx, 4):
                axes[i].set_visible(False)

            # 调整布局
            plt.tight_layout()

            # 保存图片
            safe_batch_name = "".join(c for c in batch_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            if not safe_batch_name:  # 如果清理后为空，使用默认名称
                safe_batch_name = f"batch_{hash(batch_name) % 10000}"

            filename = f"Z-score分布图_{safe_batch_name}_{timestamp}.png"
            filepath = os.path.join(zscore_dir, filename)

            # 确保文件路径长度不超过系统限制
            if len(filepath) > 250:
                filename = f"Z-score分布图_{timestamp}_{hash(batch_name) % 10000}.png"
                filepath = os.path.join(zscore_dir, filename)

            try:
                # 尝试使用较低的DPI保存
                plt.savefig(filepath, dpi=150, bbox_inches='tight', format='png')
                print(f"已保存Z-score分布图: {filename}")
            except Exception as e:
                print(f"保存Z-score分布图失败: {e}")
                # 尝试备用保存方法
                try:
                    backup_filename = f"Z-score分布图_backup_{timestamp}.png"
                    backup_filepath = os.path.join(zscore_dir, backup_filename)
                    plt.savefig(backup_filepath, dpi=100, format='png')
                    print(f"使用备用方法保存: {backup_filename}")
                except Exception as e2:
                    print(f"备用保存方法也失败: {e2}")

            plt.close(fig)

        print(f"Z-score分布图已保存到: {zscore_dir}")



    def _select_reference_channel_by_pca(self, samples, features=None):
        """使用PCA选择最具代表性的样本

        Args:
            samples (pd.DataFrame): 待选样本
            features (list): 用于PCA的特征列表，默认为None时使用预设特征

        Returns:
            pd.Series: 参考通道数据
        """
        # 如果未指定特征，使用默认特征集
        if features is None:
            features = CONFIG["REFERENCE_CHANNEL_CONFIG"]["pca"]["default_features"]

        # 确保有足够的样本和特征
        available_features = [f for f in features if f in samples.columns]
        if len(available_features) < 2 or len(samples) < 3:
            print("可用特征或样本不足，回退到传统方法")
            return self._select_reference_channel_from_subset(samples, use_multi_feature=False)

        # 准备数据
        X = samples[available_features].copy()
        X = X.fillna(X.median())

        # 标准化
        scaler = StandardScaler()
        X_scaled = scaler.fit_transform(X)

        # 应用PCA
        pca = PCA(n_components=CONFIG["REFERENCE_CHANNEL_CONFIG"]["pca"]["n_components"])
        principal_components = pca.fit_transform(X_scaled)

        # 计算到中心点的距离
        center = np.mean(principal_components, axis=0)
        distances = np.sqrt(np.sum((principal_components - center)**2, axis=1))

        # 找到距离中心最近的样本
        most_central_idx = np.argmin(distances)
        reference_sample = samples.iloc[most_central_idx]

        # 可选：可视化PCA结果
        if CONFIG["REFERENCE_CHANNEL_CONFIG"]["pca"]["visualization_enabled"]:
            # 获取批次信息用于命名
            batch_name = "未知批次"
            if '统一批次' in samples.columns and not samples['统一批次'].empty:
                batch_name = samples['统一批次'].iloc[0]
            elif '批次' in samples.columns and not samples['批次'].empty:
                batch_name = samples['批次'].iloc[0]

            self._visualize_pca_result(principal_components, most_central_idx, samples.index, batch_name=batch_name)

        print(f"PCA参考通道选择: 使用特征 {available_features}")
        print(f"选择的参考通道: {reference_sample['主机']}-{reference_sample['通道']}")

        return reference_sample

    def _select_reference_channel_by_capacity_retention(self, data):
        """使用容量保留率曲线比较选择参考通道

        通过比较每个通道的容量保留率曲线与批次平均曲线的均方误差(MSE)，
        选择MSE最小的通道作为参考通道，即最能代表批次整体表现的通道。

        实现了以下理论框架：
        1. 提取每个通道从1C首圈到当前圈数的容量保留率数据
        2. 确定共同的循环范围，确保公平比较
        3. 计算平均容量保留率曲线
        4. 计算每个通道与平均曲线的加权均方误差(MSE)
        5. 选择MSE最小的通道作为参考通道

        Args:
            data: 候选参考通道数据

        Returns:
            参考通道数据或None
        """
        if data.empty or len(data) < 2:
            return None if data.empty else data.iloc[0]

        print(f"\n开始容量保留率曲线比较分析...")

        # 获取配置
        config = CONFIG["REFERENCE_CHANNEL_CONFIG"]["capacity_retention"]

        # 检查通道数量是否满足最小要求
        if len(data) < config["min_channels"]:
            print(f"通道数量不足({len(data)} < {config['min_channels']})，无法进行有效比较")
            return None

        try:
            # 尝试查找处理中的批次信息，便于调试
            batch_info = "未知批次"
            if '批次' in data.columns and not data['批次'].empty:
                batch_info = data['批次'].iloc[0]
            elif '统一批次' in data.columns and not data['统一批次'].empty:
                batch_info = data['统一批次'].iloc[0]

            print(f"正在处理批次: {batch_info}")

            # 打印所有文件名/文件路径信息，便于调试
            if '文件名' in data.columns:
                print("批次包含的文件:")
                for i, fname in enumerate(data['文件名'].unique()):
                    print(f"  {i+1}. {fname}")
            elif '文件路径' in data.columns:
                print("批次包含的文件路径:")
                for i, fpath in enumerate(data['文件路径'].unique()):
                    print(f"  {i+1}. {fpath}")

            # 确定是否使用原始放电容量数据重新计算保留率
            if config["use_raw_capacity"] and '1C首圈编号' in data.columns:
                print("使用原始放电容量数据重新计算保留率...")
                # 直接使用实例的folder_path
                if self.folder_path and os.path.exists(self.folder_path):
                    print(f"使用文件夹路径: {self.folder_path}")
                    return self._select_reference_channel_by_raw_capacity(data, self.folder_path)
                else:
                    print("警告: 无法获取有效的文件夹路径，使用预计算的容量保留率数据")
                    return self._select_reference_channel_by_retention_columns(data)
            else:
                print("使用预计算的容量保留率数据...")
                return self._select_reference_channel_by_retention_columns(data)
        except Exception as e:
            import traceback
            print(f"容量保留率曲线比较分析过程中出错: {str(e)}")
            print(traceback.format_exc())

            # 尝试获取当前正在处理的文件信息
            try:
                if '文件名' in data.columns:
                    filenames = data['文件名'].unique()
                    print(f"出错批次包含的文件: {filenames}")
                elif '文件路径' in data.columns:
                    filepaths = data['文件路径'].unique()
                    print(f"出错批次包含的文件路径: {filepaths}")

                # 打印数据的主要列信息，以帮助调试
                print(f"数据包含的列: {data.columns.tolist()}")
                print(f"数据行数: {len(data)}")
                if len(data) > 0:
                    print("第一行数据样例:")
                    for col in ['主机', '通道', '批次', '统一批次', '1C首圈编号', '当前圈数']:
                        if col in data.columns:
                            print(f"  {col}: {data[col].iloc[0]}")
            except Exception as debug_error:
                print(f"尝试获取调试信息时出错: {str(debug_error)}")

            # 出错时返回None
            print("由于错误，返回None")
            return None

    def _select_reference_channel_by_retention_columns(self, data):
        """使用预计算的容量保留率列选择参考通道

        Args:
            data: 候选参考通道数据

        Returns:
            参考通道数据或None
        """
        print("\n======开始执行参考通道选择过程======")
        print(f"数据行数: {len(data)}")

        # 打印所有通道信息以便调试
        print("\n所有候选参考通道信息:")
        for i, (idx, row) in enumerate(data.iterrows()):
            channel_info = f"#{i+1} 索引={idx}, 主机={row.get('主机', 'N/A')}, 通道={row.get('通道', 'N/A')}"
            batch_info = f"批次={row.get('统一批次', 'N/A')}"
            file_info = ""
            if '文件名' in data.columns:
                file_info = f", 文件名={row.get('文件名', 'N/A')}"
            elif '文件路径' in data.columns:
                file_info = f", 文件路径={row.get('文件路径', 'N/A')}"
            print(f"{channel_info}, {batch_info}{file_info}")

        # 获取配置
        config = CONFIG["REFERENCE_CHANNEL_CONFIG"]["capacity_retention"]
        retention_columns = config["retention_columns"]

        # 检查是否有足够的容量保留率数据
        available_columns = [col for col in retention_columns if col in data.columns]
        if not available_columns:
            print("没有可用的容量保留率数据列")
            return None

        print(f"使用的容量保留率列: {available_columns}")

        # 提取容量保留率数据
        retention_data = {}
        for channel_id, channel_data in data.groupby(['主机', '通道']):
            # 创建通道标识
            channel_key = f"{channel_id[0]}-{channel_id[1]}"

            # 提取该通道的容量保留率数据
            channel_retention = {}
            for col in available_columns:
                if col in channel_data.columns and pd.notna(channel_data[col].iloc[0]):
                    # 提取循环次数和容量保留率
                    if col == "当前容量保持" and "当前圈数" in channel_data.columns:
                        cycle = channel_data["当前圈数"].iloc[0]
                        if pd.notna(cycle):
                            channel_retention[int(cycle)] = channel_data[col].iloc[0]
                    elif col == "100容量保持":
                        channel_retention[100] = channel_data[col].iloc[0]
                    elif col == "200容量保持":
                        channel_retention[200] = channel_data[col].iloc[0]

            # 只有当有足够的数据点时才添加
            if len(channel_retention) >= 2:
                retention_data[channel_key] = channel_retention

        # 检查是否有足够的通道数据
        if len(retention_data) < 2:
            print(f"没有足够的通道数据进行比较，只有{len(retention_data)}个通道有容量保留率数据")
            # 如果只有一个通道有数据，直接返回该通道
            if len(retention_data) == 1:
                channel_key = list(retention_data.keys())[0]
                print(f"只有一个通道有数据，尝试返回: {channel_key}")

                # 使用统一的通道解析方法
                host, channel = self._extract_host_and_channel(channel_key)

                if host and channel:
                    match = self._try_match_channel(host, channel, data)
                    if match is not None:
                        print(f"通道匹配成功! 主机={host}, 通道={channel}")
                        return match

                # 如果所有尝试都失败，回退到旧方法
                channel_parts = channel_key.split('-')
                match_data = data[(data['主机'] == channel_parts[0]) & (data['通道'] == '-'.join(channel_parts[1:]))]
                if not match_data.empty:
                    return match_data.iloc[0]
            return None

        # 计算所有通道的循环次数范围
        all_cycles = set()
        for channel_data in retention_data.values():
            all_cycles.update(channel_data.keys())

        # 按循环次数排序
        all_cycles = sorted(all_cycles)
        print(f"发现的循环次数: {all_cycles}")

        # 动态确定循环范围或使用配置的固定范围
        if config["dynamic_range"] and len(all_cycles) >= 2:
            # 找到所有通道共有的最小循环范围
            min_cycle = min(all_cycles)
            max_cycle = max(all_cycles)
            print(f"动态确定的循环范围: {min_cycle} - {max_cycle}")
        else:
            # 使用配置的固定范围
            min_cycle = max(min(all_cycles), config["min_cycles"])
            max_cycle = min(max(all_cycles), config["max_cycles"])
            print(f"配置的循环范围: {min_cycle} - {max_cycle}")

        # 如果循环次数太少，无法进行有效比较
        if max_cycle - min_cycle + 1 < config["min_cycles"]:
            print(f"有效循环范围太小({max_cycle - min_cycle + 1} < {config['min_cycles']})，无法进行有效比较")
            return None

        # 创建插值函数，为每个通道生成完整的容量保留率曲线
        from scipy import interpolate
        interpolated_curves = {}

        # 定义插值的循环次数范围
        interp_cycles = np.arange(min_cycle, max_cycle + 1, config["cycle_step"])

        for channel_key, channel_data in retention_data.items():
            # 只有当通道有足够的数据点时才进行插值
            if len(channel_data) >= 2:
                cycles = np.array(list(channel_data.keys()))
                retentions = np.array(list(channel_data.values()))

                # 创建插值函数
                if len(cycles) >= 3 and config["interpolation_method"] == "cubic":
                    # 至少需要3个点才能进行三次样条插值
                    f = interpolate.interp1d(cycles, retentions, kind='cubic',
                                            bounds_error=False, fill_value="extrapolate")
                else:
                    # 否则使用线性插值
                    f = interpolate.interp1d(cycles, retentions, kind='linear',
                                            bounds_error=False, fill_value="extrapolate")

                # 生成插值曲线
                interpolated_curve = f(interp_cycles)
                interpolated_curves[channel_key] = interpolated_curve

        # 计算平均曲线
        if interpolated_curves:
            all_curves = np.array(list(interpolated_curves.values()))
            mean_curve = np.mean(all_curves, axis=0)

            # 计算每个通道与平均曲线的均方误差(MSE)
            mse_scores = {}
            for channel_key, curve in interpolated_curves.items():
                # 根据配置决定是否使用加权MSE
                if config["use_weighted_mse"]:
                    # 计算权重
                    weights = self._calculate_cycle_weights(interp_cycles, min_cycle, max_cycle, config)
                    # 计算加权MSE
                    mse = np.average((curve - mean_curve) ** 2, weights=weights)
                else:
                    # 计算普通MSE
                    mse = np.mean((curve - mean_curve) ** 2)

                mse_scores[channel_key] = mse

            # 找到MSE最小的通道
            best_channel_key = min(mse_scores, key=mse_scores.get)
            print(f"容量保留率曲线MSE最小的通道: {best_channel_key}, MSE={mse_scores[best_channel_key]:.4f}")

            # 可视化比较结果
            if self.verbose:
                self._visualize_capacity_retention_curves(interp_cycles, interpolated_curves,
                                                        mean_curve, best_channel_key,
                                                        config["use_weighted_mse"],
                                                        weights if config["use_weighted_mse"] else None)

            # 尝试在数据中查找最佳通道
            print(f"尝试在数据中查找通道: {best_channel_key}")

            try:
                # 使用统一的通道解析方法
                host, channel = self._extract_host_and_channel(best_channel_key)

                if host and channel:
                    print(f"解析结果: 主机={host}, 通道={channel}")
                    match = self._try_match_channel(host, channel, data)
                    if match is not None:
                        print(f"通道匹配成功! 主机={host}, 通道={channel}")
                        return match

                # 如果匹配失败，打印调试信息
                print(f"无法在数据中找到通道 {best_channel_key}，匹配失败")
                print("尝试返回第一个可用通道作为备选")
                if not data.empty:
                    return data.iloc[0]
            except Exception as e:
                import traceback
                print(f"通道匹配过程中出错: {str(e)}")
                print(traceback.format_exc())

                # 记录完整错误现场信息
                print(f"\n错误现场信息:")
                print(f"最佳通道标识: {best_channel_key}")
                print(f"数据行数: {len(data)}")
                if len(data) > 0:
                    print(f"数据列: {data.columns.tolist()}")
                    if '主机' in data.columns and '通道' in data.columns:
                        print("主机-通道组合:")
                        for idx, row in data.head(5).iterrows():
                            print(f"  {row['主机']}-{row['通道']}")

                print("由于错误，尝试返回第一个可用通道")
                if not data.empty:
                    return data.iloc[0]
                return None

        print("容量保留率曲线比较未找到合适的参考通道")
        print("回退到使用预计算的容量保留率列")
        return None  # 修改为直接返回None，避免无限递归

    def _select_reference_channel_by_raw_capacity(self, data, folder_path):
        """使用原始放电容量数据重新计算保留率并选择参考通道

        按照理论框架，从1C首圈到当前圈数提取放电容量数据，
        计算容量保留率 R_i(k) = C_i(k)/C_i(c_i) × 100%，
        其中 C_i(k) 是第 k 个循环的放电容量，c_i 是1C首圈的循环号。

        Args:
            data: 候选参考通道数据
            folder_path: 数据文件夹路径

        Returns:
            参考通道数据或None
        """
        if data.empty or len(data) < 2:
            return None if data.empty else data.iloc[0]

        print(f"\n开始从原始放电容量数据计算容量保留率曲线...")

        # 获取配置
        config = CONFIG["REFERENCE_CHANNEL_CONFIG"]["capacity_retention"]

        # 检查通道数量是否满足最小要求
        if len(data) < config["min_channels"]:
            print(f"通道数量不足({len(data)} < {config['min_channels']})，无法进行有效比较")
            return None

        # 检查是否有1C首圈编号列
        if '1C首圈编号' not in data.columns:
            print("数据中缺少'1C首圈编号'列，无法确定1C首圈")
            return self._select_reference_channel_by_retention_columns(data)

        # 提取每个通道的信息
        retention_curves = {}  # 存储每个通道的容量保留率曲线
        max_cycles = 0  # 记录最大循环次数

        # 遍历每个通道
        for channel_id, channel_data in data.groupby(['主机', '通道']):
            channel_key = f"{channel_id[0]}-{channel_id[1]}"

            # 获取1C首圈编号
            if pd.isna(channel_data['1C首圈编号'].iloc[0]):
                print(f"通道 {channel_key} 的1C首圈编号为空，跳过")
                continue

            one_c_cycle = int(channel_data['1C首圈编号'].iloc[0])
            print(f"通道 {channel_key} 的1C首圈编号: {one_c_cycle}")

            # 获取Excel文件路径
            file_path = None
            for index, row in channel_data.iterrows():
                # 尝试从数据中获取文件路径
                if hasattr(row, 'file_path') and pd.notna(row['file_path']):
                    file_path = row['file_path']
                    break

            if file_path is None:
                # 如果数据中没有文件路径，尝试在文件夹中查找匹配的文件
                device_id = channel_data['主机'].iloc[0]
                channel_id = channel_data['通道'].iloc[0]

                # 在文件夹中查找匹配的Excel文件 - 同时匹配主机和通道
                excel_files = []
                for f in os.listdir(folder_path):
                    if f.endswith('.xlsx'):
                        # 使用统一的主机通道解析方法
                        file_host, file_channel = self._extract_host_and_channel(f)
                        if file_host == device_id and file_channel == channel_id:
                            excel_files.append(f)

                if excel_files:
                    file_path = os.path.join(folder_path, excel_files[0])
                    print(f"找到匹配的Excel文件: {file_path}")
                else:
                    print(f"无法找到通道 {channel_key} 对应的Excel文件，跳过")
                    print(f"  需要匹配: 主机={device_id}, 通道={channel_id}")
                    continue

            # 读取Excel文件中的循环数据
            try:
                print(f"读取文件: {file_path}")
                cycle_df = pd.read_excel(file_path, sheet_name=CONFIG['cycle_sheet_name'],
                                        engine=CONFIG['excel_engine'])

                # 检查是否有必要的列
                required_cols = ['循环序号', '放电比容量(mAh/g)']
                if not all(col in cycle_df.columns for col in required_cols):
                    print(f"文件 {file_path} 缺少必要的列: {required_cols}")
                    continue

                # 提取放电容量数据
                required_cols = ['循环序号', '放电比容量(mAh/g)']
                optional_cols = []

                # 检查是否需要包含能量保留率
                include_energy = config.get("include_energy", False)
                if include_energy:
                    optional_cols.append('放电比能量(mWh/g)')

                # 检查是否需要包含电压保留率
                include_voltage = config.get("include_voltage", False)
                if include_voltage:
                    optional_cols.append('放电中值电压(V)')

                # 确定要提取的列
                cols_to_extract = required_cols.copy()
                for col in optional_cols:
                    if col in cycle_df.columns:
                        cols_to_extract.append(col)
                    else:
                        print(f"警告: 文件 {file_path} 中缺少列 {col}，将不包含该指标")

                # 提取数据
                cycle_data = cycle_df[cols_to_extract].copy()

                # 重命名列
                rename_dict = {'循环序号': 'cycle', '放电比容量(mAh/g)': 'discharge_capacity'}
                if '放电比能量(mWh/g)' in cycle_data.columns:
                    rename_dict['放电比能量(mWh/g)'] = 'discharge_energy'
                if '放电中值电压(V)' in cycle_data.columns:
                    rename_dict['放电中值电压(V)'] = 'discharge_voltage'
                cycle_data.rename(columns=rename_dict, inplace=True)

                # 过滤无效数据
                cycle_data = cycle_data.dropna(subset=['cycle', 'discharge_capacity'])
                if cycle_data.empty:
                    print(f"文件 {file_path} 没有有效的放电容量数据")
                    continue

                # 确保循环序号是整数
                cycle_data['cycle'] = cycle_data['cycle'].astype(int)

                # 获取1C首圈的数据
                one_c_rows = cycle_data[cycle_data['cycle'] == one_c_cycle]
                if one_c_rows.empty:
                    print(f"文件 {file_path} 中找不到1C首圈 {one_c_cycle} 的数据")
                    continue

                # 获取1C首圈的放电容量
                one_c_capacity = one_c_rows['discharge_capacity'].iloc[0]

                # 计算容量保留率
                cycle_data['capacity_retention'] = cycle_data['discharge_capacity'] / one_c_capacity * 100

                # 计算能量保留率（如果有数据）
                if 'discharge_energy' in cycle_data.columns:
                    one_c_energy = one_c_rows['discharge_energy'].iloc[0]
                    if pd.notna(one_c_energy) and one_c_energy > 0:
                        cycle_data['energy_retention'] = cycle_data['discharge_energy'] / one_c_energy * 100

                # 计算电压保留率（如果有数据）
                if 'discharge_voltage' in cycle_data.columns:
                    one_c_voltage = one_c_rows['discharge_voltage'].iloc[0]
                    if pd.notna(one_c_voltage) and one_c_voltage > 0:
                        cycle_data['voltage_retention'] = cycle_data['discharge_voltage'] / one_c_voltage * 100

                # 更新最大循环次数
                max_cycles = max(max_cycles, cycle_data['cycle'].max())

                # 存储保留率曲线
                retention_curves[channel_key] = {
                    'cycles': cycle_data['cycle'].values,
                    'capacity_retention': cycle_data['capacity_retention'].values,
                    'one_c_cycle': one_c_cycle
                }

                # 如果有能量保留率数据，也存储
                if 'energy_retention' in cycle_data.columns:
                    retention_curves[channel_key]['energy_retention'] = cycle_data['energy_retention'].values

                # 如果有电压保留率数据，也存储
                if 'voltage_retention' in cycle_data.columns:
                    retention_curves[channel_key]['voltage_retention'] = cycle_data['voltage_retention'].values

                print(f"成功提取通道 {channel_key} 的保留率曲线，共 {len(cycle_data)} 个数据点")

            except Exception as e:
                print(f"处理文件 {file_path} 时出错: {str(e)}")
                continue

        # 检查是否有足够的通道数据
        if len(retention_curves) < 2:
            print(f"没有足够的通道数据进行比较，只有 {len(retention_curves)} 个通道有容量保留率数据")
            if len(retention_curves) == 1:
                channel_key = list(retention_curves.keys())[0]
                channel_parts = channel_key.split('-')
                return data[(data['主机'] == channel_parts[0]) & (data['通道'] == channel_parts[1])].iloc[0]
            return None

        # 确定共同的循环范围
        min_cycles = []
        for channel_key, curve_data in retention_curves.items():
            # 计算从1C首圈开始的循环数
            cycles_after_one_c = curve_data['cycles'] - curve_data['one_c_cycle'] + 1
            # 只考虑1C首圈及之后的循环
            valid_cycles = cycles_after_one_c[cycles_after_one_c > 0]
            if len(valid_cycles) > 0:
                min_cycles.append(len(valid_cycles))

        if not min_cycles:
            print("没有通道有1C首圈之后的循环数据")
            return None

        # 找到所有通道从1C首圈开始的最小循环数
        common_cycles = min(min_cycles)
        print(f"所有通道从1C首圈开始的最小循环数: {common_cycles}")

        # 如果循环数太少，无法进行有效比较
        if common_cycles < config["min_cycles"]:
            print(f"共同循环数太少({common_cycles} < {config['min_cycles']})，无法进行有效比较")
            return None

        # 创建插值函数，为每个通道生成完整的保留率曲线
        from scipy import interpolate
        interpolated_curves = {
            'capacity': {},  # 容量保留率曲线
            'energy': {},    # 能量保留率曲线
            'voltage': {}    # 电压保留率曲线
        }

        # 定义插值的循环位置范围（从0到common_cycles-1）
        cycle_positions = np.arange(common_cycles)

        for channel_key, curve_data in retention_curves.items():
            # 计算从1C首圈开始的循环位置
            one_c_cycle = curve_data['one_c_cycle']
            cycles = curve_data['cycles']

            # 处理容量保留率
            capacity_retention = curve_data['capacity_retention']

            # 只考虑1C首圈及之后的循环
            valid_indices = cycles >= one_c_cycle
            valid_cycles = cycles[valid_indices]
            valid_capacity_retention = capacity_retention[valid_indices]

            # 计算循环位置（从0开始）
            cycle_positions_actual = valid_cycles - one_c_cycle

            # 只有当通道有足够的数据点时才进行插值
            if len(cycle_positions_actual) >= 2:
                # 创建容量保留率插值函数
                if len(cycle_positions_actual) >= 3 and config["interpolation_method"] == "cubic":
                    # 至少需要3个点才能进行三次样条插值
                    f_capacity = interpolate.interp1d(cycle_positions_actual, valid_capacity_retention, kind='cubic',
                                                    bounds_error=False, fill_value="extrapolate")
                else:
                    # 否则使用线性插值
                    f_capacity = interpolate.interp1d(cycle_positions_actual, valid_capacity_retention, kind='linear',
                                                    bounds_error=False, fill_value="extrapolate")

                # 生成容量保留率插值曲线
                interpolated_capacity = f_capacity(cycle_positions)
                interpolated_curves['capacity'][channel_key] = interpolated_capacity

                # 处理能量保留率（如果有）
                if 'energy_retention' in curve_data and config.get("include_energy", False):
                    energy_retention = curve_data['energy_retention']
                    valid_energy_retention = energy_retention[valid_indices]

                    # 创建能量保留率插值函数
                    if len(cycle_positions_actual) >= 3 and config["interpolation_method"] == "cubic":
                        f_energy = interpolate.interp1d(cycle_positions_actual, valid_energy_retention, kind='cubic',
                                                      bounds_error=False, fill_value="extrapolate")
                    else:
                        f_energy = interpolate.interp1d(cycle_positions_actual, valid_energy_retention, kind='linear',
                                                      bounds_error=False, fill_value="extrapolate")

                    # 生成能量保留率插值曲线
                    interpolated_energy = f_energy(cycle_positions)
                    interpolated_curves['energy'][channel_key] = interpolated_energy

                # 处理电压保留率（如果有）
                if 'voltage_retention' in curve_data and config.get("include_voltage", False):
                    voltage_retention = curve_data['voltage_retention']
                    valid_voltage_retention = voltage_retention[valid_indices]

                    # 创建电压保留率插值函数
                    if len(cycle_positions_actual) >= 3 and config["interpolation_method"] == "cubic":
                        f_voltage = interpolate.interp1d(cycle_positions_actual, valid_voltage_retention, kind='cubic',
                                                       bounds_error=False, fill_value="extrapolate")
                    else:
                        f_voltage = interpolate.interp1d(cycle_positions_actual, valid_voltage_retention, kind='linear',
                                                       bounds_error=False, fill_value="extrapolate")

                    # 生成电压保留率插值曲线
                    interpolated_voltage = f_voltage(cycle_positions)
                    interpolated_curves['voltage'][channel_key] = interpolated_voltage

        # 计算平均曲线和MSE
        mean_curves = {}
        mse_scores = {}
        combined_mse_scores = {}

        # 检查是否有足够的数据进行比较
        if not interpolated_curves['capacity']:
            print("没有足够的容量保留率数据进行比较")
            return None

        # 计算容量保留率平均曲线
        capacity_curves = np.array(list(interpolated_curves['capacity'].values()))
        mean_curves['capacity'] = np.mean(capacity_curves, axis=0)

        # 计算能量保留率平均曲线（如果有）
        if interpolated_curves['energy'] and config.get("include_energy", False):
            energy_curves = np.array(list(interpolated_curves['energy'].values()))
            mean_curves['energy'] = np.mean(energy_curves, axis=0)

        # 计算电压保留率平均曲线（如果有）
        if interpolated_curves['voltage'] and config.get("include_voltage", False):
            voltage_curves = np.array(list(interpolated_curves['voltage'].values()))
            mean_curves['voltage'] = np.mean(voltage_curves, axis=0)

        # 获取权重配置
        capacity_weight = config.get("capacity_weight", 0.6)
        energy_weight = config.get("energy_weight", 0.3)
        voltage_weight = config.get("voltage_weight", 0.1)

        # 确保权重和为1
        total_weight = capacity_weight
        if interpolated_curves['energy'] and config.get("include_energy", False):
            total_weight += energy_weight
        if interpolated_curves['voltage'] and config.get("include_voltage", False):
            total_weight += voltage_weight

        if total_weight != 1.0:
            capacity_weight /= total_weight
            energy_weight /= total_weight
            voltage_weight /= total_weight

        # 计算每个通道的MSE
        for channel_key in interpolated_curves['capacity'].keys():
            # 计算容量保留率MSE
            capacity_curve = interpolated_curves['capacity'][channel_key]

            # 根据配置决定是否使用加权MSE
            if config["use_weighted_mse"]:
                # 计算权重
                weights = self._calculate_cycle_weights(cycle_positions, 0, common_cycles-1, config)
                # 计算加权MSE
                capacity_mse = np.average((capacity_curve - mean_curves['capacity']) ** 2, weights=weights)
            else:
                # 计算普通MSE
                capacity_mse = np.mean((capacity_curve - mean_curves['capacity']) ** 2)

            # 存储容量保留率MSE
            mse_scores[channel_key] = {'capacity': capacity_mse}

            # 初始化综合MSE
            combined_mse = capacity_weight * capacity_mse

            # 计算能量保留率MSE（如果有）
            if channel_key in interpolated_curves['energy'] and config.get("include_energy", False):
                energy_curve = interpolated_curves['energy'][channel_key]

                if config["use_weighted_mse"]:
                    energy_mse = np.average((energy_curve - mean_curves['energy']) ** 2, weights=weights)
                else:
                    energy_mse = np.mean((energy_curve - mean_curves['energy']) ** 2)

                mse_scores[channel_key]['energy'] = energy_mse
                combined_mse += energy_weight * energy_mse

            # 计算电压保留率MSE（如果有）
            if channel_key in interpolated_curves['voltage'] and config.get("include_voltage", False):
                voltage_curve = interpolated_curves['voltage'][channel_key]

                if config["use_weighted_mse"]:
                    voltage_mse = np.average((voltage_curve - mean_curves['voltage']) ** 2, weights=weights)
                else:
                    voltage_mse = np.mean((voltage_curve - mean_curves['voltage']) ** 2)

                mse_scores[channel_key]['voltage'] = voltage_mse
                combined_mse += voltage_weight * voltage_mse

            # 存储综合MSE
            combined_mse_scores[channel_key] = combined_mse

        # 找到综合MSE最小的通道
        if combined_mse_scores:
            best_channel_key = min(combined_mse_scores, key=combined_mse_scores.get)
            print(f"保留率曲线综合MSE最小的通道: {best_channel_key}, MSE={combined_mse_scores[best_channel_key]:.4f}")

            # 输出各指标的MSE
            print(f"  容量保留率MSE: {mse_scores[best_channel_key]['capacity']:.4f}")
            if 'energy' in mse_scores[best_channel_key]:
                print(f"  能量保留率MSE: {mse_scores[best_channel_key]['energy']:.4f}")
            if 'voltage' in mse_scores[best_channel_key]:
                print(f"  电压保留率MSE: {mse_scores[best_channel_key]['voltage']:.4f}")

            # 可视化比较结果
            if self.verbose:
                self._visualize_retention_curves(cycle_positions, interpolated_curves, mean_curves,
                                               best_channel_key, config)

            # 返回对应的通道数据
            print(f"尝试在数据中查找通道: {best_channel_key}")

            try:
                # 使用新的提取函数获取主机和通道
                host, channel = self._extract_host_and_channel(best_channel_key)
                print(f"通道标识拆分为: 主机={host}, 通道={channel}")

                # 打印所有主机和通道信息，便于调试
                print("\n数据中的所有主机-通道对:")
                for i, (idx, row) in enumerate(data.iterrows()):
                    if '主机' in row and '通道' in row:
                        print(f"  #{i+1} 索引={idx}, 主机='{row['主机']}', 通道='{row['通道']}'")

                # 尝试精确匹配
                best_channel_data = data[(data['主机'] == host) & (data['通道'] == channel)]

                # 如果精确匹配未找到结果，记录警告并尝试模糊匹配
                if best_channel_data.empty:
                    print(f"警告: 无法在数据中找到精确匹配的通道: 主机='{host}', 通道='{channel}'")
                    print(f"数据中的主机列唯一值: {data['主机'].unique()}")
                    print(f"数据中的通道列唯一值: {data['通道'].unique()}")

                    # 打印格式信息，帮助检测隐藏字符或格式问题
                    print("\n格式比较:")
                    print(f"要匹配的主机: '{host}', 长度={len(host)}, ASCII码=[{','.join([str(ord(c)) for c in host])}]")
                    for h in data['主机'].unique():
                        print(f"数据中的主机: '{h}', 长度={len(h)}, ASCII码=[{','.join([str(ord(c)) for c in str(h)])}]")

                    print(f"要匹配的通道: '{channel}', 长度={len(channel)}, ASCII码=[{','.join([str(ord(c)) for c in channel])}]")
                    for c in data['通道'].unique():
                        print(f"数据中的通道: '{c}', 长度={len(c)}, ASCII码=[{','.join([str(ord(c)) for c in str(c)])}]")

                    # 尝试模糊匹配
                    print("\n尝试模糊匹配 - 清理格式后比较:")
                    clean_host = str(host).strip()
                    clean_channel = str(channel).strip()

                    for idx, row in data.iterrows():
                        if '主机' in row and '通道' in row:
                            data_host = str(row['主机']).strip()
                            data_channel = str(row['通道']).strip()

                            if data_host == clean_host and data_channel == clean_channel:
                                print(f"找到模糊匹配! 索引={idx}")
                                return row

                    # 打印所有文件信息，帮助调试
                    if '文件名' in data.columns:
                        print("\n所有文件名:")
                        for i, fname in enumerate(data['文件名'].unique()):
                            print(f"  {i+1}. {fname}")
                    if '文件路径' in data.columns:
                        print("\n所有文件路径:")
                        for i, fpath in enumerate(data['文件路径'].unique()):
                            print(f"  {i+1}. {fpath}")

                    print("无法找到匹配的通道，返回None")
                    return None

                if not best_channel_data.empty:
                    print(f"容量保留率曲线比较方法成功选择参考通道: {best_channel_key}")
                    print(f"对应的数据行: 主机={best_channel_data.iloc[0]['主机']}, 通道={best_channel_data.iloc[0]['通道']}")
                    return best_channel_data.iloc[0]
            except Exception as e:
                import traceback
                print(f"通道匹配过程中出错: {str(e)}")
                print(traceback.format_exc())

                # 记录完整错误现场信息
                print(f"\n错误现场信息:")
                print(f"最佳通道标识: {best_channel_key}")
                print(f"数据行数: {len(data)}")
                if len(data) > 0:
                    print(f"数据列: {data.columns.tolist()}")
                    if '主机' in data.columns and '通道' in data.columns:
                        print("主机-通道组合:")
                        for idx, row in data.head(5).iterrows():
                            print(f"  {row['主机']}-{row['通道']}")

                return None
        else:
            print("无法计算MSE，没有足够的数据")

        print("从原始放电容量数据计算容量保留率曲线未找到合适的参考通道")
        print("回退到使用预计算的容量保留率列")
        return self._select_reference_channel_by_retention_columns(data)

    def _extract_unified_batch(self, batch_id):
        """从批次ID中提取统一批次

        按照原始代码的逻辑：
        blank1['统一批次'] = [('-'.join(x.split('-')[:-3])) for x in blank1['批次']] # 去除后3位

        Args:
            batch_id: 批次ID

        Returns:
            统一批次
        """
        if not isinstance(batch_id, str):
            return batch_id

        # 按照原始代码逻辑：去除后3段
        parts = batch_id.split('-')
        if len(parts) > 3:
            unified_batch = '-'.join(parts[:-3])
            return unified_batch
        else:
            # 如果段数不足3段，返回原始批次
            return batch_id

    def _extract_batch_like_original(self, filename):
        """按照原始代码的逻辑提取批次信息

        原始逻辑：
        a3 = '-'.join((j.split(sep='_')[0]).split(sep='-')[5:]) + j.split(sep='_')[1]

        Args:
            filename: 文件名

        Returns:
            批次ID字符串
        """
        try:
            # 按下划线分割
            underscore_parts = filename.split('_')
            if len(underscore_parts) >= 2:
                # 第一部分：下划线前，从第6段（索引5）开始
                first_part_segments = underscore_parts[0].split('-')[5:]
                first_part = '-'.join(first_part_segments)

                # 第二部分：第一个下划线后的部分
                second_part = underscore_parts[1]

                # 组合完整批次
                batch_id = first_part + '-' + second_part

                return batch_id
            else:
                # 如果没有下划线，使用备用方案
                parts = filename.split('-')
                if len(parts) >= 6:
                    return '-'.join(parts[5:])
                else:
                    return filename.split('.')[0]  # 使用文件名（去除扩展名）

        except Exception as e:
            print(f"批次提取异常: {str(e)}")
            return filename.split('.')[0]  # 备用方案

    def _extract_host_and_channel(self, channel_key):
        """从通道标识中提取主机和通道信息

        使用最简单的字符检测方法：
        - 如果第一段包含"." → IP格式 → 前2段是主机，第3-4段是通道
        - 否则 → 标准格式 → 前3段是主机，第4-5段是通道

        Args:
            channel_key: 通道标识

        Returns:
            主机和通道的元组
        """
        if not isinstance(channel_key, str) or not channel_key:
            return "", ""

        parts = channel_key.split('-')

        # 超简单判断：第一部分有"."就是IP格式
        if '.' in parts[0] and len(parts) >= 4:
            # IP格式：192.168.110.236-270060-7-5-...
            host = '-'.join(parts[:2])      # 前2段是主机
            channel = '-'.join(parts[2:4])  # 第3-4段是通道
            return host, channel
        elif len(parts) >= 5:
            # 标准格式：M2-PC2-036-8-1-...
            host = '-'.join(parts[:3])      # 前3段是主机
            channel = '-'.join(parts[3:5])  # 第4-5段是通道
            return host, channel
        else:
            # 兜底方案
            return channel_key, ""

    def _calculate_cycle_weights(self, cycles, min_cycle=None, max_cycle=None, config=None):
        """计算循环权重

        根据配置的权重方法计算每个循环的权重，
        使后期循环的偏差对MSE的贡献更大。

        Args:
            cycles: 循环次数数组
            min_cycle: 最小循环次数（可选，未使用）
            max_cycle: 最大循环次数（可选，未使用）
            config: 配置参数

        Returns:
            权重数组
        """
        # 如果未提供配置，使用默认配置
        if config is None:
            config = CONFIG["REFERENCE_CHANNEL_CONFIG"]["capacity_retention"]

        n_cycles = len(cycles)
        weight_method = config["weight_method"]
        weight_factor = config["weight_factor"]
        late_emphasis = config["late_cycles_emphasis"]

        if weight_method == "constant":
            # 恒定权重
            weights = np.ones(n_cycles)
        elif weight_method == "linear":
            # 线性增长权重: w_j = 1 + weight_factor * j/(n-1)
            weights = 1 + weight_factor * np.linspace(0, 1, n_cycles)
        elif weight_method == "exp":
            # 指数增长权重: w_j = exp(weight_factor * j/(n-1))
            weights = np.exp(weight_factor * np.linspace(0, 1, n_cycles))
        else:
            # 默认使用线性权重
            weights = 1 + weight_factor * np.linspace(0, 1, n_cycles)

        # 对后期循环额外加权
        if late_emphasis > 1.0 and n_cycles > 10:
            # 确定后期循环的起始位置（最后30%的循环）
            late_start = int(n_cycles * 0.7)
            # 对后期循环应用额外权重
            weights[late_start:] *= late_emphasis

        # 归一化权重，使其和为n_cycles
        weights = weights * n_cycles / np.sum(weights)

        return weights

    def _visualize_retention_curves(self, cycles, curves, mean_curves, best_channel_key, config):
        """可视化保留率曲线比较结果

        Args:
            cycles: 循环次数数组
            curves: 各通道的保留率曲线字典（包含容量、能量、电压）
            mean_curves: 平均曲线字典（包含容量、能量、电压）
            best_channel_key: 最佳通道的键
            config: 配置参数
        """
        # 确定需要绘制的指标数量
        metrics = ['capacity']
        if config.get("include_energy", False) and 'energy' in curves and curves['energy']:
            metrics.append('energy')
        if config.get("include_voltage", False) and 'voltage' in curves and curves['voltage']:
            metrics.append('voltage')

        # 确定子图数量和布局
        n_metrics = len(metrics)
        n_rows = n_metrics + 1  # 额外一行用于权重或偏差

        # 创建图表
        fig, axes = plt.subplots(n_rows, 1, figsize=(12, 5 * n_rows),
                                gridspec_kw={'height_ratios': [3] * n_metrics + [1]})

        # 如果只有一个指标，确保axes是列表
        if n_metrics == 1:
            axes = [axes[0], axes[1]]

        # 指标名称映射
        metric_names = {
            'capacity': '容量保留率',
            'energy': '能量保留率',
            'voltage': '电压保留率'
        }

        # 颜色映射
        colors = {
            'capacity': 'blue',
            'energy': 'green',
            'voltage': 'purple'
        }

        # 绘制各指标的保留率曲线
        for i, metric in enumerate(metrics):
            ax = axes[i]

            # 绘制各通道的曲线
            for channel_key, curve in curves[metric].items():
                if channel_key == best_channel_key:
                    ax.plot(cycles, curve, color='red', linewidth=2,
                           label=f"{channel_key} (最佳)")
                else:
                    ax.plot(cycles, curve, color=colors[metric], alpha=0.3, linewidth=1)

            # 绘制平均曲线
            ax.plot(cycles, mean_curves[metric], 'k--', linewidth=2, label="批次平均")

            # 设置标题和标签
            ax.set_title(f"{metric_names[metric]}曲线比较")
            ax.set_xlabel("循环次数")
            ax.set_ylabel(f"{metric_names[metric]} (%)")
            ax.grid(True, linestyle='--', alpha=0.7)
            ax.legend()

        # 在最后一个子图中绘制权重分布或偏差
        ax_last = axes[-1]

        # 根据配置决定是否使用加权MSE
        if config["use_weighted_mse"]:
            # 计算权重
            weights = self._calculate_cycle_weights(cycles, 0, len(cycles)-1, config)
            # 绘制权重分布
            ax_last.bar(cycles, weights, alpha=0.7, color='green')
            ax_last.set_title("MSE计算权重分布")
            ax_last.set_xlabel("循环次数")
            ax_last.set_ylabel("权重")
            ax_last.grid(True, linestyle='--', alpha=0.5)
        else:
            # 如果没有使用权重，显示各通道与平均曲线的偏差
            # 使用第一个指标（通常是容量保留率）
            primary_metric = metrics[0]
            for channel_key, curve in curves[primary_metric].items():
                deviation = np.abs(curve - mean_curves[primary_metric])
                if channel_key == best_channel_key:
                    ax_last.plot(cycles, deviation, 'r-', linewidth=2,
                                label=f"{channel_key} (最佳)")
                else:
                    ax_last.plot(cycles, deviation, 'b-', alpha=0.3, linewidth=1)

            ax_last.set_title(f"各通道与平均{metric_names[primary_metric]}曲线的偏差")
            ax_last.set_xlabel("循环次数")
            ax_last.set_ylabel("偏差 (百分点)")
            ax_last.grid(True, linestyle='--', alpha=0.5)
            ax_last.legend()

        # 调整子图之间的间距
        plt.tight_layout()

        # 保存图表
        save_dir = self.folder_path
        save_path = os.path.join(save_dir, "retention_curves_comparison.png")
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        print(f"已保存保留率曲线比较图表: {save_path}")

        # 不显示图表，直接关闭
        plt.close()

    def _visualize_capacity_retention_curves(self, cycles, curves, mean_curve, best_channel_key,
                                      weighted=False, weights=None):
        """可视化容量保留率曲线比较结果（兼容旧版本）

        Args:
            cycles: 循环次数数组
            curves: 各通道的容量保留率曲线字典
            mean_curve: 平均曲线
            best_channel_key: 最佳通道的键
            weighted: 是否使用加权MSE
            weights: 权重数组
        """
        # 创建一个包含两个子图的图表
        _, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10), gridspec_kw={'height_ratios': [3, 1]})

        # 在第一个子图中绘制容量保留率曲线
        for channel_key, curve in curves.items():
            if channel_key == best_channel_key:
                ax1.plot(cycles, curve, 'r-', linewidth=2, label=f"{channel_key} (最佳)")
            else:
                ax1.plot(cycles, curve, 'b-', alpha=0.3, linewidth=1)

        # 绘制平均曲线
        ax1.plot(cycles, mean_curve, 'k--', linewidth=2, label="批次平均")

        ax1.set_title("容量保留率曲线比较")
        ax1.set_xlabel("循环次数")
        ax1.set_ylabel("容量保留率 (%)")
        ax1.grid(True, linestyle='--', alpha=0.7)
        ax1.legend()

        # 在第二个子图中绘制权重分布（如果使用加权MSE）
        if weighted and weights is not None:
            ax2.bar(cycles, weights, alpha=0.7, color='green')
            ax2.set_title("MSE计算权重分布")
            ax2.set_xlabel("循环次数")
            ax2.set_ylabel("权重")
            ax2.grid(True, linestyle='--', alpha=0.5)
        else:
            # 如果没有使用权重，显示各通道与平均曲线的偏差
            for channel_key, curve in curves.items():
                deviation = np.abs(curve - mean_curve)
                if channel_key == best_channel_key:
                    ax2.plot(cycles, deviation, 'r-', linewidth=2, label=f"{channel_key} (最佳)")
                else:
                    ax2.plot(cycles, deviation, 'b-', alpha=0.3, linewidth=1)

            ax2.set_title("各通道与平均曲线的偏差")
            ax2.set_xlabel("循环次数")
            ax2.set_ylabel("偏差 (百分点)")
            ax2.grid(True, linestyle='--', alpha=0.5)
            ax2.legend()

        # 调整子图之间的间距
        plt.tight_layout()

        # 保存图表
        save_dir = self.folder_path
        save_path = os.path.join(save_dir, "capacity_retention_comparison.png")
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        print(f"已保存容量保留率曲线比较图表: {save_path}")

        # 不显示图表，直接关闭
        plt.close()

    def _visualize_pca_result(self, pca_results, central_idx, sample_ids, batch_name=None, save_path=None):
        """可视化PCA结果

        Args:
            pca_results (np.ndarray): PCA结果
            central_idx (int): 中心样本索引
            sample_ids (array-like): 样本ID
            batch_name (str): 批次名称，用于文件命名
            save_path (str): 保存路径，默认为None时自动生成
        """
        plt.figure(figsize=(10, 8))
        plt.scatter(pca_results[:, 0], pca_results[:, 1], alpha=0.7)
        plt.scatter(pca_results[central_idx, 0], pca_results[central_idx, 1],
                color='red', s=100, marker='*', label='参考通道')

        # 为每个点添加标签
        for i, sample_id in enumerate(sample_ids):
            plt.annotate(str(sample_id), (pca_results[i, 0], pca_results[i, 1]),
                        fontsize=8, alpha=0.7)

        plt.title('电池样本PCA分析')
        plt.xlabel('主成分1')
        plt.ylabel('主成分2')
        plt.grid(True, alpha=0.3)
        plt.legend()
        plt.tight_layout()

        # 如果未指定保存路径，将图片保存到数据可视化文件夹
        if save_path is None:
            # 生成文件名
            if batch_name:
                # 清理批次名称中的特殊字符
                safe_batch_name = "".join(c for c in batch_name if c.isalnum() or c in ('-', '_')).rstrip()
                filename = f'PCA分析_{safe_batch_name}.png'
            else:
                filename = 'PCA分析_未知批次.png'

            # 将图片保存到全局变量中，稍后在export_results中统一保存
            if not hasattr(self, '_pca_plots'):
                self._pca_plots = []

            # 保存图片数据和文件名
            import io
            buf = io.BytesIO()
            plt.savefig(buf, format='png', dpi=300, bbox_inches='tight')
            buf.seek(0)

            self._pca_plots.append({
                'filename': filename,
                'data': buf.getvalue(),
                'batch_name': batch_name or '未知批次'
            })

            print(f"PCA分析图表已准备保存: {filename}")
        else:
            # 如果指定了保存路径，直接保存
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            print(f"已保存PCA分析图表: {save_path}")

        # 不显示图形，只关闭
        plt.close()

    def _multi_feature_reference_selection(self, data, features=['首放', '首效', '首圈电压']):
        """
        使用多特征PCA降维选择参考通道

        Parameters:
        -----------
        data: DataFrame
            候选参考通道数据
        features: list
            用于选择的特征列表

        Returns:
        --------
        Series: 选中的参考通道数据
        """
        if data.empty:
            return None

        if len(data) == 1:
            return data.iloc[0]

        # 确保所有特征都存在
        available_features = [f for f in features if f in data.columns]
        if not available_features:
            # 如果没有有效特征，回退到仅使用首放
            if '首放' in data.columns:
                available_features = ['首放']
            else:
                print("无可用特征进行多特征参考通道选择")
                return data.iloc[0]

        # 确保没有缺失值
        valid_data = data[available_features].dropna()
        if len(valid_data) < 2:
            # 数据不足，回退到原方法
            print("有效数据不足进行PCA，回退到传统方法")
            return data.iloc[0]

        try:
            # 数据规范化
            normalized_data = (valid_data - valid_data.mean()) / valid_data.std()

            # PCA降维
            n_components = min(len(available_features), 2)  # 取最多2个主成分
            pca = PCA(n_components=n_components)
            X_pca = pca.fit_transform(normalized_data)

            # 计算中心点
            centroid = np.mean(X_pca, axis=0)

            # 计算距离
            distances = [np.linalg.norm(x - centroid) for x in X_pca]
            ref_idx = np.argmin(distances)

            # 获取原始索引
            original_idx = valid_data.index[ref_idx]
            ref_row = data.loc[original_idx]
            print(f"多特征选择参考通道: 基于{available_features}的PCA分析")
            return ref_row
        except Exception as e:
            print(f"PCA分析异常: {str(e)}，回退到传统方法")
            # 出错时回退到原始方法
            if '首放' in data.columns:
                subset_discharge_mean = data['首放'].mean()
                diff_series = abs(data['首放'] - subset_discharge_mean)
                best_idx = diff_series.idxmin()
                return data.loc[best_idx]
            else:
                return data.iloc[0]

    def plot_analysis_visualizations(self, save_dir=None):
        """
        生成数据分析可视化图表

        Parameters:
        -----------
        save_dir: str, optional
            图表保存目录，默认为当前目录
        """
        if self.all_cycle_data.empty:
            print("无数据可供可视化")
            return

        if save_dir is None:
            save_dir = os.path.dirname(os.path.abspath(__name__))

        # 确保目录存在
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)

        print(f"正在生成可视化图表，保存到: {save_dir}")

        # 尝试导入seaborn，如果不可用则使用matplotlib
        try:
            import seaborn as sns
            has_seaborn = True
        except ImportError:
            print("警告: seaborn库未安装，将使用基本matplotlib绘图")
            has_seaborn = False

        # 1. 首放容量分布图
        try:
            if '首放' in self.all_cycle_data.columns:
                plt.figure(figsize=(10, 6))
                if has_seaborn:
                    sns.histplot(self.all_cycle_data['首放'].dropna(), bins=50, kde=True)
                else:
                    plt.hist(self.all_cycle_data['首放'].dropna(), bins=50, alpha=0.7)

                plt.title('首放容量分布')
                plt.xlabel('首放容量 (mAh/g)')
                plt.ylabel('频率')

                # 保存图片
                save_path = os.path.join(save_dir, 'discharge_capacity_distribution.png')
                plt.savefig(save_path, dpi=300, bbox_inches='tight')
                print(f"已保存图表: {save_path}")

                # 不显示图形，直接关闭
                plt.close()
        except Exception as e:
            print(f"绘制首放容量分布图异常: {str(e)}")

        # 2. 按1C状态分组的首效箱线图
        try:
            if '首效' in self.all_cycle_data.columns and '1C状态' in self.all_cycle_data.columns:
                valid_data = self.all_cycle_data[self.all_cycle_data['1C状态'].notna()]

                plt.figure(figsize=(10, 6))
                if has_seaborn:
                    sns.boxplot(x='1C状态', y='首效', data=valid_data)
                else:
                    # 使用matplotlib绘制简单箱线图
                    groups = valid_data.groupby('1C状态')['首效']
                    positions = range(len(groups))
                    data = [group for _, group in groups]
                    plt.boxplot(data, positions=positions)
                    plt.xticks(positions, groups.groups.keys())

                plt.title('不同1C状态的首效分布')
                plt.xlabel('1C状态')
                plt.ylabel('首效 (%)')

                # 保存图片
                save_path = os.path.join(save_dir, 'boxplot_首效_by_1C状态.png')
                plt.savefig(save_path, dpi=300, bbox_inches='tight')
                print(f"已保存图表: {save_path}")

                # 不显示图形，直接关闭
                plt.close()
        except Exception as e:
            print(f"绘制箱线图异常: {str(e)}")

        # 3. 多特征散点图 - PCA
        try:
            features = ['首放', '首效', '首圈电压']
            available_features = [f for f in features if f in self.all_cycle_data.columns]

            if len(available_features) >= 2:
                X = self.all_cycle_data[available_features].dropna()

                if len(X) > 5:  # 确保有足够数据点
                    # 数据规范化
                    X_norm = (X - X.mean()) / X.std()

                    # PCA降维
                    pca = PCA(n_components=2)
                    X_pca = pca.fit_transform(X_norm)

                    # 添加1C状态
                    if '1C状态' in self.all_cycle_data.columns:
                        status = self.all_cycle_data.loc[X.index, '1C状态'].fillna('未知')
                    else:
                        status = pd.Series(['数据'] * len(X), index=X.index)

                    # 创建可视化数据框
                    plot_df = pd.DataFrame({
                        'PC1': X_pca[:, 0],
                        'PC2': X_pca[:, 1],
                        '状态': status
                    })

                    plt.figure(figsize=(10, 8))
                    if has_seaborn:
                        sns.scatterplot(x='PC1', y='PC2', hue='状态', data=plot_df, s=80, alpha=0.7)
                    else:
                        # 使用matplotlib绘制散点图
                        for status_name, group in plot_df.groupby('状态'):
                            plt.scatter(group['PC1'], group['PC2'], label=status_name, alpha=0.7)

                    plt.title('电池参数PCA降维可视化')
                    plt.xlabel(f'主成分1 ({pca.explained_variance_ratio_[0]:.2%}方差)')
                    plt.ylabel(f'主成分2 ({pca.explained_variance_ratio_[1]:.2%}方差)')
                    plt.grid(True, linestyle='--', alpha=0.7)
                    plt.legend()

                    # 保存图片
                    save_path = os.path.join(save_dir, 'pca_visualization.png')
                    plt.savefig(save_path, dpi=300, bbox_inches='tight')
                    print(f"已保存图表: {save_path}")

                    # 不显示图形，直接关闭
                    plt.close()
        except Exception as e:
            print(f"绘制PCA散点图异常: {str(e)}")

        # 4. 1C首圈分布图 - 新增
        try:
            # 检查是否有1C数据
            if '1C首圈编号' in self.all_cycle_data.columns and '模式' in self.all_cycle_data.columns:
                # 检查是否有1C模式数据
                has_1c_data = any(mode in CONFIG["MODE_CONFIG"]["one_c_modes"] for mode in self.all_cycle_data['模式'].unique())
                if has_1c_data:
                    # 绘制并保存1C首圈分布图
                    fig = self.plot_1c_distribution()
                    if fig is not None:  # 确认有图可保存
                        save_path = os.path.join(save_dir, '1c_first_cycle_distribution.png')
                        plt.savefig(save_path, dpi=300, bbox_inches='tight')
                        print(f"已保存图表: {save_path}")

                        # 不显示图形，直接关闭
                        plt.close()
                    else:
                        print("无有效1C首圈数据可绘制分布图")
                else:
                    print("数据集中没有1C模式的数据")
        except Exception as e:
            print(f"绘制1C首圈分布图异常: {str(e)}")

        print("所有可视化图表已生成完毕")

    def _calculate_basic_statistics(self, series, batch, group):
        """计算基本统计值 - 使用所有有效样品计算均值"""
        # 获取该批次在原始数据中的总数量（异常检测前）
        original_data = self.all_cycle_data[
            (self.all_cycle_data['系列'] == series) &
            (self.all_cycle_data['统一批次'] == batch)
        ]
        total_count = len(original_data)

        # 当前group是异常检测后的数据，len(group)是异常检测后剩余的数据数量
        filtered_count = len(group)

        # 处理活性物质列 - 添加错误处理
        active_material = '-'
        if '活性物质' in group.columns:
            try:
                # 尝试转换为数值并计算平均值
                numeric_values = pd.to_numeric(group['活性物质'], errors='coerce')
                if not numeric_values.dropna().empty:
                    active_material = round(numeric_values.mean(), 2)
            except Exception as e:
                print(f"计算活性物质平均值时出错: {str(e)}，批次: {batch}")

        return {
            '系列': series,
            '统一批次': batch,
            '上架时间': group['上架时间'].iloc[0] if '上架时间' in group.columns else '-',
            '总数据': total_count,
            '首周有效数据': filtered_count,  # 异常检测后剩余的数据数量
            '首充': round(group['首充'].mean(), 2) if '首充' in group.columns else None,
            '首放': round(group['首放'].mean(), 2) if '首放' in group.columns else None,
            '首效': round(group['首效'].mean(), 2) if '首效' in group.columns else None,
            '首圈电压': round(group['首圈电压'].mean(), 3) if '首圈电压' in group.columns else None,
            '首圈能量': round(group['首圈能量'].mean(), 2) if '首圈能量' in group.columns else None,
            '活性物质': active_material
        }

    def _calculate_cycle_statistics(self, stats, group):
        """计算循环数据统计值

        Args:
            stats: 统计值字典（将被修改）
            group: 分组数据
        """
        print("\n==== 开始计算循环统计值 ====")
        print(f"统一批次: {stats['统一批次']}, 数据行数: {len(group)}")

        try:
            # Cycle2/3/4的平均值计算部分保持不变
            cycle_columns = {
                'Cycle2充电比容量': 'Cycle2充电比容量',
                'Cycle2': 'Cycle2',
                'Cycle3充电比容量': 'Cycle3充电比容量',
                'Cycle3': 'Cycle3',
                'Cycle4充电比容量': 'Cycle4充电比容量',
                'Cycle4': 'Cycle4',
                'Cycle5充电比容量': 'Cycle5充电比容量',
                'Cycle5': 'Cycle5',
                'Cycle6充电比容量': 'Cycle6充电比容量',
                'Cycle6': 'Cycle6',
                'Cycle7充电比容量': 'Cycle7充电比容量',
                'Cycle7': 'Cycle7'
            }

            for key, col in cycle_columns.items():
                if col in group.columns:
                    non_null_values = group[col].dropna()
                    if not non_null_values.empty:
                        stats[key] = round(non_null_values.mean(), 2)

            # 计算1C相关数据 - 使用参考通道的实际值
            # 初始化1C相关统计字段
            stats['1C首周有效数据'] = 0
            stats['1C参考通道'] = None
            stats['1C状态'] = None
            stats['1C首圈编号'] = None
            stats['1C首充'] = None
            stats['1C首放'] = None
            stats['1C首效'] = None
            stats['1C倍率比'] = None
            stats['参考通道当前圈数'] = None  # 确保初始化参考通道当前圈数字段
            stats['当前容量保持'] = None
            stats['电压衰减率mV/周'] = None
            stats['当前电压保持'] = None
            stats['当前能量保持'] = None

            if '模式' in group.columns:
                # 提取所有1C模式的数据
                all_1c_data = group[group['模式'].isin(CONFIG["MODE_CONFIG"]["one_c_modes"])]

                # 详细调试信息 - 统一批次和原始数据
                print(f"\n============= 统一批次分析: {stats['统一批次']} =============")
                print(f"该批次总样品数: {len(group)}, 其中1C模式样品数: {len(all_1c_data)}")

                if not all_1c_data.empty:
                    # 详细列出每个样品的首效和状态
                    print("\n样品详情:")
                    for idx, row in all_1c_data.iterrows():
                        print(f"  样品 {row['主机']}-{row['通道']}: 首效={row['1C首效']}%, 状态={row['1C状态']}")

                # 检查是否有缺失的通道
                if self.verbose:
                    print("\n检查是否有缺失的通道:")
                    # 获取该批次在原始数据中的所有通道
                    batch_mask = (self.all_cycle_data['统一批次'] == stats['统一批次'])
                    all_channels_in_batch = self.all_cycle_data[batch_mask][['主机', '通道']].drop_duplicates()
                    print(f"原始数据中该批次的所有通道:")
                    for i, (_, row) in enumerate(all_channels_in_batch.iterrows()):
                        print(f"  {i+1}. 主机={row['主机']}, 通道={row['通道']}")

                    # 获取当前分组中的所有通道
                    current_channels = group[['主机', '通道']].drop_duplicates()
                    print(f"当前分组中的所有通道:")
                    for i, (_, row) in enumerate(current_channels.iterrows()):
                        print(f"  {i+1}. 主机={row['主机']}, 通道={row['通道']}")

                    # 找出缺失的通道
                    all_channels_set = set(tuple(row) for _, row in all_channels_in_batch.iterrows())
                    current_channels_set = set(tuple(row) for _, row in current_channels.iterrows())
                    missing_channels = all_channels_set - current_channels_set

                    if missing_channels:
                        print(f"缺失的通道数: {len(missing_channels)}")
                        print("缺失的通道:")
                        for host, channel in missing_channels:
                            # 查找原始数据中的行
                            row = self.all_cycle_data[(self.all_cycle_data['主机'] == host) & (self.all_cycle_data['通道'] == channel)].iloc[0]
                            print(f"  主机={host}, 通道={channel}, 首放={row['首放']}, 首效={row['首效']}, 批次={row['批次']}")
                    else:
                        print("没有缺失的通道")

                # 1. 准确分离不同状态的样品子集
                # 正常样品：首效 ≥ 85%
                normal_samples = all_1c_data[
                    (all_1c_data['1C状态'] == '正常')
                ]

                # 首效低样品：80% ≤ 首效 < 85%
                low_eff_samples = all_1c_data[
                    (all_1c_data['1C状态'] == '首效低')
                ]

                # 首效过低样品：首效 < 80%
                very_low_eff_samples = all_1c_data[
                    (all_1c_data['1C状态'] == '首效过低')
                ]

                # 过充样品
                overcharge_samples = all_1c_data[
                    (all_1c_data['1C状态'] == '1C过充')
                ]

                # 打印状态分布信息，帮助调试
                print(f"\n状态分布:")
                print(f"  正常样品: {len(normal_samples)}个")
                print(f"  首效低样品: {len(low_eff_samples)}个")
                print(f"  首效过低样品: {len(very_low_eff_samples)}个")
                print(f"  过充样品: {len(overcharge_samples)}个")

                # 2. 确定参考通道和批次状态 - 修改后的优先级逻辑
                reference_data = pd.DataFrame()
                batch_1c_status = None

                if not normal_samples.empty:  # 优先使用正常样品
                    reference_data = normal_samples
                    batch_1c_status = "正常"
                    print(f"\n【决策】: 选择正常样品作为参考 ({len(normal_samples)}个)")

                elif not low_eff_samples.empty:  # 修改: 只要有首效低样品就选择，无需>50%的条件
                    reference_data = low_eff_samples
                    batch_1c_status = "首效低"
                    print(f"\n【决策】: 选择首效低样品作为参考 ({len(low_eff_samples)}个)")

                elif not very_low_eff_samples.empty:  # 只有首效过低样品
                    batch_1c_status = "首效过低"
                    print(f"\n【决策】: 仅有首效过低样品，无参考通道")

                # 3. 更新批次1C状态
                stats['1C状态'] = batch_1c_status
                print(f"\n批次状态设为: {batch_1c_status}")

                # 4. 如果有可用参考数据，选择参考通道并记录其数据
                if not reference_data.empty:
                    stats['1C首周有效数据'] = len(reference_data)
                    print(f"\n可选参考样品数: {len(reference_data)}")

                    # 选择参考通道
                    print(f"\n开始参考通道选择...")

                    # 记录所有通道信息，便于出错时调试
                    try:
                        print("\n==== 详细通道信息 ====")
                        for idx, channel_row in reference_data.iterrows():
                            print(f"通道索引 {idx}:")
                            print(f"  主机: {channel_row['主机']}")
                            print(f"  通道: {channel_row['通道']}")
                            print(f"  批次: {channel_row['批次']}")
                            if '文件名' in channel_row:
                                print(f"  文件名: {channel_row['文件名']}")
                            elif '文件路径' in channel_row:
                                print(f"  文件路径: {channel_row['文件路径']}")
                    except Exception as debug_error:
                        print(f"打印通道信息时出错: {str(debug_error)}")

                    print("==== 通道信息结束 ====\n")

                    try:
                        # 根据配置的方法优先级选择参考通道
                        ref_channel = None
                        method_priority = CONFIG["REFERENCE_CHANNEL_CONFIG"]["method_priority"]
                        method_results = {}  # 存储各方法的结果，用于比较

                        # 尝试所有配置的方法
                        for method in method_priority:
                            try:
                                if method == "capacity_retention" and CONFIG["REFERENCE_CHANNEL_CONFIG"]["capacity_retention"]["enabled"]:
                                    print("尝试使用容量保留率曲线比较方法...")
                                    method_results["capacity_retention"] = self._select_reference_channel_by_capacity_retention(reference_data)
                                    if method_results["capacity_retention"] is not None:
                                        print("容量保留率曲线比较方法成功选择参考通道")

                                elif method == "pca":
                                    print("尝试使用PCA多特征选择方法...")
                                    method_results["pca"] = self._select_reference_channel_by_pca(reference_data)
                                    if method_results["pca"] is not None:
                                        print("PCA多特征方法成功选择参考通道")

                                elif method == "traditional":
                                    print("尝试使用传统方法选择...")
                                    method_results["traditional"] = self._select_reference_channel_from_subset(reference_data, use_multi_feature=False)
                                    if method_results["traditional"] is not None:
                                        print("传统方法成功选择参考通道")
                            except Exception as method_error:
                                print(f"{method}方法选择参考通道失败: {str(method_error)}")
                                continue

                        # 按优先级选择参考通道
                        for method in method_priority:
                            if method in method_results and method_results[method] is not None:
                                ref_channel = method_results[method]
                                print(f"根据优先级选择了{method}方法的结果作为参考通道")
                                break

                        # 如果所有方法都失败，使用第一个通道作为参考
                        if ref_channel is None and not reference_data.empty:
                            print("所有方法都未能选择参考通道，使用第一个通道作为参考")
                            ref_channel = reference_data.iloc[0]

                        # 如果有多个方法成功，输出比较信息
                        successful_methods = [m for m in method_results if method_results[m] is not None]
                        if len(successful_methods) > 1:
                            print("\n多个方法成功选择了参考通道，比较结果:")
                            for method in successful_methods:
                                channel = method_results[method]
                                print(f"  - {method}: {channel['主机']}-{channel['通道']}")
                            # 找出实际使用的方法
                            used_method = None
                            for method in method_priority:
                                if method in method_results and method_results[method] is not None:
                                    if method_results[method] is ref_channel:
                                        used_method = method
                                        break

                            print(f"最终选择: {ref_channel['主机']}-{ref_channel['通道']} (使用{used_method or method_priority[0]}方法)")

                            # 如果容量保留率曲线比较方法和其他方法选择了不同的通道，输出警告
                            if "capacity_retention" in method_results and method_results["capacity_retention"] is not None:
                                capacity_channel = method_results["capacity_retention"]
                                if ref_channel is not capacity_channel:
                                    print(f"警告: 容量保留率曲线比较方法选择的通道与最终选择的通道不同")
                                    print(f"  - 容量保留率方法: {capacity_channel['主机']}-{capacity_channel['通道']}")
                                    print(f"  - 最终选择: {ref_channel['主机']}-{ref_channel['通道']}")
                                    print("建议检查配置的方法优先级，或考虑手动指定参考通道")
                    except Exception as selection_error:
                        print(f"选择参考通道过程中出错: {str(selection_error)}")
                        import traceback
                        print(traceback.format_exc())

                        # 尝试直接使用第一个通道作为参考
                        try:
                            print("尝试使用第一个可用通道作为参考...")
                            if not reference_data.empty:
                                ref_channel = reference_data.iloc[0]
                                print(f"选择了第一个通道作为参考: {ref_channel['主机']}-{ref_channel['通道']}")
                            else:
                                ref_channel = None
                                print("没有可用的参考通道")
                        except Exception as fallback_error:
                            print(f"使用第一个通道作为参考失败: {str(fallback_error)}")
                            ref_channel = None

                    if ref_channel is not None:
                        try:
                            # 设置参考通道
                            stats['1C参考通道'] = f"{ref_channel['主机']}-{ref_channel['通道']}"
                            print(f"\n【结果】: 成功选择参考通道: {stats['1C参考通道']}")

                            # 使用参考通道的实际值
                            print("\n参考通道数据:")
                            # 1C相关数据 - 直接使用参考通道的实际值
                            for key in ['1C首圈编号', '1C首充', '1C首放', '1C首效', '1C倍率比']:
                                if key in ref_channel and pd.notna(ref_channel[key]):
                                    stats[key] = ref_channel[key]
                                    print(f"  {key}: {stats[key]}")

                            # 容量保持率相关数据 - 直接使用参考通道的实际值
                            for field in ['当前圈数', '当前容量保持', '电压衰减率mV/周', '当前电压保持', '当前能量保持']:
                                if field in ref_channel and pd.notna(ref_channel[field]):
                                    if field == '当前圈数':
                                        stats['参考通道当前圈数'] = ref_channel[field]  # 特别处理参考通道当前圈数
                                        print(f"  参考通道当前圈数: {stats['参考通道当前圈数']}")
                                    else:
                                        stats[field] = ref_channel[field]
                                        print(f"  {field}: {stats[field]}")
                        except Exception as data_error:
                            print(f"设置参考通道数据时出错: {str(data_error)}")
                    else:
                        print("\n【结果】: 参考通道选择失败，无参考通道")
                else:
                    print("\n无可用参考样品")

                print("============= 批次分析结束 =============\n")

        except Exception as e:
            import traceback
            print(f"计算统计值时发生错误: {str(e)}")
            print(traceback.format_exc())
            print("尝试继续处理其他数据...")

        return stats

    def _select_reference_channel_from_subset(self, data_subset, use_multi_feature=True):
        """从数据子集中选择参考通道

        Args:
            data_subset: 用于选择参考通道的数据子集 (DataFrame)
            use_multi_feature: 是否使用多特征PCA方法选择参考通道

        Returns:
            参考通道所在行的数据 (pd.Series) 或 None
        """
        if data_subset.empty or '首放' not in data_subset.columns:
            return None

        if len(data_subset) == 1:
            return data_subset.iloc[0]

        try:
            # 使用多特征选择
            if use_multi_feature:
                try:
                    # 尝试使用多特征方法
                    features = ['首放', '首效', '首圈电压']
                    available_features = [f for f in features if f in data_subset.columns]

                    if len(available_features) >= 2:  # 至少要有2个特征才能用PCA
                        return self._multi_feature_reference_selection(data_subset, features=available_features)
                except Exception as e:
                    print(f"多特征参考通道选择失败: {str(e)}, 回退到传统方法")

            # 回退到传统方法: 基于首放选择
            subset_discharge_mean = data_subset['首放'].mean()
            print(f"使用传统方法, 首放平均值: {subset_discharge_mean:.2f}")

            # 计算每个通道与子集平均值的差异
            try:
                indexed_subset = data_subset.set_index(['主机', '通道'], drop=False)
                diff_series = abs(indexed_subset['首放'] - subset_discharge_mean)

                # 找到差异最小的通道的复合索引
                min_diff_idx_tuple = diff_series.idxmin()
                print(f"找到差异最小的通道: 主机={min_diff_idx_tuple[0]}, 通道={min_diff_idx_tuple[1]}")

                # 从原始data_subset中获取对应行
                best_channel_row = data_subset[
                    (data_subset['主机'] == min_diff_idx_tuple[0]) &
                    (data_subset['通道'] == min_diff_idx_tuple[1])
                ]

                if not best_channel_row.empty:
                    print(f"传统方法选择的参考通道: {min_diff_idx_tuple[0]}-{min_diff_idx_tuple[1]}")
                    return best_channel_row.iloc[0]
                else:
                    print(f"警告: 无法在数据中找到通道: 主机={min_diff_idx_tuple[0]}, 通道={min_diff_idx_tuple[1]}")
                    # 失败时使用首行
                    print(f"使用第一行作为参考通道")
                    return data_subset.iloc[0]
            except Exception as index_error:
                print(f"使用索引查找参考通道时出错: {str(index_error)}")
                print(f"尝试直接使用首放最接近平均值的行")

                # 计算直接使用首放列的差异
                diff_values = abs(data_subset['首放'] - subset_discharge_mean)
                closest_idx = diff_values.idxmin()
                return data_subset.loc[closest_idx]

        except Exception as outer_error:
            print(f"选择参考通道过程中出现严重错误: {str(outer_error)}")
            print(f"回退到使用第一个通道")

            if not data_subset.empty:
                return data_subset.iloc[0]
            return None

    # ===== 数据导出方法 =====
    def export_results(self, start_time, generate_visuals=False, output_dir=None):
        """导出结果到Excel文件

        Args:
            start_time: 起始时间标记
            generate_visuals: 是否生成高级可视化图表
            output_dir: 输出目录，如果不指定则使用当前目录
        """
        # 检查是否有数据可以导出
        has_data = (not self.all_cycle_data.empty or
                   not self.all_first_cycle.empty or
                   not self.all_error_data.empty or
                   not self.inconsistent_data.empty or
                   not self.statistics_data.empty)

        if not has_data:
            print("\n警告: 没有任何数据可以导出!")
            print("请检查指定的文件夹中是否包含符合格式要求的Excel文件")
            return

        try:
            # 排序所有数据框
            self._sort_dataframes()

            # 确定输出目录
            if output_dir is None:
                output_dir = os.getcwd()

            print(f"导出结果到目录: {output_dir}")

            # 创建Excel工作簿
            output_filename = f"LIMS数据汇总表-{start_time}-{time.strftime('%m%d', time.localtime())}_1.xlsx"
            output_path = os.path.join(output_dir, output_filename)

            print(f"\n正在创建汇总表: {output_path}")

            # 创建Excel写入器
            writer = pd.ExcelWriter(output_path)

            # 写入各个工作表
            sheets_written = 0

            # 写入第1周循环数据
            if not self.all_first_cycle.empty:
                self._write_to_excel(writer, self.all_first_cycle, "第1周循环ing")
                print(f"已写入'第1周循环ing'表格，共{len(self.all_first_cycle)}行数据")
                sheets_written += 1
            else:
                # 创建空的DataFrame并写入
                empty_df = pd.DataFrame(columns=CONFIG["EXCEL_COLS"].get('first_cycle', ['主机', '通道', '批次', '首放', '首效']))
                self._write_to_excel(writer, empty_df, "第1周循环ing")
                print("已写入空的'第1周循环ing'表格")
                sheets_written += 1

            # 写入循环2圈以上数据
            if not self.all_cycle_data.empty:
                self._write_to_excel(writer, self.all_cycle_data, "循环2圈以上")
                print(f"已写入'循环2圈以上'表格，共{len(self.all_cycle_data)}行数据")
                sheets_written += 1
            else:
                # 创建空的DataFrame并写入，使用'cycle'键，确保与'main'键内容一致
                empty_df = pd.DataFrame(columns=CONFIG["EXCEL_COLS"].get('cycle', ['主机', '通道', '批次', '首放', '首效', '当前容量保持']))
                self._write_to_excel(writer, empty_df, "循环2圈以上")
                print("已写入空的'循环2圈以上'表格")
                sheets_written += 1

            # 写入首周充放异常数据
            if not self.all_error_data.empty:
                self._write_to_excel(writer, self.all_error_data, "首周充放异常")
                print(f"已写入'首周充放异常'表格，共{len(self.all_error_data)}行数据")
                sheets_written += 1
            else:
                # 创建空的DataFrame并写入
                empty_df = pd.DataFrame(columns=CONFIG["EXCEL_COLS"].get('error', ['主机', '通道', '批次', '异常类型', '异常值']))
                self._write_to_excel(writer, empty_df, "首周充放异常")
                print("已写入空的'首周充放异常'表格")
                sheets_written += 1

            # 写入首放一致性差待复测数据
            if not self.inconsistent_data.empty:
                self._write_to_excel(writer, self.inconsistent_data, '首放一致性差待复测')
                print(f"已写入'首放一致性差待复测'表格，共{len(self.inconsistent_data)}行数据")
                sheets_written += 1
            else:
                # 创建空的DataFrame并写入，使用'inconsistent'键
                empty_df = pd.DataFrame(columns=CONFIG["EXCEL_COLS"].get('inconsistent', ['主机', '通道', '批次', '首放差异', '首效差异']))
                self._write_to_excel(writer, empty_df, '首放一致性差待复测')
                print("已写入空的'首放一致性差待复测'表格")
                sheets_written += 1

            # 写入统计数据
            if not self.statistics_data.empty:
                # 使用预定义列顺序输出统计数据
                stat_cols = [col for col in CONFIG["EXCEL_COLS"]['statistics'] if col in self.statistics_data.columns]
                self._write_to_excel(writer, self.statistics_data[stat_cols], '统计数据(剔除异常点)')
                print(f"已写入'统计数据(剔除异常点)'表格，共{len(self.statistics_data)}行数据")
                sheets_written += 1
            else:
                # 创建空的DataFrame并写入
                empty_df = pd.DataFrame(columns=CONFIG["EXCEL_COLS"].get('statistics', []))
                self._write_to_excel(writer, empty_df, '统计数据(剔除异常点)')
                print("已写入空的'统计数据(剔除异常点)'表格")
                sheets_written += 1

            # 检查是否有任何工作表被写入
            if sheets_written == 0:
                print("警告: 没有任何数据被写入Excel文件!")
                return

            # 保存文件
            writer._save()
            print(f"[Congrats！]所有数据已汇总到文件: {output_path}")
            print(f"您可以手动打开该文件查看结果")

            # 关闭日志系统
            if hasattr(self, 'logger'):
                self.logger.close()

        except Exception as e:
            print(f"导出结果时出错: {str(e)}")
            import traceback
            print(traceback.format_exc())

        # 在保存文件完成后添加
        if generate_visuals:
            print("\n开始生成高级数据分析可视化图表...")
            # 在输出目录中创建可视化子文件夹
            vis_dir = os.path.join(output_dir, f"数据可视化-{start_time}")

            # 确保可视化目录存在
            if not os.path.exists(vis_dir):
                os.makedirs(vis_dir)

            try:
                # 首先保存PCA图片（如果有的话）
                if hasattr(self, '_pca_plots') and self._pca_plots:
                    print(f"保存 {len(self._pca_plots)} 个PCA分析图片...")
                    for pca_plot in self._pca_plots:
                        pca_path = os.path.join(vis_dir, pca_plot['filename'])
                        with open(pca_path, 'wb') as f:
                            f.write(pca_plot['data'])
                        print(f"已保存PCA图片: {pca_plot['filename']} (批次: {pca_plot['batch_name']})")

                # 然后生成其他可视化图表
                self.plot_analysis_visualizations(save_dir=vis_dir)
                print(f"可视化图表已保存至: {vis_dir}")
                print(f"您可以手动打开该文件夹查看可视化结果")
            except Exception as e:
                print(f"生成可视化图表时出错: {str(e)}")
                print("请检查是否安装了必要的库，如matplotlib和seaborn")

    def _sort_dataframes(self):
        """排序所有数据框"""
        if not self.all_cycle_data.empty:
            self.all_cycle_data = self.all_cycle_data.sort_values(['系列', '批次'], ascending=True)
            self.all_cycle_data.index = range(1, len(self.all_cycle_data) + 1)

        if not self.all_first_cycle.empty:
            self.all_first_cycle = self.all_first_cycle.sort_values(['系列', '批次'], ascending=True)
            self.all_first_cycle.index = range(1, len(self.all_first_cycle) + 1)

        if not self.all_error_data.empty:
            self.all_error_data = self.all_error_data.sort_values(['系列', '批次'], ascending=True)
            self.all_error_data.index = range(1, len(self.all_error_data) + 1)

        if not self.inconsistent_data.empty:
            self.inconsistent_data = self.inconsistent_data.sort_values(['系列', '上架时间'], ascending=True)
            self.inconsistent_data.index = range(1, len(self.inconsistent_data) + 1)

    def _write_to_excel(self, writer, df, sheet_name):
        """将DataFrame写入Excel

        Args:
            writer: Excel写入器
            df: 待写入的DataFrame
            sheet_name: 工作表名称
        """
        # 即使DataFrame为空也写入
        if df.empty:
            # 确保DataFrame至少有列名
            if len(df.columns) == 0:
                # 如果没有列名，使用默认列名
                if sheet_name == "第1周循环ing":
                    df = pd.DataFrame(columns=CONFIG["EXCEL_COLS"].get('first_cycle', ['主机', '通道', '批次', '首放', '首效']))
            elif sheet_name == "循环2圈以上":
                # 使用'cycle'键，不再使用'main'键
                df = pd.DataFrame(columns=CONFIG["EXCEL_COLS"].get('cycle', ['主机', '通道', '批次', '首放', '首效', '当前容量保持']))
            elif sheet_name == "首周充放异常":
                df = pd.DataFrame(columns=CONFIG["EXCEL_COLS"].get('error', ['主机', '通道', '批次', '异常类型', '异常值']))
            elif sheet_name == "首放一致性差待复测":
                # 使用'inconsistent'键
                df = pd.DataFrame(columns=CONFIG["EXCEL_COLS"].get('inconsistent', ['主机', '通道', '批次', '首放差异', '首效差异']))
            elif sheet_name == "统计数据(剔除异常点)":
                df = pd.DataFrame(columns=CONFIG["EXCEL_COLS"].get('statistics', ['系列', '统一批次', '样品数', '首放平均值', '首效平均值']))
            else:
                df = pd.DataFrame(columns=['主机', '通道', '批次'])

            # 写入空的DataFrame
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            return

        # 创建副本以避免修改原始DataFrame
        df_to_write = df.copy()

        # 将特定列的None值转换为空字符串，以在Excel中显示为空白
        columns_to_clean_for_none = ['1C参考通道', '1C状态']
        for col in columns_to_clean_for_none:
            if col in df_to_write.columns:
                df_to_write[col] = df_to_write[col].fillna('')

        # 写入Excel，无需设置列宽
        df_to_write.to_excel(writer, sheet_name=sheet_name, index=False)

    # ===== 可视化方法 =====
    def plot_boxplot(self, x, y, df, y_min, y_max, title=None):
        """绘制箱线图

        Args:
            x: x轴数据列名
            y: y轴数据列名
            df: 数据框
            y_min: y轴最小值
            y_max: y轴最大值
            title: 图表标题

        Returns:
            matplotlib图形对象
        """
        plt.figure(figsize=(10, 6))
        ax = sns.boxplot(x=x, y=y, data=df, fliersize=12, whis=1, fill=False)
        sns.pointplot(x=x, y=y, data=df, errorbar=None, color='r', linestyles='--')
        plt.scatter(data=df, x=x, y=y, color='m', marker='o')

        plt.xticks(rotation=30)
        ax.set(ylim=(y_min, y_max))

        if title:
            plt.title(title)

        plt.tight_layout()
        return plt.gcf()

    def plot_1c_distribution(self, save_path=None):
        """绘制1C首圈分布图

        Args:
            save_path: 保存路径，可选

        Returns:
            matplotlib图形对象
        """
        if self.all_cycle_data.empty or '1C首圈编号' not in self.all_cycle_data.columns:
            print("没有1C首圈数据可供分析")
            return None

        # 只分析有效的1C数据
        valid_data = self.all_cycle_data[
            self.all_cycle_data['模式'].isin(CONFIG["MODE_CONFIG"]["one_c_modes"]) &
            self.all_cycle_data['1C首圈编号'].notna()
        ]

        if valid_data.empty:
            print("没有有效的1C首圈数据")
            return None

        plt.figure(figsize=(10, 6))

        # 尝试导入seaborn，如果不可用则使用matplotlib
        try:
            import seaborn as sns
            # 绘制1C首圈分布
            ax = sns.countplot(x='1C首圈编号', data=valid_data, palette='viridis')

            # 添加数值标签
            for p in ax.patches:
                ax.annotate(f'{int(p.get_height())}',
                        (p.get_x() + p.get_width() / 2., p.get_height()),
                        ha = 'center', va = 'bottom', fontsize=10)
        except ImportError:
            print("警告: seaborn库未安装，使用基本matplotlib绘图")
            # 使用matplotlib绘制条形图
            counts = valid_data['1C首圈编号'].value_counts().sort_index()
            ax = plt.bar(counts.index, counts.values)

            # 添加数值标签
            for i, v in enumerate(counts.values):
                plt.text(counts.index[i], v, str(v), ha='center', va='bottom')

        plt.title('1C首圈分布')
        plt.xlabel('1C首圈编号')
        plt.ylabel('样品数量')
        plt.tight_layout()

        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')

        return plt.gcf()

    def _try_match_channel(self, host, channel, data):
        """尝试多种方式匹配通道

        Args:
            host: 主机标识
            channel: 通道标识
            data: 数据DataFrame

        Returns:
            匹配的行或None
        """
        print(f"尝试匹配通道: 主机='{host}', 通道='{channel}'")

        # 1. 精确匹配
        match = data[(data['主机'] == host) & (data['通道'] == channel)]
        if not match.empty:
            print(f"精确匹配成功!")
            return match.iloc[0]

        # 2. 清理空格后匹配
        clean_host = str(host).strip()
        clean_channel = str(channel).strip()
        match = data[(data['主机'].str.strip() == clean_host) &
                    (data['通道'].str.strip() == clean_channel)]
        if not match.empty:
            print(f"清理空格后匹配成功!")
            return match.iloc[0]

        # 3. 尝试转换格式匹配 - 比如通道中的中文破折号替换为英文破折号
        alt_host = str(host).replace('－', '-').replace('—', '-')
        alt_channel = str(channel).replace('－', '-').replace('—', '-')
        match = data[(data['主机'] == alt_host) & (data['通道'] == alt_channel)]
        if not match.empty:
            print(f"替换特殊字符后匹配成功!")
            return match.iloc[0]

        # 4. 打印调试信息，帮助诊断为什么匹配失败
        print(f"所有匹配尝试失败。")
        print(f"要匹配的主机: '{host}', 通道: '{channel}'")
        print(f"数据中的主机值: {data['主机'].unique()}")
        print(f"数据中的通道值: {data['通道'].unique()}")

        return None

def identify_series_from_filename(file_name):
    """从文件名中提取系列标识

    Args:
        file_name: 文件名

    Returns:
        系列标识或None
    """
    parts = file_name.split('-')
    if len(parts) >= 6:
        # 提取第六位作为系列标识
        series_part = parts[5].strip()
        # 提取字母部分作为系列ID
        series_id = re.sub(r'[^a-zA-Z]', '', series_part)
        if series_id:
            # 如果提取出了字母，使用第一个字母作为系列标识
            return series_id[0].upper()
        else:
            # 如果没有字母，尝试提取数字部分
            numbers = re.sub(r'[^0-9]', '', series_part)
            if numbers:
                return f"N{numbers}"
    return None

def auto_detect_series(folder_path):
    """自动检测文件夹中所有的系列标识，并按系列分组文件

    Args:
        folder_path: 数据文件夹路径

    Returns:
        包含系列标识和对应文件列表的字典
    """
    print(f"开始自动检测系列标识，扫描文件夹: {folder_path}")

    # 获取所有xlsx文件，排除汇总表(_1.xlsx结尾的文件)
    xlsx_files = [x for x in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, x)) and
                 x.endswith('.xlsx') and not x.endswith('_1.xlsx')]

    print(f"找到 {len(xlsx_files)} 个Excel文件，开始提取系列标识...")

    # 提取所有可能的系列标识
    series_identifiers = {}
    default_series_found = False

    # 首先使用CONFIG中已定义的系列
    for series_name, rules in CONFIG.get("SERIES_CONFIG", {}).get("series", {}).items():
        series_identifiers[series_name] = {
            'config': rules,
            'files': []
        }

    # 然后扫描文件名提取新的系列
    for file_name in xlsx_files:
        file_path = os.path.join(folder_path, file_name)
        # 首先检查文件是否匹配已知系列
        file_matched = False
        for series_name, series_data in series_identifiers.items():
            rules = series_data['config']
            # 检查包含规则
            include_match = True
            if 'include' in rules:
                include_match = any(pattern in file_name for pattern in rules['include'])

            # 检查排除规则
            exclude_match = False
            if 'exclude' in rules:
                exclude_match = any(pattern in file_name for pattern in rules['exclude'])

            # 如果满足包含规则且不满足排除规则，则文件已匹配
            if include_match and not exclude_match:
                series_identifiers[series_name]['files'].append(file_path)
                file_matched = True
                break

        # 如果文件不匹配任何已知系列，尝试自动识别
        if not file_matched:
            series_id = identify_series_from_filename(file_name)
            if series_id:
                if series_id not in series_identifiers:
                    # 创建新的系列规则
                    pattern = f"-{series_id}-"
                    series_identifiers[series_id] = {
                        'config': {
                            'include': [pattern]
                        },
                        'files': [file_path]
                    }
                    print(f"发现新的系列标识: {series_id}, 匹配模式: {pattern}")
                else:
                    # 添加到已有系列
                    series_identifiers[series_id]['files'].append(file_path)
            else:
                # 如果无法识别系列，标记使用默认系列
                default_series_found = True

    # 如果有无法识别系列的文件，确保默认系列存在
    if default_series_found:
        default_series = CONFIG.get("SERIES_CONFIG", {}).get("default_series", "默认系列")
        if default_series not in series_identifiers:
            series_identifiers[default_series] = {
                'config': {},
                'files': []
            }
            print(f"使用默认系列: {default_series} 用于无法识别系列的文件")

    # 将每个系列的文件数量打印出来
    print(f"完成系列识别，共找到 {len(series_identifiers)} 个系列")
    for series_name, series_data in series_identifiers.items():
        file_count = len(series_data['files'])
        print(f"  - {series_name}: {file_count} 个文件")

    # 只返回系列和对应的文件列表，简化后续处理
    result = {}
    for series_name, series_data in series_identifiers.items():
        result[series_name] = series_data['files']

    return result


# ===== 主程序入口 =====
def main(folder_path):
    """主程序入口

    Args:
        folder_path: 数据文件夹路径
    """
    print(f"使用文件夹路径: {folder_path}")

    # 基本路径验证
    if not os.path.exists(folder_path):
        print(f"\n错误: 指定的文件夹不存在: {folder_path}")
        return

    # 检查文件夹中是否有Excel文件
    excel_files = [f for f in os.listdir(folder_path)
                  if os.path.isfile(os.path.join(folder_path, f)) and
                  os.path.splitext(f)[1].lower() == '.xlsx']

    if not excel_files:
        print(f"\n警告: 指定的文件夹中没有Excel文件: {folder_path}")
        print("将只生成空的汇总表")
    else:
        print(f"文件夹中共有 {len(excel_files)} 个Excel文件")

    # 记录开始时间
    start_time = time.time()
    start_time_str = time.strftime('%H%M%S', time.localtime())

    # 实例化数据处理器，传递文件夹路径
    processor = BatteryDataProcessor(folder_path=folder_path)

    # 记录主函数开始到调试日志
    processor.logger.log_debug("主函数开始执行")
    processor.logger.log_debug(f"开始时间: {start_time_str}")

    # 自动检测系列组，如果有多个系列，则提取数字作为批次号
    file_groups = auto_detect_series(folder_path)
    processor.logger.log_debug(f"检测到文件组: {list(file_groups.keys())}")

    # 处理所有文件
    processor.process_all_files(file_groups)

    # 处理异常文件
    processor.process_error_files()

    # 处理只有一个循环的文件
    processor.process_first_cycle_files()

    # 计算处理的统计数据
    processor.calculate_statistics()

    # 导出结果，并生成可视化
    processor.export_results(start_time=start_time_str, generate_visuals=True, output_dir=folder_path)

    # 输出执行时间
    end_time = time.time()
    execution_time = end_time - start_time
    print(f"\n总执行时间: {execution_time:.2f} 秒")

    # 关闭日志系统
    processor.logger.close()

    return processor

def filter_files(directory='.', pattern='*.xlsx', exclude_pattern='_1.xlsx', additional_filters=None):
    """过滤文件列表

    Args:
        directory: 文件目录
        pattern: 文件模式
        exclude_pattern: 排除的文件模式
        additional_filters: 额外的过滤条件字典

    Returns:
        过滤后的文件列表
    """
    import glob

    # 获取所有匹配pattern的文件
    all_files = glob.glob(os.path.join(directory, pattern))

    # 排除指定模式的文件
    filtered_files = [f for f in all_files if exclude_pattern not in f]

    # 应用额外的过滤条件
    if additional_filters:
        for key, value in additional_filters.items():
            if isinstance(value, list):
                # 包含任一指定字符串
                filtered_files = [f for f in filtered_files if any(v in f for v in value)]
            elif isinstance(value, dict) and 'exclude' in value:
                # 排除指定字符串
                filtered_files = [f for f in filtered_files if all(v not in f for v in value['exclude'])]
            else:
                # 包含指定字符串
                filtered_files = [f for f in filtered_files if value in f]

    return filtered_files


def plot_comparison(dataframes, columns, title="数据比较", figsize=(12, 8)):
    """绘制多个数据框指定列的比较图

    Args:
        dataframes: 字典，键为数据框名称，值为数据框
        columns: 要比较的列列表
        title: 图表标题
        figsize: 图表大小

    Returns:
        matplotlib图表对象
    """


    # 创建图表
    fig, axes = plt.subplots(len(columns), 1, figsize=figsize)
    if len(columns) == 1:
        axes = [axes]

    # 为每列数据绘制子图
    for i, column in enumerate(columns):
        ax = axes[i]

        # 绘制每个数据框的数据
        for name, df in dataframes.items():
            if column in df.columns:
                    sns.kdeplot(df[column].dropna(), ax=ax, label=name)
            else:
                # 使用matplotlib的直方图代替密度图
                    values = df[column].dropna()
                    ax.hist(values, bins=20, alpha=0.5, density=True, label=name)

        ax.set_title(f"{column}分布")
        ax.set_xlabel(column)
        ax.set_ylabel("概率密度")
        ax.legend()

    plt.tight_layout()
    plt.suptitle(title, fontsize=16, y=1.02)

    return fig

# ===== 程序执行入口 =====
if __name__ == "__main__":
    print("=" * 80)
    print("LIMS数据处理程序 - 电池数据分析工具")
    print("=" * 80)

    # 提示用户输入文件夹路径
    folder_path = input("请输入数据文件夹路径: ").strip()

    if not folder_path:
        print("错误: 必须指定文件夹路径")
        exit(1)

    print(f"使用路径: {folder_path}")

    # 提示用户选择异常检测方法
    print("\n请选择异常数据剔除方法:")
    print("1. 改良箱线图方法")
    print("2. Z-score+MAD方法")

    while True:
        choice = input("请输入选择 (1 或 2): ").strip()
        if choice == '1':
            CONFIG["OUTLIER_DETECTION"]["method"] = "boxplot"
            print("已选择: 改良箱线图方法")
            break
        elif choice == '2':
            CONFIG["OUTLIER_DETECTION"]["method"] = "zscore_mad"
            print("已选择: Z-score+MAD方法")
            break
        else:
            print("无效选择，请输入 1 或 2")

    # 调用主程序
    try:
        main(folder_path)
        print("\n程序执行完成！")
    except Exception as e:
        import traceback
        print("\n程序执行过程中发生错误:")
        print(traceback.format_exc())
        print("请检查数据文件是否正确，或联系开发人员。")