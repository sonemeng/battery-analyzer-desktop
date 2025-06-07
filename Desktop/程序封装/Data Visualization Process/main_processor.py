#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主程序处理器 - 集成所有模块的完整数据处理流程
"""

import os
import sys
import time
import glob
from typing import Dict, List, Optional, Any
import pandas as pd
from tqdm import tqdm

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 导入模块化系统
from modules.config_parser import Config
from modules.file_parser import FileParser
from modules.data_processor import DataProcessor
from modules.outlier_detection import OutlierDetector
from utils.logger import ProcessingLogger


class MainProcessor:
    """主程序处理器 - 集成所有模块的完整数据处理流程"""

    def __init__(self, config):
        """初始化主程序处理器

        Args:
            config: 配置对象
        """
        self.config = config
        self.logger = ProcessingLogger(config.input_folder)

        # 初始化各个模块
        self.file_parser = FileParser(config, self.logger)
        self.data_processor = DataProcessor(config, self.logger)
        self.outlier_detector = OutlierDetector(config, self.logger)
        
        # 数据容器
        self.all_cycle_data = pd.DataFrame()
        self.first_cycle_files = []  # 仅1圈的文件
        self.error_files = []  # 异常数据文件
        self.inconsistent_data = pd.DataFrame()
        self.statistics_data = pd.DataFrame()
        
        # 处理统计
        self.total_processed = 0
        self.total_successful = 0
        self.start_time = None
        
        self.logger.log_info("主程序处理器初始化完成")
        self.logger.log_debug(f"输入文件夹: {config.input_folder}")
        self.logger.log_debug(f"异常检测方法: {config.outlier_method}")
        self.logger.log_debug(f"参考通道方法: {config.reference_channel_method}")
    
    def run(self) -> bool:
        """运行完整的数据处理流程

        Returns:
            bool: 处理是否成功
        """
        try:
            self.start_time = time.time()

            self.logger.log_info("=" * 80)
            self.logger.log_info("开始数据处理流程（使用模块化引擎）")
            self.logger.log_info("=" * 80)

            # 1. 验证输入文件夹
            if not self._validate_input_folder():
                return False

            # 2. 发现和分组文件
            file_groups = self._discover_and_group_files()
            if not file_groups:
                self.logger.log_warning("未发现任何可处理的文件")
                return False

            # 3. 处理所有文件
            self._process_all_files(file_groups)

            # 4. 处理异常文件
            self._process_error_files()

            # 5. 处理首圈文件
            self._process_first_cycle_files()

            # 6. 异常检测
            if hasattr(self, 'all_cycle_data') and not self.all_cycle_data.empty:
                self._perform_outlier_detection()

            # 7. 计算统计数据
            self._calculate_statistics()

            # 8. 导出结果
            start_time_str = time.strftime('%H%M%S', time.localtime(self.start_time))
            self._export_results(start_time_str)

            # 9. 输出处理总结
            self._print_processing_summary()

            return True

        except Exception as e:
            self.logger.log_error(f"数据处理流程发生错误: {str(e)}")
            import traceback
            self.logger.log_debug(f"详细错误信息:\n{traceback.format_exc()}")
            return False

        finally:
            # 确保日志系统正确关闭
            if hasattr(self, 'logger'):
                self.logger.close()
    
    # 以下方法保留用于未来扩展
    def _validate_input_folder(self) -> bool:
        """验证输入文件夹
        
        Returns:
            bool: 验证是否通过
        """
        if not os.path.exists(self.config.input_folder):
            self.logger.log_error(f"输入文件夹不存在: {self.config.input_folder}")
            return False
        
        # 检查是否有Excel文件
        excel_files = glob.glob(os.path.join(self.config.input_folder, "*.xlsx"))
        if not excel_files:
            self.logger.log_warning(f"输入文件夹中没有Excel文件: {self.config.input_folder}")
            self.logger.log_info("将只生成空的汇总表")
        else:
            self.logger.log_info(f"发现 {len(excel_files)} 个Excel文件")
        
        return True
    
    def _discover_and_group_files(self) -> Dict[str, List[str]]:
        """发现和分组文件
        
        Returns:
            Dict[str, List[str]]: 按系列分组的文件路径字典
        """
        self.logger.log_info("开始文件发现和分组...")
        
        # 获取所有Excel文件
        excel_files = glob.glob(os.path.join(self.config.input_folder, "*.xlsx"))
        
        # 按系列分组
        file_groups = {}
        for file_path in excel_files:
            try:
                series = self.file_parser._identify_series_from_filename(os.path.basename(file_path))
                if series not in file_groups:
                    file_groups[series] = []
                file_groups[series].append(file_path)
            except Exception as e:
                self.logger.log_warning(f"文件分组失败: {os.path.basename(file_path)}, 错误: {str(e)}")
                # 添加到默认组
                if 'UNKNOWN' not in file_groups:
                    file_groups['UNKNOWN'] = []
                file_groups['UNKNOWN'].append(file_path)
        
        # 输出分组结果
        for series, files in file_groups.items():
            self.logger.log_info(f"系列 {series}: {len(files)} 个文件")
        
        return file_groups
    
    def _process_all_files(self, file_groups: Dict[str, List[str]]):
        """处理所有文件
        
        Args:
            file_groups: 按系列分组的文件路径字典
        """
        self.logger.log_info("开始处理所有文件...")
        
        all_results = []
        
        for series_name, files in file_groups.items():
            if not files:
                self.logger.log_info(f"系列 {series_name} 没有文件，跳过处理")
                continue
            
            self.logger.log_info(f"正在处理系列 {series_name}，共 {len(files)} 个文件")
            
            series_results = []
            series_successful = 0
            
            # 使用进度条处理文件
            for file_path in tqdm(files, desc=f"处理{series_name}系列", ncols=100):
                self.total_processed += 1
                
                try:
                    result = self._process_single_file(file_path, series_name)
                    if result:
                        series_results.append(result)
                        series_successful += 1
                        self.total_successful += 1
                        
                except Exception as e:
                    self.logger.log_warning(f"处理文件失败: {os.path.basename(file_path)}, 错误: {str(e)}")
            
            self.logger.log_info(f"系列 {series_name} 处理完成: {series_successful}/{len(files)} 个文件成功")
            
            if series_results:
                all_results.extend(series_results)
        
        # 合并所有结果
        if all_results:
            # 直接从字典列表创建DataFrame
            self.all_cycle_data = pd.DataFrame(all_results)
            self.logger.log_info(f"总共处理了 {len(all_results)} 条有效数据记录")
        else:
            self.logger.log_warning("没有有效的数据记录")
    
    def _process_single_file(self, file_path: str, series_name: str) -> Optional[List[Any]]:
        """处理单个文件
        
        Args:
            file_path: 文件路径
            series_name: 系列名称
            
        Returns:
            Optional[List[Any]]: 处理后的数据记录，如果失败则返回None
        """
        file_name = os.path.basename(file_path)
        
        try:
            # 1. 解析文件信息
            file_info = self.file_parser.parse_file_info(file_path)
            if not file_info:
                self.logger.log_warning(f"文件信息解析失败: {file_name}")
                return None
            
            # 2. 读取循环数据
            cycle_df = self.file_parser.read_cycle_data(file_path)
            if cycle_df is None or cycle_df.empty:
                self.logger.log_warning(f"循环数据读取失败或为空: {file_name}")
                return None
            
            # 3. 检查数据有效性
            if len(cycle_df) == 1:
                self.logger.log_debug(f"文件只有1个循环，添加到首圈文件列表: {file_name}")
                self.first_cycle_files.append((file_path, series_name))
                return None
            
            # 4. 检查首圈数据是否异常
            if self.data_processor.is_abnormal_first_cycle(cycle_df):
                self.logger.log_warning(f"首圈数据异常，添加到异常文件列表: {file_name}")
                self.error_files.append((file_path, series_name))
                return None

            # 5. 处理循环数据
            result = self.data_processor.process_cycle_data(cycle_df, file_info, series_name)
            if result:
                self.logger.log_debug(f"文件处理成功: {file_name}")
                return result  # 直接返回字典
            else:
                self.logger.log_warning(f"循环数据处理失败: {file_name}")
                return None
                
        except Exception as e:
            self.logger.log_error(f"处理文件时发生异常: {file_name}, 错误: {str(e)}")
            return None

    def _process_error_files(self):
        """处理异常文件"""
        if not self.error_files:
            self.logger.log_info("没有异常文件需要处理")
            return

        self.logger.log_info(f"开始处理 {len(self.error_files)} 个异常文件...")

        error_results = []
        for file_path, series_name in self.error_files:
            try:
                file_name = os.path.basename(file_path)
                self.logger.log_debug(f"处理异常文件: {file_name}")

                # 解析文件信息
                file_info = self.file_parser.parse_file_info(file_path)
                if not file_info:
                    continue

                # 读取循环数据
                cycle_df = self.file_parser.read_cycle_data(file_path)
                if cycle_df is None or cycle_df.empty:
                    continue

                # 创建异常数据记录
                error_record = {
                    '系列': series_name,
                    '主机': file_info['device_id'],
                    '通道': file_info['channel_id'],
                    '批次': file_info['batch_id'],
                    '上架时间': file_info['shelf_time'],
                    '模式': file_info['mode'],
                    '活性物质': file_info['mass'],
                    '异常原因': '首圈数据异常',
                    '文件名': file_name
                }

                # 如果有数据，添加首圈信息
                if len(cycle_df) > 0:
                    first_cycle = cycle_df.iloc[0]
                    error_record.update({
                        '首充': first_cycle.get('充电比容量(mAh/g)', 0),
                        '首放': first_cycle.get('放电比容量(mAh/g)', 0),
                        '首圈电压': first_cycle.get('放电中值电压(V)', 0),
                        '首圈能量': first_cycle.get('放电比能量(mWh/g)', 0)
                    })

                    # 计算首效
                    if error_record['首充'] > 0:
                        error_record['首效'] = (error_record['首放'] / error_record['首充']) * 100

                error_results.append(error_record)

            except Exception as e:
                self.logger.log_warning(f"处理异常文件失败: {os.path.basename(file_path)}, 错误: {str(e)}")

        if error_results:
            self.inconsistent_data = pd.DataFrame(error_results)
            self.logger.log_info(f"异常文件处理完成，共 {len(error_results)} 条记录")

    def _process_first_cycle_files(self):
        """处理首圈文件"""
        if not self.first_cycle_files:
            self.logger.log_info("没有首圈文件需要处理")
            return

        self.logger.log_info(f"开始处理 {len(self.first_cycle_files)} 个首圈文件...")

        first_cycle_results = []
        for file_path, series_name in self.first_cycle_files:
            try:
                file_name = os.path.basename(file_path)
                self.logger.log_debug(f"处理首圈文件: {file_name}")

                # 解析文件信息
                file_info = self.file_parser.parse_file_info(file_path)
                if not file_info:
                    continue

                # 读取循环数据
                cycle_df = self.file_parser.read_cycle_data(file_path)
                if cycle_df is None or cycle_df.empty:
                    continue

                # 创建首圈数据记录
                first_cycle = cycle_df.iloc[0]
                first_cycle_record = {
                    '系列': series_name,
                    '主机': file_info['device_id'],
                    '通道': file_info['channel_id'],
                    '批次': file_info['batch_id'],
                    '上架时间': file_info['shelf_time'],
                    '模式': file_info['mode'],
                    '活性物质': file_info['mass'],
                    '首充': first_cycle.get('充电比容量(mAh/g)', 0),
                    '首放': first_cycle.get('放电比容量(mAh/g)', 0),
                    '首圈电压': first_cycle.get('放电中值电压(V)', 0),
                    '首圈能量': first_cycle.get('放电比能量(mWh/g)', 0),
                    '当前圈数': 1,
                    '文件名': file_name
                }

                # 计算首效
                if first_cycle_record['首充'] > 0:
                    first_cycle_record['首效'] = (first_cycle_record['首放'] / first_cycle_record['首充']) * 100

                first_cycle_results.append(first_cycle_record)

            except Exception as e:
                self.logger.log_warning(f"处理首圈文件失败: {os.path.basename(file_path)}, 错误: {str(e)}")

        if first_cycle_results:
            # 将首圈数据添加到主数据中（如果主数据为空）或创建单独的首圈数据表
            if self.all_cycle_data.empty:
                # 如果主数据为空，将首圈数据作为主数据
                self.all_cycle_data = pd.DataFrame(first_cycle_results)
                self.logger.log_info(f"首圈文件处理完成，作为主数据，共 {len(first_cycle_results)} 条记录")
            else:
                # 否则记录到统计数据中
                self.logger.log_info(f"首圈文件处理完成，共 {len(first_cycle_results)} 条记录")

    def _perform_outlier_detection(self):
        """执行异常检测"""
        if self.all_cycle_data.empty:
            self.logger.log_info("没有数据需要进行异常检测")
            return

        self.logger.log_info("开始异常检测...")
        original_count = len(self.all_cycle_data)

        try:
            # 执行异常检测
            cleaned_data = self.outlier_detector.detect_and_remove_outliers(self.all_cycle_data)

            if cleaned_data is not None:
                removed_count = original_count - len(cleaned_data)
                self.all_cycle_data = cleaned_data

                self.logger.log_info(f"异常检测完成: 原始数据 {original_count} 条，移除 {removed_count} 条，剩余 {len(cleaned_data)} 条")
            else:
                self.logger.log_warning("异常检测返回空结果，保持原始数据")

        except Exception as e:
            self.logger.log_error(f"异常检测过程中发生错误: {str(e)}")
            self.logger.log_info("保持原始数据不变")

    def _calculate_statistics(self):
        """计算统计数据"""
        self.logger.log_info("开始计算统计数据...")

        try:
            if self.all_cycle_data.empty:
                self.logger.log_warning("没有数据用于统计计算")
                return

            # 按系列和批次分组统计
            stats_list = []

            for series in self.all_cycle_data['系列'].unique():
                series_data = self.all_cycle_data[self.all_cycle_data['系列'] == series]

                for batch in series_data['批次'].unique():
                    batch_data = series_data[series_data['批次'] == batch]

                    if batch_data.empty:
                        continue

                    # 计算统计指标
                    stats = {
                        '系列': series,
                        '批次': batch,
                        '样品数量': len(batch_data),
                        '首放平均值': batch_data['首放'].mean(),
                        '首放标准差': batch_data['首放'].std(),
                        '首效平均值': batch_data['首效'].mean(),
                        '首效标准差': batch_data['首效'].std(),
                        '首圈电压平均值': batch_data['首圈电压'].mean(),
                        '首圈电压标准差': batch_data['首圈电压'].std(),
                        '当前容量保持平均值': batch_data['当前容量保持'].mean(),
                        '当前容量保持标准差': batch_data['当前容量保持'].std()
                    }

                    # 添加1C相关统计
                    one_c_data = batch_data[batch_data['1C首圈编号'].notna()]
                    if not one_c_data.empty:
                        stats.update({
                            '1C样品数量': len(one_c_data),
                            '1C首放平均值': one_c_data['1C首放'].mean(),
                            '1C首效平均值': one_c_data['1C首效'].mean(),
                            '1C倍率比平均值': one_c_data['1C倍率比'].mean()
                        })

                    stats_list.append(stats)

            if stats_list:
                self.statistics_data = pd.DataFrame(stats_list)
                self.logger.log_info(f"统计数据计算完成，共 {len(stats_list)} 个批次的统计")
            else:
                self.logger.log_warning("没有生成统计数据")

        except Exception as e:
            self.logger.log_error(f"统计数据计算过程中发生错误: {str(e)}")

    def _export_results(self, start_time_str: str):
        """导出结果

        Args:
            start_time_str: 开始时间字符串
        """
        self.logger.log_info("开始导出结果...")

        try:
            # 生成输出文件名
            output_filename = f"电池数据汇总表-{start_time_str}.xlsx"
            output_path = os.path.join(self.config.input_folder, output_filename)

            # 使用ExcelWriter写入多个工作表
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # 主数据表
                if not self.all_cycle_data.empty:
                    self.all_cycle_data.to_excel(writer, sheet_name='主数据', index=False)
                    self.logger.log_info(f"主数据表导出完成: {len(self.all_cycle_data)} 条记录")

                # 统计数据表
                if not self.statistics_data.empty:
                    self.statistics_data.to_excel(writer, sheet_name='统计数据', index=False)
                    self.logger.log_info(f"统计数据表导出完成: {len(self.statistics_data)} 条记录")

                # 异常数据表
                if not self.inconsistent_data.empty:
                    self.inconsistent_data.to_excel(writer, sheet_name='异常数据', index=False)
                    self.logger.log_info(f"异常数据表导出完成: {len(self.inconsistent_data)} 条记录")

                # 如果所有表都为空，创建一个空的主数据表
                if (self.all_cycle_data.empty and
                    self.statistics_data.empty and
                    self.inconsistent_data.empty):
                    empty_df = pd.DataFrame(columns=['系列', '主机', '通道', '批次', '说明'])
                    empty_df.loc[0] = ['', '', '', '', '未发现有效数据']
                    empty_df.to_excel(writer, sheet_name='主数据', index=False)
                    self.logger.log_info("导出空数据表")

            self.logger.log_info(f"结果导出完成: {output_path}")

        except Exception as e:
            self.logger.log_error(f"结果导出过程中发生错误: {str(e)}")

    def _print_processing_summary(self):
        """输出处理总结"""
        end_time = time.time()
        execution_time = end_time - self.start_time if self.start_time else 0

        self.logger.log_info("=" * 80)
        self.logger.log_info("数据处理完成总结")
        self.logger.log_info("=" * 80)
        self.logger.log_info(f"总处理文件数: {self.total_processed}")
        self.logger.log_info(f"成功处理文件数: {self.total_successful}")
        self.logger.log_info(f"主数据记录数: {len(self.all_cycle_data)}")
        self.logger.log_info(f"异常文件数: {len(self.error_files)}")
        self.logger.log_info(f"首圈文件数: {len(self.first_cycle_files)}")
        self.logger.log_info(f"统计批次数: {len(self.statistics_data) if not self.statistics_data.empty else 0}")
        self.logger.log_info(f"总执行时间: {execution_time:.2f} 秒")
        self.logger.log_info("=" * 80)
