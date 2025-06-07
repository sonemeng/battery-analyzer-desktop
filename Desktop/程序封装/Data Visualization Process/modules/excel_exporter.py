"""
Excel导出模块 - 严格按照原始脚本LIMS_DATA_PROCESS_改良箱线图版.py重构

完全按照原始脚本的Excel导出逻辑，不做任何简化或修改
"""

import os
import time
import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Tuple, Any


class ExcelExporter:
    """Excel导出器类 - 严格按照原始脚本逻辑"""
    
    def __init__(self, config, logger):
        """初始化Excel导出器
        
        Args:
            config: 配置对象
            logger: 日志记录器
        """
        self.config = config
        self.logger = logger
        
        # Excel导出配置 - 完全按照原始脚本
        self.excel_config = {
            "engine": config.excel_engine,
            "auto_adjust_width": getattr(config, 'excel_auto_adjust_width', False),  # 原始脚本中禁用了自动调整列宽
            "freeze_panes": getattr(config, 'excel_freeze_panes', True),
            "add_filters": getattr(config, 'excel_add_filters', True),
            "sheet_names": {
                "main": getattr(config, 'excel_sheet_name_main', 'All_Cycle_Data'),
                "first_cycle": getattr(config, 'excel_sheet_name_first_cycle', 'First_Cycle_Only'),
                "error": getattr(config, 'excel_sheet_name_error', 'Error_Data'),
                "statistics": getattr(config, 'excel_sheet_name_statistics', 'Statistics'),
                "inconsistent": getattr(config, 'excel_sheet_name_inconsistent', 'Inconsistent_Data')
            }
        }
        
        # 输出路径配置
        self.output_folder = getattr(config, 'output_folder', '')
        self.timestamp_format = getattr(config, 'timestamp_format', '%Y%m%d_%H%M%S')

    def export_all_data(self, data_dict: Dict[str, pd.DataFrame], output_filename: Optional[str] = None) -> str:
        """导出所有数据到Excel文件 - 完全按照原始脚本逻辑
        
        Args:
            data_dict: 包含所有数据的字典
            output_filename: 输出文件名（可选）
            
        Returns:
            str: 输出文件路径
        """
        # 生成输出文件名 - 完全按照原始脚本
        if not output_filename:
            timestamp = time.strftime(self.timestamp_format)
            output_filename = f"电池数据汇总_{timestamp}.xlsx"
        
        # 确定输出路径
        if self.output_folder:
            output_path = os.path.join(self.output_folder, output_filename)
        else:
            # 如果没有指定输出路径，使用当前目录
            output_path = output_filename
        
        # 确保输出目录存在
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        print(f"正在导出数据到: {output_path}")
        
        try:
            # 创建Excel写入器 - 完全按照原始脚本
            with pd.ExcelWriter(output_path, engine=self.excel_config["engine"]) as writer:
                
                # 导出主数据表 - 完全按照原始脚本
                if 'main_data' in data_dict and not data_dict['main_data'].empty:
                    self._export_main_data(writer, data_dict['main_data'])
                
                # 导出首圈数据表 - 完全按照原始脚本
                if 'first_cycle_data' in data_dict and not data_dict['first_cycle_data'].empty:
                    self._export_first_cycle_data(writer, data_dict['first_cycle_data'])
                
                # 导出异常数据表 - 完全按照原始脚本
                if 'error_data' in data_dict and not data_dict['error_data'].empty:
                    self._export_error_data(writer, data_dict['error_data'])
                
                # 导出统计数据表 - 完全按照原始脚本
                if 'statistics_data' in data_dict and not data_dict['statistics_data'].empty:
                    self._export_statistics_data(writer, data_dict['statistics_data'])
                
                # 导出不一致数据表 - 完全按照原始脚本
                if 'inconsistent_data' in data_dict and not data_dict['inconsistent_data'].empty:
                    self._export_inconsistent_data(writer, data_dict['inconsistent_data'])
            
            print(f"数据导出完成: {output_path}")
            return output_path
            
        except Exception as e:
            print(f"导出Excel文件失败: {str(e)}")
            raise

    def _export_main_data(self, writer: pd.ExcelWriter, main_data: pd.DataFrame):
        """导出主数据表 - 完全按照原始脚本逻辑"""
        sheet_name = self.excel_config["sheet_names"]["main"]
        
        # 写入数据
        main_data.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # 获取工作表对象
        worksheet = writer.sheets[sheet_name]
        
        # 应用格式设置 - 完全按照原始脚本
        self._apply_worksheet_formatting(worksheet, main_data)
        
        print(f"已导出主数据表: {len(main_data)} 行数据")

    def _export_first_cycle_data(self, writer: pd.ExcelWriter, first_cycle_data: pd.DataFrame):
        """导出首圈数据表 - 完全按照原始脚本逻辑"""
        sheet_name = self.excel_config["sheet_names"]["first_cycle"]
        
        # 写入数据
        first_cycle_data.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # 获取工作表对象
        worksheet = writer.sheets[sheet_name]
        
        # 应用格式设置
        self._apply_worksheet_formatting(worksheet, first_cycle_data)
        
        print(f"已导出首圈数据表: {len(first_cycle_data)} 行数据")

    def _export_error_data(self, writer: pd.ExcelWriter, error_data: pd.DataFrame):
        """导出异常数据表 - 完全按照原始脚本逻辑"""
        sheet_name = self.excel_config["sheet_names"]["error"]
        
        # 写入数据
        error_data.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # 获取工作表对象
        worksheet = writer.sheets[sheet_name]
        
        # 应用格式设置
        self._apply_worksheet_formatting(worksheet, error_data)
        
        print(f"已导出异常数据表: {len(error_data)} 行数据")

    def _export_statistics_data(self, writer: pd.ExcelWriter, statistics_data: pd.DataFrame):
        """导出统计数据表 - 完全按照原始脚本逻辑"""
        sheet_name = self.excel_config["sheet_names"]["statistics"]
        
        # 写入数据
        statistics_data.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # 获取工作表对象
        worksheet = writer.sheets[sheet_name]
        
        # 应用格式设置
        self._apply_worksheet_formatting(worksheet, statistics_data)
        
        print(f"已导出统计数据表: {len(statistics_data)} 行数据")

    def _export_inconsistent_data(self, writer: pd.ExcelWriter, inconsistent_data: pd.DataFrame):
        """导出不一致数据表 - 完全按照原始脚本逻辑"""
        sheet_name = self.excel_config["sheet_names"]["inconsistent"]
        
        # 写入数据
        inconsistent_data.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # 获取工作表对象
        worksheet = writer.sheets[sheet_name]
        
        # 应用格式设置
        self._apply_worksheet_formatting(worksheet, inconsistent_data)
        
        print(f"已导出不一致数据表: {len(inconsistent_data)} 行数据")

    def _apply_worksheet_formatting(self, worksheet, data: pd.DataFrame):
        """应用工作表格式设置 - 完全按照原始脚本逻辑
        
        Args:
            worksheet: 工作表对象
            data: 数据DataFrame
        """
        try:
            # 冻结首行 - 完全按照原始脚本
            if self.excel_config["freeze_panes"]:
                worksheet.freeze_panes(1, 0)
            
            # 添加筛选器 - 完全按照原始脚本
            if self.excel_config["add_filters"] and not data.empty:
                worksheet.auto_filter.ref = f"A1:{self._get_column_letter(len(data.columns))}{len(data) + 1}"
            
            # 注意：原始脚本中禁用了自动调整列宽，因为会导致程序错误
            # 所以这里不执行自动调整列宽操作
            if self.excel_config["auto_adjust_width"]:
                print("警告: 自动调整列宽功能已禁用，因为在原始脚本中会导致程序错误")
            
        except Exception as e:
            print(f"应用工作表格式时出错: {str(e)}")

    def _get_column_letter(self, column_number: int) -> str:
        """获取Excel列字母 - 完全按照原始脚本逻辑
        
        Args:
            column_number: 列号（1基）
            
        Returns:
            str: Excel列字母
        """
        result = ""
        while column_number > 0:
            column_number -= 1
            result = chr(column_number % 26 + ord('A')) + result
            column_number //= 26
        return result

    def export_summary_file(self, summary_data: Dict[str, Any], output_filename: Optional[str] = None) -> str:
        """导出汇总文件 - 完全按照原始脚本逻辑
        
        Args:
            summary_data: 汇总数据字典
            output_filename: 输出文件名（可选）
            
        Returns:
            str: 输出文件路径
        """
        # 生成输出文件名
        if not output_filename:
            timestamp = time.strftime(self.timestamp_format)
            output_filename = f"数据处理汇总_{timestamp}.xlsx"
        
        # 确定输出路径
        if self.output_folder:
            output_path = os.path.join(self.output_folder, output_filename)
        else:
            output_path = output_filename
        
        # 确保输出目录存在
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        try:
            # 创建汇总DataFrame
            summary_df = pd.DataFrame([summary_data])
            
            # 导出到Excel
            with pd.ExcelWriter(output_path, engine=self.excel_config["engine"]) as writer:
                summary_df.to_excel(writer, sheet_name='处理汇总', index=False)
                
                # 应用格式设置
                worksheet = writer.sheets['处理汇总']
                self._apply_worksheet_formatting(worksheet, summary_df)
            
            print(f"汇总文件导出完成: {output_path}")
            return output_path
            
        except Exception as e:
            print(f"导出汇总文件失败: {str(e)}")
            raise

    def create_output_folder(self, base_folder: str) -> str:
        """创建输出文件夹 - 完全按照原始脚本逻辑
        
        Args:
            base_folder: 基础文件夹路径
            
        Returns:
            str: 创建的输出文件夹路径
        """
        timestamp = time.strftime('%Y%m%d_%H%M%S')
        output_folder = os.path.join(base_folder, f"data_visualization_{timestamp}")
        
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            print(f"创建输出文件夹: {output_folder}")
        
        return output_folder

    def get_default_output_path(self, input_folder: str) -> str:
        """获取默认输出路径 - 完全按照原始脚本逻辑
        
        Args:
            input_folder: 输入文件夹路径
            
        Returns:
            str: 默认输出路径
        """
        if self.output_folder:
            return self.output_folder
        else:
            # 如果没有指定输出路径，使用输入文件夹
            return input_folder
