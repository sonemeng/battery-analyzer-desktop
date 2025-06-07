#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
日志系统模块
从原始脚本中提取的日志功能，支持同时输出到控制台和文件
"""

import os
import sys
import time
from typing import List, TextIO


class TeeOutput:
    """同时输出到控制台和文件的类"""
    
    def __init__(self, *files: TextIO):
        """初始化TeeOutput
        
        Args:
            *files: 要同时写入的文件对象列表
        """
        self.files = files

    def write(self, text: str):
        """写入文本到所有文件
        
        Args:
            text: 要写入的文本
        """
        for file in self.files:
            file.write(text)
            file.flush()

    def flush(self):
        """刷新所有文件缓冲区"""
        for file in self.files:
            file.flush()


class ProcessingLogger:
    """处理日志管理器
    
    负责管理所有类型的日志输出，包括：
    - 主要处理日志（同时输出到控制台和文件）
    - 异常检测详细日志
    - 调试详细日志
    """
    
    def __init__(self, output_dir: str = None):
        """初始化日志管理器
        
        Args:
            output_dir: 输出目录，如果为None则使用当前工作目录
        """
        self.output_dir = output_dir or os.getcwd()
        self.timestamp = time.strftime('%Y%m%d_%H%M%S')

        # 创建日志文件夹
        self.log_dir = os.path.join(self.output_dir, f"处理日志-{self.timestamp}")
        os.makedirs(self.log_dir, exist_ok=True)

        # 创建不同类型的日志文件（使用无缓冲模式确保实时写入）
        self.main_log_file = open(
            os.path.join(self.log_dir, "主要处理日志.txt"), 
            "w", 
            encoding='utf-8', 
            buffering=1
        )
        self.outlier_log_file = open(
            os.path.join(self.log_dir, "异常检测详细日志.txt"), 
            "w", 
            encoding='utf-8', 
            buffering=1
        )
        self.debug_log_file = open(
            os.path.join(self.log_dir, "调试详细日志.txt"), 
            "w", 
            encoding='utf-8', 
            buffering=1
        )

        # 保存原始stdout
        self.original_stdout = sys.stdout

        # 设置tee输出（同时输出到控制台和主日志）
        sys.stdout = TeeOutput(self.original_stdout, self.main_log_file)

        print(f"日志系统已启动，日志保存到: {self.log_dir}")
        print("=" * 80)

    def log_outlier_detection(self, message: str):
        """记录异常检测相关信息
        
        Args:
            message: 要记录的消息
        """
        timestamp = time.strftime('%H:%M:%S')
        self.outlier_log_file.write(f"[{timestamp}] {message}\n")
        self.outlier_log_file.flush()

    def log_debug(self, message: str):
        """记录调试信息
        
        Args:
            message: 要记录的调试消息
        """
        timestamp = time.strftime('%H:%M:%S')
        self.debug_log_file.write(f"[{timestamp}] {message}\n")
        self.debug_log_file.flush()

    def log_info(self, message: str):
        """记录一般信息（输出到主日志）
        
        Args:
            message: 要记录的信息
        """
        print(message)

    def log_warning(self, message: str):
        """记录警告信息
        
        Args:
            message: 警告消息
        """
        warning_msg = f"警告: {message}"
        print(warning_msg)
        self.log_debug(f"WARNING: {message}")

    def log_error(self, message: str):
        """记录错误信息
        
        Args:
            message: 错误消息
        """
        error_msg = f"错误: {message}"
        print(error_msg)
        self.log_debug(f"ERROR: {message}")

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

    def get_log_dir(self) -> str:
        """获取日志目录路径
        
        Returns:
            str: 日志目录路径
        """
        return self.log_dir

    def __enter__(self):
        """上下文管理器入口"""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器出口"""
        self.close()
