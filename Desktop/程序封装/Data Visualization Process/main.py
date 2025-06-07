#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LIMS数据处理程序 - 电池数据分析工具
主程序入口

主要功能：
1. 解析配置参数（命令行或交互模式）
2. 协调各个功能模块的执行
3. 提供统一的错误处理和日志记录
4. 确保配置参数正确传递给各个模块

版本: v2.0 (模块化版本)
日期: 2024
"""

import os
import sys
from typing import Optional

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from modules.config_parser import ConfigParser
from main_processor import MainProcessor


def interactive_mode() -> Optional[list]:
    """交互模式 - 让用户选择文件夹和配置

    Returns:
        Optional[list]: 用户选择的参数列表，如果取消则返回None
    """
    print("=" * 80)
    print("LIMS数据处理程序 - 电池数据分析工具")
    print("=" * 80)
    print()

    # 获取文件夹路径
    while True:
        folder_path = input("请输入数据文件夹路径 (或输入 'q' 退出): ").strip()

        if folder_path.lower() == 'q':
            print("用户取消操作")
            return None

        if not folder_path:
            print("错误: 必须指定文件夹路径")
            continue

        # 移除引号（如果有）
        folder_path = folder_path.strip('"\'')

        if not os.path.exists(folder_path):
            print(f"错误: 指定的文件夹不存在: {folder_path}")
            continue

        print(f"使用路径: {folder_path}")
        break

    # 选择异常检测方法
    print("\n请选择异常数据剔除方法:")
    print("1. 改良箱线图方法 (推荐)")
    print("2. Z-score+MAD方法")

    outlier_method = 'boxplot'  # 默认值
    while True:
        choice = input("请输入选择 (1 或 2，直接回车使用默认): ").strip()

        if choice == '' or choice == '1':
            outlier_method = 'boxplot'
            print("已选择: 改良箱线图方法")
            break
        elif choice == '2':
            outlier_method = 'zscore_mad'
            print("已选择: Z-score+MAD方法")
            break
        else:
            print("无效选择，请输入 1 或 2")

    # 选择参考通道方法
    print("\n请选择参考通道选择方法:")
    print("1. 传统方法 (基于首放容量)")
    print("2. PCA多特征分析")
    print("3. 保留率曲线MSE比较 (推荐)")

    reference_method = 'retention_curve_mse'  # 默认值
    while True:
        choice = input("请输入选择 (1、2 或 3，直接回车使用默认): ").strip()

        if choice == '' or choice == '3':
            reference_method = 'retention_curve_mse'
            print("已选择: 保留率曲线MSE比较")
            break
        elif choice == '1':
            reference_method = 'traditional'
            print("已选择: 传统方法")
            break
        elif choice == '2':
            reference_method = 'pca'
            print("已选择: PCA多特征分析")
            break
        else:
            print("无效选择，请输入 1、2 或 3")

    # 是否显示详细输出
    verbose = False
    verbose_choice = input("\n是否显示详细输出? (y/N): ").strip().lower()
    if verbose_choice in ['y', 'yes', '是']:
        verbose = True
        print("已启用详细输出")

    print("\n开始处理...")

    # 构建参数列表
    args = [
        '--input_folder', folder_path,
        '--outlier_method', outlier_method,
        '--reference_channel_method', reference_method
    ]

    if verbose:
        args.append('--verbose')

    return args


def main():
    """主程序入口"""
    try:
        # 检查是否有命令行参数
        if len(sys.argv) == 1:
            # 没有命令行参数，使用交互模式
            args = interactive_mode()
            if args is None:
                return 0
        else:
            # 有命令行参数，直接使用
            args = sys.argv[1:]

        # 解析配置
        config_parser = ConfigParser()
        config = config_parser.parse_arguments(args)

        print(f"\n配置信息:")
        print(f"  输入文件夹: {config.input_folder}")
        print(f"  异常检测方法: {config.outlier_method}")
        print(f"  参考通道方法: {config.reference_channel_method}")
        print(f"  详细输出: {config.verbose}")
        print(f"  默认1C圈数: {config.default_1c_cycle}")
        print()

        # 创建主处理器并运行
        processor = MainProcessor(config)
        success = processor.run()

        if success:
            print("\n🎉 程序执行完成！")
            return 0
        else:
            print("\n❌ 程序执行失败，请检查日志文件")
            return 1

    except KeyboardInterrupt:
        print("\n\n用户中断程序执行")
        return 1
    except Exception as e:
        print(f"\n❌ 程序执行过程中发生错误: {str(e)}")
        import traceback
        print("\n详细错误信息:")
        print(traceback.format_exc())
        print("\n请检查数据文件是否正确，或联系开发人员。")
        return 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
