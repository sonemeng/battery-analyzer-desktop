"""
数据处理模块 - 严格按照原始脚本LIMS_DATA_PROCESS_改良箱线图版.py重构
"""

import os
import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Tuple, Any
from tqdm import tqdm


class DataProcessor:
    """数据处理器类 - 严格按照原始脚本逻辑"""

    def __init__(self, config, logger):
        self.config = config
        self.logger = logger

        # Excel列配置 - 完全按照原始脚本
        self.excel_cols = {
            'main': ['系列', '主机', '通道', '批次', '上架时间', '模式', '活性物质',
                    '首充', '首放', '首效', '首圈电压', '首圈能量', 'Cycle2', 'Cycle2充电比容量',
                    'Cycle3', 'Cycle3充电比容量', '1C首圈编号', '1C首充', '1C首放', '1C首效', '1C状态', '1C倍率比',
                    'Cycle4', 'Cycle4充电比容量', 'Cycle5', 'Cycle5充电比容量', 'Cycle6', 'Cycle6充电比容量',
                    'Cycle7', 'Cycle7充电比容量', '当前圈数', '当前容量保持', '电压衰减率mV/周',
                    '当前电压保持', '当前能量保持', '100容量保持', '100电压保持', '100能量保持',
                    '200容量保持', '200电压保持', '200能量保持'],
            'first_cycle': ['系列', '主机', '通道', '批次', '上架时间', '首充', '首放'],
            'error': ['系列', '主机', '通道', '批次', '上架时间', '首充', '首放', '当前圈数'],
            'statistics': ['系列', '统一批次', '上架时间', '总数据', '首周有效数据', '首充', '首放', '首效',
                        '首圈电压', '首圈能量', '活性物质', 'Cycle2', 'Cycle2充电比容量', 'Cycle3', 'Cycle3充电比容量',
                        'Cycle4', 'Cycle4充电比容量', 'Cycle5', 'Cycle5充电比容量', 'Cycle6', 'Cycle6充电比容量',
                        'Cycle7', 'Cycle7充电比容量', '1C首周有效数据', '1C参考通道', '1C首圈编号', '1C首充', '1C首放',
                        '1C首效', '1C倍率比', '1C状态', '参考通道当前圈数', '当前容量保持', '电压衰减率mV/周', '当前电压保持', '当前能量保持']
        }

        # 1C阈值配置
        self.one_c_thresholds = {
            "ratio_threshold": config.one_c_ratio_threshold,
            "discharge_diff_threshold": config.one_c_discharge_diff_threshold,
            "overcharge_threshold": config.one_c_overcharge_threshold,
            "very_low_efficiency_threshold": config.one_c_very_low_efficiency_threshold,
            "low_efficiency_threshold": config.one_c_low_efficiency_threshold
        }

        # 测试模式配置
        self.one_c_modes = ['-1C-']

        # 数据容器
        self.first_cycle_files = []
        self.error_files = []
        self.all_cycle_data = pd.DataFrame(columns=self.excel_cols['main'])
        self.all_first_cycle = pd.DataFrame(columns=self.excel_cols['first_cycle'])
        self.all_error_data = pd.DataFrame(columns=self.excel_cols['error'])
        self.statistics_data = pd.DataFrame()

        self.verbose = getattr(config, 'verbose', False)

    def process_all_files(self, file_groups: Dict[str, List[str]], file_parser) -> Tuple[int, int]:
        """处理所有电池数据文件"""
        total_processed = 0
        total_successful = 0

        for series_name, files in file_groups.items():
            if not files:
                print(f'系列 {series_name} 没有文件，跳过处理')
                continue

            print(f'正在处理系列 {series_name}，共 {len(files)} 个文件')
            results = []
            successful_files = 0

            for file_path in tqdm(files, desc=f"处理{series_name}系列", ncols=100):
                total_processed += 1
                try:
                    data = self._process_single_file(file_path, series_name, file_parser)
                    if data:
                        results.append(data)
                        successful_files += 1
                        total_successful += 1
                except Exception as e:
                    if self.verbose:
                        print(f"处理文件失败: {os.path.basename(file_path)}, 错误: {str(e)}")

            print(f"系列 {series_name} 处理完成: {successful_files}/{len(files)} 个文件成功")

            if results:
                cycle_df = pd.DataFrame(results, columns=self.excel_cols['main'])
                self.all_cycle_data = pd.concat([self.all_cycle_data, cycle_df], ignore_index=True)
                print(f"系列 {series_name} 添加了 {len(results)} 条数据记录")

        print(f"\n数据处理总结:")
        print(f"总共处理文件: {total_processed} 个")
        print(f"成功处理文件: {total_successful} 个")
        print(f"总有效数据记录: {len(self.all_cycle_data)} 条")

        return total_processed, total_successful

    def _process_single_file(self, file_path: str, series_name: str, file_parser) -> Optional[List[Any]]:
        """处理单个文件"""
        file_name = os.path.basename(file_path)

        if not os.path.exists(file_path):
            print(f"文件不存在: {file_path}")
            return None

        # 读取循环数据
        cycle_df = file_parser.read_cycle_data(file_path)
        if cycle_df is None:
            return None

        # 提取文件信息
        try:
            file_info = file_parser.parse_file_info(file_path)
            if not file_info or file_info.get('device_id') == 'error':
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
            if file_parser.is_abnormal_first_cycle(cycle_df):
                print(f"警告: 首圈数据异常，添加到异常文件列表: {file_name}")
                self.error_files.append((file_path, series_name))
                return None
        except Exception as e:
            print(f"检查首圈数据异常失败: {file_name}, 错误: {str(e)}")
            return None

        # 处理循环数据
        try:
            result = self._process_cycle_data(cycle_df, file_info, series_name)
            return result
        except Exception as e:
            print(f"处理循环数据失败: {file_name}, 错误: {str(e)}")
            return None

    def _process_cycle_data(self, cycle_df: pd.DataFrame, file_info: Dict[str, Any], series_name: str) -> Optional[List[Any]]:
        """处理循环数据 - 完全按照原始脚本逻辑"""
        try:
            # 基础数据提取
            first_charge = cycle_df.loc[0, '充电比容量(mAh/g)']
            first_discharge = cycle_df.loc[0, '放电比容量(mAh/g)']
            first_efficiency = (first_discharge / first_charge * 100) if first_charge > 0 else 0
            first_voltage = cycle_df.loc[0, '放电中值电压(V)']
            first_energy = cycle_df.loc[0, '放电比能量(mWh/g)']

            # 构建基础数据行
            result_row = [
                series_name,  # 系列
                file_info['device_id'],  # 主机
                file_info['channel_id'],  # 通道
                file_info['batch_id'],  # 批次
                file_info['shelf_time'],  # 上架时间
                file_info['mode'],  # 模式
                file_info.get('mass', ''),  # 活性物质
                first_charge,  # 首充
                first_discharge,  # 首放
                first_efficiency,  # 首效
                first_voltage,  # 首圈电压
                first_energy,  # 首圈能量
            ]

            # 添加Cycle2-7数据
            for cycle_num in range(2, 8):
                if cycle_num - 1 < len(cycle_df):
                    cycle_discharge = cycle_df.loc[cycle_num - 1, '放电比容量(mAh/g)']
                    cycle_charge = cycle_df.loc[cycle_num - 1, '充电比容量(mAh/g)']
                    result_row.extend([cycle_discharge, cycle_charge])
                else:
                    result_row.extend(['', ''])

            # 1C相关数据处理
            one_c_data = self._process_one_c_data(cycle_df, file_info)
            result_row.extend(one_c_data)

            # 当前圈数和保持率计算
            current_cycle = len(cycle_df)

            # 获取1C首圈编号用于所有保持率计算
            one_c_cycle_num = one_c_data[0] if one_c_data[0] != '' else None

            current_capacity_retention = self._calculate_capacity_retention(cycle_df, file_info, one_c_cycle_num)
            voltage_decay_rate = self._calculate_voltage_decay_rate(cycle_df, file_info, one_c_cycle_num)
            current_voltage_retention = self._calculate_voltage_retention(cycle_df, file_info, one_c_cycle_num)
            current_energy_retention = self._calculate_energy_retention(cycle_df, file_info, one_c_cycle_num)

            result_row.extend([
                current_cycle,  # 当前圈数
                current_capacity_retention,  # 当前容量保持
                voltage_decay_rate,  # 电压衰减率mV/周
                current_voltage_retention,  # 当前电压保持
                current_energy_retention,  # 当前能量保持
            ])

            # 100圈和200圈保持率
            retention_100 = self._calculate_retention_at_cycle(cycle_df, 100, file_info, one_c_cycle_num)
            retention_200 = self._calculate_retention_at_cycle(cycle_df, 200, file_info, one_c_cycle_num)
            result_row.extend(retention_100 + retention_200)

            return result_row

        except Exception as e:
            print(f"处理循环数据时出错: {str(e)}")
            return None

    def _process_one_c_data(self, cycle_df: pd.DataFrame, file_info: Dict[str, Any]) -> List[Any]:
        """处理1C相关数据 - 完全按照原始脚本逻辑"""
        mode = file_info.get('mode', '')

        if mode in self.one_c_modes:
            # 1C模式处理
            one_c_cycle_num = self._find_one_c_cycle(cycle_df)
            if one_c_cycle_num > 0:
                one_c_charge = cycle_df.loc[one_c_cycle_num - 1, '充电比容量(mAh/g)']
                one_c_discharge = cycle_df.loc[one_c_cycle_num - 1, '放电比容量(mAh/g)']
                one_c_efficiency = (one_c_discharge / one_c_charge * 100) if one_c_charge > 0 else 0
                one_c_status = self._determine_one_c_status(one_c_charge, one_c_discharge, one_c_efficiency)
                one_c_ratio = self._calculate_one_c_ratio(cycle_df, one_c_cycle_num)

                return [one_c_cycle_num, one_c_charge, one_c_discharge, one_c_efficiency, one_c_status, one_c_ratio]
            else:
                return ['', '', '', '', '无1C', '']
        else:
            # 非1C模式
            return ['', '', '', '', '非1C', '']

    def _find_one_c_cycle(self, cycle_df: pd.DataFrame) -> int:
        """查找1C首圈编号 - 完全按照原始脚本逻辑"""
        # 从第3圈开始查找（索引2）
        for i in range(2, len(cycle_df)):
            charge = cycle_df.loc[i, '充电比容量(mAh/g)']
            discharge = cycle_df.loc[i, '放电比容量(mAh/g)']

            # 检查是否满足1C条件
            if self._is_valid_one_c_cycle(charge, discharge):
                return i + 1  # 返回1基索引

        return 0  # 未找到

    def _is_valid_one_c_cycle(self, charge: float, discharge: float) -> bool:
        """检查是否为有效的1C循环 - 完全按照原始脚本逻辑"""
        # 检查过充
        if charge > self.one_c_thresholds["overcharge_threshold"]:
            return False

        # 检查效率
        efficiency = (discharge / charge * 100) if charge > 0 else 0
        if efficiency < self.one_c_thresholds["very_low_efficiency_threshold"]:
            return False

        return True

    def _determine_one_c_status(self, charge: float, discharge: float, efficiency: float) -> str:
        """确定1C状态 - 完全按照原始脚本逻辑"""
        if charge > self.one_c_thresholds["overcharge_threshold"]:
            return "过充"
        elif efficiency < self.one_c_thresholds["very_low_efficiency_threshold"]:
            return "极低效"
        elif efficiency < self.one_c_thresholds["low_efficiency_threshold"]:
            return "低效"
        else:
            return "正常"

    def _calculate_one_c_ratio(self, cycle_df: pd.DataFrame, one_c_cycle_num: int) -> float:
        """计算1C倍率比 - 完全按照原始脚本逻辑"""
        if one_c_cycle_num <= 1:
            return 0

        first_discharge = cycle_df.loc[0, '放电比容量(mAh/g)']
        one_c_discharge = cycle_df.loc[one_c_cycle_num - 1, '放电比容量(mAh/g)']

        if first_discharge > 0:
            return one_c_discharge / first_discharge
        else:
            return 0

    def _calculate_capacity_retention(self, cycle_df: pd.DataFrame, file_info: Dict[str, Any], one_c_cycle_num: Optional[int] = None) -> float:
        """计算当前容量保持率 - 严格按照原始脚本逻辑

        Args:
            cycle_df: 循环数据DataFrame
            file_info: 文件信息字典
            one_c_cycle_num: 1C首圈编号（1基索引）

        Returns:
            float: 当前容量保持率（百分比）
        """
        if len(cycle_df) <= 4:  # 原始脚本：只有当圈数>4时才计算
            return None

        mode = file_info.get('mode', '')
        total_cycles = len(cycle_df)
        current_discharge = cycle_df.loc[total_cycles - 1, '放电比容量(mAh/g)']

        # 0.1C模式 - 完全按照原始脚本第925行
        if mode == '-0.1C-':
            first_discharge = cycle_df.loc[0, '放电比容量(mAh/g)']
            capacity_retention = 100 * current_discharge / first_discharge
            return round(capacity_retention, 1)

        # 1C模式 - 完全按照原始脚本第949行
        elif mode in self.one_c_modes:
            # 确定1C首圈索引
            if one_c_cycle_num is not None and one_c_cycle_num > 0:
                one_c_idx = one_c_cycle_num - 1  # 转换为0基索引
            else:
                one_c_idx = min(3, total_cycles - 1)  # 默认第4圈或最大可用圈

            # 确保有足够的循环数据
            if total_cycles > one_c_idx + 1:
                one_c_discharge = cycle_df.loc[one_c_idx, '放电比容量(mAh/g)']
                capacity_retention = 100 * current_discharge / one_c_discharge
                return round(capacity_retention, 1)
            else:
                return None

        # 其他模式，使用首圈作为基准
        else:
            first_discharge = cycle_df.loc[0, '放电比容量(mAh/g)']
            capacity_retention = 100 * current_discharge / first_discharge
            return round(capacity_retention, 1)

    def _calculate_voltage_decay_rate(self, cycle_df: pd.DataFrame, file_info: Dict[str, Any], one_c_cycle_num: Optional[int] = None) -> float:
        """计算电压衰减率 - 严格按照原始脚本逻辑

        Args:
            cycle_df: 循环数据DataFrame
            file_info: 文件信息字典
            one_c_cycle_num: 1C首圈编号（1基索引）

        Returns:
            float: 电压衰减率（mV/周）
        """
        if len(cycle_df) <= 4:  # 原始脚本：只有当圈数>4时才计算
            return 0

        mode = file_info.get('mode', '')
        total_cycles = len(cycle_df)
        current_voltage = cycle_df.loc[total_cycles - 1, '放电中值电压(V)']

        # 0.1C模式 - 完全按照原始脚本第926行
        if mode == '-0.1C-':
            first_voltage = cycle_df.loc[0, '放电中值电压(V)']
            voltage_decay_rate = (first_voltage - current_voltage) * 1000 / (total_cycles - 1)
            return round(voltage_decay_rate, 1)

        # 1C模式 - 完全按照原始脚本第950行
        elif mode in self.one_c_modes:
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

    def _calculate_voltage_retention(self, cycle_df: pd.DataFrame, file_info: Dict[str, Any], one_c_cycle_num: Optional[int] = None) -> float:
        """计算电压保持率 - 严格按照原始脚本逻辑

        Args:
            cycle_df: 循环数据DataFrame
            file_info: 文件信息字典
            one_c_cycle_num: 1C首圈编号（1基索引）

        Returns:
            float: 电压保持率（百分比）
        """
        if len(cycle_df) <= 4:  # 原始脚本：只有当圈数>4时才计算
            return None

        mode = file_info.get('mode', '')
        total_cycles = len(cycle_df)
        current_voltage = cycle_df.loc[total_cycles - 1, '放电中值电压(V)']

        # 0.1C模式 - 完全按照原始脚本第927行
        if mode == '-0.1C-':
            first_voltage = cycle_df.loc[0, '放电中值电压(V)']
            voltage_retention = 100 * current_voltage / first_voltage
            return round(voltage_retention, 1)

        # 1C模式 - 完全按照原始脚本第951行
        elif mode in self.one_c_modes:
            # 确定1C首圈索引
            if one_c_cycle_num is not None and one_c_cycle_num > 0:
                one_c_idx = one_c_cycle_num - 1  # 转换为0基索引
            else:
                one_c_idx = min(3, total_cycles - 1)  # 默认第4圈或最大可用圈

            # 确保有足够的循环数据
            if total_cycles > one_c_idx + 1:
                one_c_voltage = cycle_df.loc[one_c_idx, '放电中值电压(V)']
                voltage_retention = 100 * current_voltage / one_c_voltage
                return round(voltage_retention, 1)
            else:
                return None

        # 其他模式，使用首圈作为基准
        else:
            first_voltage = cycle_df.loc[0, '放电中值电压(V)']
            voltage_retention = 100 * current_voltage / first_voltage
            return round(voltage_retention, 1)

    def _calculate_energy_retention(self, cycle_df: pd.DataFrame, file_info: Dict[str, Any], one_c_cycle_num: Optional[int] = None) -> float:
        """计算能量保持率 - 严格按照原始脚本逻辑

        Args:
            cycle_df: 循环数据DataFrame
            file_info: 文件信息字典
            one_c_cycle_num: 1C首圈编号（1基索引）

        Returns:
            float: 能量保持率（百分比）
        """
        if len(cycle_df) <= 4:  # 原始脚本：只有当圈数>4时才计算
            return None

        mode = file_info.get('mode', '')
        total_cycles = len(cycle_df)
        current_energy = cycle_df.loc[total_cycles - 1, '放电比能量(mWh/g)']

        # 0.1C模式 - 完全按照原始脚本第928行
        if mode == '-0.1C-':
            first_energy = cycle_df.loc[0, '放电比能量(mWh/g)']
            energy_retention = 100 * current_energy / first_energy
            return round(energy_retention, 1)

        # 1C模式 - 完全按照原始脚本第952行
        elif mode in self.one_c_modes:
            # 确定1C首圈索引
            if one_c_cycle_num is not None and one_c_cycle_num > 0:
                one_c_idx = one_c_cycle_num - 1  # 转换为0基索引
            else:
                one_c_idx = min(3, total_cycles - 1)  # 默认第4圈或最大可用圈

            # 确保有足够的循环数据
            if total_cycles > one_c_idx + 1:
                one_c_energy = cycle_df.loc[one_c_idx, '放电比能量(mWh/g)']
                energy_retention = 100 * current_energy / one_c_energy
                return round(energy_retention, 1)
            else:
                return None

        # 其他模式，使用首圈作为基准
        else:
            first_energy = cycle_df.loc[0, '放电比能量(mWh/g)']
            energy_retention = 100 * current_energy / first_energy
            return round(energy_retention, 1)

    def _calculate_retention_at_cycle(self, cycle_df: pd.DataFrame, target_cycle: int, file_info: Dict[str, Any], one_c_cycle_num: Optional[int] = None) -> List[Any]:
        """计算指定循环数的保持率 - 严格按照原始脚本逻辑

        Args:
            cycle_df: 循环数据DataFrame
            target_cycle: 目标循环数（100或200）
            file_info: 文件信息字典
            one_c_cycle_num: 1C首圈编号（1基索引）

        Returns:
            List[Any]: [容量保持率, 电压保持率, 能量保持率]
        """
        mode = file_info.get('mode', '')
        total_cycles = len(cycle_df)

        # 0.1C模式：直接使用第100圈或第200圈
        if mode == '-0.1C-':
            if total_cycles <= target_cycle:
                return ['', '', '']

            first_discharge = cycle_df.loc[0, '放电比容量(mAh/g)']
            first_voltage = cycle_df.loc[0, '放电中值电压(V)']
            first_energy = cycle_df.loc[0, '放电比能量(mWh/g)']

            target_discharge = cycle_df.loc[target_cycle - 1, '放电比容量(mAh/g)']
            target_voltage = cycle_df.loc[target_cycle - 1, '放电中值电压(V)']
            target_energy = cycle_df.loc[target_cycle - 1, '放电比能量(mWh/g)']

            capacity_retention = round(100 * target_discharge / first_discharge, 1) if first_discharge > 0 else 0
            voltage_retention = round(100 * target_voltage / first_voltage, 1) if first_voltage > 0 else 0
            energy_retention = round(100 * target_energy / first_energy, 1) if first_energy > 0 else 0

            return [capacity_retention, voltage_retention, energy_retention]

        # 1C模式：从1C首圈开始计算 - 完全按照原始脚本第955-966行
        elif mode in self.one_c_modes:
            # 确定1C首圈索引
            if one_c_cycle_num is not None and one_c_cycle_num > 0:
                one_c_idx = one_c_cycle_num - 1  # 转换为0基索引
            else:
                one_c_idx = min(3, total_cycles - 1)  # 默认第4圈或最大可用圈

            # 计算目标循环索引：1C首圈 + target_cycle
            target_idx = one_c_idx + target_cycle

            if total_cycles <= target_idx:
                return ['', '', '']

            # 使用1C首圈作为基准
            one_c_discharge = cycle_df.loc[one_c_idx, '放电比容量(mAh/g)']
            one_c_voltage = cycle_df.loc[one_c_idx, '放电中值电压(V)']
            one_c_energy = cycle_df.loc[one_c_idx, '放电比能量(mWh/g)']

            # 获取目标循环数据
            target_discharge = cycle_df.loc[target_idx, '放电比容量(mAh/g)']
            target_voltage = cycle_df.loc[target_idx, '放电中值电压(V)']
            target_energy = cycle_df.loc[target_idx, '放电比能量(mWh/g)']

            capacity_retention = round(100 * target_discharge / one_c_discharge, 1) if one_c_discharge > 0 else 0
            voltage_retention = round(100 * target_voltage / one_c_voltage, 1) if one_c_voltage > 0 else 0
            energy_retention = round(100 * target_energy / one_c_energy, 1) if one_c_energy > 0 else 0

            return [capacity_retention, voltage_retention, energy_retention]

        # 其他模式，使用首圈作为基准
        else:
            if total_cycles <= target_cycle:
                return ['', '', '']

            first_discharge = cycle_df.loc[0, '放电比容量(mAh/g)']
            first_voltage = cycle_df.loc[0, '放电中值电压(V)']
            first_energy = cycle_df.loc[0, '放电比能量(mWh/g)']

            target_discharge = cycle_df.loc[target_cycle - 1, '放电比容量(mAh/g)']
            target_voltage = cycle_df.loc[target_cycle - 1, '放电中值电压(V)']
            target_energy = cycle_df.loc[target_cycle - 1, '放电比能量(mWh/g)']

            capacity_retention = round(100 * target_discharge / first_discharge, 1) if first_discharge > 0 else 0
            voltage_retention = round(100 * target_voltage / first_voltage, 1) if first_voltage > 0 else 0
            energy_retention = round(100 * target_energy / first_energy, 1) if first_energy > 0 else 0

            return [capacity_retention, voltage_retention, energy_retention]

    def get_results(self) -> Dict[str, pd.DataFrame]:
        """获取所有处理结果 - 完全按照原始脚本逻辑"""
        return {
            'main_data': self.all_cycle_data,
            'first_cycle_data': self.all_first_cycle,
            'error_data': self.all_error_data,
            'statistics_data': self.statistics_data
        }

    def is_abnormal_first_cycle(self, cycle_df: pd.DataFrame) -> bool:
        """检查首圈数据是否异常
        
        Args:
            cycle_df: 循环数据DataFrame
            
        Returns:
            bool: 是否异常
        """
        if cycle_df.empty:
            return True
        
        first_cycle = cycle_df.iloc[0]
        
        # 获取首圈数据
        charge_capacity = first_cycle.get('充电比容量(mAh/g)', 0)
        discharge_capacity = first_cycle.get('放电比容量(mAh/g)', 0)
        
        # 检查是否超出异常阈值
        if (charge_capacity > self.config.abnormal_high_charge or
            charge_capacity < self.config.abnormal_low_charge):
            self.logger.log_debug(f"首充容量异常: {charge_capacity}")
            return True

        if discharge_capacity < self.config.abnormal_low_discharge:
            self.logger.log_debug(f"首放容量异常: {discharge_capacity}")
            return True
        
        return False
    
    def process_cycle_data(self, cycle_df: pd.DataFrame, file_info: Dict[str, Any], series_name: str) -> Optional[Dict[str, Any]]:
        """处理循环数据
        
        Args:
            cycle_df: 循环数据DataFrame
            file_info: 文件信息字典
            series_name: 系列名称
            
        Returns:
            Optional[Dict[str, Any]]: 处理后的数据记录
        """
        try:
            if cycle_df.empty:
                return None
            
            # 基础数据记录
            data = {
                '系列': series_name,
                '主机': file_info['device_id'],
                '通道': file_info['channel_id'],
                '批次': file_info['batch_id'],
                '上架时间': file_info['shelf_time'],
                '模式': file_info['mode'],
                '活性物质': file_info['mass']
            }
            
            # 处理首圈数据
            first_cycle = cycle_df.iloc[0]
            data.update({
                '首充': first_cycle.get('充电比容量(mAh/g)', 0),
                '首放': first_cycle.get('放电比容量(mAh/g)', 0),
                '首圈电压': first_cycle.get('放电中值电压(V)', 0),
                '首圈能量': first_cycle.get('放电比能量(mWh/g)', 0)
            })
            
            # 计算首效
            if data['首充'] > 0:
                data['首效'] = (data['首放'] / data['首充']) * 100
            else:
                data['首效'] = 0
            
            # 处理后续循环数据
            self._process_subsequent_cycles(cycle_df, data)
            
            # 1C识别和处理
            self._process_1c_identification(cycle_df, data, file_info['mode'])
            
            # 计算容量保持率
            self._calculate_capacity_retention(cycle_df, data)
            
            return data
            
        except Exception as e:
            self.logger.log_error(f"处理循环数据时发生错误: {str(e)}")
            return None
    
    def _process_subsequent_cycles(self, cycle_df: pd.DataFrame, data: Dict[str, Any]):
        """处理后续循环数据
        
        Args:
            cycle_df: 循环数据DataFrame
            data: 数据记录字典
        """
        # 处理Cycle2-7数据
        for cycle_num in range(2, 8):
            if len(cycle_df) >= cycle_num:
                cycle_data = cycle_df.iloc[cycle_num - 1]
                data[f'Cycle{cycle_num}'] = cycle_data.get('放电比容量(mAh/g)', 0)
                data[f'Cycle{cycle_num}充电比容量'] = cycle_data.get('充电比容量(mAh/g)', 0)
            else:
                data[f'Cycle{cycle_num}'] = 0
                data[f'Cycle{cycle_num}充电比容量'] = 0
        
        # 当前圈数
        data['当前圈数'] = len(cycle_df)
    
    def _process_1c_identification(self, cycle_df: pd.DataFrame, data: Dict[str, Any], mode: str):
        """1C识别和处理
        
        Args:
            cycle_df: 循环数据DataFrame
            data: 数据记录字典
            mode: 测试模式
        """
        # 初始化1C相关字段
        data.update({
            '1C首圈编号': None,
            '1C首充': 0,
            '1C首放': 0,
            '1C首效': 0,
            '1C状态': '未识别',
            '1C倍率比': 0
        })
        
        if self.config.verbose:
            self.logger.log_debug(f"开始1C识别，模式: {mode}")
        
        # 检查是否需要1C识别
        if mode in self.config.mode_one_c_modes:
            self._identify_1c_cycle(cycle_df, data)
        else:
            # 非1C模式，使用默认处理
            self._process_non_1c_mode(cycle_df, data)
    
    def _identify_1c_cycle(self, cycle_df: pd.DataFrame, data: Dict[str, Any]):
        """识别1C循环
        
        Args:
            cycle_df: 循环数据DataFrame
            data: 数据记录字典
        """
        found_1c = False
        
        # 从第2圈开始查找1C
        for i in range(1, min(len(cycle_df), 10)):  # 最多查找到第10圈
            current_cycle = cycle_df.iloc[i]
            current_discharge = current_cycle.get('放电比容量(mAh/g)', 0)
            
            if current_discharge > 0 and data['首放'] > 0:
                # 计算比值和差值
                discharge_ratio = current_discharge / data['首放']
                discharge_diff = abs(data['首放'] - current_discharge)
                
                # 检查是否满足1C条件
                if (discharge_ratio < self.config.ratio_threshold and
                    discharge_diff > self.config.discharge_diff_threshold):
                    
                    # 找到1C循环
                    data['1C首圈编号'] = i + 1
                    data['1C首充'] = current_cycle.get('充电比容量(mAh/g)', 0)
                    data['1C首放'] = current_discharge
                    data['1C倍率比'] = discharge_ratio
                    
                    # 计算1C首效
                    if data['1C首充'] > 0:
                        data['1C首效'] = (data['1C首放'] / data['1C首充']) * 100
                    
                    # 判断1C状态
                    if data['1C首充'] > self.config.overcharge_threshold:
                        data['1C状态'] = '过充'
                    elif data['1C首效'] < self.config.very_low_efficiency_threshold:
                        data['1C状态'] = '首效过低'
                    elif data['1C首效'] < self.config.low_efficiency_threshold:
                        data['1C状态'] = '首效偏低'
                    else:
                        data['1C状态'] = '正常'
                    
                    found_1c = True
                    break
        
        if not found_1c:
            # 使用默认1C圈数
            default_cycle = self.config.default_1c_cycle
            if len(cycle_df) >= default_cycle:
                default_cycle_data = cycle_df.iloc[default_cycle - 1]
                data['1C首圈编号'] = default_cycle
                data['1C首充'] = default_cycle_data.get('充电比容量(mAh/g)', 0)
                data['1C首放'] = default_cycle_data.get('放电比容量(mAh/g)', 0)

                if data['1C首充'] > 0:
                    data['1C首效'] = (data['1C首放'] / data['1C首充']) * 100

                if data['首放'] > 0:
                    data['1C倍率比'] = data['1C首放'] / data['首放']

                # 判断状态
                if data['1C首充'] > self.config.overcharge_threshold:
                    data['1C状态'] = '过充'
                elif data['1C首效'] < self.config.very_low_efficiency_threshold:
                    data['1C状态'] = '首效过低'
                elif data['1C首效'] < self.config.low_efficiency_threshold:
                    data['1C状态'] = '首效偏低'
                else:
                    data['1C状态'] = '正常'

                self.logger.log_debug(f"未找到满足条件的1C首圈，使用默认第{default_cycle}圈，状态为{data['1C状态']}")
    
    def _process_non_1c_mode(self, cycle_df: pd.DataFrame, data: Dict[str, Any]):
        """处理非1C模式
        
        Args:
            cycle_df: 循环数据DataFrame
            data: 数据记录字典
        """
        # 非1C模式不进行1C识别，保持默认值
        data['1C状态'] = '非1C模式'
    
    def _calculate_capacity_retention(self, cycle_df: pd.DataFrame, data: Dict[str, Any]):
        """计算容量保持率
        
        Args:
            cycle_df: 循环数据DataFrame
            data: 数据记录字典
        """
        if len(cycle_df) < 2:
            data.update({
                '当前容量保持': 100,
                '电压衰减率mV/周': 0,
                '当前电压保持': 100,
                '当前能量保持': 100,
                '100容量保持': 0,
                '100电压保持': 0,
                '100能量保持': 0,
                '200容量保持': 0,
                '200电压保持': 0,
                '200能量保持': 0
            })
            return
        
        # 计算当前容量保持率
        current_cycle = cycle_df.iloc[-1]
        current_discharge = current_cycle.get('放电比容量(mAh/g)', 0)
        
        if data['首放'] > 0:
            data['当前容量保持'] = (current_discharge / data['首放']) * 100
        else:
            data['当前容量保持'] = 0
        
        # 计算电压和能量保持率
        current_voltage = current_cycle.get('放电中值电压(V)', 0)
        current_energy = current_cycle.get('放电比能量(mWh/g)', 0)
        
        if data['首圈电压'] > 0:
            data['当前电压保持'] = (current_voltage / data['首圈电压']) * 100
        else:
            data['当前电压保持'] = 0
        
        if data['首圈能量'] > 0:
            data['当前能量保持'] = (current_energy / data['首圈能量']) * 100
        else:
            data['当前能量保持'] = 0
        
        # 计算电压衰减率 (简化计算)
        if len(cycle_df) > 1 and data['首圈电压'] > 0:
            voltage_decay = (data['首圈电压'] - current_voltage) / len(cycle_df) * 1000  # mV/周
            data['电压衰减率mV/周'] = voltage_decay
        else:
            data['电压衰减率mV/周'] = 0
        
        # 计算特定循环的保持率
        for target_cycle in [100, 200]:
            if len(cycle_df) >= target_cycle:
                target_data = cycle_df.iloc[target_cycle - 1]
                target_discharge = target_data.get('放电比容量(mAh/g)', 0)
                target_voltage = target_data.get('放电中值电压(V)', 0)
                target_energy = target_data.get('放电比能量(mWh/g)', 0)
                
                if data['首放'] > 0:
                    data[f'{target_cycle}容量保持'] = (target_discharge / data['首放']) * 100
                else:
                    data[f'{target_cycle}容量保持'] = 0
                
                if data['首圈电压'] > 0:
                    data[f'{target_cycle}电压保持'] = (target_voltage / data['首圈电压']) * 100
                else:
                    data[f'{target_cycle}电压保持'] = 0
                
                if data['首圈能量'] > 0:
                    data[f'{target_cycle}能量保持'] = (target_energy / data['首圈能量']) * 100
                else:
                    data[f'{target_cycle}能量保持'] = 0
            else:
                data[f'{target_cycle}容量保持'] = 0
                data[f'{target_cycle}电压保持'] = 0
                data[f'{target_cycle}能量保持'] = 0
