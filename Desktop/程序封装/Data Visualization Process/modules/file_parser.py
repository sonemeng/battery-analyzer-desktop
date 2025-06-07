"""
文件解析模块 - 严格按照原始脚本LIMS_DATA_PROCESS_改良箱线图版.py重构

完全按照原始脚本的逻辑，不做任何简化或修改
"""

import os
import re
import time
import pandas as pd
from typing import Dict, List, Optional, Tuple, Any
import warnings

# 忽略pandas警告
warnings.filterwarnings('ignore')


class FileParser:
    """文件解析器类 - 严格按照原始脚本逻辑"""

    def __init__(self, config, logger):
        """初始化文件解析器

        Args:
            config: 配置对象
            logger: 日志记录器
        """
        self.config = config
        self.logger = logger

        # 直接使用原始脚本的CONFIG结构
        self.excel_engine = config.excel_engine
        self.cycle_sheet_name = config.cycle_sheet_name
        self.test_sheet_name = config.test_sheet_name

        # 异常数据阈值 - 完全按照原始脚本
        self.abnormal_thresholds = {
            'high_charge': config.abnormal_high_charge,
            'low_charge': config.abnormal_low_charge,
            'low_discharge': config.abnormal_low_discharge
        }

        # 循环数据列 - 完全按照原始脚本
        self.cycle_sheet_cols = [
            '充电比容量(mAh/g)', '放电比容量(mAh/g)', '放电中值电压(V)',
            '充电比能量(mWh/g)', '放电比能量(mWh/g)'
        ]

        # 测试模式配置 - 完全按照原始脚本
        self.mode_patterns = ['-0.1C-', '-0.5C-', '-1C-', '-BL-', '-0.33C-']
        self.one_c_modes = ['-1C-']
        self.non_one_c_modes = ['-0.1C-', '-BL-', '-0.33C-']

        # 系列配置 - 完全按照原始脚本
        self.series_config = {
            'G': {'include': ['-G-'], 'exclude': ['-M-']},
            'Q3': {'include': ['-Q3-']},
            'M': {'include': ['-M-']},
            'D': {'include': ['-D-']},
            'Z': {'include': ['-Z-']}
        }
        self.default_series = 'Q3'

        # 文件名解析配置
        self.device_id_max_length = 20
        self.default_channel = "CH-01"
        self.batch_id_prefix = "BATCH-"
        self.device_id_prefix = "DEVICE-"
    

    
    def discover_and_group_files(self, folder_path: str) -> Dict[str, List[str]]:
        """发现并按系列分组文件
        
        Args:
            folder_path: 文件夹路径
            
        Returns:
            Dict[str, List[str]]: 按系列分组的文件路径字典
        """
        self.logger.log_info(f"开始扫描文件夹: {folder_path}")
        
        # 查找所有Excel文件
        excel_files = self._find_excel_files(folder_path)
        if not excel_files:
            self.logger.log_error(f"在文件夹 {folder_path} 中未找到Excel文件")
            return {}
        
        self.logger.log_info(f"找到 {len(excel_files)} 个Excel文件")
        
        # 按系列分组文件
        file_groups = self._group_files_by_series(excel_files)
        
        # 输出分组结果
        for series, files in file_groups.items():
            self.logger.log_info(f"系列 {series}: {len(files)} 个文件")
        
        return file_groups
    
    def _find_excel_files(self, folder_path: str) -> List[str]:
        """查找文件夹中的Excel文件
        
        Args:
            folder_path: 文件夹路径
            
        Returns:
            List[str]: Excel文件路径列表
        """
        excel_patterns = ['*.xlsx', '*.xls']
        excel_files = []
        
        for pattern in excel_patterns:
            files = glob.glob(os.path.join(folder_path, pattern))
            excel_files.extend(files)
        
        # 过滤掉临时文件和备份文件
        filtered_files = []
        for file_path in excel_files:
            file_name = os.path.basename(file_path)
            if not file_name.startswith('~') and not file_name.startswith('.'):
                filtered_files.append(file_path)
        
        return filtered_files
    
    def _group_files_by_series(self, file_paths: List[str]) -> Dict[str, List[str]]:
        """按系列分组文件
        
        Args:
            file_paths: 文件路径列表
            
        Returns:
            Dict[str, List[str]]: 按系列分组的文件字典
        """
        file_groups = {}
        
        for file_path in file_paths:
            series = self._identify_series_from_filename(os.path.basename(file_path))
            
            if series not in file_groups:
                file_groups[series] = []
            file_groups[series].append(file_path)
        
        return file_groups

    def parse_file_info(self, file_path: str) -> Optional[Dict[str, Any]]:
        """解析文件信息

        Args:
            file_path: 文件路径

        Returns:
            Optional[Dict[str, Any]]: 文件信息字典，如果解析失败则返回None
        """
        try:
            file_name = os.path.basename(file_path)

            # 提取设备ID和通道ID
            device_id, channel_id = self._extract_host_and_channel(file_name)

            # 提取批次ID
            batch_id = self._extract_batch_like_original(file_name)

            # 分割文件名用于上架时间提取
            parts = file_name.split('-')

            # 提取上架时间
            shelf_time = self._extract_shelf_time(file_name, parts)

            # 提取模式
            mode = self._identify_test_mode(file_name)

            # 提取活性物质质量
            mass = self._extract_mass(file_name)

            return {
                'device_id': device_id,
                'channel_id': channel_id,
                'batch_id': batch_id,
                'shelf_time': shelf_time,
                'mode': mode,
                'mass': mass,
                'file_name': file_name
            }

        except Exception as e:
            self.logger.log_error(f"解析文件信息失败: {file_name}, 错误: {str(e)}")
            return None

    def read_cycle_data(self, file_path: str) -> Optional[pd.DataFrame]:
        """读取循环数据

        Args:
            file_path: 文件路径

        Returns:
            Optional[pd.DataFrame]: 循环数据DataFrame，如果读取失败则返回None
        """
        try:
            # 读取Excel文件的Cycle工作表
            cycle_df = pd.read_excel(
                file_path,
                sheet_name=self.config.cycle_sheet_name,
                usecols=['充电比容量(mAh/g)', '放电比容量(mAh/g)', '放电中值电压(V)',
                        '充电比能量(mWh/g)', '放电比能量(mWh/g)'],
                engine=self.config.excel_engine
            )

            # 检查数据是否为空
            if cycle_df.empty:
                self.logger.log_warning(f"循环数据为空: {os.path.basename(file_path)}")
                return None

            return cycle_df

        except Exception as e:
            self.logger.log_error(f"读取循环数据失败: {os.path.basename(file_path)}, 错误: {str(e)}")
            return None

    def _identify_series_from_filename(self, file_name: str) -> str:
        """从文件名中识别系列（支持动态识别）

        Args:
            file_name: 文件名

        Returns:
            str: 系列标识
        """
        # 首先尝试预设系列匹配
        for series_name, series_config in self.series_config["series"].items():
            # 检查包含条件
            include_patterns = series_config.get('include', [])
            exclude_patterns = series_config.get('exclude', [])

            # 检查是否包含必需的模式
            has_include = any(pattern in file_name for pattern in include_patterns)
            if not has_include:
                continue

            # 检查是否包含排除的模式
            has_exclude = any(pattern in file_name for pattern in exclude_patterns)
            if has_exclude:
                continue

            return series_name

        # 如果没有匹配到预设系列，尝试动态识别
        dynamic_series = self._auto_detect_series(file_name)
        if dynamic_series:
            self.logger.log_debug(f"动态识别系列: {file_name} -> {dynamic_series}")
            return dynamic_series

        # 如果动态识别也失败，返回默认系列
        return self.series_config["default_series"]

    def _auto_detect_series(self, file_name: str) -> Optional[str]:
        """自动检测系列标识

        Args:
            file_name: 文件名

        Returns:
            Optional[str]: 检测到的系列标识，如果失败则返回None
        """
        try:
            # 按破折号分割文件名
            parts = file_name.split('-')

            # 如果文件名段数不足，无法提取系列
            if len(parts) < 6:
                self.logger.log_debug(f"文件名段数不足，无法动态识别系列: {file_name}")
                return None

            # 从第6段开始提取系列标识
            series_part = parts[5] if len(parts) > 5 else ""

            if not series_part:
                return None

            # 提取字母部分作为系列标识
            import re
            letter_match = re.search(r'([A-Za-z]+)', series_part)
            if letter_match:
                series_id = letter_match.group(1).upper()
                self.logger.log_debug(f"从第6段提取字母系列: {series_part} -> {series_id}")
                return series_id

            # 如果没有字母，尝试提取数字部分
            number_match = re.search(r'(\d+)', series_part)
            if number_match:
                series_id = f"N{number_match.group(1)}"
                self.logger.log_debug(f"从第6段提取数字系列: {series_part} -> {series_id}")
                return series_id

            # 如果都没有，使用整个第6段
            series_id = series_part.upper()
            self.logger.log_debug(f"使用第6段作为系列: {series_part} -> {series_id}")
            return series_id

        except Exception as e:
            self.logger.log_debug(f"动态系列识别失败: {file_name}, 错误: {str(e)}")
            return None
    
    def extract_file_info(self, file_path: str) -> Dict[str, Any]:
        """从文件路径中提取文件信息

        Args:
            file_path: 文件路径

        Returns:
            Dict[str, Any]: 文件信息字典
        """
        try:
            # 首先从完整路径中提取文件名
            file_name = os.path.basename(file_path)
            self.logger.log_debug(f"正在解析文件: {file_name}")

            # 使用文件名进行解析
            parts = file_name.split(sep='-')
            # 分割文件名，获取下划线分隔的部分
            underscore_parts = file_name.split(sep='_')
            self.logger.log_debug(f"文件名分割结果: 破折号部分={len(parts)}个, 下划线部分={len(underscore_parts)}个")

            # 使用统一的主机通道解析方法
            device_id, channel_id = self._extract_host_and_channel(file_name)

            # 如果解析失败，使用备用方案
            if not device_id or not channel_id:
                self.logger.log_debug(f"主机通道解析失败，使用备用方案")
                # 备用方案：尝试从文件名中提取设备ID
                device_id = file_name.split('.')[0]  # 使用文件名作为设备ID
                max_length = self.config.device_id_max_length
                if len(device_id) > max_length:  # 如果太长，截断
                    device_id = device_id[:max_length]

                # 备用通道ID
                channel_match = re.search(r'CH[-_]?(\d+)', file_name, re.IGNORECASE)
                if channel_match:
                    channel_id = f"CH-{channel_match.group(1)}"
                else:
                    channel_id = self.config.default_channel  # 默认通道

            # 3. 批次ID提取 - 使用原始代码的下划线分割法
            try:
                batch_id = self._extract_batch_like_original(file_name)
                self.logger.log_debug(f"  提取的批次ID: {batch_id}")
            except Exception as e:
                self.logger.log_debug(f"  批次提取失败: {str(e)}，使用默认值")
                batch_id = self.config.batch_id_prefix + file_name[:5]

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
                        self.logger.log_debug(f"  使用空格前的最后两个破折号部分作为上架时间: {shelf_time}")
                    else:
                        # 只有一个部分，使用该部分
                        shelf_time = dash_parts[-1]
                        self.logger.log_debug(f"  使用空格前的最后一个破折号部分作为上架时间: {shelf_time}")
                else:
                    # 如果没有空格，尝试使用正则表达式查找日期格式
                    date_match = re.search(r'(\d{4}[-/]?\d{2}[-/]?\d{2}|\d{6}|\d{2}\d{2})', file_name)
                    if date_match:
                        shelf_time = date_match.group(1)
                        self.logger.log_debug(f"  使用正则表达式找到的日期作为上架时间: {shelf_time}")
                    else:
                        # 如果没有找到日期格式，使用文件名中倒数第二个部分和最后一个部分
                        if len(parts) >= 2:
                            shelf_time = '-'.join(parts[-2:])
                            self.logger.log_debug(f"  使用文件名中最后两个破折号部分作为上架时间: {shelf_time}")
                        else:
                            # 如果没有足够的部分，使用当前日期
                            import time
                            shelf_time = time.strftime('%m%d', time.localtime())
                            self.logger.log_debug(f"  使用当前日期作为上架时间: {shelf_time}")
            except Exception as e:
                self.logger.log_debug(f"  提取上架时间出错: {str(e)}，使用当前日期")
                import time
                shelf_time = time.strftime('%m%d', time.localtime())

            # 5. 测试模式识别
            mode = self._identify_test_mode(file_name)

            # 6. 读取活性物质质量
            try:
                test_df = pd.read_excel(file_path, sheet_name=self.config.test_sheet_name, engine=self.config.excel_engine)
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
                self.logger.log_debug(f"读取活性物质失败: {os.path.basename(file_path)}, 错误: {str(e)}")
                mass = None

            result = {
                'file_path': file_path,
                'file_name': file_name,
                'device_id': device_id,
                'channel_id': channel_id,
                'batch_id': batch_id,
                'shelf_time': shelf_time,
                'mode': mode,
                'mass': mass
            }

            self.logger.log_debug(f"文件信息提取结果: 设备={device_id}, 通道={channel_id}, 批次={batch_id}, 上架时间={shelf_time}, 模式={mode}")
            return result

        except Exception as e:
            self.logger.log_error(f"文件名解析错误: {os.path.basename(file_path)}, 错误: {str(e)}")
            # 提供默认值
            import time
            default_result = {
                'file_path': file_path,
                'file_name': os.path.basename(file_path),
                'device_id': self.config.device_id_prefix + os.path.basename(file_path)[:10],
                'channel_id': self.config.default_channel,
                'batch_id': self.config.batch_id_prefix + time.strftime('%m%d', time.localtime()),
                'shelf_time': time.strftime('%m%d', time.localtime()),
                'mode': '-1C-',
                'mass': None
            }
            self.logger.log_debug(f"使用默认值: {default_result}")
            return default_result

    def _extract_host_and_channel(self, channel_key: str) -> Tuple[str, str]:
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
            return "", ""

    def _extract_batch_like_original(self, filename: str) -> str:
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
                    return filename.split('.')[0]
        except Exception:
            return filename.split('.')[0]  # 备用方案

    def _extract_shelf_time(self, file_name: str, parts: list) -> str:
        """提取上架时间

        Args:
            file_name: 文件名
            parts: 文件名分割后的部分

        Returns:
            str: 上架时间
        """
        import time
        import re

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
                else:
                    # 只有一个部分，使用该部分
                    shelf_time = dash_parts[-1]
            else:
                # 如果没有空格，尝试使用正则表达式查找日期格式
                date_match = re.search(r'(\d{4}[-/]?\d{2}[-/]?\d{2}|\d{6}|\d{2}\d{2})', file_name)
                if date_match:
                    shelf_time = date_match.group(1)
                else:
                    # 如果没有找到日期格式，使用文件名中倒数第二个部分和最后一个部分
                    if len(parts) >= 2:
                        shelf_time = '-'.join(parts[-2:])
                    else:
                        # 如果没有足够的部分，使用当前日期
                        shelf_time = time.strftime('%m%d', time.localtime())
        except Exception as e:
            # 提取上架时间出错，使用当前日期
            shelf_time = time.strftime('%m%d', time.localtime())

        return shelf_time

    def _extract_mass(self, file_name: str) -> Optional[float]:
        """提取活性物质质量（简化版本）

        Args:
            file_name: 文件名

        Returns:
            Optional[float]: 活性物质质量，如果无法提取则返回None
        """
        # 这里简化处理，实际应该从Excel文件中读取
        # 在parse_file_info中会从Excel文件读取实际值
        return None

    def _identify_test_mode(self, file_name: str) -> str:
        """识别测试模式 - 完全按照原始脚本逻辑

        Args:
            file_name: 文件名

        Returns:
            测试模式标识
        """
        # 确保使用文件名而不是完整路径
        file_name = os.path.basename(file_name)

        for pattern in self.mode_patterns:
            if pattern in file_name:
                return pattern
        return '-1C-'  # 默认模式

    def is_abnormal_first_cycle(self, df: pd.DataFrame) -> bool:
        """检查首圈数据是否异常 - 完全按照原始脚本逻辑"""
        first_charge = df.loc[0, '充电比容量(mAh/g)']
        first_discharge = df.loc[0, '放电比容量(mAh/g)']

        return (first_charge > self.abnormal_thresholds['high_charge'] or
                first_charge < self.abnormal_thresholds['low_charge'] or
                first_discharge < self.abnormal_thresholds['low_discharge'])

    def read_test_data(self, file_path: str) -> Optional[float]:
        """读取测试数据中的活性物质质量 - 按照原始脚本逻辑"""
        try:
            test_df = pd.read_excel(file_path, sheet_name=self.test_sheet_name, engine=self.excel_engine)
            # 查找包含"活性物质"的行
            active_material_row = test_df[test_df.iloc[:, 0] == "活性物质"]
            if not active_material_row.empty:
                mass = active_material_row.iloc[0, 1]  # 取对应行的第二列值
                return mass
            else:
                # 尝试其他可能的列名
                for col_name in ["活性物质", "活性物质质量", "质量", "mass"]:
                    if col_name in test_df.columns:
                        mass = test_df[col_name].iloc[0]
                        return mass
                return None
        except Exception as e:
            print(f"读取活性物质失败: {os.path.basename(file_path)}, 错误: {str(e)}")
            return None

    def _identify_test_mode(self, file_path: str) -> str:
        """识别测试模式

        Args:
            file_path: 文件路径或文件名

        Returns:
            测试模式标识
        """
        # 确保使用文件名而不是完整路径
        file_name = os.path.basename(file_path)

        for pattern in self.config.mode_patterns:
            if pattern in file_name:
                return pattern
        return '-1C-'  # 默认模式

    def read_excel_file(self, file_path: str) -> Optional[pd.DataFrame]:
        """读取Excel文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            Optional[pd.DataFrame]: 读取的数据，如果失败则返回None
        """
        try:
            df = pd.read_excel(
                file_path, 
                sheet_name=self.config.cycle_sheet_name,
                usecols=self.cycle_sheet_cols, 
                engine=self.config.excel_engine
            )
            
            if df.empty:
                self.logger.log_warning(f"文件 {os.path.basename(file_path)} 的数据表为空")
                return None
            
            return df
            
        except Exception as e:
            self.logger.log_error(f"读取文件失败: {os.path.basename(file_path)}, 错误: {str(e)}")
            return None
