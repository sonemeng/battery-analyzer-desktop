# LIMS数据处理程序 - 模块化版本 v2.0

## 📋 概述

这是LIMS数据处理程序的模块化重构版本，将原始的单文件脚本重构为多个独立模块，提高了代码的可维护性、可扩展性和可测试性。

## 🏗️ 架构设计

### 核心模块

```
modules/
├── config_parser.py      # 配置解析器 - 处理所有配置参数
├── file_parser.py        # 文件解析器 - 文件发现和信息提取
├── data_processor.py     # 数据处理器 - 核心数据处理逻辑
├── outlier_detection.py  # 异常检测器 - 箱线图和Z-score方法
└── __init__.py

utils/
├── logger.py             # 日志系统 - 多级日志记录
└── __init__.py

main_processor.py         # 主处理器 - 集成所有模块
main.py                   # 主程序入口 - 命令行和交互模式
run.py                    # 快速启动脚本
```

### 主要改进

1. **模块化设计**: 将单一脚本拆分为功能独立的模块
2. **配置系统**: 统一的配置管理，支持命令行参数和交互模式
3. **日志系统**: 多级日志记录，便于调试和问题追踪
4. **错误处理**: 健壮的异常处理和恢复机制
5. **进度跟踪**: 实时处理进度和状态反馈

## 🚀 使用方法

### 方法1: 双击启动（推荐）
```
双击 "启动程序.bat" 文件
```

### 方法2: 命令行启动
```bash
# 交互模式（推荐新手）
python main.py

# 命令行模式（推荐高级用户）
python main.py --input_folder "C:/Data" --outlier_method boxplot --verbose

# 快速启动
python run.py
```

### 方法3: 直接调用主处理器
```python
from modules.config_parser import ConfigParser
from main_processor import MainProcessor

# 创建配置
config_parser = ConfigParser()
config = config_parser.parse_arguments(['--input_folder', 'your_data_folder'])

# 运行处理
processor = MainProcessor(config)
success = processor.run()
```

## ⚙️ 配置参数

### 高频调节参数（用户经常需要调整）
- `outlier_method`: 异常检测方法 (boxplot/zscore_mad)
- `reference_channel_method`: 参考通道选择方法 (traditional/pca/retention_curve_mse)
- `very_low_efficiency_threshold`: 首效过低阈值 (默认: 80%)
- `low_efficiency_threshold`: 首效低阈值 (默认: 85%)
- `default_1c_cycle`: 默认1C首圈编号 (默认: 3)

### 中频调节参数（偶尔需要调整）
- `boxplot_threshold_discharge`: 箱线图首放极差阈值 (默认: 10)
- `zscore_threshold_discharge`: Z-score首放阈值 (默认: 3.0)
- `capacity_retention_capacity_weight`: 容量保留率权重 (默认: 0.6)
- `capacity_retention_voltage_weight`: 电压保留率权重 (默认: 0.1)
- `capacity_retention_energy_weight`: 能量保留率权重 (默认: 0.3)

### 低频调节参数（很少需要调整）
- `excel_engine`: Excel读取引擎 (默认: calamine)
- `output_format`: 输出格式 (默认: xlsx)
- `verbose`: 详细输出 (默认: False)

## 📊 输出文件

程序会在输入文件夹中生成以下文件：

1. **电池数据汇总表-HHMMSS.xlsx**: 主要结果文件
   - 主数据表: 处理后的所有有效数据
   - 统计数据表: 按系列和批次的统计信息
   - 异常数据表: 被标记为异常的数据记录

2. **处理日志-YYYYMMDD_HHMMSS/**: 日志文件夹
   - 主要处理日志.txt: 主要处理过程和结果
   - 异常检测详细日志.txt: 异常检测的详细信息
   - 调试详细日志.txt: 完整的调试信息

## 🔧 开发者信息

### 模块依赖关系
```
main.py
├── main_processor.py
│   ├── modules/config_parser.py
│   ├── modules/file_parser.py
│   ├── modules/data_processor.py
│   ├── modules/outlier_detection.py
│   └── utils/logger.py
└── 交互模式界面
```

### 扩展开发
- 添加新的异常检测方法: 修改 `outlier_detection.py`
- 添加新的参考通道选择方法: 修改 `data_processor.py`
- 添加新的配置参数: 修改 `config_parser.py`
- 添加新的输出格式: 创建新的导出模块

## 📝 版本历史

### v2.0 (模块化版本)
- 完全重构为模块化架构
- 统一配置管理系统
- 改进的日志和错误处理
- 支持命令行和交互模式
- 更好的代码组织和可维护性

### v1.0 (原始版本)
- 单文件脚本实现
- 基础功能完整
- 文件: `LIMS_DATA_PROCESS_改良箱线图版.py`

## 🆘 故障排除

### 常见问题
1. **导入错误**: 确保所有模块文件都在正确位置
2. **配置错误**: 检查配置参数是否正确设置
3. **文件读取失败**: 确认Excel文件格式和工作表名称
4. **内存不足**: 对于大数据集，考虑分批处理

### 获取帮助
```bash
python main.py --help  # 查看所有可用参数
```

## 📄 许可证

本程序为内部使用工具，请遵守相关使用规定。
