@echo off
chcp 65001 >nul
echo ========================================
echo 电池数据分析程序 - 自动安装依赖包
echo ========================================
echo.

echo 正在检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 错误：未找到Python环境
    echo 请先安装Python 3.7或更高版本
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

echo ✅ Python环境检查通过
echo.

echo 正在安装依赖包...
echo ----------------------------------------

echo 📦 安装numpy...
pip install numpy>=1.21.0

echo 📦 安装pandas...
pip install pandas>=1.3.0

echo 📦 安装matplotlib...
pip install matplotlib>=3.5.0

echo 📦 安装seaborn...
pip install seaborn>=0.11.0

echo 📦 安装openpyxl...
pip install openpyxl>=3.0.0

echo 📦 安装python-calamine...
pip install python-calamine>=0.4.0

echo 📦 安装scikit-learn...
pip install scikit-learn>=1.0.0

echo 📦 安装tqdm...
pip install tqdm>=4.62.0

echo 📦 安装xlsxwriter (可选)...
pip install xlsxwriter>=3.0.0

echo.
echo ========================================
echo ✅ 依赖包安装完成！
echo ========================================
echo.
echo 现在可以运行主程序了：
echo python LIMS_DATA_PROCESS_改良箱线图版.py
echo.
pause
