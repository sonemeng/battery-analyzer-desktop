@echo off
chcp 65001 >nul
echo ========================================
echo LIMS数据处理程序启动器 v2.0
echo 模块化版本
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请先安装Python 3.8或更高版本
    pause
    exit /b 1
)

REM 显示Python版本
echo 检测到Python版本:
python --version

REM 检查依赖是否安装
echo.
echo 正在检查依赖...
python -c "import pandas, openpyxl, tqdm, matplotlib, seaborn, scikit-learn, scipy" >nul 2>&1
if errorlevel 1 (
    echo 警告: 部分依赖未安装，正在尝试自动安装...
    call 自动安装依赖.bat
    if errorlevel 1 (
        echo 依赖安装失败，请手动运行"自动安装依赖.bat"
        pause
        exit /b 1
    )
)

echo 依赖检查完成！
echo.

REM 启动主程序
echo 启动LIMS数据处理程序...
echo ========================================
python run.py

echo.
echo ========================================
echo 程序执行完成
echo ========================================
pause
