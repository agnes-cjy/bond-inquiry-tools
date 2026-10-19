@echo off
chcp 65001 >nul
echo ============================================
echo   债券询价双表合并工具 - 一键打包
echo ============================================
echo.

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未找到 Python，请先安装 Python 3.10+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] 安装依赖...
pip install pyinstaller pandas openpyxl xlrd -q

echo [2/3] 开始打包（约1-2分钟）...
pyinstaller --onefile --windowed --name "债券询价双表合并工具" --add-data "bond_merge_processor.py;." --hidden-import openpyxl.cell._writer bond_merge_gui.py

if %errorlevel% neq 0 (
    echo [错误] 打包失败，请检查上方错误信息
    pause
    exit /b 1
)

echo [3/3] 打包完成！
echo.
echo exe 文件位置: dist\债券询价双表合并工具.exe
echo.
start explorer dist
pause
