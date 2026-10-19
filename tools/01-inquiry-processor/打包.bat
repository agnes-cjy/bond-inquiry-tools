@echo off
chcp 65001 >nul
echo ============================================
echo   债券询价处理工具 v6 - 一键打包 EXE
echo ============================================
echo.

REM 1. 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.8+
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)
echo [OK] Python 已就绪

REM 2. 安装依赖
echo [安装] 正在安装依赖...
pip install -r requirements.txt -q
if errorlevel 1 (
    echo [错误] 依赖安装失败，请检查网络
    pause
    exit /b 1
)
echo [OK] 依赖安装完成

REM 3. PyInstaller 打包
echo [打包] 正在生成 EXE（可能需要 1-2 分钟）...
pyinstaller --onefile --windowed --name "债券询价处理工具_v6" --add-data "bond_inquiry_processor.py;." --hidden-import openpyxl.cell._writer bond_gui.py
if errorlevel 1 (
    echo [错误] 打包失败
    pause
    exit /b 1
)

echo.
echo ============================================
echo   打包成功！
echo   EXE 位置：dist\债券询价处理工具_v6.exe
echo ============================================
echo.
echo 将 dist 文件夹中的 EXE 拷贝到任意位置即可使用。
pause
