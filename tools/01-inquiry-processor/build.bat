@echo off
REM 手动分步构建脚本（用于排查打包问题）
chcp 65001 >nul

echo [1/3] 安装依赖
pip install pandas openpyxl xlrd pyinstaller

echo [2/3] 清理旧构建
if exist build rd /s /q build
if exist dist rd /s /q dist
del /q *.spec 2>nul

echo [3/3] PyInstaller 打包
pyinstaller --onefile --windowed --name "债券询价处理工具_v6" --add-data "bond_inquiry_processor.py;." --hidden-import openpyxl.cell._writer bond_gui.py

echo.
echo 完成。EXE 在 dist\ 目录下。
