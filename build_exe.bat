@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d %~dp0

py -3.14 -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto :error

set "DATA_ARGS=--add-data BS5000;BS5000"
if exist "BS2800" set "DATA_ARGS=%DATA_ARGS% --add-data BS2800;BS2800"

py -3.14 -m PyInstaller ^
  --noconfirm --clean --windowed --onefile ^
  --name 化免校准报告自动生成工具 ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --hidden-import win32com ^
  --hidden-import win32com.client ^
  %DATA_ARGS% ^
  run_gui.py
if errorlevel 1 goto :error

echo.
echo 打包完成：dist\化免校准报告自动生成工具.exe
echo 默认输出目录为 EXE 同级 result 文件夹。
pause
exit /b 0

:error
echo.
echo 打包失败，请查看上方错误信息。
pause
exit /b 1
