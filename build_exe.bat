@echo off
chcp 65001 >nul
cd /d %~dp0
py -3.14 -m pip install -r requirements.txt pyinstaller
py -3.14 -m PyInstaller --noconfirm --clean --windowed --onefile --name BS_Calibration_Report_Generator run_gui.py
pause
