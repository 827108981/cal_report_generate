@echo off
chcp 65001 >nul
cd /d %~dp0
python -m pip install -r requirements.txt
python run_gui.py
pause
