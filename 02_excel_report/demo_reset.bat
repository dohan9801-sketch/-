@echo off
chcp 65001 > nul
cd /d "%~dp0"
python make_sample_data.py
pause
