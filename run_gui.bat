@echo off
cd /d "%~dp0"
python excel_image_inspector_gui.py
if errorlevel 1 pause
