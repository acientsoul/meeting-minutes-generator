@echo off
chcp 65001 > nul
cd /d "%~dp0"

start "" pythonw gui_stable.py
exit
