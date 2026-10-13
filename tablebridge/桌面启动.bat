@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist .venv py -3 -m venv .venv
.venv\Scripts\python -m pip install -r requirements-desktop.txt
if errorlevel 1 (
  pause
  exit /b 1
)
.venv\Scripts\python desktop.py
if errorlevel 1 pause
