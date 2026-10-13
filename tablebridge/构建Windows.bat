@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist .venv py -3 -m venv .venv
.venv\Scripts\python -m pip install -r requirements-desktop.txt
if errorlevel 1 goto failed
.venv\Scripts\python -m PyInstaller --noconfirm --clean TableBridge.spec
if errorlevel 1 goto failed
echo Build completed: dist\TableBridge\TableBridge.exe
pause
exit /b 0
:failed
echo Build failed. See diagnostics above.
pause
exit /b 1
