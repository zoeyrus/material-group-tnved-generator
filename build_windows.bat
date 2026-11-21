@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 (
 echo Install Python 3.10 or newer from python.org with the Python launcher.
 pause
 exit /b 1
)
py -3 -m venv .venv
if errorlevel 1 goto fail
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto fail
.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --windowed --name TNVED_JSON --collect-all openpyxl --collect-all xlrd tnved_json.py
if errorlevel 1 goto fail
echo EXE ready: dist\TNVED_JSON.exe
explorer dist
pause
exit /b 0
:fail
echo Build failed. Check messages above.
pause
exit /b 1
