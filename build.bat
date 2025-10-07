@echo off
REM Build standalone exe with PyInstaller (Windows)
pip install --upgrade pip
pip install pyinstaller
pyinstaller --onefile --windowed --name PhanTichDAO main.py
echo Done. Check dist\PhanTichDAO.exe
pause
