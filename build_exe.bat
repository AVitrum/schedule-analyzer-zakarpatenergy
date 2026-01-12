@echo off
chcp 65001 > nul
echo ========================================
echo Створення EXE файлу...
echo ========================================
echo.

python -m PyInstaller --onefile --windowed --name "ScheduleAnalyzer" main.py

echo.
echo ========================================
echo Готово! EXE файл знаходиться в папці dist/
echo ========================================
pause
