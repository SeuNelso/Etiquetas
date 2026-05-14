@echo off
cd /d "%~dp0"

where python >nul 2>&1
if errorlevel 1 (
    echo Instale Python em https://www.python.org/downloads/ e marque "Add python.exe to PATH".
    pause
    exit /b 1
)

python -m pip install -r requirements.txt
python -m PyInstaller --onefile --windowed --name Etiquetas --hidden-import win32print --hidden-import openpyxl.cell._writer --collect-submodules openpyxl --clean --noconfirm etiquetas.py

echo.
echo Executavel: dist\Etiquetas.exe
pause
