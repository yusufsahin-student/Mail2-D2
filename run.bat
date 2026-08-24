@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Ilk kurulum yapiliyor...
    py -m venv .venv
    if errorlevel 1 goto :error
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    if errorlevel 1 goto :error
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 goto :error
)

".venv\Scripts\python.exe" app.py
if errorlevel 1 goto :error
exit /b 0

:error
echo.
echo Uygulama baslatilamadi. Yukaridaki hata mesajini kontrol edin.
pause
exit /b 1

