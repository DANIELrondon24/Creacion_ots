@echo off
cd /d "%~dp0"
echo Activando entorno virtual y ejecutando main.py...
if exist "venv\Scripts\python.exe" (
    venv\Scripts\python.exe main.py
) else (
    python main.py
)
pause