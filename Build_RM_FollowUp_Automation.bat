@echo off
setlocal
cd /d "%~dp0"
pyinstaller --noconfirm --clean rm_tracker_app.spec
if errorlevel 1 (
    echo Build failed. Install PyInstaller with: python -m pip install pyinstaller
    pause
    exit /b 1
)
echo Build complete. The executable is in the dist folder.
pause
