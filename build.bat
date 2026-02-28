@echo off
echo Building Vehicle Monitor Executable...

:: Ensure pyinstaller is available
py -m pip install pystray colorama pandas openpyxl pillow
py -m pip install pyinstaller

echo Running PyInstaller...
py -m PyInstaller --noconfirm --onedir --windowed --icon "app_icon.ico" "vehicle_monitor.py"

echo.
echo Build Complete!
echo You can find your executable inside the "dist/vehicle_monitor" folder.
pause
