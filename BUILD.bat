@echo off
echo Building CA_Update.exe...
echo.

REM Install PyInstaller if needed
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

echo.
echo Building... This may take a few minutes...
echo.

pyinstaller --name=CA_Update --onefile --windowed --add-data="CA_Tracking_System.py;." --hidden-import=pandas --hidden-import=openpyxl --hidden-import=tkinter --hidden-import=xlrd --collect-all=pandas --collect-all=openpyxl --clean CA_Tracking_GUI.py

if errorlevel 1 (
    echo.
    echo Build failed!
    pause
    exit /b 1
)

echo.
echo Done! Your .exe is in the 'dist' folder: dist\CA_Update.exe
echo.
pause

