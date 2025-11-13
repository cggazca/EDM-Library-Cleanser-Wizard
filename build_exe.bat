@echo off
echo ========================================
echo EDM Library Wizard - Executable Builder
echo ========================================
echo.

echo Installing required packages...
pip install -r requirements_wizard.txt

echo.
echo Building executable...
pyinstaller --onefile --windowed --name "EDM_Library_Wizard" edm_wizard.py

echo.
echo ========================================
echo Build Complete!
echo ========================================
echo.
echo The executable is located at: dist\EDM_Library_Wizard.exe
echo.
pause
