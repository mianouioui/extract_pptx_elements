@echo off
REM Build standalone Windows executable
REM Prerequisites: pip install pyinstaller

cd /d "%~dp0.."

echo === Building Windows executable ===
if not exist build\specs mkdir build\specs
pyinstaller --onefile --name extract_pptx_elements --specpath build\specs extract_pptx_elements.py

echo.
echo Done! Binary at: dist\extract_pptx_elements.exe
echo Run: dist\extract_pptx_elements.exe --help
dir dist\extract_pptx_elements.exe
