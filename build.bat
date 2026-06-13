@echo off
REM ============================================================
REM  Optional LOCAL build script (alternative to the GitHub
REM  cloud build). Run on a Windows PC that has Python installed.
REM  Staff who use the app do NOT need Python - just the .exe.
REM ============================================================
cd /d "%~dp0"

echo.
echo [1/3] Installing required packages...
py -m pip install --quiet --upgrade pandas openpyxl python-pptx pyinstaller
if errorlevel 1 (
    echo.
    echo ERROR: Could not install packages. Is Python installed?
    echo Download it from https://www.python.org/downloads/ and tick
    echo "Add python.exe to PATH" during installation, then run this again.
    pause
    exit /b 1
)

echo [2/3] Building QuoteGenerator.exe (this takes a few minutes)...
py -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name "QuoteGenerator" ^
    --add-data "Proposal  PERSONAL AUTO 23mar2026.pptx;." ^
    --add-data "HO3 Proposal 01july2026.pptx;." ^
    --add-data "ho3_mapping.json;." ^
    app.py
if errorlevel 1 (
    echo.
    echo ERROR: Build failed. See messages above.
    pause
    exit /b 1
)

echo [3/3] Done!
echo.
echo Your finished app is here:  dist\QuoteGenerator.exe
echo The templates are embedded inside it - the .exe is all staff need.
echo (To use an updated template later, just place the new .pptx file
echo  next to the .exe - it takes priority over the embedded one.)
echo.
pause
