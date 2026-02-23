@echo off
title Movie Rater - Setup and Build
color 0A

echo ============================================
echo   Movie Rater - Windows Setup Script
echo ============================================
echo.

REM Detect Python
SET PYTHON=
FOR %%P IN (python python3) DO (
    IF NOT DEFINED PYTHON (
        %%P --version >nul 2>&1 && SET PYTHON=%%P
    )
)
IF NOT DEFINED PYTHON (
    echo [ERROR] Python not found.
    echo Install Python 3.10+ from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during install.
    pause
    exit /b 1
)
FOR /F "tokens=*" %%i IN ('%PYTHON% --version') DO SET PYVER=%%i
echo [OK] %PYVER%
echo.

IF NOT EXIST "movie_rater.py" (
    echo [ERROR] movie_rater.py not found in this folder.
    echo Place setup_windows.bat and movie_rater.py in the same folder.
    pause
    exit /b 1
)
echo [OK] movie_rater.py found.
echo.

echo Preparing assets folder...
IF NOT EXIST "assets\" mkdir assets
SET LOGO_SRC=C:\Users\admin\Desktop\Movie Rater Project\logo
IF EXIST "%LOGO_SRC%\mrl_noBg.png" (
    copy /Y "%LOGO_SRC%\mrl_noBg.png" "assets\mrl_noBg.png" >nul
    echo [OK] mrl_noBg.png copied to assets\
) ELSE (
    echo [WARN] mrl_noBg.png not found - app launches without custom logo.
)
IF EXIST "%LOGO_SRC%\mrl_noBg.ico" (
    copy /Y "%LOGO_SRC%\mrl_noBg.ico" "assets\mrl_noBg.ico" >nul
    echo [OK] mrl_noBg.ico  copied to assets\
) ELSE (
    echo [WARN] mrl_noBg.ico not found
)
echo.

echo [1/4] Upgrading pip...
%PYTHON% -m pip install --upgrade pip --quiet
echo [OK] pip upgraded.
echo.

echo [2/4] Installing dependencies...
echo       customtkinter, pandas, openpyxl, Pillow, pyinstaller
echo.
%PYTHON% -m pip install customtkinter pandas openpyxl Pillow pyinstaller
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Package install failed. Check your internet connection.
    pause
    exit /b 1
)
echo.
echo [OK] All packages installed.
echo.

echo [3/4] Verifying imports...
%PYTHON% -c "import customtkinter, pandas, openpyxl, PIL, PyInstaller; print('[OK] All libraries verified.')"
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Import check failed.
    echo Run: %PYTHON% -m pip install customtkinter pandas openpyxl Pillow pyinstaller
    pause
    exit /b 1
)
echo.

echo [4/4] Building Movie Rater.exe ...
echo       This takes 2-5 minutes. Please wait.
echo.

SET ICO_FLAG=
IF EXIST "assets\mrl_noBg.ico" SET ICO_FLAG=--icon "assets\mrl_noBg.ico"

%PYTHON% -m PyInstaller --onefile --windowed --name "Movie Rater" --add-data "assets;assets" --collect-all customtkinter --hidden-import=openpyxl --hidden-import=openpyxl.styles --hidden-import=openpyxl.utils --hidden-import=openpyxl.workbook --hidden-import=pandas --hidden-import=pandas.io.formats.excel --hidden-import=PIL --hidden-import=PIL.Image %ICO_FLAG% movie_rater.py

IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Build failed. Try these fixes:
    echo   1. Disable antivirus temporarily, delete build\ and dist\ folders, retry.
    echo   2. Right-click this .bat and Run as Administrator.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   BUILD COMPLETE
echo ============================================
echo.
echo   Your app : dist\Movie Rater.exe
echo.
echo   Move it anywhere. No Python needed to run.
echo   First launch creates:
echo   Documents\Movie Rating DB.xlsx
echo ============================================
echo.
explorer dist
pause