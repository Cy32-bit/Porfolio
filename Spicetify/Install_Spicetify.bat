@echo off
setlocal EnableDelayedExpansion

:: Configuration
set "APP_NAME=SPICETIFY"
set "APP_DESC=Spotify Customization CLI"
set "SOURCE_URL=https://raw.githubusercontent.com/spicetify/cli/main/install.ps1"

:: Initialize
title %APP_NAME% Installer
cls
call :init_colors

:: Main
call :draw_header
call :draw_info
call :draw_warning
call :execute_command
call :draw_footer

pause > nul
exit /b

:: Subroutines
:init_colors
    for /F %%a in ('echo prompt $E ^| cmd') do set "ESC=%%a"
    set "RESET=%ESC%[0m"
    set "BOLD=%ESC%[1m"
    set "CYAN=%ESC%[36m"
    set "GREEN=%ESC%[32m"
    set "YELLOW=%ESC%[33m"
    set "RED=%ESC%[31m"
    set "WHITE=%ESC%[37m"
    set "GRAY=%ESC%[90m"
exit /b

:draw_header
    echo.
    echo %CYAN%    ////////////////////////////////////////////////////////////%RESET%
    echo %CYAN%   //%RESET%                                                        %CYAN%//%RESET%
    echo %CYAN%  //%RESET%   %BOLD%%WHITE%  SSSS  PPPP   III  CCC EEEEE TTTTT III FFFFF Y   Y  %RESET%%CYAN%//%RESET%
    echo %CYAN% //%RESET%    %BOLD%%WHITE% S     P   P   I  C    E       T    I  F      Y Y   %RESET%%CYAN%//%RESET%
    echo %CYAN%//%RESET%     %BOLD%%WHITE%  SSS  PPPP    I  C    EEE     T    I  FFF     Y    %RESET%%CYAN%//%RESET%
    echo %CYAN%\\%RESET%     %BOLD%%WHITE%     S P       I  C    E       T    I  F       Y    %RESET%%CYAN%\\%RESET%
    echo %CYAN% \\%RESET%    %BOLD%%WHITE% SSSS  P      III  CCC EEEEE   T   III F       Y    %RESET%%CYAN%\\%RESET%
    echo %CYAN%  \\%RESET%                                                       %CYAN%\\%RESET%
    echo %CYAN%   \\%RESET%              %GRAY%%APP_DESC%%RESET%                    %CYAN%\\%RESET%
    echo %CYAN%    \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\%RESET%
    echo.
exit /b

:draw_info
    echo %CYAN%  +----------------------------------------------------------+%RESET%
    echo %CYAN%  :%RESET% %WHITE%System%RESET%                                                     %CYAN%:%RESET%
    echo %CYAN%  :%RESET%   Platform    Windows                                      %CYAN%:%RESET%
    echo %CYAN%  :%RESET%   Shell       PowerShell                                   %CYAN%:%RESET%
    echo %CYAN%  :%RESET%   Source      github.com/spicetify/cli                     %CYAN%:%RESET%
    echo %CYAN%  +----------------------------------------------------------+%RESET%
    echo.
exit /b

:draw_warning
    echo %YELLOW%  +----------------------------------------------------------+%RESET%
    echo %YELLOW%  :%RESET% %WHITE%NOTICE%RESET%                                                    %YELLOW%:%RESET%
    echo %YELLOW%  :%RESET% This script downloads and executes code from the internet. %YELLOW%:%RESET%
    echo %YELLOW%  :%RESET% Ensure you trust the source before continuing.             %YELLOW%:%RESET%
    echo %YELLOW%  +----------------------------------------------------------+%RESET%
    echo.
    set /p "confirm=  Continue? [Y/N] "
    if /I not "%confirm%"=="Y" exit /b 1
    echo.
exit /b

:execute_command
    echo %GREEN%  [EXECUTING]%RESET% %GRAY%%SOURCE_URL%%RESET%
    echo.
    
    powershell -ExecutionPolicy Bypass -Command "iwr -useb %SOURCE_URL% | iex"
    
    if %ERRORLEVEL% EQU 0 (
        echo.
        echo %GREEN%  [SUCCESS]%RESET% Installation completed.
        echo.
        echo %WHITE%  Next steps:%RESET%
        echo    1. Restart Spotify if running
        echo    2. Run 'spicetify --help' for commands
        echo    3. Visit spicetify.app/docs for documentation
    ) else (
        echo.
        echo %RED%  [ERROR]%RESET% Installation failed. Code: %ERRORLEVEL%
        echo    Check your connection and try again.
    )
    echo.
exit /b

:draw_footer
    echo %CYAN%  ----------------------------------------------------------%RESET%
    echo %GRAY%  Press any key to exit...%RESET%
exit /b