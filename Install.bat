@echo off
setlocal EnableDelayedExpansion
title ExhiBytes Setup

echo.
echo  ==============================================
echo   ExhiBytes ^| Legal Exhibit Bundle Management
echo   Built by Tan Sze Yao
echo  ==============================================
echo.

:: ── Require Administrator ────────────────────────────────────────────────────
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo  ERROR: This installer must be run as Administrator.
    echo.
    echo  Please right-click Install.bat and choose
    echo  "Run as administrator", then try again.
    echo.
    pause
    exit /b 1
)

:: ── Check for LibreOffice ────────────────────────────────────────────────────
echo  Checking for LibreOffice...
set "LO_FOUND=0"

reg query "HKLM\SOFTWARE\LibreOffice" >nul 2>&1
if not errorlevel 1 set "LO_FOUND=1"

if "%LO_FOUND%"=="0" (
    reg query "HKLM\SOFTWARE\WOW6432Node\LibreOffice" >nul 2>&1
    if not errorlevel 1 set "LO_FOUND=1"
)

if "%LO_FOUND%"=="1" (
    echo  LibreOffice is already installed. Skipping.
) else (
    echo  LibreOffice was not detected.
    set "LO_MSI=%~dp0LibreOffice_26.2.2_Win_x86-64.msi"
    if exist "!LO_MSI!" (
        echo.
        echo  LibreOffice is required to convert Word, Excel, and PowerPoint
        echo  files to PDF inside ExhiBytes.
        echo.
        choice /C YN /M "  Install LibreOffice now? (Recommended)"
        if !errorlevel!==1 (
            echo.
            echo  Installing LibreOffice — please wait, this may take
            echo  several minutes...
            msiexec /i "!LO_MSI!" /qb! /norestart
            if !errorlevel! neq 0 (
                echo.
                echo  WARNING: LibreOffice installation encountered an issue.
                echo  You can install it manually later from libreoffice.org
            ) else (
                echo  LibreOffice installed successfully.
            )
        ) else (
            echo.
            echo  Skipping LibreOffice. Office document conversion will not
            echo  be available until LibreOffice is installed.
        )
    ) else (
        echo.
        echo  NOTE: LibreOffice installer not found in this folder.
        echo  Office document conversion will not work until you install
        echo  LibreOffice from libreoffice.org
    )
)

:: ── Install ExhiBytes ────────────────────────────────────────────────────────
echo.
echo  Installing ExhiBytes...

set "INSTALL_DIR=%ProgramFiles%\ExhiBytes"

if not exist "%INSTALL_DIR%" (
    mkdir "%INSTALL_DIR%"
)

xcopy /E /Y /I /Q "%~dp0app\*" "%INSTALL_DIR%\" >nul
if %errorlevel% neq 0 (
    echo.
    echo  ERROR: Failed to copy ExhiBytes files.
    echo  Make sure the "app" folder is present next to Install.bat.
    pause
    exit /b 1
)

:: ── Create Desktop Shortcut ──────────────────────────────────────────────────
echo  Creating shortcuts...

set "DESKTOP=%PUBLIC%\Desktop"
set "EXE=%INSTALL_DIR%\ExhiBytes.exe"
set "STARTMENU=%ProgramData%\Microsoft\Windows\Start Menu\Programs"

powershell -NoProfile -NonInteractive -Command ^
  "$ws = New-Object -ComObject WScript.Shell; " ^
  "$s = $ws.CreateShortcut('%DESKTOP%\ExhiBytes.lnk'); " ^
  "$s.TargetPath = '%EXE%'; " ^
  "$s.WorkingDirectory = '%INSTALL_DIR%'; " ^
  "$s.IconLocation = '%INSTALL_DIR%\resources\icon.ico,0'; " ^
  "$s.Description = 'ExhiBytes - Legal Exhibit Bundle Management'; " ^
  "$s.Save()" >nul 2>&1

powershell -NoProfile -NonInteractive -Command ^
  "$ws = New-Object -ComObject WScript.Shell; " ^
  "$s = $ws.CreateShortcut('%STARTMENU%\ExhiBytes.lnk'); " ^
  "$s.TargetPath = '%EXE%'; " ^
  "$s.WorkingDirectory = '%INSTALL_DIR%'; " ^
  "$s.IconLocation = '%INSTALL_DIR%\resources\icon.ico,0'; " ^
  "$s.Description = 'ExhiBytes - Legal Exhibit Bundle Management'; " ^
  "$s.Save()" >nul 2>&1

:: ── Clear ALL Windows icon caches and restart Explorer ──────────────────────
echo  Refreshing icon cache...
taskkill /f /im explorer.exe >nul 2>&1
timeout /t 1 /nobreak >nul
if exist "%LOCALAPPDATA%\IconCache.db" del /F /Q "%LOCALAPPDATA%\IconCache.db" >nul 2>&1
for /f "delims=" %%f in ('dir /b "%LOCALAPPDATA%\Microsoft\Windows\Explorer\iconcache*.db" 2^>nul') do (
    del /F /Q "%LOCALAPPDATA%\Microsoft\Windows\Explorer\%%f" >nul 2>&1
)
for /f "delims=" %%f in ('dir /b "%LOCALAPPDATA%\Microsoft\Windows\Explorer\thumbcache*.db" 2^>nul') do (
    del /F /Q "%LOCALAPPDATA%\Microsoft\Windows\Explorer\%%f" >nul 2>&1
)
start explorer.exe
timeout /t 2 /nobreak >nul

:: ── Done ─────────────────────────────────────────────────────────────────────
echo.
echo  ==============================================
echo   ExhiBytes installed successfully!
echo.
echo   Launch from the Desktop shortcut or
echo   Start Menu ^> ExhiBytes
echo  ==============================================
echo.
pause
