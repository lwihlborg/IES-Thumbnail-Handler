@echo off
:: IES Thumbnail Handler - Uninstallation Script
:: No Administrator required

echo ============================================
echo   IES Photometry Thumbnail Handler
echo   Uninstallation Script
echo ============================================
echo.

set "INSTALL_DIR=%LocalAppData%\IESThumbnailHandler"
set "CLSID={7B3FC2A1-E8D4-4F5B-9A2C-8E1D3F6B5C4A}"
set "THUMBPROV={e357fccd-a995-4576-b01f-234630154e96}"

echo Removing registry entries...

:: Remove CLSID registration
reg delete "HKCU\Software\Classes\CLSID\%CLSID%" /f >nul 2>&1

:: Remove thumbnail handler registrations
reg delete "HKCU\Software\Classes\.ies\ShellEx\%THUMBPROV%" /f >nul 2>&1
reg delete "HKCU\Software\Classes\.IES\ShellEx\%THUMBPROV%" /f >nul 2>&1

:: Try to remove empty ShellEx keys
reg delete "HKCU\Software\Classes\.ies\ShellEx" /f >nul 2>&1
reg delete "HKCU\Software\Classes\.IES\ShellEx" /f >nul 2>&1

:: Remove from approved list
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved" /v "%CLSID%" /f >nul 2>&1

echo Removing installation files...
if exist "%INSTALL_DIR%" (
    rmdir /s /q "%INSTALL_DIR%" >nul 2>&1
)

echo.
echo Restarting Windows Explorer...
taskkill /f /im explorer.exe >nul 2>&1
timeout /t 2 /nobreak >nul
start explorer.exe

echo.
echo ============================================
echo   Uninstallation Complete!
echo ============================================
echo.
echo The IES thumbnail handler has been removed.
echo.
pause
