@echo off
:: IES Thumbnail Handler - Installation Script
:: Installs per-user, no Administrator required

setlocal EnableDelayedExpansion

echo ============================================
echo   IES Photometry Thumbnail Handler
echo   Installation Script
echo ============================================
echo.

set "SCRIPT_DIR=%~dp0"
set "INSTALL_DIR=%LocalAppData%\IESThumbnailHandler"
set "DLL_PATH=%INSTALL_DIR%\IESThumbnailHandler.dll"
set "SOURCE_FILE=%SCRIPT_DIR%IESThumbnailProvider.cs"
set "CLSID={7B3FC2A1-E8D4-4F5B-9A2C-8E1D3F6B5C4A}"
set "THUMBPROV={e357fccd-a995-4576-b01f-234630154e96}"

:: Find csc.exe (C# compiler included with .NET Framework)
set "CSC="
for /f "delims=" %%i in ('dir /b /ad /o-n "%SystemRoot%\Microsoft.NET\Framework64\v4*" 2^>nul') do (
    if exist "%SystemRoot%\Microsoft.NET\Framework64\%%i\csc.exe" (
        set "CSC=%SystemRoot%\Microsoft.NET\Framework64\%%i\csc.exe"
        goto :found_csc
    )
)
for /f "delims=" %%i in ('dir /b /ad /o-n "%SystemRoot%\Microsoft.NET\Framework\v4*" 2^>nul') do (
    if exist "%SystemRoot%\Microsoft.NET\Framework\%%i\csc.exe" (
        set "CSC=%SystemRoot%\Microsoft.NET\Framework\%%i\csc.exe"
        goto :found_csc
    )
)

echo ERROR: Could not find .NET Framework C# compiler (csc.exe)
echo Please ensure .NET Framework 4.x is installed.
pause
exit /b 1

:found_csc
echo Found compiler: %CSC%
echo.

:: Create installation directory
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

:: Compile the DLL
echo Compiling IESThumbnailHandler.dll...
"%CSC%" /target:library /out:"%DLL_PATH%" /reference:System.Drawing.dll "%SOURCE_FILE%"

if errorlevel 1 (
    echo ERROR: Compilation failed.
    pause
    exit /b 1
)

echo Compilation successful.
echo.

:: Register COM server via mscoree.dll (required for .NET managed assemblies)
echo Registering COM server...

set "CODEBASE_PATH=%DLL_PATH:\=/%"
reg add "HKCU\Software\Classes\CLSID\%CLSID%" /ve /d "IES Photometry Thumbnail Handler" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /ve /d "mscoree.dll" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /v "ThreadingModel" /d "Both" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /v "Assembly" /d "IESThumbnailHandler, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /v "Class" /d "IESThumbnailHandler.IESThumbnailProvider" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /v "RuntimeVersion" /d "v4.0.30319" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /v "CodeBase" /d "file:///%CODEBASE_PATH%" /f >nul

echo.
echo Registering file associations...

:: Register .ies extension
reg add "HKCU\Software\Classes\.ies" /ve /d "IESFile" /f >nul
reg add "HKCU\Software\Classes\.ies" /v "Content Type" /d "application/x-ies" /f >nul
reg add "HKCU\Software\Classes\.ies" /v "PerceivedType" /d "document" /f >nul

:: Register .IES (uppercase)
reg add "HKCU\Software\Classes\.IES" /ve /d "IESFile" /f >nul

:: Register ProgID
reg add "HKCU\Software\Classes\IESFile" /ve /d "IES Photometry File" /f >nul

:: Register thumbnail handler
reg add "HKCU\Software\Classes\.ies\ShellEx\%THUMBPROV%" /ve /d "%CLSID%" /f >nul
reg add "HKCU\Software\Classes\.IES\ShellEx\%THUMBPROV%" /ve /d "%CLSID%" /f >nul

:: Mark as approved (per-user)
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved" /v "%CLSID%" /d "IES Photometry Thumbnail Handler" /f >nul

echo.
echo Clearing thumbnail cache...
ie4uinit.exe -ClearIconCache >nul 2>&1
ie4uinit.exe -show >nul 2>&1

:: Kill and restart Explorer to apply changes
echo.
echo Restarting Windows Explorer...
taskkill /f /im explorer.exe >nul 2>&1
timeout /t 2 /nobreak >nul
start explorer.exe

echo.
echo ============================================
echo   Installation Complete!
echo ============================================
echo.
echo IES file thumbnails should now appear in Windows Explorer.
echo.
echo If thumbnails don't appear immediately:
echo   1. Open Explorer and navigate to View ^> Options
echo   2. Go to the View tab
echo   3. Make sure "Always show icons, never thumbnails" is UNCHECKED
echo   4. Clear thumbnail cache: run "cleanmgr" and select Thumbnails
echo.
pause
