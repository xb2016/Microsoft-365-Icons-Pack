@echo off

:: -----------------------
:: Get Administrator Privileges
:: -----------------------
mode con cols=20 lines=3
color A
echo. UAC Request...
echo UACTEST > "%windir%\System32\Moedog.uac"
if exist "%windir%\System32\Moedog.uac" goto GOTADMIN
echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\GetAdmin.vbs"
echo UAC.ShellExecute "%~S0", "", "", "runas", 1 >> "%temp%\GetAdmin.vbs"
"%temp%\GetAdmin.vbs"
exit /b

:GOTADMIN
if exist "%temp%\GetAdmin.vbs" ( del "%temp%\GetAdmin.vbs" )
if exist "%windir%\System32\Moedog.uac" ( del "%windir%\System32\Moedog.uac" )
pushd "%~dp0"
cls
mode con cols=60 lines=24
title Microsoft 365 Icon Switcher - by Moedog
setlocal EnableDelayedExpansion

:: -----------------------
:: Directory Configuration
:: -----------------------
set "FILES=accicons.exe grv_icons.exe joticon.exe outicon.exe pj11icon.exe pptico.exe pubs.exe visicon.exe wordicon.exe xlicons.exe"
set "VERSIONS=O2003 O2007 O2010 O2013-2016 O2019-2024 O2025"
:: 64 bit os, 64 bit office
set "TARGET1=C:\Program Files\Microsoft Office\root\vfs\Windows\Installer\{90160000-000F-0000-1000-0000000FF1CE}"
:: 64 bit os, 32 bit office
set "TARGET2=C:\Program Files (x86)\Microsoft Office\root\vfs\Windows\Installer\{90160000-000F-0000-0000-0000000FF1CE}"
:: 32 bit os, 32 bit office
set "TARGET3=C:\Program Files\Microsoft Office\root\vfs\Windows\Installer\{90160000-000F-0000-0000-0000000FF1CE}"
set "BACKUP_BASE=C:\ProgramData\OfficeIconsBackup"

:: -----------------------
:: Read Office Version
:: -----------------------
set "OFFICE_VER=UNKNOWN"
for /f "tokens=2,*" %%A in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v VersionToReport 2^>nul') do (
    set "OFFICE_VER=%%B"
)
if "%OFFICE_VER%"=="UNKNOWN" (
    for /f "tokens=2,*" %%A in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v ClientVersionToReport 2^>nul') do (
        set "OFFICE_VER=%%B"
    )
)
if "%OFFICE_VER%"=="UNKNOWN" (
    for /f "tokens=2,*" %%A in ('reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration" /v VersionToReport 2^>nul') do (
        set "OFFICE_VER=%%B"
    )
)
if "%OFFICE_VER%"=="UNKNOWN" (
    for /f "tokens=2,*" %%A in ('reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration" /v ClientVersionToReport 2^>nul') do (
        set "OFFICE_VER=%%B"
    )
)
set "OFFICE_VER=%OFFICE_VER: =%"

:: -----------------------
:: Clean Non-Current Version Backups
:: -----------------------
for /f "delims=" %%B in ('dir /b /ad "%BACKUP_BASE%\backup_*" 2^>nul') do (
    if /i not "%%~nxB"=="backup_%OFFICE_VER%" (
        rd /s /q "%BACKUP_BASE%\%%~nxB" 2>nul
    )
)
set "BACKUP_EXISTS=0"
if exist "%BACKUP_BASE%\backup_%OFFICE_VER%\" set "BACKUP_EXISTS=1"

:: -----------------------
:: Check Target Directory
:: -----------------------
set "TARGET="
if exist "%TARGET3%\" set "TARGET=%TARGET3%"
if exist "%TARGET2%\" set "TARGET=%TARGET2%"
if exist "%TARGET1%\" set "TARGET=%TARGET1%"
if "%TARGET%"=="" (
    cls
	echo.
	echo. ==========================================================
	echo.
    echo. Could not detect any of the following directories.
    echo. Please check installation path and try again.
	echo.
	echo.   [1] !TARGET1!
	echo.
	echo.   [2] !TARGET2!
	echo.
	echo.   [3] !TARGET3!
	echo.
	echo. ==========================================================
	echo.
	echo. Press any key to exit... & pause >nul
    exit /b
)

:: -----------------------
:: Main Menu
:: -----------------------
:MENU
cls
echo.
echo. ==========================================================
echo.
echo. Office Version: %OFFICE_VER%    Backup Status: %BACKUP_EXISTS%
echo.
echo. Detected Target Directory: %TARGET%
echo.
echo. Please choose an option:
set /a idx=0
for %%V in (%VERSIONS%) do (
    set /a idx+=1
    set "opt!idx!=%%V"
    echo.   [!idx!] Use icons from %%V
)
set /a next=idx+1
if "%BACKUP_EXISTS%"=="1" (
    echo.   [%next%] Restore Backup ^(%BACKUP_BASE%^)
    set /a restoreOpt=%next%
    set /a next+=1
)
echo.   [%next%] Exit
set /a exitOpt=%next%
echo.
echo. ==========================================================
echo.
set /p "CHOICE=Enter the number and press Enter:"
if "%CHOICE%"=="" goto MENU
if "%BACKUP_EXISTS%"=="1" if "%CHOICE%"=="%restoreOpt%" goto RESTORE
if "%CHOICE%"=="%exitOpt%" exit /b
set "SELECTED_VERSION="
for /l %%N in (1,1,%idx%) do (
    if "%CHOICE%"=="%%N" set "SELECTED_VERSION=!opt%%N!"
)
if "%SELECTED_VERSION%"=="" goto MENU
if not exist "%cd%\%SELECTED_VERSION%\" (
    echo.
    echo. Source directory missing: %cd%\%SELECTED_VERSION%
    echo.
	echo. Press any key to exit... & pause >nul
    exit /b
)
cls
echo.
echo. ==========================================================
echo.

:: -----------------------
:: Backup
:: -----------------------
if "%BACKUP_EXISTS%"=="1" goto REPLACE
set "BACKUP_TARGET=%BACKUP_BASE%\backup_%OFFICE_VER%"
if not exist "%BACKUP_TARGET%\" md "%BACKUP_TARGET%"
echo. Backing up target files to: %BACKUP_TARGET%
echo.
for %%F in (%FILES%) do if exist "%TARGET%\%%F" copy /y "%TARGET%\%%F" "%BACKUP_TARGET%\%%F" >nul

:: -----------------------
:: Replace
:: -----------------------
:REPLACE
set "SRC=%cd%\%SELECTED_VERSION%"
echo. Starting replacement (Source: %SRC%)...
for %%F in (%FILES%) do (
    if exist "%SRC%\%%F" (
        echo.   - %%F
        copy /y "%SRC%\%%F" "%TARGET%\%%F" >nul 2>&1 || (
            echo.     Trying to take ownership and retry...
            takeown /f "%TARGET%\%%F" >nul 2>&1
            icacls "%TARGET%\%%F" /grant administrators:F >nul 2>&1
            copy /y "%SRC%\%%F" "%TARGET%\%%F" >nul 2>&1
        )
    ) else echo Source file missing: %SRC%\%%F
)
echo.
echo. Replacement completed, you may need to clear thumbnail
echo. cache and restart Explorer for changes to take effect.
echo.
echo. ==========================================================
echo.

:: -----------------------
:: Clean Thumbnail Cache
:: -----------------------
:CLEAN
echo. Press any key to execute cleanup... & pause >nul
taskkill /f /im explorer.exe >nul 2>&1
ping 127.0.0.1 -n 2 >nul
attrib -h -s -r "%localappdata%\IconCache.db" >nul 2>&1
attrib /s /d -h -s -r "%localappdata%\Microsoft\Windows\Explorer\*" >nul 2>&1
del /f /q "%localappdata%\IconCache.db" >nul 2>&1
del /f /q "%localappdata%\Microsoft\Windows\Explorer\thumbcache_*.db" >nul 2>&1
del /f /q "%localappdata%\Microsoft\Windows\Explorer\iconcache_*.db" >nul 2>&1
ping 127.0.0.1 -n 2 >nul
start explorer.exe
echo.
echo. Cleanup completed.
echo. The Start Menu application list icons may not change.
echo. Refer to https://moedog.org/1331.html for a solution.
echo.
echo. Press any key to exit... & pause >nul
exit /b

:: -----------------------
:: Restore Backup
:: -----------------------
:RESTORE
if exist "%BACKUP_BASE%\backup_%OFFICE_VER%\" (
    echo.
    echo. Restoring from backup_%OFFICE_VER%...
    for %%F in (%FILES%) do if exist "%BACKUP_BASE%\backup_%OFFICE_VER%\%%F" copy /y "%BACKUP_BASE%\backup_%OFFICE_VER%\%%F" "%TARGET%\%%F" >nul
    echo. Restore complete, you may need to clear thumbnail
    echo. cache and restart Explorer for changes to take effect.
    echo.
    rd /s /q "%BACKUP_BASE%" 2>nul
    goto CLEAN
) else (
    echo.
    echo. Backup file missing: %cd%\%SELECTED_VERSION%
    echo.
    echo. Press any key to return... & pause >nul
    goto MENU
)
