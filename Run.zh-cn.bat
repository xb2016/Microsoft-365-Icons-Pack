@echo off

:: -----------------------
:: ��ȡ����ԱȨ��
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
title Microsoft 365 Icon Switcher - by С��-��
setlocal EnableDelayedExpansion

:: -----------------------
:: Ŀ¼����
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
:: ��ȡ Office �汾��
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
:: ����ǵ�ǰ�汾����
:: -----------------------
for /f "delims=" %%B in ('dir /b /ad "%BACKUP_BASE%\backup_*" 2^>nul') do (
    if /i not "%%~nxB"=="backup_%OFFICE_VER%" (
        rd /s /q "%BACKUP_BASE%\%%~nxB" 2>nul
    )
)
set "BACKUP_EXISTS=0"
if exist "%BACKUP_BASE%\backup_%OFFICE_VER%\" set "BACKUP_EXISTS=1"

:: -----------------------
:: ���Ŀ��Ŀ¼
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
    echo. δ��⵽������һĿ¼�����鰲װ·�������ԡ�
	echo.
	echo.   [1] !TARGET1!
	echo.
	echo.   [2] !TARGET2!
	echo.
	echo.   [3] !TARGET3!
	echo.
	echo. ==========================================================
    echo.
	echo.��������˳�����... & pause >nul
    exit /b
)

:: -----------------------
:: ���˵�
:: -----------------------
:MENU
cls
echo.
echo. ==========================================================
echo.
echo. Office �汾��%OFFICE_VER%    ����״̬��%BACKUP_EXISTS%
echo.
echo. ��⵽Ŀ��Ŀ¼��%TARGET%
echo.
echo. ��ѡ����Ҫִ�еĲ�����
set /a idx=0
for %%V in (%VERSIONS%) do (
    set /a idx+=1
    set "opt!idx!=%%V"
    echo.   [!idx!] ʹ�� %%V ��ͼ��
)
set /a next=idx+1
if "%BACKUP_EXISTS%"=="1" (
    echo.   [%next%] �ӱ��ݻ�ԭ ^(%BACKUP_BASE%^)
    set /a restoreOpt=%next%
    set /a next+=1
)
echo.   [%next%] �˳�
set /a exitOpt=%next%
echo.
echo. ==========================================================
echo.
set /p "CHOICE=������ѡ���Ӧ�����ֲ��س���"
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
    echo. ԴĿ¼��ʧ��%cd%\%SELECTED_VERSION%
    echo.
	echo.��������˳�����... & pause >nul
    exit /b
)
cls
echo.
echo. ==========================================================
echo.

:: -----------------------
:: ����
:: -----------------------
if "%BACKUP_EXISTS%"=="1" goto REPLACE
set "BACKUP_TARGET=%BACKUP_BASE%\backup_%OFFICE_VER%"
if not exist "%BACKUP_TARGET%\" md "%BACKUP_TARGET%"
echo. ����Ŀ���ļ�����%BACKUP_TARGET%
echo.
for %%F in (%FILES%) do if exist "%TARGET%\%%F" copy /y "%TARGET%\%%F" "%BACKUP_TARGET%\%%F" >nul

:: -----------------------
:: �滻
:: -----------------------
:REPLACE
set "SRC=%cd%\%SELECTED_VERSION%"
echo. ��ʼ�滻 (Դ��%SRC%)...
for %%F in (%FILES%) do (
    if exist "%SRC%\%%F" (
        echo.   - %%F
        copy /y "%SRC%\%%F" "%TARGET%\%%F" >nul 2>&1 || (
            echo.     ���Ի�ȡ����Ȩ������...
            takeown /f "%TARGET%\%%F" >nul 2>&1
            icacls "%TARGET%\%%F" /grant administrators:F >nul 2>&1
            copy /y "%SRC%\%%F" "%TARGET%\%%F" >nul 2>&1
        )
    ) else echo Դ�ļ���ʧ��%SRC%\%%F
)
echo.
echo. �滻��ɣ���Ҫ�������ͼ���沢������Դ������������Ч
echo.
echo. ==========================================================
echo.

:: -----------------------
:: ��������ͼ����
:: -----------------------
:CLEAN
echo.�������ִ������... & pause >nul
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
echo. ������ɡ���ʼ�˵�Ӧ���б��ͼ�����δ����
echo. ��ο� https://moedog.org/1331.html ���
echo.
echo.��������˳�����... & pause >nul
exit /b

:: -----------------------
:: �ָ�����
:: -----------------------
:RESTORE
if exist "%BACKUP_BASE%\backup_%OFFICE_VER%\" (
    echo.
    echo. ���ڴ� backup_%OFFICE_VER% �ָ�...
    for %%F in (%FILES%) do if exist "%BACKUP_BASE%\backup_%OFFICE_VER%\%%F" copy /y "%BACKUP_BASE%\backup_%OFFICE_VER%\%%F" "%TARGET%\%%F" >nul
    echo. �ָ���ɣ���Ҫ�������ͼ���沢������Դ������������Ч
    echo.
    rd /s /q "%BACKUP_BASE%" 2>nul
	goto CLEAN
) else (
    echo.
    echo. �����ļ���ʧ��%cd%\%SELECTED_VERSION%
    echo.
    echo.�����������... & pause >nul
    goto MENU
)
