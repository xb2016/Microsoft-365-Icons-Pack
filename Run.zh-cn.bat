@echo off

:: -----------------------
:: 获取管理员权限
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
title Microsoft 365 Icon Switcher - by 小白-白
setlocal EnableDelayedExpansion

:: -----------------------
:: 目录配置
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
:: 读取 Office 版本号
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
:: 清理非当前版本备份
:: -----------------------
for /f "delims=" %%B in ('dir /b /ad "%BACKUP_BASE%\backup_*" 2^>nul') do (
    if /i not "%%~nxB"=="backup_%OFFICE_VER%" (
        rd /s /q "%BACKUP_BASE%\%%~nxB" 2>nul
    )
)
set "BACKUP_EXISTS=0"
if exist "%BACKUP_BASE%\backup_%OFFICE_VER%\" set "BACKUP_EXISTS=1"

:: -----------------------
:: 检测目标目录
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
    echo. 未检测到以下任一目录，请检查安装路径后重试。
	echo.
	echo.   [1] !TARGET1!
	echo.
	echo.   [2] !TARGET2!
	echo.
	echo.   [3] !TARGET3!
	echo.
	echo. ==========================================================
    echo.
	echo.按任意键退出程序... & pause >nul
    exit /b
)

:: -----------------------
:: 主菜单
:: -----------------------
:MENU
cls
echo.
echo. ==========================================================
echo.
echo. Office 版本：%OFFICE_VER%    备份状态：%BACKUP_EXISTS%
echo.
echo. 检测到目标目录：%TARGET%
echo.
echo. 请选择需要执行的操作：
set /a idx=0
for %%V in (%VERSIONS%) do (
    set /a idx+=1
    set "opt!idx!=%%V"
    echo.   [!idx!] 使用 %%V 的图标
)
set /a next=idx+1
if "%BACKUP_EXISTS%"=="1" (
    echo.   [%next%] 从备份还原 ^(%BACKUP_BASE%^)
    set /a restoreOpt=%next%
    set /a next+=1
)
echo.   [%next%] 退出
set /a exitOpt=%next%
echo.
echo. ==========================================================
echo.
set /p "CHOICE=请输入选项对应的数字并回车："
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
    echo. 源目录丢失：%cd%\%SELECTED_VERSION%
    echo.
	echo.按任意键退出程序... & pause >nul
    exit /b
)
cls
echo.
echo. ==========================================================
echo.

:: -----------------------
:: 备份
:: -----------------------
if "%BACKUP_EXISTS%"=="1" goto REPLACE
set "BACKUP_TARGET=%BACKUP_BASE%\backup_%OFFICE_VER%"
if not exist "%BACKUP_TARGET%\" md "%BACKUP_TARGET%"
echo. 备份目标文件至：%BACKUP_TARGET%
echo.
for %%F in (%FILES%) do if exist "%TARGET%\%%F" copy /y "%TARGET%\%%F" "%BACKUP_TARGET%\%%F" >nul

:: -----------------------
:: 替换
:: -----------------------
:REPLACE
set "SRC=%cd%\%SELECTED_VERSION%"
echo. 开始替换 (源：%SRC%)...
for %%F in (%FILES%) do (
    if exist "%SRC%\%%F" (
        echo.   - %%F
        copy /y "%SRC%\%%F" "%TARGET%\%%F" >nul 2>&1 || (
            echo.     尝试获取所有权并重试...
            takeown /f "%TARGET%\%%F" >nul 2>&1
            icacls "%TARGET%\%%F" /grant administrators:F >nul 2>&1
            copy /y "%SRC%\%%F" "%TARGET%\%%F" >nul 2>&1
        )
    ) else echo 源文件丢失：%SRC%\%%F
)
echo.
echo. 替换完成，需要清除缩略图缓存并重启资源管理器才能生效
echo.
echo. ==========================================================
echo.

:: -----------------------
:: 清理缩略图缓存
:: -----------------------
:CLEAN
echo.按任意键执行清理... & pause >nul
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
echo. 清理完成。开始菜单应用列表的图标可能未更改
echo. 请参考 https://moedog.org/1331.html 解决
echo.
echo.按任意键退出程序... & pause >nul
exit /b

:: -----------------------
:: 恢复备份
:: -----------------------
:RESTORE
if exist "%BACKUP_BASE%\backup_%OFFICE_VER%\" (
    echo.
    echo. 正在从 backup_%OFFICE_VER% 恢复...
    for %%F in (%FILES%) do if exist "%BACKUP_BASE%\backup_%OFFICE_VER%\%%F" copy /y "%BACKUP_BASE%\backup_%OFFICE_VER%\%%F" "%TARGET%\%%F" >nul
    echo. 恢复完成，需要清除缩略图缓存并重启资源管理器才能生效
    echo.
    rd /s /q "%BACKUP_BASE%" 2>nul
	goto CLEAN
) else (
    echo.
    echo. 备份文件丢失：%cd%\%SELECTED_VERSION%
    echo.
    echo.按任意键返回... & pause >nul
    goto MENU
)
