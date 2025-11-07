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
mode con cols=60 lines=24
title Microsoft 365 Icon Switcher - by 小白-白
setlocal EnableDelayedExpansion

:: -----------------------
:: 配置
:: -----------------------
:: 图标文件夹
set "VERSIONS=O2003 O2007 O2010 O2013-2016 O2019-2024 O2025"
:: 图标文件名
set "FILES=accicons.exe grv_icons.exe joticon.exe outicon.exe pj11icon.exe pptico.exe pubs.exe visicon.exe wordicon.exe xlicons.exe"
:: 备份/数据目录
set "BACKUP_BASE=C:\ProgramData\OfficeIconsBackup"
:: 开始菜单目录
set "START_MENU=C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
:: 开始菜单快捷方式名
set "LNK_NAME="Access" "Excel" "OneDrive for Business" "OneNote" "Outlook (classic)" "PowerPoint" "Project" "Publisher" "Skype for Business" "Sticky Notes (new)" "Visio" "Word""
:: 目标目录
set "TARGET1=C:\Program Files\Microsoft Office\root\vfs\Windows\Installer\{90160000-000F-0000-1000-0000000FF1CE}"
set "TARGET2=C:\Program Files (x86)\Microsoft Office\root\vfs\Windows\Installer\{90160000-000F-0000-0000-0000000FF1CE}"
set "TARGET3=C:\Program Files\Microsoft Office\root\vfs\Windows\Installer\{90160000-000F-0000-0000-0000000FF1CE}"

:: -----------------------
:: 读取 Office 版本号
:: -----------------------
cls
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
:: 读取 Windows 版本
:: -----------------------
set "WINDOWS_VER=UNKNOWN"
for /f "tokens=2,*" %%A in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName 2^>nul') do (
    set "WINDOWS_VER=%%B"
)
set "IS_WIN10=0"
echo %WINDOWS_VER% | findstr /i "Windows 10" >nul
if "%errorlevel%"=="0" set "IS_WIN10=1"

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
    goto EXIT
)

:: -----------------------
:: 主菜单
:: -----------------------
:MENU
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
if "%IS_WIN10%"=="1" (
    echo.   [%next%] 强制更改开始屏幕磁贴 ^(Windows 10^)
    set /a shortcutOpt=%next%
    set /a next+=1
)
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
if "%CHOICE%"=="%shortcutOpt%" goto SHORTCUT
set "SELECTED_VERSION="
for /l %%N in (1,1,%idx%) do (
    if "%CHOICE%"=="%%N" set "SELECTED_VERSION=!opt%%N!"
)
if "%SELECTED_VERSION%"=="" goto MENU
set "SRC=%cd%\Icons\%SELECTED_VERSION%"
if not exist "%SRC%\" (
    echo.
    echo. 源目录丢失：%SRC%
    goto EXIT
)

:: -----------------------
:: 备份
:: -----------------------
cls
echo.
echo. ==========================================================
if "%BACKUP_EXISTS%"=="1" goto REPLACE
set "BACKUP_TARGET=%BACKUP_BASE%\backup_%OFFICE_VER%"
if not exist "%BACKUP_TARGET%\" md "%BACKUP_TARGET%"
echo.
echo. 备份目标文件至：%BACKUP_TARGET%
for %%F in (%FILES%) do if exist "%TARGET%\%%F" copy /y "%TARGET%\%%F" "%BACKUP_TARGET%\%%F" >nul

:: -----------------------
:: 替换
:: -----------------------
:REPLACE
echo.
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

:: -----------------------
:: 清理缩略图缓存
:: -----------------------
:CLEAN
echo.
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
cls
echo.
echo. ==========================================================
echo.
echo. 清理完成。开始菜单应用列表的图标可能未更改
echo.
echo. 需要自行在快捷方式属性中重设图标来强制刷新
echo.
if not "%CHOICE%"=="%restoreOpt%" if "%IS_WIN10%"=="1" (
    echo. 对于 Windows 10，即使刷新也无法更改开始菜单磁贴图标
    echo. 
    echo. ==========================================================
    echo.
    echo.按任意键查看解决方案... & pause >nul
    goto SHORTCUT
)
echo. 详情请参考 https://moedog.org/1331.html
echo.
echo. ==========================================================
goto REFLASHSHORTCUT

:: -----------------------
:: 强制更改开始屏幕磁贴 (Windows 10)
:: -----------------------
:SHORTCUT
cls
echo.
echo. ==========================================================
echo.
echo. 对于 Windows 10，开始菜单应用列表 (左侧) 会遵循设置的快捷
echo. 方式图标
echo.
echo. 但是开始菜单磁贴 (右侧) 会优先读取原始可执行文件的图标
echo.
echo. 因此需要修改快捷方式目标，通过一个 VBS 中转调用来避免读取
echo. 原始图标
echo.
echo. 详情请参考 https://moedog.org/1331.html
echo.
echo. ==========================================================
echo.
echo.按任意键修改开始菜单快捷方式... & pause >nul
cls
echo.
echo. ==========================================================
echo.
echo. 开始修改快捷方式...
set "VBS_SRC=%cd%\Tools\LaunchOfficeApp.vbs"
set "VBS_TARGET=%BACKUP_BASE%\LaunchOfficeApp.vbs"
if not exist "%BACKUP_BASE%\" md "%BACKUP_BASE%"
if exist "%VBS_SRC%" copy /y "%VBS_SRC%" "%VBS_TARGET%" >nul
if not exist "%VBS_TARGET%" (
    echo. 文件复制失败：%VBS_SRC% → %VBS_SRC%
    goto EXIT
)
for %%A in (%LNK_NAME%) do (
    set "lnkFile=%START_MENU%\%%~A.lnk"
    if exist "!lnkFile!" (
        echo   - %%~A.lnk
        cscript //nologo "%cd%\Tools\EditShortcut.vbs" "!lnkFile!" "%VBS_TARGET%"
    )
)
echo.
echo. 修改完成
echo.
echo. ==========================================================
goto REFLASHSHORTCUT

:: -----------------------
:: 刷新开始菜单快捷方式
:: -----------------------
:REFLASHSHORTCUT
for %%A in (%LNK_NAME%) do (
    set "lnkFile=%START_MENU%\%%~A.lnk"
    if exist "!lnkFile!" (
        cscript //nologo "%cd%\Tools\ReflashShortcut.vbs" "!lnkFile!"
    )
)
goto EXIT

:: -----------------------
:: 退出程序
:: -----------------------
:EXIT
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
    rd /s /q "%BACKUP_BASE%\backup_%OFFICE_VER%" 2>nul
    goto CLEAN
) else (
    echo.
    echo. 备份文件丢失：%cd%\%SELECTED_VERSION%
    echo.
    echo.按任意键返回... & pause >nul
    goto MENU
)
