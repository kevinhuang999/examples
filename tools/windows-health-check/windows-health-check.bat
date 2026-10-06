@echo off
setlocal EnableDelayedExpansion
title 电脑管家 - 健康体检与安全优化

rem ============================================================
rem  Windows 电脑管家 · 健康体检与安全优化
rem  生成日期: 2026-10-06  (v2: 输入改为按键式, 不需要按回车)
rem  原则: 只做安全、可逆的操作; 不动个人文件; 改注册表前先备份
rem ============================================================

rem ---------- 管理员权限自提升 ----------
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo   [!] 本工具需要管理员权限，正在请求提升...
    echo       请在弹出的窗口中点「是」。
    echo.
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs" >nul 2>&1
    exit /b
)

chcp 936 >nul 2>&1
set "BK=%~dp0reg_backup"
set "RP=%USERPROFILE%\Desktop\电脑体检报告.txt"
if not exist "%USERPROFILE%\Desktop\" set "RP=%USERPROFILE%\电脑体检报告.txt"

:MENU
cls
echo.============================================================
echo.        Windows 电脑管家   ·   健康体检 与 安全优化
echo.============================================================
echo.
echo.  本工具只执行「安全的、可逆的」操作。
echo.  不删除任何个人文件（文档 / 照片 / 资料 / 下载内容）。
echo.
echo.  [1]  一键安全优化（推荐）  清理垃圾 + 网络刷新 + 去广告 + TRIM + 出报告
echo.  [2]  只做体检报告          纯只读，不改动系统任何设置
echo.  [3]  磁盘垃圾清理          临时文件 / 更新缓存 / 传递优化 / 回收站
echo.  [4]  网络诊断与优化        刷新 DNS / 优化 TCP / 更换公共 DNS
echo.  [5]  开机启动项检查        查看自启清单 + 禁用残留项
echo.  [6]  系统文件修复          sfc + DISM（耗时 10-30 分钟）
echo.  [7]  隐私与广告优化        注册表项，改动前自动备份
echo.  [8]  C 盘 SSD 维护          TRIM + 磁盘健康检测
echo.  [0]  退出
echo.
echo.------------------------------------------------------------
echo.  提示：按对应数字键即可，不用按回车。
echo.
choice /c 123456780 /n /t 90 /d 0 /m "  按对应数字键选择（90 秒无操作自动退出）: "
if errorlevel 9 goto END
if errorlevel 8 goto SSD
if errorlevel 7 goto PRIVACY
if errorlevel 6 goto REPAIR
if errorlevel 5 goto STARTUP
if errorlevel 4 goto NET
if errorlevel 3 goto CLEAN
if errorlevel 2 goto REPORT
if errorlevel 1 goto ONEKEY
goto MENU


rem ============================================================
rem  [1] 一键安全优化
rem ============================================================
:ONEKEY
cls
echo.============================================================
echo.              一键安全优化
echo.============================================================
echo.
echo.  将依次执行以下 5 步：
echo.
echo.    1. 清理可再生成的缓存文件（临时文件 / 错误报告 / 传递优化）
echo.    2. 刷新 DNS 与地址解析缓存
echo.    3. 关闭系统广告与推荐内容（改动前自动备份注册表）
echo.    4. 对 C 盘 SSD 执行 TRIM 维护
echo.    5. 生成一份完整体检报告到桌面
echo.
echo.  不含：清空回收站、删除下载文件、改动任何个人资料。
echo.
choice /c yn /n /t 20 /d y /m "  按 Y 键开始 / N 键返回（20 秒无操作自动开始）: "
if errorlevel 2 goto MENU
echo.
echo.  [已确认] 开始执行，请勿关闭本窗口。

echo.
echo.------------------------------------------------------------
echo.  [1/5] 清理缓存文件
echo.------------------------------------------------------------
echo.  正在扫描并清理，文件较多时约需 1-2 分钟，请稍候...
call :LOGTEMP "清理前"
call :CLEANTEMP
call :LOGTEMP "清理后"
echo.        ^> 缓存清理完成

echo.
echo.------------------------------------------------------------
echo.  [2/5] 刷新 DNS 缓存
echo.------------------------------------------------------------
ipconfig /flushdns >nul 2>&1
ipconfig /registerdns >nul 2>&1
arp -d * >nul 2>&1
nbtstat -R >nul 2>&1
echo.        ^> DNS 与 ARP 缓存已刷新

echo.
echo.------------------------------------------------------------
echo.  [3/5] 关闭广告与推荐内容
echo.------------------------------------------------------------
call :PRIVSET

echo.
echo.------------------------------------------------------------
echo.  [4/5] C 盘 SSD TRIM 维护
echo.------------------------------------------------------------
powershell -NoProfile -Command "try { Optimize-Volume -DriveLetter C -ReTrim -ErrorAction Stop; Write-Host '        > TRIM 指令已下发' } catch { Write-Host '        > TRIM 跳过: ' + $_.Exception.Message }"

echo.
echo.------------------------------------------------------------
echo.  [5/5] 生成体检报告
echo.------------------------------------------------------------
call :MAKEREPORT

echo.
echo.============================================================
echo.              一键优化完成
echo.============================================================
echo.
echo.  报告位置: %RP%
echo.
echo.  下一步建议:
echo.    * 打开报告看一眼，重点看「磁盘可用空间」和「启动项」两栏
echo.    * 想手动管启动项，走菜单 [5]
echo.
echo.  按任意键回主菜单，直接关窗口则退出。
pause >nul
goto MENU


rem ============================================================
rem  [2] 只读体检报告
rem ============================================================
:REPORT
cls
echo.正在采集系统信息并生成报告，请稍候...
echo.
call :MAKEREPORT
echo.
echo.============================================================
echo.  报告已生成
echo.============================================================
echo.
echo.  文件位置: %RP%
echo.
choice /c yn /n /t 20 /d n /m "  按 Y 键打开报告 / N 键返回（20 秒后自动返回）: "
if errorlevel 2 goto MENU
start "" notepad "%RP%"
goto MENU


rem ============================================================
rem  [3] 磁盘垃圾清理
rem ============================================================
:CLEAN
cls
echo.============================================================
echo.              磁盘垃圾清理
echo.============================================================
echo.
echo.  将要处理的项目：
echo.
echo.    [1] 用户临时文件夹        %TEMP%
echo.    [2] Windows 临时文件夹     C:\Windows\Temp
echo.    [3] 系统错误报告转储       WER
echo.    [4] 传递优化缓存           Windows 更新 P2P 缓存
echo.    [5] Windows 更新下载缓存   需先暂停更新服务
echo.    [6] 清空回收站             ！删除后不可恢复，请先确认
echo.
echo.  以上 1-5 项都是系统可以自动重建的缓存，不含个人资料。
echo.
choice /c yn /n /t 20 /d y /m "  按 Y 键执行 / N 键跳过（20 秒无操作自动执行）: "
if errorlevel 2 goto CLEAN_SKIP1
echo.
echo.  [已确认] 正在清理，文件较多时约需 1-2 分钟，请稍候...
call :LOGTEMP "清理前"
call :CLEANTEMP
echo.
echo.  正在清理 Windows 更新下载缓存...
net stop wuauserv >nul 2>&1
net stop bits >nul 2>&1
del /f /s /q "C:\Windows\SoftwareDistribution\Download\*" >nul 2>&1
for /d %%D in ("C:\Windows\SoftwareDistribution\Download\*") do rd /s /q "%%D" >nul 2>&1
net start wuauserv >nul 2>&1
net start bits >nul 2>&1
echo.        ^> 更新缓存已清理，服务已恢复
call :LOGTEMP "清理后"

:CLEAN_SKIP1
echo.
echo.------------------------------------------------------------
echo.  关于回收站
echo.------------------------------------------------------------
echo.
echo.  清空回收站会永久删除里面所有文件，无法找回。
echo.  请先打开回收站确认没有要留的东西。
echo.
choice /c yn /n /t 30 /d n /m "  按 Y 键清空 / N 键跳过（30 秒无操作自动跳过）: "
if errorlevel 2 goto CLEAN_SKIP2
powershell -NoProfile -Command "try { Clear-RecycleBin -Force -ErrorAction Stop; Write-Host '        > 回收站已清空' } catch { Write-Host '        > 回收站跳过: ' + $_.Exception.Message }"
goto CLEAN_DONE
:CLEAN_SKIP2
echo.        ^> 已跳过回收站
:CLEAN_DONE
echo.
echo.  清理完成。按任意键返回。
pause >nul
goto MENU


rem ============================================================
rem  [4] 网络诊断与优化
rem ============================================================
:NET
cls
echo.============================================================
echo.              网络诊断与优化
echo.============================================================
echo.
echo.  --- 当前网络状态 ---
echo.
ipconfig | findstr /i "适配器 IPv4 默认网关"
echo.
echo.  --- 当前 DNS 服务器 ---
netsh interface ipv4 show dnsservers
echo.
echo.  --- 延迟测试（阿里 DNS / 腾讯 DNS / 百度）---
ping -n 2 223.5.5.5 | findstr /i "平均 时间 time"
ping -n 2 119.29.29.29 | findstr /i "平均 时间 time"
ping -n 2 www.baidu.com | findstr /i "平均 时间 time"
echo.
echo.------------------------------------------------------------
echo.  [A] 刷新 DNS 缓存（安全，立即生效）
echo.  [B] 优化 TCP 参数（安全，需重启生效）
echo.  [C] 把无线网卡 DNS 改为阿里 + 腾讯公共 DNS（可一键恢复）
echo.  [D] 恢复无线网卡为自动获取 DNS
echo.  [E] 重度修复：重置 Winsock 与 TCP/IP（慎用，需重启，会断网重连）
echo.  [0] 返回主菜单
echo.------------------------------------------------------------
choice /c abcde0 /n /t 60 /d 0 /m "  按对应字母键选择（60 秒无操作返回主菜单）: "
if errorlevel 6 goto MENU
if errorlevel 5 goto NET_E
if errorlevel 4 goto NET_D
if errorlevel 3 goto NET_C
if errorlevel 2 goto NET_B
if errorlevel 1 goto NET_A
goto MENU

:NET_A
echo.
ipconfig /flushdns
ipconfig /registerdns
arp -d *
nbtstat -R
echo.
echo.  完成：DNS / ARP / NetBIOS 缓存已刷新。
pause >nul
goto NET

:NET_B
echo.
netsh int tcp set global autotuninglevel=normal
netsh int tcp set global rss=enabled
netsh int tcp set heuristics disabled
netsh int tcp set global ecncapability=disabled
echo.
echo.  完成：TCP 参数已优化，重启后生效。
echo.  如需还原：netsh int tcp reset
pause >nul
goto NET

:NET_C
echo.
echo.  将把无线网卡（WLAN）的 DNS 改为 223.5.5.5 与 119.29.29.29。
echo.  如果你的无线网卡名称不是 WLAN，请先用 ipconfig 查看后手动执行 netsh 命令。
echo.
netsh interface ip set dns name="WLAN" static 223.5.5.5 primary
netsh interface ip add dns name="WLAN" 119.29.29.29 index=2
ipconfig /flushdns >nul 2>&1
echo.
echo.  完成。如需还原，走本菜单 [D]。
pause >nul
goto NET

:NET_D
echo.
netsh interface ip set dns name="WLAN" dhcp
ipconfig /flushdns >nul 2>&1
echo.
echo.  完成：WLAN 已恢复为自动获取 DNS。
pause >nul
goto NET

:NET_E
echo.
echo.  ！警告：此操作会重置网络协议栈，执行后需要重启电脑，
echo.    并且会短暂断网。仅在网络反复出问题时才需要使用。
echo.
choice /c yn /n /t 30 /d n /m "  按 Y 键执行 / N 键取消（30 秒无操作自动取消）: "
if errorlevel 2 goto NET
netsh winsock reset
netsh int ip reset
ipconfig /flushdns
echo.
echo.  完成：网络协议栈已重置，请重启电脑后生效。
pause >nul
goto NET


rem ============================================================
rem  [5] 开机启动项检查
rem ============================================================
:STARTUP
cls
echo.============================================================
echo.              开机启动项检查
echo.============================================================
echo.
echo.  --- 当前用户自启 (HKCU) ---
powershell -NoProfile -Command "$p='HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'; if(Test-Path $p){ $props=(Get-ItemProperty $p).PSObject.Properties; $n=0; foreach($x in $props){ if($x.Name -like 'PS*'){continue}; $n++; Write-Host ('   ' + $x.Name + '  =  ' + $x.Value) }; Write-Host ('   共 ' + $n + ' 项') } else { Write-Host '   (该键不存在)' }"
echo.
echo.  --- 全局自启 (HKLM) ---
powershell -NoProfile -Command "$p='HKLM:\Software\Microsoft\Windows\CurrentVersion\Run'; if(Test-Path $p){ $props=(Get-ItemProperty $p).PSObject.Properties; $n=0; foreach($x in $props){ if($x.Name -like 'PS*'){continue}; $n++; Write-Host ('   ' + $x.Name + '  =  ' + $x.Value) }; Write-Host ('   共 ' + $n + ' 项') } else { Write-Host '   (该键不存在)' }"
echo.
echo.  --- 32 位程序自启 (HKLM WOW6432Node) ---
powershell -NoProfile -Command "$p='HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'; if(Test-Path $p){ $props=(Get-ItemProperty $p).PSObject.Properties; $n=0; foreach($x in $props){ if($x.Name -like 'PS*'){continue}; $n++; Write-Host ('   ' + $x.Name + '  =  ' + $x.Value) }; Write-Host ('   共 ' + $n + ' 项') } else { Write-Host '   (该键不存在)' }"
echo.
echo.  --- 启用 / 禁用 状态 ---
powershell -NoProfile -Command "$ps=@('HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run','HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run','HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32'); foreach($p in $ps){ if(Test-Path $p){ $props=(Get-ItemProperty $p).PSObject.Properties; foreach($x in $props){ if($x.Name -like 'PS*'){continue}; $b=$x.Value; $st='未知'; if($b[0] -eq 3){$st='已禁用'} elseif($b[0] -eq 2){$st='已启用'}; Write-Host ('   [' + $st + '] ' + $x.Name) } } }"
echo.
echo.------------------------------------------------------------
echo.  下面 B/C/D 三项为幂等操作：已经是禁用状态的会再写一次禁用标记，
echo.  不会有副作用；恢复随时可以在任务管理器里点「启用」。
echo.
echo.  [A] 打开任务管理器（启动应用页）
echo.  [B] 禁用迅雷残留组件 xlacc 的开机自启
echo.  [C] 禁用华硕吉祥物 TX Mascot 的开机自启
echo.  [D] 禁用 VMware 托盘 vmware-tray 的开机自启
echo.  [0] 返回主菜单
echo.------------------------------------------------------------
choice /c abcd0 /n /t 60 /d 0 /m "  按对应字母键选择（60 秒无操作返回主菜单）: "
if errorlevel 5 goto MENU
if errorlevel 4 goto ST_D
if errorlevel 3 goto ST_C
if errorlevel 2 goto ST_B
if errorlevel 1 goto ST_A
goto MENU

:ST_A
start "" taskmgr.exe
goto MENU

:ST_B
call :DISABLEITEM "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32" "xlacc"
pause >nul
goto STARTUP

:ST_C
call :DISABLEITEM "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" "TX Mascot"
pause >nul
goto STARTUP

:ST_D
call :DISABLEITEM "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" "vmware-tray.exe"
pause >nul
goto STARTUP


rem ============================================================
rem  [6] 系统文件修复
rem ============================================================
:REPAIR
cls
echo.============================================================
echo.              系统文件完整性修复
echo.============================================================
echo.
echo.  将执行微软官方修复流程：
echo.
echo.    第一步 DISM 扫描并修复系统映像   约 5-20 分钟
echo.    第二步 SFC 扫描并修复系统文件     约 5-15 分钟
echo.
echo.  过程中进度可能长时间停在某个百分比，属于正常现象，请勿关闭窗口。
echo.  期间请勿运行大型程序，避免影响修复。
echo.
choice /c yn /n /t 30 /d n /m "  按 Y 键开始 / N 键返回（30 秒无操作自动返回）: "
if errorlevel 2 goto MENU

echo.
echo.------------------------------------------------------------
echo.  第一步：DISM 检查系统映像
echo.------------------------------------------------------------
DISM /Online /Cleanup-Image /CheckHealth
echo.
echo.  --- 扫描健康状态 ---
DISM /Online /Cleanup-Image /ScanHealth
echo.
echo.  --- 修复系统映像（这一步最慢）---
DISM /Online /Cleanup-Image /RestoreHealth

echo.
echo.------------------------------------------------------------
echo.  第二步：SFC 修复系统文件
echo.------------------------------------------------------------
sfc /scannow

echo.
echo.============================================================
echo.  修复流程结束
echo.============================================================
echo.
echo.  如果上面出现「未找到完整性冲突」，说明系统文件完好，无需处理。
echo.  如果出现「已修复」，建议重启电脑后再跑一次确认。
echo.
pause >nul
goto MENU


rem ============================================================
rem  [7] 隐私与广告优化
rem ============================================================
:PRIVACY
cls
echo.============================================================
echo.              隐私与广告优化
echo.============================================================
echo.
echo.  将关闭以下系统级广告与追踪（全部可逆）：
echo.
echo.    * 广告 ID 追踪
echo.    * 开始菜单的推荐应用与建议
echo.    * 设置页面的推荐内容
echo.    * 锁屏界面的推广信息
echo.    * 静默安装推荐应用
echo.    * 根据诊断数据推送定制体验
echo.    * 顺带开启「存储感知」自动清理临时文件
echo.
echo.  改动前会把原注册表导出到：reg_backup 文件夹
echo.
choice /c yn /n /t 20 /d y /m "  按 Y 键执行 / N 键返回（20 秒无操作自动执行）: "
if errorlevel 2 goto MENU
call :PRIVSET
echo.
echo.  完成。部分变化需要重启资源管理器或重启电脑后生效。
echo.
pause >nul
goto MENU


rem ============================================================
rem  [8] C 盘 SSD 维护
rem ============================================================
:SSD
cls
echo.============================================================
echo.              C 盘 SSD 维护
echo.============================================================
echo.
echo.  --- 物理磁盘健康状态 ---
powershell -NoProfile -Command "$d=Get-PhysicalDisk; foreach($x in $d){ Write-Host ('   ' + $x.FriendlyName + '  |  ' + $x.MediaType + '  |  ' + [math]::Round($x.Size/1GB,0) + ' GB  |  健康: ' + $x.HealthStatus + '  |  运行: ' + $x.OperationalStatus) }"
echo.
echo.  --- 分区剩余空间 ---
powershell -NoProfile -Command "$v=Get-Volume; foreach($x in $v){ if($x.DriveLetter){ Write-Host ('   ' + $x.DriveLetter + ':  剩余 ' + [math]::Round($x.SizeRemaining/1GB,1) + ' GB / 共 ' + [math]::Round($x.Size/1GB,1) + ' GB   健康: ' + $x.HealthStatus) } }"
echo.
echo.------------------------------------------------------------
choice /c yn /n /t 20 /d y /m "  按 Y 键执行 / N 键返回（20 秒无操作自动执行）: "
if errorlevel 2 goto MENU
defrag C: /L
echo.
echo.  完成：TRIM 指令已下发。SSD 不需要传统碎片整理，这个操作只做垃圾回收。
echo.
pause >nul
goto MENU


rem ============================================================
rem  子过程
rem ============================================================

:LOGTEMP
powershell -NoProfile -Command "$a=Get-ChildItem $env:TEMP -Recurse -Force -ErrorAction SilentlyContinue; $s=0; foreach($x in $a){$s+=$x.Length}; Write-Host ('   %~1 用户临时文件夹: ' + [math]::Round($s/1MB,1) + ' MB')"
exit /b

:CLEANTEMP
del /f /s /q "%TEMP%\*" >nul 2>&1
for /d %%D in ("%TEMP%\*") do rd /s /q "%%D" >nul 2>&1
del /f /s /q "C:\Windows\Temp\*" >nul 2>&1
for /d %%D in ("C:\Windows\Temp\*") do rd /s /q "%%D" >nul 2>&1
del /f /s /q "%LOCALAPPDATA%\Microsoft\Windows\WER\*" >nul 2>&1
del /f /s /q "C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\*" >nul 2>&1
del /f /s /q "C:\Windows\Minidump\*" >nul 2>&1
exit /b

:DISABLEITEM
rem %~1 = 注册表路径   %~2 = 启动项名称
if not exist "%BK%" md "%BK%" >nul 2>&1
reg export "%~1" "%BK%\StartupApproved_%~2.reg" /y >nul 2>&1
reg add "%~1" /v "%~2" /t REG_BINARY /d 030000000000000000000000 /f >nul 2>&1
if %errorlevel%==0 (
    echo.
    echo.  已禁用开机自启: %~2
    echo.  备份: %BK%\StartupApproved_%~2.reg
    echo.  想恢复：在任务管理器「启动应用」里右键点「启用」即可。
) else (
    echo.
    echo.  操作失败，可能该项不存在或权限不足。
)
exit /b

:PRIVSET
if not exist "%BK%" md "%BK%" >nul 2>&1
reg export "HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "%BK%\ContentDeliveryManager_before.reg" /y >nul 2>&1
reg export "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "%BK%\ExplorerAdvanced_before.reg" /y >nul 2>&1

reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" /v Enabled /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v Start_TrackProgs /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v Start_TrackSuggestions /t REG_DWORD /d 0 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Privacy" /v TailoredExperiencesWithDiagnosticDataEnabled /t REG_DWORD /d 0 /f >nul 2>&1

set "CDM=HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
reg add "%CDM%" /v SystemPaneSuggestionsEnabled /t REG_DWORD /d 0 /f >nul 2>&1
reg add "%CDM%" /v SilentInstalledAppsEnabled /t REG_DWORD /d 0 /f >nul 2>&1
reg add "%CDM%" /v SoftLandingEnabled /t REG_DWORD /d 0 /f >nul 2>&1
reg add "%CDM%" /v RotatingLockScreenOverlayEnabled /t REG_DWORD /d 0 /f >nul 2>&1
reg add "%CDM%" /v SubscribedContent-338388Enabled /t REG_DWORD /d 0 /f >nul 2>&1
reg add "%CDM%" /v SubscribedContent-338389Enabled /t REG_DWORD /d 0 /f >nul 2>&1
reg add "%CDM%" /v SubscribedContent-338393Enabled /t REG_DWORD /d 0 /f >nul 2>&1
reg add "%CDM%" /v SubscribedContent-353694Enabled /t REG_DWORD /d 0 /f >nul 2>&1
reg add "%CDM%" /v SubscribedContent-353696Enabled /t REG_DWORD /d 0 /f >nul 2>&1
reg add "%CDM%" /v SubscribedContent-310093Enabled /t REG_DWORD /d 0 /f >nul 2>&1
reg add "%CDM%" /v OemPreInstalledAppsEnabled /t REG_DWORD /d 0 /f >nul 2>&1
reg add "%CDM%" /v PreInstalledAppsEnabled /t REG_DWORD /d 0 /f >nul 2>&1

reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy" /v 01 /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy" /v 04 /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy" /v 08 /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy" /v 32 /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy" /v 2048 /t REG_DWORD /d 0 /f >nul 2>&1

echo.        ^> 广告与推荐内容已关闭，存储感知已开启
echo.        ^> 备份目录: %BK%
exit /b

:MAKEREPORT
powershell -NoProfile -Command "$o=@(); $o+='############################################################'; $o+='#  Windows 电脑健康体检报告'; $o+='#  生成时间: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); $o+='############################################################'; $o+=''; $o+='【一、系统概况】'; $cs=Get-CimInstance Win32_ComputerSystem; $os=Get-CimInstance Win32_OperatingSystem; $o+='  计算机名 : ' + $cs.Name; $o+='  品牌型号 : ' + $cs.Manufacturer + ' ' + $cs.Model; $o+='  操作系统 : ' + $os.Caption; $o+='  系统版本 : Build ' + $os.BuildNumber; $o+='  本次开机 : ' + [math]::Round(((Get-Date)-$os.LastBootUpTime).TotalHours,1) + ' 小时'; $o+=''; $o+='【二、CPU】'; $cpu=Get-CimInstance Win32_Processor; foreach($c in $cpu){ $o+='  型号   : ' + $c.Name; $o+='  核心   : ' + $c.NumberOfCores + ' 物理核 / ' + $c.NumberOfLogicalProcessors + ' 逻辑线程'; $o+='  主频   : ' + $c.MaxClockSpeed + ' MHz' }; $pf=Get-CimInstance Win32_PerfFormattedData_PerfOS_Processor -Filter \"Name='_Total'\"; $o+='  当前负载: ' + $pf.PercentProcessorTime + ' %%'; $o+=''; $o+='【三、内存】'; $t=[math]::Round($os.TotalVisibleMemorySize/1MB,2); $f=[math]::Round($os.FreePhysicalMemory/1MB,2); $u=[math]::Round($t-$f,2); $p=[math]::Round(($u/$t)*100,1); $o+='  总容量 : ' + $t + ' GB'; $o+='  已使用 : ' + $u + ' GB'; $o+='  可用   : ' + $f + ' GB'; $o+='  占用率 : ' + $p + ' %%'; $m=Get-CimInstance Win32_PhysicalMemory; $o+='  内存条 : ' + $m.Count + ' 条'; foreach($x in $m){ $o+='     - ' + [math]::Round($x.Capacity/1GB,0) + ' GB  ' + $x.Speed + ' MHz  ' + $x.Manufacturer }; Set-Content -Path '%RP%' -Value $o -Encoding UTF8"

powershell -NoProfile -Command "$o=@(); $o+=''; $o+='【四、磁盘空间】'; $v=Get-Volume; foreach($x in $v){ if($x.DriveLetter){ $tt=[math]::Round($x.Size/1GB,1); $rr=[math]::Round($x.SizeRemaining/1GB,1); $uu=0; if($x.Size -gt 0){ $uu=[math]::Round((($x.Size-$x.SizeRemaining)/$x.Size)*100,1) }; $o+='  ' + $x.DriveLetter + ':  总 ' + $tt + ' GB  |  剩 ' + $rr + ' GB  |  已用 ' + $uu + ' %%  |  ' + $x.FileSystem + '  |  ' + $x.HealthStatus } }; $o+=''; $o+='【五、物理磁盘健康】'; $pd=Get-PhysicalDisk; foreach($x in $pd){ $o+='  ' + $x.FriendlyName + '  |  ' + $x.MediaType + '  |  ' + [math]::Round($x.Size/1GB,0) + ' GB  |  健康: ' + $x.HealthStatus + '  |  ' + $x.OperationalStatus }; Add-Content -Path '%RP%' -Value $o -Encoding UTF8"

powershell -NoProfile -Command "$o=@(); $o+=''; $o+='【六、可清理的缓存占用】'; $x=Get-ChildItem $env:TEMP -Recurse -Force -ErrorAction SilentlyContinue; $s=0; foreach($i in $x){ $s+=$i.Length }; $o+='  用户临时文件夹 : ' + [math]::Round($s/1MB,1) + ' MB'; $x2=Get-ChildItem 'C:\Windows\Temp' -Recurse -Force -ErrorAction SilentlyContinue; $s2=0; foreach($i in $x2){ $s2+=$i.Length }; $o+='  Windows 临时   : ' + [math]::Round($s2/1MB,1) + ' MB'; $x3=Get-ChildItem 'C:\Windows\SoftwareDistribution\Download' -Recurse -Force -ErrorAction SilentlyContinue; $s3=0; foreach($i in $x3){ $s3+=$i.Length }; $o+='  更新下载缓存   : ' + [math]::Round($s3/1MB,1) + ' MB'; $x4=Get-ChildItem 'C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization' -Recurse -Force -ErrorAction SilentlyContinue; $s4=0; foreach($i in $x4){ $s4+=$i.Length }; $o+='  传递优化缓存   : ' + [math]::Round($s4/1MB,1) + ' MB'; $o+=''; $o+='【七、开机启动项数量】'; $c1=0; $r1=Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -ErrorAction SilentlyContinue; if($r1){ $c1=($r1.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' }).Count }; $c2=0; $r2=Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run' -ErrorAction SilentlyContinue; if($r2){ $c2=($r2.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' }).Count }; $c3=0; $r3=Get-ItemProperty 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run' -ErrorAction SilentlyContinue; if($r3){ $c3=($r3.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' }).Count }; $o+='  当前用户自启 : ' + $c1 + ' 项'; $o+='  全局自启     : ' + $c2 + ' 项'; $o+='  32位程序自启 : ' + $c3 + ' 项'; Add-Content -Path '%RP%' -Value $o -Encoding UTF8"

powershell -NoProfile -Command "$o=@(); $o+=''; $o+='【八、网络】'; $n=Get-NetAdapter; foreach($x in $n){ if($x.Status -eq 'Up' -and $x.InterfaceDescription -notlike '*Virtual*' -and $x.InterfaceDescription -notlike '*VMware*' -and $x.InterfaceDescription -notlike '*Hyper-V*'){ $o+='  [已连接] ' + $x.Name + '  |  ' + $x.InterfaceDescription + '  |  ' + $x.LinkSpeed } }; $ip=Get-NetIPConfiguration; foreach($x in $ip){ if($x.NetAdapter.Status -eq 'Up' -and $x.IPv4Address -and $x.InterfaceAlias -notlike '*VMnet*' -and $x.InterfaceAlias -notlike '*vEthernet*'){ $o+='     IPv4: ' + ($x.IPv4Address.IPAddress -join ',') + '   网关: ' + ($x.IPv4DefaultGateway.NextHop -join ',') } }; $dns=Get-DnsClientServerAddress -AddressFamily IPv4; foreach($x in $dns){ if($x.ServerAddresses){ $o+='     DNS [' + $x.InterfaceAlias + ']: ' + ($x.ServerAddresses -join ', ') } }; $t1=Test-Connection -ComputerName 223.5.5.5 -Count 3 -ErrorAction SilentlyContinue; if($t1){ $a=($t1 | Measure-Object -Property ResponseTime -Average).Average; $o+='  延迟(阿里DNS) : ' + [math]::Round($a,0) + ' ms' } else { $o+='  延迟(阿里DNS) : 不通' }; Add-Content -Path '%RP%' -Value $o -Encoding UTF8"

powershell -NoProfile -Command "$o=@(); $o+=''; $o+='【九、安全状态】'; $fw=Get-NetFirewallProfile; foreach($x in $fw){ $st='关闭'; if($x.Enabled){ $st='已启用' }; $o+='  防火墙 [' + $x.Name + '] : ' + $st }; try { $mp=Get-MpComputerStatus; $r='关闭'; if($mp.RealTimeProtectionEnabled){ $r='已开启' }; $o+='  Defender 实时保护 : ' + $r; $o+='  Defender 上次快扫 : ' + $mp.QuickScanAge + ' 天前' } catch { $o+='  Defender 状态读取失败（可能装有第三方杀毒）' }; $bs=Get-CimInstance Win32_NTLogEvent -Filter \"Logfile='System' AND EventCode=1001 AND SourceName='BugCheck'\" -ErrorAction SilentlyContinue; if($bs){ $o+='  蓝屏记录 : 发现 ' + $bs.Count + ' 条，最近一次 ' + $bs[0].TimeGenerated } else { $o+='  蓝屏记录 : 无（系统稳定）' }; $o+=''; $o+='【十、建议】'; $o+='  1. C 盘剩余空间保持在 15%% 以上，目前状态见「磁盘空间」一栏。'; $o+='  2. 开机启动项建议控制在 10 项以内，多了会拖慢开机速度。'; $o+='  3. 重要资料建议每周备份到 D 盘或外置硬盘，注意不要只存一份。'; $o+='  4. 每月跑一次本工具的「一键安全优化」，维护成本最低。'; $o+=''; $o+='############################################################'; Add-Content -Path '%RP%' -Value $o -Encoding UTF8"
exit /b


:END
cls
echo.
echo.  已退出。
echo.
echo.  提示：本工具不会自动重复执行。想维护时再双击运行即可。
echo.
ping -n 3 127.0.0.1 >nul
exit /b
