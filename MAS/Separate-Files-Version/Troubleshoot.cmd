@set masver=2.5
@setlocal DisableDelayedExpansion
@echo off



::=================================================================================================
::
::   此脚本是“Microsoft 激活脚本”（MAS）项目中的一部分。
::
::   主    页：mass grave[.]dev
::      Email：windowsaddict@protonmail.com
::
::=================================================================================================



::========================================================================================================================================

::  设置路径变量，如果在系统中配置错误时会有所帮助

set "PATH=%SystemRoot%\System32;%SystemRoot%\System32\wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\;%LocalAppData%\Microsoft\WindowsApps\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "PATH=%SystemRoot%\Sysnative;%SystemRoot%\Sysnative\wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%LocalAppData%\Microsoft\WindowsApps\;%PATH%"
)

::  如果脚本是由 x64 位 Windows 上的 x86 进程启动的，则将使用 x64 进程重新启动脚本
::  或者如果它是由 ARM64 Windows 上的 x86/ARM32 进程启动的，则将使用 ARM64 进程

set "_cmdf=%~f0"
for %%# in (%*) do (
if /i "%%#"=="r1" set r1=1
if /i "%%#"=="r2" set r2=1
if /i "%%#"=="-qedit" (
reg add HKCU\Console /v QuickEdit /t REG_DWORD /d "1" /f %nul1%
rem 查看下方的管理员提升代码了解它为什么在这里
)
)

if exist %SystemRoot%\Sysnative\cmd.exe if not defined r1 (
setlocal EnableDelayedExpansion
for %%# in (wt.exe) do @if "%%~$PATH:#"=="" start %SystemRoot%\Sysnative\cmd.exe /c ""!_cmdf!" %* r1" && exit /b
start wt.exe new-tab %SystemRoot%\Sysnative\cmd.exe /c ""!_cmdf!" %* r1" && exit /b
)

::  使用 ARM32 进程重新启动脚本（如果此脚本是由 ARM64 Windows 上的 x64 进程启动的）

if exist %SystemRoot%\SysArm32\cmd.exe if %PROCESSOR_ARCHITECTURE%==AMD64 if not defined r2 (
setlocal EnableDelayedExpansion
for %%# in (wt.exe) do @if "%%~$PATH:#"=="" start %SystemRoot%\SysArm32\cmd.exe /c ""!_cmdf!" %* r2" && exit /b
start wt.exe new-tab %SystemRoot%\SysArm32\cmd.exe /c ""!_cmdf!" %* r2" && exit /b
)

::========================================================================================================================================

set "blank="
set "mas=ht%blank%tps%blank%://mass%blank%grave.dev/"

::  检查 Null 服务是否正常工作，这对批处理脚本很重要

sc query Null | find /i "RUNNING"
if %errorlevel% NEQ 0 (
echo:
echo Null 服务未运行，脚本可能会崩溃……
echo:
echo:
echo 帮助 - %mas%troubleshoot.html
echo:
echo:
ping 127.0.0.1 -n 10
)
cls

::  检查 LF 行尾

pushd "%~dp0"
>nul findstr /v "$" "%~nx0" && (
echo:
echo 错误：脚本存在以 LF 行结束的问题，或者在脚本末尾缺少空行。
echo:
ping 127.0.0.1 -n 6 >nul
popd
exit /b
)
popd

::========================================================================================================================================

cls
color 07
title 疑难解答 %masver%

set _args=
set _elev=

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="-el"                    set _elev=1
)
)

set "nul1=1>nul"
set "nul2=2>nul"
set "nul6=2^>nul"
set "nul=>nul 2>&1"

set psc=powershell.exe
set winbuild=1
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G

set _NCS=1
if %winbuild% LSS 10586 set _NCS=0
if %winbuild% GEQ 10586 reg query "HKCU\Console" /v ForceV2 %nul2% | find /i "0x0" %nul1% && (set _NCS=0)

call :_colorprep

set "nceline=echo: &echo ==== 错误 ==== &echo:"
set "eline=echo: &call :_color %Red% "==== 错误 ====" &echo:"
set "line=_________________________________________________________________________________________________"
if %~z0 GEQ 200000 (set "_exitmsg=返回") else (set "_exitmsg=退出")

::========================================================================================================================================

if %winbuild% LSS 7600 (
%nceline%
echo 检测到不受支持的操作系统版本 [%winbuild%]。
echo 此项目仅支持 Windows 7/8/8.1/10/11 和它们对应的服务器版本。
goto at_done
)

for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" (
%nceline%
echo 在系统中未找到 powershell.exe。
goto at_done
)

::========================================================================================================================================

::  修复路径名称中的特殊字符限制

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

set _PSarg="""%~f0""" -el %_args%

set "_ttemp=%userprofile%\AppData\Local\Temp"

setlocal EnableDelayedExpansion

::========================================================================================================================================

echo "!_batf!" | find /i "!_ttemp!" %nul1% && (
if /i not "!_work!"=="!_ttemp!" (
%nceline%
echo 脚本是从 temp 文件夹启动的，
echo 最有可能的原因是，你是直接从压缩文件运行脚本。
echo:
echo 请解压压缩文件并从解压文件夹中启动脚本。
goto at_done
)
)

::========================================================================================================================================

::  将脚本提升为管理员权限及传递参数并防止循环

%nul1% fltmc || (
if not defined _elev for %%# in (wt.exe) do @if "%%~$PATH:#"=="" %psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && exit /b
if not defined _elev %psc% "start wt.exe -arg 'new-tab cmd.exe /c \"!_PSarg:'=''!\"' -verb runas" && exit /b
%nceline%
echo 此脚本需要管理员权限。
echo 为此，右键单击此脚本并选择“以管理员身份运行”。
goto at_done
)

::========================================================================================================================================

::  此代码仅禁用此 cmd.exe 会话的快速编辑，而不对注册表进行永久更改
::  添加它的原因是单击脚本窗口会暂停操作并导致脚本因错误而停止的混乱

for %%# in (%_args%) do (if /i "%%#"=="-qedit" set quedit=1)

reg query HKCU\Console /v QuickEdit %nul2% | find /i "0x0" %nul1% || if not defined quedit (
reg add HKCU\Console /v QuickEdit /t REG_DWORD /d "0" /f %nul1%
for %%# in (wt.exe) do @if "%%~$PATH:#"=="" start cmd.exe /c ""!_batf!" %_args% -qedit" && exit /b
start wt.exe new-tab cmd.exe /c ""!_batf!" %_args% -qedit" && exit /b
rem 快速编辑重置代码应添加在脚本的开头而不是此处，因为在某些情况下需要时间来反映
exit /b
)

::========================================================================================================================================

::  检查更新

set -=
set old=

for /f "delims=[] tokens=2" %%# in ('ping -4 -n 1 updatecheck.mass%-%grave.dev') do (
if not [%%#]==[] (echo "%%#" | find "127.69" %nul1% && (echo "%%#" | find "127.69.%masver%" %nul1% || set old=1))
)

if defined old (
echo ________________________________________________
%eline%
echo echo 你正在运行旧版本的 MAS 版本 %masver%
echo ________________________________________________
echo:
echo [1] 下载最新版本 MAS
echo [0] 仍然继续
echo:
call :_color %_Green% "请输入一个菜单选项 [1,0] ："
choice /C:10 /N
if !errorlevel!==2 rem
if !errorlevel!==1 (start ht%-%tps://github.com/mass%-%gravel/Microsoft-Acti%-%vation-Scripts & start %mas% & exit /b)
)
cls

::========================================================================================================================================

setlocal DisableDelayedExpansion

::  检测桌面位置

set desktop=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "desktop=%%b"
if not defined desktop for /f "delims=" %%a in ('%psc% "& {write-host $([Environment]::GetFolderPath('Desktop'))}"') do call set "desktop=%%a"

if not defined desktop (
%eline%
echo 桌面位置未被检测到，正在中止……
goto at_done
)

setlocal EnableDelayedExpansion

::========================================================================================================================================

:at_menu

cls
color 07
title 疑难解答 %masver%
mode con cols=77 lines=30

echo:
echo:
echo:
echo:
echo:       _______________________________________________________________
echo:                                                   
call :_color2 %_White% "             [1] " %_Green% "帮助"
echo:             ___________________________________________________
echo:                                                                      
echo:             [2] Dism RestoreHealth
echo:             [3] SFC Scannow
echo:                                                                      
echo:             [4] 修复 WMI
echo:             [5] 修复许可
echo:             [6] 修复 WPA 注册表
echo:             ___________________________________________________
echo:
echo:             [0] %_exitmsg%
echo:       _______________________________________________________________
echo:          
call :_color2 %_White% "            " %_Green% "请输入一个菜单选项 ："
choice /C:1234560 /N
set _erl=%errorlevel%

if %_erl%==7 exit /b
if %_erl%==6 start %mas%fix-wpa-registry.html &goto at_menu
if %_erl%==5 goto:retokens
if %_erl%==4 goto:fixwmi
if %_erl%==3 goto:sfcscan
if %_erl%==2 goto:dism_rest
if %_erl%==1 start %mas%troubleshoot.html &goto at_menu
goto :at_menu

::========================================================================================================================================

:dism_rest

cls
mode 98, 30
title Dism /Online /Cleanup-Image /RestoreHealth

if %winbuild% LSS 9200 (
%eline%
echo 检测到不受支持的操作系统版本。
echo 仅限 Windows 8/8.1/10/11 及其对应的 Server 支持此命令。
goto :at_back
)

set _int=
for %%a in (l.root-servers.net resolver1.opendns.com download.windowsupdate.com google.com) do if not defined _int (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (if not [%%#]==[] set _int=1)
)

echo:
if defined _int (
echo      正在检查 Internet 连接        [已连接]
) else (
call :_color2 %_White% "     " %Red% "正在检查 Internet 连接        [未连接]"
)

echo %line%
echo:
echo      Dism 使用 Windows 更新来提供修复损坏所需的文件。
echo      这将需要 5-15 分钟或更长时间……
echo %line%
echo:
echo      注：
echo:
call :_color2 %_White% "     - " %Gray% "请确保 Internet 已连接。"
call :_color2 %_White% "     - " %Gray% "请确保 Windows 更新正常工作。"
echo:
echo %line%
echo:
choice /C:09 /N /M ">    [9] 继续 [0] 返回 ："
if %errorlevel%==1 goto at_menu

cls
mode 110, 30
%psc% Stop-Service TrustedInstaller -force %nul%

set _time=
for /f %%a in ('%psc% "Get-Date -format HH_mm_ss"') do set _time=%%a
echo:
echo 正在应用命令，
echo Dism /Online /Cleanup-Image /RestoreHealth
Dism /Online /Cleanup-Image /RestoreHealth

%psc% Stop-Service TrustedInstaller -force %nul%

if not exist "!desktop!\AT_Logs\" md "!desktop!\AT_Logs\" %nul%

call :compresslog cbs\CBS.log RHealth_CBS %nul%
call :compresslog DISM\dism.log RHealth_DISM %nul%

if not exist "!desktop!\AT_Logs\RHealth_CBS_%_time%.cab" (
copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "!desktop!\AT_Logs\RHealth_CBS_%_time%.log" %nul%
)

if not exist "!desktop!\AT_Logs\RHealth_DISM_%_time%.cab" (
copy /y /b "%SystemRoot%\logs\DISM\dism.log" "!desktop!\AT_Logs\RHealth_DISM_%_time%.log" %nul%
)

echo:
call :_color %Gray% "CBS 和 DISM 日志已被复制到桌面上的 AT_Logs 文件夹中。"
goto :at_back

::========================================================================================================================================

:sfcscan

cls
mode 98, 30
title sfc /scannow

echo:
echo %line%
echo:    
echo      系统文件检查器（System File Checker）将修复丢失或损坏的系统文件。
echo      这将需要 10-15 分钟或更长时间……
echo:
echo      如果 SFC 无法修复某些问题，则再次运行该命令以查看下一次是否能够修复。有时可能需
echo      要运行 sfc /scannow 命令 3 次，每次之后重新启动 PC 才能完全修复它能够修复的所有
echo      问题。
echo:   
echo %line%
echo:
choice /C:09 /N /M ">    [9] 继续 [0] 返回 ："
if %errorlevel%==1 goto at_menu

cls
%psc% Stop-Service TrustedInstaller -force %nul%

set _time=
for /f %%a in ('%psc% "Get-Date -format HH_mm_ss"') do set _time=%%a
echo:
echo 正在应用命令，
echo sfc /scannow
sfc /scannow

%psc% Stop-Service TrustedInstaller -force %nul%

if not exist "!desktop!\AT_Logs\" md "!desktop!\AT_Logs\" %nul%

call :compresslog cbs\CBS.log SFC_CBS %nul%

if not exist "!desktop!\AT_Logs\SFC_CBS_%_time%.cab" (
copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "!desktop!\AT_Logs\SFC_CBS_%_time%.log" %nul%
)

echo:
call :_color %Gray% "CBS 和主要提取的日志被复制到桌面上的 AT_Logs 文件夹中。"
goto :at_back

::========================================================================================================================================

:retokens

cls
mode con cols=125 lines=32
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=200;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}"
title 修复许可（ClipSVC ^+ Office vNext ^+ SPP ^+ OSPP）

echo:
echo %line%
echo:   
echo      注：
echo:
echo       - 它有助于解决激活问题。
echo:
echo       - 此选项将会，
echo            - 反激活 Windows 和 Office，你可能需要重新激活
echo              如果 Windows 是使用主板 / OEM / 数字许可证激活的，请不要担心
echo:
echo            - 清理 ClipSVC、Office vNext、SPP 和 OSPP 许可
echo            - 修复令牌文件夹和注册表的 SPP 权限
echo            - 触发 Office 的修复选项。
echo:
call :_color2 %_White% "      - " %Red% "仅在必要时应用。"
echo:
echo %line%
echo:
choice /C:09 /N /M ">    [9] 继续 [0] 返回 ："
if %errorlevel%==1 goto at_menu

::========================================================================================================================================

::  重建 ClipSVC 许可证

cls
:cleanlicensing

echo:
echo %line%
echo:
call :_color %Blue% "正在重建 ClipSVC 许可证"
echo:

if %winbuild% LSS 10240 (
echo ClipSVC 许可证重建仅在 Win 10/11 和服务器对应版本上受支持。
echo 正在跳过……
goto :cleanvnext
)

%psc% "(([WMISEARCHER]'SELECT Name FROM SoftwareLicensingProduct WHERE LicenseStatus=1 AND GracePeriodRemaining=0 AND PartialProductKey IS NOT NULL').Get()).Name" %nul2% | findstr /i "Windows" %nul1% && (
echo Windows 已永久激活。
echo 正在跳过重建 ClipSVC 许可证……
goto :cleanvnext
)

echo 正在停止 ClipSVC 服务……
%psc% Stop-Service ClipSVC -force %nul%
timeout /t 2 %nul%

echo:
echo 正在将命令应用于清理 ClipSVC 许可证……
echo rundll32 clipc.dll,ClipCleanUpState

rundll32 clipc.dll,ClipCleanUpState

if %winbuild% LEQ 10240 (
echo [成功]
) else (
if exist "%ProgramData%\Microsoft\Windows\ClipSVC\tokens.dat" (
call :_color %Red% "[失败]"
) else (
echo [成功]
)
)

::  下方的注册表项（易失性和受保护）在 ClipSVC 许可证清理命令之后创建，并将在系统重新启动后自动删除。
::  需要删除它才能激活系统，无需重新启动。

set "RegKey=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState"
set "_ident=HKU\S-1-5-19\SOFTWARE\Microsoft\IdentityCRL"

reg query "%RegKey%" %nul% && %nul% call :regownstart
reg delete "%RegKey%" /f %nul% 

echo:
echo 正在删除易失性和受保护的注册表项……
echo [%RegKey%]
reg query "%RegKey%" %nul% && (
call :_color %Red% "[失败]"
echo 重新启动系统，这将会自动删除此注册表项。
) || (
echo [成功]
)

::   清除 HWID 令牌相关注册表以修复激活，以防出现任何损坏

echo:
echo 正在删除 IdentityCRL 注册表项……
echo [%_ident%]
reg delete "%_ident%" /f %nul%
reg query "%_ident%" %nul% && (
call :_color %Red% "[失败]"
) || (
echo [成功]
)

%psc% Stop-Service ClipSVC -force %nul%

::  重建 ClipSVC 文件夹以修复权限问题

echo:
if %winbuild% GTR 10240 (
echo 正在删除文件夹 %ProgramData%\Microsoft\Windows\ClipSVC\
rmdir /s /q "C:\ProgramData\Microsoft\Windows\ClipSvc" %nul%

if exist "%ProgramData%\Microsoft\Windows\ClipSVC\" (
call :_color %Red% "[失败]"
) else (
echo [成功]
)

echo:
echo Rebuilding Folder %ProgramData%\Microsoft\Windows\ClipSVC\
%psc% Start-Service ClipSVC %nul%
timeout /t 3 %nul%
if not exist "%ProgramData%\Microsoft\Windows\ClipSVC\" timeout /t 5 %nul%
if not exist "%ProgramData%\Microsoft\Windows\ClipSVC\" (
call :_color %Red% "[失败]"
) else (
echo [成功]
)
)

echo:
echo 正在重启 [wlidsvc LicenseManager] 服务……
for %%# in (wlidsvc LicenseManager) do (%psc% Restart-Service %%# %nul%)

::========================================================================================================================================

::  查找 Office vNext 许可证块的残余并将其删除，因为它会阻止非 vNext 许可证的显示
::  https://learn.microsoft.com/en-us/office/troubleshoot/activation/reset-office-365-proplus-activation-state

:cleanvnext

echo:
echo %line%
echo:
call :_color %Blue% "正在清理 Office vNext 许可证"
echo:

setlocal DisableDelayedExpansion
set "_Local=%LocalAppData%"
setlocal EnableDelayedExpansion

attrib -R "!ProgramData!\Microsoft\Office\Licenses" %nul%
attrib -R "!_Local!\Microsoft\Office\Licenses" %nul%

if exist "!ProgramData!\Microsoft\Office\Licenses\" (
rd /s /q "!ProgramData!\Microsoft\Office\Licenses\" %nul%
if exist "!ProgramData!\Microsoft\Office\Licenses\" (
echo 删除失败 - !ProgramData!\Microsoft\Office\Licenses\
) else (
echo 已删除文件夹 - !ProgramData!\Microsoft\Office\Licenses\
)
) else (
echo 未找到 - !ProgramData!\Microsoft\Office\Licenses\
)

if exist "!_Local!\Microsoft\Office\Licenses\" (
rd /s /q "!_Local!\Microsoft\Office\Licenses\" %nul%
if exist "!_Local!\Microsoft\Office\Licenses\" (
echo 删除失败 - !_Local!\Microsoft\Office\Licenses\
) else (
echo 已删除文件夹 - !_Local!\Microsoft\Office\Licenses\
)
) else (
echo 未找到 - !_Local!\Microsoft\Office\Licenses\
)


echo:
for /f "tokens=* delims=" %%a in ('%psc% "Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' | ForEach-Object { Split-Path -Path $_.PSPath -Leaf }" %nul6%') do (if defined _sid (set "_sid=!_sid! HKU\%%a") else (set "_sid=HKU\%%a"))

set regfound=
for %%# in (HKCU !_sid!) do (
for %%A in (
%%#\Software\Microsoft\Office\16.0\Common\Licensing
%%#\Software\Microsoft\Office\16.0\Common\Identity
%%#\Software\Microsoft\Office\16.0\Registration
) do (
reg query %%A %nul% && (
set regfound=1
reg delete %%A /f %nul% && (
echo 已删除注册表 - %%A
) || (
echo 删除失败 - %%A
)
)
)
)
if not defined regfound echo 未找到 - Office vNext 注册表

::========================================================================================================================================

::  重建 SPP 许可令牌

echo:
echo %line%
echo:
call :_color %Blue% "正在重建 SPP 许可令牌"
echo:

call :scandat check

if not defined token (
call :_color %Red% "未找到 tokens.dat 文件。"
) else (
echo tokens.dat 文件：[%token%]
)

echo:
set wpainfo=
for /f "delims=" %%a in ('%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':wpatest\:.*';iex ($f[1]);" %nul6%') do (set wpainfo=%%a)
echo "%wpainfo%" | find /i "Error Found" %nul% && (
call :_color %Red% "WPA 注册错误：%wpainfo%"
) || (
echo WPA 注册表数：%wpainfo%
)

set tokenstore=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v TokenStore %nul6%') do call set "tokenstore=%%b"

::  检查 sppsvc 权限并应用修补程序

if %winbuild% GEQ 10240 (

echo:
echo 正在检查 SPP 权限相关问题……
call :checkperms

if defined permerror (

mkdir "%tokenstore%" %nul%
set "d=$sddl = 'O:BAG:BAD:PAI(A;OICI;FA;;;SY)(A;OICI;FA;;;BA)(A;OICIIO;GR;;;BU)(A;;FR;;;BU)(A;OICI;FA;;;S-1-5-80-123231216-2592883651-3715271367-3753151631-4175906628)';"
set "d=!d! $AclObject = New-Object System.Security.AccessControl.DirectorySecurity;"
set "d=!d! $AclObject.SetSecurityDescriptorSddlForm($sddl);"
set "d=!d! Set-Acl -Path %tokenstore% -AclObject $AclObject;"
%psc% "!d!" %nul%

for %%# in (
"HKLM:\SYSTEM\WPA_QueryValues, EnumerateSubKeys, WriteKey"
"HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform_SetValue"
) do for /f "tokens=1,2 delims=_" %%A in (%%#) do (
set "d=$acl = Get-Acl '%%A';"
set "d=!d! $rule = New-Object System.Security.AccessControl.RegistryAccessRule ('NT Service\sppsvc', '%%B', 'ContainerInherit, ObjectInherit','None','Allow');"
set "d=!d! $acl.ResetAccessRule($rule);"
set "d=!d! $acl.SetAccessRule($rule);"
set "d=!d! Set-Acl -Path '%%A' -AclObject $acl"
%psc% "!d!" %nul%
)

call :checkperms
if defined permerror (
call :_color %Red% "[修复失败]"
) else (
echo [修复成功]
)
) else (
echo [未发现错误]
)
)

echo:
echo 正在停止 sppsvc 服务   ...
%psc% Stop-Service sppsvc -force %nul%

echo:
call :scandat delete
call :scandat check

if defined token (
echo:
call :_color %Red% "删除 .dat 文件失败。"
echo:
)

echo:
echo 正在重新安装系统许可 [slmgr /rilc]……
cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
if %errorlevel% NEQ 0 cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
if %errorlevel% EQU 0 (
echo [成功]
) else (
call :_color %Red% "[失败]"
)

call :scandat check

echo:
if not defined token (
call :_color %Red% "重建 tokens.dat 文件失败。"
) else (
echo tokens.dat 文件已成功重建。
)

::========================================================================================================================================

::  重建 OSPP 令牌

echo:
echo %line%
echo:
call :_color %Blue% "正在重建 OSPP 许可令牌"
echo:

sc qc osppsvc %nul% || (
echo 未安装基于 OSPP 的 Office
echo 正在跳过重建 OSPP 令牌……
goto :repairoffice
)

call :scandatospp check

if not defined token (
call :_color %Red% "未找到 tokens.dat 文件。"
) else (
echo tokens.dat 文件：[%token%]
)

echo:
echo 正在停止 osppsvc 服务……
%psc% Stop-Service osppsvc -force %nul%

echo:
call :scandatospp delete
call :scandatospp check

if defined token (
echo:
call :_color %Red% "删除 .dat 文件失败。"
echo:
)

echo:
echo 正在启动 osppsvc 服务以生成 tokens.dat
%psc% Start-Service osppsvc %nul%
call :scandatospp check
if not defined token (
%psc% Stop-Service osppsvc -force %nul%
%psc% Start-Service osppsvc %nul%
timeout /t 3 %nul%
)

call :scandatospp check

echo:
if not defined token (
call :_color %Red% "重建 tokens.dat 文件失败。"
) else (
echo tokens.dat 文件已成功重建。
)

::========================================================================================================================================

:repairoffice

echo:
echo %line%
echo:
call :_color %Blue% "正在修复 Office 许可证"
echo:

for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PROCESSOR_ARCHITECTURE') do set arch=%%b

if /i "%arch%"=="ARM64" (
echo:
echo 已发现 ARM64 Windows。
echo 你需要在 Office 的 Windows 设置中使用修复选项。
echo:
start ms-settings:appsfeatures
goto :repairend
)

if /i "%arch%"=="x86" (
set arch=X86
) else (
set arch=X64
)

for %%# in (68 86) do (
for %%A in (msi14 msi15 msi16 c2r14 c2r15 c2r16) do (set %%A_%%#=&set %%Arepair%%#=)
)

set _68=HKLM\SOFTWARE\Microsoft\Office
set _86=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office

reg query %_68%\14.0\CVH /f Click2run /k %nul% && (set "c2r14_68=Office 14.0 C2R x86/x64"  & set "c2r14repair68=")
reg query %_86%\14.0\CVH /f Click2run /k %nul% && (set "c2r14_86=Office 14.0 C2R x86"      & set "c2r14repair86=")

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi14_86=Office 14.0 MSI x86"      & set "msi14repair86=%systemdrive%\Program Files (x86)\Common Files\Microsoft Shared\OFFICE14\Office Setup Controller\Setup.exe")
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi14_68=Office 14.0 MSI x86/x64"  & set "msi14repair68=%systemdrive%\Program Files\Common Files\microsoft shared\OFFICE14\Office Setup Controller\Setup.exe")
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi15_86=Office 15.0 MSI x86"      & set "msi15repair86=%systemdrive%\Program Files (x86)\Common Files\Microsoft Shared\OFFICE15\Office Setup Controller\Setup.exe")
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi15_68=Office 15.0 MSI x86/x64"  & set "msi15repair68=%systemdrive%\Program Files\Common Files\microsoft shared\OFFICE15\Office Setup Controller\Setup.exe")
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi16_86=Office 16.0 MSI x86"      & set "msi16repair86=%systemdrive%\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\Office Setup Controller\Setup.exe")
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set "msi16_68=Office 16.0 MSI x86/x64"  & set "msi16repair68=%systemdrive%\Program Files\Common Files\Microsoft Shared\OFFICE16\Office Setup Controller\Setup.exe")

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\15.0\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (set "c2r15_86=Office 15.0 C2R x86"      & set "c2r15repair86=%systemdrive%\Program Files\Microsoft Office 15\Client%arch%\integratedoffice.exe")
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\15.0\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (set "c2r15_68=Office 15.0 C2R x86/x64"  & set "c2r15repair68=%systemdrive%\Program Files\Microsoft Office 15\Client%arch%\integratedoffice.exe")
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms"    (set "c2r16_86=Office 16.0 C2R x86"      & set "c2r16repair86=%systemdrive%\Program Files\Microsoft Office 15\Client%arch%\OfficeClickToRun.exe")
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms"    (set "c2r16_68=Office 16.0 C2R x86/x64"  & set "c2r16repair68=%systemdrive%\Program Files\Microsoft Office 15\Client%arch%\OfficeClickToRun.exe")

set uwp16=
if %winbuild% GEQ 10240 (
%psc% "Get-AppxPackage -name "Microsoft.Office.Desktop"" | find /i "Office" %nul1% && set uwp16=Office 16.0 UWP
)

set /a counter=0
echo 正在检查已安装的 Office 版本……
echo:

for %%# in (
"%msi14_68%"
"%msi14_86%"
"%msi15_68%"
"%msi15_86%"
"%msi16_68%"
"%msi16_86%"
"%c2r14_68%"
"%c2r14_86%"
"%c2r15_68%"
"%c2r15_86%"
"%c2r16_68%"
"%c2r16_86%"
"%uwp16%"
) do (
if not "%%#"=="""" (
set insoff=%%#
set insoff=!insoff:"=!
echo [!insoff!]
set /a counter+=1
)
)

if %counter% GTR 1 (
%eline%
echo 已找到多个 office 版本。
echo 建议只安装 office 的一个版本。
echo ________________________________________________________________
echo:
)

if %counter% EQU 0 (
echo:
echo 没有找到已安装的 Office。
goto :repairend
echo:
) else (
echo:
call :_color %_Yellow% "将弹出一个窗口，在此窗口中你需要选择 [快速] 修复选项……"
call :_color %_Yellow% "请按任意键继续执行……"
echo:
pause %nul1%
)

if defined uwp16 (
echo:
echo 注：正在跳过针对 Office 16.0 UWP 的修复。
echo     你需要在 Windows 设置中使用重置选项。
echo ________________________________________________________________
echo:
start ms-settings:appsfeatures
)

set c2r14=
if defined c2r14_68 set c2r14=1
if defined c2r14_86 set c2r14=1

if defined c2r14 (
echo:
echo 注：正在跳过 Office 14.0 C2R 的修复
echo     你需要在 Windows 设置中使用修复选项。
echo ________________________________________________________________
echo:
start appwiz.cpl
)

if defined msi14_68 if exist "%msi14repair68%" echo 正在运行 - "%msi14repair68%"                    & "%msi14repair68%"
if defined msi14_86 if exist "%msi14repair86%" echo 正在运行 - "%msi14repair86%"                    & "%msi14repair86%"
if defined msi15_68 if exist "%msi15repair68%" echo 正在运行 - "%msi15repair68%"                    & "%msi15repair68%"
if defined msi15_86 if exist "%msi15repair86%" echo 正在运行 - "%msi15repair86%"                    & "%msi15repair86%"
if defined msi16_68 if exist "%msi16repair68%" echo 正在运行 - "%msi16repair68%"                    & "%msi16repair68%"
if defined msi16_86 if exist "%msi16repair86%" echo 正在运行 - "%msi16repair86%"                    & "%msi16repair86%"
if defined c2r15_68 if exist "%c2r15repair68%" echo 正在运行 - "%c2r15repair68%" REPAIRUI RERUNMODE & "%c2r15repair68%" REPAIRUI RERUNMODE
if defined c2r15_86 if exist "%c2r15repair86%" echo 正在运行 - "%c2r15repair86%" REPAIRUI RERUNMODE & "%c2r15repair86%" REPAIRUI RERUNMODE
if defined c2r16_68 if exist "%c2r16repair68%" echo 正在运行 - "%c2r16repair68%" scenario=Repair    & "%c2r16repair68%" scenario=Repair
if defined c2r16_86 if exist "%c2r16repair86%" echo 正在运行 - "%c2r16repair86%" scenario=Repair    & "%c2r16repair86%" scenario=Repair

:repairend

echo:
echo %line%
echo:
echo:
call :_color %Green% "已完成"
goto :at_back

::========================================================================================================================================

:fixwmi

cls
mode 98, 34
title 修复 WMI

::  https://techcommunity.microsoft.com/t5/ask-the-performance-team/wmi-repository-corruption-or-not/ba-p/375484

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
%eline%
echo 不建议在 Windows Server 上重建 WMI。正在中止……
goto :at_back
)

for %%# in (wmic.exe) do @if "%%~$PATH:#"=="" (
%eline%
echo 在系统中找不到 WMIC.exe 文件。正在中止……
goto :at_back
)

echo:
echo 正在检查 WMI
call :checkwmi

::  首先应用基本修复并检查

if defined error (
%psc% Stop-Service Winmgmt -force %nul%
winmgmt /salvagerepository %nul%
call :checkwmi
)

if not defined error (
echo [正在工作]
echo 无需应用此选项。正在中止……
goto :at_back
)

call :_color %Red% "[未响应]"

set _corrupt=
sc start Winmgmt %nul%
if %errorlevel% EQU 1060 set _corrupt=1
sc query Winmgmt %nul% || set _corrupt=1
for %%G in (DependOnService Description DisplayName ErrorControl ImagePath ObjectName Start Type) do if not defined _corrupt (reg query HKLM\SYSTEM\CurrentControlSet\Services\Winmgmt /v %%G %nul% || set _corrupt=1)

echo:
if defined _corrupt (
%eline%
echo Winmgmt 服务未安装。正在中止……
goto :at_back
)

echo 正在停止 Winmgmt 服务
sc config Winmgmt start= disabled %nul%
if %errorlevel% EQU 0 (
echo [成功]
) else (
call :_color %Red% "[失败] 正在中止……"
sc config Winmgmt start= auto %nul%
goto :at_back
)

echo:
echo 正在停止 Winmgmt 服务
%psc% Stop-Service Winmgmt -force %nul%
%psc% Stop-Service Winmgmt -force %nul%
%psc% Stop-Service Winmgmt -force %nul%
sc query Winmgmt | find /i "STOPPED" %nul% && (
echo [成功]
) || (
call :_color %Red% "[失败]"
echo:
call :_color %Blue% "建议选择 [重新启动] 选项，然后再次应用修复 WMI 选项。"
echo %line%
echo:
choice /C:21 /N /M "> [1] 重新启动  [2] 恢复更改 ："
if !errorlevel!==1 (sc config Winmgmt start= auto %nul%&goto :at_back)
echo:
echo 正在重启……
shutdown -t 5 -r
exit
)

echo:
echo 正在删除 WMI 存储库
rmdir /s /q "%windir%\System32\wbem\repository\" %nul%
if exist "%windir%\System32\wbem\repository\" (
call :_color %Red% "[失败]"
) else (
echo [成功]
)

echo:
echo 正在启用 Winmgmt 服务
sc config Winmgmt start= auto %nul%
if %errorlevel% EQU 0 (
echo [成功]
) else (
call :_color %Red% "[失败]"
)

call :checkwmi
if not defined error (
echo:
echo 正在检查 WMI
call :_color %Green% "[正在工作]"
goto :at_back
)

echo:
echo 正在注册 .dll 并编译 .mof 和 .mfl
call :registerobj %nul%

echo:
echo 正在检查 WMI
call :checkwmi
if defined error (
call :_color %Red% "[未响应]"
echo:
echo 运行 [Dism RestoreHealth] 和 [SFC Scannow] 选项，并确保没有错误。
) else (
call :_color %Green% "[正在工作]"
)

goto :at_back

:registerobj

::  https://eskonr.com/2012/01/how-to-fix-wmi-issues-automatically/

%psc% Stop-Service Winmgmt -force %nul%
cd /d %systemroot%\system32\wbem\
regsvr32 /s %systemroot%\system32\scecli.dll
regsvr32 /s %systemroot%\system32\userenv.dll
mofcomp cimwin32.mof
mofcomp cimwin32.mfl
mofcomp rsop.mof
mofcomp rsop.mfl
for /f %%s in ('dir /b /s *.dll') do regsvr32 /s %%s
for /f %%s in ('dir /b *.mof') do mofcomp %%s
for /f %%s in ('dir /b *.mfl') do mofcomp %%s

winmgmt /salvagerepository
winmgmt /resetrepository
exit /b

:checkwmi

::  https://learn.microsoft.com/en-us/windows/win32/wmisdk/wmi-error-constants

set error=
wmic path Win32_ComputerSystem get CreationClassName /value %nul2% | find /i "computersystem" %nul1%
if %errorlevel% NEQ 0 (set error=1& exit /b)
winmgmt /verifyrepository %nul%
if %errorlevel% NEQ 0 (set error=1& exit /b)

cscript //nologo %windir%\system32\slmgr.vbs /dlv %nul%
cmd /c exit /b %errorlevel%
echo "0x%=ExitCode%" | findstr /i "0x800410 0x800440" %nul1%
if %errorlevel% EQU 0 set error=1
exit /b

::========================================================================================================================================

:at_back

echo:
echo %line%
echo:
call :_color %_Yellow% "请按任意键返回……"
pause %nul1%
goto :at_menu

::========================================================================================================================================

:at_done

echo:
echo 请按任意键%_exitmsg%脚本……
pause %nul1%
exit /b

::========================================================================================================================================

:compresslog

::  https://stackoverflow.com/a/46268232

set "ddf="%SystemRoot%\Temp\ddf""
%nul% del /q /f %ddf%
echo/.New Cabinet>%ddf%
echo/.set Cabinet=ON>>%ddf%
echo/.set CabinetFileCountThreshold=0;>>%ddf%
echo/.set Compress=ON>>%ddf%
echo/.set CompressionType=LZX>>%ddf%
echo/.set CompressionLevel=7;>>%ddf%
echo/.set CompressionMemory=21;>>%ddf%
echo/.set FolderFileCountThreshold=0;>>%ddf%
echo/.set FolderSizeThreshold=0;>>%ddf%
echo/.set GenerateInf=OFF>>%ddf%
echo/.set InfFileName=nul>>%ddf%
echo/.set MaxCabinetSize=0;>>%ddf%
echo/.set MaxDiskFileCount=0;>>%ddf%
echo/.set MaxDiskSize=0;>>%ddf%
echo/.set MaxErrors=1;>>%ddf%
echo/.set RptFileName=nul>>%ddf%
echo/.set UniqueFiles=ON>>%ddf%
for /f "tokens=* delims=" %%D in ('dir /a:-D/b/s "%SystemRoot%\logs\%1"') do (
 echo/"%%~fD"  /inf=no;>>%ddf%
)
makecab /F %ddf% /D DiskDirectory1="" /D CabinetNameTemplate="!desktop!\AT_Logs\%2_%_time%.cab"
del /q /f %ddf%
exit /b

::========================================================================================================================================

::  此代码检查 HKLM\SYSTEM\WPA 中是否存在无效的注册表项。即使在正常运行的系统上也可能会出现此问题。

:wpatest:
$wpaKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64').OpenSubKey("SYSTEM\\WPA")
$count = $wpaKey.SubKeyCount

$osVersion = [System.Environment]::OSVersion.Version
$minBuildNumber = 14393

if ($osVersion.Build -ge $minBuildNumber) {
    $subkeyHashTable = @{}
    foreach ($subkeyName in $wpaKey.GetSubKeyNames()) {
        $keyNumber = $subkeyName -replace '.*-', ''
        $subkeyHashTable[$keyNumber] = $true
    }
    for ($i=1; $i -le $count; $i++) {
        if (-not $subkeyHashTable.ContainsKey("$i")) {
            Write-Host "Total Keys $count. Error Found- $i key does not exist"
			$wpaKey.Close()
            exit
        }
    }
}
$wpaKey.GetSubKeyNames() | ForEach-Object {
    $subkey = $wpaKey.OpenSubKey($_)
    $p = $subkey.GetValueNames()
    if (($p | Where-Object { $subkey.GetValueKind($_) -eq [Microsoft.Win32.RegistryValueKind]::Binary }).Count -eq 0) {
        Write-Host "Total Keys $count. Error Found- Binary Data is corrupt"
		$wpaKey.Close()
        exit
    }
}
$count
$wpaKey.Close()
:wpatest:

::========================================================================================================================================

:checkperms

set permerror=
if not exist "%tokenstore%\" set permerror=1

for %%# in (
"%tokenstore%"
"HKLM:\SYSTEM\WPA"
"HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
) do if not defined permerror (
%psc% "$acl = Get-Acl '%%#'; if ($acl.Access.Where{ $_.IdentityReference -eq 'NT SERVICE\sppsvc' -and $_.AccessControlType -eq 'Deny' -or $acl.Access.IdentityReference -notcontains 'NT SERVICE\sppsvc'}) {Exit 2}" %nul%
if !errorlevel!==2 set permerror=1
)
exit /b

::========================================================================================================================================

:scandat

set token=
for %%# in (
%Systemdrive%\Windows\System32\spp\store_test\2.0\
%Systemdrive%\Windows\System32\spp\store\
%Systemdrive%\Windows\System32\spp\store\2.0\
%Systemdrive%\Windows\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\
) do (

if %1==check (
if exist %%#tokens.dat set token=%%#tokens.dat
)

if %1==delete (
if exist %%# (
%nul% dir /a-d /s "%%#*.dat" && (
attrib -r -s -h "%%#*.dat" /S
del /S /F /Q "%%#*.dat"
)
)
)
)
exit /b

:scandatospp

set token=
for %%# in (
%ProgramData%\Microsoft\OfficeSoftwareProtectionPlatform\
) do (

if %1==check (
if exist %%#tokens.dat set token=%%#tokens.dat
)

if %1==delete (
if exist %%# (
%nul% dir /a-d /s "%%#*.dat" && (
attrib -r -s -h "%%#*.dat" /S
del /S /F /Q "%%#*.dat"
)
)
)
)
exit /b

::========================================================================================================================================

:regownstart

%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':regown\:.*';iex ($f[1]);"
exit /b

::  以下是获取易失性注册表项的所有权并将其删除的代码
::  HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState

:regown:
$AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1)
$ModuleBuilder = $AssemblyBuilder.DefineDynamicModule(2, $False)
$TypeBuilder = $ModuleBuilder.DefineType(0)

$TypeBuilder.DefinePInvokeMethod('RtlAdjustPrivilege', 'ntdll.dll', 'Public, Static', 1, [int], @([int], [bool], [bool], [bool].MakeByRefType()), 1, 3) | Out-Null
$TypeBuilder.CreateType()::RtlAdjustPrivilege(9, $true, $false, [ref]$false) | Out-Null

$SID = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')
$IDN = ($SID.Translate([System.Security.Principal.NTAccount])).Value
$Admin = New-Object System.Security.Principal.NTAccount($IDN)

$path = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState'
$key = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64').OpenSubKey($path, 'ReadWriteSubTree', 'takeownership')

$acl = $key.GetAccessControl()
$acl.SetOwner($Admin)
$key.SetAccessControl($acl)

$rule = New-Object System.Security.AccessControl.RegistryAccessRule($Admin,"FullControl","Allow")
$acl.SetAccessRule($rule)
$key.SetAccessControl($acl)
:regown:

::========================================================================================================================================

:_color

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[0m
) else (
call :batcol %~1 "%~2"
)
exit /b

:_color2

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[%~3%~4%esc%[0m
) else (
call :batcol %~1 "%~2" %~3 "%~4"
)
exit /b

::=======================================

::  纯批处理方法的彩色文本
::  感谢 @dbenham 和 @jeb
::  stackoverflow.com/a/10407642

:batcol

pushd %_coltemp%
if not exist "'" (<nul >"'" set /p "=.")
setlocal
set "s=%~2"
set "t=%~4"
call :_batcol %1 s %3 t
del /f /q "'"
del /f /q "`.txt"
popd
exit /b

:_batcol

setlocal EnableDelayedExpansion
set "s=!%~2!"
set "t=!%~4!"
for /f delims^=^ eol^= %%i in ("!s!") do (
  if "!" equ "" setlocal DisableDelayedExpansion
    >`.txt (echo %%i\..\')
    findstr /a:%~1 /f:`.txt "."
    <nul set /p "=%_BS%%_BS%%_BS%%_BS%%_BS%%_BS%%_BS%"
)
if "%~4"=="" echo(&exit /b
setlocal EnableDelayedExpansion
for /f delims^=^ eol^= %%i in ("!t!") do (
  if "!" equ "" setlocal DisableDelayedExpansion
    >`.txt (echo %%i\..\')
    findstr /a:%~3 /f:`.txt "."
    <nul set /p "=%_BS%%_BS%%_BS%%_BS%%_BS%%_BS%%_BS%"
)
echo(
exit /b

::=======================================

:_colorprep

if %_NCS% EQU 1 (
for /F %%a in ('echo prompt $E ^| cmd') do set "esc=%%a"

set     "Red="41;97m""
set    "Gray="100;97m""
set   "Black="30m""
set   "Green="42;97m""
set    "Blue="44;97m""
set  "Yellow="43;97m""
set "Magenta="45;97m""

set    "_Red="40;91m""
set  "_Green="40;92m""
set   "_Blue="40;94m""
set  "_White="40;37m""
set "_Yellow="40;93m""

exit /b
)

for /f %%A in ('"prompt $H&for %%B in (1) do rem"') do set "_BS=%%A %%A"
set "_coltemp=%SystemRoot%\Temp"

set     "Red="CF""
set    "Gray="8F""
set   "Black="00""
set   "Green="2F""
set    "Blue="1F""
set  "Yellow="6F""
set "Magenta="5F""

set    "_Red="0C""
set  "_Green="0A""
set   "_Blue="09""
set  "_White="07""
set "_Yellow="0E""

exit /b

::========================================================================================================================================
::  下方保留空行
