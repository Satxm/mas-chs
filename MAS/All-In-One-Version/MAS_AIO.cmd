@set masver=2.5
@setlocal DisableDelayedExpansion
@echo off


::  对于命令行开关，请查看 mass grave[.]dev/command_line_switches.html
::  如果你想更好地理解脚本，请从 MAS 独立文件版本读取。


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
reg add HKCU\Console /v QuickEdit /t REG_DWORD /d "1" /f 1>nul
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
title 微软激活脚本 %masver%

set _args=
set _elev=
set _MASunattended=

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="-el"                    set _elev=1
)
)

if defined _args echo "%_args%" | find /i "/" >nul && set _MASunattended=1

::========================================================================================================================================

set "nul1=1>nul"
set "nul2=2>nul"
set "nul6=2^>nul"
set "nul=>nul 2>&1"

set winbuild=1
set psc=powershell.exe
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G

set _NCS=1
if %winbuild% LSS 10586 set _NCS=0
if %winbuild% GEQ 10586 reg query "HKCU\Console" /v ForceV2 %nul2% | find /i "0x0" %nul1% && (set _NCS=0)

call :_colorprep

set "nceline=echo: &echo ==== 错误 ==== &echo:"
set "eline=echo: &call :_color %Red% "==== 错误 ====" &echo:"

::========================================================================================================================================

if %winbuild% LSS 7600 (
%nceline%
echo 检测到不受支持的操作系统版本 [%winbuild%]。
echo 此项目仅支持 Windows 7/8/8.1/10/11 和它们对应的服务器版本。
goto MASend
)

for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" (
%nceline%
echo 在系统中未找到 powershell.exe。
echo 正在中止……
goto MASend
)

::========================================================================================================================================

::  修复路径名称中的特殊字符限制

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

set _PSarg="""%~f0""" -el %_args%

set "_ttemp=%userprofile%\AppData\Local\Temp"
set "_Local=%LocalAppData%"

setlocal EnableDelayedExpansion

::========================================================================================================================================

echo "!_batf!" | find /i "!_ttemp!" %nul1% && (
if /i not "!_work!"=="!_ttemp!" (
%nceline%
echo 脚本是从 temp 文件夹启动的，
echo 最有可能的原因是，你是直接从压缩文件运行脚本。
echo:
echo 请解压压缩文件并从解压文件夹中启动脚本。
goto MASend
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
goto MASend
)

if not exist "%SystemRoot%\Temp\" mkdir "%SystemRoot%\Temp" %nul%

::========================================================================================================================================

::  此代码仅禁用此 cmd.exe 会话的快速编辑，而不对注册表进行永久更改
::  添加它的原因是单击脚本窗口会暂停操作并导致脚本因错误而停止的混乱

if defined _MASunattended set quedit=1
for %%# in (%_args%) do (if /i "%%#"=="-qedit" set quedit=1)

reg query HKCU\Console /v QuickEdit %nul2% | find /i "0x0" %nul1% || if not defined quedit (
reg add HKCU\Console /v QuickEdit /t REG_DWORD /d "0" /f %nul1%
for %%# in (wt.exe) do @if "%%~$PATH:#"=="" start cmd.exe /c ""!_batf!" %_args% -qedit" && exit /b
start wt.exe new-tab cmd.exe /c ""!_batf!" %_args% -qedit" && exit /b
rem 快速编辑重置代码应添加在脚本的开头而不是此处，因为在某些情况下需要时间来反映
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
if not defined _MASunattended (
echo [1] 下载最新版本 MAS
echo [0] 仍然继续
echo:
call :_color %_Green% "请输入一个菜单选项 [1,0] ："
choice /C:10 /N
if !errorlevel!==2 rem
if !errorlevel!==1 (start ht%-%tps://github.com/mass%-%gravel/Microsoft-Acti%-%vation-Scripts & start %mas% & exit /b)
)
)
cls

::========================================================================================================================================

::  在无人参与模式下运行使用参数的脚本

set _elev=
if defined _args echo "%_args%" | find /i "/S" %nul% && (set "_silent=%nul%") || (set _silent=)
if defined _args echo "%_args%" | find /i "/" %nul% && (
echo "%_args%" | find /i "/HWID"   %nul% && (setlocal & cls & (call :HWIDActivation   %_args% %_silent%) & endlocal)
echo "%_args%" | find /i "/KMS38"  %nul% && (setlocal & cls & (call :KMS38Activation  %_args% %_silent%) & endlocal)
echo "%_args%" | find /i "/KMS-"   %nul% && (setlocal & cls & (call :KMSActivation    %_args% %_silent%) & endlocal)
echo "%_args%" | find /i "/Ohook"  %nul% && (setlocal & cls & (call :OhookActivation  %_args% %_silent%) & endlocal)
exit /b
)

::========================================================================================================================================

setlocal DisableDelayedExpansion

::  检测桌面位置

set _desktop_=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "_desktop_=%%b"
if not defined _desktop_ for /f "delims=" %%a in ('%psc% "& {write-host $([Environment]::GetFolderPath('Desktop'))}"') do call set "_desktop_=%%a"

setlocal EnableDelayedExpansion

::========================================================================================================================================

:MainMenu

cls
color 07
title 微软激活脚本 %masver%
mode 76, 30

echo:
echo:
echo:
echo:
echo:       ______________________________________________________________
echo:
echo:                 激 活 方 式：
echo:
echo:             [1] HWID        ^|  Windows           ^|      永久
echo:             [2] Ohook       ^|  Office            ^|      永久
echo:             [3] KMS38       ^|  Windows           ^|   2038 年
echo:             [4] 在线 KMS    ^|  Windows / Office  ^|    180 天
echo:             __________________________________________________      
echo:
echo:             [5] 激活状态
echo:             [6] 疑难解答
echo:             [7] 附加选项
echo:             [8] 帮助
echo:             [0] 退出
echo:       ______________________________________________________________
echo:
call :_color2 %_White% "          " %_Green% "请输入一个菜单选项 [1,2,3,4,5,6,7,8,0] ："
choice /C:123456780 /N
set _erl=%errorlevel%

if %_erl%==9 exit /b
if %_erl%==8 start %mas%troubleshoot.html & goto :MainMenu
if %_erl%==7 goto:Extras
if %_erl%==6 setlocal & call :troubleshoot      & cls & endlocal & goto :MainMenu
if %_erl%==5 setlocal & call :_Check_Status_wmi & cls & endlocal & goto :MainMenu
if %_erl%==4 setlocal & call :KMSActivation     & cls & endlocal & goto :MainMenu
if %_erl%==3 setlocal & call :KMS38Activation   & cls & endlocal & goto :MainMenu
if %_erl%==2 setlocal & call :OhookActivation   & cls & endlocal & goto :MainMenu
if %_erl%==1 setlocal & call :HWIDActivation    & cls & endlocal & goto :MainMenu
goto :MainMenu

::========================================================================================================================================

:Extras

cls
title 附加选项
mode 76, 30
echo:
echo:
echo:
echo:
echo:
echo:       ______________________________________________________________
echo:
echo:             [1] 更改 Windows 版本
echo:
echo:             [2] 解压 $OEM$ 文件夹
echo:
echo:             [3] 激活状态 [vbs]
echo:
echo:             [4] 下载正版 Windows / Office
echo:             __________________________________________________      
echo:                                                                     
echo:             [0] 返回到主菜单
echo:       ______________________________________________________________
echo:
call :_color2 %_White% "           " %_Green% "请输入一个菜单选项 [1,2,3,4,0] ："
choice /C:12340 /N
set _erl=%errorlevel%

if %_erl%==5 goto :MainMenu
if %_erl%==4 start %mas%genuine-installation-media.html & goto :Extras
if %_erl%==3 setlocal & call :_Check_Status_vbs & cls & endlocal & goto :Extras
if %_erl%==2 goto:Extract$OEM$
if %_erl%==1 setlocal & call :change_edition    & cls & endlocal & goto :Extras
goto :Extras

::========================================================================================================================================

:Extract$OEM$

cls
title 解压 $OEM$ 文件夹
mode 76, 30

if not exist "!_desktop_!\" (
%eline%
echo 桌面位置未被检测到，正在中止……
echo _____________________________________________________
echo:
call :_color %_Yellow% "请按任意键返回……"
pause >nul
goto Extras
)

if exist "!_desktop_!\$OEM$\" (
%eline%
echo $OEM$ 文件夹已经存在于桌面上。
echo _____________________________________________________
echo:
call :_color %_Yellow% "请按任意键返回……"
pause >nul
goto Extras
)

:Extract$OEM$2

cls
title 解压 $OEM$ 文件夹
mode 78, 30
echo:
echo:
echo:
echo:
echo:                        将 $OEM$ 文件夹解压缩到桌面上
echo:           ________________________________________________________
echo:
echo:             [1] HWID
echo:             [2] Ohook
echo:             [3] KMS38
echo:             [4] 在线 KMS
echo:
echo:             [5] HWID       （Windows） ^+ Ohook      （Office）
echo:             [6] HWID       （Windows） ^+ 在线 KMS   （Office）
echo:             [7] KMS38      （Windows） ^+ Ohook      （Office）
echo:             [8] KMS38      （Windows） ^+ 在线 KMS   （Office）
echo:             [9] 在线 KMS   （Windows） ^+ Ohook      （Office）
echo:
call :_color2 %_White% "              [R] " %_Green% "自述文件"
echo:             [0] 返回
echo:           ________________________________________________________
echo:  
call :_color2 %_White% "           " %_Green% "请输入一个菜单选项："
choice /C:123456789R0 /N
set _erl=%errorlevel%

if %_erl%==11 goto:Extras
if %_erl%==10 start %mas%oem-folder.html &goto:Extract$OEM$2
if %_erl%==9 (set "_oem=在线 KMS [Windows] + Ohook [Office]" & set "para=/KMS-ActAndRenewalTask /KMS-Windows /Ohook" &goto:Extract$OEM$3)
if %_erl%==8 (set "_oem=KMS38 [Windows] + 在线 KMS [Office]" & set "para=/KMS38 /KMS-ActAndRenewalTask /KMS-Office" &goto:Extract$OEM$3)
if %_erl%==7 (set "_oem=KMS38 [Windows] + Ohook [Office]" & set "para=/KMS38 /Ohook" &goto:Extract$OEM$3)
if %_erl%==6 (set "_oem=HWID [Windows] + 在线 KMS [Office]" & set "para=/HWID /KMS-ActAndRenewalTask /KMS-Office" &goto:Extract$OEM$3)
if %_erl%==5 (set "_oem=HWID [Windows] + Ohook [Office]" & set "para=/HWID /Ohook" &goto:Extract$OEM$3)
if %_erl%==4 (set "_oem=在线 KMS" & set "para=/KMS-ActAndRenewalTask /KMS-WindowsOffice" &goto:Extract$OEM$3)
if %_erl%==3 (set "_oem=KMS38" & set "para=/KMS38" &goto:Extract$OEM$3)
if %_erl%==2 (set "_oem=Ohook" & set "para=/Ohook" &goto:Extract$OEM$3)
if %_erl%==1 (set "_oem=HWID" & set "para=/HWID" &goto:Extract$OEM$3)
goto :Extract$OEM$2

::========================================================================================================================================

:Extract$OEM$3

cls
set "_dir=!_desktop_!\$OEM$\$$\Setup\Scripts"
md "!_dir!\"
copy /y /b "!_batf!" "!_dir!\MAS_AIO.cmd" %nul%

(
echo @echo off
echo fltmc ^>nul ^|^| exit /b
echo call "%%~dp0MAS_AIO.cmd" %para%
echo cd \
echo ^(goto^) 2^>nul ^& ^(if "%%~dp0"=="%%SystemRoot%%\Setup\Scripts\" rd /s /q "%%~dp0"^)
)>"!_dir!\SetupComplete.cmd"

set _error=
if not exist "!_dir!\MAS_AIO.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1

if defined _error (
%eline%
echo 解压 $OEM$ 文件夹到桌面上失败。
) else (
echo:
call :_color %Blue% "%_oem%"
call :_color %Green% "$OEM$ 文件夹已在桌面上成功创建。"
)
echo "%_oem%" | find /i "KMS38" 1>nul && (
echo:
echo 对于 KMS38 激活服务器 Cor/Acor 版本（无 GUI 版本），
echo 查看此页面 %mas%oem-folder
)
echo ___________________________________________________________________
echo:
call :_color %_Yellow% "请按任意键返回……"
pause >nul
goto Extras

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:HWIDActivation
@setlocal DisableDelayedExpansion
@echo off

::  若要激活，请使用“/HWID”参数运行脚本，或在以下行中将参数 0 更改为 1
set _act=0

::  要在当前版本不支持 HWID 激活时禁用更改版本，请将参数从 0 更改为 1，或使用“/HWID-NoEditionChange”参数运行脚本
set _NoEditionChange=0

::  如果在上面几行中更改了值或使用参数，脚本将会在无人值守模式下运行

::========================================================================================================================================

cls
color 07
title HWID 激活 %masver%

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="/HWID"                  set _act=1
if /i "%%A"=="/HWID-NoEditionChange"  set _NoEditionChange=1
if /i "%%A"=="-el"                    set _elev=1
)
)

for %%A in (%_act% %_NoEditionChange%) do (if "%%A"=="1" set _unattended=1)

::========================================================================================================================================

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

if %_NCS% EQU 1 (
for /F %%a in ('echo prompt $E ^| cmd') do set "esc=%%a"
set     "Red="41;97m""
set    "Gray="100;97m""
set   "Green="42;97m""
set    "Blue="44;97m""
set  "_White="40;37m""
set  "_Green="40;92m""
set "_Yellow="40;93m""
) else (
set     "Red="Red" "white""
set    "Gray="Darkgray" "white""
set   "Green="DarkGreen" "white""
set    "Blue="Blue" "white""
set  "_White="Black" "Gray""
set  "_Green="Black" "Green""
set "_Yellow="Black" "Yellow""
)

set "nceline=echo: &echo ==== 错误 ==== &echo:"
set "eline=echo: &call :dk_color %Red% "==== 错误 ====" &echo:"
if %~z0 GEQ 200000 (
set "_exitmsg=返回"
set "_fixmsg=请返回主菜单，选择疑难解答并运行修复许可选项。"
) else (
set "_exitmsg=退出"
set "_fixmsg=在 MAS 文件夹中，请运行疑难解答脚本并选择修复许可选项。"
)

::========================================================================================================================================

if %winbuild% LSS 10240 (
%eline%
echo 检测到不受支持的操作系统版本 [%winbuild%]。
echo HWID 激活仅支持 Windows 10/11。
echo 请使用在线 KMS 激活选项。
goto dk_done
)

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
%eline%
echo HWID 激活不支持 Windows Server。
echo 请使用 KMS38 或在线 KMS 选项。
goto dk_done
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

cls
mode 110, 34
if exist "%Systemdrive%\Windows\System32\spp\store_test\" mode 134, 34
title HWID 激活 %masver%

echo:
echo 正在初始化……

::  检查 PowerShell

%psc% $ExecutionContext.SessionState.LanguageMode %nul2% | find /i "Full" %nul1% || (
%eline%
%psc% $ExecutionContext.SessionState.LanguageMode
echo:
echo PowerShell 不可用，正在中止……
echo 如果你对Powershell施加了限制，请撤销这些更改。
echo:
echo 请查看此页面以获得帮助。 %mas%troubleshoot
goto dk_done
)

::========================================================================================================================================

call :dk_product
call :dk_ckeckwmic

::  显示潜在的脚本卡住情况的信息

sc start sppsvc %nul%
if %errorlevel% NEQ 1056 if %errorlevel% NEQ 0 (
echo:
echo 错误代码：%errorlevel%
call :dk_color %Red% "启动 [sppsvc] 服务失败，其余的进程可能需要很长时间……"
echo:
)

::========================================================================================================================================

::  检查系统是否已永久激活

call :dk_checkperm
if defined _perm (
cls
echo ___________________________________________________________________________________________
echo:
call :dk_color2 %_White% "     " %Green% "正在检查：%winos% 已永久激活。"
call :dk_color2 %_White% "     " %Gray% "不需要执行激活。"
echo ___________________________________________________________________________________________
if %_unattended%==1 goto dk_done
echo:
choice /C:10 /N /M ">    [1] 激活 [0] %_exitmsg% ："
if errorlevel 2 exit /b
)
cls

::========================================================================================================================================

::  检查评估版本

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul2% | find /i "Eval" %nul1% && (
%eline%
echo [%winos% ^| %winbuild%]
echo:
echo 评估版本无法激活。 
echo 你需要安装 %winos% 的完整版本。
echo:
echo 请从此处下载，
echo %mas%genuine-installation-media.html
goto dk_done
)
)

::========================================================================================================================================

call :dk_checksku

if not defined osSKU (
%eline%
echo 未正确检测到 SKU 值。正在中止……
goto dk_done
)

::========================================================================================================================================

set error=

cls
echo:
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PROCESSOR_ARCHITECTURE') do set arch=%%b
for /f "tokens=6-7 delims=[]. " %%i in ('ver') do if "%%j"=="" (set fullbuild=%%i) else (set fullbuild=%%i.%%j)
echo 正在检查操作系统信息                    [%winos% ^| %fullbuild% ^| %arch%]

::  检查 Internet 连接

set _int=
for %%a in (l.root-servers.net resolver1.opendns.com download.windowsupdate.com google.com) do if not defined _int (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (if not [%%#]==[] set _int=1)
)

if not defined _int (
%psc% "If([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet){Exit 0}Else{Exit 1}"
if !errorlevel!==0 (set _int=1&set ping_f= 但是Ping失败)
)

if defined _int (
echo 正在检查 Internet 连接                  [已连接%ping_f%]
) else (
set error=1
call :dk_color %Red% "正在检查 Internet 连接                  [未连接]"
)

::========================================================================================================================================

::  检查 Windows Script Host

set _WSH=1
reg query "HKCU\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled %nul2% | find /i "0x0" %nul1% && (set _WSH=0)
reg query "HKLM\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled %nul2% | find /i "0x0" %nul1% && (set _WSH=0)

if %_WSH% EQU 0 (
reg add "HKLM\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f %nul%
reg add "HKCU\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f %nul%
if not "%arch%"=="x86" reg add "HKLM\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f /reg:32 %nul%
echo 正在启用 Windows Script Host            [成功]
)

::========================================================================================================================================

echo 正在初始化诊断测试……

set "_serv=ClipSVC wlidsvc sppsvc KeyIso LicenseManager Winmgmt DoSvc UsoSvc CryptSvc BITS TrustedInstaller wuauserv"
if %winbuild% GEQ 17134 set "_serv=%_serv% WaaSMedicSvc"

::  Client License Service (ClipSVC)
::  Microsoft Account Sign-in Assistant
::  Software Protection
::  CNG Key Isolation
::  Windows License Manager Service
::  Windows Management Instrumentation
::  Delivery Optimization
::  Update Orchestrator Service
::  Cryptographic Services
::  Background Intelligent Transfer Service
::  Windows Modules Installer
::  Windows Update
::  Windows Update Medic Service

call :dk_errorcheck

::  检查 Windows 更新商店应用阻止程序

set updatesblock=
reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer /v SettingsPageVisibility %nul2% | find /i "windowsupdate" %nul% && set updatesblock=1
reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdateSysprepInProgress %nul% && set updatesblock=1
reg query HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate /s %nul2% | findstr /i "NoAutoUpdate DisableWindowsUpdateAccess" %nul% && set updatesblock=1

if defined updatesblock call :dk_color %Gray% "正在检查 Windows 更新 Blocker           [已找到]"

if defined applist echo: %serv_e% | find /i "wuauserv" %nul% && (
call :dk_color %Blue% "Windows 更新不起作用。如果你已禁用它，请启用它。"
reg query HKLM\SYSTEM\CurrentControlSet\Services\wuauserv /v WubLock %nul% && call :dk_color %Blue% "已使用Sordum Windows Update Blocker 工具阻止更新。"
)

reg query "HKLM\SOFTWARE\Policies\Microsoft\WindowsStore" /v DisableStoreApps %nul2% | find /i "0x1" %nul% && (
call :dk_color %Gray% "正在检查商店应用 Blocker                [已找到]"
)

::========================================================================================================================================

::  检查密钥

set key=
set altkey=
set skufound=
set changekey=
set altapplist=
set altedition=
set notworking=

if defined applist call :hwiddata key
if not defined key (
for /f "delims=" %%a in ('%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':getactivationid\:.*';iex ($f[1]);"') do (set altapplist=%%a)
if defined altapplist call :hwiddata key
)

if defined notworking call :hwidfallback
if not defined key call :hwidfallback

if defined altkey (set key=%altkey%&set changekey=1&set notworking=)

if defined notworking if defined notfoundaltactID (
call :dk_color %Red% "正在检查 HWID 的备用版本                [未找到 %altedition% 的激活 ID]"
)

if not defined key (
%eline%
echo [%winos% ^| %winbuild% ^| SKU：%osSKU%]
if not defined skufound (
echo 在支持的产品列表中找不到此产品。
) else (
echo %SystemRoot%\System32\spp\tokens\skus\ 中找不到所需的许可证文件
)
echo 请确保你使用的是此脚本的更新版本。
echo %mas%
echo:
goto dk_done
)

if defined notworking set error=1

::========================================================================================================================================

::  安装密钥

echo:
if defined changekey (
call :dk_color %Blue% "[%altedition%] 版本的产品密钥将会被用于启用 HWID 激活。"
echo:
)

if %_wmic% EQU 1 wmic path SoftwareLicensingService where __CLASS='SoftwareLicensingService' call InstallProductKey ProductKey="%key%" %nul%
if %_wmic% EQU 0 %psc% "(([WMISEARCHER]'SELECT Version FROM SoftwareLicensingService').Get()).InstallProductKey('%key%')" %nul%
if not %errorlevel%==0 cscript //nologo %windir%\system32\slmgr.vbs /ipk %key% %nul%
set errorcode=%errorlevel%
cmd /c exit /b %errorcode%
if %errorcode% NEQ 0 set "errorcode=[0x%=ExitCode%]"

if %errorcode% EQU 0 (
call :dk_refresh
echo 正在安装通用产品密钥                    [%key%] [成功]
) else (
call :dk_color %Red% "正在安装通用产品密钥                    [%key%] [失败] %errorcode%"
if not defined error (
if defined altapplist call :dk_color %Red% "无法找到此密钥的激活 ID。"
call :dk_color %Blue% "%_fixmsg%"
set showfix=1
)
set error=1
)

::========================================================================================================================================

::  将 Windows 区域更改为美国避免激活问题，因为 Windows 应用商店许可证在许多国家/地区不可用

for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Control Panel\International\Geo" /v Name %nul6%') do set "name=%%b"
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Control Panel\International\Geo" /v Nation %nul6%') do set "nation=%%b"

set regionchange=
if not "%name%"=="US" (
set regionchange=1
%psc% "Set-WinHomeLocation -GeoId 244" %nul%
if !errorlevel! EQU 0 (
echo 正在更改 Windows 区域为 USA             [成功]
) else (
call :dk_color %Red% "正在更改 Windows 区域为 USA             [失败]"
)
)

::========================================================================================================================================

::  生成 GenuineTicket.xml 并应用
::  应用票证的最正确方法是重新启动 ClipSVC 服务，但我们无法以此方式检查日志详细信息
::  为了获取日志详细信息并正确应用票证，脚本将安装票证 2 次（重新启动服务 + clipup -v -o）

if not exist %SystemRoot%\system32\ClipUp.exe (
call :dk_color %Red% "正在检查 ClipUp.exe 文件                [未找到，中止进程]"
call :dk_color2 %Blue% "查看此页面以获取帮助" %_Yellow% " %mas%troubleshoot"
goto :dl_final
)

set "tdir=%ProgramData%\Microsoft\Windows\ClipSVC\GenuineTicket"
if not exist "%tdir%\" md "%tdir%\" %nul%

if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%
if exist "%tdir%\*.xml" del /f /q "%tdir%\*.xml" %nul%
if exist "%ProgramData%\Microsoft\Windows\ClipSVC\Install\Migration\*" del /f /q "%ProgramData%\Microsoft\Windows\ClipSVC\Install\Migration\*" %nul%

call :hwiddata ticket

copy /y /b "%tdir%\GenuineTicket" "%tdir%\GenuineTicket.xml" %nul%

if not exist "%tdir%\GenuineTicket.xml" (
call :dk_color %Red% "正在生成 GenuineTicket.xml              [失败，中止进程]"
echo [%encoded%]
if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%
goto :dl_final
) else (
echo 正在生成 GenuineTicket.xml              [成功]
)

set "_xmlexist=if exist "%tdir%\GenuineTicket.xml""

%_xmlexist% (
%psc% Restart-Service ClipSVC %nul%
%_xmlexist% timeout /t 2 %nul%
%_xmlexist% timeout /t 2 %nul%

%_xmlexist% (
set error=1
if exist "%tdir%\*.xml" del /f /q "%tdir%\*.xml" %nul%
call :dk_color %Red% "正在安装 GenuineTicket.xml              [重新启动 ClipSVC 服务失败，正在等待……]"
)
)

copy /y /b "%tdir%\GenuineTicket" "%tdir%\GenuineTicket.xml" %nul%
clipup -v -o

set rebuildinfo=

if not exist %ProgramData%\Microsoft\Windows\ClipSVC\tokens.dat (
set error=1
set rebuildinfo=1
call :dk_color %Red% "正在检查 ClipSVC tokens.dat             [未找到]"
)

%_xmlexist% (
set error=1
set rebuildinfo=1
call :dk_color %Red% "正在安装 GenuineTicket.xml              [使用 clipup -v -o 失败]"
)

if exist "%ProgramData%\Microsoft\Windows\ClipSVC\Install\Migration\*.xml" (
set error=1
set rebuildinfo=1
call :dk_color %Red% "检查票证迁移                            [失败]"
)

if defined applist if not defined showfix if defined rebuildinfo (
set showfix=1
call :dk_color %Blue% "%_fixmsg%"
)

if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%

::==========================================================================================================================================

call :dk_product

echo:
echo 正在激活……

call :dk_act
call :dk_checkperm
if defined _perm (
echo:
call :dk_color %Green% "%winos% 已使用数字权利永久激活。"
goto :dl_final
)

::==========================================================================================================================================

::  扩展许可服务器测试，防止未找到的错误和激活失败

set resfail=
if not defined error (

ipconfig /flushdns %nul%
set "tls=$Tls12 = [Enum]::ToObject([System.Net.SecurityProtocolType], 3072); [System.Net.ServicePointManager]::SecurityProtocol = $Tls12;"

for %%# in (
login.live.com/ppsecure/deviceaddcredential.srf
purchase.mp.microsoft.com/v7.0/users/me/orders
) do if not defined resfail (
set "d1=Add-Type -AssemblyName System.Net.Http;"
set "d1=!d1! $client = [System.Net.Http.HttpClient]::new();"
set "d1=!d1! $response = $client.GetAsync('https://%%#').GetAwaiter().GetResult();"
set "d1=!d1! $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()"
%psc% "!tls! !d1!" %nul2% | findstr /i "PurchaseFD DeviceAddResponse" %nul1% || set resfail=1
)

if not defined resfail (
%psc% "!tls! irm https://licensing.mp.microsoft.com/v7.0/licenses/content -Method POST" | find /i "traceId" %nul1% || set resfail=1
)

if defined resfail (
set error=1
echo:
call :dk_color %Red% "正在检查许可服务器                      [连接失败]"
call :dk_color2 %Blue% "查看此页面以获取帮助" %_Yellow% " %mas%licensing-servers-issue"
)
)

::==========================================================================================================================================

::  清除与商店 ID 相关的注册表以修复激活，防止出现任何损坏

if not defined error (
echo:
set "_ident=HKU\S-1-5-19\SOFTWARE\Microsoft\IdentityCRL"
reg delete "!_ident!" /f %nul%
reg query "!_ident!" %nul% && (
call :dk_color %Red% "正在删除注册表                          [失败] [!_ident!]"
) || (
echo 正在删除注册表                          [成功] [!_ident!]
)

REM 刷新某些服务和许可证状态

for %%# in (wlidsvc LicenseManager sppsvc) do (%psc% Restart-Service %%# %nul%)
call :dk_refresh
call :dk_act
call :dk_checkperm
)

REM 检查与 Internet 相关的错误代码

if not defined error if not defined _perm (
echo "%error_code%" | findstr /i "0x80072e 0x80072f 0x800704cf" %nul% && (
set error=1
echo:
call :dk_color %Red% "正在检查 Internet 问题                  [已找到] %error_code%"
call :dk_color2 %Blue% "查看此页面以获取帮助" %_Yellow% " %mas%licensing-servers-issue"
)
)

::==========================================================================================================================================

echo:
if defined _perm (
call :dk_color %Green% "%winos% 已使用数字权利永久激活。"
) else (
call :dk_color %Red% "激活失败 %error_code%"
if defined notworking (
call :dk_color %Blue% "在编写此内容时，此产品不支持 HWID 激活。"
call :dk_color %Blue% "使用 KMS38 激活选项。"
) else (
if not defined error call :dk_color %Blue% "%_fixmsg%"
call :dk_color2 %Blue% "查看此页面以获取帮助" %_Yellow% " %mas%troubleshoot"
)
)

::========================================================================================================================================

:dl_final

echo:

if defined regionchange (
%psc% "Set-WinHomeLocation -GeoId %nation%" %nul%
if !errorlevel! EQU 0 (
echo 正在恢复 Windows 区域                   [成功]
) else (
call :dk_color %Red% "正在恢复 Windows 区域                   [失败] [%name% - %nation%]"
)
)

if %osSKU%==175 call :dk_color %Red% "%winos% 版本不支持在非 Azure 平台上激活。"

goto :dk_done

::========================================================================================================================================

::  检查 SKU 值

:dk_checksku

set osSKU=
set slcSKU=
set wmiSKU=
set regSKU=

if %winbuild% GEQ 14393 (set info=Kernel-BrandingInfo) else (set info=Kernel-ProductInfo)
set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('SLGetWindowsInformationDWORD', 'slc.dll', 'Public, Static', 1, [int], @([String], [int].MakeByRefType()), 1, 3);
set d1=%d1% $Sku = 0; [void]$TypeBuilder.CreateType()::SLGetWindowsInformationDWORD('%info%', [ref]$Sku); $Sku
for /f "delims=" %%s in ('"%psc% %d1%"') do if not errorlevel 1 (set slcSKU=%%s)
if "%slcSKU%"=="0" set slcSKU=
if 1%slcSKU% NEQ +1%slcSKU% set slcSKU=

for /f "tokens=3 delims=." %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions" /v OSProductPfn %nul6%') do set "regSKU=%%a"
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%a in ('"wmic Path Win32_OperatingSystem Get OperatingSystemSKU /format:LIST" %nul6%') do if not errorlevel 1 set "wmiSKU=%%a"
if %_wmic% EQU 0 for /f "tokens=1" %%a in ('%psc% "([WMI]'Win32_OperatingSystem=@').OperatingSystemSKU" %nul6%') do if not errorlevel 1 set "wmiSKU=%%a"

set osSKU=%slcSKU%
if not defined osSKU set osSKU=%wmiSKU%
if not defined osSKU set osSKU=%regSKU%
exit /b

::  检查 Windows 永久激活状态

:dk_checkperm

if %_wmic% EQU 1 wmic path SoftwareLicensingProduct where (LicenseStatus='1' and GracePeriodRemaining='0' and PartialProductKey is not NULL) get Name /value %nul2% | findstr /i "Windows" %nul1% && set _perm=1||set _perm=
if %_wmic% EQU 0 %psc% "(([WMISEARCHER]'SELECT Name FROM SoftwareLicensingProduct WHERE LicenseStatus=1 AND GracePeriodRemaining=0 AND PartialProductKey IS NOT NULL').Get()).Name | %% {echo ('Name='+$_)}" %nul2% | findstr /i "Windows" %nul1% && set _perm=1||set _perm=
exit /b

::  刷新许可证状态

:dk_refresh

if %_wmic% EQU 1 wmic path SoftwareLicensingService where __CLASS='SoftwareLicensingService' call RefreshLicenseStatus %nul%
if %_wmic% EQU 0 %psc% "$null=(([WMICLASS]'SoftwareLicensingService').GetInstances()).RefreshLicenseStatus()" %nul%
exit /b

::  激活命令

:dk_act

set error_code=
if %_wmic% EQU 1 wmic path SoftwareLicensingProduct where "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey<>null" call Activate %nul%
if %_wmic% EQU 0 %psc% "(([WMISEARCHER]'SELECT ID FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND PartialProductKey IS NOT NULL').Get()).Activate()" %nul%
if not %errorlevel%==0 cscript //nologo %windir%\system32\slmgr.vbs /ato %nul%
set error_code=%errorlevel%
cmd /c exit /b %error_code%
if %error_code% NEQ 0 (set "error_code=[错误代码：0x%=ExitCode%]") else (set error_code=)
exit /b

::  获取 Windows 激活 ID

:dk_actids

set applist=
if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f') get ID /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f''').Get()).ID ^| %% {echo ('ID='+$_)}" %nul6%')"
%chkapp% do (if defined applist (call set "applist=!applist! %%a") else (call set "applist=%%a"))
exit /b

::  从许可文件中获取激活 ID（如果未通过 WMI 找到）

:getactivationid:
$folderPath = "$env:windir\System32\spp\tokens\skus"
$files = Get-ChildItem -Path $folderPath -Recurse -Filter "*.xrm-ms"
$guids = @()
foreach ($file in $files) {
    $content = Get-Content -Path $file.FullName -Raw
    $matches = [regex]::Matches($content, 'name="productSkuId">\{([0-9a-fA-F\-]+)\}')
    foreach ($match in $matches) {
        $guids += $match.Groups[1].Value
    }
}
$guids = $guids | Select-Object -Unique
$guidsString = $guids -join " "
$guidsString
:getactivationid:

::  检查 wmic.exe

:dk_ckeckwmic

set _wmic=0
for %%# in (wmic.exe) do @if not "%%~$PATH:#"=="" (
wmic path Win32_ComputerSystem get CreationClassName /value %nul2% | find /i "computersystem" %nul1% && set _wmic=1
)
exit /b

::  获取产品名称（WMI/REG 方法并非在所有条件下都可靠，因此使用 winbrand.dll 方法）

:dk_product

call :dk_reflection

set d1=%ref% $meth = $TypeBuilder.DefinePInvokeMethod('BrandingFormatString', 'winbrand.dll', 'Public, Static', 1, [String], @([String]), 1, 3);
set d1=%d1% $meth.SetImplementationFlags(128); $TypeBuilder.CreateType()::BrandingFormatString('%%WINDOWS_LONG%%')

set winos=
for /f "delims=" %%s in ('"%psc% %d1%"') do if not errorlevel 1 (set winos=%%s)
echo "%winos%" | find /i "Windows" %nul1% || (
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName %nul6%') do set "winos=%%b"
if %winbuild% GEQ 22000 (
set winos=!winos:Windows 10=Windows 11!
)
)
exit /b

::  PowerShell 中使用反射代码的常见行

:dk_reflection

set ref=$AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1);
set ref=%ref% $ModuleBuilder = $AssemblyBuilder.DefineDynamicModule(2, $False);
set ref=%ref% $TypeBuilder = $ModuleBuilder.DefineType(0);
exit /b

::========================================================================================================================================

:dk_errorcheck

set showfix=

::  检查损坏的服务

set serv_cor=
for %%# in (%_serv%) do (
set _corrupt=
sc start %%# %nul%
if !errorlevel! EQU 1060 set _corrupt=1
sc query %%# %nul% || set _corrupt=1
for %%G in (DependOnService Description DisplayName ErrorControl ImagePath ObjectName Start Type) do if not defined _corrupt (
reg query HKLM\SYSTEM\CurrentControlSet\Services\%%# /v %%G %nul% || set _corrupt=1
if /i %%#==TrustedInstaller if /i %%G==DependOnService set _corrupt=
)

if defined _corrupt (if defined serv_cor (set "serv_cor=!serv_cor! %%#") else (set "serv_cor=%%#"))
)

if defined serv_cor (
set error=1
call :dk_color %Red% "正在检查损坏的服务                      [%serv_cor%]"
)

::========================================================================================================================================

::  检查禁用的服务

set serv_ste=
for %%# in (%_serv%) do (
sc start %%# %nul%
if !errorlevel! EQU 1058 (if defined serv_ste (set "serv_ste=!serv_ste! %%#") else (set "serv_ste=%%#"))
)

::  将禁用的服务启动类型更改为默认值

set serv_csts=
set serv_cste=

if defined serv_ste (
for %%# in (%serv_ste%) do (
if /i %%#==ClipSVC          (reg add "HKLM\SYSTEM\CurrentControlSet\Services\%%#" /v "Start" /t REG_DWORD /d "3" /f %nul% & sc config %%# start= demand %nul%)
if /i %%#==wlidsvc          sc config %%# start= demand %nul%
if /i %%#==sppsvc           (reg add "HKLM\SYSTEM\CurrentControlSet\Services\%%#" /v "Start" /t REG_DWORD /d "2" /f %nul% & sc config %%# start= delayed-auto %nul%)
if /i %%#==KeyIso           sc config %%# start= demand %nul%
if /i %%#==LicenseManager   sc config %%# start= demand %nul%
if /i %%#==Winmgmt          sc config %%# start= auto %nul%
if /i %%#==DoSvc            sc config %%# start= delayed-auto %nul%
if /i %%#==UsoSvc           sc config %%# start= delayed-auto %nul%
if /i %%#==CryptSvc         sc config %%# start= auto %nul%
if /i %%#==BITS             sc config %%# start= delayed-auto %nul%
if /i %%#==wuauserv         sc config %%# start= demand %nul%
if /i %%#==WaaSMedicSvc     sc config %%# start= demand %nul%
if !errorlevel!==0 (
if defined serv_csts (set "serv_csts=!serv_csts! %%#") else (set "serv_csts=%%#")
) else (
if defined serv_cste (set "serv_cste=!serv_cste! %%#") else (set "serv_cste=%%#")
)
)
)

if defined serv_csts call :dk_color %Gray% "正在启用禁用的服务                      [成功] [%serv_csts%]"

if defined serv_cste (
set error=1
call :dk_color %Red% "正在启用禁用的服务                      [失败] [%serv_cste%]"
)

::========================================================================================================================================

::  检查服务是否能够运行
::  添加获取正确状态和错误代码的解决方法，因为 sc 查询在某些情况下不会输出正确的结果

set serv_e=
for %%# in (%_serv%) do (
set errorcode=
set checkerror=

sc query %%# | find /i "RUNNING" %nul% || (
%psc% Start-Service %%# %nul%
set errorcode=!errorlevel!
sc query %%# | find /i "RUNNING" %nul% || set checkerror=1
)

sc start %%# %nul%
if !errorlevel! NEQ 1056 if !errorlevel! NEQ 0 (set errorcode=!errorlevel!&set checkerror=1)
if defined checkerror if defined serv_e (set "serv_e=!serv_e!, %%#-!errorcode!") else (set "serv_e=%%#-!errorcode!")
)

if defined serv_e (
set error=1
call :dk_color %Red% "正在启动服务                            [失败] [%serv_e%]"
echo %serv_e% | findstr /i "ClipSVC-1058 sppsvc-1058" %nul% && (
call :dk_color %Blue% "请重新启动系统修复已禁用服务错误 1058。"
set showfix=1
)
)

::========================================================================================================================================

::  各类错误检查

if defined safeboot_option (
set error=1
set showfix=1
call :dk_color2 %Red% "正在检查引导模式                        [%safeboot_option%] " %Blue% "[系统正在安全模式下运行。将以正常模式运行。]"
)


for /f "skip=2 tokens=2*" %%A in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State" /v ImageState') do (set imagestate=%%B)
if /i not "%imagestate%"=="IMAGE_STATE_COMPLETE" (
set error=1
call :dk_color %Red% "检查 Windows 安装状态                   [%imagestate%]"
echo "%imagestate%" | find /i "RESEAL" %nul% && (
set showfix=1
call :dk_color %Blue% "你需要它在正常模式下运行,以防你在审计模式运行。"
)
)


reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WinPE" /v InstRoot %nul% && (
set error=1
set showfix=1
call :dk_color2 %Red% "正在检查 WinPE                          " %Blue% "[系统正在 WinPE 模式下运行。将以正常模式运行。]"
)


set wpainfo=
set wpaerror=
for /f "delims=" %%a in ('%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':wpatest\:.*';iex ($f[1]);" %nul6%') do (set wpainfo=%%a)
echo "%wpainfo%" | find /i "Error Found" %nul% && (
set error=1
set wpaerror=1
call :dk_color %Red% "正在检查WPA注册表错误                   [%wpainfo%]"
) || (
echo 正在检查WPA注册表总数                   [%wpainfo%]
)


DISM /English /Online /Get-CurrentEdition %nul%
set dism_error=%errorlevel%
cmd /c exit /b %dism_error%
if %dism_error% NEQ 0 set "dism_error=0x%=ExitCode%"
if %dism_error% NEQ 0 (
call :dk_color %Red% "正在检查 DISM                           [未响应] [%dism_error%]"
)


if not defined officeact if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" (
set error=1
set showfix=1
call :dk_color %Red% "正在检查评估程序包                      [非评估许可证被安装在 Windows 评估版本中]"
call :dk_color %Blue% "无法激活 Windows 评估版本，不同的许可证安装可能会导致错误。"
call :dk_color %Blue% "推荐安装 %winos% 的完整版本。"
call :dk_color %Blue% "你可以从 %mas%genuine-installation-media.html 下载"
)


set osedition=
for /f "skip=2 tokens=3" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul6%') do set "osedition=%%a"

::  解决了在1607至1709版本中将专业教育版显示为专业版的问题。

if "%osSKU%"=="164" set osedition=ProfessionalEducation
if "%osSKU%"=="165" set osedition=ProfessionalEducationN

if not defined officeact (
if not defined osedition (
call :dk_color %Red% "检查 Windows 版本名称                   [注册表中未找到]"
) else (

if not exist "%SystemRoot%\System32\spp\tokens\skus\%osedition%\%osedition%*.xrm-ms" (
set error=1
call :dk_color %Red% "检查许可证文件                          [未找到] [%osedition%]"
)

if not exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*-%osedition%-*.mum" (
set error=1
call :dk_color %Red% "检查包文件                              [未找到] [%osedition%]"
)
)
)


cscript //nologo %windir%\system32\slmgr.vbs /dlv %nul%
set error_code=%errorlevel%
cmd /c exit /b %error_code%
if %error_code% NEQ 0 set "error_code=0x%=ExitCode%"
if %error_code% NEQ 0 (
set error=1
call :dk_color %Red% "正在检查 slmgr /dlv                     [未响应] %error_code%"
)


for %%# in (wmic.exe) do @if "%%~$PATH:#"=="" (
call :dk_color %Gray% "正在检查 WMIC.exe                       [未找到]"
)


set wmifailed=
if %_wmic% EQU 1 wmic path Win32_ComputerSystem get CreationClassName /value %nul2% | find /i "computersystem" %nul1%
if %_wmic% EQU 0 %psc% "Get-CIMInstance -Class Win32_ComputerSystem | Select-Object -Property CreationClassName" %nul2% | find /i "computersystem" %nul1%

if %errorlevel% NEQ 0 set wmifailed=1
echo "%error_code%" | findstr /i "0x800410 0x800440" %nul1% && set wmifailed=1& ::  https://learn.microsoft.com/en-us/windows/win32/wmisdk/wmi-error-constants
if defined wmifailed (
set error=1
call :dk_color %Red% "正在检查 WMI                            [未响应]"
call :dk_color %Blue% "在 MAS 中，请转到疑难解答并运行修复 WMI 选项。"
set showfix=1
)


%nul% set /a "sum=%slcSKU%+%regSKU%+%wmiSKU%"
set /a "sum/=3"
if not defined officeact if not "%sum%"=="%slcSKU%" (
call :dk_color %Red% 正在检查 SLC/WMI/REG SKU                [发现差异 - SLC:%slcSKU% WMI:%wmiSKU% Reg:%regSKU%]"
)


reg query "HKU\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\PersistedTSReArmed" %nul% && (
set error=1
set showfix=1
call :dk_color2 %Red% "正在检查 Rearm                          " %Blue% "[需要重新启动系统]"
)


reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState" %nul% && (
set error=1
set showfix=1
call :dk_color2 %Red% "正在检查 ClipSVC                        " %Blue% "[需要重新启动系统]"
)


for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "SkipRearm" %nul6%') do if /i %%b NEQ 0x0 (
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "SkipRearm" /t REG_DWORD /d "0" /f %nul%
call :dk_color %Red% "正在检查 SkipRearm                      [默认值 0 未找到，更改为 0]"
%psc% Restart-Service sppsvc %nul%
set error=1
)


reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Plugins\Objects\msft:rm/algorithm/hwid/4.0" /f ba02fed39662 /d %nul% || (
call :dk_color %Red% "正在检查 SPP 注册表键值                 [已找到不正确的模块 ID]"
call :dk_color %Blue% "可能是由 Gaming Spoofer 引起的。帮助：%mas%troubleshoot"
set error=1
set showfix=1
)


set tokenstore=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v TokenStore %nul6%') do call set "tokenstore=%%b"
if not exist "%tokenstore%\" (
set error=1
REM 此代码仅在缺少令牌文件夹时创建令牌文件夹，并为其设置默认权限
mkdir "%tokenstore%" %nul%
set "d=$sddl = 'O:BAG:BAD:PAI(A;OICI;FA;;;SY)(A;OICI;FA;;;BA)(A;OICIIO;GR;;;BU)(A;;FR;;;BU)(A;OICI;FA;;;S-1-5-80-123231216-2592883651-3715271367-3753151631-4175906628)';"
set "d=!d! $AclObject = New-Object System.Security.AccessControl.DirectorySecurity;"
set "d=!d! $AclObject.SetSecurityDescriptorSddlForm($sddl);"
set "d=!d! Set-Acl -Path %tokenstore% -AclObject $AclObject;"
%psc% "!d!" %nul%
call :dk_color %Gray% "正在检查 SPP Token 文件夹               [未找到。立即创建] [%tokenstore%\]"
)


call :dk_actids
if not defined applist (
%psc% Stop-Service sppsvc %nul%
cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
if !errorlevel! NEQ 0 cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
call :dk_refresh
call :dk_actids
if not defined applist (
set error=1
call :dk_color %Red% "正在检查激活 ID                         [未找到]"
)
)


if exist "%tokenstore%\" if not exist "%tokenstore%\tokens.dat" (
set error=1
call :dk_color %Red% "正在检查 SPP tokens.dat                 [未找到] [%tokenstore%\]"
)


if not exist %SystemRoot%\system32\sppsvc.exe (
set error=1
set showfix=1
call :dk_color %Red% "正在检查 sppsvc.exe 文件                [未找到]"
)


::  这段代码检查NT SERVICE\sppsvc 是否有权限访问令牌文件夹和所需的注册表项。这通常是由 Gaming Spoofer 引起的。

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
if defined permerror (
set error=1
set showfix=1
call :dk_color %Red% "正在检查 SPP 权限                       [发现错误]"
call :dk_color %Blue% "%_fixmsg%"
)

::  如果所需的服务未被禁用或损坏 + 是否有任何错误 + slmgr /dlv 错误级别不为零 + 之前未显示修复程序，则执行以下检查

if not defined serv_cor if not defined serv_cste if defined error if /i not %error_code%==0 if not defined showfix (
set showfix=1
call :dk_color %Blue% "%_fixmsg%"
if not defined permerror call :dk_color %Blue% "如果激活仍然失败，之后运行修复 WPA 注册表选项。"
)

if not defined showfix if defined wpaerror (
set showfix=1
call :dk_color %Blue% "如果激活失败，请返回主菜单，选择“故障排除”并运行“修复WPA注册表”选项。"
)

exit /b

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

:dk_color

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[0m
) else (
%psc% write-host -back '%1' -fore '%2' '%3'
)
exit /b

:dk_color2

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[%~3%~4%esc%[0m
) else (
%psc% write-host -back '%1' -fore '%2' '%3' -NoNewline; write-host -back '%4' -fore '%5' '%6'
)
exit /b

::========================================================================================================================================

:dk_done

echo:
if %_unattended%==1 timeout /t 2 & exit /b
call :dk_color %_Yellow% "请按任意键%_exitmsg%脚本……"
pause %nul1%
exit /b

::========================================================================================================================================

::  第 1 列 = 激活 ID
::  第 2 列 = 通用零售/OEM/MAK 密钥
::  第 3 列 = SKU ID
::  第 4 列 = 密钥部分编号
::  第 5 列 = 票证签名值。它就是这样，它没有被编码。（请查阅 mass grave[.]dev/hwid.html#Manual_Activation 了解它是如何生成的）
::  第 6 列 = 1 = 激活无效（在撰写本文时）、0 = 激活有效
::  第 7 列 = 密钥类型
::  第 8 列 = WMI 版本 ID（仅供参考）
::  第 9 列 = 版本名称，以防止相同的版本 ID 用于具有不同密钥的不同操作系统版本
::  分隔符  = _


:hwiddata

set f=
for %%# in (
8b351c9c-f398-4515-9900-09df49427262_XG%f%VPP-NMH%f%47-7T%f%THJ-W3F%f%W7-8H%f%V2C___4_X19-99683_HGNKjkKcKQHO6n8srMUrDh/MElffBZarLqCMD9rWtgFKf3YzYOLDPEMGhuO/auNMKCeiU7ebFbQALS/MyZ7TvidMQ2dvzXeXXKzPBjfwQx549WJUU7qAQ9Txg9cR9SAT8b12Pry2iBk+nZWD9VtHK3kOnEYkvp5WTCTsrSi6Re4_0_OEM:NONSLP_Enterprise
c83cef07-6b72-4bbc-a28f-a00386872839_3V%f%6Q6-NQX%f%CX-V8%f%YXR-9QC%f%YV-QP%f%FCT__27_X19-98746_NHn2n0N1UfVf00CfaI5LCDMDsKdVAWpD/HAfUrcTAKsw9d2Sks4h5MhyH/WUx+B6dFi8ol7D3AHorR8y9dqVS1Bd2FdZNJl/tTR1PGwYn6KL88NS19aHmFNdX8s4438vaa+Ty8Qk8EDcwm/wscC8lQmi3/RgUKYdyGFvpbGSVlk_0_Volume:MAK_EnterpriseN
4de7cb65-cdf1-4de9-8ae8-e3cce27b9f2c_VK%f%7JG-NPH%f%TM-C9%f%7JM-9MP%f%GT-3V%f%66T__48_X19-98841_Yl/jNfxJ1SnaIZCIZ4m6Pf3ySNoQXifNeqfltNaNctx+onwiivOx7qcSn8dFtURzgMzSOFnsRQzb5IrvuqHoxWWl1S3JIQn56FvKsvSx7aFXIX3+2Q98G1amPV/WEQ0uHA5d7Ya6An+g0Z0zRP7evGoomTs4YuweaWiZQjQzSpA_0_____Retail_Professional
9fbaf5d6-4d83-4422-870d-fdda6e5858aa_2B%f%87N-8KF%f%HP-DK%f%V6R-Y2C%f%8J-PK%f%CKT__49_X19-98859_Ge0mRQbW8ALk7T09V+1k1yg66qoS0lhkgPIROOIOgxKmWPAvsiLAYPKDqM4+neFCA/qf1dHFmdh0VUrwFBPYsK251UeWuElj4bZFVISL6gUt1eZwbGfv5eurQ0i+qZiFv+CcQOEFsd5DD4Up6xPLLQS3nAXODL5rSrn2sHRoCVY_0_____Retail_ProfessionalN
f742e4ff-909d-4fe9-aacb-3231d24a0c58_4C%f%PRK-NM3%f%K3-X6%f%XXQ-RXX%f%86-WX%f%CHW__98_X19-98877_vel4ytVtnE8FhvN87Cflz9sbh5QwHD1YGOeej9QP7hF3vlBR4EX2/S/09gRneeXVbQnjDOCd2KFMKRUWHLM7ZhFBk8AtlG+kvUawPZ+CIrwrD3mhi7NMv8UX/xkLK3HnBupMEuEwsMJgCUD8Pn6om1mEiQebHBAqu4cT7GN9Y0g_0_____Retail_CoreN
1d1bac85-7365-4fea-949a-96978ec91ae0_N2%f%434-X9D%f%7W-8P%f%F6X-8DV%f%9T-8T%f%YMD__99_X19-99652_Nv17eUTrr1TmUX6frlI7V69VR6yWb7alppCFJPcdjfI+xX4/Cf2np3zm7jmC+zxFb9nELUs477/ydw2KCCXFfM53bKpBQZKHE5+MdGJGxebOCcOtJ3hrkDJtwlVxTQmUgk5xnlmpk8PHg82M2uM5B7UsGLxGKK4d3hi0voSyKeI_0_____Retail_CoreCountrySpecific
3ae2cc14-ab2d-41f4-972f-5e20142771dc_BT%f%79Q-G7N%f%6G-PG%f%BYW-4YW%f%X6-6F%f%4BT_100_X19-99661_FV2Eao/R5v8sGrfQeOjQ4daokVlNOlqRCDZXuaC45bQd5PsNU3t1b4AwWeYM8TAwbHauzr4tPG0UlsUqUikCZHy0poROx35bBBMBym6Zbm9wDBVyi7nCzBtwS86eOonQ3cU6WfZxhZRze0POdR33G3QTNPrnVIM2gf6nZJYqDOA_0_____Retail_CoreSingleLanguage
2b1f36bb-c1cd-4306-bf5c-a0367c2d97d8_YT%f%MG3-N6D%f%KC-DK%f%B77-7M9%f%GH-8H%f%VX7_101_X19-98868_GH/jwFxIcdQhNxJIlFka8c1H48PF0y7TgJwaryAUzqSKXynONLw7MVciDJFVXTkCjbXSdxLSWpPIC50/xyy1rAf8aC7WuN/9cRNAvtFPC1IVAJaMeq1vf4mCqRrrxJQP6ZEcuAeHFzLe/LLovGWCd8rrs6BbBwJXCvAqXImvycQ_0_____Retail_Core
2a6137f3-75c0-4f26-8e3e-d83d802865a4_XK%f%CNC-J26%f%Q9-KF%f%HD2-FKT%f%HY-KD%f%72Y_119_X19-99606_hci78IRWDLBtdbnAIKLDgV9whYgtHc1uYyp9y6FszE9wZBD5Nc8CUD2pI2s2RRd3M04C4O7M3tisB3Ov/XVjpAbxlX3MWfUR5w4MH0AphbuQX0p5MuHEDYyfqlRgBBRzOKePF06qfYvPQMuEfDpKCKFwNojQxBV8O0Arf5zmrIw_0_OEM:NONSLP_PPIPro
e558417a-5123-4f6f-91e7-385c1c7ca9d4_YN%f%MGQ-8RY%f%V3-4P%f%GQ3-C8X%f%TP-7C%f%FBY_121_X19-98886_x9tPFDZmjZMf29zFeHV5SHbXj8Wd8YAcCn/0hbpLcId4D7OWqkQKXxXHIegRlwcWjtII0sZ6WYB0HQV2KH3LvYRnWKpJ5SxeOgdzBIJ6fhegYGGyiXsBv9sEb3/zidPU6ZK9LugVGAcRZ6HQOiXyOw+Yf5H35iM+2oDZXSpjvJw_0_____Retail_Education
c5198a66-e435-4432-89cf-ec777c9d0352_84%f%NGF-MHB%f%T6-FX%f%BX8-QWJ%f%K7-DR%f%R8H_122_X19-98892_jkL4YZkmBCJtvL1fT30ZPBcjmzshBSxjwrE0Q00AZ1hYnhrH+npzo1MPCT6ZRHw19ZLTz7wzyBb0qqcBVbtEjZW0Xs2MYLxgriyoONkhnPE6KSUJBw7C0enFVLHEqnVu/nkaOFfockN3bc+Eouw6W2lmHjklPHc9c6Clo04jul0_0_____Retail_EducationN
f6e29426-a256-4316-88bf-cc5b0f95ec0c_PJ%f%B47-8PN%f%2T-MC%f%GDY-JTY%f%3D-CB%f%CPV_125_X23-50331_OPGhsyx+Ctw7w/KLMRNrY+fNBmKPjUG0R9RqkWk4e8ez+ExSJxSLLex5WhO5QSNgXLmEra+cCsN6C638aLjIdH2/L7D+8z/C6EDgRvbHMmidHg1lX3/O8lv0JudHkGtHJYewjorn/xXGY++vOCTQdZNk6qzEgmYSvPehKfdg8js_1_Volume:MAK_EnterpriseS_Ge
cce9d2de-98ee-4ce2-8113-222620c64a27_KC%f%NVH-YKW%f%X8-GJ%f%JB9-H9F%f%DT-6F%f%7W2_125_X22-66075_GCqWmJOsTVun9z4QkE9n2XqBvt3ZWSPl9QmIh9Q2mXMG/QVt2IE7S+ES/NWlyTSNjLVySr1D2sGjxgEzy9kLwn7VENQVJ736h1iOdMj/3rdqLMSpTa813+nPSQgKpqJ3uMuvIvRP0FdB7Y4qt8qf9kNKK25A1QknioD/6YubL/4_1_Volume:MAK_EnterpriseS_VB
d06934ee-5448-4fd1-964a-cd077618aa06_43%f%TBQ-NH9%f%2J-XK%f%TM7-KT3%f%KK-P3%f%9PB_125_X21-83233_EpB6qOCo8pRgO5kL4vxEHck2J1vxyd9OqvxUenDnYO9AkcGWat/D74ZcFg5SFlIya1U8l5zv+tsvZ4wAvQ1IaFW1PwOKJLOaGgejqZ41TIMdFGGw+G+s1RHsEnrWr3UOakTodby1aIMUMoqf3NdaM5aWFo8fOmqWC5/LnCoighs_0_OEM:NONSLP_EnterpriseS_RS5
706e0cfd-23f4-43bb-a9af-1a492b9f1302_NK%f%96Y-D9C%f%D8-W4%f%4CQ-R8Y%f%TK-DY%f%JWX_125_X21-05035_ntcKmazIvLpZOryft28gWBHu1nHSbR+Gp143f/BiVe+BD2UjHBZfSR1q405xmQZsygz6VRK6+zm8FPR++71pkmArgCLhodCQJ5I4m7rAJNw/YX99pILphi1yCRcvHsOTGa825GUVXgf530tHT6hr0HQ1lGeGgG1hPekpqqBbTlg_0_OEM:NONSLP_EnterpriseS_RS1
faa57748-75c8-40a2-b851-71ce92aa8b45_FW%f%N7H-PF9%f%3Q-4G%f%GP8-M8R%f%F3-MD%f%WWW_125_X19-99617_Fe9CDClilrAmwwT7Yhfx67GafWRQEpwyj8R+a4eaTqbpPcAt7d1hv1rx8Sa9AzopEGxIrb7IhiPoDZs0XaT1HN0/olJJ/MnD73CfBP4sdQdLTsSJE3dKMWYTQHpnjqRaS/pNBYRr8l9Mv8yfcP8uS2MjIQ1cRTqRmC7WMpShyCg_0_OEM:NONSLP_EnterpriseS_TH
837766ff-61c5-427d-87c3-a2acbd44767a_XF%f%C77-XNR%f%XM-2Q%f%36W-FCM%f%9T-YH%f%DJ9_126_X23-50304_h6V6Q4DL/hlvcD3GyVxrVfP1BEL4a5TdyNCMlbq/OZnky/HowuRAcHMpN59fwqLS98+7WEDooWCrxriXcATwo0fwOGs/fEfP/Pa5SKP+Xnng1eoPm1PkjuZaqA8p2dPQv32wJ0u3QW7VMQM9BzzpyqtNAsqNS/wl7vfN7tyLbDo_1_Volume:MAK_EnterpriseSN_Ge
2c060131-0e43-4e01-adc1-cf5ad1100da8_RQ%f%FNW-9TP%f%M3-JQ%f%73T-QV4%f%VQ-DV%f%9PT_126_X22-66108_w/HFPDNCz4EogszDYZ8xUJh8aylfpgh6gzm9k8JSteprY5UumLc5n6KUwiSE3/5NaiI9gZ3xmTJq+g1OSPsdGwhuA+8LA2pQhA+wU8VO/ZaYxe1T4WF6oip/c0n6xA1sx/mWYNwd/WBDJpslTw5NRNLc5wWh0FV5RtxCaXE07lM_1_Volume:MAK_EnterpriseSN_VB
e8f74caa-03fb-4839-8bcc-2e442b317e53_M3%f%3WV-NHY%f%3C-R7%f%FPM-BQG%f%PT-23%f%9PG_126_X21-83264_Fl7tjifybEI9hArxMVFKqIqmI6mrCZy4EtJyVjpo2eSfeMTBli55+E0i2AaPfE2FJknUig7HuiNC1Pu2IWZcj5ShVFQEKPY6K//RucX8oPQfh0zK5r1aNJNvV4gMlqvOyGD8sXttLBZv8wg1w/++cNk/z38DE2shiDf7LYnK4w0_1_Volume:MAK_EnterpriseSN_RS5
3d1022d8-969f-4222-b54b-327f5a5af4c9_2D%f%BW3-N2P%f%JG-MV%f%HW3-G7T%f%DK-9H%f%KR4_126_X21-04921_zLPNvcl1iqOefy0VLg+WZgNtRNhuGpn8+BFKjMqjaNOSKiuDcR6GNDS5FF1Aqk6/e6shJ+ohKzuwrnmYq3iNQ3I2MBlYjM5kuNfKs8Vl9dCjSpQr//GBGps6HtF2xrG/2g/yhtYC7FbtGDIE16uOeNKFcVg+XMb0qHE/5Etyfd8_0_Volume:MAK_EnterpriseSN_RS1
60c243e1-f90b-4a1b-ba89-387294948fb6_NT%f%X6B-BRY%f%C2-K6%f%786-F6M%f%VQ-M7%f%V2X_126_X19-98770_kbXfe0z9Vi1S0yfxMWzI5+UtWsJKzxs7wLGUDLjrckFDn1bDQb4MvvuCK1w+Qrq33lemiGpNDspa+ehXiYEeSPFcCvUBpoMlGBFfzurNCHWiv3o1k3jBoawJr/VoDoVZfxhkps0fVoubf9oy6C6AgrkZ7PjCaS58edMcaUWvYYg_0_Volume:MAK_EnterpriseSN_TH
01eb852c-424d-4060-94b8-c10d799d7364_3X%f%P6D-CRN%f%D4-DR%f%YM2-GM8%f%4D-4G%f%G8Y_139_X23-37869_PVW0XnRJnsWYjTqxb6StCi2tge/uUwegjdiFaFUiZpwdJ620RK+MIAsSq5S+egXXzIWNntoy2fB6BO8F1wBFmxP/mm/3rn5C33jtF5QrbNqY7X9HMbqSiC7zhs4v4u2Xa4oZQx8JQkwr8Q2c/NgHrOJKKRASsSckhunxZ+WVEuM_1_____Retail_ProfessionalCountrySpecific_Zn
eb6d346f-1c60-4643-b960-40ec31596c45_DX%f%G7C-N36%f%C4-C4%f%HTG-X4T%f%3X-2Y%f%V77_161_X21-43626_MaVqTkRrGnOqYizl15whCOKWzx01+BZTVAalvEuHXM+WV55jnIfhWmd/u1GqCd5OplqXdU959zmipK2Iwgu2nw/g91nW//sQiN/cUcvg1Lxo6pC3gAo1AjTpHmGIIf9XlZMYlD+Vl6gXsi/Auwh3yrSSFh5s7gOczZoDTqQwHXA_0_____Retail_ProfessionalWorkstation
89e87510-ba92-45f6-8329-3afa905e3e83_WY%f%PNQ-8C4%f%67-V2%f%W6J-TX4%f%WX-WT%f%2RQ_162_X21-43644_JVGQowLiCcPtGY9ndbBDV+rTu/q5ljmQTwQWZgBIQsrAeQjLD8jLEk/qse7riZ7tMT6PKFVNXeWqF7PhLAmACbE8O3Lvp65XMd/Oml9Daynj5/4n7unsffFHIHH8TGyO5j7xb4dkFNqC5TX3P8/1gQEkTIdZEOTQQXFu0L2SP5c_0_____Retail_ProfessionalWorkstationN
62f0c100-9c53-4e02-b886-a3528ddfe7f6_8P%f%TT6-RNW%f%4C-6V%f%7J2-C2D%f%3X-MH%f%BPB_164_X21-04955_CEDgxI8f/fxMBiwmeXw5Of55DG32sbGALzHihXkdbYTDaE3pY37oAA4zwGHALzAFN/t254QImGPYR6hATgl+Cp804f7serJqiLeXY965Zy67I4CKIMBm49lzHLFJeDnVTjDB0wVyN29pvgO3+HLhZ22KYCpkRHFFMy2OKxS68Yc_0_____Retail_ProfessionalEducation
13a38698-4a49-4b9e-8e83-98fe51110953_GJ%f%TYN-HDM%f%QY-FR%f%R76-HVG%f%C7-QP%f%F8P_165_X21-04956_r35zp9OfxKSBcTxKWon3zFtbOiCufAPo6xRGY5DJqCRFKdB0jgZalNQitvjmaZ/Rlez2vjRJnEart4LrvyW4d9rrukAjR3+c3UkeTKwoD3qBl9AdRJbXCa2BdsoXJs1WVS4w4LuVzpB/SZDuggZt0F2DlMB427F5aflook/n1pY_0_____Retail_ProfessionalEducationN
df96023b-dcd9-4be2-afa0-c6c871159ebe_NJ%f%CF7-PW8%f%QT-33%f%24D-688%f%JX-2Y%f%V66_175_X21-41295_rVpetYUmiRB48YJfCvJHiaZapJ0bO8gQDRoql+rq5IobiSRu//efV1VXqVpBkwILQRKgKIVONSTUF5y2TSxlDLbDSPKp7UHfbz17g6vRKLwOameYEz0ZcK3NTbApN/cMljHvvF/mBag1+sHjWu+eoFzk8H89k9nw8LMeVOPJRDc_0_____Retail_ServerRdsh
d4ef7282-3d2c-4cf0-9976-8854e64a8d1e_V3%f%WVW-N2P%f%V2-CG%f%WC3-34Q%f%GF-VM%f%J2C_178_X21-32983_Xzme9hDZR6H0Yx0deURVdE6LiTOkVqWng5W/OTbkxRc0rq+mSYpo/f/yqhtwYlrkBPWx16Yok5Bvcb34vbKHvEAtxfYp4te20uexLzVOtBcoeEozARv4W/6MhYfl+llZtR5efsktj4N4/G4sVbuGvZ9nzNfQO9TwV6NGgGEj2Ec_0_____Retail_Cloud
af5c9381-9240-417d-8d35-eb40cd03e484_NH%f%9J3-68W%f%K7-6F%f%B93-4K3%f%DF-DJ%f%4F6_179_X21-32987_QGRDZOU/VZhYLOSdp2xDnFs8HInNZctcQlWCIrORVnxTQr55IJwN4vK3PJHjkfRLQ/bgUrcEIhyFbANqZFUq8yD1YNubb2bjNORgI/m8u85O9V7nDGtxzO/viEBSWyEHnrzLKKWYqkRQKbbSW3ungaZR0Ti5O2mAUI4HzAFej50_0_____Retail_CloudN
8ab9bdd1-1f67-4997-82d9-8878520837d9_XQ%f%QYW-NFF%f%MW-XJ%f%PBH-K87%f%32-CK%f%FFD_188_X21-99378_djy0od0uuKd2rrIl+V1/2+MeRltNgW7FEeTNQsPMkVSL75NBphgoso4uS0JPv2D7Y1iEEvmVq6G842Kyt52QOwXgFWmP/IQ6Sq1dr+fHK/4Et7bEPrrGBEZoCfWqk0kdcZRPBij2KN6qCRWhrk1hX2g+U40smx/EYCLGh9HCi24_0_____OEM:DM_IoTEnterprise
ed655016-a9e8-4434-95d9-4345352c2552_QP%f%M6N-7J2%f%WJ-P8%f%8HH-P3Y%f%RH-YY%f%74H_191_X21-99682_qHs/PzfhYWdtSys2edzcz4h+Qs8aDqb8BIiQ/mJ/+0uyoJh1fitbRCIgiFh2WAGZXjdgB8hZeheNwHibd8ChXaXg4u+0XlOdFlaDTgTXblji8fjETzDBk9aGkeMCvyVXRuUYhTSdp83IqGHz7XuLwN2p/6AUArx9JZCoLGV8j3w_0_OEM:NONSLP_IoTEnterpriseS_VB
6c4de1b8-24bb-4c17-9a77-7b939414c298_CG%f%K42-GYN%f%6Y-VD%f%22B-BX9%f%8W-J8%f%JXD_191_X23-12617_J/fpIRynsVQXbp4qZNKp6RvOgZ/P2klILUKQguMlcwrBZybwNkHg/kM5LNOF/aDzEktbPnLnX40GEvKkYT6/qP4cMhn/SOY0/hYOkIdR34ilzNlVNq5xP7CMjCjaUYJe+6ydHPK6FpOuEoWOYYP5BZENKNGyBy4w4shkMAw19mA_0_OEM:NONSLP_IoTEnterpriseS_Ge
d4bdc678-0a4b-4a32-a5b3-aaa24c3b0f24_K9%f%VKN-3BG%f%WV-Y6%f%24W-MCR%f%MQ-BH%f%DCD_202_X22-53884_kyoNx2s93U6OUSklB1xn+GXcwCJO1QTEtACYnChi8aXSoxGQ6H2xHfUdHVCwUA1OR0UeNcRrMmOzZBOEUBtdoGWSYPg9AMjvxlxq9JOzYAH+G6lT0UbCWgMSGGrqdcIfmshyEak3aUmsZK6l+uIAFCCZZ/HbbCRkkHC5rWKstMI_0_____Retail_CloudEditionN
92fb8726-92a8-4ffc-94ce-f82e07444653_KY%f%7PN-VR6%f%RX-83%f%W6Y-6DD%f%YQ-T6%f%R4W_203_X22-53847_gD6HnT4jP4rcNu9u83gvDiQq1xs7QSujcDbo60Di5iSVa9/ihZ7nlhnA0eDEZfnoDXriRiPPqc09T6AhSnFxLYitAkOuPJqL5UMobIrab9dwTKlowqFolxoHhLOO4V92Hsvn/9JLy7rEzoiAWHhX/0cpMr3FCzVYPeUW1OyLT1A_0_____Retail_CloudEdition
5a85300a-bfce-474f-ac07-a30983e3fb90_N9%f%79K-XWD%f%77-YW%f%3GB-HBG%f%H6-D3%f%2MH_205_X23-15042_blZopkUuayCTgZKH4bOFiisH9GTAHG5/js6UX/qcMWWc3sWNxKSX1OLp1k3h8Xx1cFuvfG/fNAw/I83ssEtPY+A0Gx1JF4QpRqsGOqJ5ruQ2tGW56CJcCVHkB+i46nJAD759gYmy3pEYMQbmpWbhLx3MJ6kvwxKfU+0VCio8k50_0_____OEM:DM_IoTEnterpriseSK
80083eae-7031-4394-9e88-4901973d56fe_P8%f%Q7T-WNK%f%7X-PM%f%FXY-VXH%f%BG-RR%f%K69_206_X23-62084_habUJ0hhAG0P8iIKaRQ74/wZQHyAdFlwHmrejNjOSRG08JeqilJlTM6V8G9UERLJ92/uMDVHIVOPXfN8Zdh8JuYO8oflPnqymIRmff/pU+Gpb871jV2JDA4Cft5gmn+ictKoN4VoSfEZRR+R5hzF2FsoCExDNNw6gLdjtiX94uA_0_____OEM:DM_IoTEnterpriseK
) do (
for /f "tokens=1-9 delims=_" %%A in ("%%#") do (

REM 检测密钥

if %1==key if %osSKU%==%%C if not defined key (
set skufound=1
echo "!applist! !altapplist!" | find /i "%%A" %nul1% && (
if %%F==1 set notworking=1
set key=%%B
)
)

REM 生成票证

if %1==ticket if "%key%"=="%%B" (
set "string=OSMajorVersion=5;OSMinorVersion=1;OSPlatformId=2;PP=0;Pfn=Microsoft.Windows.%%C.%%D_8wekyb3d8bbwe;PKeyIID=465145217131314304264339481117862266242033457260311819664735280;$([char]0)"
for /f "tokens=* delims=" %%i in ('%psc% [conv%f%ert]::ToBas%f%e64String([Text.En%f%coding]::Uni%f%code.GetBytes("""!string!"""^)^)') do set "encoded=%%i"
echo "!encoded!" | find "AAAA" %nul1% || exit /b

<nul set /p "=<?xml version="1.0" encoding="utf-8"?><genuineAuthorization xmlns="http://www.microsoft.com/DRM/SL/GenuineAuthorization/1.0"><version>1.0</version><genuineProperties origin="sppclient"><properties>OA3xOriginalProductId=;OA3xOriginalProductKey=;SessionId=!encoded!;TimeStampClient=2022-10-11T12:00:00Z</properties><signatures><signature name="clientLockboxKey" method="rsa-sha256">%%E=</signature></signatures></genuineProperties></genuineAuthorization>" >"%tdir%\GenuineTicket"
)

)
)
exit /b

::========================================================================================================================================

::  如果当前版本不支持 HWID 激活，以下代码用于获取备用版本名称和密钥

::  第 1 列 = 当前 SKU ID
::  第 2 列 = 当前版本名称
::  第 3 列 = 当前版本激活 ID
::  第 4 列 = 备用版本激活 ID
::  第 5 列 = 备用版本 HWID 密钥
::  第 6 列 = 备用版本名称
::  分隔符  = _


:hwidfallback

set notfoundaltactID=
if %_NoEditionChange%==1 exit /b

for %%# in (
125_EnterpriseS-2021_______________cce9d2de-98ee-4ce2-8113-222620c64a27_ed655016-a9e8-4434-95d9-4345352c2552_QPM%f%6N-7J2%f%WJ-P8%f%8HH-P3Y%f%RH-YY%f%74H_IoTEnterpriseS-2021
125_EnterpriseS-2024_______________f6e29426-a256-4316-88bf-cc5b0f95ec0c_6c4de1b8-24bb-4c17-9a77-7b939414c298_CGK%f%42-GYN%f%6Y-VD%f%22B-BX9%f%8W-J8%f%JXD_IoTEnterpriseS-2024
138_ProfessionalSingleLanguage_____a48938aa-62fa-4966-9d44-9f04da3f72f2_4de7cb65-cdf1-4de9-8ae8-e3cce27b9f2c_VK7%f%JG-NPH%f%TM-C9%f%7JM-9MP%f%GT-3V%f%66T_Professional
139_ProfessionalCountrySpecific____f7af7d09-40e4-419c-a49b-eae366689ebd_4de7cb65-cdf1-4de9-8ae8-e3cce27b9f2c_VK7%f%JG-NPH%f%TM-C9%f%7JM-9MP%f%GT-3V%f%66T_Professional
139_ProfessionalCountrySpecific-Zn_01eb852c-424d-4060-94b8-c10d799d7364_4de7cb65-cdf1-4de9-8ae8-e3cce27b9f2c_VK7%f%JG-NPH%f%TM-C9%f%7JM-9MP%f%GT-3V%f%66T_Professional
) do (
for /f "tokens=1-6 delims=_" %%A in ("%%#") do if %osSKU%==%%A (
echo "!applist! !altapplist!" | find /i "%%C" %nul1% && (
echo "!applist!" | find /i "%%D" %nul1% && (
set altkey=%%E
set altedition=%%F
) || (
set altedition=%%F
set notfoundaltactID=1
)
)
)
)
exit /b

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:OhookActivation
@setlocal DisableDelayedExpansion
@echo off

::  若要使用 Ohook 激活激活 Office，请使用“/Ohook”参数运行脚本，或在以下行中将 0 更改为 1
set _act=0

::  若要移除 Ohook 激活，请使用 /Ohook-Uninstall 参数运行脚本，或在以下行中将 0 更改为 1
set _rem=0

::  如果在上面几行中更改了值或使用参数，脚本将会在无人值守模式下运行

::========================================================================================================================================

cls
color 07
title Ohook 激活 %masver%

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="/Ohook"                  set _act=1
if /i "%%A"=="/Ohook-Uninstall"        set _rem=1
if /i "%%A"=="-el"                     set _elev=1
)
)

for %%A in (%_act% %_rem%) do (if "%%A"=="1" set _unattended=1)

::========================================================================================================================================

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

if %_NCS% EQU 1 (
for /F %%a in ('echo prompt $E ^| cmd') do set "esc=%%a"
set     "Red="41;97m""
set    "Gray="100;97m""
set   "Green="42;97m""
set    "Blue="44;97m""
set  "_White="40;37m""
set  "_Green="40;92m""
set "_Yellow="40;93m""
) else (
set     "Red="Red" "white""
set    "Gray="Darkgray" "white""
set   "Green="DarkGreen" "white""
set    "Blue="Blue" "white""
set  "_White="Black" "Gray""
set  "_Green="Black" "Green""
set "_Yellow="Black" "Yellow""
)

set "nceline=echo: &echo ==== 错误 ==== &echo:"
set "eline=echo: &call :dk_color %Red% "==== 错误 ====" &echo:"
if %~z0 GEQ 200000 (
set "_exitmsg=返回"
set "_fixmsg=请返回主菜单，选择疑难解答并运行修复许可选项。"
) else (
set "_exitmsg=退出"
set "_fixmsg=在 MAS 文件夹中，请运行疑难解答脚本并选择修复许可选项。"
)

::========================================================================================================================================

if %winbuild% LSS 9200 (
%eline%
echo 检测到不受支持的操作系统版本 [%winbuild%]。
echo Ohook 激活在 Windows 8 及更高版本及其服务器对应方上受支持。
goto dk_done
)

::========================================================================================================================================

::  修复路径名称中的特殊字符限制

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

set _PSarg="""%~f0""" -el %_args%

set "_ttemp=%userprofile%\AppData\Local\Temp"
set "_Local=%LocalAppData%"
setlocal EnableDelayedExpansion

::========================================================================================================================================

if %_rem%==1 goto :oh_uninstall

:oh_menu

if %_unattended%==0 (
cls
mode 76, 25
title Ohook 激活 %masver%

echo:
echo:
echo:
echo:
echo         ____________________________________________________________
echo:
echo                 [1] 安装 Ohook Office 激活
echo:
echo                 [2] 卸载
echo                 ____________________________________________
echo:
echo                 [3] 下载 Office
echo:
echo                 [0] %_exitmsg%
echo         ____________________________________________________________
echo: 
call :dk_color2 %_White% "              " %_Green% "请输入一个菜单选项 [1,2,3,0]"
choice /C:1230 /N
set _el=!errorlevel!
if !_el!==4  exit /b
if !_el!==3  start %mas%genuine-installation-media.html &goto :oh_menu
if !_el!==2  goto :oh_uninstall
if !_el!==1  goto :oh_menu2
goto :oh_menu
)

::========================================================================================================================================

:oh_menu2

cls
mode 130, 32
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=32;$B.Height=300;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}"

title Ohook 激活 %masver%

echo:
echo 正在初始化……

::  检查 PowerShell

%psc% $ExecutionContext.SessionState.LanguageMode %nul2% | find /i "Full" %nul1% || (
%eline%
%psc% $ExecutionContext.SessionState.LanguageMode
echo:
echo PowerShell 不可用，正在中止……
echo 如果你对Powershell施加了限制，请撤销这些更改。
echo:
echo 请查看此页面以获得帮助。 %mas%troubleshoot
goto dk_done
)

::========================================================================================================================================

call :dk_product
call :dk_ckeckwmic

::  显示潜在的脚本卡住情况的信息

sc start sppsvc %nul%
if %errorlevel% NEQ 1056 if %errorlevel% NEQ 0 (
echo:
echo 错误代码：%errorlevel%
call :dk_color %Red% "启动 [sppsvc] 服务失败，其余的进程可能需要很长时间……"
echo:
)

::========================================================================================================================================

set error=

cls
echo:
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PROCESSOR_ARCHITECTURE') do set osarch=%%b
for /f "tokens=6-7 delims=[]. " %%i in ('ver') do if "%%j"=="" (set fullbuild=%%i) else (set fullbuild=%%i.%%j)
echo 正在检查操作系统信息                    [%winos% ^| %fullbuild% ^| %osarch%]

::========================================================================================================================================

::  检查 Windows Script Host

set _WSH=1
reg query "HKCU\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled %nul2% | find /i "0x0" %nul1% && (set _WSH=0)
reg query "HKLM\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled %nul2% | find /i "0x0" %nul1% && (set _WSH=0)

if %_WSH% EQU 0 (
reg add "HKLM\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f %nul%
reg add "HKCU\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f %nul%
if not "%arch%"=="x86" reg add "HKLM\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f /reg:32 %nul%
echo 正在启用 Windows Script Host            [成功]
)

::========================================================================================================================================

echo 正在初始化诊断测试……

set "_serv=sppsvc Winmgmt"
set officeact=1
call :dk_errorcheck

::  检查不支持的 office 版本

set o14msi=
set o14c2r=
set o16uwp=

set _68=HKLM\SOFTWARE\Microsoft\Office
set _86=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o14msi=Office 2010 MSI )
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\14.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o14msi=Office 2010 MSI )
%nul% reg query %_68%\14.0\CVH /f Click2run /k         && set o14c2r=Office 2010 C2R 
%nul% reg query %_86%\14.0\CVH /f Click2run /k         && set o14c2r=Office 2010 C2R 

if %winbuild% GEQ 10240 %psc% "Get-AppxPackage -name "Microsoft.Office.Desktop"" | find /i "Office" %nul1% && set o16uwp=Office UWP 

if not "%o14msi%%o14c2r%%o16uwp%"=="" (
echo:
call :dk_color %Red% "正在检查不受支持的 Office 安装          [ %o14msi%%o14c2r%%o16uwp%]"
)

::========================================================================================================================================

::  检查受支持的 office 版本

call :oh_getpath

sc query ClickToRunSvc %nul%
set error1=%errorlevel%

if defined o16c2r if %error1% EQU 1060 (
call :dk_color %Red% "正在检查 ClickToRun 服务                [未找到，已找到 Office 16.0 文件]"
set o16c2r=
set error=1
)

sc query OfficeSvc %nul%
set error2=%errorlevel%

if defined o15c2r if %error1% EQU 1060 if %error2% EQU 1060 (
call :dk_color %Red% "正在检查 ClickToRun 服务                [未找到，已找到 Office 15.0 文件]"
set o15c2r=
set error=1
)

if "%o16c2r%%o15c2r%%o16msi%%o15msi%"=="" (
set error=1
echo:
if not "%o14msi%%o14c2r%%o16uwp%"=="" (
call :dk_color %Red% "正在检查受支持的 Office 安装            [未找到]"
) else (
call :dk_color %Red% "正在检查已安装的 Office                 [未找到]"
)

if %winbuild% GEQ 10240 %psc% "Get-AppxPackage -name "Microsoft.MicrosoftOfficeHub"" | find /i "Office" %nul1% && (
echo:
echo 你只安装了 Office 仪表板应用，你需要安装完整的 Office 版本。
)
echo:
call :dk_color %Blue% "请从以下 URL 下载并安装 Office，然后重试。"
echo:
echo %mas%genuine-installation-media.html
goto dk_done
)

set multioffice=
if not "%o16c2r%%o15c2r%%o16msi%%o15msi%"=="1" set multioffice=1
if not "%o14msi%%o14c2r%%o16uwp%"=="" set multioffice=1

if defined multioffice (
call :dk_color %Gray% "正在检查是否有多个 Office 安装          [已找到。最好只安装一个版本]"
)

::========================================================================================================================================

::  处理 Office 15.0 C2R

if not defined o15c2r goto :starto16c2r

call :oh_reset
call :oh_actids

set oVer=15
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg% /v InstallPath" %nul6%') do (set "_oRoot=%%b\root")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\Configuration /v Platform" %nul6%') do (set "_oArch=%%b")
if not defined _oArch for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\propertyBag /v Platform" %nul6%') do (set "_oArch=%%b")

echo "%o15c2r_reg%" | find /i "Wow6432Node" %nul1% && (set _tok=10) || (set _tok=9)
for /f "tokens=%_tok% delims=\" %%a in ('reg query %o15c2r_reg%\ProductReleaseIDs\Active %nul6% ^| findstr /i "Retail Volume"') do (
echo "!_oIds!" | find /i " %%a " %nul1% || (set "_oIds= !_oIds! %%a ")
)

set "_oLPath=%_oRoot%\Licenses"
set "_oIntegrator=%_oRoot%\integration\integrator.exe"

if [%_oArch%]==[x64] (set "_hookPath=%_oRoot%\vfs\System"    & set "_hook=sppc64.dll")
if [%_oArch%]==[x86] (set "_hookPath=%_oRoot%\vfs\SystemX86" & set "_hook=sppc32.dll")
if not [%osarch%]==[x86] (
if [%_oArch%]==[x64] set "_sppcPath=%SystemRoot%\System32\sppc.dll"
if [%_oArch%]==[x86] set "_sppcPath=%SystemRoot%\SysWOW64\sppc.dll"
) else (
set "_sppcPath=%SystemRoot%\System32\sppc.dll"
)

echo:
echo 正在激活 Office 15.0 %_oArch% C2R……

if not defined _oIds (
call :dk_color %Red% "正在检查已安装产品                      [产品 ID 未找到。正在中止激活……]"
set error=1
goto :starto16c2r
)

call :oh_process
call :oh_hookinstall

::========================================================================================================================================

:starto16c2r

::  处理 Office 16.0 C2R

if not defined o16c2r goto :startmsi

call :oh_reset
call :oh_actids

set oVer=16
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg% /v InstallPath" %nul6%') do (set "_oRoot=%%b\root")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg%\Configuration /v Platform" %nul6%') do (set "_oArch=%%b")

echo "%o16c2r_reg%" | find /i "Wow6432Node" %nul1% && (set _tok=9) || (set _tok=8)
for /f "tokens=%_tok% delims=\" %%a in ('reg query "%o16c2r_reg%\ProductReleaseIDs" /s /f ".16" /k %nul6% ^| findstr /i "Retail Volume"') do (
echo "!_oIds!" | find /i " %%a " %nul1% || (set "_oIds= !_oIds! %%a ")
)
set _oIds=%_oIds:.16=%

set "_oLPath=%_oRoot%\Licenses16"
set "_oIntegrator=%_oRoot%\integration\integrator.exe"

if [%_oArch%]==[x64] (set "_hookPath=%_oRoot%\vfs\System"    & set "_hook=sppc64.dll")
if [%_oArch%]==[x86] (set "_hookPath=%_oRoot%\vfs\SystemX86" & set "_hook=sppc32.dll")
if not [%osarch%]==[x86] (
if [%_oArch%]==[x64] set "_sppcPath=%SystemRoot%\System32\sppc.dll"
if [%_oArch%]==[x86] set "_sppcPath=%SystemRoot%\SysWOW64\sppc.dll"
) else (
set "_sppcPath=%SystemRoot%\System32\sppc.dll"
)

echo:
echo 正在激活 Office 16.0 %_oArch% C2R……

if not defined _oIds (
call :dk_color %Red% "正在检查已安装产品                      [产品 ID 未找到。正在中止激活……]"
set error=1
goto :startmsi
)

call :oh_process
call :oh_hookinstall

::========================================================================================================================================

::  查找 Office vNext 许可证块的残余并将其删除，因为它会阻止非 vNext 许可证的显示
::  https://learn.microsoft.com/en-us/office/troubleshoot/activation/reset-office-365-proplus-activation-state

set _sid=
set sub_next=

for /f "tokens=* delims=" %%a in ('%psc% "Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' | ForEach-Object { Split-Path -Path $_.PSPath -Leaf }" %nul6%') do (if defined _sid (set "_sid=!_sid! HKU\%%a") else (set "_sid=HKU\%%a"))

if not defined _sid (
call :dk_color %Red% "检查用户账号 SID                        [未找到]"
)

dir /b /s /a:-d "!_Local!\Microsoft\Office\Licenses\*" %nul% && set sub_next=1
dir /b /s /a:-d "!ProgramData!\Microsoft\Office\Licenses\*" %nul% && set sub_next=1

for %%# in (!_sid! HKCU) do if not defined sub_next (
reg query %%#\Software\Microsoft\Office\16.0\Common\Licensing\LicensingNext /v MigrationToV5Done %nul2% | find /i "0x1" %nul% && (
reg query %%#\Software\Microsoft\Office\16.0\Common\Licensing\LicensingNext %nul2% | findstr /i "volume retail" %nul2% | findstr /i "0x2 0x3" %nul% && (
set sub_next=1
)
)
)

if defined sub_next (
rmdir /s /q "!_Local!\Microsoft\Office\Licenses\" %nul%
rmdir /s /q "!ProgramData!\Microsoft\Office\Licenses\" %nul%
for %%# in (!_sid! HKCU) do (
reg delete %%#\Software\Microsoft\Office\16.0\Common\Licensing /f %nul%
reg delete %%#\Software\Microsoft\Office\16.0\Common\Identity /f %nul%
reg delete %%#\Software\Microsoft\Office\16.0\Registration /f %nul%
)
)

if defined sub_next echo 正在移除 Office vNext 许可              [成功]

::========================================================================================================================================

::  订阅产品会尝试验证许可证，并可能显示横幅“检查此设备的许可证状态时出现问题”。
::  复原注册表项可以跳过此检查

if defined o16c2r (
for %%# in (!_sid! HKCU) do (reg delete %%#\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency /f %nul%)
for %%# in (!_sid! HKCU) do (
reg query "%%#\Volatile Environment" %nul% && (
reg add %%#\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency /v "TimeOfLastHeartbeatFailure" /t REG_SZ /d "2040-01-01T00:00:00Z" /f %nul%
)
)
echo 添加注册表以跳过许可证检查              [成功]
)

::========================================================================================================================================

::  mass grave[.]dev/office-license-is-not-genuine.html
::  为批量产品添加注册表项，以便不会显示“非正版”横幅
::  脚本已使用 MAK 而不是 GVLK，因此无论如何都不会显示，但如果 Office 为批量产品安装默认 GVLK 宽限密钥，则会添加注册表项

echo "%_oIds%" | find /i "Volume" %nul1% && (
if %winbuild% GEQ 9200 (
if not [%osarch%]==[x86] (
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f /reg:32 %nul%
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f /v KeyManagementServiceName /t REG_SZ /d "10.0.0.10" /reg:32 %nul%
)
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f %nul%
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f /v KeyManagementServiceName /t REG_SZ /d "10.0.0.10" %nul%
echo 添加注册表以防止显示“非正版”横幅        [成功]
)
)

::========================================================================================================================================

:startmsi

if defined o15msi call :oh_processmsi 15 %o15msi_reg%
if defined o16msi call :oh_processmsi 16 %o16msi_reg%

::========================================================================================================================================

::  卸载其他 / 宽限期密钥

set upk_result=0
set allapplist=

if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (ApplicationID='0ff1ce15-a989-479d-af46-f275c6370663' and PartialProductKey is not null) get ID /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM SoftwareLicensingProduct WHERE ApplicationID=''0ff1ce15-a989-479d-af46-f275c6370663'' AND PartialProductKey IS NOT NULL').Get()).ID ^| %% {echo ('ID='+$_)}" %nul6%')"
%chkapp% do (if defined allapplist (call set "allapplist=!allapplist! %%a") else (call set "allapplist=%%a"))

for %%# in (%allapplist%) do (
echo "%_allactid%" | find /i "%%#" %nul1% || (
cscript //nologo %windir%\system32\slmgr.vbs /upk %%# %nul% && (
set upk_result=1
) || (
set error=1
set upk_result=2
)
)
)

if not %upk_result%==0 echo:
if %upk_result%==1 echo 正在卸载其他/宽限期密钥                 [成功]
if %upk_result%==2 call :dk_color %Red% "正在卸载其他/宽限期密钥                 [失败]"

::========================================================================================================================================

::  刷新 Windows Insider Preview 许可证
::  在 Insider 版本中需要它，否则 office 可能无法激活

if exist "%windir%\system32\spp\store_test\2.0\tokens.dat" (
cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
if !errorlevel! NEQ 0 cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
)

::========================================================================================================================================

echo:
if not defined error (
call :dk_color %Green% "Office 已永久激活。"
echo 帮助：%mas%troubleshoot
) else (
call :dk_color %Red% "检测到一些错误。"
if not defined ierror if not defined showfix if not defined serv_cor if not defined serv_cste call :dk_color %Blue% "%_fixmsg%"
echo:
call :dk_color2 %Blue% "查看此页面以获取帮助" %_Yellow% " %mas%troubleshoot"
)

goto :dk_done

::========================================================================================================================================

:oh_uninstall

cls
mode 99, 28
title 卸载 Ohook 激活 %masver%

set _present=
set _unerror=
call :oh_reset
call :oh_getpath

echo:
echo 正在卸载 Ohook 激活……
echo:

if defined o16c2r_reg (for /f "skip=2 tokens=2*" %%a in ('"reg query %o16c2r_reg% /v InstallPath" %nul6%') do (set "_16CHook=%%b\root\vfs"))
if defined o15c2r_reg (for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg% /v InstallPath" %nul6%') do (set "_15CHook=%%b\root\vfs"))
if defined o16msi_reg (for /f "skip=2 tokens=2*" %%a in ('"reg query %o16msi_reg%\Common\InstallRoot /v Path" %nul6%') do (set "_16MHook=%%b"))
if defined o15msi_reg (for /f "skip=2 tokens=2*" %%a in ('"reg query %o15msi_reg%\Common\InstallRoot /v Path" %nul6%') do (set "_15MHook=%%b"))

if defined _16CHook (if exist "%_16CHook%\System\sppc*dll"    (set _present=1& del /s /f /q "%_16CHook%\System\sppc*dll"    & if exist "%_16CHook%\System\sppc*dll"    set _unerror=1))
if defined _16CHook (if exist "%_16CHook%\SystemX86\sppc*dll" (set _present=1& del /s /f /q "%_16CHook%\SystemX86\sppc*dll" & if exist "%_16CHook%\SystemX86\sppc*dll" set _unerror=1))
if defined _15CHook (if exist "%_15CHook%\System\sppc*dll"    (set _present=1& del /s /f /q "%_15CHook%\System\sppc*dll"    & if exist "%_15CHook%\System\sppc*dll"    set _unerror=1))
if defined _15CHook (if exist "%_15CHook%\SystemX86\sppc*dll" (set _present=1& del /s /f /q "%_15CHook%\SystemX86\sppc*dll" & if exist "%_15CHook%\SystemX86\sppc*dll" set _unerror=1))
if defined _16MHook (if exist "%_16MHook%sppc*dll"            (set _present=1& del /s /f /q "%_16MHook%sppc*dll"            & if exist "%_16MHook%sppc*dll"            set _unerror=1))
if defined _15MHook (if exist "%_15MHook%sppc*dll"            (set _present=1& del /s /f /q "%_15MHook%sppc*dll"            & if exist "%_15MHook%sppc*dll"            set _unerror=1))

for %%# in (15 16) do (
for %%A in ("%ProgramFiles%" "%ProgramW6432%" "%ProgramFiles(x86)%") do (
if exist "%%~A\Microsoft Office\Office%%#\sppc*dll" (set _present=1& del /s /f /q "%%~A\Microsoft Office\Office%%#\sppc*dll" & if exist "%%~A\Microsoft Office\Office%%#\sppc*dll" set _unerror=1)
)
)

for %%# in (System SystemX86) do (
for %%G in ("Office 15" "Office") do (
for %%A in ("%ProgramFiles%" "%ProgramW6432%" "%ProgramFiles(x86)%") do (
if exist "%%~A\Microsoft %%~G\root\vfs\%%#\sppc*dll" (set _present=1& del /s /f /q "%%~A\Microsoft %%~G\root\vfs\%%#\sppc*dll" & if exist "%%~A\Microsoft %%~G\root\vfs\%%#\sppc*dll" set _unerror=1)
)
)
)

reg query HKCU\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency %nul% && (
echo:
echo 正在删除 - 用于跳过许可证检查的注册表项
reg delete HKCU\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency /f

for /f "tokens=* delims=" %%a in ('%psc% "Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' | ForEach-Object { Split-Path -Path $_.PSPath -Leaf }" %nul6%') do (if defined _sid (set "_sid=!_sid! %%a") else (set "_sid=%%a"))
for %%# in (!_sid!) do (reg query HKU\%%#\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency %nul% && (
reg delete HKU\%%#\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency /f
)
)
)

reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" %nul% && (
echo:
echo 正在删除 - 防止“非正版”横幅的注册表项
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f
)

reg query "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" %nul% && (
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f
)

echo __________________________________________________________________________________________
echo:

if not defined _present (
echo Ohook 激活未安装。
) else (
if defined _unerror (
call :dk_color %Red% "卸载 Ohook 激活失败。"
call :dk_color %Blue% "如果 Office 应用正在运行，请关闭它们，然后重试。"
) else (
call :dk_color %Green% "卸载 Ohook 激活成功。"
)
)
echo __________________________________________________________________________________________

goto :dk_done

::========================================================================================================================================

:oh_reset

set _oRoot=
set _oArch=
set _oIds=
set _oLPath=
set _hookPath=
set _hook=
set _sppcPath=
set _key=
set _actid=
set _prod=
set _lic=
set _License=
exit /b

::========================================================================================================================================

:oh_getpath

set o16c2r=
set o15c2r=
set o16msi=
set o15msi=

set _68=HKLM\SOFTWARE\Microsoft\Office
set _86=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms"    (set o16c2r=1&set o16c2r_reg=%_86%\ClickToRun)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms"    (set o16c2r=1&set o16c2r_reg=%_68%\ClickToRun)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\15.0\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (set o15c2r=1&set o15c2r_reg=%_86%\15.0\ClickToRun)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\15.0\ClickToRun /v InstallPath" %nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (set o15c2r=1&set o15c2r_reg=%_68%\15.0\ClickToRun)

for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o16msi=1&set o16msi_reg=%_86%\16.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\16.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o16msi=1&set o16msi_reg=%_68%\16.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_86%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o15msi=1&set o15msi_reg=%_86%\15.0)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_68%\15.0\Common\InstallRoot /v Path" %nul6%') do if exist "%%b\EntityPicker.dll" (set o15msi=1&set o15msi_reg=%_68%\15.0)

exit /b

::========================================================================================================================================

:oh_installkey

if %_wmic% EQU 1 wmic path SoftwareLicensingService where __CLASS='SoftwareLicensingService' call InstallProductKey ProductKey="%_key%" %nul%
if %_wmic% EQU 0 %psc% "(([WMISEARCHER]'SELECT Version FROM SoftwareLicensingService').Get()).InstallProductKey('%_key%')" %nul%
if not %errorlevel%==0 cscript //nologo %windir%\system32\slmgr.vbs /ipk %_key% %nul%
set errorcode=%errorlevel%
cmd /c exit /b %errorcode%
if %errorcode% NEQ 0 set "errorcode=[0x%=ExitCode%]"

if %errorcode% EQU 0 (
call :dk_refresh
echo 正在安装通用产品密钥                    [%_key%] [%_prod%] [%_lic%] [成功]
) else (
call :dk_color %Red% "正在安装通用产品密钥                    [%_key%] [%_prod%] [失败] %errorcode%"
if not defined error (
call :dk_color %Blue% "%_fixmsg%"
set showfix=1
)
set error=1
)

exit /b

::========================================================================================================================================

:oh_installlic

if not defined _oLPath exit /b

if %oVer%==16 (
"!_oIntegrator!" /I /License PRIDName=%_License%.16 PidKey=%_key% %nul%
) else (
"!_oIntegrator!" /I /License PRIDName=%_License% PidKey=%_key% %nul%
)

call :oh_actids
echo "!oapplist!" | find /i "!_actid!" %nul1% && (
call :dk_color %Gray% "正在安装缺失的许可证文件                [Office %oVer%.0 %_prod%] [成功]"
exit /b
)

::  回退到 /ilc 方法以安装许可证，以防 integrator.exe 不可用

set _License=%_License:XVolume=XC2RVL_%

set _License=%_License:O365EduCloudRetail=O365EduCloudEDUR_%

set _License=%_License:ProjectProRetail=ProjectProO365R_%
set _License=%_License:ProjectStdRetail=ProjectStdO365R_%
set _License=%_License:VisioProRetail=VisioProO365R_%
set _License=%_License:VisioStdRetail=VisioStdO365R_%

if defined _preview set _License=%_License:Volume=PreviewVL_%

set _License=%_License:Retail=R_%
set _License=%_License:Volume=VL_%

for %%# in ("!_oLPath!\client-issuance-*.xrm-ms") do (
cscript //nologo %windir%\system32\slmgr.vbs /ilc "!_oLPath!\%%~nx#" %nul%
)
cscript //nologo %windir%\system32\slmgr.vbs /ilc "!_oLPath!\pkeyconfig-office.xrm-ms" %nul%

for %%# in ("!_oLPath!\%_License%*.xrm-ms") do (
cscript //nologo %windir%\system32\slmgr.vbs /ilc "!_oLPath!\%%~nx#" %nul%
)

call :oh_actids
echo "!oapplist!" | find /i "!_actid!" %nul1% && (
call :dk_color %Gray% "正在安装缺失的许可证文件                [Office %oVer%.0 %_prod%] [使用 /ilc 方法成功]"
) || (
set error=1
call :dk_color %Red% "正在安装缺失的许可证文件                [Office %oVer%.0 %_prod%] [失败]"
)

exit /b

::========================================================================================================================================

:oh_hookinstall

set ierror=
set hasherror=

if %_hook%==sppc32.dll set offset=2564
if %_hook%==sppc64.dll set offset=3076

del /s /q "%_hookPath%\sppcs.dll" %nul%
del /s /q "%_hookPath%\sppc.dll" %nul%

if exist "%_hookPath%\sppcs.dll" set ierror=1
if exist "%_hookPath%\sppc.dll" set ierror=1

mklink "%_hookPath%\sppcs.dll" "%_sppcPath%" %nul%
if not %errorlevel%==0 set ierror=1

if not exist "%_hookPath%\sppc.dll" call :oh_extractdll "%_hookPath%\sppc.dll" "%offset%"
if not exist "%_hookPath%\sppc.dll" set ierror=1

echo:
if not defined ierror (
echo 正在用符号链接系统 sppc.dll 至          ["%_hookPath%\sppcs.dll"] [成功]
echo 正在解压自定义 %_hook% 到            ["%_hookPath%\sppc.dll"] [成功]
) else (
set error=1
call :dk_color %Red% "正在用符号链接系统 sppc.dll             [失败]"
call :dk_color %Red% "正在解压自定义 %_hook%               [失败]"
echo ["%_hookPath%\sppc.dll"]
echo:
call :dk_color %Blue% "请关闭所有 Office 应用（包括 Outlook），然后重试。"
call :dk_color %Blue% "如果仍未解决，请重新启动系统并再试一次。"
)

if not defined ierror (
if defined hasherror (
set error=1
set ierror=1
call :dk_color %Red% "正在修改自定义 %_hook% 的哈希值      [失败]"
) else (
echo 正在修改自定义 %_hook% 的哈希值      [成功]
)
)

exit /b

::========================================================================================================================================

:oh_process

for %%# in (%_oIds%) do (
set _key=
set _actid=
set _lic=
set _preview=
set _License=%%#

echo %%# | find /i "2024" %nul% && (
if exist "!_oLPath!\ProPlus2024PreviewVL_*.xrm-ms" if not exist "!_oLPath!\ProPlus2024VL_*.xrm-ms" set _preview=-Preview
)
set _prod=%%#!_preview!

call :ohookdata getinfo !_prod!

if not [!_key!]==[] (
echo "!oapplist!" | find /i "!_actid!" %nul1% || call :oh_installlic
call :oh_installkey
) else (
set error=1
call :dk_color %Red% "检查脚本中产品列表              [Office %oVer%.0 %%# 未在脚本中找到]"
call :dk_color %Blue% 请确保你使用的是 MAS 脚本的最新版本。
)
)

exit /b

::========================================================================================================================================

:oh_msiproducts

set msitemp=%SystemRoot%\Temp\_msitemp.txt

if %oVer%==15 set _psmsikey=%o15msi_reg:HKLM\=HKLM:%
if %oVer%==16 set _psmsikey=%o16msi_reg:HKLM\=HKLM:%

if exist %msitemp% del /f /q %msitemp%
%psc% "$Key = '%_psmsikey%\Registration\{*FF1CE}'; $keydata = Get-ItemProperty -Path $Key -Name "DigitalProductID"; $binaryData = $keydata."DigitalProductID"; $stringData = [System.Text.Encoding]::Unicode.GetString($binaryData);$stringData" >>%msitemp%

if exist %msitemp% call :ohookdata getmsiprod
if exist %msitemp% del /f /q %msitemp%

exit /b

::========================================================================================================================================

:oh_processmsi

::  处理 Office MSI 版本

call :oh_reset
call :oh_actids

set oVer=%1
for /f "skip=2 tokens=2*" %%a in ('"reg query %2\Common\InstallRoot /v Path" %nul6%') do (set "_oRoot=%%b")
if "%_oRoot:~-1%"=="\" set "_oRoot=%_oRoot:~0,-1%"

echo "%2" | find /i "Wow6432Node" %nul1% && set _oArch=x86
if not [%osarch%]==[x86] if not defined _oArch set _oArch=x64
if [%osarch%]==[x86] set _oArch=x86

if [%_oArch%]==[x64] (set "_hookPath=%_oRoot%" & set "_hook=sppc64.dll")
if [%_oArch%]==[x86] (set "_hookPath=%_oRoot%" & set "_hook=sppc32.dll")
if not [%osarch%]==[x86] (
if [%_oArch%]==[x64] set "_sppcPath=%SystemRoot%\System32\sppc.dll"
if [%_oArch%]==[x86] set "_sppcPath=%SystemRoot%\SysWOW64\sppc.dll"
) else (
set "_sppcPath=%SystemRoot%\System32\sppc.dll"
)

call :oh_msiproducts

echo:
echo 正在激活 Office %1.0 %_oArch% MSI 版本……

if not defined _oIds (
set error=1
call :dk_color %Red% "正在检查已安装产品                      [产品 ID 未找到。正在中止激活……]"
exit /b
)

call :oh_process
call :oh_hookinstall

exit /b

::========================================================================================================================================

::  获取 Office 激活 ID

:oh_actids

set oapplist=
if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (ApplicationID='0ff1ce15-a989-479d-af46-f275c6370663') get ID /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM SoftwareLicensingProduct WHERE ApplicationID=''0ff1ce15-a989-479d-af46-f275c6370663''').Get()).ID ^| %% {echo ('ID='+$_)}" %nul6%')"
%chkapp% do (if defined oapplist (call set "oapplist=!oapplist! %%a") else (call set "oapplist=%%a"))
exit /b

::========================================================================================================================================

::  第 1 列 = Office 版本号
::  第 2 列 = 激活 ID
::  第 3 列 = 通用密钥。给出优先顺序，Retail:TB:Sub > Retail > OEM:NONSLP > Volume:MAK > Volume:GVLK
::  第 4 列 = 许可证描述的最后一部分
::  第 5 列 = 版本
::  分隔符  = "_"

:ohookdata

set f=
for %%# in (
15_ab4d047b-97cf-4126-a69f-34df08e2f254_B7%f%RFY-7N%f%XPK-Q43%f%42-Y9%f%X2H-3JX%f%4X_Retail________AccessRetail
15_4374022d-56b8-48c1-9bb7-d8f2fc726343_9M%f%F9G-CN%f%32B-HV7%f%XT-9X%f%J8T-9KV%f%F4_MAK___________AccessVolume
15_1b1d9bd5-12ea-4063-964c-16e7e87d6e08_NT%f%889-MB%f%H4X-8MD%f%4H-X8%f%R2D-WQH%f%F8_Retail________ExcelRetail
15_ac1ae7fd-b949-4e04-a330-849bc40638cf_Y3%f%N36-YC%f%HDK-XYW%f%BG-KY%f%QVV-BDT%f%J2_MAK___________ExcelVolume
15_cfaf5356-49e3-48a8-ab3c-e729ab791250_BM%f%K4W-6N%f%88B-BP9%f%QR-PH%f%FCK-MG7%f%GF_Retail________GrooveRetail
15_4825ac28-ce41-45a7-9e6e-1fed74057601_RN%f%84D-7H%f%CWY-FTC%f%BK-JM%f%XWM-HT7%f%GJ_MAK___________GrooveVolume
15_c02fb62e-1cd5-4e18-ba25-e0480467ffaa_2W%f%QNF-GB%f%K4B-XVG%f%6F-BB%f%MX7-M4F%f%2Y_OEM-Perp______HomeBusinessPipcRetail
15_a2b90e7a-a797-4713-af90-f0becf52a1dd_YW%f%D4R-CN%f%KVT-VG8%f%VJ-93%f%33B-RCW%f%9F_Subscription__HomeBusinessRetail
15_f2de350d-3028-410a-bfae-283e00b44d0e_6W%f%W3N-BD%f%GM9-PCC%f%HD-9Q%f%PP9-P34%f%QG_Subscription__HomeStudentRetail
15_44984381-406e-4a35-b1c3-e54f499556e2_RV%f%7NQ-HY%f%3WW-7CK%f%WH-QT%f%VMW-29V%f%HC_Retail________InfoPathRetail
15_9e016989-4007-42a6-8051-64eb97110cf2_C4%f%TGN-QQ%f%W6Y-FYK%f%XC-6W%f%JW7-X73%f%VG_MAK___________InfoPathVolume
15_9103f3ce-1084-447a-827e-d6097f68c895_6M%f%DN4-WF%f%3FV-4WH%f%3Q-W6%f%99V-RGC%f%MY_PrepidBypass__LyncAcademicRetail
15_ff693bf4-0276-4ddb-bb42-74ef1a0c9f4d_N4%f%2BF-CB%f%Y9F-W2C%f%7R-X3%f%97X-DYF%f%QW_PrepidBypass__LyncEntryRetail
15_fada6658-bfc6-4c4e-825a-59a89822cda8_89%f%P23-2N%f%K2R-JXM%f%2M-3Q%f%8R8-BWM%f%3Y_Retail________LyncRetail
15_e1264e10-afaf-4439-a98b-256df8bb156f_3W%f%KCD-RN%f%489-4M7%f%XJ-GJ%f%2GQ-YBF%f%Q6_MAK___________LyncVolume
15_69ec9152-153b-471a-bf35-77ec88683eae_VN%f%WHF-FK%f%FBW-Q2R%f%GD-HY%f%HWF-R3H%f%H2_Subscription__MondoRetail
15_f33485a0-310b-4b72-9a0e-b1d605510dbd_2Y%f%NYQ-FQ%f%MVG-CB8%f%KW-6X%f%KYD-M7R%f%RJ_MAK___________MondoVolume
15_3391e125-f6e4-4b1e-899c-a25e6092d40d_4T%f%GWV-6N%f%9P6-G2H%f%8Y-2H%f%WKB-B4F%f%F4_Bypass________OneNoteFreeRetail
15_8b524bcc-67ea-4876-a509-45e46f6347e8_3K%f%XXQ-PV%f%N2C-8P7%f%YY-HC%f%V88-GVG%f%Q6_Retail________OneNoteRetail
15_b067e965-7521-455b-b9f7-c740204578a2_JD%f%MWF-NJ%f%C7B-HRC%f%HY-WF%f%T8G-BPX%f%D9_MAK___________OneNoteVolume
15_12004b48-e6c8-4ffa-ad5a-ac8d4467765a_9N%f%4RQ-CF%f%8R2-HBV%f%CB-J3%f%C9V-94P%f%4D_Retail________OutlookRetail
15_8d577c50-ae5e-47fd-a240-24986f73d503_HN%f%G29-GG%f%WRG-RFC%f%8C-JT%f%FP4-2J9%f%FH_MAK___________OutlookVolume
15_5aab8561-1686-43f7-9ff5-2c861da58d17_9C%f%YB3-NF%f%MRW-YFD%f%G6-XC%f%7TF-BY3%f%6J_OEM-Perp______PersonalPipcRetail
15_17e9df2d-ed91-4382-904b-4fed6a12caf0_2N%f%CQJ-MF%f%RMH-TXV%f%83-J7%f%V4C-RVR%f%WC_Retail________PersonalRetail
15_31743b82-bfbc-44b6-aa12-85d42e644d5b_HV%f%MN2-KP%f%HQH-DVQ%f%MK-7B%f%3CM-FGB%f%FC_Retail________PowerPointRetail
15_e40dcb44-1d5c-4085-8e8f-943f33c4f004_47%f%DKN-HP%f%JP7-RF9%f%M3-VC%f%YT2-TMQ%f%4G_MAK___________PowerPointVolume
15_064383fa-1538-491c-859b-0ecab169a0ab_N3%f%QMM-GK%f%DT3-JQG%f%X6-7X%f%3MQ-4GB%f%G3_Retail________ProPlusRetail
15_2b88c4f2-ea8f-43cd-805e-4d41346e18a7_QK%f%HNX-M9%f%GGH-T3Q%f%MW-YP%f%K4Q-QRP%f%9V_MAK___________ProPlusVolume
15_4e26cac1-e15a-4467-9069-cb47b67fe191_CF%f%9DD-6C%f%NW2-BJW%f%JQ-CV%f%CFX-Y7T%f%XD_OEM-Perp______ProfessionalPipcRetail
15_44bc70e2-fb83-4b09-9082-e5557e0c2ede_MB%f%QBN-CQ%f%PT6-PXR%f%MC-TY%f%JFR-3C8%f%MY_Retail________ProfessionalRetail
15_2f72340c-b555-418d-8b46-355944fe66b8_WP%f%Y8N-PD%f%PY4-FC7%f%TF-KM%f%P7P-KWY%f%FY_Subscription__ProjectProRetail
15_ed34dc89-1c27-4ecd-8b2f-63d0f4cedc32_WF%f%CT2-NB%f%FQ7-JD7%f%VV-MF%f%JX6-6F2%f%CM_MAK___________ProjectProVolume
15_58d95b09-6af6-453d-a976-8ef0ae0316b1_NT%f%HQT-VK%f%K6W-BRB%f%87-HV%f%346-Y96%f%W8_Subscription__ProjectStdRetail
15_2b9e4a37-6230-4b42-bee2-e25ce86c8c7a_3C%f%NQX-T3%f%4TY-99R%f%H4-C4%f%YD2-KWY%f%GV_MAK___________ProjectStdVolume
15_c3a0814a-70a4-471f-af37-2313a6331111_TW%f%NCJ-YR%f%84W-X7P%f%PF-6D%f%PRP-D67%f%VC_Retail________PublisherRetail
15_38ea49f6-ad1d-43f1-9888-99a35d7c9409_DJ%f%PHV-NC%f%JV6-GWP%f%T6-K2%f%6JX-C7G%f%X6_MAK___________PublisherVolume
15_ba3e3833-6a7e-445a-89d0-7802a9a68588_3N%f%Y6J-WH%f%T3F-47B%f%DV-JH%f%F36-234%f%3W_PrepidBypass__SPDRetail
15_32255c0a-16b4-4ce2-b388-8a4267e219eb_V6%f%VWN-KC%f%2HR-YYD%f%D6-9V%f%7HQ-7T7%f%VP_Retail________StandardRetail
15_a24cca51-3d54-4c41-8a76-4031f5338cb2_9T%f%N6B-PC%f%YH4-MCV%f%DQ-KT%f%83C-TMQ%f%7T_MAK___________StandardVolume
15_a56a3b37-3a35-4bbb-a036-eee5f1898eee_NV%f%K2G-2M%f%Y4G-7JX%f%2P-7D%f%6F2-VFQ%f%BR_Subscription__VisioProRetail
15_3e4294dd-a765-49bc-8dbd-cf8b62a4bd3d_YN%f%7CF-XR%f%H6R-CGK%f%RY-GK%f%PV3-BG7%f%WF_MAK___________VisioProVolume
15_980f9e3e-f5a8-41c8-8596-61404addf677_NC%f%RB7-VP%f%48F-43F%f%YY-62%f%P3R-367%f%WK_Subscription__VisioStdRetail
15_44a1f6ff-0876-4edb-9169-dbb43101ee89_RX%f%63Y-4N%f%FK2-XTY%f%C8-C6%f%B3W-YPX%f%PJ_MAK___________VisioStdVolume
15_191509f2-6977-456f-ab30-cf0492b1e93a_NB%f%77V-RP%f%FQ6-PMM%f%KQ-T8%f%7DV-M4D%f%84_Retail________WordRetail
15_9cedef15-be37-4ff0-a08a-13a045540641_RP%f%HPB-Y7%f%NC4-3VY%f%FM-DW%f%7VD-G8Y%f%J8_MAK___________WordVolume
15_6337137e-7c07-4197-8986-bece6a76fc33_2P%f%3C9-BQ%f%NJH-VCV%f%PH-YD%f%Y6M-43J%f%PQ_Subscription__O365BusinessRetail
15_537ea5b5-7d50-4876-bd38-a53a77caca32_J2%f%W28-TN%f%9C8-26P%f%WV-F7%f%J4G-72X%f%CB_Subscription1_O365HomePremRetail
15_149dbce7-a48e-44db-8364-a53386cd4580_2N%f%382-D6%f%PKK-QTX%f%4D-2J%f%JYK-M96%f%P2_Subscription1_O365ProPlusRetail
15_bacd4614-5bef-4a5e-bafc-de4c788037a2_HN%f%8JP-87%f%TQJ-PBF%f%3P-Y6%f%6KC-W2K%f%9V_Subscription1_O365SmallBusPremRetail
16_bfa358b0-98f1-4125-842e-585fa13032e6_WH%f%K4N-YQ%f%GHB-XWX%f%CC-G3%f%HYC-6JF%f%94_Retail________AccessRetail
16_9d9faf9e-d345-4b49-afce-68cb0a539c7c_RN%f%B7V-P4%f%8F4-3FY%f%Y6-2P%f%3R3-63B%f%QV_PrepidBypass__AccessRuntimeRetail
16_3b2fa33f-cd5a-43a5-bd95-f49f3f546b0b_JJ%f%2Y4-N8%f%KM3-Y8K%f%Y3-Y2%f%2FR-R3K%f%VK_MAK___________AccessVolume
16_424d52ff-7ad2-4bc7-8ac6-748d767b455d_RK%f%JBN-VW%f%TM2-BDK%f%XX-RK%f%QFD-JTY%f%Q2_Retail________ExcelRetail
16_685062a7-6024-42e7-8c5f-6bb9e63e697f_FV%f%GNR-X8%f%2B2-6PR%f%JM-YT%f%4W7-8HV%f%36_MAK___________ExcelVolume
16_c02fb62e-1cd5-4e18-ba25-e0480467ffaa_2W%f%QNF-GB%f%K4B-XVG%f%6F-BB%f%MX7-M4F%f%2Y_OEM-Perp______HomeBusinessPipcRetail
16_86834d00-7896-4a38-8fae-32f20b86fa2b_HM%f%6FM-NV%f%F78-KV9%f%PM-F3%f%6B8-D9M%f%XD_Retail________HomeBusinessRetail
16_c28acdb8-d8b3-4199-baa4-024d09e97c99_PN%f%PRV-F2%f%627-Q8J%f%VC-3D%f%GR9-WTY%f%RK_Retail________HomeStudentRetail
16_e2127526-b60c-43e0-bed1-3c9dc3d5a468_YW%f%D4R-CN%f%KVT-VG8%f%VJ-93%f%33B-RC3%f%B8_Retail________HomeStudentVNextRetail
16_69ec9152-153b-471a-bf35-77ec88683eae_VN%f%WHF-FK%f%FBW-Q2R%f%GD-HY%f%HWF-R3H%f%H2_Subscription__MondoRetail
16_2cd0ea7e-749f-4288-a05e-567c573b2a6c_FM%f%TQQ-84%f%NR8-274%f%4R-MX%f%F4P-PGY%f%R3_MAK___________MondoVolume
16_436366de-5579-4f24-96db-3893e4400030_XY%f%NTG-R9%f%6FY-369%f%HX-YF%f%PHY-F9C%f%PM_Bypass________OneNoteFreeRetail
16_83ac4dd9-1b93-40ed-aa55-ede25bb6af38_FX%f%F6F-CN%f%C26-W64%f%3C-K6%f%KB7-6XX%f%W3_Retail________OneNoteRetail
16_23b672da-a456-4860-a8f3-e062a501d7e8_9T%f%YVN-D7%f%6HK-BVM%f%WT-Y7%f%G88-9TP%f%PV_MAK___________OneNoteVolume
16_5a670809-0983-4c2d-8aad-d3c2c5b7d5d1_7N%f%4KG-P2%f%QDH-86V%f%9C-DJ%f%FVF-369%f%W9_Retail________OutlookRetail
16_50059979-ac6f-4458-9e79-710bcb41721a_7Q%f%PNR-3H%f%FDG-YP6%f%T9-JQ%f%CKQ-KKX%f%XC_MAK___________OutlookVolume
16_5aab8561-1686-43f7-9ff5-2c861da58d17_9C%f%YB3-NF%f%MRW-YFD%f%G6-XC%f%7TF-BY3%f%6J_OEM-Perp______PersonalPipcRetail
16_a9f645a1-0d6a-4978-926a-abcb363b72a6_FT%f%7VF-XB%f%N92-HPD%f%JV-RH%f%MBY-6VK%f%BF_Retail________PersonalRetail
16_f32d1284-0792-49da-9ac6-deb2bc9c80b6_N7%f%GCB-WQ%f%T7K-QRH%f%WG-TT%f%PYD-7T9%f%XF_Retail________PowerPointRetail
16_9b4060c9-a7f5-4a66-b732-faf248b7240f_X3%f%RT9-ND%f%G64-VMK%f%2M-KQ%f%6XY-DPF%f%GV_MAK___________PowerPointVolume
16_de52bd50-9564-4adc-8fcb-a345c17f84f9_GM%f%43N-F7%f%42Q-6JD%f%DK-M6%f%22J-J8G%f%DV_Retail________ProPlusRetail
16_c47456e3-265d-47b6-8ca0-c30abbd0ca36_FN%f%VK8-8D%f%VCJ-F7X%f%3J-KG%f%VQB-RC2%f%QY_MAK___________ProPlusVolume
16_4e26cac1-e15a-4467-9069-cb47b67fe191_CF%f%9DD-6C%f%NW2-BJW%f%JQ-CV%f%CFX-Y7T%f%XD_OEM-Perp______ProfessionalPipcRetail
16_d64edc00-7453-4301-8428-197343fafb16_NX%f%FTK-YD%f%9Y7-X9M%f%MJ-9B%f%WM6-J2Q%f%VH_Retail________ProfessionalRetail
16_2f72340c-b555-418d-8b46-355944fe66b8_WP%f%Y8N-PD%f%PY4-FC7%f%TF-KM%f%P7P-KWY%f%FY_Subscription__ProjectProRetail
16_82f502b5-b0b0-4349-bd2c-c560df85b248_PK%f%C3N-8F%f%99H-28M%f%VY-J4%f%RYY-CWG%f%DH_MAK___________ProjectProVolume
16_16728639-a9ab-4994-b6d8-f81051e69833_JB%f%NPH-YF%f%2F7-Q9Y%f%29-86%f%CTG-C9Y%f%GV_MAKC2R________ProjectProXVolume
16_58d95b09-6af6-453d-a976-8ef0ae0316b1_NT%f%HQT-VK%f%K6W-BRB%f%87-HV%f%346-Y96%f%W8_Subscription__ProjectStdRetail
16_82e6b314-2a62-4e51-9220-61358dd230e6_4T%f%GWV-6N%f%9P6-G2H%f%8Y-2H%f%WKB-B4G%f%93_MAK___________ProjectStdVolume
16_431058f0-c059-44c5-b9e7-ed2dd46b6789_N3%f%W2Q-69%f%MBT-27R%f%D9-BH%f%8V3-JT2%f%C8_MAKC2R________ProjectStdXVolume
16_6e0c1d99-c72e-4968-bcb7-ab79e03e201e_WK%f%WND-X6%f%G9G-CDM%f%TV-CP%f%GYJ-6MV%f%BF_Retail________PublisherRetail
16_fcc1757b-5d5f-486a-87cf-c4d6dedb6032_9Q%f%VN2-PX%f%XRX-8V4%f%W8-Q7%f%926-TJG%f%D8_MAK___________PublisherVolume
16_9103f3ce-1084-447a-827e-d6097f68c895_6M%f%DN4-WF%f%3FV-4WH%f%3Q-W6%f%99V-RGC%f%MY_PrepidBypass__SkypeServiceBypassRetail
16_971cd368-f2e1-49c1-aedd-330909ce18b6_4N%f%4D8-3J%f%7Y3-YYW%f%7C-73%f%HD2-V8R%f%HY_PrepidBypass__SkypeforBusinessEntryRetail
16_418d2b9f-b491-4d7f-84f1-49e27cc66597_PB%f%J79-77%f%NY4-VRG%f%FG-Y8%f%WYC-CKC%f%RC_Retail________SkypeforBusinessRetail
16_03ca3b9a-0869-4749-8988-3cbc9d9f51bb_DM%f%TCJ-KN%f%RKR-JV8%f%TQ-V2%f%CR2-VFT%f%FH_MAK___________SkypeforBusinessVolume
16_4a31c291-3a12-4c64-b8ab-cd79212be45e_2F%f%PWN-4H%f%6CM-KD8%f%QQ-8H%f%CHC-P9X%f%YW_Retail________StandardRetail
16_0ed94aac-2234-4309-ba29-74bdbb887083_WH%f%GMQ-JN%f%MGT-MDQ%f%VF-WD%f%R69-KQB%f%WC_MAK___________StandardVolume
16_a56a3b37-3a35-4bbb-a036-eee5f1898eee_NV%f%K2G-2M%f%Y4G-7JX%f%2P-7D%f%6F2-VFQ%f%BR_Subscription__VisioProRetail
16_295b2c03-4b1c-4221-b292-1411f468bd02_NR%f%KT9-C8%f%GP2-XDY%f%XQ-YW%f%72K-MG9%f%2B_MAK___________VisioProVolume
16_0594dc12-8444-4912-936a-747ca742dbdb_G9%f%8Q2-B6%f%N77-CFH%f%9J-K8%f%24G-XQC%f%C4_MAKC2R________VisioProXVolume
16_980f9e3e-f5a8-41c8-8596-61404addf677_NC%f%RB7-VP%f%48F-43F%f%YY-62%f%P3R-367%f%WK_Subscription__VisioStdRetail
16_44151c2d-c398-471f-946f-7660542e3369_XN%f%CJB-YY%f%883-JRW%f%64-DP%f%XMX-JXC%f%R6_MAK___________VisioStdVolume
16_1d1c6879-39a3-47a5-9a6d-aceefa6a289d_B2%f%HTN-JP%f%H8C-J6Y%f%6V-HC%f%HKB-43M%f%GT_MAKC2R________VisioStdXVolume
16_cacaa1bf-da53-4c3b-9700-11738ef1c2a5_P8%f%K82-NQ%f%7GG-JKY%f%8T-6V%f%HVY-88G%f%GD_Retail________WordRetail
16_c3000759-551f-4f4a-bcac-a4b42cbf1de2_YH%f%MWC-YN%f%6V9-WJP%f%XD-3W%f%QKP-TMV%f%CV_MAK___________WordVolume
16_518687bd-dc55-45b9-8fa6-f918e1082e83_WR%f%YJ6-G3%f%NP7-7VH%f%94-8X%f%7KP-JB7%f%HC_Retail________Access2019Retail
16_385b91d6-9c2c-4a2e-86b5-f44d44a48c5f_6F%f%WHX-NK%f%YXK-BW3%f%4Q-7X%f%C9F-Q9P%f%X7_MAK-AE________Access2019Volume
16_22e6b96c-1011-4cd5-8b35-3c8fb6366b86_FG%f%QNJ-JW%f%JCG-7Q8%f%MG-RM%f%RGJ-9TQ%f%VF_PrepidBypass__AccessRuntime2019Retail
16_c201c2b7-02a1-41a8-b496-37c72910cd4a_KB%f%PNW-64%f%CMM-8KW%f%CB-23%f%F44-8B7%f%HM_Retail________Excel2019Retail
16_05cb4e1d-cc81-45d5-a769-f34b09b9b391_8N%f%T4X-GQ%f%MCK-62X%f%4P-TW%f%6QP-YKP%f%YF_MAK-AE________Excel2019Volume
16_7fe09eef-5eed-4733-9a60-d7019df11cac_QB%f%N2Y-9B%f%284-9KW%f%78-K4%f%8PB-R62%f%YT_Retail________HomeBusiness2019Retail
16_4539aa2c-5c31-4d47-9139-543a868e5741_XN%f%WPM-32%f%XQC-Y7Q%f%JC-QG%f%GBV-YY7%f%JK_Retail________HomeStudent2019Retail
16_20e359d5-927f-47c0-8a27-38adbdd27124_WR%f%43D-NM%f%WQQ-HCQ%f%R2-VK%f%XDR-37B%f%7H_Retail________Outlook2019Retail
16_92a99ed8-2923-4cb7-a4c5-31da6b0b8cf3_RN%f%3QB-GT%f%6D7-YB3%f%VH-F3%f%RPB-3GQ%f%YB_MAK-AE________Outlook2019Volume
16_2747b731-0f1f-413e-a92d-386ec1277dd8_NM%f%BY8-V3%f%CV7-BX6%f%K6-29%f%22Y-43M%f%7T_Retail________Personal2019Retail
16_7e63cc20-ba37-42a1-822d-d5f29f33a108_HN%f%27K-JH%f%J8R-7T7%f%KK-WJ%f%YC3-FM7%f%MM_Retail________PowerPoint2019Retail
16_13c2d7bf-f10d-42eb-9e93-abf846785434_29%f%GNM-VM%f%33V-WR2%f%3K-HG%f%2DT-KTQ%f%YR_MAK-AE________PowerPoint2019Volume
16_a3072b8f-adcc-4e75-8d62-fdeb9bdfae57_BN%f%4XJ-R9%f%DYY-96W%f%48-YK%f%8DM-MY7%f%PY_Retail________ProPlus2019Retail
16_6755c7a7-4dfe-46f5-bce8-427be8e9dc62_T8%f%YBN-4Y%f%V3X-KK2%f%4Q-QX%f%BD7-T3C%f%63_MAK-AE________ProPlus2019Volume
16_1717c1e0-47d3-4899-a6d3-1022db7415e0_9N%f%XDK-MR%f%Y98-2VJ%f%V8-GF%f%73J-TQ9%f%FK_Retail________Professional2019Retail
16_0d270ef7-5aaf-4370-a372-bc806b96adb7_JD%f%TNC-PP%f%77T-T9H%f%2W-G4%f%J2J-VH8%f%JK_Retail________ProjectPro2019Retail
16_d4ebadd6-401b-40d5-adf4-a5d4accd72d1_TB%f%XBD-FN%f%WKJ-WRH%f%BD-KB%f%PHH-XD9%f%F2_MAK-AE________ProjectPro2019Volume
16_bb7ffe5f-daf9-4b79-b107-453e1c8427b5_R3%f%JNT-8P%f%BDP-MTW%f%CK-VD%f%2V8-HMK%f%F9_Retail________ProjectStd2019Retail
16_fdaa3c03-dc27-4a8d-8cbf-c3d843a28ddc_RB%f%RFX-MQ%f%NDJ-4XF%f%HF-7Q%f%VDR-JHX%f%GC_MAK-AE________ProjectStd2019Volume
16_f053a7c7-f342-4ab8-9526-a1d6e5105823_4Q%f%C36-NW%f%3YH-D2Y%f%9D-RJ%f%PC7-VVB%f%9D_Retail________Publisher2019Retail
16_40055495-be00-444e-99cc-07446729b53e_K8%f%F2D-NB%f%M32-BF2%f%6V-YC%f%KFJ-29Y%f%9W_MAK-AE________Publisher2019Volume
16_b639e55c-8f3e-47fe-9761-26c6a786ad6b_JB%f%DKF-6N%f%CD6-49K%f%3G-2T%f%V79-BKP%f%73_Retail________SkypeforBusiness2019Retail
16_15a430d4-5e3f-4e6d-8a0a-14bf3caee4c7_9M%f%NQ7-YP%f%Q3B-6WJ%f%XM-G8%f%3T3-CBB%f%DK_MAK-AE________SkypeforBusiness2019Volume
16_f88cfdec-94ce-4463-a969-037be92bc0e7_N9%f%722-BV%f%9H6-WTJ%f%TT-FP%f%B93-978%f%MK_PrepidBypass__SkypeforBusinessEntry2019Retail
16_fdfa34dd-a472-4b85-bee6-cf07bf0aaa1c_ND%f%GVM-MD%f%27H-2XH%f%VC-KD%f%DX2-YKP%f%74_Retail________Standard2019Retail
16_beb5065c-1872-409e-94e2-403bcfb6a878_NT%f%3V6-XM%f%BK7-Q66%f%MF-VM%f%KR4-FC3%f%3M_MAK-AE________Standard2019Volume
16_a6f69d68-5590-4e02-80b9-e7233dff204e_2N%f%WVW-QG%f%F4T-9CP%f%MB-WY%f%DQ9-7XP%f%79_Retail________VisioPro2019Retail
16_f41abf81-f409-4b0d-889d-92b3e3d7d005_33%f%YF4-GN%f%CQ3-J6G%f%DM-J6%f%7P3-FM7%f%QP_MAK-AE________VisioPro2019Volume
16_4a582021-18c2-489f-9b3d-5186de48f1cd_26%f%3WK-3N%f%797-7R4%f%37-28%f%BKG-3V8%f%M8_Retail________VisioStd2019Retail
16_933ed0e3-747d-48b0-9c2c-7ceb4c7e473d_BG%f%NHX-QT%f%PRJ-F9C%f%9G-R8%f%QQG-8T2%f%7F_MAK-AE________VisioStd2019Volume
16_72cee1c2-3376-4377-9f25-4024b6baadf8_JX%f%R8H-NJ%f%3MK-X66%f%W8-78%f%CWD-QRV%f%R2_Retail________Word2019Retail
16_fe5fe9d5-3b06-4015-aa35-b146f85c4709_9F%f%36R-PN%f%VHH-3DX%f%GQ-7C%f%D2H-R9D%f%3V_MAK-AE________Word2019Volume
16_f634398e-af69-48c9-b256-477bea3078b5_P2%f%86B-N3%f%XYP-36Q%f%RQ-29%f%CMP-RVX%f%9M_Retail________Access2021Retail
16_ae17db74-16b0-430b-912f-4fe456e271db_JB%f%H3N-P9%f%7FP-FRT%f%JD-MG%f%K2C-VFW%f%G6_MAK-AE________Access2021Volume
16_fb099c19-d48b-4a2f-a160-4383011060aa_V6%f%QFB-7N%f%7G9-PF7%f%W9-M8%f%FQM-MY8%f%G9_Retail________Excel2021Retail
16_9da1ecdb-3a62-4273-a234-bf6d43dc0778_WN%f%YR4-KM%f%R9H-KVC%f%8W-7H%f%J8B-K79%f%DQ_MAK-AE________Excel2021Volume
16_38b92b63-1dff-4be7-8483-2a839441a2bc_JM%f%99N-4M%f%MD8-DQC%f%GJ-VM%f%YFY-R63%f%YK_Subscription__HomeBusiness2021Retail
16_2f258377-738f-48dd-9397-287e43079958_N3%f%CWD-38%f%XVH-KRX%f%2Y-YR%f%P74-6RB%f%B2_Subscription__HomeStudent2021Retail
16_279706f4-3a4b-4877-949b-f8c299cf0cc5_NB%f%2TQ-3Y%f%79C-77C%f%6M-QM%f%Y7H-7QY%f%8P_Retail________OneNote2021Retail
16_ecea2cfa-d406-4a7f-be0d-c6163250d126_4N%f%CWR-9V%f%92Y-34V%f%B2-RP%f%THR-YTG%f%R7_Retail________Outlook2021Retail
16_45bf67f9-0fc8-4335-8b09-9226cef8a576_JQ%f%9MJ-QY%f%N6B-67P%f%X9-GY%f%FVY-QJ6%f%TB_MAK-AE________Outlook2021Volume
16_8f89391e-eedb-429d-af90-9d36fbf94de6_RR%f%RYB-DN%f%749-GCP%f%W4-9H%f%6VK-HCH%f%PT_Retail________Personal2021Retail
16_c9bf5e86-f5e3-4ac6-8d52-e114a604d7bf_3K%f%XXQ-PV%f%N2C-8P7%f%YY-HC%f%V88-GVM%f%96_Retail1_______PowerPoint2021Retail
16_716f2434-41b6-4969-ab73-e61e593a3875_39%f%G2N-3B%f%D9C-C4X%f%CM-BD%f%4QG-FVY%f%DY_MAK-AE________PowerPoint2021Volume
16_c2f04adf-a5de-45c5-99a5-f5fddbda74a8_8W%f%XTP-MN%f%628-KY4%f%4G-VJ%f%WCK-C7P%f%CF_Retail________ProPlus2021Retail
16_3f180b30-9b05-4fe2-aa8d-0c1c4790f811_RN%f%HJY-DT%f%FXW-HW9%f%F8-49%f%82D-MD2%f%CW_MAK-AE1_______ProPlus2021Volume
16_96097a68-b5c5-4b19-8600-2e8d6841a0db_JR%f%JNJ-33%f%M7C-R73%f%X3-P9%f%XF7-R9F%f%6M_MAK-AE________ProPlusSPLA2021Volume
16_711e48a6-1a79-4b00-af10-73f4ca3aaac4_DJ%f%PHV-NC%f%JV6-GWP%f%T6-K2%f%6JX-C7P%f%BG_Retail________Professional2021Retail
16_3747d1d5-55a8-4bc3-b53d-19fff1913195_QK%f%HNX-M9%f%GGH-T3Q%f%MW-YP%f%K4Q-QRW%f%MV_Retail________ProjectPro2021Retail
16_17739068-86c4-4924-8633-1e529abc7efc_HV%f%C34-CV%f%NPG-RVC%f%MT-X2%f%JRF-CR7%f%RK_MAK-AE1_______ProjectPro2021Volume
16_4ea64dca-227c-436b-813f-b6624be2d54c_2B%f%96V-X9%f%NJY-WFB%f%RC-Q8%f%MP2-7CH%f%RR_Retail________ProjectStd2021Retail
16_84313d1e-47c8-4e27-8ced-0476b7ee46c4_3C%f%NQX-T3%f%4TY-99R%f%H4-C4%f%YD2-KW6%f%WH_MAK-AE________ProjectStd2021Volume
16_b769b746-53b1-4d89-8a68-41944dafe797_CD%f%NFG-77%f%T8D-VKQ%f%JX-B7%f%KT3-KK2%f%8V_Retail1_______Publisher2021Retail
16_a0234cfe-99bd-4586-a812-4f296323c760_2K%f%XJH-3N%f%HTW-RDB%f%PX-QF%f%RXJ-MTG%f%XF_MAK-AE________Publisher2021Volume
16_c3fb48b2-1fd4-4dc8-af39-819edf194288_DV%f%BXN-HF%f%T43-CVP%f%RQ-J8%f%9TF-VMM%f%HG_Retail________SkypeforBusiness2021Retail
16_6029109c-ceb8-4ee5-b324-f8eb2981e99a_R3%f%FCY-NH%f%GC7-CBP%f%VP-8Q%f%934-YTG%f%XG_MAK-AE________SkypeforBusiness2021Volume
16_9e7e7b8e-a0e7-467b-9749-d0de82fb7297_HX%f%NXB-J4%f%JGM-TCF%f%44-2X%f%2CV-FJV%f%VH_Retail________Standard2021Retail
16_223a60d8-9002-4a55-abac-593f5b66ca45_2C%f%JN4-C9%f%XK2-HFP%f%Q6-YH%f%498-82T%f%XH_MAK-AE________Standard2021Volume
16_b99ba8c4-e257-4b70-a31a-8bd308ce7073_BQ%f%WDW-NJ%f%9YF-P7Y%f%79-H6%f%DCT-MKQ%f%9C_MAK-AE________StandardSPLA2021Volume
16_814014d3-c30b-4f63-a493-3708e0dc0ba8_T6%f%P26-NJ%f%VBR-76B%f%K8-WB%f%CDY-TX3%f%BC_Retail________VisioPro2021Retail
16_c590605a-a08a-4cc7-8dc2-f1ffb3d06949_JN%f%KBX-MH%f%9P4-K8Y%f%YV-8C%f%G2Y-VQ2%f%C8_MAK-AE________VisioPro2021Volume
16_16d43989-a5ef-47e2-9ff1-272784caee24_89%f%NYY-KB%f%93R-7X2%f%2F-93%f%QDF-DJ6%f%YM_Retail________VisioStd2021Retail
16_d55f90ee-4ba2-4d02-b216-1300ee50e2af_BW%f%43B-4P%f%NFP-V63%f%7F-23%f%TR2-J47%f%TX_MAK-AE________VisioStd2021Volume
16_fb33d997-4aa3-494e-8b58-03e9ab0f181d_VN%f%CC4-CJ%f%QVK-BKX%f%34-77%f%Y8H-CYX%f%MR_Retail________Word2021Retail
16_0c728382-95fb-4a55-8f12-62e605f91727_BJ%f%G97-NW%f%3GM-8QQ%f%Q7-FH%f%76G-686%f%XM_MAK-AE________Word2021Volume
16_8fdb1f1e-663f-4f2e-8fdb-7c35aee7d5ea_GN%f%XWX-DF%f%797-B2J%f%T3-82%f%W27-KHP%f%XT_MAK-AE________ProPlus2024Volume-Preview
16_33b11b14-91fd-4f7b-b704-e64a055cf601_X8%f%6XX-N3%f%QMW-B4W%f%GQ-QC%f%B69-V26%f%KW_MAK_AE________ProjectPro2024Volume-Preview
16_eb074198-7384-4bdd-8e6c-c3342dac8435_DW%f%99Y-H7%f%NT6-6B2%f%9D-8J%f%Q8F-R3Q%f%T7_MAK_AE________VisioPro2024Volume-Preview
16_6337137e-7c07-4197-8986-bece6a76fc33_2P%f%3C9-BQ%f%NJH-VCV%f%PH-YD%f%Y6M-43J%f%PQ_Subscription__O365BusinessRetail
16_2f5c71b4-5b7a-4005-bb68-f9fac26f2ea3_W6%f%2NQ-26%f%7QR-RTF%f%74-PF%f%2MH-JQM%f%TH_Subscription__O365EduCloudRetail
16_537ea5b5-7d50-4876-bd38-a53a77caca32_J2%f%W28-TN%f%9C8-26P%f%WV-F7%f%J4G-72X%f%CB_Subscription1_O365HomePremRetail
16_149dbce7-a48e-44db-8364-a53386cd4580_2N%f%382-D6%f%PKK-QTX%f%4D-2J%f%JYK-M96%f%P2_Subscription1_O365ProPlusRetail
16_bacd4614-5bef-4a5e-bafc-de4c788037a2_HN%f%8JP-87%f%TQJ-PBF%f%3P-Y6%f%6KC-W2K%f%9V_Subscription1_O365SmallBusPremRetail
) do (
for /f "tokens=1-5 delims=_" %%A in ("%%#") do (

if %1==getinfo if not defined _key (
if %oVer%==%%A if /i "%2"=="%%E" (
set _key=%%C
set _actid=%%B
set _allactid=!_allactid! %%B
set _lic=%%D
if %oVer%==16 (echo "%%D" | find /i "Subscription" %nul% && set _sublic=1)
)
)

if %1==getmsiprod if %oVer%==%%A (
find /i "%%E" %msitemp% %nul% && (
if defined _oIds (set _oIds=!_oIds! %%E) else (set _oIds=%%E)
)
)

)
)
exit /b

::========================================================================================================================================

::  此代码用于修改 sppc dll 文件的时间戳值以更改校验和
::  这样做是为了降低防病毒软件的潜在误报检测。在每次安装时，它将安装一个唯一的 sppc dll 文件

:oh_extractdll

set b=
%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':%_hook%\:.*';$bytes = [Con%b%vert]::FromBas%b%e64String($f[1]); $PePath='%1'; $offset='%2'; $m=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':hexedit\:.*';iex ($m[1]);" %nul2% | find /i "Error found" %nul1% && set hasherror=1
exit /b

:hexedit:
# 使用内存流对字节执行操作
$MemoryStream = New-Object System.IO.MemoryStream
$Writer = New-Object System.IO.BinaryWriter($MemoryStream)
$Writer.Write($bytes)

# 定义动态程序集、模块和类型
$AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1)
$ModuleBuilder = $AssemblyBuilder.DefineDynamicModule(2, $False)
$TypeBuilder = $ModuleBuilder.DefineType(0)

# 定义 P/调用方法
[void]$TypeBuilder.DefinePInvokeMethod('MapFileAndCheckSum', 'imagehlp.dll', 'Public, Static', [Reflection.CallingConventions]::Standard, [int], @([string], [int].MakeByRefType(), [int].MakeByRefType()), [Runtime.InteropServices.CallingConvention]::Winapi, [Runtime.InteropServices.CharSet]::Auto)

# 创建类型
$Imagehlp = $TypeBuilder.CreateType()

# 偏移信息
$timestampOffset = 136
$exportTimestampOffset = $offset
$checkSumOffset = 216

# 计算时间戳
$currentTimestamp = [DateTime]::UtcNow
$unixTimestamp = [int]($currentTimestamp - (Get-Date -Year 1970 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0)).TotalSeconds

# 更改时间戳
$Writer.BaseStream.Position = $timestampOffset
$Writer.Write($unixTimestamp)

$Writer.BaseStream.Position = $exportTimestampOffset
$Writer.Write($unixTimestamp)

$Writer.Flush()

# 将内存流的当前状态写入临时文件
$tempFilePath = [System.IO.Path]::Combine($env:windir, "Temp", [System.IO.Path]::GetRandomFileName())
[System.IO.File]::WriteAllBytes($tempFilePath, $MemoryStream.ToArray())

# 使用临时文件更新哈希
[int]$HeaderSum = 0
[int]$CheckSum = 0
[void]$Imagehlp::MapFileAndCheckSum($tempFilePath, [ref]$HeaderSum, [ref]$CheckSum)

# 如果校验和不匹配，请更新内存流中的校验和
if ($HeaderSum -ne $CheckSum) {
    $Writer.BaseStream.Position = $checkSumOffset
    $Writer.Write($CheckSum)
    $Writer.Flush()
} else {
    Write-host 发现错误
}

# 删除临时文件
Remove-Item -Path $tempFilePath -Force

# 获取修改后的字节
$modifiedBytes = $MemoryStream.ToArray()

# 将修改后的字节写入最终文件
[System.IO.File]::WriteAllBytes($PePath, $modifiedBytes)

[void]$Imagehlp::MapFileAndCheckSum($PePath, [ref]$HeaderSum, [ref]$CheckSum)
if ($HeaderSum -ne $CheckSum) {
    Write-host 发现错误
}

$MemoryStream.Close()
:hexedit:

::========================================================================================================================================
::
::  下面的文本块以 base64 格式编码
::  标签“sppc64.dll”和“sppc32.dll”中的块包含以下文件
::
::  e6ac83560c19ec7eb868c50ea97ea0ed5632a397a9f43c17e24e6de4a694d118 *sppc32.dll
::  c6df24deef2e83813dee9c81ddd9793a3d60c117a4e8e231b82e32b3192927e7 *sppc64.dll
::
::  The files are encoded in base64 to make MAS AIO version.
::
::  mass grave[.]dev/ohook
::  在这里，你可以找到文件源代码以及如何重建相同的 sppc.dll 文件的信息
::
::  stackoverflow.com/a/35335273
::  在这里，你可以检查如何从 base64 中提取 sppc.dll 文件
::
::  如有任何其他问题，请随时通过以下方式与我们联系 mass grave[.]dev/contactus
::
::========================================================================================================================================

:sppc32.dll:
TVqQAAMAAAAEAAAA//8AALgAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAA4fug4AtAnNIbgBTM0hVGhpcyBwcm9ncmFtIGNhbm5vdCBiZSBydW4gaW4gRE9TIG1vZGUuDQ0KJAAAAAAAAABQRQAATAEHAMDc0GQAAAAAAAAAAOAA
DiMLAQIoAAIAAAAeAAAAAAAAABAAAAAQAAAAAAAAAACAagAQAAAAAgAABAAAAAEAAAAGAAAAAAAAAACQAAAABAAAi9MAAAIAQAEAACAAABAAAAAAEAAAEAAAAAAAABAAAAAAQAAAjRAAAABgAAAYAQAAAHAAAHgDAAAAAAAAAAAAAAAAAAAAAAAAAIAAABQAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABsYAAAHAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC50ZXh0AAAAcAEAAAAQAAAAAgAAAAQAAAAAAAAAAAAAAAAAACAAAGAucmRhdGEAABgAAAAAIAAAAAIAAAAGAAAAAAAAAAAAAAAA
AABAAABALmVoX2ZyYW2AAAAAADAAAAACAAAACAAAAAAAAAAAAAAAAAAAQAAAQC5lZGF0YQAAjRAAAABAAAAAEgAAAAoAAAAAAAAAAAAAAAAAAEAAAEAuaWRhdGEAABgBAAAAYAAAAAIAAAAcAAAAAAAAAAAAAAAAAABAAADALnJzcmMAAAB4AwAAAHAAAAAEAAAAHgAA
AAAAAAAAAAAAAAAAQAAAwC5yZWxvYwAAFAAAAACAAAAAAgAAACIAAAAAAAAAAAAAAAAAAEAAAEIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALgBAAAAwgwAVYnlVlONRfCD7DDHRfAAAAAA
iUQkFI1F9IlEJBCLRQzHRCQMAAAAAIlEJASLRQjHRCQIACCAaokEJMdF9AAAAADoAgEAAIs1eGCAaoPsGIXAicOLRfB0CokEJDHb/9ZR6zKLVfTHRCQECiCAaokEJIlUJAj/FYBggGqD7AyFwItF8IkEJHQK/9a7AQAAAFLrA//WUI1l+InYW15dw1WJ5VdWU4PsPItF
GIt1HIlEJBCLRRSJdCQUiUQkDItFEIlEJAiLRQyJRCQEi0UIiQQk6HwAAAAxyYPsGInHhcB1XItFGDkIdlVr2SiLBgHYg3gQAHRFiUQkBItFCIlN5IkEJOj7/v//i03khcB1LAMex0MQAQAAAMdDFAAAAADHQxgAAAAAx0McAAAAAMdDIAAAAADHQyQAAAAAQeukjWX0
ifhbXl9dwhgAkP8lcGCAapCQ/yVsYIBqkJD/////AAAAAP////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAATgBhAG0AZQAAAEcAcgBhAGMAZQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAUAAAAAAAAAAF6UgABfAgBGwwEBIgBAAAQAAAAHAAAAODf//8IAAAAAAAAACQAAAAwAAAA
1N///50AAAAAQQ4IhQJCDQVIhgODBAKPw0HGQcUMBAQoAAAAWAAAAEng//+qAAAAAEEOCIUCQg0FRocDhgSDBQKbw0HGQcdBxQwEBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAADA3NBkAAAAAMZCAAABAAAAQwAAAEMAAAAoQAAANEEAAEBCAADPQgAA70IAAAVDAAApQwAAXUMAAKFDAADpQwAAF0QAADVEAABnRAAAnUQAAONEAAAtRQAAYUUAAJ9FAADTRQAADUYAADtGAABxRgAAr0YAAM9GAAD7RgAApRAAAFFHAABvRwAA
n0cAANNHAAARSAAATUgAAG9IAAClSAAAzUgAAAVJAABBSQAAbUkAAKdJAAC7SQAA+0kAADlKAABPSgAAdUoAAJ1KAADTSgAAB0sAAD1LAABpSwAApUsAAONLAAANTAAAOUwAAIlMAADRTAAAEU0AAFlNAACjTQAA8U0AABtOAABHTgAAh04AALtOAADnTgAAK08AAFtP
AAC1TwAA608AACdQAABdUAAA4kIAAP1CAAAaQwAARkMAAIJDAADIQwAAA0QAAClEAABRRAAAhUQAAMNEAAALRQAASkUAAINFAAC8RQAA80UAACdGAABZRgAAk0YAAMJGAADoRgAAGUcAADFHAABjRwAAikcAALxHAAD1RwAAMkgAAGFIAACNSAAAvEgAAOxIAAAmSQAA
WkkAAI1JAAC0SQAA3kkAAB1KAABHSgAAZUoAAIxKAAC7SgAA8EoAACVLAABWSwAAiksAAMdLAAD7SwAAJkwAAGRMAACwTAAA9EwAADhNAACBTQAAzU0AAAlOAAA0TgAAak4AAKROAADUTgAADE8AAEZPAACLTwAA008AAAxQAABFUAAAeFAAAAAAAQACAAMABAAFAAYA
BwAIAAkACgALAAwADQAOAA8AEAARABIAEwAUABUAFgAXABgAGQAaABsAHAAdAB4AHwAgACEAIgAjACQAJQAmACcAKAApACoAKwAsAC0ALgAvADAAMQAyADMANAA1ADYANwA4ADkAOgA7ADwAPQA+AD8AQABBAEIAc3BwYy5kbGwAU1BQQ1MuU0xDYWxsU2VydmVyAFNM
Q2FsbFNlcnZlcgBTUFBDUy5TTENsb3NlAFNMQ2xvc2UAU1BQQ1MuU0xDb25zdW1lUmlnaHQAU0xDb25zdW1lUmlnaHQAU1BQQ1MuU0xEZXBvc2l0TWlncmF0aW9uQmxvYgBTTERlcG9zaXRNaWdyYXRpb25CbG9iAFNQUENTLlNMRGVwb3NpdE9mZmxpbmVDb25maXJt
YXRpb25JZABTTERlcG9zaXRPZmZsaW5lQ29uZmlybWF0aW9uSWQAU1BQQ1MuU0xEZXBvc2l0T2ZmbGluZUNvbmZpcm1hdGlvbklkRXgAU0xEZXBvc2l0T2ZmbGluZUNvbmZpcm1hdGlvbklkRXgAU1BQQ1MuU0xEZXBvc2l0U3RvcmVUb2tlbgBTTERlcG9zaXRTdG9y
ZVRva2VuAFNQUENTLlNMRmlyZUV2ZW50AFNMRmlyZUV2ZW50AFNQUENTLlNMR2F0aGVyTWlncmF0aW9uQmxvYgBTTEdhdGhlck1pZ3JhdGlvbkJsb2IAU1BQQ1MuU0xHYXRoZXJNaWdyYXRpb25CbG9iRXgAU0xHYXRoZXJNaWdyYXRpb25CbG9iRXgAU1BQQ1MuU0xH
ZW5lcmF0ZU9mZmxpbmVJbnN0YWxsYXRpb25JZABTTEdlbmVyYXRlT2ZmbGluZUluc3RhbGxhdGlvbklkAFNQUENTLlNMR2VuZXJhdGVPZmZsaW5lSW5zdGFsbGF0aW9uSWRFeABTTEdlbmVyYXRlT2ZmbGluZUluc3RhbGxhdGlvbklkRXgAU1BQQ1MuU0xHZXRBY3Rp
dmVMaWNlbnNlSW5mbwBTTEdldEFjdGl2ZUxpY2Vuc2VJbmZvAFNQUENTLlNMR2V0QXBwbGljYXRpb25JbmZvcm1hdGlvbgBTTEdldEFwcGxpY2F0aW9uSW5mb3JtYXRpb24AU1BQQ1MuU0xHZXRBcHBsaWNhdGlvblBvbGljeQBTTEdldEFwcGxpY2F0aW9uUG9saWN5
AFNQUENTLlNMR2V0QXV0aGVudGljYXRpb25SZXN1bHQAU0xHZXRBdXRoZW50aWNhdGlvblJlc3VsdABTUFBDUy5TTEdldEVuY3J5cHRlZFBJREV4AFNMR2V0RW5jcnlwdGVkUElERXgAU1BQQ1MuU0xHZXRHZW51aW5lSW5mb3JtYXRpb24AU0xHZXRHZW51aW5lSW5m
b3JtYXRpb24AU1BQQ1MuU0xHZXRJbnN0YWxsZWRQcm9kdWN0S2V5SWRzAFNMR2V0SW5zdGFsbGVkUHJvZHVjdEtleUlkcwBTUFBDUy5TTEdldExpY2Vuc2UAU0xHZXRMaWNlbnNlAFNQUENTLlNMR2V0TGljZW5zZUZpbGVJZABTTEdldExpY2Vuc2VGaWxlSWQAU1BQ
Q1MuU0xHZXRMaWNlbnNlSW5mb3JtYXRpb24AU0xHZXRMaWNlbnNlSW5mb3JtYXRpb24AU0xHZXRMaWNlbnNpbmdTdGF0dXNJbmZvcm1hdGlvbgBTUFBDUy5TTEdldFBLZXlJZABTTEdldFBLZXlJZABTUFBDUy5TTEdldFBLZXlJbmZvcm1hdGlvbgBTTEdldFBLZXlJ
bmZvcm1hdGlvbgBTUFBDUy5TTEdldFBvbGljeUluZm9ybWF0aW9uAFNMR2V0UG9saWN5SW5mb3JtYXRpb24AU1BQQ1MuU0xHZXRQb2xpY3lJbmZvcm1hdGlvbkRXT1JEAFNMR2V0UG9saWN5SW5mb3JtYXRpb25EV09SRABTUFBDUy5TTEdldFByb2R1Y3RTa3VJbmZv
cm1hdGlvbgBTTEdldFByb2R1Y3RTa3VJbmZvcm1hdGlvbgBTUFBDUy5TTEdldFNMSURMaXN0AFNMR2V0U0xJRExpc3QAU1BQQ1MuU0xHZXRTZXJ2aWNlSW5mb3JtYXRpb24AU0xHZXRTZXJ2aWNlSW5mb3JtYXRpb24AU1BQQ1MuU0xJbnN0YWxsTGljZW5zZQBTTElu
c3RhbGxMaWNlbnNlAFNQUENTLlNMSW5zdGFsbFByb29mT2ZQdXJjaGFzZQBTTEluc3RhbGxQcm9vZk9mUHVyY2hhc2UAU1BQQ1MuU0xJbnN0YWxsUHJvb2ZPZlB1cmNoYXNlRXgAU0xJbnN0YWxsUHJvb2ZPZlB1cmNoYXNlRXgAU1BQQ1MuU0xJc0dlbnVpbmVMb2Nh
bEV4AFNMSXNHZW51aW5lTG9jYWxFeABTUFBDUy5TTExvYWRBcHBsaWNhdGlvblBvbGljaWVzAFNMTG9hZEFwcGxpY2F0aW9uUG9saWNpZXMAU1BQQ1MuU0xPcGVuAFNMT3BlbgBTUFBDUy5TTFBlcnNpc3RBcHBsaWNhdGlvblBvbGljaWVzAFNMUGVyc2lzdEFwcGxp
Y2F0aW9uUG9saWNpZXMAU1BQQ1MuU0xQZXJzaXN0UlRTUGF5bG9hZE92ZXJyaWRlAFNMUGVyc2lzdFJUU1BheWxvYWRPdmVycmlkZQBTUFBDUy5TTFJlQXJtAFNMUmVBcm0AU1BQQ1MuU0xSZWdpc3RlckV2ZW50AFNMUmVnaXN0ZXJFdmVudABTUFBDUy5TTFJlZ2lz
dGVyUGx1Z2luAFNMUmVnaXN0ZXJQbHVnaW4AU1BQQ1MuU0xTZXRBdXRoZW50aWNhdGlvbkRhdGEAU0xTZXRBdXRoZW50aWNhdGlvbkRhdGEAU1BQQ1MuU0xTZXRDdXJyZW50UHJvZHVjdEtleQBTTFNldEN1cnJlbnRQcm9kdWN0S2V5AFNQUENTLlNMU2V0R2VudWlu
ZUluZm9ybWF0aW9uAFNMU2V0R2VudWluZUluZm9ybWF0aW9uAFNQUENTLlNMVW5pbnN0YWxsTGljZW5zZQBTTFVuaW5zdGFsbExpY2Vuc2UAU1BQQ1MuU0xVbmluc3RhbGxQcm9vZk9mUHVyY2hhc2UAU0xVbmluc3RhbGxQcm9vZk9mUHVyY2hhc2UAU1BQQ1MuU0xV
bmxvYWRBcHBsaWNhdGlvblBvbGljaWVzAFNMVW5sb2FkQXBwbGljYXRpb25Qb2xpY2llcwBTUFBDUy5TTFVucmVnaXN0ZXJFdmVudABTTFVucmVnaXN0ZXJFdmVudABTUFBDUy5TTFVucmVnaXN0ZXJQbHVnaW4AU0xVbnJlZ2lzdGVyUGx1Z2luAFNQUENTLlNMcEF1
dGhlbnRpY2F0ZUdlbnVpbmVUaWNrZXRSZXNwb25zZQBTTHBBdXRoZW50aWNhdGVHZW51aW5lVGlja2V0UmVzcG9uc2UAU1BQQ1MuU0xwQmVnaW5HZW51aW5lVGlja2V0VHJhbnNhY3Rpb24AU0xwQmVnaW5HZW51aW5lVGlja2V0VHJhbnNhY3Rpb24AU1BQQ1MuU0xw
Q2xlYXJBY3RpdmF0aW9uSW5Qcm9ncmVzcwBTTHBDbGVhckFjdGl2YXRpb25JblByb2dyZXNzAFNQUENTLlNMcERlcG9zaXREb3dubGV2ZWxHZW51aW5lVGlja2V0AFNMcERlcG9zaXREb3dubGV2ZWxHZW51aW5lVGlja2V0AFNQUENTLlNMcERlcG9zaXRUb2tlbkFj
dGl2YXRpb25SZXNwb25zZQBTTHBEZXBvc2l0VG9rZW5BY3RpdmF0aW9uUmVzcG9uc2UAU1BQQ1MuU0xwR2VuZXJhdGVUb2tlbkFjdGl2YXRpb25DaGFsbGVuZ2UAU0xwR2VuZXJhdGVUb2tlbkFjdGl2YXRpb25DaGFsbGVuZ2UAU1BQQ1MuU0xwR2V0R2VudWluZUJs
b2IAU0xwR2V0R2VudWluZUJsb2IAU1BQQ1MuU0xwR2V0R2VudWluZUxvY2FsAFNMcEdldEdlbnVpbmVMb2NhbABTUFBDUy5TTHBHZXRMaWNlbnNlQWNxdWlzaXRpb25JbmZvAFNMcEdldExpY2Vuc2VBY3F1aXNpdGlvbkluZm8AU1BQQ1MuU0xwR2V0TVNQaWRJbmZv
cm1hdGlvbgBTTHBHZXRNU1BpZEluZm9ybWF0aW9uAFNQUENTLlNMcEdldE1hY2hpbmVVR1VJRABTTHBHZXRNYWNoaW5lVUdVSUQAU1BQQ1MuU0xwR2V0VG9rZW5BY3RpdmF0aW9uR3JhbnRJbmZvAFNMcEdldFRva2VuQWN0aXZhdGlvbkdyYW50SW5mbwBTUFBDUy5T
THBJQUFjdGl2YXRlUHJvZHVjdABTTHBJQUFjdGl2YXRlUHJvZHVjdABTUFBDUy5TTHBJc0N1cnJlbnRJbnN0YWxsZWRQcm9kdWN0S2V5RGVmYXVsdEtleQBTTHBJc0N1cnJlbnRJbnN0YWxsZWRQcm9kdWN0S2V5RGVmYXVsdEtleQBTUFBDUy5TTHBQcm9jZXNzVk1Q
aXBlTWVzc2FnZQBTTHBQcm9jZXNzVk1QaXBlTWVzc2FnZQBTUFBDUy5TTHBTZXRBY3RpdmF0aW9uSW5Qcm9ncmVzcwBTTHBTZXRBY3RpdmF0aW9uSW5Qcm9ncmVzcwBTUFBDUy5TTHBUcmlnZ2VyU2VydmljZVdvcmtlcgBTTHBUcmlnZ2VyU2VydmljZVdvcmtlcgBT
UFBDUy5TTHBWTEFjdGl2YXRlUHJvZHVjdABTTHBWTEFjdGl2YXRlUHJvZHVjdAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFBgAAAAAAAAAAAAAOhgAABsYAAAXGAAAAAAAAAAAAAA
+GAAAHhgAABkYAAAAAAAAAAAAAAMYQAAgGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAiGAAAKpgAAAAAAAAyGAAAAAAAADUYAAAAAAAAIhgAACqYAAAAAAAAMhgAAAAAAAA1GAAAAAAAAACAFNMR2V0TGljZW5zaW5nU3RhdHVzSW5mb3JtYXRpb24AAQBTTEdldFByb2R1
Y3RTa3VJbmZvcm1hdGlvbgAA3QNMb2NhbEZyZWUARwFTdHJTdHJOSVcAAGAAAABgAABzcHBjcy5kbGwAAAAUYAAAS0VSTkVMMzIuZGxsAAAAAChgAABTSExXQVBJLmRsbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABABAAAAAYAACAAAAAAAAAAAAAAAAAAAABAAEAAAAwAACAAAAAAAAAAAAAAAAAAAABAAkEAABIAAAAWHAAABwDAAAAAAAAAAAAABwDNAAAAFYAUwBfAFYARQBSAFMASQBPAE4AXwBJAE4A
RgBPAAAAAAC9BO/+AAABAAMAAAAAAAAAAwAAAAAAAAAAAAAAAAAAAAQABAACAAAAAAAAAAAAAAAAAAAAfAIAAAEAUwB0AHIAaQBuAGcARgBpAGwAZQBJAG4AZgBvAAAAWAIAAAEAMAA0ADAAOQAwADQARQA0AAAAegAtAAEAQwBvAG0AcABhAG4AeQBOAGEAbQBlAAAA
AABBAG4AbwBtAGEAbABvAHUAcwAgAFMAbwBmAHQAdwBhAHIAZQAgAEQAZQB0AGUAcgBpAG8AcgBhAHQAaQBvAG4AIABDAG8AcgBwAG8AcgBhAHQAaQBvAG4AAAAAAD4ACwABAEYAaQBsAGUARABlAHMAYwByAGkAcAB0AGkAbwBuAAAAAABvAGgAbwBvAGsAIABTAFAA
UABDAAAAAAAwAAgAAQBGAGkAbABlAFYAZQByAHMAaQBvAG4AAAAAADAALgAzAC4AMAAuADAAAAAqAAUAAQBJAG4AdABlAHIAbgBhAGwATgBhAG0AZQAAAHMAcABwAGMAAAAAAIwANAABAEwAZQBnAGEAbABDAG8AcAB5AHIAaQBnAGgAdAAAAKkAIAAyADAAMgAzACAA
QQBuAG8AbQBhAGwAbwB1AHMAIABTAG8AZgB0AHcAYQByAGUAIABEAGUAdABlAHIAaQBvAHIAYQB0AGkAbwBuACAAQwBvAHIAcABvAHIAYQB0AGkAbwBuAAAAOgAJAAEATwByAGkAZwBpAG4AYQBsAEYAaQBsAGUAbgBhAG0AZQAAAHMAcABwAGMALgBkAGwAbAAAAAAA
LAAGAAEAUAByAG8AZAB1AGMAdABOAGEAbQBlAAAAAABvAGgAbwBvAGsAAAA0AAgAAQBQAHIAbwBkAHUAYwB0AFYAZQByAHMAaQBvAG4AAAAwAC4AMwAuADAALgAwAAAARAAAAAEAVgBhAHIARgBpAGwAZQBJAG4AZgBvAAAAAAAkAAQAAABUAHIAYQBuAHMAbABhAHQA
aQBvAG4AAAAAAAkE5AQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAQAAAUAAAAOzBQMHEwfjBSMVoxAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
:sppc32.dll:

:========================================================================================================================================

:sppc64.dll:
TVqQAAMAAAAEAAAA//8AALgAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAA4fug4AtAnNIbgBTM0hVGhpcyBwcm9ncmFtIGNhbm5vdCBiZSBydW4gaW4gRE9TIG1vZGUuDQ0KJAAAAAAAAABQRQAAZIYHAMDc0GQAAAAAAAAAAPAA
LiILAgIoAAIAAAAeAAAAAAAAABAAAAAQAAAAAJIxAgAAAAAQAAAAAgAABAAAAAAAAAAGAAAAAAAAAACQAAAABAAA39AAAAIAYAEAACAAAAAAAAAQAAAAAAAAAAAQAAAAAAAAEAAAAAAAAAAAAAAQAAAAAFAAAI0QAAAAcAAAUAEAAACAAAB4AwAAADAAACQAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAiHAAADgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAudGV4dAAAAHABAAAAEAAAAAIAAAAEAAAAAAAAAAAAAAAAAAAgAABgLnJkYXRhAAAgAAAAACAAAAAC
AAAABgAAAAAAAAAAAAAAAAAAQAAAQC5wZGF0YQAAJAAAAAAwAAAAAgAAAAgAAAAAAAAAAAAAAAAAAEAAAEAueGRhdGEAACQAAAAAQAAAAAIAAAAKAAAAAAAAAAAAAAAAAABAAABALmVkYXRhAACNEAAAAFAAAAASAAAADAAAAAAAAAAAAAAAAAAAQAAAQC5pZGF0YQAA
UAEAAABwAAAAAgAAAB4AAAAAAAAAAAAAAAAAAEAAAMAucnNyYwAAAHgDAAAAgAAAAAQAAAAgAAAAAAAAAAAAAAAAAABAAADAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALgBAAAAw0FUU0iD7EhFMclMjQXpDwAASI1E
JDjHRCQ0AAAAAEiJRCQoSI1EJDRIiUQkIEjHRCQ4AAAAAOj/AAAASItMJDhIix1TYAAAhcBBicR0B//TRTHk6yhEi0QkNEiNFaMPAAD/FUNgAABIi0wkOEiFwHQK/9NBvAEAAADrAv/TRIngSIPESFtBXMNBVUFUVVdWU0iD7Dgx9kyLrCSQAAAASIusJJgAAABMiWwk
IEiJz0iJbCQo6IoAAABBicSFwHVEQTl1AHY+SGveKEiLVQBIAdqDehAAdChIifnoIv///4XAdRxIA10ASMdDEAEAAABIx0MYAAAAAEjHQyAAAAAASP/G67xEieBIg8Q4W15fXUFcQV3DkJCQkJCQkP8lel8AAJCQDx+EAAAAAAD/JXpfAACQkA8fhAAAAAAA/yVKXwAA
kJD/JTpfAACQkP//////////AAAAAAAAAAD//////////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAATgBhAG0AZQAAAEcAcgBhAGMAZQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAABhAAAABAAAAGEAAAjhAAAARAAACOEAAAGREAABBAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAEAAAABBwMAB4IDMALAAAABDAcADGIIMAdgBnAFUATAAtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMDc0GQAAAAAxlIAAAEAAABDAAAAQwAAAChQAAA0UQAAQFIAAM9SAADvUgAABVMAAClTAABdUwAAoVMAAOlTAAAXVAAANVQAAGdU
AACdVAAA41QAAC1VAABhVQAAn1UAANNVAAANVgAAO1YAAHFWAACvVgAAz1YAAPtWAACOEAAAUVcAAG9XAACfVwAA01cAABFYAABNWAAAb1gAAKVYAADNWAAABVkAAEFZAABtWQAAp1kAALtZAAD7WQAAOVoAAE9aAAB1WgAAnVoAANNaAAAHWwAAPVsAAGlbAAClWwAA
41sAAA1cAAA5XAAAiVwAANFcAAARXQAAWV0AAKNdAADxXQAAG14AAEdeAACHXgAAu14AAOdeAAArXwAAW18AALVfAADrXwAAJ2AAAF1gAADiUgAA/VIAABpTAABGUwAAglMAAMhTAAADVAAAKVQAAFFUAACFVAAAw1QAAAtVAABKVQAAg1UAALxVAADzVQAAJ1YAAFlW
AACTVgAAwlYAAOhWAAAZVwAAMVcAAGNXAACKVwAAvFcAAPVXAAAyWAAAYVgAAI1YAAC8WAAA7FgAACZZAABaWQAAjVkAALRZAADeWQAAHVoAAEdaAABlWgAAjFoAALtaAADwWgAAJVsAAFZbAACKWwAAx1sAAPtbAAAmXAAAZFwAALBcAAD0XAAAOF0AAIFdAADNXQAA
CV4AADReAABqXgAApF4AANReAAAMXwAARl8AAItfAADTXwAADGAAAEVgAAB4YAAAAAABAAIAAwAEAAUABgAHAAgACQAKAAsADAANAA4ADwAQABEAEgATABQAFQAWABcAGAAZABoAGwAcAB0AHgAfACAAIQAiACMAJAAlACYAJwAoACkAKgArACwALQAuAC8AMAAxADIA
MwA0ADUANgA3ADgAOQA6ADsAPAA9AD4APwBAAEEAQgBzcHBjLmRsbABTUFBDUy5TTENhbGxTZXJ2ZXIAU0xDYWxsU2VydmVyAFNQUENTLlNMQ2xvc2UAU0xDbG9zZQBTUFBDUy5TTENvbnN1bWVSaWdodABTTENvbnN1bWVSaWdodABTUFBDUy5TTERlcG9zaXRNaWdy
YXRpb25CbG9iAFNMRGVwb3NpdE1pZ3JhdGlvbkJsb2IAU1BQQ1MuU0xEZXBvc2l0T2ZmbGluZUNvbmZpcm1hdGlvbklkAFNMRGVwb3NpdE9mZmxpbmVDb25maXJtYXRpb25JZABTUFBDUy5TTERlcG9zaXRPZmZsaW5lQ29uZmlybWF0aW9uSWRFeABTTERlcG9zaXRP
ZmZsaW5lQ29uZmlybWF0aW9uSWRFeABTUFBDUy5TTERlcG9zaXRTdG9yZVRva2VuAFNMRGVwb3NpdFN0b3JlVG9rZW4AU1BQQ1MuU0xGaXJlRXZlbnQAU0xGaXJlRXZlbnQAU1BQQ1MuU0xHYXRoZXJNaWdyYXRpb25CbG9iAFNMR2F0aGVyTWlncmF0aW9uQmxvYgBT
UFBDUy5TTEdhdGhlck1pZ3JhdGlvbkJsb2JFeABTTEdhdGhlck1pZ3JhdGlvbkJsb2JFeABTUFBDUy5TTEdlbmVyYXRlT2ZmbGluZUluc3RhbGxhdGlvbklkAFNMR2VuZXJhdGVPZmZsaW5lSW5zdGFsbGF0aW9uSWQAU1BQQ1MuU0xHZW5lcmF0ZU9mZmxpbmVJbnN0
YWxsYXRpb25JZEV4AFNMR2VuZXJhdGVPZmZsaW5lSW5zdGFsbGF0aW9uSWRFeABTUFBDUy5TTEdldEFjdGl2ZUxpY2Vuc2VJbmZvAFNMR2V0QWN0aXZlTGljZW5zZUluZm8AU1BQQ1MuU0xHZXRBcHBsaWNhdGlvbkluZm9ybWF0aW9uAFNMR2V0QXBwbGljYXRpb25J
bmZvcm1hdGlvbgBTUFBDUy5TTEdldEFwcGxpY2F0aW9uUG9saWN5AFNMR2V0QXBwbGljYXRpb25Qb2xpY3kAU1BQQ1MuU0xHZXRBdXRoZW50aWNhdGlvblJlc3VsdABTTEdldEF1dGhlbnRpY2F0aW9uUmVzdWx0AFNQUENTLlNMR2V0RW5jcnlwdGVkUElERXgAU0xH
ZXRFbmNyeXB0ZWRQSURFeABTUFBDUy5TTEdldEdlbnVpbmVJbmZvcm1hdGlvbgBTTEdldEdlbnVpbmVJbmZvcm1hdGlvbgBTUFBDUy5TTEdldEluc3RhbGxlZFByb2R1Y3RLZXlJZHMAU0xHZXRJbnN0YWxsZWRQcm9kdWN0S2V5SWRzAFNQUENTLlNMR2V0TGljZW5z
ZQBTTEdldExpY2Vuc2UAU1BQQ1MuU0xHZXRMaWNlbnNlRmlsZUlkAFNMR2V0TGljZW5zZUZpbGVJZABTUFBDUy5TTEdldExpY2Vuc2VJbmZvcm1hdGlvbgBTTEdldExpY2Vuc2VJbmZvcm1hdGlvbgBTTEdldExpY2Vuc2luZ1N0YXR1c0luZm9ybWF0aW9uAFNQUENT
LlNMR2V0UEtleUlkAFNMR2V0UEtleUlkAFNQUENTLlNMR2V0UEtleUluZm9ybWF0aW9uAFNMR2V0UEtleUluZm9ybWF0aW9uAFNQUENTLlNMR2V0UG9saWN5SW5mb3JtYXRpb24AU0xHZXRQb2xpY3lJbmZvcm1hdGlvbgBTUFBDUy5TTEdldFBvbGljeUluZm9ybWF0
aW9uRFdPUkQAU0xHZXRQb2xpY3lJbmZvcm1hdGlvbkRXT1JEAFNQUENTLlNMR2V0UHJvZHVjdFNrdUluZm9ybWF0aW9uAFNMR2V0UHJvZHVjdFNrdUluZm9ybWF0aW9uAFNQUENTLlNMR2V0U0xJRExpc3QAU0xHZXRTTElETGlzdABTUFBDUy5TTEdldFNlcnZpY2VJ
bmZvcm1hdGlvbgBTTEdldFNlcnZpY2VJbmZvcm1hdGlvbgBTUFBDUy5TTEluc3RhbGxMaWNlbnNlAFNMSW5zdGFsbExpY2Vuc2UAU1BQQ1MuU0xJbnN0YWxsUHJvb2ZPZlB1cmNoYXNlAFNMSW5zdGFsbFByb29mT2ZQdXJjaGFzZQBTUFBDUy5TTEluc3RhbGxQcm9v
Zk9mUHVyY2hhc2VFeABTTEluc3RhbGxQcm9vZk9mUHVyY2hhc2VFeABTUFBDUy5TTElzR2VudWluZUxvY2FsRXgAU0xJc0dlbnVpbmVMb2NhbEV4AFNQUENTLlNMTG9hZEFwcGxpY2F0aW9uUG9saWNpZXMAU0xMb2FkQXBwbGljYXRpb25Qb2xpY2llcwBTUFBDUy5T
TE9wZW4AU0xPcGVuAFNQUENTLlNMUGVyc2lzdEFwcGxpY2F0aW9uUG9saWNpZXMAU0xQZXJzaXN0QXBwbGljYXRpb25Qb2xpY2llcwBTUFBDUy5TTFBlcnNpc3RSVFNQYXlsb2FkT3ZlcnJpZGUAU0xQZXJzaXN0UlRTUGF5bG9hZE92ZXJyaWRlAFNQUENTLlNMUmVB
cm0AU0xSZUFybQBTUFBDUy5TTFJlZ2lzdGVyRXZlbnQAU0xSZWdpc3RlckV2ZW50AFNQUENTLlNMUmVnaXN0ZXJQbHVnaW4AU0xSZWdpc3RlclBsdWdpbgBTUFBDUy5TTFNldEF1dGhlbnRpY2F0aW9uRGF0YQBTTFNldEF1dGhlbnRpY2F0aW9uRGF0YQBTUFBDUy5T
TFNldEN1cnJlbnRQcm9kdWN0S2V5AFNMU2V0Q3VycmVudFByb2R1Y3RLZXkAU1BQQ1MuU0xTZXRHZW51aW5lSW5mb3JtYXRpb24AU0xTZXRHZW51aW5lSW5mb3JtYXRpb24AU1BQQ1MuU0xVbmluc3RhbGxMaWNlbnNlAFNMVW5pbnN0YWxsTGljZW5zZQBTUFBDUy5T
TFVuaW5zdGFsbFByb29mT2ZQdXJjaGFzZQBTTFVuaW5zdGFsbFByb29mT2ZQdXJjaGFzZQBTUFBDUy5TTFVubG9hZEFwcGxpY2F0aW9uUG9saWNpZXMAU0xVbmxvYWRBcHBsaWNhdGlvblBvbGljaWVzAFNQUENTLlNMVW5yZWdpc3RlckV2ZW50AFNMVW5yZWdpc3Rl
ckV2ZW50AFNQUENTLlNMVW5yZWdpc3RlclBsdWdpbgBTTFVucmVnaXN0ZXJQbHVnaW4AU1BQQ1MuU0xwQXV0aGVudGljYXRlR2VudWluZVRpY2tldFJlc3BvbnNlAFNMcEF1dGhlbnRpY2F0ZUdlbnVpbmVUaWNrZXRSZXNwb25zZQBTUFBDUy5TTHBCZWdpbkdlbnVp
bmVUaWNrZXRUcmFuc2FjdGlvbgBTTHBCZWdpbkdlbnVpbmVUaWNrZXRUcmFuc2FjdGlvbgBTUFBDUy5TTHBDbGVhckFjdGl2YXRpb25JblByb2dyZXNzAFNMcENsZWFyQWN0aXZhdGlvbkluUHJvZ3Jlc3MAU1BQQ1MuU0xwRGVwb3NpdERvd25sZXZlbEdlbnVpbmVU
aWNrZXQAU0xwRGVwb3NpdERvd25sZXZlbEdlbnVpbmVUaWNrZXQAU1BQQ1MuU0xwRGVwb3NpdFRva2VuQWN0aXZhdGlvblJlc3BvbnNlAFNMcERlcG9zaXRUb2tlbkFjdGl2YXRpb25SZXNwb25zZQBTUFBDUy5TTHBHZW5lcmF0ZVRva2VuQWN0aXZhdGlvbkNoYWxs
ZW5nZQBTTHBHZW5lcmF0ZVRva2VuQWN0aXZhdGlvbkNoYWxsZW5nZQBTUFBDUy5TTHBHZXRHZW51aW5lQmxvYgBTTHBHZXRHZW51aW5lQmxvYgBTUFBDUy5TTHBHZXRHZW51aW5lTG9jYWwAU0xwR2V0R2VudWluZUxvY2FsAFNQUENTLlNMcEdldExpY2Vuc2VBY3F1
aXNpdGlvbkluZm8AU0xwR2V0TGljZW5zZUFjcXVpc2l0aW9uSW5mbwBTUFBDUy5TTHBHZXRNU1BpZEluZm9ybWF0aW9uAFNMcEdldE1TUGlkSW5mb3JtYXRpb24AU1BQQ1MuU0xwR2V0TWFjaGluZVVHVUlEAFNMcEdldE1hY2hpbmVVR1VJRABTUFBDUy5TTHBHZXRU
b2tlbkFjdGl2YXRpb25HcmFudEluZm8AU0xwR2V0VG9rZW5BY3RpdmF0aW9uR3JhbnRJbmZvAFNQUENTLlNMcElBQWN0aXZhdGVQcm9kdWN0AFNMcElBQWN0aXZhdGVQcm9kdWN0AFNQUENTLlNMcElzQ3VycmVudEluc3RhbGxlZFByb2R1Y3RLZXlEZWZhdWx0S2V5
AFNMcElzQ3VycmVudEluc3RhbGxlZFByb2R1Y3RLZXlEZWZhdWx0S2V5AFNQUENTLlNMcFByb2Nlc3NWTVBpcGVNZXNzYWdlAFNMcFByb2Nlc3NWTVBpcGVNZXNzYWdlAFNQUENTLlNMcFNldEFjdGl2YXRpb25JblByb2dyZXNzAFNMcFNldEFjdGl2YXRpb25JblBy
b2dyZXNzAFNQUENTLlNMcFRyaWdnZXJTZXJ2aWNlV29ya2VyAFNMcFRyaWdnZXJTZXJ2aWNlV29ya2VyAFNQUENTLlNMcFZMQWN0aXZhdGVQcm9kdWN0AFNMcFZMQWN0aXZhdGVQcm9kdWN0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAUHAAAAAAAAAAAAAAIHEAAIhwAABocAAAAAAAAAAAAAAwcQAAoHAAAHhwAAAAAAAAAAAAAERxAACwcAAAAAAAAAAAAAAAAAAAAAAAAAAAAADAcAAAAAAAAOJwAAAAAAAAAAAAAAAAAAAAcQAAAAAAAAAAAAAAAAAA
DHEAAAAAAAAAAAAAAAAAAMBwAAAAAAAA4nAAAAAAAAAAAAAAAAAAAABxAAAAAAAAAAAAAAAAAAAMcQAAAAAAAAAAAAAAAAAAAgBTTEdldExpY2Vuc2luZ1N0YXR1c0luZm9ybWF0aW9uAAEAU0xHZXRQcm9kdWN0U2t1SW5mb3JtYXRpb24AAOgDTG9jYWxGcmVlAFEB
U3RyU3RyTklXAABwAAAAcAAAc3BwY3MuZGxsAAAAFHAAAEtFUk5FTDMyLmRsbAAAAAAocAAAU0hMV0FQSS5kbGwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAEAAAABgAAIAAAAAAAAAAAAAAAAAAAAEAAQAAADAAAIAAAAAAAAAAAAAA
AAAAAAEACQQAAEgAAABYgAAAHAMAAAAAAAAAAAAAHAM0AAAAVgBTAF8AVgBFAFIAUwBJAE8ATgBfAEkATgBGAE8AAAAAAL0E7/4AAAEAAwAAAAAAAAADAAAAAAAAAAAAAAAAAAAABAAEAAIAAAAAAAAAAAAAAAAAAAB8AgAAAQBTAHQAcgBpAG4AZwBGAGkAbABlAEkA
bgBmAG8AAABYAgAAAQAwADQAMAA5ADAANABFADQAAAB6AC0AAQBDAG8AbQBwAGEAbgB5AE4AYQBtAGUAAAAAAEEAbgBvAG0AYQBsAG8AdQBzACAAUwBvAGYAdAB3AGEAcgBlACAARABlAHQAZQByAGkAbwByAGEAdABpAG8AbgAgAEMAbwByAHAAbwByAGEAdABpAG8A
bgAAAAAAPgALAAEARgBpAGwAZQBEAGUAcwBjAHIAaQBwAHQAaQBvAG4AAAAAAG8AaABvAG8AawAgAFMAUABQAEMAAAAAADAACAABAEYAaQBsAGUAVgBlAHIAcwBpAG8AbgAAAAAAMAAuADMALgAwAC4AMAAAACoABQABAEkAbgB0AGUAcgBuAGEAbABOAGEAbQBlAAAA
cwBwAHAAYwAAAAAAjAA0AAEATABlAGcAYQBsAEMAbwBwAHkAcgBpAGcAaAB0AAAAqQAgADIAMAAyADMAIABBAG4AbwBtAGEAbABvAHUAcwAgAFMAbwBmAHQAdwBhAHIAZQAgAEQAZQB0AGUAcgBpAG8AcgBhAHQAaQBvAG4AIABDAG8AcgBwAG8AcgBhAHQAaQBvAG4A
AAA6AAkAAQBPAHIAaQBnAGkAbgBhAGwARgBpAGwAZQBuAGEAbQBlAAAAcwBwAHAAYwAuAGQAbABsAAAAAAAsAAYAAQBQAHIAbwBkAHUAYwB0AE4AYQBtAGUAAAAAAG8AaABvAG8AawAAADQACAABAFAAcgBvAGQAdQBjAHQAVgBlAHIAcwBpAG8AbgAAADAALgAzAC4A
MAAuADAAAABEAAAAAQBWAGEAcgBGAGkAbABlAEkAbgBmAG8AAAAAACQABAAAAFQAcgBhAG4AcwBsAGEAdABpAG8AbgAAAAAACQTkBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
:sppc64.dll:

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:KMS38Activation
@setlocal DisableDelayedExpansion
@echo off

::  若要激活，请使用 /KMS38 参数运行脚本或将以下行参数由 0 更改为 1
set _act=0

::  若要移除 KMS38 保护，请使用 /KMS38-RemoveProtection 参数运行脚本或将以下行参数由 0 更改为 1
set _rem=0

::  若要在当前版本不支持 KMS38 激活时禁用更改版本，请将参数的值由 0 更改为 1，或使用“/KMS38-NoEditionChange”参数运行脚本
set _NoEditionChange=0

::  如果在上面几行中更改了值或使用参数，脚本将会在无人值守模式下运行

::========================================================================================================================================

cls
color 07
title KMS38 激活 %masver%

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="/KMS38"                  set _act=1
if /i "%%A"=="/KMS38-RemoveProtection" set _rem=1
if /i "%%A"=="/KMS38-NoEditionChange"  set _NoEditionChange=1
if /i "%%A"=="-el"                     set _elev=1
)
)

for %%A in (%_act% %_rem% %_NoEditionChange%) do (if "%%A"=="1" set _unattended=1)

::========================================================================================================================================

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

if %_NCS% EQU 1 (
for /F %%a in ('echo prompt $E ^| cmd') do set "esc=%%a"
set     "Red="41;97m""
set    "Gray="100;97m""
set   "Green="42;97m""
set    "Blue="44;97m""
set  "_White="40;37m""
set  "_Green="40;92m""
set "_Yellow="40;93m""
) else (
set     "Red="Red" "white""
set    "Gray="Darkgray" "white""
set   "Green="DarkGreen" "white""
set    "Blue="Blue" "white""
set  "_White="Black" "Gray""
set  "_Green="Black" "Green""
set "_Yellow="Black" "Yellow""
)

set _k38=
set "nceline=echo: &echo ==== 错误 ==== &echo:"
set "eline=echo: &call :dk_color %Red% "==== 错误 ====" &echo:"
if %~z0 GEQ 200000 (
set "_exitmsg=返回"
set "_fixmsg=请返回主菜单，选择疑难解答并运行修复许可选项。"
) else (
set "_exitmsg=退出"
set "_fixmsg=在 MAS 文件夹中，请运行疑难解答脚本并选择修复许可选项。"
)

set "specific_kms=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f"

::========================================================================================================================================

if %winbuild% LSS 14393 (
%eline%
echo 检测到不受支持的操作系统版本 [%winbuild%]。
echo KMS38 激活支持 Windows 10/11/Server，Build 14393 及以上版本。
goto dk_done
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

if %_rem%==1 goto :k_uninstall

:k_menu

if %_unattended%==0 (
cls
mode 76, 25
title KMS38 激活 %masver%

echo:
echo:
echo:
echo:
echo         ____________________________________________________________
echo:
echo                 [1] KMS38 激活
echo                 ____________________________________________
echo:
echo                 [2] 移除 KM38 保护
echo:
echo                 [0] %_exitmsg%
echo         ____________________________________________________________
echo: 
call :dk_color2 %_White% "              " %_Green% "请输入一个菜单选项 [1,2,0]"
choice /C:120 /N
set _el=!errorlevel!
if !_el!==3  exit /b
if !_el!==2  goto :k_uninstall
if !_el!==1  goto :k_menu2
goto :k_menu
)

::========================================================================================================================================

:k_menu2

cls
mode 110, 34
if exist "%Systemdrive%\Windows\System32\spp\store_test\" mode 134, 34
title KMS38 激活 %masver%

echo:
echo 正在初始化……

::  检查 PowerShell

%psc% $ExecutionContext.SessionState.LanguageMode %nul2% | find /i "Full" %nul1% || (
%eline%
%psc% $ExecutionContext.SessionState.LanguageMode
echo:
echo PowerShell 不可用，正在中止……
echo 如果你对Powershell施加了限制，请撤销这些更改。
echo:
echo 请查看此页面以获得帮助。 %mas%troubleshoot
goto dk_done
)

::========================================================================================================================================

call :dk_product
call :dk_ckeckwmic

::  显示潜在的脚本卡住情况的信息

sc start sppsvc %nul%
if %errorlevel% NEQ 1056 if %errorlevel% NEQ 0 (
echo:
echo 错误代码：%errorlevel%
call :dk_color %Red% "启动 [sppsvc] 服务失败，其余的进程可能需要很长时间……"
echo:
)

::========================================================================================================================================

::  检查系统是否已永久激活

call :dk_checkperm
if defined _perm (
cls
echo ___________________________________________________________________________________________
echo:
call :dk_color2 %_White% "     " %Green% "正在检查：%winos% 已永久激活。"
call :dk_color2 %_White% "     " %Gray% "不需要执行激活。"
echo ___________________________________________________________________________________________
if %_unattended%==1 goto dk_done
echo:
choice /C:10 /N /M ">    [1] 激活 [0] %_exitmsg% ："
if errorlevel 2 exit /b
)
cls

::========================================================================================================================================

::  检查评估版本

set _eval=
set _evalserv=

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" set _eval=1
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalEdition~*.mum" set _evalserv=1
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalCorEdition~*.mum" set _eval=1 & set _evalserv=1

if defined _eval (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul2% | find /i "Eval" %nul1% && (
%eline%
echo [%winos% ^| %winbuild%]
if defined _evalserv (
echo Server 评估版本无法激活。请将其转换为完整的 Server 操作系统。
echo:
echo 在 MAS 中，请转到附加功能并使用“更改版本”选项。
) else (
echo 评估版本无法激活。 
echo 你需要安装 %winos% 的完整版本。
echo:
echo 请从此处下载，
echo %mas%genuine-installation-media.html
)
goto dk_done
)
)

::========================================================================================================================================

::  检查 clipup.exe 以检测和激活服务器 cor 和 acor 版本

set a_cor=
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*CorEdition~*.mum" if not exist "%systemroot%\System32\clipup.exe" set a_cor=1

if defined a_cor (
if not exist "!_work!\clipup.exe" (
%eline%
echo clipup.exe 在 Server Cor/Acor [无 GUI] 版本中不存在。
echo 它是 KMS38 激活必要组件。
echo 查看下方的页面了解如何激活它。
echo %mas%kms38.html
goto dk_done
)
)

::========================================================================================================================================

call :dk_checksku

if not defined osSKU (
%eline%
echo 未正确检测到 SKU 值。正在中止……
goto dk_done
)

::========================================================================================================================================

set error=

cls
echo:
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PROCESSOR_ARCHITECTURE') do set arch=%%b
for /f "tokens=6-7 delims=[]. " %%i in ('ver') do if "%%j"=="" (set fullbuild=%%i) else (set fullbuild=%%i.%%j)
echo 正在检查操作系统信息                    [%winos% ^| %fullbuild% ^| %arch%]

::========================================================================================================================================

::  检查 Windows Script Host

set _WSH=1
reg query "HKCU\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled %nul2% | find /i "0x0" %nul1% && (set _WSH=0)
reg query "HKLM\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled %nul2% | find /i "0x0" %nul1% && (set _WSH=0)

if %_WSH% EQU 0 (
reg add "HKLM\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f %nul%
reg add "HKCU\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f %nul%
if not "%arch%"=="x86" reg add "HKLM\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f /reg:32 %nul%
echo 正在启用 Windows Script Host            [成功]
)

::========================================================================================================================================

echo 正在初始化诊断测试……

set "_serv=ClipSVC sppsvc KeyIso Winmgmt"

::  Client License Service (ClipSVC)
::  Software Protection
::  CNG Key Isolation
::  Windows Management Instrumentation

call :dk_errorcheck

::========================================================================================================================================

::  检查 SKU 值 / 在多个地方检查，查找损坏的版本更改

set _gvlk=
call :dk_channel
if /i "Volume:GVLK"=="%_channel%" set _gvlk=1

::  检查密钥

set key=
set pkey=
set altkey=
set skufound=
set changekey=
set altedition=

call :kms38data getkey
if not defined key call :dk_gvlk %nul%
if defined applist if not defined key call :kms38fallback

if defined altkey (set key=%altkey%&set changekey=1)

set /a UBR=0
if %osSKU%==191 if defined altkey if defined altedition (
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v UBR 2^>nul') do if not errorlevel 1 set /a UBR=%%b
if %winbuild% GEQ 19044 if !UBR! LSS 2788 (
call :dk_color %Blue% "对于 IotEnterpriseS KMS38 激活，Windows 必须更新到 Build 19044.2788 或更高版本。"
)
)

if not defined key if defined notfoundaltactID (
call :dk_color %Red% "正在检查 KMS38 的备用版本               [未找到 %altedition% 的激活 ID]"
)

if not defined key if not defined _gvlk (
%eline%
echo [%winos% ^| %winbuild% ^| SKU:%osSKU%]
if not defined skufound (
echo 在支持的产品列表中找不到此产品。
) else (
echo 未安装所需的License文件。
)
echo 请确保你使用的是此脚本的更新版本。
echo %mas%
echo:
goto dk_done
)

::========================================================================================================================================

::  Install key

echo:
if defined changekey (
call :dk_color %Blue% "[%altedition%] 版本产品密钥将用于启用 KMS38 激活。"
echo:
)

set _partial=
if not defined key (
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%# in ('wmic path SoftwareLicensingProduct where "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey<>null" Get PartialProductKey /value %nul6%') do set "_partial=%%#"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%# in ('%psc% "(([WMISEARCHER]'SELECT PartialProductKey FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND PartialProductKey IS NOT NULL').Get()).PartialProductKey | %% {echo ('PartialProductKey='+$_)}" %nul6%') do set "_partial=%%#"
call echo 正在检查已安装的产品密钥                [部分产品密钥 - %%_partial%%] [Volume:GVLK]
)

set error_code=
if defined key (
if %_wmic% EQU 1 wmic path SoftwareLicensingService where __CLASS='SoftwareLicensingService' call InstallProductKey ProductKey="%key%" %nul%
if %_wmic% EQU 0 %psc% "(([WMISEARCHER]'SELECT Version FROM SoftwareLicensingService').Get()).InstallProductKey('%key%')" %nul%
if not !errorlevel!==0 cscript //nologo %windir%\system32\slmgr.vbs /ipk %key% %nul%
set error_code=!errorlevel!
cmd /c exit /b !error_code!
if !error_code! NEQ 0 set "error_code=[0x!=ExitCode!]"

if !error_code! EQU 0 (
call :dk_refresh
echo 正在安装 KMS 客户端安装程序密钥         [%key%] [成功]
) else (
call :dk_color %Red% "正在安装 KMS 客户端安装程序密钥         [%key%] [失败] !error_code!"
if not defined error (
call :dk_color %Blue% "%_fixmsg%"
set showfix=1
)
set error=1
)
)

::========================================================================================================================================

::  检查激活 ID 用于设置特定 KMS 主机

set app=
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get ID /VALUE" %nul6%') do call set "app=%%a"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND Description like ''%%KMSCLIENT%%'' AND PartialProductKey IS NOT NULL').Get()).ID | %% {echo ('ID='+$_)}" %nul6%') do call set "app=%%a"

if not defined app (
call :dk_color %Red% "正在检查已安装的 GVLK 激活 ID           [未找到] 正在中止……"
call :dk_color2 %Blue% "查看此页面以获取帮助" %_Yellow% " %mas%troubleshoot"
goto :dk_done
)

::========================================================================================================================================

::  将特定 KMS 主机设置为本地主机
::  通过这样做，全球 KMS IP 无法代替 KMS38 激活，但可以与 Office 和其他 Windows 版本一起使用

echo:
%nul% reg delete "HKLM\%specific_kms%" /f
%nul% reg delete "HKU\S-1-5-20\%specific_kms%" /f

%nul% reg query "HKLM\%specific_kms%" && (
%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':regdel\:.*';iex ($f[1]);"
%nul% reg delete "HKLM\%specific_kms%" /f
)

set k_error=
%nul% reg add "HKLM\%specific_kms%\%app%" /f /v KeyManagementServiceName /t REG_SZ /d "127.0.0.2" || set k_error=1
%nul% reg add "HKLM\%specific_kms%\%app%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" || set k_error=1

if not defined k_error (
echo 正在添加特定 KMS 主机                   [LocalHost 127.0.0.2] [成功]
) else (
call :dk_color %Red% "正在添加特定 KMS 主机                   [LocalHost 127.0.0.2] [失败]"
)

::========================================================================================================================================

::  将 clipup.exe 复制到 System32 目录以激活 Server Cor/Acor 版本

if defined a_cor (
set "_clipup=%systemroot%\System32\clipup.exe"
pushd "!_work!\"
copy /y /b "ClipUp.exe" "!_clipup!" %nul%
popd

echo:
if exist "!_clipup!" (
echo 正在复制 clipup.exe 文件到              [%systemroot%\System32\] [成功]
) else (
call :dk_color %Red% "正在复制 clipup.exe 文件到              [%systemroot%\System32\] [失败] 正在中止……"
goto :k_final
)
)

::========================================================================================================================================

::  生成 GenuineTicket.xml 并应用
::  应用票证的最正确方法是重新启动 ClipSVC 服务，但我们无法以此方式检查日志详细信息
::  为了获取日志详细信息并正确应用票证，脚本将安装票证 2 次（重新启动服务 + clipup -v -o）

if not exist %SystemRoot%\system32\ClipUp.exe (
call :dk_color %Red% "正在检查 ClipUp.exe 文件                [未找到，中止进程]"
call :dk_color2 %Blue% "查看此页面以获取帮助" %_Yellow% " %mas%troubleshoot"
goto :k_final
)

set "tdir=%ProgramData%\Microsoft\Windows\ClipSVC\GenuineTicket"
if not exist "%tdir%\" md "%tdir%\" %nul%

if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%
if exist "%tdir%\*.xml" del /f /q "%tdir%\*.xml" %nul%
if exist "%ProgramData%\Microsoft\Windows\ClipSVC\Install\Migration\*" del /f /q "%ProgramData%\Microsoft\Windows\ClipSVC\Install\Migration\*" %nul%

::  签名值按原样，未经过编码
::  会话 ID 采用 Base64 编码格式。它的解码值是“OSMajorVersion=5;OSMinorVersion=1;OSPlatformId=2;PP=0;GVLKExp=2038-01-19T03:14:07Z;DownlevelGenuineState=1;”
::  请查看 mass grave[.]dev/kms38.html#Manual_Activation 了解它如何生成

set "signature=C52iGEoH+1VqzI6kEAqOhUyrWuEObnivzaVjyef8WqItVYd/xGDTZZ3bkxAI9hTpobPFNJyJx6a3uriXq3HVd7mlXfSUK9ydeoUdG4eqMeLwkxeb6jQWJzLOz41rFVSMtBL0e+ycCATebTaXS4uvFYaDHDdPw2lKY8ADj3MLgsA="
set "sessionId=TwBTAE0AYQBqAG8AcgBWAGUAcgBzAGkAbwBuAD0ANQA7AE8AUwBNAGkAbgBvAHIAVgBlAHIAcwBpAG8AbgA9ADEAOwBPAFMAUABsAGEAdABmAG8AcgBtAEkAZAA9ADIAOwBQAFAAPQAwADsARwBWAEwASwBFAHgAcAA9ADIAMAAzADgALQAwADEALQAxADkAVAAwADMAOgAxADQAOgAwADcAWgA7AEQAbwB3AG4AbABlAHYAZQBsAEcAZQBuAHUAaQBuAGUAUwB0AGEAdABlAD0AMQA7AAAA"
<nul set /p "=<?xml version="1.0" encoding="utf-8"?><genuineAuthorization xmlns="http://www.microsoft.com/DRM/SL/GenuineAuthorization/1.0"><version>1.0</version><genuineProperties origin="sppclient"><properties>OA3xOriginalProductId=;OA3xOriginalProductKey=;SessionId=%sessionId%;TimeStampClient=2022-10-11T12:00:00Z</properties><signatures><signature name="clientLockboxKey" method="rsa-sha256">%signature%</signature></signatures></genuineProperties></genuineAuthorization>" >"%tdir%\GenuineTicket"

copy /y /b "%tdir%\GenuineTicket" "%tdir%\GenuineTicket.xml" %nul%

if not exist "%tdir%\GenuineTicket.xml" (
call :dk_color %Red% "正在生成 GenuineTicket.xml              [失败，中止进程]"
if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%
goto :k_final
) else (
echo 正在生成 GenuineTicket.xml              [成功]
)

set "_xmlexist=if exist "%tdir%\GenuineTicket.xml""

::  停止 sppsvc

%psc% Stop-Service sppsvc %nul%

sc query sppsvc | find /i "STOPPED" %nul% && (
echo 正在停止 sppsvc 服务                    [成功]
) || (
call :dk_color %Gray% "正在停止 sppsvc 服务                    [失败]"
)

%_xmlexist% (
%psc% Restart-Service ClipSVC %nul%
%_xmlexist% timeout /t 2 %nul%
%_xmlexist% timeout /t 2 %nul%

%_xmlexist% (
set error=1
if exist "%tdir%\*.xml" del /f /q "%tdir%\*.xml" %nul%
call :dk_color %Red% "正在安装 GenuineTicket.xml              [重新启动 ClipSVC 服务失败，正在等待……]"
)
)

copy /y /b "%tdir%\GenuineTicket" "%tdir%\GenuineTicket.xml" %nul%
clipup -v -o

set rebuildinfo=

if not exist %ProgramData%\Microsoft\Windows\ClipSVC\tokens.dat (
set error=1
set rebuildinfo=1
call :dk_color %Red% "正在检查 ClipSVC tokens.dat             [未找到]"
)

%_xmlexist% (
set error=1
set rebuildinfo=1
call :dk_color %Red% "正在安装 GenuineTicket.xml              [使用 clipup -v -o 失败]"
)

if exist "%ProgramData%\Microsoft\Windows\ClipSVC\Install\Migration\*.xml" (
set error=1
set rebuildinfo=1
call :dk_color %Red% "检查票证迁移                            [失败]"
)

if defined applist if not defined showfix if defined rebuildinfo (
set showfix=1
call :dk_color %Blue% "%_fixmsg%"
)

if exist "%tdir%\Genuine*" del /f /q "%tdir%\Genuine*" %nul%

::==========================================================================================================================================

call :dk_product

echo:
echo 正在激活……
echo:

call :k_checkexp
if defined _k38 (
call :k_actinfo
goto :k_final
)

::  清除 180 天 KMS 激活锁，具有特定于 Windows SKU 的重置，无需重新启动系统

if %_wmic% EQU 1 wmic path SoftwareLicensingProduct where ID='%app%' call ReArmsku %nul%
if %_wmic% EQU 0 %psc% "$null=([WMI]'SoftwareLicensingProduct=''%app%''').ReArmsku()" %nul%

if %errorlevel%==0 (
echo 正在应用 SKU-ID Rearm                   [成功]
) else (
call :dk_color %Red% "正在应用 SKU-ID Rearm                   [失败]"
)
call :dk_refresh

echo:
call :k_checkexp
if defined _k38 (
call :k_actinfo
goto :k_final
)

call :dk_color %Red% "激活失败"
if not defined error call :dk_color %Blue% "%_fixmsg%"
call :dk_color2 %Blue% "查看此页面以获取帮助" %_Yellow% " %mas%troubleshoot"

::========================================================================================================================================

:k_final

::  如果未完成激活，移除所添加的特定 KMS 主机（本地主机）

echo:
if not defined _k38 (
%nul% reg delete "HKLM\%specific_kms%" /f
%nul% reg delete "HKU\S-1-5-20\%specific_kms%" /f
%nul% reg query "HKLM\%specific_kms%" && (
call :dk_color %Red% "正在移除所添加的特定 KMS 主机           [失败]"
) || (
echo 正在移除所添加的特定 KMS 主机           [成功]
)
)

::  保护 KMS38（如果用户选择且条件正确）

if defined _k38 (
%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':regdel\:.*';& ([ScriptBlock]::Create($f[1])) -protect;"
%nul% reg delete "HKLM\%specific_kms%" /f
%nul% reg query "HKLM\%specific_kms%" && (
echo 通过 KMS 保护 KMS38                     [成功] [已锁定注册表项]
) || (
call :dk_color %Red% "通过 KMS 保护 KMS38                     [锁定注册表项失败]"
)
)

::  默认情况下 clipup.exe 在 server cor 和 acor 版本中不存在，已使用此脚本复制到相应位置

if defined a_cor if exist "%_clipup%" del /f /q "%_clipup%" %nul%

if defined a_cor (
if exist "%_clipup%" (
call :dk_color %Red% "正在删除已复制的 clipup.exe 文件        [失败]"
) else (
echo 正在删除已复制的 clipup.exe 文件        [成功]
)
)

for %%# in (175 407) do if %osSKU%==%%# (
call :dk_color %Red% "%winos% 版本不支持在非 Azure 平台上激活。"
)

goto :dk_done

::========================================================================================================================================

:k_uninstall

cls
mode 99, 28
title 移除 KMS38 保护 %masver%

%nul% reg delete "HKLM\%specific_kms%" /f
%nul% reg delete "HKU\S-1-5-20\%specific_kms%" /f

%nul% reg query "HKLM\%specific_kms%" && (
%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':regdel\:.*';iex ($f[1]);"
%nul% reg delete "HKLM\%specific_kms%" /f
)

echo:
%nul% reg query "HKLM\%specific_kms%" && (
call :dk_color %Red% "正在移除特定 KMS 主机                   [失败]"
) || (
echo 正在移除特定 KMS 主机                   [成功]
)

goto :dk_done

::========================================================================================================================================

::  运行此代码以保护/撤消以下注册表项以进行 KMS38 保护
::  HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f

::  KMS38 保护将会停止 180 天 KMS 激活从而取代 KMS38 激活

:regdel:
param (
    [switch]$protect
)

$SID = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')
$Admin = ($SID.Translate([System.Security.Principal.NTAccount])).Value

if($protect) {
$ruleArgs = @("$Admin", "Delete, SetValue", "ContainerInherit", "None", "Deny")
} else {
$ruleArgs = @("$Admin", "FullControl", "Allow")
}

$path = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f'
$key = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64').OpenSubKey($path, 'ReadWriteSubTree', 'ChangePermissions')
$acl = $key.GetAccessControl()

$rule = [System.Security.AccessControl.RegistryAccessRule]::new.Invoke($ruleArgs)
$acl.ResetAccessRule($rule)
$key.SetAccessControl($acl)
:regdel:

::========================================================================================================================================

::  检查 KMS 激活状态

:k_actinfo

set xpr=
for /f "tokens=* delims=" %%# in ('%psc% "$([DateTime]::Now.addMinutes(%gpr%)).ToString('yyyy-MM-dd HH:mm:ss')" %nul6%') do set "xpr=%%#"
call :dk_color %Green% "%winos% 已激活至 !xpr!"
exit /b

::  检查剩余的 KMS 激活宽限期

:k_checkexp

set gpr=0
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%# in ('"wmic path SoftwareLicensingProduct where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get GracePeriodRemaining /VALUE" %nul6%') do set "gpr=%%#"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%# in ('%psc% "(([WMISEARCHER]'SELECT GracePeriodRemaining FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND Description like ''%%KMSCLIENT%%'' AND PartialProductKey IS NOT NULL').Get()).GracePeriodRemaining | %% {echo ('GracePeriodRemaining='+$_)}" %nul6%') do set "gpr=%%#"
if %gpr% GTR 259200 (set _k38=1) else (set _k38=)
exit /b

::  检查 Windows 已安装密钥的通道

:dk_channel

if %_wmic% EQU 1 for /f "tokens=2 delims==" %%# in ('wmic path SoftwareLicensingProduct where "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey<>null" Get ProductKeyChannel /value %nul6%') do set "_channel=%%#"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%# in ('%psc% "(([WMISEARCHER]'SELECT ProductKeyChannel FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND PartialProductKey IS NOT NULL').Get()).ProductKeyChannel | %% {echo ('ProductKeyChannel='+$_)}" %nul6%') do set "_channel=%%#"
exit /b

::========================================================================================================================================

::  从 pkeyhelper.dll 获取产品序列用于未来的新版本
::  它工作在 Windows 10 1803（17134）及以上版本。

:dk_pkey

call :dk_reflection

set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('SkuGetProductKeyForEdition', 'pkeyhelper.dll', 'Public, Static', 1, [int], @([int], [String], [String].MakeByRefType(), [String].MakeByRefType()), 1, 3);
set d1=%d1% $out = ''; [void]$TypeBuilder.CreateType()::SkuGetProductKeyForEdition(%1, %2, [ref]$out, [ref]$null); $out

set pkey=
for /f %%a in ('%psc% "%d1%"') do if not errorlevel 1 (set pkey=%%a)
exit /b

::  获取从 pkeyhelper.dll 中提取的密钥的通道名称

:dk_pkeychannel

set k=%1
set m=[Runtime.InteropServices.Marshal]
set p=%SystemRoot%\System32\spp\tokens\pkeyconfig\pkeyconfig.xrm-ms

set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('PidGenX', 'pidgenx.dll', 'Public, Static', 1, [int], @([String], [String], [String], [int], [IntPtr], [IntPtr], [IntPtr]), 1, 3);
set d1=%d1% $r = [byte[]]::new(0x04F8); $r[0] = 0xF8; $r[1] = 0x04; $f = %m%::AllocHGlobal(0x04F8); %m%::Copy($r, 0, $f, 0x04F8);
set d1=%d1% [void]$TypeBuilder.CreateType()::PidGenX('%k%', '%p%', '00000', 0, 0, 0, $f); %m%::Copy($f, $r, 0, 0x04F8); %m%::FreeHGlobal($f); [Text.Encoding]::Unicode.GetString($r, 1016, 128)

set pkeychannel=
for /f %%a in ('%psc% "%d1%"') do if not errorlevel 1 (set pkeychannel=%%a)
exit /b

:dk_gvlk

for %%# in (pkeyhelper.dll) do @if "%%~$PATH:#"=="" exit /b
for %%# in (Volume:GVLK) do (
call :dk_pkey %osSKU% '%%#'
if defined pkey call :dk_pkeychannel !pkey!
if /i [!pkeychannel!]==[%%#] (
set key=!pkey!
exit /b
)
)
exit /b

::========================================================================================================================================

::  第 1 列 = 激活 ID
::  第 2 列 = GVLK（通用批量许可密钥）
::  第 3 列 = SKU ID
::  第 4 列 = WMI 版本 ID（仅供参考）
::  第 5 列 = 内部版本分支名称以防相同的版本 ID 在不同的操作系统版本中使用不同的密钥（仅供参考）
::  分隔符  = "_"

:kms38data

set f=
for %%# in (
73111121-5638-40f6-bc11-f1d7b0d64300_NP%f%PR9-FWD%f%CX-D2%f%C8J-H872%f%K-2Y%f%T43___4_Enterprise
7dc26449-db21-4e09-ba37-28f2958506a6_DP%f%NXD-67Y%f%Y9-WW%f%FJJ-RYH9%f%9-RM%f%832___7_ServerStandard_Ge
9bd77860-9b31-4b7b-96ad-2564017315bf_VD%f%YBN-27W%f%PP-V4%f%HQT-9VMD%f%4-VM%f%K7H___7_ServerStandard_FE
de32eafd-aaee-4662-9444-c1befb41bde2_N6%f%9G4-B89%f%J2-4G%f%8F4-WWYC%f%C-J4%f%64C___7_ServerStandard_RS5
8c1c5410-9f39-4805-8c9d-63a07706358f_WC%f%2BQ-8NR%f%M3-FD%f%DYY-2BFG%f%V-KH%f%KQY___7_ServerStandard_RS1
c052f164-cdf6-409a-a0cb-853ba0f0f55a_CN%f%FDQ-2BW%f%8H-9V%f%4WM-TKCP%f%D-MD%f%2QF___8_ServerDatacenter_Ge
ef6cfc9f-8c5d-44ac-9aad-de6a2ea0ae03_WX%f%4NM-KYW%f%YW-QJ%f%JR4-XV3Q%f%B-6V%f%M33___8_ServerDatacenter_FE
34e1ae55-27f8-4950-8877-7a03be5fb181_WM%f%DGN-G9P%f%QG-XV%f%VXX-R3X4%f%3-63%f%DFG___8_ServerDatacenter_RS5
21c56779-b449-4d20-adfc-eece0e1ad74b_CB%f%7KF-BWN%f%84-R7%f%R2Y-793K%f%2-8X%f%DDG___8_ServerDatacenter_RS1
e272e3e2-732f-4c65-a8f0-484747d0d947_DP%f%H2V-TTN%f%VB-4X%f%9Q3-TJR4%f%H-KH%f%JW4__27_EnterpriseN
2de67392-b7a7-462a-b1ca-108dd189f588_W2%f%69N-WFG%f%WX-YV%f%C9B-4J6C%f%9-T8%f%3GX__48_Professional
a80b5abf-76ad-428b-b05d-a47d2dffeebf_MH%f%37W-N47%f%XK-V7%f%XM9-C722%f%7-GC%f%QG9__49_ProfessionalN
034d3cbb-5d4b-4245-b3f8-f84571314078_WV%f%DHN-86M%f%7X-46%f%6P6-VHXV%f%7-YY%f%726__50_ServerSolution_RS5
2b5a1b0f-a5ab-4c54-ac2f-a6d94824a283_JC%f%KRF-N37%f%P4-C2%f%D82-9YXR%f%T-4M%f%63B__50_ServerSolution_RS1
7b9e1751-a8da-4f75-9560-5fadfe3d8e38_3K%f%HY7-WNT%f%83-DG%f%QKR-F7HP%f%R-84%f%4BM__98_CoreN
a9107544-f4a0-4053-a96a-1479abdef912_PV%f%MJN-6DF%f%Y6-9C%f%CP6-7BKT%f%T-D3%f%WVR__99_CoreCountrySpecific
cd918a57-a41b-4c82-8dce-1a538e221a83_7H%f%NRX-D7K%f%GG-3K%f%4RQ-4WPJ%f%4-YT%f%DFH_100_CoreSingleLanguage
58e97c99-f377-4ef1-81d5-4ad5522b5fd8_TX%f%9XD-98N%f%7V-6W%f%MQ6-BX7F%f%G-H8%f%Q99_101_Core
7b4433f4-b1e7-4788-895a-c45378d38253_QN%f%4C6-GBJ%f%D2-FB%f%422-GHWJ%f%K-GJ%f%G2R_110_ServerCloudStorage
8de8eb62-bbe0-40ac-ac17-f75595071ea3_GR%f%FBW-QND%f%C4-6Q%f%BHG-CCK3%f%B-2P%f%R88_120_ServerARM64_RS5
43d9af6e-5e86-4be8-a797-d072a046896c_K9%f%FYF-G6N%f%CK-73%f%M32-XMVP%f%Y-F9%f%DRR_120_ServerARM64_RS4
e0c42288-980c-4788-a014-c080d2e1926e_NW%f%6C2-QMP%f%VW-D7%f%KKK-3GKT%f%6-VC%f%FB2_121_Education
3c102355-d027-42c6-ad23-2e7ef8a02585_2W%f%H4N-8QG%f%BV-H2%f%2JP-CT43%f%Q-MD%f%WWJ_122_EducationN
32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee_M7%f%XTQ-FN8%f%P6-TT%f%KYV-9D4C%f%C-J4%f%62D_125_EnterpriseS_RS5,VB,Ge
2d5a5a60-3040-48bf-beb0-fcd770c20ce0_DC%f%PHK-NFM%f%TC-H8%f%8MJ-PFHP%f%Y-QJ%f%4BJ_125_EnterpriseS_RS1
7b51a46c-0c04-4e8f-9af4-8496cca90d5e_WN%f%MTR-4C8%f%8C-JK%f%8YV-HQ7T%f%2-76%f%DF9_125_EnterpriseS_TH1
7103a333-b8c8-49cc-93ce-d37c09687f92_92%f%NFX-8DJ%f%QP-P6%f%BBQ-THF9%f%C-7C%f%G2H_126_EnterpriseSN_RS5,VB,Ge
9f776d83-7156-45b2-8a5c-359b9c9f22a3_QF%f%FDN-GRT%f%3P-VK%f%WWX-X7T3%f%R-8B%f%639_126_EnterpriseSN_RS1
87b838b7-41b6-4590-8318-5797951d8529_2F%f%77B-TNF%f%GY-69%f%QQF-B8YK%f%P-D6%f%9TJ_126_EnterpriseSN_TH1
39e69c41-42b4-4a0a-abad-8e3c10a797cc_QF%f%ND9-D3Y%f%9C-J3%f%KKY-6RPV%f%P-2D%f%PYV_145_ServerDatacenterACor_FE
90c362e5-0da1-4bfd-b53b-b87d309ade43_6N%f%MRW-2C8%f%FM-D2%f%4W7-TQWM%f%Y-CW%f%H2D_145_ServerDatacenterACor_RS5
e49c08e7-da82-42f8-bde2-b570fbcae76c_2H%f%XDN-KRX%f%HB-GP%f%YC7-YCKF%f%J-7F%f%VDG_145_ServerDatacenterACor_RS3
f5e9429c-f50b-4b98-b15c-ef92eb5cff39_67%f%KN8-4FY%f%JW-24%f%87Q-MQ2J%f%7-4C%f%4RG_146_ServerStandardACor_FE
73e3957c-fc0c-400d-9184-5f7b6f2eb409_N2%f%KJX-J94%f%YW-TQ%f%VFB-DG9Y%f%T-72%f%4CC_146_ServerStandardACor_RS5
61c5ef22-f14f-4553-a824-c4b31e84b100_PT%f%XN8-JFH%f%JM-4W%f%C78-MPCB%f%R-9W%f%4KR_146_ServerStandardACor_RS3
82bbc092-bc50-4e16-8e18-b74fc486aec3_NR%f%G8B-VKK%f%3Q-CX%f%VCJ-9G2X%f%F-6Q%f%84J_161_ProfessionalWorkstation
4b1571d3-bafb-4b40-8087-a961be2caf65_9F%f%NHH-K3H%f%BT-3W%f%4TD-6383%f%H-6X%f%YWF_162_ProfessionalWorkstationN
3f1afc82-f8ac-4f6c-8005-1d233e606eee_6T%f%P4R-GNP%f%TD-KY%f%YHQ-7B7D%f%P-J4%f%47Y_164_ProfessionalEducation
5300b18c-2e33-4dc2-8291-47ffcec746dd_YV%f%WGF-BXN%f%MC-HT%f%QYQ-CPQ9%f%9-66%f%QFC_165_ProfessionalEducationN
45b5aff2-60a0-42f2-bc4b-ec6e5f7b527e_QN%f%7G3-4RM%f%92-MT%f%6QR-PR96%f%6-FV%f%YV7_168_ServerAzureCor_Ge
8c8f0ad3-9a43-4e05-b840-93b8d1475cbc_6N%f%379-GGT%f%MK-23%f%C6M-XVVT%f%C-CK%f%FRQ_168_ServerAzureCor_FE
a99cc1f0-7719-4306-9645-294102fbff95_FD%f%NH6-VW9%f%RW-BX%f%PJ7-4XTY%f%G-23%f%9TB_168_ServerAzureCor_RS5
3dbf341b-5f6c-4fa7-b936-699dce9e263f_VP%f%34G-4NP%f%PG-79%f%JTQ-864T%f%4-R3%f%MQX_168_ServerAzureCor_RS1
e0b2d383-d112-413f-8a80-97f373a5820c_YY%f%VX9-NTF%f%WV-6M%f%DM3-9PT4%f%T-4M%f%68B_171_EnterpriseG
e38454fb-41a4-4f59-a5dc-25080e354730_44%f%RPN-FTY%f%23-9V%f%TTB-MP9B%f%X-T8%f%4FV_172_EnterpriseGN
ec868e65-fadf-4759-b23e-93fe37f2cc29_CP%f%WHC-NT2%f%C7-VY%f%W78-DHDB%f%2-PG%f%3GK_175_ServerRdsh_RS5
e4db50ea-bda1-4566-b047-0ca50abc6f07_7N%f%BT4-WGB%f%QX-MP%f%4H7-QXFF%f%8-YP%f%3KX_175_ServerRdsh_RS3
0df4f814-3f57-4b8b-9a9d-fddadcd69fac_NB%f%TWJ-3DR%f%69-3C%f%4V8-C26M%f%C-GQ%f%9M6_183_CloudE
59eb965c-9150-42b7-a0ec-22151b9897c5_KB%f%N8V-HFG%f%Q4-MG%f%XVD-347P%f%6-PD%f%QGT_191_IoTEnterpriseS_VB,NI
d30136fc-cb4b-416e-a23d-87207abc44a9_6X%f%N7V-PCB%f%DC-BD%f%BRH-8DQY%f%7-G6%f%R44_202_CloudEditionN
ca7df2e3-5ea0-47b8-9ac1-b1be4d8edd69_37%f%D7F-N49%f%CB-WQ%f%R8W-TBJ7%f%3-FM%f%8RX_203_CloudEdition
c2e946d1-cfa2-4523-8c87-30bc696ee584_NQ%f%8HH-FTD%f%TM-6V%f%GY7-TQ3D%f%V-XF%f%BV2_407_ServerTurbine_Ge
19b5e0fb-4431-46bc-bac1-2f1873e4ae73_NT%f%BV8-9K7%f%Q8-V2%f%7C6-M2BT%f%V-KH%f%MXV_407_ServerTurbine_RS5
) do (
for /f "tokens=1-5 delims=_" %%A in ("%%#") do if %osSKU%==%%C (
set skufound=1
if %1==getkey if not defined key echo "!applist!" | find /i "%%A" %nul1% && set key=%%B
)
)
exit /b

::========================================================================================================================================

::  如果当前版本不支持 KMS38 激活，以下代码用于获取备用版本名称和密钥

::  第 1 列 = 当前 SKU ID
::  第 2 列 = 当前版本名称
::  第 3 列 = 当前版本激活 ID
::  第 4 列 = 备用版本激活 ID
::  第 5 列 = 备用版本 GVLK
::  第 6 列 = 备用版本名称
::  分隔符  = _


:kms38fallback

set notfoundaltactID=
if %_NoEditionChange%==1 exit /b

for %%# in (
188_IoTEnterprise__________________8ab9bdd1-1f67-4997-82d9-8878520837d9_73111121-5638-40f6-bc11-f1d7b0d64300_NPP%f%R9-FWD%f%CX-D2%f%C8J-H872%f%K-2Y%f%T43_Enterprise
206_IoTEnterpriseK_________________80083eae-7031-4394-9e88-4901973d56fe_73111121-5638-40f6-bc11-f1d7b0d64300_NPP%f%R9-FWD%f%CX-D2%f%C8J-H872%f%K-2Y%f%T43_Enterprise
191_IoTEnterpriseS-2021____________ed655016-a9e8-4434-95d9-4345352c2552_32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee_M7X%f%TQ-FN8%f%P6-TT%f%KYV-9D4C%f%C-J4%f%62D_EnterpriseS-2021
205_IoTEnterpriseSK________________d4f9b41f-205c-405e-8e08-3d16e88e02be_59eb965c-9150-42b7-a0ec-22151b9897c5_KBN%f%8V-HFG%f%Q4-MG%f%XVD-347P%f%6-PD%f%QGT_IoTEnterpriseS
138_ProfessionalSingleLanguage_____a48938aa-62fa-4966-9d44-9f04da3f72f2_2de67392-b7a7-462a-b1ca-108dd189f588_W26%f%9N-WFG%f%WX-YV%f%C9B-4J6C%f%9-T8%f%3GX_Professional
139_ProfessionalCountrySpecific____f7af7d09-40e4-419c-a49b-eae366689ebd_2de67392-b7a7-462a-b1ca-108dd189f588_W26%f%9N-WFG%f%WX-YV%f%C9B-4J6C%f%9-T8%f%3GX_Professional
139_ProfessionalCountrySpecific-Zn_01eb852c-424d-4060-94b8-c10d799d7364_2de67392-b7a7-462a-b1ca-108dd189f588_W26%f%9N-WFG%f%WX-YV%f%C9B-4J6C%f%9-T8%f%3GX_Professional
) do (
for /f "tokens=1-6 delims=_" %%A in ("%%#") do if %osSKU%==%%A (
echo "!applist!" | find /i "%%C" %nul1% && (
echo "!applist!" | find /i "%%D" %nul1% && (
set altkey=%%E
set altedition=%%F
) || (
set altedition=%%F
set notfoundaltactID=1
)
)
)
)
exit /b

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:KMSActivation
@setlocal DisableDelayedExpansion
@echo off

cls
color 07
title 在线 KMS 激活 %masver%

::  你不应编辑此处下方的任何内容。

set WMI_VBS=0
set _Debug=0
set Silent=0
set Logger=0
set AutoR2V=1
set SkipKMS38=1
set vNextOverride=1
set ActWindows=1
set ActOffice=1

set _uni=
set _args=
set _elev=
set _renetask=
set _renacttask=
set _unattended=
set _unattendedact=

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args (
echo "%_args%" | find /i "/KMS" >nul && set _unattended=1

for %%A in (%_args%) do (
if /i "%%A"=="-el"  (set _elev=1
) else if /i "%%A"=="/KMS-RenewalTask"  (set _renetask=1
) else if /i "%%A"=="/KMS-ActAndRenewalTask" (set _renacttask=1
) else if /i "%%A"=="/KMS-Uninstall" (set _uni=1
) else if /i "%%A"=="/KMS-Windows"   (set ActWindows=1&set ActOffice=0&set _unattendedact=1
) else if /i "%%A"=="/KMS-Office"   (set ActWindows=0&set ActOffice=1&set _unattendedact=1
) else if /i "%%A"=="/KMS-WindowsOffice"  (set ActWindows=1&set ActOffice=1&set _unattendedact=1
) else if /i "%%A"=="/KMS-KeepvNext"  (set vNextOverride=0
) else if /i "%%A"=="/KMS-Debug"   (set _Debug=1
) else if /i "%%A"=="/KMS-Logger"   (set Logger=1&set Silent=1
)
)
)

::========================================================================================================================================

set "nul=>nul 2>&1"
set psc=powershell.exe
set winbuild=1
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G

set _NCS=1
if %winbuild% LSS 10586 set _NCS=0
if %winbuild% GEQ 10586 reg query "HKCU\Console" /v ForceV2 2>nul | find /i "0x0" 1>nul && (set _NCS=0)

call :_colorprep
set "_buf={$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=300;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}"

set "nceline=echo. &echo ==== 错误 ==== &echo."
set "eline=echo. &call :_color %Red% "==== 错误 ====" &echo."
if %_Debug% EQU 1 set _unattended=1

::========================================================================================================================================

::  修复路径名称中的特殊字符限制

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

set _PSarg="""%~f0""" -el %_args%

set "_ttemp=%userprofile%\AppData\Local\Temp"
set "_Local=%LocalAppData%"

setlocal EnableDelayedExpansion

::========================================================================================================================================

if %~z0 GEQ 300000 (set "_exitmsg=返回") else (set "_exitmsg=退出")

::  检查非 x86 Windows

set notx86=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PROCESSOR_ARCHITECTURE') do set arch=%%b
if /i not "%arch%"=="x86" set notx86=1

::========================================================================================================================================

for %%# in (wmic.exe) do @if "%%~$PATH:#"=="" (
%nceline%
echo 在系统中找不到 wmic.exe。
if %winbuild% GEQ 22621 echo 确保 WMIC 已在可选功能中启用。
goto Done
)

wmic path Win32_ComputerSystem get CreationClassName /value 2>nul | find /i "ComputerSystem" 1>nul || (
%nceline%
echo wmic.exe 在系统中未响应。
echo:
echo 在 MAS 中，请转到疑难解答并运行修复 WMI 选项。
goto Done
)

set _WSH=1
reg query "HKCU\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled 2>nul | find /i "0x0" 1>nul && (set _WSH=0)
reg query "HKLM\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled 2>nul | find /i "0x0" 1>nul && (set _WSH=0)

if %_WSH% EQU 0 (
reg add "HKLM\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f %nul%
reg add "HKCU\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f %nul%
if defined notx86 reg add "HKLM\Software\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f /reg:32 %nul%
)

::========================================================================================================================================

if defined _uni goto _Complete_Uninstall

if defined _renetask set ActTask=&call:RenTask&timeout /t 2
cls
if defined _renacttask set ActTask=1&call:RenTask&timeout /t 2
cls
if defined _unattended if not defined _unattendedact goto Done

::========================================================================================================================================

set "_title=在线 KMS 激活 %masver%"
set _gui=

:_KMS_Menu

set sub_next=0
set sub_o365=0
set sub_proj=0
set sub_vsio=0
set kNext=HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\LicensingNext
reg query %kNext% /v MigrationToV5Done 2>nul | find /i "0x1" %nul% && call :officeSub %nul%

set _tskinstalled=
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\taskcache\tasks" /f Path /s | find /i "\Activation-Renewal" >nul && (
find /i "Ver:1.9" "%ProgramFiles%\Activation-Renewal\Activation_task.cmd" %nul% && set _tskinstalled=1
)

set _oldtsk=
if not defined _tskinstalled (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\taskcache\tasks" /f Path /s | findstr /i "\Activation-Renewal \Online_KMS_Activation_Script-Renewal" >nul && (
set _oldtsk=1
)
)

if defined _unattended (
call :Activation_Start
timeout /t 2
goto Done
)

cls
set _gui=1
title %_title%
mode con: cols=76 lines=30

echo.
echo.
echo.
echo.
echo.       ______________________________________________________________
echo.
echo.              [1] 激活     - Windows
echo.              [2] 激活     - Office
echo.              [3] 激活     - 所有
echo.
if defined _tskinstalled call :_color2 %_White% "              [4] 安装自动续期              " %_Green% "[已安装]"
if defined _oldtsk       call :_color2 %_White% "              [4] 安装自动续期              " %_Red% "[旧安装]"
if not defined _tskinstalled if not defined _oldtsk echo.              [4] 安装自动续期
echo.              [5] 卸载
echo.              _______________________________________________  
echo.
if %_Debug%==0 (
echo.              [6] 启用调试模式              [否]
) else (
call :_color2 %_White% "              [6] 启用调试模式              " %_Red% "[是]"
)
if %vNextOverride% EQU 1 (
if %sub_next% EQU 1 (
call :_color2 %_White% "              [7] 覆盖 Office vNext         " %_Red% "[是]"
) else (
echo               [7] 覆盖 Office vNext         [是]
)
) else (
if %sub_next% EQU 1 (
call :_color2 %_White% "              [7] 覆盖 Office vNext         " %_Yellow% "[否]"
) else (
echo               [7] 覆盖 Office vNext         [否]
)
)
echo.              _______________________________________________       
echo.
echo.              [0] %_exitmsg%
echo.       ______________________________________________________________
echo.
call :_color2 %_White% "           " %_Green% "请输入一个菜单选项 [1,2,3,4,5,6,7,0]"
choice /C:12345670 /N
set _el=%errorlevel%

if %_el%==8 exit /b
if %_el%==7 (if %vNextOverride% EQU 0 (set vNextOverride=1) else (set vNextOverride=0))&goto _KMS_Menu
if %_el%==6 (if %_Debug%==0 (set _Debug=1) else (set _Debug=0)) &goto _KMS_Menu
if %_el%==5 call:_Complete_Uninstall&cls&goto _KMS_Menu
if %_el%==4 set ActTask=&call:RenTask&goto _KMS_Menu
if %_el%==3 cls&setlocal&set "ActWindows=1"&set "ActOffice=1"&call :Activation_Start&endlocal&cls&goto _KMS_Menu
if %_el%==2 cls&setlocal&set "ActWindows=0"&set "ActOffice=1"&call :Activation_Start&endlocal&cls&goto _KMS_Menu
if %_el%==1 cls&setlocal&set "ActWindows=1"&set "ActOffice=0"&call :Activation_Start&endlocal&cls&goto _KMS_Menu
goto _KMS_Menu

::========================================================================================================================================

:Done

if defined _unattended exit /b

echo.
echo 请按任意键退出脚本……
pause >nul
exit /b

:=========================================================================================================================================

:Activation_Start

@setlocal DisableDelayedExpansion

set nil=
for %%# in (SppE%nil%xtComObj.exe,sppsvc.exe,osppsvc.exe) do (
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ima%nil%ge File Execu%nil%tion Options\%%#" /f %nul%)
)

call :Clear-KMS-Cache %nul%

set "_Null=1>nul 2>nul"
set KMS_Port=1688
if %_Debug% EQU 1 set _unattended=1
set "_run=nul"
if %Logger% EQU 1 set _run="%~dpn0_Silent.log"

set "SysPath=%SystemRoot%\System32"
set "Path=%SystemRoot%\System32;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\;%LocalAppData%\Microsoft\WindowsApps\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "SysPath=%SystemRoot%\Sysnative"
set "Path=%SystemRoot%\Sysnative;%SystemRoot%\Sysnative\Wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%LocalAppData%\Microsoft\WindowsApps\;%Path%"
)
set "_bit=64"
set "_wow=1"
if /i "%PROCESSOR_ARCHITECTURE%"=="amd64" set "xBit=x64"&set "xOS=x64"
if /i "%PROCESSOR_ARCHITECTURE%"=="arm64" set "xBit=x86"&set "xOS=A64"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if "%PROCESSOR_ARCHITEW6432%"=="" set "xBit=x86"&set "xOS=x86"&set "_wow=0"&set "_bit=32"
if /i "%PROCESSOR_ARCHITEW6432%"=="amd64" set "xBit=x64"&set "xOS=x64"
if /i "%PROCESSOR_ARCHITEW6432%"=="arm64" set "xBit=x86"&set "xOS=A64"
if not defined xBit set "xBit=x64"&set "xOS=x64"

set _cwmi=0
for %%# in (wmic.exe) do @if not "%%~$PATH:#"=="" (
wmic path Win32_ComputerSystem get CreationClassName /value 2>nul | find /i "ComputerSystem" 1>nul && set _cwmi=1
)

set "_Local=%LocalAppData%"
set "_temp=%SystemRoot%\Temp"
set "_log=%~dpn0"
set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"
set _UNC=0
if "%_work:~0,2%"=="\\" (
set _UNC=1
) else (
net use %~d0 %_Null%
if not errorlevel 1 set _UNC=1
)
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "_dsk=%%b"
if exist "%PUBLIC%\Desktop\desktop.ini" set "_dsk=%PUBLIC%\Desktop"
set "_mO21a=检测到 Office 2021 C2R 零售版本已经激活"
set "_mO19a=检测到 Office 2019 C2R 零售版本已经激活"
set "_mO16a=检测到 Office 2016 C2R 零售版本已经激活"
set "_mO15a=检测到 Office 2013 C2R 零售版本已经激活"
set "_mO21c=检测到 Office 2021 C2R 零售版本无法转换为批量版本"
set "_mO19c=检测到 Office 2019 C2R 零售版本无法转换为批量版本"
set "_mO16c=检测到 Office 2016 C2R 零售版本无法转换为批量版本"
set "_mO15c=检测到 Office 2013 C2R 零售版本无法转换为批量版本"
set "_mO14c=检测到 Office 2010 C2R 零售版本不支持此脚本"
set "_mO14m=检测到 Office 2010 MSI 零售版本不支持此脚本"
set "_mO15m=检测到 Office 2013 MSI 零售版本不支持此脚本"
set "_mO16m=检测到 Office 2016 MSI 零售版本不支持此脚本"
set "_mOuwp=检测到 Office 365/2016 UWP 不支持此脚本"
set DO15Ids=ProPlus,Standard,Access,Lync,Excel,Groove,InfoPath,OneNote,Outlook,PowerPoint,Publisher,Word
set DO16Ids=ProPlus,Standard,Access,SkypeforBusiness,Excel,Outlook,PowerPoint,Publisher,Word
set LV16Ids=Mondo,ProPlus,ProjectPro,VisioPro,Standard,ProjectStd,VisioStd,Access,SkypeforBusiness,OneNote,Excel,Outlook,PowerPoint,Publisher,Word
set LR16Ids=%LV16Ids%,Professional,HomeBusiness,HomeStudent,O365Business,O365SmallBusPrem,O365HomePrem,O365EduCloud
set "ESUEditions=Enterprise,EnterpriseE,EnterpriseN,Professional,ProfessionalE,ProfessionalN,Ultimate,UltimateE,UltimateN"
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
set "ESUEditions=ServerDatacenter,ServerDatacenterCore,ServerDatacenterV,ServerDatacenterVCore,ServerStandard,ServerStandardCore,ServerStandardV,ServerStandardVCore,ServerEnterprise,ServerEnterpriseCore,ServerEnterpriseV,ServerEnterpriseVCore"
)
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
set UBR=0
if %winbuild% GEQ 7601 for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v UBR 2^>nul') do if not errorlevel 1 set /a UBR=%%b
set "_csq=cscript.exe //NoLogo //Job:WmiQuery "%~nx0?.wsf""
set "_csm=cscript.exe //NoLogo //Job:WmiMethod "%~nx0?.wsf""
set "_csp=cscript.exe //NoLogo //Job:WmiPKey "%~nx0?.wsf""
set "_csd=cscript.exe //NoLogo //Job:MPS "%~nx0?.wsf""
if %_cwmi% EQU 0 set WMI_VBS=1
if %WMI_VBS% EQU 0 (
set "_zz1=wmic path"
set "_zz2=where"
set "_zz3=get"
set "_zz4=/value"
set "_zz5=("
set "_zz6=)"
set "_zz7="wmic path"
set "_zz8=/value""
) else (
set "_zz1=%_csq%"
set "_zz2="
set "_zz3="
set "_zz4="
set "_zz5=""
set "_zz6=""
set "_zz7=%_csq%"
set "_zz8="
)

setlocal EnableDelayedExpansion
pushd "!_work!"

if not defined _unattended (
mode con cols=98 lines=31
%psc% "&%_buf%"
title %_title%
) else (
title 在线 KMS 激活 %masver%
)

if defined _gui if %_Debug%==1 mode con cols=98 lines=30

if %_Debug% EQU 0 (
  set "_Nul1=1>nul"
  set "_Nul2=2>nul"
  set "_Nul6=2^>nul"
  set "_Nul3=1>nul 2>nul"
  set "_Pause=pause >nul"
  if %Silent% EQU 0 (call :Begin) else (call :Begin >!_run! 2>&1)
) else (
  set "_Nul1="
  set "_Nul2="
  set "_Nul6="
  set "_Nul3="
  set "_log=!_dsk!\%~n0"
  if %Silent% EQU 0 (
  echo.
  echo 正在调试模式下运行……
  if not defined _args (echo 当完成之后，此窗口将会关闭) else (echo请稍候……)
  echo.
  echo 正在写入调试日志到：
  echo "!_log!_Debug.log"
  )
  @echo on
  @prompt $G
  @call :Begin >"!_log!_tmp.log" 2>&1 &cmd /u /c type "!_log!_tmp.log">"!_log!_Debug.log"&del "!_log!_tmp.log"
)
@echo off
if defined _gui if %_Debug%==1 (
echo.
call :_color %_Yellow% "请按任意键返回……"
pause >nul
exit /b
)
@exit /b

:Begin

::========================================================================================================================================

set act_failed=0
set /a act_attempt=0

echo.
echo 正在初始化……

::  检查 Internet 连接。即使禁用 ICMP 回显也可以工作。

call :setserv
for %%a in (%srvlist%) do (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (
if not [%%#]==[] goto IntConnected
)
)

nslookup dns.msftncsi.com 2>nul | find "131.107.255.255" 1>nul
if [%errorlevel%]==[0] goto IntConnected

cls
if %_Debug%==1 (
echo 错误：Internet 未连接。
exit /b
)

if defined _unattended (
echo.
call :_color %_Red% "Internet 未连接，无论如何仍继续此过程。"
) else (
%eline%
echo Internet 未连接。
echo:
call :_color %_Yellow% "请按任意键返回……"
pause >nul
exit /b
)

:IntConnected

call :getserv

::========================================================================================================================================

set "_wApp=55c92734-d682-4d71-983e-d6ec3f16059f"
set "_oApp=0ff1ce15-a989-479d-af46-f275c6370663"
set "_oA14=59a52881-a989-479d-af46-f275c6370663"
set "IFEO=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
set "OPPk=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
set SSppHook=0
for /f %%A in ('dir /b /ad %SysPath%\spp\tokens\skus') do (
  if %winbuild% GEQ 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*GVLK*.xrm-ms" set SSppHook=1
  if %winbuild% LSS 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*VLKMS*.xrm-ms" set SSppHook=1
  if %winbuild% LSS 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*VL-BYPASS*.xrm-ms" set SSppHook=1
)
set OsppHook=1
sc query osppsvc %_Nul3%
if %errorlevel% EQU 1060 set OsppHook=0

set ESU_KMS=0
if %winbuild% LSS 9200 for /f %%A in ('dir /b /ad %SysPath%\spp\tokens\channels') do (
  if exist "%SysPath%\spp\tokens\channels\%%A\*VL-BYPASS*.xrm-ms" set ESU_KMS=1
)
if %ESU_KMS% EQU 1 (set "adoff=and LicenseDependsOn is NULL"&set "addon=and LicenseDependsOn is not NULL") else (set "adoff="&set "addon=")
set ESU_EDT=0
if %ESU_KMS% EQU 1 for %%A in (%ESUEditions%) do (
  if exist "%SysPath%\spp\tokens\skus\Security-SPP-Component-SKU-%%A\*.xrm-ms" set ESU_EDT=1
)
::  if %ESU_EDT% EQU 1 set SSppHook=1
set ESU_ADD=0

if %winbuild% GEQ 9200 (
  set OSType=Win8
  set SppVer=SppExtComObj.exe
) else if %winbuild% GEQ 7600 (
  set OSType=Win7
  set SppVer=sppsvc.exe
) else (
  goto :UnsupportedVersion
)
if %OSType% EQU Win8 reg query "%IFEO%\sppsvc.exe" %_Nul3% && (
reg delete "%IFEO%\sppsvc.exe" /f %_Nul3%
call :StopService sppsvc
)

if %ActWindows% EQU 0 if %ActOffice% EQU 0 set ActWindows=1
set _AUR=1
if %winbuild% GEQ 9600 (
  reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v NoGenTicket /t REG_DWORD /d 1 /f %_Nul3%
  if %winbuild% EQU 14393 reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v NoAcquireGT /t REG_DWORD /d 1 /f %_Nul3%
)
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc

:ReturnHook
call :UpdateOSPPEntry osppsvc.exe

SET Win10Gov=0
SET "EditionWMI="
SET "EditionID="
IF %winbuild% LSS 14393 if %SSppHook% NEQ 0 GOTO :Main
SET "RegKey=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages"
SET "Pattern=Microsoft-Windows-*Edition~31bf3856ad364e35"
SET "EditionPKG=FFFFFFFF"
FOR /F "TOKENS=8 DELIMS=\" %%A IN ('REG QUERY "%RegKey%" /f "%Pattern%" /k %_Nul6% ^| FIND /I "CurrentVersion"') DO (
  REG QUERY "%RegKey%\%%A" /v "CurrentState" %_Nul2% | FIND /I "0x70" %_Nul1% && (
    FOR /F "TOKENS=3 DELIMS=-~" %%B IN ('ECHO %%A') DO SET "EditionPKG=%%B"
  )
)
IF /I "%EditionPKG:~-7%"=="Edition" (
SET "EditionID=%EditionPKG:~0,-7%"
) ELSE (
FOR /F "TOKENS=3 DELIMS=: " %%A IN ('DISM /English /Online /Get-CurrentEdition %_Nul6% ^| FIND /I "Current Edition :"') DO SET "EditionID=%%A"
)
net start sppsvc /y %_Nul3%
set "_qr=%_zz7% SoftwareLicensingProduct %_zz2% %_zz5%ApplicationID='%_wApp%' %adoff% AND PartialProductKey is not NULL%_zz6% %_zz3% LicenseFamily %_zz8%"
FOR /F "TOKENS=2 DELIMS==" %%A IN ('%_qr% %_Nul6%') DO SET "EditionWMI=%%A"
IF "%EditionWMI%"=="" (
IF %winbuild% GEQ 17063 FOR /F "SKIP=2 TOKENS=2*" %%A IN ('REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionId') DO SET "EditionID=%%B"
IF %winbuild% LSS 14393 (
  FOR /F "SKIP=2 TOKENS=2*" %%A IN ('REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionId') DO SET "EditionID=%%B"
  GOTO :Main
  )
)
IF NOT "%EditionWMI%"=="" SET "EditionID=%EditionWMI%"
IF /I "%EditionID%"=="IoTEnterprise" SET "EditionID=Enterprise"
IF /I "%EditionID%"=="IoTEnterpriseS" IF %winbuild% LSS 22610 (
SET "EditionID=EnterpriseS"
IF %winbuild% GEQ 19041 IF %UBR% GEQ 2788 SET "EditionID=IoTEnterpriseS"
)
IF /I "%EditionID%"=="ProfessionalSingleLanguage" SET "EditionID=Professional"
IF /I "%EditionID%"=="ProfessionalCountrySpecific" SET "EditionID=Professional"
IF /I "%EditionID%"=="EnterpriseG" SET Win10Gov=1
IF /I "%EditionID%"=="EnterpriseGN" SET Win10Gov=1

:Main
if defined EditionID (set "_winos=Windows %EditionID% edition") else (set "_winos=检测到 Windows")
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName %_Nul6%') do if not errorlevel 1 set "_winos=%%b"
set "nKMS=不支持 KMS 激活……"
set "nEval=评估版本无法激活。请安装完整版本 Windows 操作系统。"
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" set _eval=1
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalEdition~*.mum" set "nEval=Server 评估版本无法激活。请转换为完整的 Server 操作系统。"
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalCorEdition~*.mum" set _eval=1&set "nEval=Server 评估版本无法激活。请转换为完整的 Server 操作系统。"
set "_C16R="
reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath %_Nul3% && for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (
reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v ProductReleaseIds %_Nul3% && set "_C16R=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
)
if not defined _C16R reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v InstallPath %_Nul3% && for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (
reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration /v ProductReleaseIds %_Nul3% && set "_C16R=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
)
set "_C15R="
reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath %_Nul3% && for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (
reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds %_Nul3% && call set "_C15R=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration"
if not defined _C15R reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid %_Nul3% && call set "_C15R=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag"
)
set "_C14R="
if %_wow%==0 (reg query HKLM\SOFTWARE\Microsoft\Office\14.0\CVH /f Click2run /k %_Nul3% && set "_C14R=1") else (reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\CVH /f Click2run /k %_Nul3% && set "_C14R=1")
for %%A in (14,15,16,19,21) do call :officeLoc %%A
if %_O14MSI% EQU 1 set "_C14R="

set S_OK=1
call :RunSPP
if %ActOffice% NEQ 0 call :RunOSPP
if %ActOffice% EQU 0 (echo.&echo Office 激活已关闭……)

if exist "!_temp!\crv*.txt" del /f /q "!_temp!\crv*.txt"
if exist "!_temp!\*chk.txt" del /f /q "!_temp!\*chk.txt"
if exist "!_temp!\slmgr.vbs" del /f /q "!_temp!\slmgr.vbs"
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc

sc start sppsvc trigger=timer;sessionid=0 %_Nul3%

goto TheEnd

:RunSPP
set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService
set W1nd0ws=1
set WinPerm=0
set WinVL=0
set Off1ce=0
set RanR2V=0
set aC2R21=0
set aC2R19=0
set aC2R16=0
set aC2R15=0
if %winbuild% GEQ 9200 if %ActOffice% NEQ 0 call :sppoff
set "_qr=%_zz1% %spp% %_zz2% %_zz5%Description like '%%KMSCLIENT%%' %_zz6% %_zz3% Name %_zz4%"
%_qr% %_Nul2% | findstr /i Windows %_Nul1% && (set WinVL=1)
if %WinVL% EQU 0 (
if %ActWindows% EQU 0 (
  echo.&echo Windows 激活已关闭……
  ) else (
  if %SSppHook% EQU 0 (
    echo.&echo %_winos% %nKMS%
    if defined _eval echo %nEval%
    ) else (
    echo.&echo 检查 Windows 的 KMS 激活 ID 失败。&call :CheckWS
    exit /b
    )
  )
)
if %WinVL% EQU 0 if %Off1ce% EQU 0 exit /b
set _gvlk=0
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL%_zz6% %_zz3% Name %_zz4%"
if %winbuild% GEQ 10240 %_qr% %_Nul2% | findstr /i Windows %_Nul1% && (set _gvlk=1)
set gpr=0
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL%_zz6% %_zz3% GracePeriodRemaining %_zz8%"
if %winbuild% GEQ 10240 if %SkipKMS38% NEQ 0 if %_gvlk% EQU 1 for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set "gpr=%%A"
set "_qr=%_zz1% %spp% %_zz2% "ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL" %_zz3% LicenseFamily %_zz4%"
if %gpr% NEQ 0 if %gpr% GTR 259200 (
set W1nd0ws=0
%_qr% %_Nul2% | findstr /i EnterpriseG %_Nul1% && (call set W1nd0ws=1)
)
set "_qr=%_zz7% %sps% %_zz3% Version %_zz8%"
for /f "tokens=2 delims==" %%A in ('%_qr%') do set slsv=%%A
reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" %_Nul3%
reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" %_Nul3%
if %winbuild% GEQ 9200 (
if not %xOS%==x86 (
reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" /reg:32 %_Nul3%
reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" /reg:32 %_Nul3%
reg delete "HKLM\%SPPk%\%_oApp%" /f /reg:32 %_Null%
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" /reg:32 %_Nul3%
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" /reg:32 %_Nul3%
)
reg delete "HKLM\%SPPk%\%_oApp%" /f %_Null%
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" %_Nul3%
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" %_Nul3%
)
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%' %_zz6% %_zz3% ID %_zz8%"
if %W1nd0ws% EQU 0 for /f "tokens=2 delims==" %%G in ('%_qr%') do (set app=%%G&call :sppchkwin)
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%' %adoff% %_zz6% %_zz3% ID %_zz8%"
if %W1nd0ws% EQU 1 if %ActWindows% NEQ 0 for /f "tokens=2 delims==" %%G in ('%_qr%') do (set app=%%G&call :sppchkwin)
::  set "_qr=%_zz7% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%' %addon% %_zz6% %_zz3% ID %_zz8%"
::  if %ESU_EDT% EQU 1 if %ActWindows% NEQ 0 for /f "tokens=2 delims==" %%G in ('%_qr%') do (set app=%%G&call :esuchk)
if %W1nd0ws% EQU 1 if %ActWindows% EQU 0 (echo.&echo Windows 激活已关闭……)
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ApplicationID='%_oApp%' and Description like '%%KMSCLIENT%%' %_zz6% %_zz3% ID %_zz8%"
if %Off1ce% EQU 1 if %ActOffice% NEQ 0 for /f "tokens=2 delims==" %%G in ('%_qr%') do (set app=%%G&call :sppchkoff 1)
reg delete "HKLM\%SPPk%" /f /v DisableDnsPublishing %_Null%
reg delete "HKLM\%SPPk%" /f /v DisableKeyManagementServiceHostCaching %_Null%
exit /b

:sppoff
set OffUWP=0
if %winbuild% GEQ 10240 reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msoxmled.exe" %_Nul3% && (
dir /b "%ProgramFiles%\WindowsApps\Microsoft.Office.Desktop*" %_Nul3% && set OffUWP=1
if not %xOS%==x86 dir /b "%ProgramW6432%\WindowsApps\Microsoft.Office.Desktop*" %_Nul3% && set OffUWP=1
)
rem 无已安装的方案
if %loc_off21% EQU 0 if %loc_off19% EQU 0 if %loc_off16% EQU 0 if %loc_off15% EQU 0 (
if %winbuild% GEQ 9200 (
  if %OffUWP% EQU 0 (echo.&echo 未检测到已安装的 Office 2013-2021 产品……) else (echo.&echo %_mOuwp%)
  exit /b
  )
if %winbuild% LSS 9200 (if %loc_off14% EQU 0 (echo.&echo 未检测到已安装的 Office %aword% 产品……&exit /b))
)
if %vNextOverride% EQU 1 if %AutoR2V% EQU 1 (
set sub_o365=0
set sub_proj=0
set sub_vsio=0
if %sub_next% EQU 1 reg delete HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing /f %_Nul3%
)
set Off1ce=1
set _sC2R=sppoff
set _fC2R=ReturnSPP

set vol_off14=0&set vol_off15=0&set vol_off16=0&set vol_off19=0&set vol_off21=0
set "_qr=%_zz1% %spp% %_zz2% %_zz5%Description like '%%KMSCLIENT%%' AND NOT Name like '%%MondoR_KMS_Automation%%' %_zz6% %_zz3% Name %_zz4%"
%_qr% > "!_temp!\sppchk.txt" 2>&1
find /i "Office 21" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off21=1)
find /i "Office 19" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off19=1)
find /i "Office 16" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off16=1)
find /i "Office 15" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off15=1)
if %winbuild% LSS 9200 find /i "Office 14" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off14=1)
for %%A in (14,15,16,19,21) do if !loc_off%%A! EQU 0 set vol_off%%A=0
set "_qr=%_zz1% %spp% %_zz2% "ApplicationID='%_oApp%' AND LicenseFamily like 'Office16O365%%'" %_zz3% LicenseFamily %_zz4%"
if %vol_off16% EQU 1 find /i "Office16MondoVL_KMS_Client" "!_temp!\sppchk.txt" %_Nul1% && (
%_qr% %_Nul2% | find /i "O365" %_Nul1% || (set vol_off16=0)
)
set "_qr=%_zz1% %spp% %_zz2% "ApplicationID='%_oApp%' AND LicenseFamily like 'OfficeO365%%'" %_zz3% LicenseFamily %_zz4%"
if %vol_off15% EQU 1 find /i "OfficeMondoVL_KMS_Client" "!_temp!\sppchk.txt" %_Nul1% && (
%_qr% %_Nul2% | find /i "O365" %_Nul1% || (set vol_off15=0)
)

set ret_off14=0&set ret_off15=0&set ret_off16=0&set ret_off19=0&set ret_off21=0
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ApplicationID='%_oApp%' AND NOT Name like '%%O365%%' %_zz6% %_zz3% Name %_zz4%"
%_qr% > "!_temp!\sppchk.txt" 2>&1
find /i "R_Retail" "!_temp!\sppchk.txt" %_Nul2% | find /i "Office 21" %_Nul1% && (set ret_off21=1)
find /i "R_Retail" "!_temp!\sppchk.txt" %_Nul2% | find /i "Office 19" %_Nul1% && (set ret_off19=1)
find /i "R_Retail" "!_temp!\sppchk.txt" %_Nul2% | find /i "Office 16" %_Nul1% && (set ret_off16=1)
find /i "R_Retail" "!_temp!\sppchk.txt" %_Nul2% | find /i "Office 15" %_Nul1% && (set ret_off15=1)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ApplicationID='%_oA14%'%_zz6% %_zz3% Description %_zz4%"
if %winbuild% LSS 9200 if %vol_off14% EQU 0 %_qr% %_Nul2% | findstr /i channel %_Nul1% && (set ret_off14=1)

set run_off21=0&set prr_off21=0&set prv_off21=0
if %loc_off21% EQU 1 if %ret_off21% EQU 1 if %_O16MSI% EQU 0 if %vol_off21% EQU 0 set run_off21=1
if %loc_off21% EQU 1 if %ret_off21% EQU 1 if %_O16MSI% EQU 0 if %vol_off21% EQU 1 (
for %%a in (%DO16Ids%) do find /i "Office21%%a2021R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off21+=1
  find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off21+=1
  )
for %%a in (Professional) do find /i "Office21%%a2021R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off21+=1
  find /i "Office21ProPlus2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off21+=1
  )
for %%a in (HomeBusiness,HomeStudent) do find /i "Office21%%a2021R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off21+=1
  find /i "Office21Standard2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off21+=1
  )
if %sub_proj% EQU 0 for %%a in (ProjectPro,ProjectStd) do find /i "Office21%%a2021R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off21+=1
  find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off21+=1
  )
if %sub_vsio% EQU 0 for %%a in (VisioPro,VisioStd) do find /i "Office21%%a2021R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off21+=1
  find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off21+=1
  )
)
if %loc_off21% EQU 1 if %ret_off21% EQU 1 if %_O16MSI% EQU 0 if %vol_off21% EQU 1 if %prv_off21% LSS %prr_off21% (set vol_off21=0&set run_off21=1)

set run_off19=0&set prr_off19=0&set prv_off19=0
if %loc_off19% EQU 1 if %ret_off19% EQU 1 if %_O16MSI% EQU 0 if %vol_off19% EQU 0 set run_off19=1
if %loc_off19% EQU 1 if %ret_off19% EQU 1 if %_O16MSI% EQU 0 if %vol_off19% EQU 1 (
for %%a in (%DO16Ids%) do find /i "Office19%%a2019R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off19+=1
  find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off19+=1
  )
for %%a in (Professional) do find /i "Office19%%a2019R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off19+=1
  find /i "Office19ProPlus2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off19+=1
  )
for %%a in (HomeBusiness,HomeStudent) do find /i "Office19%%a2019R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off19+=1
  find /i "Office19Standard2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off19+=1
  )
if %sub_proj% EQU 0 for %%a in (ProjectPro,ProjectStd) do find /i "Office19%%a2019R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off19+=1
  find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off19+=1
  )
if %sub_vsio% EQU 0 for %%a in (VisioPro,VisioStd) do find /i "Office19%%a2019R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off19+=1
  find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off19+=1
  )
)
if %loc_off19% EQU 1 if %ret_off19% EQU 1 if %_O16MSI% EQU 0 if %vol_off19% EQU 1 if %prv_off19% LSS %prr_off19% (set vol_off19=0&set run_off19=1)

set run_off16=0&set prr_off16=0&set prv_off16=0
if %loc_off16% EQU 1 if %ret_off16% EQU 1 if %_O16MSI% EQU 0 if defined _C16R (
for %%a in (%DO16Ids%) do find /i "Office16%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off16+=1
  if %vol_off16% EQU 1 if %vol_off21% EQU 0 if %vol_off19% EQU 0 find /i "Office16%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off21% EQU 1 find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off19% EQU 1 find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  )
for %%a in (Professional) do find /i "Office16%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off16+=1
  if %vol_off16% EQU 1 if %vol_off21% EQU 0 if %vol_off19% EQU 0 find /i "Office16ProPlusVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off21% EQU 1 find /i "Office21ProPlus2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off19% EQU 1 find /i "Office19ProPlus2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  )
for %%a in (HomeBusiness,HomeStudent) do find /i "Office16%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off16+=1
  if %vol_off16% EQU 1 if %vol_off21% EQU 0 if %vol_off19% EQU 0 find /i "Office16StandardVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off21% EQU 1 find /i "Office21Standard2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off19% EQU 1 find /i "Office19Standard2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  )
if %sub_proj% EQU 0 for %%a in (ProjectPro,ProjectStd) do find /i "Office16%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off16+=1
  if %vol_off16% EQU 1 if %vol_off21% EQU 0 if %vol_off19% EQU 0 find /i "Office16%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off21% EQU 1 find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off19% EQU 1 find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  )
if %sub_vsio% EQU 0 for %%a in (VisioPro,VisioStd) do find /i "Office16%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off16+=1
  if %vol_off16% EQU 1 if %vol_off21% EQU 0 if %vol_off19% EQU 0 find /i "Office16%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off21% EQU 1 find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off19% EQU 1 find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  )
)
if %loc_off16% EQU 1 if %ret_off16% EQU 1 if %_O16MSI% EQU 0 if defined _C16R if %prv_off16% LSS %prr_off16% (set vol_off16=0&set run_off16=1)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ApplicationID='%_oApp%' AND LicenseFamily like 'Office16O365%%' %_zz6% %_zz3% LicenseFamily %_zz4%"
if %loc_off16% EQU 1 if %run_off16% EQU 0 if %sub_o365% EQU 0 if defined _C16R %_qr% %_Nul2% | find /i "O365" %_Nul1% && (
find /i "Office16MondoVL" "!_temp!\sppchk.txt" %_Nul1% || set run_off16=1
)

set run_off15=0&set prr_off15=0&set prv_off15=0
if %loc_off15% EQU 1 if %ret_off15% EQU 1 if %_O15MSI% EQU 0 if %vol_off15% EQU 0 if defined _C15R set run_off15=1
if %loc_off15% EQU 1 if %ret_off15% EQU 1 if %_O15MSI% EQU 0 if %vol_off15% EQU 1 if defined _C15R (
for %%a in (%DO15Ids%) do find /i "Office%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off15+=1
  find /i "Office%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off15+=1
  )
for %%a in (Professional) do find /i "Office%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off15+=1
  find /i "OfficeProPlusVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off15+=1
  )
for %%a in (HomeBusiness,HomeStudent) do find /i "Office%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off15+=1
  find /i "OfficeStandardVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off15+=1
  )
if %sub_proj% EQU 0 for %%a in (ProjectPro,ProjectStd) do find /i "Office%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off15+=1
  find /i "Office%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off15+=1
  )
if %sub_vsio% EQU 0 for %%a in (VisioPro,VisioStd) do find /i "Office%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off15+=1
  find /i "Office%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off15+=1
  )
)
if %loc_off15% EQU 1 if %ret_off15% EQU 1 if %_O15MSI% EQU 0 if %vol_off15% EQU 1 if defined _C15R if %prv_off15% LSS %prr_off15% (set vol_off15=0&set run_off15=1)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ApplicationID='%_oApp%' AND LicenseFamily like 'OfficeO365%%' %_zz6% %_zz3% LicenseFamily %_zz4%"
if %loc_off15% EQU 1 if %run_off15% EQU 0 if defined _C15R %_qr% %_Nul2% | find /i "O365" %_Nul1% && (
find /i "OfficeMondoVL" "!_temp!\sppchk.txt" %_Nul1% || set run_off15=1
)

set vol_offgl=1
if %vol_off21% EQU 0 if %vol_off19% EQU 0 if %vol_off16% EQU 0 if %vol_off15% EQU 0 (
if %winbuild% GEQ 9200 set vol_offgl=0
if %winbuild% LSS 9200 if %vol_off14% EQU 0 set vol_offgl=0
)
rem 混合批量版本 + 零售版本方案
if %run_off21% EQU 1 if %AutoR2V% EQU 1 if %RanR2V% EQU 0 goto :C2RR2V
if %run_off19% EQU 1 if %AutoR2V% EQU 1 if %RanR2V% EQU 0 goto :C2RR2V
if %run_off16% EQU 1 if %AutoR2V% EQU 1 if %RanR2V% EQU 0 goto :C2RR2V
if %run_off15% EQU 1 if %AutoR2V% EQU 1 if %RanR2V% EQU 0 goto :C2RR2V
rem 全支持批量方案 + 不支持的消息
if %loc_off16% EQU 0 if %ret_off16% EQU 1 if %_O16MSI% EQU 0 if %OffUWP% EQU 1 (echo.&echo %_mOuwp%)
if %vol_offgl% EQU 1 (
if %ret_off16% EQU 1 if %_O16MSI% EQU 1 (echo.&echo %_mO16m%)
if %ret_off15% EQU 1 if %_O15MSI% EQU 1 (echo.&echo %_mO15m%)
if %winbuild% LSS 9200 if %loc_off14% EQU 1 if %vol_off14% EQU 0 (if defined _C14R (echo.&echo %_mO14c%) else if %_O14MSI% EQU 1 (if %ret_off14% EQU 1 echo.&echo %_mO14m%))
exit /b
)
set Off1ce=0
rem 零售 C2R 方案
if %AutoR2V% EQU 1 if %RanR2V% EQU 0 goto :C2RR2V
:ReturnSPP
rem 零售 MSI/C2R 方案或失败的 C2R-R2V 方案
if %loc_off21% EQU 1 if %vol_off21% EQU 0 (
if %aC2R21% EQU 1 (echo.&echo %_mO21a%) else (echo.&echo %_mO21c%)
)
if %loc_off19% EQU 1 if %vol_off19% EQU 0 (
if %aC2R19% EQU 1 (echo.&echo %_mO19a%) else (echo.&echo %_mO19c%)
)
if %loc_off16% EQU 1 if %vol_off16% EQU 0 (
if defined _C16R (if %aC2R16% EQU 1 (echo.&echo %_mO16a%) else (if %sub_o365% EQU 0 echo.&echo %_mO16c%)) else if %_O16MSI% EQU 1 (if %ret_off16% EQU 1 echo.&echo %_mO16m%)
)
if %loc_off15% EQU 1 if %vol_off15% EQU 0 (
if defined _C15R (if %aC2R15% EQU 1 (echo.&echo %_mO15a%) else (echo.&echo %_mO15c%)) else if %_O15MSI% EQU 1 (if %ret_off15% EQU 1 echo.&echo %_mO15m%)
)
if %winbuild% LSS 9200 if %loc_off14% EQU 1 if %vol_off14% EQU 0 (
if defined _C14R (echo.&echo %_mO14c%) else if %_O14MSI% EQU 1 (if %ret_off14% EQU 1 echo.&echo %_mO14m%)
)
exit /b

:sppchkoff
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='%app%'%_zz6% %_zz3% Name %_zz4%"
%_qr% > "!_temp!\sppchk.txt"
if %winbuild% LSS 9200 find /i "Office 14" "!_temp!\sppchk.txt" %_Nul1% && (if %loc_off14% EQU 0 exit /b)
find /i "Office 15" "!_temp!\sppchk.txt" %_Nul1% && (if %loc_off15% EQU 0 exit /b)
find /i "Office 16" "!_temp!\sppchk.txt" %_Nul1% && (if %loc_off16% EQU 0 exit /b)
find /i "Office 19" "!_temp!\sppchk.txt" %_Nul1% && (if %loc_off19% EQU 0 exit /b)
find /i "Office 21" "!_temp!\sppchk.txt" %_Nul1% && (if %loc_off21% EQU 0 exit /b)
if %1 EQU 1 (set _officespp=1) else (set _officespp=0)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%PartialProductKey is not NULL%_zz6% %_zz3% ID %_zz4%"
%_qr% %_Nul2% | findstr /i "%app%" %_Nul1% && (echo.&call :activate&exit /b)
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ID='%app%'%_zz6% %_zz3% Name %_zz8%"
for /f "tokens=3 delims==, " %%G in ('%_qr%') do set OffVer=%%G
call :offchk%OffVer%
exit /b

:sppchkwin
set _officespp=0
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL%_zz6% %_zz3% Name %_zz4%"
if %winbuild% GEQ 14393 if %WinPerm% EQU 0 if %_gvlk% EQU 0 %_qr% %_Nul2% | findstr /i Windows %_Nul1% && (set _gvlk=1)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='%app%'%_zz6% %_zz3% LicenseStatus %_zz4%"
%_qr% %_Nul2% | findstr "1" %_Nul1% && (echo.&call :activate&exit /b)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%PartialProductKey is not NULL%_zz6% %_zz3% ID %_zz4%"
%_qr% %_Nul2% | findstr /i "%app%" %_Nul1% && (echo.&call :activate&exit /b)
if %winbuild% GEQ 14393 if %_gvlk% EQU 1 exit /b
if %WinPerm% EQU 1 exit /b
if %winbuild% LSS 10240 (call :winchk&exit /b)
for %%A in (
b71515d9-89a2-4c60-88c8-656fbcca7f3a,af43f7f0-3b1e-4266-a123-1fdb53f4323b,075aca1f-05d7-42e5-a3ce-e349e7be7078
11a37f09-fb7f-4002-bd84-f3ae71d11e90,43f2ab05-7c87-4d56-b27c-44d0f9a3dabd,2cf5af84-abab-4ff0-83f8-f040fb2576eb
6ae51eeb-c268-4a21-9aae-df74c38b586d,ff808201-fec6-4fd4-ae16-abbddade5706,34260150-69ac-49a3-8a0d-4a403ab55763
4dfd543d-caa6-4f69-a95f-5ddfe2b89567,5fe40dd6-cf1f-4cf2-8729-92121ac2e997,903663f7-d2ab-49c9-8942-14aa9e0a9c72
2cc171ef-db48-4adc-af09-7c574b37f139,5b2add49-b8f4-42e0-a77c-adad4efeeeb1
) do (
if /i '%app%' EQU '%%A' exit /b
)
if not defined EditionID (call :winchk&exit /b)
if %winbuild% LSS 14393 (call :winchk&exit /b)
if /i '%app%' EQU '32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee' if /i %EditionID% NEQ EnterpriseS exit /b
if /i '%app%' EQU 'ca7df2e3-5ea0-47b8-9ac1-b1be4d8edd69' if /i %EditionID% NEQ CloudEdition exit /b
if /i '%app%' EQU 'd30136fc-cb4b-416e-a23d-87207abc44a9' if /i %EditionID% NEQ CloudEditionN exit /b
if /i '%app%' EQU '0df4f814-3f57-4b8b-9a9d-fddadcd69fac' if /i %EditionID% NEQ CloudE exit /b
if /i '%app%' EQU 'e0c42288-980c-4788-a014-c080d2e1926e' if /i %EditionID% NEQ Education exit /b
if /i '%app%' EQU '73111121-5638-40f6-bc11-f1d7b0d64300' if /i %EditionID% NEQ Enterprise exit /b
if /i '%app%' EQU '2de67392-b7a7-462a-b1ca-108dd189f588' if /i %EditionID% NEQ Professional exit /b
if /i '%app%' EQU '3f1afc82-f8ac-4f6c-8005-1d233e606eee' if /i %EditionID% NEQ ProfessionalEducation exit /b
if /i '%app%' EQU '82bbc092-bc50-4e16-8e18-b74fc486aec3' if /i %EditionID% NEQ ProfessionalWorkstation exit /b
if /i '%app%' EQU '3c102355-d027-42c6-ad23-2e7ef8a02585' if /i %EditionID% NEQ EducationN exit /b
if /i '%app%' EQU 'e272e3e2-732f-4c65-a8f0-484747d0d947' if /i %EditionID% NEQ EnterpriseN exit /b
if /i '%app%' EQU 'a80b5abf-76ad-428b-b05d-a47d2dffeebf' if /i %EditionID% NEQ ProfessionalN exit /b
if /i '%app%' EQU '5300b18c-2e33-4dc2-8291-47ffcec746dd' if /i %EditionID% NEQ ProfessionalEducationN exit /b
if /i '%app%' EQU '4b1571d3-bafb-4b40-8087-a961be2caf65' if /i %EditionID% NEQ ProfessionalWorkstationN exit /b
if /i '%app%' EQU '58e97c99-f377-4ef1-81d5-4ad5522b5fd8' if /i %EditionID% NEQ Core exit /b
if /i '%app%' EQU 'cd918a57-a41b-4c82-8dce-1a538e221a83' if /i %EditionID% NEQ CoreSingleLanguage exit /b
if /i '%app%' EQU 'ec868e65-fadf-4759-b23e-93fe37f2cc29' if /i %EditionID% NEQ ServerRdsh exit /b
if /i '%app%' EQU 'e4db50ea-bda1-4566-b047-0ca50abc6f07' if /i %EditionID% NEQ ServerRdsh exit /b
set "_qr=%_zz1% %spp% %_zz2% "Description like '%%KMSCLIENT%%'" %_zz3% ID %_zz4%"
if /i "%app%" EQU "e4db50ea-bda1-4566-b047-0ca50abc6f07" (
%_qr% | findstr /i "ec868e65-fadf-4759-b23e-93fe37f2cc29" %_Nul3% && (exit /b)
)
call :winchk
exit /b

:winchk
if not defined tok (if %winbuild% GEQ 9200 (set "tok=4") else (set "tok=7"))
set "_qr=%_zz1% %spp% %_zz2% %_zz5%LicenseStatus='1' and Description like '%%KMSCLIENT%%' %adoff% %_zz6% %_zz3% Name %_zz4%"
%_qr% %_Nul2% | findstr /i "Windows" %_Nul3% && (exit /b)
echo.
set "_qr=%_zz1% %spp% %_zz2% %_zz5%LicenseStatus='1' and GracePeriodRemaining='0' %adoff% and PartialProductKey is not NULL%_zz6% %_zz3% Name %_zz4%"
%_qr% %_Nul2% | findstr /i "Windows" %_Nul3% && (
set WinPerm=1
)
set WinOEM=0
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and LicenseStatus='1' %adoff% %_zz6% %_zz3% Name %_zz4%"
if %WinPerm% EQU 0 %_qr% %_Nul2% | findstr /i "Windows" %_Nul3% && set WinOEM=1
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and LicenseStatus='1' %adoff% %_zz6% %_zz3% Description %_zz8%"
if %WinOEM% EQU 1 (
for /f "tokens=%tok% delims=, " %%G in ('%_qr%') do set "channel=%%G"
for %%A in (VOLUME_MAK, RETAIL, OEM_DM, OEM_SLP, OEM_COA, OEM_COA_SLP, OEM_COA_NSLP, OEM_NONSLP, OEM) do if /i "%%A"=="!channel!" set WinPerm=1
)
if %WinPerm% EQU 0 (
copy /y %SysPath%\slmgr.vbs "!_temp!\slmgr.vbs" %_Nul3%
cscript //nologo "!_temp!\slmgr.vbs" /xpr %_Nul2% | findstr /i "permanently" %_Nul3% && set WinPerm=1
)
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ApplicationID='%_wApp%' and LicenseStatus='1' %adoff% %_zz6% %_zz3% Name %_zz8%"
if %WinPerm% EQU 1 (
for /f "tokens=2 delims==" %%x in ('%_qr%') do echo 正在检查：%%x
echo 产品已永久激活。
exit /b
)
call :insKey
exit /b

:esuchk
set _officespp=0
set ESU_ADD=1
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='%app%'%_zz6% %_zz3% LicenseStatus %_zz4%"
%_qr% %_Nul2% | findstr "1" %_Nul1% && (echo.&call :activate&exit /b)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='77db037b-95c3-48d7-a3ab-a9c6d41093e0'%_zz6% %_zz3% LicenseStatus %_zz4%"
if /i "%app%" EQU "3fcc2df2-f625-428d-909a-1f76efc849b6" (
%_qr% %_Nul2% | findstr "1" %_Nul1% && (exit /b)
)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='0e00c25d-8795-4fb7-9572-3803d91b6880'%_zz6% %_zz3% LicenseStatus %_zz4%"
if /i "%app%" EQU "dadfcd24-6e37-47be-8f7f-4ceda614cece" (
%_qr% %_Nul2% | findstr "1" %_Nul1% && (exit /b)
)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='4220f546-f522-46df-8202-4d07afd26454'%_zz6% %_zz3% LicenseStatus %_zz4%"
if /i "%app%" EQU "0c29c85e-12d7-4af8-8e4d-ca1e424c480c" (
%_qr% %_Nul2% | findstr "1" %_Nul1% && (exit /b)
)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='553673ed-6ddf-419c-a153-b760283472fd'%_zz6% %_zz3% LicenseStatus %_zz4%"
if /i "%app%" EQU "f2b21bfc-a6b0-4413-b4bb-9f06b55f2812" (
%_qr% %_Nul2% | findstr "1" %_Nul1% && (exit /b)
)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='04fa0286-fa74-401e-bbe9-fbfbb158010d'%_zz6% %_zz3% LicenseStatus %_zz4%"
if /i "%app%" EQU "bfc078d0-8c7f-475c-8519-accc46773113" (
%_qr% %_Nul2% | findstr "1" %_Nul1% && (exit /b)
)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='16c08c85-0c8b-4009-9b2b-f1f7319e45f9'%_zz6% %_zz3% LicenseStatus %_zz4%"
if /i "%app%" EQU "23c6188f-c9d8-457e-81b6-adb6dacb8779" (
%_qr% %_Nul2% | findstr "1" %_Nul1% && (exit /b)
)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%ID='8e7bfb1e-acc1-4f56-abae-b80fce56cd4b'%_zz6% %_zz3% LicenseStatus %_zz4%"
if /i "%app%" EQU "e7cce015-33d6-41c1-9831-022ba63fe1da" (
%_qr% %_Nul2% | findstr "1" %_Nul1% && (exit /b)
)
set "_qr=%_zz1% %spp% %_zz2% %_zz5%PartialProductKey is not NULL%_zz6% %_zz3% ID %_zz4%"
%_qr% %_Nul2% | findstr /i "%app%" %_Nul1% && (echo.&call :activate&exit /b)
call :insKey
exit /b

:RunOSPP
set spp=OfficeSoftwareProtectionProduct
set sps=OfficeSoftwareProtectionService
set Off1ce=0
set RanR2V=0
set aC2R21=0
set aC2R19=0
set aC2R16=0
set aC2R15=0
if %winbuild% LSS 9200 (set "aword=2010-2021") else (set "aword=2010")
if %OsppHook% EQU 0 (echo.&echo 未检测到已安装的 Office %aword% 产品……&exit /b)
if %winbuild% GEQ 9200 if %loc_off14% EQU 0 (echo.&echo 未检测到已安装的 Office %aword% 产品……&exit /b)
set err_offsvc=0
net start osppsvc /y %_Nul3% || (
sc start osppsvc %_Nul3%
if !errorlevel! EQU 1053 set err_offsvc=1
)
if %err_offsvc% EQU 1 (echo.&echo 错误：osppsvc 服务未运行……&exit /b)
if %winbuild% GEQ 9200 call :oppoff
if %winbuild% LSS 9200 call :sppoff
if %Off1ce% EQU 0 exit /b
set "vPrem="&set "vProf="
set "_qr=%_zz7% %spp% %_zz2% %_zz5%LicenseFamily='OfficeVisioPrem-MAK'%_zz6% %_zz3% LicenseStatus %_zz8%"
if %loc_off14% EQU 1 for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set vPrem=%%A
set "_qr=%_zz7% %spp% %_zz2% %_zz5%LicenseFamily='OfficeVisioPro-MAK'%_zz6% %_zz3% LicenseStatus %_zz8%"
if %loc_off14% EQU 1 for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set vProf=%%A
set "_qr=%_zz7% %sps% %_zz3% Version %_zz8%"
for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set slsv=%%A
reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" %_Nul3%
reg add "HKLM\%OPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" %_Nul3%
set "_qr=%_zz7% %spp% %_zz2% %_zz5%Description like '%%KMSCLIENT%%' %_zz6% %_zz3% ID %_zz8%"
for /f "tokens=2 delims==" %%G in ('%_qr%') do (set app=%%G&call :sppchkoff 2)
reg delete "HKLM\%OPPk%" /f /v DisableDnsPublishing %_Null%
reg delete "HKLM\%OPPk%" /f /v DisableKeyManagementServiceHostCaching %_Null%
exit /b

:oppoff
set "_qr=%_zz1% %spp% %_zz3% Description %_zz4%"
%_qr% %_Nul2% | findstr /i KMSCLIENT %_Nul1% && (
set Off1ce=1
exit /b
)
set ret_off14=0
%_qr% %_Nul2% | findstr /i channel %_Nul1% && (set ret_off14=1)
if defined _C14R (echo.&echo %_mO14c%) else if %_O14MSI% EQU 1 (if %ret_off14% EQU 1 echo.&echo %_mO14m%)
exit /b

:offchk
set ls=0
set ls2=0
set ls3=0
set "_qr=%_zz7% %spp% %_zz2% %_zz5%LicenseFamily='Office%~1'%_zz6% %_zz3% LicenseStatus %_zz8%"
for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set /a ls=%%A
set "_qr=%_zz7% %spp% %_zz2% %_zz5%LicenseFamily='Office%~3'%_zz6% %_zz3% LicenseStatus %_zz8%"
if /i not "%~3"=="" for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set /a ls2=%%A
set "_qr=%_zz7% %spp% %_zz2% %_zz5%LicenseFamily='Office%~5'%_zz6% %_zz3% LicenseStatus %_zz8%"
if /i not "%~5"=="" for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set /a ls3=%%A
if "%ls3%"=="1" (
echo 正在检查：%~6
echo 产品已永久激活。
exit /b
)
if "%ls2%"=="1" (
echo 正在检查：%~4
echo 产品已永久激活。
exit /b
)
if "%ls%"=="1" (
echo 正在检查：%~2
echo 产品已永久激活。
exit /b
)
call :insKey
exit /b

:offchk21
if /i '%app%' EQU 'f3fb2d68-83dd-4c8b-8f09-08e0d950ac3b' exit /b
if /i '%app%' EQU '76093b1b-7057-49d7-b970-638ebcbfd873' exit /b
if /i '%app%' EQU 'a3b44174-2451-4cd6-b25f-66638bfb9046' exit /b
if /i '%app%' EQU 'fbdb3e18-a8ef-4fb3-9183-dffd60bd0984' (
call :offchk "21ProPlus2021VL_MAK_AE1" "Office ProPlus 2021" "21ProPlus2021VL_MAK_AE2"
exit /b
)
if /i '%app%' EQU '080a45c5-9f9f-49eb-b4b0-c3c610a5ebd3' (
call :offchk "21Standard2021VL_MAK_AE" "Office Standard 2021"
exit /b
)
if /i '%app%' EQU '76881159-155c-43e0-9db7-2d70a9a3a4ca' (
call :offchk "21ProjectPro2021VL_MAK_AE1" "Project Pro 2021" "21ProjectPro2021VL_MAK_AE2"
exit /b
)
if /i '%app%' EQU '6dd72704-f752-4b71-94c7-11cec6bfc355' (
call :offchk "21ProjectStd2021VL_MAK_AE" "Project Standard 2021"
exit /b
)
if /i '%app%' EQU 'fb61ac9a-1688-45d2-8f6b-0674dbffa33c' (
call :offchk "21VisioPro2021VL_MAK_AE" "Visio Pro 2021"
exit /b
)
if /i '%app%' EQU '72fce797-1884-48dd-a860-b2f6a5efd3ca' (
call :offchk "21VisioStd2021VL_MAK_AE" "Visio Standard 2021"
exit /b
)
call :insKey
exit /b

:offchk19
if /i '%app%' EQU '0bc88885-718c-491d-921f-6f214349e79c' exit /b
if /i '%app%' EQU 'fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9' exit /b
if /i '%app%' EQU '500f6619-ef93-4b75-bcb4-82819998a3ca' exit /b
if /i '%app%' EQU '85dd8b5f-eaa4-4af3-a628-cce9e77c9a03' (
call :offchk "19ProPlus2019VL_MAK_AE" "Office ProPlus 2019"
exit /b
)
if /i '%app%' EQU '6912a74b-a5fb-401a-bfdb-2e3ab46f4b02' (
call :offchk "19Standard2019VL_MAK_AE" "Office Standard 2019"
exit /b
)
if /i '%app%' EQU '2ca2bf3f-949e-446a-82c7-e25a15ec78c4' (
call :offchk "19ProjectPro2019VL_MAK_AE" "Project Pro 2019"
exit /b
)
if /i '%app%' EQU '1777f0e3-7392-4198-97ea-8ae4de6f6381' (
call :offchk "19ProjectStd2019VL_MAK_AE" "Project Standard 2019"
exit /b
)
if /i '%app%' EQU '5b5cf08f-b81a-431d-b080-3450d8620565' (
call :offchk "19VisioPro2019VL_MAK_AE" "Visio Pro 2019"
exit /b
)
if /i '%app%' EQU 'e06d7df3-aad0-419d-8dfb-0ac37e2bdf39' (
call :offchk "19VisioStd2019VL_MAK_AE" "Visio Standard 2019"
exit /b
)
call :insKey
exit /b

:offchk16
if /i '%app%' EQU 'd450596f-894d-49e0-966a-fd39ed4c4c64' (
call :offchk "16ProPlusVL_MAK" "Office ProPlus 2016"
exit /b
)
if /i '%app%' EQU 'dedfa23d-6ed1-45a6-85dc-63cae0546de6' (
call :offchk "16StandardVL_MAK" "Office Standard 2016"
exit /b
)
if /i '%app%' EQU '4f414197-0fc2-4c01-b68a-86cbb9ac254c' (
call :offchk "16ProjectProVL_MAK" "Project Pro 2016"
exit /b
)
if /i '%app%' EQU 'da7ddabc-3fbe-4447-9e01-6ab7440b4cd4' (
call :offchk "16ProjectStdVL_MAK" "Project Standard 2016"
exit /b
)
if /i '%app%' EQU '6bf301c1-b94a-43e9-ba31-d494598c47fb' (
call :offchk "16VisioProVL_MAK" "Visio Pro 2016"
exit /b
)
if /i '%app%' EQU 'aa2a7821-1827-4c2c-8f1d-4513a34dda97' (
call :offchk "16VisioStdVL_MAK" "Visio Standard 2016"
exit /b
)
if /i '%app%' EQU '829b8110-0e6f-4349-bca4-42803577788d' (
call :offchk "16ProjectProXC2RVL_MAKC2R" "Project Pro 2016 C2R"
exit /b
)
if /i '%app%' EQU 'cbbaca45-556a-4416-ad03-bda598eaa7c8' (
call :offchk "16ProjectStdXC2RVL_MAKC2R" "Project Standard 2016 C2R"
exit /b
)
if /i '%app%' EQU 'b234abe3-0857-4f9c-b05a-4dc314f85557' (
call :offchk "16VisioProXC2RVL_MAKC2R" "Visio Pro 2016 C2R"
exit /b
)
if /i '%app%' EQU '361fe620-64f4-41b5-ba77-84f8e079b1f7' (
call :offchk "16VisioStdXC2RVL_MAKC2R" "Visio Standard 2016 C2R"
exit /b
)
call :insKey
exit /b

:offchk15
if /i '%app%' EQU 'b322da9c-a2e2-4058-9e4e-f59a6970bd69' (
call :offchk "ProPlusVL_MAK" "Office ProPlus 2013"
exit /b
)
if /i '%app%' EQU 'b13afb38-cd79-4ae5-9f7f-eed058d750ca' (
call :offchk "StandardVL_MAK" "Office Standard 2013"
exit /b
)
if /i '%app%' EQU '4a5d124a-e620-44ba-b6ff-658961b33b9a' (
call :offchk "ProjectProVL_MAK" "Project Pro 2013"
exit /b
)
if /i '%app%' EQU '427a28d1-d17c-4abf-b717-32c780ba6f07' (
call :offchk "ProjectStdVL_MAK" "Project Standard 2013"
exit /b
)
if /i '%app%' EQU 'e13ac10e-75d0-4aff-a0cd-764982cf541c' (
call :offchk "VisioProVL_MAK" "Visio Pro 2013"
exit /b
)
if /i '%app%' EQU 'ac4efaf0-f81f-4f61-bdf7-ea32b02ab117' (
call :offchk "VisioStdVL_MAK" "Visio Standard 2013"
exit /b
)
call :insKey
exit /b

:offchk14
if /i '%app%' EQU '6f327760-8c5c-417c-9b61-836a98287e0c' (
call :offchk "ProPlus-MAK" "Office ProPlus 2010" "ProPlusAcad-MAK" "Office Professional Academic 2010"
exit /b
)
if /i '%app%' EQU '9da2a678-fb6b-4e67-ab84-60dd6a9c819a' (
call :offchk "Standard-MAK" "Office Standard 2010" "StandardAcad-MAK"  "Office Standard Academic 2010"
exit /b
)
if /i '%app%' EQU 'ea509e87-07a1-4a45-9edc-eba5a39f36af' (
call :offchk "SmallBusBasics-MAK" "Office Small Business Basics 2010"
exit /b
)
if /i '%app%' EQU 'df133ff7-bf14-4f95-afe3-7b48e7e331ef' (
call :offchk "ProjectPro-MAK" "Project Pro 2010"
exit /b
)
if /i '%app%' EQU '5dc7bf61-5ec9-4996-9ccb-df806a2d0efe' (
call :offchk "ProjectStd-MAK" "Project Standard 2010" "ProjectStd-MAK2" "Project Standard 2010"
exit /b
)
if /i '%app%' EQU '92236105-bb67-494f-94c7-7f7a607929bd' (
call :offchk "VisioPrem-MAK" "Visio Premium 2010" "VisioPro-MAK" "Visio Pro 2010"
exit /b
)
if defined vPrem exit /b
if /i '%app%' EQU 'e558389c-83c3-4b29-adfe-5e4d7f46c358' (
call :offchk "VisioPro-MAK" "Visio Pro 2010" "VisioStd-MAK" "Visio Standard 2010"
exit /b
)
if defined vProf exit /b
if /i '%app%' EQU '9ed833ff-4f92-4f36-b370-8683a4f13275' (
call :offchk "VisioStd-MAK" "Visio Standard 2010"
exit /b
)
call :insKey
exit /b

:officeLoc
set loc_off%1=0
set _O%1MSI=0
if %1 EQU 19 (
if defined _C16R reg query %_C16R% /v ProductReleaseIds %_Nul2% | findstr 2019 %_Nul1% && set loc_off%1=1
exit /b
)
if %1 EQU 21 (
if defined _C16R reg query %_C16R% /v ProductReleaseIds %_Nul2% | findstr 2021 %_Nul1% && set loc_off%1=1
exit /b
)

for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\%1.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" (
set loc_off%1=1
set _O%1MSI=1
)
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\%1.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" (
set loc_off%1=1
set _O%1MSI=1
)

if %1 EQU 16 if defined _C16R (
for /f "skip=2 tokens=2*" %%a in ('reg query %_C16R% /v ProductReleaseIds') do echo %%b> "!_temp!\c2rchk.txt"
for %%a in (%LV16Ids%,ProjectProX,ProjectStdX,VisioProX,VisioStdX) do (
  findstr /I /C:"%%aVolume" "!_temp!\c2rchk.txt" %_Nul1% && set loc_off%1=1
  )
for %%a in (%LR16Ids%) do (
  findstr /I /C:"%%aRetail" "!_temp!\c2rchk.txt" %_Nul1% && set loc_off%1=1
  )
exit /b
)

if %1 EQU 15 if defined _C15R (
set loc_off%1=1
exit /b
)

if exist "%ProgramFiles%\Microsoft Office\Office%1\OSPP.VBS" set loc_off%1=1
if not %xOS%==x86 if exist "%ProgramW6432%\Microsoft Office\Office%1\OSPP.VBS" set loc_off%1=1
if not %xOS%==x86 if exist "%ProgramFiles(x86)%\Microsoft Office\Office%1\OSPP.VBS" set loc_off%1=1
exit /b

:officeSub
reg query %kNext% | findstr /i /r ".*retail" %_Nul2% | findstr /i /v "project visio" %_Nul2% | find /i "0x2" %_Nul1% && (set sub_o365=1)
reg query %kNext% | findstr /i /r ".*retail" %_Nul2% | findstr /i /v "project visio" %_Nul2% | find /i "0x3" %_Nul1% && (set sub_o365=1)
reg query %kNext% | findstr /i /r ".*volume" %_Nul2% | findstr /i /v "project visio" %_Nul2% | find /i "0x2" %_Nul1% && (set sub_o365=1)
reg query %kNext% | findstr /i /r ".*volume" %_Nul2% | findstr /i /v "project visio" %_Nul2% | find /i "0x3" %_Nul1% && (set sub_o365=1)
reg query %kNext% | findstr /i /r "project.*" %_Nul2% | find /i "0x2" %_Nul1% && set sub_proj=1
reg query %kNext% | findstr /i /r "project.*" %_Nul2% | find /i "0x3" %_Nul1% && set sub_proj=1
reg query %kNext% | findstr /i /r "visio.*" %_Nul2% | find /i "0x2" %_Nul1% && set sub_vsio=1
reg query %kNext% | findstr /i /r "visio.*" %_Nul2% | find /i "0x3" %_Nul1% && set sub_vsio=1
if %sub_o365% EQU 1 set sub_next=1
if %sub_proj% EQU 1 set sub_next=1
if %sub_vsio% EQU 1 set sub_next=1
exit /b

:insKey
set S_OK=1
echo.
set "_key="
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ID='%app%'%_zz6% %_zz3% Name %_zz8%"
if %ESU_ADD% EQU 0 for /f "tokens=2 delims==" %%x in ('%_qr%') do echo 正在安装密钥：%%x
if %ESU_ADD% EQU 1 for /f "tokens=2 delims==f" %%x in ('%_qr%') do echo 正在安装密钥：%%x
set ESU_ADD=0
call :keys %app%
if "%_key%"=="" (echo 未找到关联的 KMS 客户端密钥&exit /b)
set "_qr=wmic path %sps% where Version='%slsv%' call InstallProductKey ProductKey="%_key%""
if %WMI_VBS% NEQ 0 set "_qr=%_csp% %sps% "%_key%""
%_qr% %_Nul3%
set ERRORCODE=%ERRORLEVEL%
if %ERRORCODE% NEQ 0 (
cmd /c exit /b %ERRORCODE%
echo 失败：0x!=ExitCode!
set S_OK=0
exit /b
)
set "_qr=wmic path %sps% where Version='%slsv%' call RefreshLicenseStatus"
if %WMI_VBS% NEQ 0 set "_qr=%_csm% "%sps%.Version='%slsv%'" RefreshLicenseStatus"
if %sps% EQU SoftwareLicensingService %_qr% %_Nul3%

:activate
set S_OK=1
if %sps% EQU SoftwareLicensingService (
if %_officespp% EQU 0 (reg delete "HKLM\%SPPk%\%_wApp%\%app%" /f %_Null%) else (reg delete "HKLM\%SPPk%\%_oApp%\%app%" /f %_Null%)
) else (
reg delete "HKLM\%OPPk%\%_oA14%\%app%" /f %_Null%
reg delete "HKLM\%OPPk%\%_oApp%\%app%" /f %_Null%
)
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ID='%app%'%_zz6% %_zz3% Name %_zz8%"
if %W1nd0ws% EQU 0 if %_officespp% EQU 0 if %sps% EQU SoftwareLicensingService (
reg add "HKLM\%SPPk%\%_wApp%\%app%" /f /v KeyManagementServiceName /t REG_SZ /d "127.0.0.2" %_Nul3%
reg add "HKLM\%SPPk%\%_wApp%\%app%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" %_Nul3%
for /f "tokens=2 delims==" %%x in ('%_qr%') do echo 正在检查：%%x
echo 产品已通过 KMS 2038 激活。
set _keepkms38=1
exit /b
)
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ID='%app%'%_zz6% %_zz3% Name %_zz8%"
if %act_attempt% LSS 1 (
if %ESU_ADD% EQU 0 for /f "tokens=2 delims==" %%x in ('%_qr%') do echo 正在激活：%%x
if %ESU_ADD% EQU 1 for /f "tokens=2 delims==f" %%x in ('%_qr%') do echo 正在激活：%%x
)

set ESU_ADD=0
set "_qr=wmic path %spp% where ID='%app%' call Activate"
if %WMI_VBS% NEQ 0 set "_qr=%_csm% "%spp%.ID='%app%'" Activate"
%_qr% %_Nul3%
call set ERRORCODE=%ERRORLEVEL%
if %act_attempt% LSS 1 if %ERRORCODE% EQU -1073418187 (
echo 产品激活失败：0xC004F035
if %OSType% EQU Win7 echo 由于使用不合格的 OEM BIOS，此电脑运行的 Windows 7 无法通过 KMS 激活。
echo 请参阅自述文件了解详细信息。
exit /b
)
if %act_attempt% LSS 1 if %ERRORCODE% EQU -1073417728 (
echo 产品激活失败：0xC004F200
echo Windows 需要重建激活相关文件。
echo 请参阅 KB2736303 了解详细信息。
exit /b
)
if %act_attempt% LSS 1 if %ERRORCODE% EQU -1073422315 (
echo 产品激活失败：0xC004E015
echo 正在运行 slmgr.vbs /rilc 缓解。
cscript //Nologo //B %SysPath%\slmgr.vbs /rilc
)
set gpr=0
set gpr2=0
set "_qr=%_zz7% %spp% %_zz2% %_zz5%ID='%app%'%_zz6% %_zz3% GracePeriodRemaining %_zz8%"
for /f "tokens=2 delims==" %%x in ('%_qr%') do (set gpr=%%x&set /a "gpr2=(%%x+1440-1)/1440")
if %act_attempt% LSS 1 if %ERRORCODE% EQU 0 if %gpr% EQU 0 (
echo 产品激活成功，但剩余期限增加失败。
if %OSType% EQU Win7 echo 这可能与 KB4487266 中描述的错误有关
exit /b
)
set Act_OK=0
if %gpr% EQU 43200 if %_officespp% EQU 0 if %winbuild% GEQ 9200 set Act_OK=1
if %gpr% EQU 64800 set Act_OK=1
if %gpr% GTR 259200 if %Win10Gov% EQU 1 set Act_OK=1
if %gpr% EQU 259200 set Act_OK=1

if %ERRORCODE% EQU 0 if %Act_OK% EQU 1 (
call :_color %_Green% "产品激活成功"
echo 剩余期限：%gpr2% 天（%gpr% 分钟）
set /a act_attempt=0
exit /b
)

if not !server_num! gtr %max_servers% (
if %act_attempt% LSS 3 (
set /a act_attempt+=1
call :getserv
%nul% reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "!KMS_IP!"
%nul% reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "!KMS_IP!"
if %winbuild% GEQ 9200 (
%nul% reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "!KMS_IP!"
if defined notx86 (
%nul% reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "!KMS_IP!" /reg:32
%nul% reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "!KMS_IP!" /reg:32
)
)
goto :activate
)
)

cmd /c exit /b %ERRORCODE%
if %ERRORCODE% NEQ 0 (
call :_color %_Red% "产品激活失败：0x!=ExitCode!"
) else (
call :_color %_Red% "产品激活失败"
)
echo 剩余期限：%gpr2% 天（%gpr% 分钟）
set S_OK=0
set act_failed=1
set /a act_attempt=0
exit /b

:StopService
sc query %1 | find /i "STOPPED" %_Nul1% || net stop %1 /y %_Nul3%
sc query %1 | find /i "STOPPED" %_Nul1% || sc stop %1 %_Nul3%
goto :eof

:UpdateOSPPEntry
if /i %1 EQU osppsvc.exe (
reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "!KMS_IP!" %_Nul3%
reg add "HKLM\%OPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" %_Nul3%
)
goto :eof

:CheckFR

set WMIe=0
call :CheckWS
if %WMIe% EQU 1 (
echo.
echo %_err%
echo 运行 WMI 查询检查失败。
)
goto :eof

:CheckWS
set "_qrw=%_zz1% Win32_ComputerSystem %_zz3% CreationClassName %_zz4%"
set "_qrs=%_zz1% SoftwareLicensingService %_zz3% Version %_zz4%"

%_qrs% %_Nul2% | findstr /r "[0-9]*\.[0-9]*\.[0-9]*\.[0-9]*" %_Nul1% || (
  set WMIe=1
  %_qrw% %_Nul2% | find /i "ComputerSystem" %_Nul1% && (
    echo 错误：SPP 没有响应
    ) || (
    echo 错误：WMI 和 SPP 没有响应
  )
)
goto :eof

:C2RR2V
set RanR2V=1
set "_SLMGR=%SysPath%\slmgr.vbs"
if %_Debug% EQU 0 (
set "_cscript=cscript //Nologo //B"
) else (
set "_cscript=cscript //Nologo"
)
set _LTSC=0
set "_tag="&set "_ons= 2016"
sc query ClickToRunSvc %_Nul3%
set error1=%errorlevel%
sc query OfficeSvc %_Nul3%
set error2=%errorlevel%
if %error1% EQU 1060 if %error2% EQU 1060 (
echo 错误：没有检测到 Office C2R 安装路径
goto :%_fC2R%
)
set _Office16=0
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (
  set _Office16=1
)
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (
  set _Office16=1
)
set _Office15=0
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (
  set _Office15=1
)
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (
  set _Office15=1
)
if %_Office16% EQU 0 if %_Office15% EQU 0 (
echo 错误：没有检测到 Office C2R 安装路径
goto :%_fC2R%
)

:Reg16istry
if %_Office16% EQU 0 goto :Reg15istry
set "_InstallRoot="
set "_ProductIds="
set "_GUID="
set "_Config="
set "_PRIDs="
set "_LicensesPath="
set "_Integrator="
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do (set "_InstallRoot=%%b\root")
if not "%_InstallRoot%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do (set "_OSPPVBS=%%b\Office16\OSPP.VBS")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v PackageGUID" %_Nul6%') do (set "_GUID=%%b")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do (set "_ProductIds=%%b")
  set "_Config=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
  set "_PRIDs=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs"
) else (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do (set "_InstallRoot=%%b\root")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do (set "_OSPPVBS=%%b\Office16\OSPP.VBS")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v PackageGUID" %_Nul6%') do (set "_GUID=%%b")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do (set "_ProductIds=%%b")
  set "_Config=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
  set "_PRIDs=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\ProductReleaseIDs"
)
set "_LicensesPath=%_InstallRoot%\Licenses16"
set "_Integrator=%_InstallRoot%\integration\integrator.exe"
for /f "skip=2 tokens=2*" %%a in ('"reg query %_PRIDs% /v ActiveConfiguration" %_Nul6%') do set "_PRIDs=%_PRIDs%\%%b"
if "%_ProductIds%"=="" (
if %_Office15% EQU 0 (echo 错误：没有检测到 Office C2R 产品 ID&goto :%_fC2R%) else (goto :Reg15istry)
)
if not exist "%_LicensesPath%\ProPlus*.xrm-ms" (
if %_Office15% EQU 0 (echo 错误：没有检测到 Office C2R 许可文件&goto :%_fC2R%) else (goto :Reg15istry)
)
if not exist "%_Integrator%" (
if %_Office15% EQU 0 (echo 错误：没有检测到 Office C2R 许可集成程序&goto :%_fC2R%) else (goto :Reg15istry)
)
if exist "%_LicensesPath%\Word2019VL_KMS_Client_AE*.xrm-ms" (set "_tag=2019"&set "_ons= 2019")
if exist "%_LicensesPath%\Word2021VL_KMS_Client_AE*.xrm-ms" (set _LTSC=1)
if %winbuild% LSS 10240 if !_LTSC! EQU 1 (set "_tag=2021"&set "_ons= 2021")
if %_Office15% EQU 0 goto :CheckC2R

:Reg15istry
set "_Install15Root="
set "_Product15Ids="
set "_Con15fig="
set "_PR15IDs="
set "_OSPP15Ready="
set "_Licenses15Path="
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul6%') do (set "_Install15Root=%%b\root")
if not "%_Install15Root%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do (set "_Product15Ids=%%b")
  set "_Con15fig=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds"
  set "_PR15IDs=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs"
  set "_OSPP15Ready=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration"
) else (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul6%') do (set "_Install15Root=%%b\root")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do (set "_Product15Ids=%%b")
  set "_Con15fig=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds"
  set "_PR15IDs=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs"
  set "_OSPP15Ready=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\Configuration"
)
set "_OSPP15ReadT=REG_SZ"
if "%_Product15Ids%"=="" (
reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid %_Nul3% && (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid" %_Nul6%') do (set "_Product15Ids=%%b")
  set "_Con15fig=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid"
  set "_OSPP15Ready=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun"
  set "_OSPP15ReadT=REG_DWORD"
  )
reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid %_Nul3% && (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid" %_Nul6%') do (set "_Product15Ids=%%b")
  set "_Con15fig=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid"
  set "_OSPP15Ready=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun"
  set "_OSPP15ReadT=REG_DWORD"
  )
)
set "_Licenses15Path=%_Install15Root%\Licenses"
if exist "%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS" (
  set "_OSPP15VBS=%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS"
) else if exist "%ProgramW6432%\Microsoft Office\Office15\OSPP.VBS" (
  set "_OSPP15VBS=%ProgramW6432%\Microsoft Office\Office15\OSPP.VBS"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS" (
  set "_OSPP15VBS=%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS"
)
if "%_Product15Ids%"=="" (
if %_Office16% EQU 0 (echo 错误：没有检测到 Office 2013 C2R 产品 ID&goto :%_fC2R%) else (goto :CheckC2R)
)
if not exist "%_Licenses15Path%\ProPlus*.xrm-ms" (
if %_Office16% EQU 0 (echo 错误：没有检测到 Office 2013 C2R 许可文件&goto :%_fC2R%) else (goto :CheckC2R)
)
if %winbuild% LSS 9200 if not exist "%_OSPP15VBS%" (
if %_Office16% EQU 0 (echo 错误：没有检测到 Office 2013 C2R 许可工具 OSPP.vbs&goto :%_fC2R%) else (goto :CheckC2R)
)

:CheckC2R
set _OMSI=0
if %_Office16% EQU 0 (
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set _OMSI=1
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set _OMSI=1
)
if %_Office15% EQU 0 (
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set _OMSI=1
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set _OMSI=1
)
if %winbuild% GEQ 9200 (
set _spp=SoftwareLicensingProduct
set _sps=SoftwareLicensingService
set "_vbsi=%_SLMGR% /ilc "
set "_vbsf=%_SLMGR% /ilc "
) else (
set _spp=OfficeSoftwareProtectionProduct
set _sps=OfficeSoftwareProtectionService
set _vbsi="!_OSPP15VBS!" /inslic:
set _vbsf="!_OSPPVBS!" /inslic:
)
set "_wmi="
set "_qr=%_zz7% %_sps% %_zz3% Version %_zz8%"
for /f "tokens=2 delims==" %%# in ('%_qr%') do set _wmi=%%#
if "%_wmi%"=="" (
echo 错误：没有检测到 %_sps% WMI 版本
call :CheckWS
goto :%_fC2R%
)
set _Retail=0
set "_ocq=ApplicationID='%_oApp%' AND LicenseStatus='1' AND PartialProductKey is not NULL"
if %WMI_VBS% EQU 0 wmic path %_spp% where (%_ocq%) get Description %_Nul2% |findstr /V /R "^$" >"!_temp!\crvRetail.txt"
set "_qr=%_csq% %_spp% "%_ocq%" Description"
if %WMI_VBS% NEQ 0 %_qr% %_Nul2% >"!_temp!\crvRetail.txt"
find /i "RETAIL channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
find /i "RETAIL(MAK) channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
find /i "TIMEBASED_SUB channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
set rancopp=0
if %_Retail% EQU 0 if %_OMSI% EQU 0 (
set rancopp=1
%_Nul3% powershell "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':cleanlicense\:.*';iex ($f[1]);"
)
set _O16O365=0
set _C16Msg=0
set _C15Msg=0
set "_qr=%_csq% %_spp% "%_ocq%" LicenseFamily"
if %_Retail% EQU 1 if %WMI_VBS% EQU 0 wmic path %_spp% where (%_ocq%) get LicenseFamily %_Nul2% |findstr /V /R "^$" >"!_temp!\crvRetail.txt"
if %_Retail% EQU 1 if %WMI_VBS% NEQ 0 %_qr% %_Nul2% >"!_temp!\crvRetail.txt"
set "_qr=%_csq% %_spp% "ApplicationID='%_oApp%'" LicenseFamily"
if %WMI_VBS% EQU 0 wmic path %_spp% where "ApplicationID='%_oApp%'" get LicenseFamily %_Nul2% |findstr /V /R "^$" >"!_temp!\crvVolume.txt" 2>&1
if %WMI_VBS% NEQ 0 %_qr% %_Nul2% >"!_temp!\crvVolume.txt" 2>&1

if %_Office16% EQU 0 goto :R15V

set _O21Ids=ProPlus2021,ProjectPro2021,VisioPro2021,Standard2021,ProjectStd2021,VisioStd2021,Access2021,SkypeforBusiness2021
set _O19Ids=ProPlus2019,ProjectPro2019,VisioPro2019,Standard2019,ProjectStd2019,VisioStd2019,Access2019,SkypeforBusiness2019
set _O16Ids=ProjectPro,VisioPro,Standard,ProjectStd,VisioStd,Access,SkypeforBusiness
set _A21Ids=Excel2021,Outlook2021,PowerPoint2021,Publisher2021,Word2021
set _A19Ids=Excel2019,Outlook2019,PowerPoint2019,Publisher2019,Word2019
set _A16Ids=Excel,Outlook,PowerPoint,Publisher,Word
set _V21Ids=%_O21Ids%,%_A21Ids%
set _V19Ids=%_O19Ids%,%_A19Ids%
set _V16Ids=Mondo,%_O16Ids%,%_A16Ids%,OneNote
set _R16Ids=%_V16Ids%,Professional,HomeBusiness,HomeStudent,O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem,O365EduCloud
set _RetIds=%_V21Ids%,Professional2021,HomeBusiness2021,HomeStudent2021,%_V19Ids%,Professional2019,HomeBusiness2019,HomeStudent2019,%_R16Ids%
set _Suites=Mondo,O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem,O365EduCloud,ProPlus,Standard,Professional,HomeBusiness,HomeStudent,ProPlus2019,Standard2019,Professional2019,HomeBusiness2019,HomeStudent2019,ProPlus2021,Standard2021,Professional2021,HomeBusiness2021,HomeStudent2021
set _PrjSKU=ProjectPro,ProjectStd,ProjectPro2019,ProjectStd2019,ProjectPro2021,ProjectStd2021
set _VisSKU=VisioPro,VisioStd,VisioPro2019,VisioStd2019,VisioPro2021,VisioStd2021

echo %_ProductIds%>"!_temp!\crvProductIds.txt"
for %%a in (%_RetIds%,ProPlus) do (
set _%%a=0
)
for %%a in (%_RetIds%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProductIds.txt" %_Nul1% && set _%%a=1
)
if !_LTSC! EQU 0 for %%a in (%_V21Ids%) do (
set _%%a=0
)
if !_LTSC! EQU 1 for %%a in (%_V21Ids%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office21%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
for %%a in (%_V19Ids%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office19%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
for %%a in (%_V16Ids%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office16%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
reg query %_PRIDs%\ProPlusRetail.16 %_Nul3% && (
  find /i "Office16ProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
reg query %_PRIDs%\ProPlusVolume.16 %_Nul3% && (
  find /i "Office16ProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
if %_Retail% EQU 1 for %%a in (%_RetIds%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office16%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aR_Sub" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aR_PIN" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aE5R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aEDUR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aO365R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aCO365R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office16%%aXC2RVL_MAKC2R" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R16=1)
  find /i "Office19%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R19=1)
  find /i "Office19%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R19=1)
  find /i "Office19%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R19=1)
  find /i "Office19%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R19=1)
  find /i "Office21%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R21=1)
  find /i "Office21%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R21=1)
  find /i "Office21%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R21=1)
  find /i "Office21%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R21=1)
  )
)
if %_Retail% EQU 1 reg query %_PRIDs%\ProPlusRetail.16 %_Nul3% && (
  find /i "Office16ProPlusR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0 & set aC2R16=1)
  find /i "Office16ProPlusR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0 & set aC2R16=1)
  find /i "Office16ProPlusMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0 & set aC2R16=1)
  find /i "Office16ProPlusVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0 & set aC2R16=1)
)
set "_qr=%_zz1% %_spp% %_zz2% "ApplicationID='%_oApp%' AND LicenseFamily like 'Office16O365%%'" %_zz3% LicenseFamily %_zz4%"
find /i "Office16MondoVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (
%_qr% %_Nul2% | find /i "O365" %_Nul1% && (
  for %%a in (O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem,O365EduCloud) do set _%%a=0
  )
)
if %sub_o365% EQU 1 (
  for %%a in (%_Suites%) do set _%%a=0
echo.
echo Microsoft Office 已使用 vNext 许可证激活。
)
if %sub_proj% EQU 1 (
  for %%a in (%_PrjSKU%) do set _%%a=0
echo.
echo Microsoft Project 已使用 vNext 许可证激活。
)
if %sub_vsio% EQU 1 (
  for %%a in (%_VisSKU%) do set _%%a=0
echo.
echo Microsoft Visio 已使用 vNext 许可证激活。
)

for %%a in (%_RetIds%,ProPlus) do if !_%%a! EQU 1 (
set _C16Msg=1
)
if %_C16Msg% EQU 1 (
echo.
echo 正在将 Office C2R 零售版本转换为批量版本：
)
if %_C16Msg% EQU 0 (if %_Office15% EQU 1 (goto :R15V) else (goto :GVLKC2R))

for %%# in ("!_LicensesPath!\client-issuance-*.xrm-ms") do (
%_cscript% %_vbsf%"!_LicensesPath!\%%~nx#"
)
%_cscript% %_vbsf%"!_LicensesPath!\pkeyconfig-office.xrm-ms"

if !_Mondo! EQU 1 (
call :InsLic Mondo
)
if !_O365ProPlus! EQU 1 (
echo O365ProPlus 2016 套件 ^<-^> Mondo 2016 许可
call :InsLic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365Business! EQU 1 if !_O365ProPlus! EQU 0 (
set _O365ProPlus=1
echo O365Business 2016 套件 ^<-^> Mondo 2016 许可
call :InsLic O365Business NCHRJ-3VPGW-X73DM-6B36K-3RQ6B
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365SmallBusPrem! EQU 1 if !_O365Business! EQU 0 if !_O365ProPlus! EQU 0 (
set _O365ProPlus=1
echo O365SmallBusPrem 2016 套件 ^<-^> Mondo 2016 许可
call :InsLic O365SmallBusPrem 3FBRX-NFP7C-6JWVK-F2YGK-H499R
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365HomePrem! EQU 1 if !_O365SmallBusPrem! EQU 0 if !_O365Business! EQU 0 if !_O365ProPlus! EQU 0 (
set _O365ProPlus=1
echo O365HomePrem 2016 套件 ^<-^> Mondo 2016 许可
call :InsLic O365HomePrem 9FNY8-PWWTY-8RY4F-GJMTV-KHGM9
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365EduCloud! EQU 1 if !_O365HomePrem! EQU 0 if !_O365SmallBusPrem! EQU 0 if !_O365Business! EQU 0 if !_O365ProPlus! EQU 0 (
set _O365ProPlus=1
echo O365EduCloud 2016 套件 ^<-^> Mondo 2016 许可
call :InsLic O365EduCloud 8843N-BCXXD-Q84H8-R4Q37-T3CPT
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365ProPlus! EQU 1 set _O16O365=1
if !_Mondo! EQU 1 if !_O365ProPlus! EQU 0 (
echo Mondo 2016 套件
call :InsLic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if %_Office15% EQU 1 (goto :R15V) else (goto :GVLKC2R)
)
if !_ProPlus2021! EQU 1 if !_O365ProPlus! EQU 0 (
echo ProPlus 2021 套件
call :InsLic ProPlus2021
)
if !_ProPlus2019! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 (
echo ProPlus 2019 套件 -^> ProPlus%_ons% 许可
call :InsLic ProPlus%_tag%
)
if !_ProPlus! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 (
echo ProPlus 2016 套件 -^> ProPlus%_ons% 许可
call :InsLic ProPlus%_tag%
)
if !_Professional2021! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 (
echo Professional 2021 套件 -^> ProPlus 2021 许可
call :InsLic ProPlus2021
)
if !_Professional2019! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 (
echo Professional 2019 套件 -^> ProPlus%_ons% 许可
call :InsLic ProPlus%_tag%
)
if !_Professional! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 (
echo Professional 2016 套件 -^> ProPlus%_ons% 许可
call :InsLic ProPlus%_tag%
)
if !_Standard2021! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 (
echo Standard 2021 套件
call :InsLic Standard2021
)
if !_Standard2019! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2021! EQU 0 (
echo Standard 2019 套件 -^> Standard%_ons% 许可
call :InsLic Standard%_tag%
)
if !_Standard! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2021! EQU 0 if !_Standard2019! EQU 0 (
echo Standard 2016 套件 -^> Standard%_ons% 许可
call :InsLic Standard%_tag%
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a2021! EQU 1 (
  echo %%a 2021 SKU
  call :InsLic %%a2021
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a2019! EQU 1 (
if !_%%a2021! EQU 0 (
  echo %%a 2019 SKU -^> %%a%_ons% 许可
  call :InsLic %%a%_tag%
  )
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a! EQU 1 (
if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 (
  echo %%a 2016 SKU -^> %%a%_ons% 许可
  call :InsLic %%a%_tag%
  )
)
for %%a in (HomeBusiness,HomeStudent) do if !_%%a2021! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2021! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 (
  set _Standard2021=1
  echo %%a 2021 套件 -^> Standard 2021 许可
  call :InsLic Standard2021
  )
)
for %%a in (HomeBusiness,HomeStudent) do if !_%%a2019! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2021! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 if !_%%a2021! EQU 0 (
  set _Standard2019=1
  echo %%a 2019 套件 -^> Standard%_ons% 许可
  call :InsLic Standard%_tag%
  )
)
for %%a in (HomeBusiness,HomeStudent) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2021! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 (
  set _Standard=1
  echo %%a 2016 套件 -^> Standard%_ons% 许可
  call :InsLic Standard%_tag%
  )
)
for %%a in (%_A21Ids%,OneNote) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2021! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 (
  echo %%a 应用
  call :InsLic %%a
  )
)
for %%a in (%_A16Ids%) do if !_%%a2019! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2021! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 if !_%%a2021! EQU 0 (
  echo %%a 2019 应用 -^> %%a%_ons% 许可
  call :InsLic %%a%_tag%
  )
)
for %%a in (%_A16Ids%) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2021! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 (
  echo %%a 2016 应用 -^> %%a%_ons% 许可
  call :InsLic %%a%_tag%
  )
)
for %%a in (Access) do if !_%%a2021! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 (
  echo %%a 2021 应用
  call :InsLic %%a2021
  )
)
for %%a in (Access) do if !_%%a2019! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_%%a2021! EQU 0 (
  echo %%a 2019 应用 -^> %%a%_ons% 许可
  call :InsLic %%a%_tag%
  )
)
for %%a in (Access) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2021! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 (
  echo %%a 2016 应用 -^> %%a%_ons% 许可
  call :InsLic %%a%_tag%
  )
)
for %%a in (SkypeforBusiness) do if !_%%a2021! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 (
  echo %%a 2021 应用
  call :InsLic %%a2021
  )
)
for %%a in (SkypeforBusiness) do if !_%%a2019! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_%%a2021! EQU 0 (
  echo %%a 2019 应用 -^> %%a%_ons% 许可
  call :InsLic %%a%_tag%
  )
)
for %%a in (SkypeforBusiness) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2021! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 (
  echo %%a 2016 应用 -^> %%a%_ons% 许可
  call :InsLic %%a%_tag%
  )
)
if %_Office15% EQU 1 (goto :R15V) else (goto :GVLKC2R)

:R15V
set _O15Ids=Standard,ProjectPro,VisioPro,ProjectStd,VisioStd,Access,Lync
set _A15Ids=Excel,Groove,InfoPath,OneNote,Outlook,PowerPoint,Publisher,Word
set _R15Ids=SPD,Mondo,%_O15Ids%,%_A15Ids%,Professional,HomeBusiness,HomeStudent,O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem
set _V15Ids=Mondo,%_O15Ids%,%_A15Ids%

echo %_Product15Ids%>"!_temp!\crvProduct15s.txt"
for %%a in (%_R15Ids%,ProPlus) do (
set _%%a=0
)
for %%a in (%_R15Ids%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProduct15s.txt" %_Nul1% && set _%%a=1
)
for %%a in (%_V15Ids%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProduct15s.txt" %_Nul1% && (
  find /i "Office%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
reg query %_PR15IDs%\Active\ProPlusRetail\x-none %_Nul3% && (
  find /i "OfficeProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
reg query %_PR15IDs%\Active\ProPlusVolume\x-none %_Nul3% && (
  find /i "OfficeProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
if %_Retail% EQU 1 for %%a in (%_R15Ids%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProduct15s.txt" %_Nul1% && (
  find /i "Office%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R15=1)
  find /i "Office%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R15=1)
  find /i "Office%%aR_Sub" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R15=1)
  find /i "Office%%aR_PIN" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R15=1)
  find /i "Office%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R15=1)
  find /i "Office%%aO365R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R15=1)
  find /i "Office%%aCO365R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R15=1)
  find /i "Office%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0 & set aC2R15=1)
  )
)
if %_Retail% EQU 1 reg query %_PR15IDs%\Active\ProPlusRetail\x-none %_Nul3% && (
  find /i "OfficeProPlusR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0 & set aC2R15=1)
  find /i "OfficeProPlusR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0 & set aC2R15=1)
  find /i "OfficeProPlusMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0 & set aC2R15=1)
  find /i "OfficeProPlusVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0 & set aC2R15=1)
)
set "_qr=%_zz1% %_spp% %_zz2% "ApplicationID='%_oApp%' AND LicenseFamily like 'OfficeO365%%'" %_zz3% LicenseFamily %_zz4%"
find /i "OfficeMondoVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (
%_qr% %_Nul2% | find /i "O365" %_Nul1% && (
  for %%a in (O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem) do set _%%a=0
  )
)

for %%a in (%_R15Ids%,ProPlus) do if !_%%a! EQU 1 (
set _C15Msg=1
)
if %_C15Msg% EQU 1 if %_C16Msg% EQU 0 (
echo.
echo 正在将 Office C2R 零售版本转换为批量版本：
)
if %_C15Msg% EQU 0 goto :GVLKC2R

for %%# in ("!_Licenses15Path!\client-issuance-*.xrm-ms") do (
%_cscript% %_vbsi%"!_Licenses15Path!\%%~nx#"
)
%_cscript% %_vbsi%"!_Licenses15Path!\pkeyconfig-office.xrm-ms"

if !_Mondo! EQU 1 (
call :Ins15Lic Mondo
)
if !_O365ProPlus! EQU 1 if !_O16O365! EQU 0 (
echo O365ProPlus 2013 套件 ^<-^> Mondo 2013 许可
call :Ins15Lic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365SmallBusPrem! EQU 1 if !_O365ProPlus! EQU 0 if !_O16O365! EQU 0 (
set _O365ProPlus=1
echo O365SmallBusPrem 2013 套件 ^<-^> Mondo 2013 许可
call :Ins15Lic O365SmallBusPrem 3FBRX-NFP7C-6JWVK-F2YGK-H499R
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365HomePrem! EQU 1 if !_O365SmallBusPrem! EQU 0 if !_O365ProPlus! EQU 0 if !_O16O365! EQU 0 (
set _O365ProPlus=1
echo O365HomePrem 2013 套件 ^<-^> Mondo 2013 许可
call :Ins15Lic O365HomePrem 9FNY8-PWWTY-8RY4F-GJMTV-KHGM9
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365Business! EQU 1 if !_O365HomePrem! EQU 0 if !_O365SmallBusPrem! EQU 0 if !_O365ProPlus! EQU 0 if !_O16O365! EQU 0 (
set _O365ProPlus=1
echo O365Business 2013 套件 ^<-^> Mondo 2013 许可
call :Ins15Lic O365Business MCPBN-CPY7X-3PK9R-P6GTT-H8P8Y
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_Mondo! EQU 1 if !_O365ProPlus! EQU 0 if !_O16O365! EQU 0 (
echo Mondo 2013 套件
call :Ins15Lic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
goto :GVLKC2R
)
if !_SPD! EQU 1 if !_Mondo! EQU 0 if !_O365ProPlus! EQU 0 (
echo SharePoint Designer 2013 应用 -^> Mondo 2013 许可
call :Ins15Lic Mondo
goto :GVLKC2R
)
if !_ProPlus! EQU 1 if !_O365ProPlus! EQU 0 (
echo ProPlus 2013 套件
call :Ins15Lic ProPlus
)
if !_Professional! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 (
echo Professional 2013 套件 -^> ProPlus 2013 许可
call :Ins15Lic ProPlus
)
if !_Standard! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 if !_Professional! EQU 0 (
echo Standard 2013 套件
call :Ins15Lic Standard
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a! EQU 1 (
echo %%a 2013 SKU
call :Ins15Lic %%a
)
for %%a in (HomeBusiness,HomeStudent) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 if !_Professional! EQU 0 if !_Standard! EQU 0 (
  set _Standard=1
  echo %%a 2013 套件 -^> Standard 2013 许可
  call :Ins15Lic Standard
  )
)
for %%a in (%_A15Ids%) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 if !_Professional! EQU 0 if !_Standard! EQU 0 (
  echo %%a 2013 应用
  call :Ins15Lic %%a
  )
)
for %%a in (Access) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 if !_Professional! EQU 0 (
  echo %%a 2013 应用
  call :Ins15Lic %%a
  )
)
for %%a in (Lync) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 (
  echo SkypeforBusiness 2015 应用
  call :Ins15Lic %%a
  )
)
goto :GVLKC2R

:InsLic
set "_ID=%1Volume"
set "_patt=%1VL_"
set "_pkey="
set "_kpey="
if not "%2"=="" (
set "_ID=%1Retail"
set "_patt=%1R_"
set "_pkey=PidKey=%2"
set "_kpey=%2"
)
reg delete %_Config% /f /v %_ID%.OSPPReady %_Nul3%
"!_Integrator!" /I /License PRIDName=%_ID%.16 %_pkey% PackageGUID="%_GUID%" PackageRoot="!_InstallRoot!" %_Nul1%

set fallback=0
set "_qr=wmic path %_spp% where ApplicationID='%_oApp%' get LicenseFamily"
if %WMI_VBS% NEQ 0 set "_qr=%_csq% %_spp% "ApplicationID='%_oApp%'" LicenseFamily"
%_qr% %_Nul2% | find /i "%_patt%" %_Nul1% || (set fallback=1)
if %fallback% equ 0 goto :IntOK

set "_lsfs="
for %%# in ("!_LicensesPath!\%_patt%*.xrm-ms") do (
set "_lsfs=!_lsfs! %%~nx#"
)
if defined _kpey (
  for %%# in ("!_LicensesPath!\%1DemoR*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1E5R*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1EDUR*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1MSDNR*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1O365R*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1CO365R*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
)
for %%# in (!_lsfs!) do (
%_cscript% %_vbsf%"!_LicensesPath!\%%#"
)
set "_qr=wmic path %_sps% where Version='%_wmi%' call InstallProductKey ProductKey="%_kpey%""
if %WMI_VBS% NEQ 0 set "_qr=%_csp% %_sps% "%_kpey%""
if defined _kpey %_qr% %_Nul3%

:IntOK
reg add %_Config% /f /v %_ID%.OSPPReady /t REG_SZ /d 1 %_Nul1%
reg query %_Config% /v ProductReleaseIds | findstr /I "%_ID%" %_Nul1%
if %errorlevel% NEQ 0 (
for /f "skip=2 tokens=2*" %%a in ('reg query %_Config% /v ProductReleaseIds') do reg add %_Config% /v ProductReleaseIds /t REG_SZ /d "%%b,%_ID%" /f %_Nul1%
)
exit /b

:Ins15Lic
set "_ID=%1Volume"
set "_patt=%1VL_"
set "_pkey="
if not "%2"=="" (
set "_ID=%1Retail"
set "_patt=%1R_"
set "_pkey=%2"
)
reg delete %_OSPP15Ready% /f /v %_ID%.OSPPReady %_Nul3%
for %%# in ("!_Licenses15Path!\%_patt%*.xrm-ms") do (
%_cscript% %_vbsi%"!_Licenses15Path!\%%~nx#"
)
set "_qr=wmic path %_sps% where Version='%_wmi%' call InstallProductKey ProductKey="%_pkey%""
if %WMI_VBS% NEQ 0 set "_qr=%_csp% %_sps% "%_pkey%""
if defined _pkey %_qr% %_Nul3%
reg add %_OSPP15Ready% /f /v %_ID%.OSPPReady /t %_OSPP15ReadT% /d 1 %_Nul1%
reg query %_Con15fig% %_Nul2% | findstr /I "%_ID%" %_Nul1%
if %errorlevel% NEQ 0 (
for /f "skip=2 tokens=2*" %%a in ('reg query %_Con15fig% %_Nul6%') do reg add %_Con15fig% /t REG_SZ /d "%%b,%_ID%" /f %_Nul1%
)
exit /b

:GVLKC2R
set _CtRMsg=0
if %_C16Msg% EQU 1 set _CtRMsg=1
if %_C15Msg% EQU 1 set _CtRMsg=1
if %_Office16% EQU 1 (
for %%a in (%_RetIds%,ProPlus) do set "_%%a="
for %%A in (19,21) do call :officeLoc %%A
)
if %_Office15% EQU 1 (
for %%a in (%_R15Ids%,ProPlus) do set "_%%a="
)
set "_qr=wmic path %_sps% where version='%_wmi%' call RefreshLicenseStatus"
if %WMI_VBS% NEQ 0 set "_qr=%_csm% "%_sps%.Version='%_wmi%'" RefreshLicenseStatus"
if %winbuild% GEQ 9200 %_qr% %_Nul3%
if exist "%SysPath%\spp\store_test\2.0\tokens.dat" if %rancopp% EQU 1 if %_CtRMsg% EQU 1 (
%_cscript% %_SLMGR% /rilc
if !ERRORLEVEL! NEQ 0 %_cscript% %_SLMGR% /rilc
)
goto :%_sC2R%

:keys
if "%~1"=="" exit /b
set yh=-
goto :%1 %_Nul2%

::  Windows 11 [Ni]
:59eb965c-9150-42b7-a0ec-22151b9897c5
set "_key=KBN8V%yh%HFGQ4%yh%MGXVD%yh%347P6%yh%PDQGT" &::  IoT Enterprise LTSC
exit /b

::  Windows 11 [Co]
:ca7df2e3-5ea0-47b8-9ac1-b1be4d8edd69
set "_key=37D7F%yh%N49CB%yh%WQR8W%yh%TBJ73%yh%FM8RX" &::  SE {Cloud}
exit /b

:d30136fc-cb4b-416e-a23d-87207abc44a9
set "_key=6XN7V%yh%PCBDC%yh%BDBRH%yh%8DQY7%yh%G6R44" &::  SE N {Cloud N}
exit /b

::  Windows 10 [RS5]
:32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee
set "_key=M7XTQ%yh%FN8P6%yh%TTKYV%yh%9D4CC%yh%J462D" &::  Enterprise LTSC 2019
exit /b

:7103a333-b8c8-49cc-93ce-d37c09687f92
set "_key=92NFX%yh%8DJQP%yh%P6BBQ%yh%THF9C%yh%7CG2H" &::  Enterprise LTSC 2019 N
exit /b

:ec868e65-fadf-4759-b23e-93fe37f2cc29
set "_key=CPWHC%yh%NT2C7%yh%VYW78%yh%DHDB2%yh%PG3GK" &::  Enterprise for Virtual Desktops
exit /b

:0df4f814-3f57-4b8b-9a9d-fddadcd69fac
set "_key=NBTWJ%yh%3DR69%yh%3C4V8%yh%C26MC%yh%GQ9M6" &::  Lean
exit /b

::  Windows 10 [RS3]
:82bbc092-bc50-4e16-8e18-b74fc486aec3
set "_key=NRG8B%yh%VKK3Q%yh%CXVCJ%yh%9G2XF%yh%6Q84J" &::  Pro Workstation
exit /b

:4b1571d3-bafb-4b40-8087-a961be2caf65
set "_key=9FNHH%yh%K3HBT%yh%3W4TD%yh%6383H%yh%6XYWF" &::  Pro Workstation N
exit /b

:e4db50ea-bda1-4566-b047-0ca50abc6f07
set "_key=7NBT4%yh%WGBQX%yh%MP4H7%yh%QXFF8%yh%YP3KX" &::  Enterprise Remote Server
exit /b

::  Windows 10 [RS2]
:e0b2d383-d112-413f-8a80-97f373a5820c
set "_key=YYVX9%yh%NTFWV%yh%6MDM3%yh%9PT4T%yh%4M68B" &::  Enterprise G
exit /b

:e38454fb-41a4-4f59-a5dc-25080e354730
set "_key=44RPN%yh%FTY23%yh%9VTTB%yh%MP9BX%yh%T84FV" &::  Enterprise G N
exit /b

::  Windows 10 [RS1]
:2d5a5a60-3040-48bf-beb0-fcd770c20ce0
set "_key=DCPHK%yh%NFMTC%yh%H88MJ%yh%PFHPY%yh%QJ4BJ" &::  Enterprise 2016 LTSB
exit /b

:9f776d83-7156-45b2-8a5c-359b9c9f22a3
set "_key=QFFDN%yh%GRT3P%yh%VKWWX%yh%X7T3R%yh%8B639" &::  Enterprise 2016 LTSB N
exit /b

:3f1afc82-f8ac-4f6c-8005-1d233e606eee
set "_key=6TP4R%yh%GNPTD%yh%KYYHQ%yh%7B7DP%yh%J447Y" &::  Pro Education
exit /b

:5300b18c-2e33-4dc2-8291-47ffcec746dd
set "_key=YVWGF%yh%BXNMC%yh%HTQYQ%yh%CPQ99%yh%66QFC" &::  Pro Education N
exit /b

::  Windows 10 [TH]
:58e97c99-f377-4ef1-81d5-4ad5522b5fd8
set "_key=TX9XD%yh%98N7V%yh%6WMQ6%yh%BX7FG%yh%H8Q99" &::  Home
exit /b

:7b9e1751-a8da-4f75-9560-5fadfe3d8e38
set "_key=3KHY7%yh%WNT83%yh%DGQKR%yh%F7HPR%yh%844BM" &::  Home N
exit /b

:cd918a57-a41b-4c82-8dce-1a538e221a83
set "_key=7HNRX%yh%D7KGG%yh%3K4RQ%yh%4WPJ4%yh%YTDFH" &::  Home Single Language
exit /b

:a9107544-f4a0-4053-a96a-1479abdef912
set "_key=PVMJN%yh%6DFY6%yh%9CCP6%yh%7BKTT%yh%D3WVR" &::  Home China
exit /b

:2de67392-b7a7-462a-b1ca-108dd189f588
set "_key=W269N%yh%WFGWX%yh%YVC9B%yh%4J6C9%yh%T83GX" &::  Pro
exit /b

:a80b5abf-76ad-428b-b05d-a47d2dffeebf
set "_key=MH37W%yh%N47XK%yh%V7XM9%yh%C7227%yh%GCQG9" &::  Pro N
exit /b

:e0c42288-980c-4788-a014-c080d2e1926e
set "_key=NW6C2%yh%QMPVW%yh%D7KKK%yh%3GKT6%yh%VCFB2" &::  Education
exit /b

:3c102355-d027-42c6-ad23-2e7ef8a02585
set "_key=2WH4N%yh%8QGBV%yh%H22JP%yh%CT43Q%yh%MDWWJ" &::  Education N
exit /b

:73111121-5638-40f6-bc11-f1d7b0d64300
set "_key=NPPR9%yh%FWDCX%yh%D2C8J%yh%H872K%yh%2YT43" &::  Enterprise
exit /b

:e272e3e2-732f-4c65-a8f0-484747d0d947
set "_key=DPH2V%yh%TTNVB%yh%4X9Q3%yh%TJR4H%yh%KHJW4" &::  Enterprise N
exit /b

:7b51a46c-0c04-4e8f-9af4-8496cca90d5e
set "_key=WNMTR%yh%4C88C%yh%JK8YV%yh%HQ7T2%yh%76DF9" &::  Enterprise 2015 LTSB
exit /b

:87b838b7-41b6-4590-8318-5797951d8529
set "_key=2F77B%yh%TNFGY%yh%69QQF%yh%B8YKP%yh%D69TJ" &::  Enterprise 2015 LTSB N
exit /b

::  Windows Server 2022 [Fe]
:9bd77860-9b31-4b7b-96ad-2564017315bf
set "_key=VDYBN%yh%27WPP%yh%V4HQT%yh%9VMD4%yh%VMK7H" &::  Standard
exit /b

:ef6cfc9f-8c5d-44ac-9aad-de6a2ea0ae03
set "_key=WX4NM%yh%KYWYW%yh%QJJR4%yh%XV3QB%yh%6VM33" &::  Datacenter
exit /b

:8c8f0ad3-9a43-4e05-b840-93b8d1475cbc
set "_key=6N379%yh%GGTMK%yh%23C6M%yh%XVVTC%yh%CKFRQ" &::  Azure Core
exit /b

:f5e9429c-f50b-4b98-b15c-ef92eb5cff39
set "_key=67KN8%yh%4FYJW%yh%2487Q%yh%MQ2J7%yh%4C4RG" &::  Standard ACor
exit /b

:39e69c41-42b4-4a0a-abad-8e3c10a797cc
set "_key=QFND9%yh%D3Y9C%yh%J3KKY%yh%6RPVP%yh%2DPYV" &::  Datacenter ACor
exit /b

::  Windows Server 2019 [RS5]
:de32eafd-aaee-4662-9444-c1befb41bde2
set "_key=N69G4%yh%B89J2%yh%4G8F4%yh%WWYCC%yh%J464C" &::  Standard
exit /b

:34e1ae55-27f8-4950-8877-7a03be5fb181
set "_key=WMDGN%yh%G9PQG%yh%XVVXX%yh%R3X43%yh%63DFG" &::  Datacenter
exit /b

:a99cc1f0-7719-4306-9645-294102fbff95
set "_key=FDNH6%yh%VW9RW%yh%BXPJ7%yh%4XTYG%yh%239TB" &::  Azure Core
exit /b

:73e3957c-fc0c-400d-9184-5f7b6f2eb409
set "_key=N2KJX%yh%J94YW%yh%TQVFB%yh%DG9YT%yh%724CC" &::  Standard ACor
exit /b

:90c362e5-0da1-4bfd-b53b-b87d309ade43
set "_key=6NMRW%yh%2C8FM%yh%D24W7%yh%TQWMY%yh%CWH2D" &::  Datacenter ACor
exit /b

:034d3cbb-5d4b-4245-b3f8-f84571314078
set "_key=WVDHN%yh%86M7X%yh%466P6%yh%VHXV7%yh%YY726" &::  Essentials
exit /b

:8de8eb62-bbe0-40ac-ac17-f75595071ea3
set "_key=GRFBW%yh%QNDC4%yh%6QBHG%yh%CCK3B%yh%2PR88" &::  ServerARM64
exit /b

:19b5e0fb-4431-46bc-bac1-2f1873e4ae73
set "_key=NTBV8%yh%9K7Q8%yh%V27C6%yh%M2BTV%yh%KHMXV" &::  Azure Datacenter - ServerTurbine
exit /b

::  Windows Server 2016 [RS4]
:43d9af6e-5e86-4be8-a797-d072a046896c
set "_key=K9FYF%yh%G6NCK%yh%73M32%yh%XMVPY%yh%F9DRR" &::  ServerARM64
exit /b

::  Windows Server 2016 [RS3]
:61c5ef22-f14f-4553-a824-c4b31e84b100
set "_key=PTXN8%yh%JFHJM%yh%4WC78%yh%MPCBR%yh%9W4KR" &::  Standard ACor
exit /b

:e49c08e7-da82-42f8-bde2-b570fbcae76c
set "_key=2HXDN%yh%KRXHB%yh%GPYC7%yh%YCKFJ%yh%7FVDG" &::  Datacenter ACor
exit /b

::  Windows Server 2016 [RS1]
:8c1c5410-9f39-4805-8c9d-63a07706358f
set "_key=WC2BQ%yh%8NRM3%yh%FDDYY%yh%2BFGV%yh%KHKQY" &::  Standard
exit /b

:21c56779-b449-4d20-adfc-eece0e1ad74b
set "_key=CB7KF%yh%BWN84%yh%R7R2Y%yh%793K2%yh%8XDDG" &::  Datacenter
exit /b

:3dbf341b-5f6c-4fa7-b936-699dce9e263f
set "_key=VP34G%yh%4NPPG%yh%79JTQ%yh%864T4%yh%R3MQX" &::  Azure Core
exit /b

:2b5a1b0f-a5ab-4c54-ac2f-a6d94824a283
set "_key=JCKRF%yh%N37P4%yh%C2D82%yh%9YXRT%yh%4M63B" &::  Essentials
exit /b

:7b4433f4-b1e7-4788-895a-c45378d38253
set "_key=QN4C6%yh%GBJD2%yh%FB422%yh%GHWJK%yh%GJG2R" &::  Cloud Storage
exit /b

::  Windows 8.1
:fe1c3238-432a-43a1-8e25-97e7d1ef10f3
set "_key=M9Q9P%yh%WNJJT%yh%6PXPY%yh%DWX8H%yh%6XWKK" &::  Core
exit /b

:78558a64-dc19-43fe-a0d0-8075b2a370a3
set "_key=7B9N3%yh%D94CG%yh%YTVHR%yh%QBPX3%yh%RJP64" &::  Core N
exit /b

:c72c6a1d-f252-4e7e-bdd1-3fca342acb35
set "_key=BB6NG%yh%PQ82V%yh%VRDPW%yh%8XVD2%yh%V8P66" &::  Core Single Language
exit /b

:db78b74f-ef1c-4892-abfe-1e66b8231df6
set "_key=NCTT7%yh%2RGK8%yh%WMHRF%yh%RY7YQ%yh%JTXG3" &::  Core China
exit /b

:ffee456a-cd87-4390-8e07-16146c672fd0
set "_key=XYTND%yh%K6QKT%yh%K2MRH%yh%66RTM%yh%43JKP" &::  Core ARM
exit /b

:c06b6981-d7fd-4a35-b7b4-054742b7af67
set "_key=GCRJD%yh%8NW9H%yh%F2CDX%yh%CCM8D%yh%9D6T9" &::  Pro
exit /b

:7476d79f-8e48-49b4-ab63-4d0b813a16e4
set "_key=HMCNV%yh%VVBFX%yh%7HMBH%yh%CTY9B%yh%B4FXY" &::  Pro N
exit /b

:096ce63d-4fac-48a9-82a9-61ae9e800e5f
set "_key=789NJ%yh%TQK6T%yh%6XTH8%yh%J39CJ%yh%J8D3P" &::  Pro with Media Center
exit /b

:81671aaf-79d1-4eb1-b004-8cbbe173afea
set "_key=MHF9N%yh%XY6XB%yh%WVXMC%yh%BTDCT%yh%MKKG7" &::  Enterprise
exit /b

:113e705c-fa49-48a4-beea-7dd879b46b14
set "_key=TT4HM%yh%HN7YT%yh%62K67%yh%RGRQJ%yh%JFFXW" &::  Enterprise N
exit /b

:0ab82d54-47f4-4acb-818c-cc5bf0ecb649
set "_key=NMMPB%yh%38DD4%yh%R2823%yh%62W8D%yh%VXKJB" &::  Embedded Industry Pro
exit /b

:cd4e2d9f-5059-4a50-a92d-05d5bb1267c7
set "_key=FNFKF%yh%PWTVT%yh%9RC8H%yh%32HB2%yh%JB34X" &::  Embedded Industry Enterprise
exit /b

:f7e88590-dfc7-4c78-bccb-6f3865b99d1a
set "_key=VHXM3%yh%NR6FT%yh%RY6RT%yh%CK882%yh%KW2CJ" &::  Embedded Industry Automotive
exit /b

:e9942b32-2e55-4197-b0bd-5ff58cba8860
set "_key=3PY8R%yh%QHNP9%yh%W7XQD%yh%G6DPH%yh%3J2C9" &::  with Bing
exit /b

:c6ddecd6-2354-4c19-909b-306a3058484e
set "_key=Q6HTR%yh%N24GM%yh%PMJFP%yh%69CD8%yh%2GXKR" &::  with Bing N
exit /b

:b8f5e3a3-ed33-4608-81e1-37d6c9dcfd9c
set "_key=KF37N%yh%VDV38%yh%GRRTV%yh%XH8X6%yh%6F3BB" &::  with Bing Single Language
exit /b

:ba998212-460a-44db-bfb5-71bf09d1c68b
set "_key=R962J%yh%37N87%yh%9VVK2%yh%WJ74P%yh%XTMHR" &::  with Bing China
exit /b

:e58d87b5-8126-4580-80fb-861b22f79296
set "_key=MX3RK%yh%9HNGX%yh%K3QKC%yh%6PJ3F%yh%W8D7B" &::  Pro for Students
exit /b

:cab491c7-a918-4f60-b502-dab75e334f40
set "_key=TNFGH%yh%2R6PB%yh%8XM3K%yh%QYHX2%yh%J4296" &::  Pro for Students N
exit /b

::  Windows Server 2012 R2
:b3ca044e-a358-4d68-9883-aaa2941aca99
set "_key=D2N9P%yh%3P6X9%yh%2R39C%yh%7RTCD%yh%MDVJX" &::  Standard
exit /b

:00091344-1ea4-4f37-b789-01750ba6988c
set "_key=W3GGN%yh%FT8W3%yh%Y4M27%yh%J84CP%yh%Q3VJ9" &::  Datacenter
exit /b

:21db6ba4-9a7b-4a14-9e29-64a60c59301d
set "_key=KNC87%yh%3J2TX%yh%XB4WP%yh%VCPJV%yh%M4FWM" &::  Essentials
exit /b

:b743a2be-68d4-4dd3-af32-92425b7bb623
set "_key=3NPTF%yh%33KPT%yh%GGBPR%yh%YX76B%yh%39KDD" &::  Cloud Storage
exit /b

::  Windows 8
:c04ed6bf-55c8-4b47-9f8e-5a1f31ceee60
set "_key=BN3D2%yh%R7TKB%yh%3YPBD%yh%8DRP2%yh%27GG4" &::  Core
exit /b

:197390a0-65f6-4a95-bdc4-55d58a3b0253
set "_key=8N2M2%yh%HWPGY%yh%7PGT9%yh%HGDD8%yh%GVGGY" &::  Core N
exit /b

:8860fcd4-a77b-4a20-9045-a150ff11d609
set "_key=2WN2H%yh%YGCQR%yh%KFX6K%yh%CD6TF%yh%84YXQ" &::  Core Single Language
exit /b

:9d5584a2-2d85-419a-982c-a00888bb9ddf
set "_key=4K36P%yh%JN4VD%yh%GDC6V%yh%KDT89%yh%DYFKP" &::  Core China
exit /b

:af35d7b7-5035-4b63-8972-f0b747b9f4dc
set "_key=DXHJF%yh%N9KQX%yh%MFPVR%yh%GHGQK%yh%Y7RKV" &::  Core ARM
exit /b

:a98bcd6d-5343-4603-8afe-5908e4611112
set "_key=NG4HW%yh%VH26C%yh%733KW%yh%K6F98%yh%J8CK4" &::  Pro
exit /b

:ebf245c1-29a8-4daf-9cb1-38dfc608a8c8
set "_key=XCVCF%yh%2NXM9%yh%723PB%yh%MHCB7%yh%2RYQQ" &::  Pro N
exit /b

:a00018a3-f20f-4632-bf7c-8daa5351c914
set "_key=GNBB8%yh%YVD74%yh%QJHX6%yh%27H4K%yh%8QHDG" &::  Pro with Media Center
exit /b

:458e1bec-837a-45f6-b9d5-925ed5d299de
set "_key=32JNW%yh%9KQ84%yh%P47T8%yh%D8GGY%yh%CWCK7" &::  Enterprise
exit /b

:e14997e7-800a-4cf7-ad10-de4b45b578db
set "_key=JMNMF%yh%RHW7P%yh%DMY6X%yh%RF3DR%yh%X2BQT" &::  Enterprise N
exit /b

:10018baf-ce21-4060-80bd-47fe74ed4dab
set "_key=RYXVT%yh%BNQG7%yh%VD29F%yh%DBMRY%yh%HT73M" &::  Embedded Industry Pro
exit /b

:18db1848-12e0-4167-b9d7-da7fcda507db
set "_key=NKB3R%yh%R2F8T%yh%3XCDP%yh%7Q2KW%yh%XWYQ2" &::  Embedded Industry Enterprise
exit /b

::  Windows Server 2012
:f0f5ec41-0d55-4732-af02-440a44a3cf0f
set "_key=XC9B7%yh%NBPP2%yh%83J2H%yh%RHMBY%yh%92BT4" &::  Standard
exit /b

:d3643d60-0c42-412d-a7d6-52e6635327f6
set "_key=48HP8%yh%DN98B%yh%MYWDG%yh%T2DCC%yh%8W83P" &::  Datacenter
exit /b

:7d5486c7-e120-4771-b7f1-7b56c6d3170c
set "_key=HM7DN%yh%YVMH3%yh%46JC3%yh%XYTG7%yh%CYQJJ" &::  MultiPoint Standard
exit /b

:95fd1c83-7df5-494a-be8b-1300e1c9d1cd
set "_key=XNH6W%yh%2V9GX%yh%RGJ4K%yh%Y8X6F%yh%QGJ2G" &::  MultiPoint Premium
exit /b

::  Windows 7
:b92e9980-b9d5-4821-9c94-140f632f6312
set "_key=FJ82H%yh%XT6CR%yh%J8D7P%yh%XQJJ2%yh%GPDD4" &::  Professional
exit /b

:54a09a0d-d57b-4c10-8b69-a842d6590ad5
set "_key=MRPKT%yh%YTG23%yh%K7D7T%yh%X2JMM%yh%QY7MG" &::  Professional N
exit /b

:5a041529-fef8-4d07-b06f-b59b573b32d2
set "_key=W82YF%yh%2Q76Y%yh%63HXB%yh%FGJG9%yh%GF7QX" &::  Professional E
exit /b

:ae2ee509-1b34-41c0-acb7-6d4650168915
set "_key=33PXH%yh%7Y6KF%yh%2VJC9%yh%XBBR8%yh%HVTHH" &::  Enterprise
exit /b

:1cb6d605-11b3-4e14-bb30-da91c8e3983a
set "_key=YDRBP%yh%3D83W%yh%TY26F%yh%D46B2%yh%XCKRJ" &::  Enterprise N
exit /b

:46bbed08-9c7b-48fc-a614-95250573f4ea
set "_key=C29WB%yh%22CC8%yh%VJ326%yh%GHFJW%yh%H9DH4" &::  Enterprise E
exit /b

:db537896-376f-48ae-a492-53d0547773d0
set "_key=YBYF6%yh%BHCR3%yh%JPKRB%yh%CDW7B%yh%F9BK4" &::  Embedded POSReady 7
exit /b

:e1a8296a-db37-44d1-8cce-7bc961d59c54
set "_key=XGY72%yh%BRBBT%yh%FF8MH%yh%2GG8H%yh%W7KCW" &::  Embedded Standard
exit /b

:aa6dd3aa-c2b4-40e2-a544-a6bbb3f5c395
set "_key=73KQT%yh%CD9G6%yh%K7TQG%yh%66MRP%yh%CQ22C" &::  Embedded ThinPC
exit /b

::  Windows Server 2008 R2
:a78b8bd9-8017-4df5-b86a-09f756affa7c
set "_key=6TPJF%yh%RBVHG%yh%WBW2R%yh%86QPH%yh%6RTM4" &::  Web
exit /b

:cda18cf3-c196-46ad-b289-60c072869994
set "_key=TT8MH%yh%CG224%yh%D3D7Q%yh%498W2%yh%9QCTX" &::  HPC
exit /b

:68531fb9-5511-4989-97be-d11a0f55633f
set "_key=YC6KT%yh%GKW9T%yh%YTKYR%yh%T4X34%yh%R7VHC" &::  Standard
exit /b

:7482e61b-c589-4b7f-8ecc-46d455ac3b87
set "_key=74YFP%yh%3QFB3%yh%KQT8W%yh%PMXWJ%yh%7M648" &::  Datacenter
exit /b

:620e2b3d-09e7-42fd-802a-17a13652fe7a
set "_key=489J6%yh%VHDMP%yh%X63PK%yh%3K798%yh%CPX3Y" &::  Enterprise
exit /b

:8a26851c-1c7e-48d3-a687-fbca9b9ac16b
set "_key=GT63C%yh%RJFQ3%yh%4GMB6%yh%BRFB9%yh%CB83V" &::  Itanium
exit /b

:f772515c-0e87-48d5-a676-e6962c3e1195
set "_key=736RG%yh%XDKJK%yh%V34PF%yh%BHK87%yh%J6X3K" &::  MultiPoint Server - ServerEmbeddedSolution
exit /b

::  Office 2021
:fbdb3e18-a8ef-4fb3-9183-dffd60bd0984
set "_key=FXYTK%yh%NJJ8C%yh%GB6DW%yh%3DYQT%yh%6F7TH" &::  Professional Plus
exit /b

:080a45c5-9f9f-49eb-b4b0-c3c610a5ebd3
set "_key=KDX7X%yh%BNVR8%yh%TXXGX%yh%4Q7Y8%yh%78VT3" &::  Standard
exit /b

:76881159-155c-43e0-9db7-2d70a9a3a4ca
set "_key=FTNWT%yh%C6WBT%yh%8HMGF%yh%K9PRX%yh%QV9H8" &::  Project Professional
exit /b

:6dd72704-f752-4b71-94c7-11cec6bfc355
set "_key=J2JDC%yh%NJCYY%yh%9RGQ4%yh%YXWMH%yh%T3D4T" &::  Project Standard
exit /b

:fb61ac9a-1688-45d2-8f6b-0674dbffa33c
set "_key=KNH8D%yh%FGHT4%yh%T8RK3%yh%CTDYJ%yh%K2HT4" &::  Visio Professional
exit /b

:72fce797-1884-48dd-a860-b2f6a5efd3ca
set "_key=MJVNY%yh%BYWPY%yh%CWV6J%yh%2RKRT%yh%4M8QG" &::  Visio Standard
exit /b

:1fe429d8-3fa7-4a39-b6f0-03dded42fe14
set "_key=WM8YG%yh%YNGDD%yh%4JHDC%yh%PG3F4%yh%FC4T4" &::  Access
exit /b

:ea71effc-69f1-4925-9991-2f5e319bbc24
set "_key=NWG3X%yh%87C9K%yh%TC7YY%yh%BC2G7%yh%G6RVC" &::  Excel
exit /b

:a5799e4c-f83c-4c6e-9516-dfe9b696150b
set "_key=C9FM6%yh%3N72F%yh%HFJXB%yh%TM3V9%yh%T86R9" &::  Outlook
exit /b

:6e166cc3-495d-438a-89e7-d7c9e6fd4dea
set "_key=TY7XF%yh%NFRBR%yh%KJ44C%yh%G83KF%yh%GX27K" &::  PowerPoint
exit /b

:aa66521f-2370-4ad8-a2bb-c095e3e4338f
set "_key=2MW9D%yh%N4BXM%yh%9VBPG%yh%Q7W6M%yh%KFBGQ" &::  Publisher
exit /b

:1f32a9af-1274-48bd-ba1e-1ab7508a23e8
set "_key=HWCXN%yh%K3WBT%yh%WJBKY%yh%R8BD9%yh%XK29P" &::  Skype for Business
exit /b

:abe28aea-625a-43b1-8e30-225eb8fbd9e5
set "_key=TN8H9%yh%M34D3%yh%Y64V9%yh%TR72V%yh%X79KV" &::  Word
exit /b

::  Office 2019
:85dd8b5f-eaa4-4af3-a628-cce9e77c9a03
set "_key=NMMKJ%yh%6RK4F%yh%KMJVX%yh%8D9MJ%yh%6MWKP" &::  Professional Plus
exit /b

:6912a74b-a5fb-401a-bfdb-2e3ab46f4b02
set "_key=6NWWJ%yh%YQWMR%yh%QKGCB%yh%6TMB3%yh%9D9HK" &::  Standard
exit /b

:2ca2bf3f-949e-446a-82c7-e25a15ec78c4
set "_key=B4NPR%yh%3FKK7%yh%T2MBV%yh%FRQ4W%yh%PKD2B" &::  Project Professional
exit /b

:1777f0e3-7392-4198-97ea-8ae4de6f6381
set "_key=C4F7P%yh%NCP8C%yh%6CQPT%yh%MQHV9%yh%JXD2M" &::  Project Standard
exit /b

:5b5cf08f-b81a-431d-b080-3450d8620565
set "_key=9BGNQ%yh%K37YR%yh%RQHF2%yh%38RQ3%yh%7VCBB" &::  Visio Professional
exit /b

:e06d7df3-aad0-419d-8dfb-0ac37e2bdf39
set "_key=7TQNQ%yh%K3YQQ%yh%3PFH7%yh%CCPPM%yh%X4VQ2" &::  Visio Standard
exit /b

:9e9bceeb-e736-4f26-88de-763f87dcc485
set "_key=9N9PT%yh%27V4Y%yh%VJ2PD%yh%YXFMF%yh%YTFQT" &::  Access
exit /b

:237854e9-79fc-4497-a0c1-a70969691c6b
set "_key=TMJWT%yh%YYNMB%yh%3BKTF%yh%644FC%yh%RVXBD" &::  Excel
exit /b

:c8f8a301-19f5-4132-96ce-2de9d4adbd33
set "_key=7HD7K%yh%N4PVK%yh%BHBCQ%yh%YWQRW%yh%XW4VK" &::  Outlook
exit /b

:3131fd61-5e4f-4308-8d6d-62be1987c92c
set "_key=RRNCX%yh%C64HY%yh%W2MM7%yh%MCH9G%yh%TJHMQ" &::  PowerPoint
exit /b

:9d3e4cca-e172-46f1-a2f4-1d2107051444
set "_key=G2KWX%yh%3NW6P%yh%PY93R%yh%JXK2T%yh%C9Y9V" &::  Publisher
exit /b

:734c6c6e-b0ba-4298-a891-671772b2bd1b
set "_key=NCJ33%yh%JHBBY%yh%HTK98%yh%MYCV8%yh%HMKHJ" &::  Skype for Business
exit /b

:059834fe-a8ea-4bff-b67b-4d006b5447d3
set "_key=PBX3G%yh%NWMT6%yh%Q7XBW%yh%PYJGG%yh%WXD33" &::  Word
exit /b

:0bc88885-718c-491d-921f-6f214349e79c
set "_key=VQ9DP%yh%NVHPH%yh%T9HJC%yh%J9PDT%yh%KTQRG" &::  Pro Plus 2019 Preview
exit /b

:fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9
set "_key=XM2V9%yh%DN9HH%yh%QB449%yh%XDGKC%yh%W2RMW" &::  Project Pro 2019 Preview
exit /b

:500f6619-ef93-4b75-bcb4-82819998a3ca
set "_key=N2CG9%yh%YD3YK%yh%936X4%yh%3WR82%yh%Q3X4H" &::  Visio Pro 2019 Preview
exit /b

:f3fb2d68-83dd-4c8b-8f09-08e0d950ac3b
set "_key=HFPBN%yh%RYGG8%yh%HQWCW%yh%26CH6%yh%PDPVF" &::  Pro Plus 2021 Preview
exit /b

:76093b1b-7057-49d7-b970-638ebcbfd873
set "_key=WDNBY%yh%PCYFY%yh%9WP6G%yh%BXVXM%yh%92HDV" &::  Project Pro 2021 Preview
exit /b

:a3b44174-2451-4cd6-b25f-66638bfb9046
set "_key=2XYX7%yh%NXXBK%yh%9CK7W%yh%K2TKW%yh%JFJ7G" &::  Visio Pro 2021 Preview
exit /b

::  Office 2016
:829b8110-0e6f-4349-bca4-42803577788d
set "_key=WGT24%yh%HCNMF%yh%FQ7XH%yh%6M8K7%yh%DRTW9" &::  Project Professional C2R-P
exit /b

:cbbaca45-556a-4416-ad03-bda598eaa7c8
set "_key=D8NRQ%yh%JTYM3%yh%7J2DX%yh%646CT%yh%6836M" &::  Project Standard C2R-P
exit /b

:b234abe3-0857-4f9c-b05a-4dc314f85557
set "_key=69WXN%yh%MBYV6%yh%22PQG%yh%3WGHK%yh%RM6XC" &::  Visio Professional C2R-P
exit /b

:361fe620-64f4-41b5-ba77-84f8e079b1f7
set "_key=NY48V%yh%PPYYH%yh%3F4PX%yh%XJRKJ%yh%W4423" &::  Visio Standard C2R-P
exit /b

:e914ea6e-a5fa-4439-a394-a9bb3293ca09
set "_key=DMTCJ%yh%KNRKX%yh%26982%yh%JYCKT%yh%P7KB6" &::  MondoR
exit /b

:9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce
set "_key=HFTND%yh%W9MK4%yh%8B7MJ%yh%B6C4G%yh%XQBR2" &::  Mondo
exit /b

:d450596f-894d-49e0-966a-fd39ed4c4c64
set "_key=XQNVK%yh%8JYDB%yh%WJ9W3%yh%YJ8YR%yh%WFG99" &::  Professional Plus
exit /b

:dedfa23d-6ed1-45a6-85dc-63cae0546de6
set "_key=JNRGM%yh%WHDWX%yh%FJJG3%yh%K47QV%yh%DRTFM" &::  Standard
exit /b

:4f414197-0fc2-4c01-b68a-86cbb9ac254c
set "_key=YG9NW%yh%3K39V%yh%2T3HJ%yh%93F3Q%yh%G83KT" &::  Project Professional
exit /b

:da7ddabc-3fbe-4447-9e01-6ab7440b4cd4
set "_key=GNFHQ%yh%F6YQM%yh%KQDGJ%yh%327XX%yh%KQBVC" &::  Project Standard
exit /b

:6bf301c1-b94a-43e9-ba31-d494598c47fb
set "_key=PD3PC%yh%RHNGV%yh%FXJ29%yh%8JK7D%yh%RJRJK" &::  Visio Professional
exit /b

:aa2a7821-1827-4c2c-8f1d-4513a34dda97
set "_key=7WHWN%yh%4T7MP%yh%G96JF%yh%G33KR%yh%W8GF4" &::  Visio Standard
exit /b

:67c0fc0c-deba-401b-bf8b-9c8ad8395804
set "_key=GNH9Y%yh%D2J4T%yh%FJHGG%yh%QRVH7%yh%QPFDW" &::  Access
exit /b

:c3e65d36-141f-4d2f-a303-a842ee756a29
set "_key=9C2PK%yh%NWTVB%yh%JMPW8%yh%BFT28%yh%7FTBF" &::  Excel
exit /b

:d8cace59-33d2-4ac7-9b1b-9b72339c51c8
set "_key=DR92N%yh%9HTF2%yh%97XKM%yh%XW2WJ%yh%XW3J6" &::  OneNote
exit /b

:ec9d9265-9d1e-4ed0-838a-cdc20f2551a1
set "_key=R69KK%yh%NTPKF%yh%7M3Q4%yh%QYBHW%yh%6MT9B" &::  Outlook
exit /b

:d70b1bba-b893-4544-96e2-b7a318091c33
set "_key=J7MQP%yh%HNJ4Y%yh%WJ7YM%yh%PFYGF%yh%BY6C6" &::  Powerpoint
exit /b

:041a06cb-c5b8-4772-809f-416d03d16654
set "_key=F47MM%yh%N3XJP%yh%TQXJ9%yh%BP99D%yh%8K837" &::  Publisher
exit /b

:83e04ee1-fa8d-436d-8994-d31a862cab77
set "_key=869NQ%yh%FJ69K%yh%466HW%yh%QYCP2%yh%DDBV6" &::  Skype for Business
exit /b

:bb11badf-d8aa-470e-9311-20eaf80fe5cc
set "_key=WXY84%yh%JN2Q9%yh%RBCCQ%yh%3Q3J3%yh%3PFJ6" &::  Word
exit /b

::  Office 2013
:dc981c6b-fc8e-420f-aa43-f8f33e5c0923
set "_key=42QTK%yh%RN8M7%yh%J3C4G%yh%BBGYM%yh%88CYV" &::  Mondo
exit /b

:b322da9c-a2e2-4058-9e4e-f59a6970bd69
set "_key=YC7DK%yh%G2NP3%yh%2QQC3%yh%J6H88%yh%GVGXT" &::  Professional Plus
exit /b

:b13afb38-cd79-4ae5-9f7f-eed058d750ca
set "_key=KBKQT%yh%2NMXY%yh%JJWGP%yh%M62JB%yh%92CD4" &::  Standard
exit /b

:4a5d124a-e620-44ba-b6ff-658961b33b9a
set "_key=FN8TT%yh%7WMH6%yh%2D4X9%yh%M337T%yh%2342K" &::  Project Professional
exit /b

:427a28d1-d17c-4abf-b717-32c780ba6f07
set "_key=6NTH3%yh%CW976%yh%3G3Y2%yh%JK3TX%yh%8QHTT" &::  Project Standard
exit /b

:e13ac10e-75d0-4aff-a0cd-764982cf541c
set "_key=C2FG9%yh%N6J68%yh%H8BTJ%yh%BW3QX%yh%RM3B3" &::  Visio Professional
exit /b

:ac4efaf0-f81f-4f61-bdf7-ea32b02ab117
set "_key=J484Y%yh%4NKBF%yh%W2HMG%yh%DBMJC%yh%PGWR7" &::  Visio Standard
exit /b

:6ee7622c-18d8-4005-9fb7-92db644a279b
set "_key=NG2JY%yh%H4JBT%yh%HQXYP%yh%78QH9%yh%4JM2D" &::  Access
exit /b

:f7461d52-7c2b-43b2-8744-ea958e0bd09a
set "_key=VGPNG%yh%Y7HQW%yh%9RHP7%yh%TKPV3%yh%BG7GB" &::  Excel
exit /b

:fb4875ec-0c6b-450f-b82b-ab57d8d1677f
set "_key=H7R7V%yh%WPNXQ%yh%WCYYC%yh%76BGV%yh%VT7GH" &::  Groove
exit /b

:a30b8040-d68a-423f-b0b5-9ce292ea5a8f
set "_key=DKT8B%yh%N7VXH%yh%D963P%yh%Q4PHY%yh%F8894" &::  InfoPath
exit /b

:1b9f11e3-c85c-4e1b-bb29-879ad2c909e3
set "_key=2MG3G%yh%3BNTT%yh%3MFW9%yh%KDQW3%yh%TCK7R" &::  Lync
exit /b

:efe1f3e6-aea2-4144-a208-32aa872b6545
set "_key=TGN6P%yh%8MMBC%yh%37P2F%yh%XHXXK%yh%P34VW" &::  OneNote
exit /b

:771c3afa-50c5-443f-b151-ff2546d863a0
set "_key=QPN8Q%yh%BJBTJ%yh%334K3%yh%93TGY%yh%2PMBT" &::  Outlook
exit /b

:8c762649-97d1-4953-ad27-b7e2c25b972e
set "_key=4NT99%yh%8RJFH%yh%Q2VDH%yh%KYG2C%yh%4RD4F" &::  Powerpoint
exit /b

:00c79ff1-6850-443d-bf61-71cde0de305f
set "_key=PN2WF%yh%29XG2%yh%T9HJ7%yh%JQPJR%yh%FCXK4" &::  Publisher
exit /b

:d9f5b1c6-5386-495a-88f9-9ad6b41ac9b3
set "_key=6Q7VD%yh%NX8JD%yh%WJ2VH%yh%88V73%yh%4GBJ7" &::  Word
exit /b

::  Office 2010
:09ed9640-f020-400a-acd8-d7d867dfd9c2
set "_key=YBJTT%yh%JG6MD%yh%V9Q7P%yh%DBKXJ%yh%38W9R" &::  Mondo
exit /b

:ef3d4e49-a53d-4d81-a2b1-2ca6c2556b2c
set "_key=7TC2V%yh%WXF6P%yh%TD7RT%yh%BQRXR%yh%B8K32" &::  Mondo2
exit /b

:6f327760-8c5c-417c-9b61-836a98287e0c
set "_key=VYBBJ%yh%TRJPB%yh%QFQRF%yh%QFT4D%yh%H3GVB" &::  Professional Plus
exit /b

:9da2a678-fb6b-4e67-ab84-60dd6a9c819a
set "_key=V7QKV%yh%4XVVR%yh%XYV4D%yh%F7DFM%yh%8R6BM" &::  Standard
exit /b

:df133ff7-bf14-4f95-afe3-7b48e7e331ef
set "_key=YGX6F%yh%PGV49%yh%PGW3J%yh%9BTGG%yh%VHKC6" &::  Project Professional
exit /b

:5dc7bf61-5ec9-4996-9ccb-df806a2d0efe
set "_key=4HP3K%yh%88W3F%yh%W2K3D%yh%6677X%yh%F9PGB" &::  Project Standard
exit /b

:92236105-bb67-494f-94c7-7f7a607929bd
set "_key=D9DWC%yh%HPYVV%yh%JGF4P%yh%BTWQB%yh%WX8BJ" &::  Visio Premium
exit /b

:e558389c-83c3-4b29-adfe-5e4d7f46c358
set "_key=7MCW8%yh%VRQVK%yh%G677T%yh%PDJCM%yh%Q8TCP" &::  Visio Professional
exit /b

:9ed833ff-4f92-4f36-b370-8683a4f13275
set "_key=767HD%yh%QGMWX%yh%8QTDB%yh%9G3R2%yh%KHFGJ" &::  Visio Standard
exit /b

:8ce7e872-188c-4b98-9d90-f8f90b7aad02
set "_key=V7Y44%yh%9T38C%yh%R2VJK%yh%666HK%yh%T7DDX" &::  Access
exit /b

:cee5d470-6e3b-4fcc-8c2b-d17428568a9f
set "_key=H62QG%yh%HXVKF%yh%PP4HP%yh%66KMR%yh%CW9BM" &::  Excel
exit /b

:8947d0b8-c33b-43e1-8c56-9b674c052832
set "_key=QYYW6%yh%QP4CB%yh%MBV6G%yh%HYMCJ%yh%4T3J4" &::  Groove - SharePoint Workspace
exit /b

:ca6b6639-4ad6-40ae-a575-14dee07f6430
set "_key=K96W8%yh%67RPQ%yh%62T9Y%yh%J8FQJ%yh%BT37T" &::  InfoPath
exit /b

:ab586f5c-5256-4632-962f-fefd8b49e6f4
set "_key=Q4Y4M%yh%RHWJM%yh%PY37F%yh%MTKWH%yh%D3XHX" &::  OneNote
exit /b

:ecb7c192-73ab-4ded-acf4-2399b095d0cc
set "_key=7YDC2%yh%CWM8M%yh%RRTJC%yh%8MDVC%yh%X3DWQ" &::  Outlook
exit /b

:45593b1d-dfb1-4e91-bbfb-2d5d0ce2227a
set "_key=RC8FX%yh%88JRY%yh%3PF7C%yh%X8P67%yh%P4VTT" &::  Powerpoint
exit /b

:b50c4f75-599b-43e8-8dcd-1081a7967241
set "_key=BFK7F%yh%9MYHM%yh%V68C7%yh%DRQ66%yh%83YTP" &::  Publisher
exit /b

:2d0882e7-a4e7-423b-8ccc-70d91e0158b1
set "_key=HVHB3%yh%C6FV7%yh%KQX9W%yh%YQG79%yh%CRY7T" &::  Word
exit /b

:ea509e87-07a1-4a45-9edc-eba5a39f36af
set "_key=D6QFG%yh%VBYP2%yh%XQHM7%yh%J97RH%yh%VVRCK" &::  Small Business Basics
exit /b

:TheEnd

if %act_failed% EQU 1 (
echo ____________________________________________________________________
echo.
call :_errorinfo
)

if not defined _tskinstalled if not defined _oldtsk (
echo.
if %winbuild% GEQ 9200 (
call :leavenonexistentkms %nul%
echo 将不存在的 IP 地址 10.0.0.10 保留作为 KMS 服务器。
) else (
call :Clear-KMS-Cache
)
)

if not [%Act_OK%]==[1] (
echo.
echo 如果有任何问题，请查看 https://mass%-%grave.dev/troubleshoot
)

if defined _unattended exit /b

echo ____________________________________________________________________
echo.
call :_color %_Yellow% "请按任意键返回……"
pause >nul
exit /b

::========================================================================================================================================

:_errorinfo

call :CheckFR

set _intcon=
for %%a in (l.root-servers.net resolver1.opendns.com download.windowsupdate.com google.com) do if not defined _intcon (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (if not [%%#]==[] set _intcon=1)
)

if not defined _intcon (
call :_color %_Red% "Internet 未连接。"
exit /b
)

if [%ERRORCODE%]==[-1073418124] (
echo 正在检查端口 1688 连接，它将需要一段时间……
echo.

set /a count=0
set _portcon=
for %%a in (%srvlist%) do if not defined _portcon if !count! LEQ 7 (
set /a count+=1
%psc% "$t = New-Object Net.Sockets.TcpClient;try{$t.Connect("""%%a""", 1688)}catch{};$t.Connected" | findstr /i true 1>nul && set _portcon=1
)

if not defined _portcon (
call :_color %Red% "端口 1688 在你的 Internet 连接中被阻止。"
echo.
echo 原因：    有可能连接了受限的 Internet [办公室/院校]，
echo           或者防火墙阻止了连接。
echo.
echo 解决方案：使用其他 Internet 连接或使用脱机 KMS
echo           https://github.com/abbodi1406/KMS_VL_ALL_AIO
) else (
echo 已通过 KMS 服务器端口 1688 测试。
echo.
echo 请确保系统文件未在你的防火墙中阻止。
echo 如果问题仍然存在，请尝试使用脱机 KMS 激活程序
echo https://github.com/abbodi1406/KMS_VL_ALL_AIO
)
echo.
)

echo 在这种情况下，KMS 服务器并非问题。
exit /b

::========================================================================================================================================

:setserv

::  多 KMS 服务器集成和服务器随机化

set srvlist=
set -=

set "srvlist=kms.zhu%-%xiaole.org kms-default.cangs%-%hui.net kms.six%-%yin.com kms.moe%-%club.org kms.cgt%-%soft.com"
set "srvlist=%srvlist% kms.id%-%ina.cn kms.moe%-%yuuko.com xinch%-%eng213618.cn kms.wl%-%rxy.cn kms.ca%-%tqu.com"
set "srvlist=%srvlist% kms.0%-%t.net.cn kms.its%-%jzx.com kms.wx%-%lost.com kms.moe%-%yuuko.top kms.gh%-%pym.com"

set n=1
for %%a in (%srvlist%) do (set %%a=&set server!n!=%%a&set /a n+=1)
set max_servers=15
set /a server_num=0
exit /b

:getserv

if %server_num% equ %max_servers% set /a server_num+=1&set KMS_IP=222.184.9.98&exit /b
set /a rand=%Random%%%(15+1-1)+1
if defined !server%rand%! goto :getserv
set KMS_IP=!server%rand%!
set !server%rand%!=1

::  获取用于激活的 KMS 服务器的 IPv4 地址，即使禁用 ICMP 回显也可以工作。
::  如果直接使用公共 KMS 服务器主机名进行激活，Microsoft 和防病毒软件可能会标记此问题。

set /a server_num+=1
(for /f "delims=[] tokens=2" %%a in ('ping -4 -n 1 %KMS_IP% 2^>nul') do set "KMS_IP=%%a"
if [%KMS_IP%]==[!KMS_IP!] for /f "delims=[] tokens=2" %%# in ('pathping -4 -h 1 -n -p 1 -q 1 -w 1 %KMS_IP% 2^>nul') do set "KMS_IP=%%#"
if not [%KMS_IP%]==[!KMS_IP!] exit /b
goto :getserv
)

:==========================================================================================================================================

:Clear-KMS-Cache

set OPPk=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform
set SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform

set _wApp=55c92734-d682-4d71-983e-d6ec3f16059f
set _oApp=0ff1ce15-a989-479d-af46-f275c6370663
set _oA14=59a52881-a989-479d-af46-f275c6370663

%nul% reg delete "HKLM\%SPPk%" /f /v KeyManagementServiceName
%nul% reg delete "HKLM\%SPPk%" /f /v KeyManagementServicePort
%nul% reg delete "HKLM\%SPPk%" /f /v DisableDnsPublishing
%nul% reg delete "HKLM\%SPPk%" /f /v DisableKeyManagementServiceHostCaching
%nul% reg delete "HKLM\%SPPk%\%_wApp%" /f
if %winbuild% GEQ 9200 (
if defined notx86 (
%nul% reg delete "HKLM\%SPPk%" /f /v KeyManagementServiceName /reg:32
%nul% reg delete "HKLM\%SPPk%" /f /v KeyManagementServicePort /reg:32
%nul% reg delete "HKLM\%SPPk%\%_oApp%" /f /reg:32
)
%nul% reg delete "HKLM\%SPPk%\%_oApp%" /f
)
if %winbuild% GEQ 9600 (
%nul% reg delete "HKU\S-1-5-20\%SPPk%\%_wApp%" /f
%nul% reg delete "HKU\S-1-5-20\%SPPk%\%_oApp%" /f
)
%nul% reg delete "HKLM\%OPPk%" /f /v KeyManagementServiceName
%nul% reg delete "HKLM\%OPPk%" /f /v KeyManagementServicePort
%nul% reg delete "HKLM\%OPPk%" /f /v DisableDnsPublishing
%nul% reg delete "HKLM\%OPPk%" /f /v DisableKeyManagementServiceHostCaching
%nul% reg delete "HKLM\%OPPk%\%_oA14%" /f
%nul% reg delete "HKLM\%OPPk%\%_oApp%" /f

::  检查 KMS38 锁

%nul% reg query "HKLM\%SPPk%\%_wApp%" && (
set error_=9
echo 完全清除 KMS 缓存失败。
reg query "HKLM\%SPPk%\%_wApp%" /s 2>nul | findstr /i "127.0.0.2" >nul && echo KMS38 激活被锁定。
) || (
echo 已成功清除 KMS 缓存。
)
exit /b

:=========================================================================================================================================

:leavenonexistentkms

reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "10.0.0.10"
reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "1688"
reg delete "HKLM\%SPPk%" /f /v DisableDnsPublishing
reg delete "HKLM\%SPPk%" /f /v DisableKeyManagementServiceHostCaching
if not defined _keepkms38 reg delete "HKLM\%SPPk%\%_wApp%" /f
if %winbuild% GEQ 9200 (
if not %xOS%==x86 (
reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "10.0.0.10" /reg:32
reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" /reg:32
reg delete "HKLM\%SPPk%\%_oApp%" /f /reg:32
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "10.0.0.10" /reg:32
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" /reg:32
)
reg delete "HKLM\%SPPk%\%_oApp%" /f
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "10.0.0.10"
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "1688"
)
if %winbuild% GEQ 9600 (
reg delete "HKU\S-1-5-20\%SPPk%\%_wApp%" /f
reg delete "HKU\S-1-5-20\%SPPk%\%_oApp%" /f
)
reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "10.0.0.10"
reg delete "HKLM\%OPPk%" /f /v KeyManagementServicePort
reg delete "HKLM\%OPPk%" /f /v DisableDnsPublishing
reg delete "HKLM\%OPPk%" /f /v DisableKeyManagementServiceHostCaching
reg delete "HKLM\%OPPk%\%_oA14%" /f
reg delete "HKLM\%OPPk%\%_oApp%" /f
goto :eof

:=========================================================================================================================================

:_Complete_Uninstall

cls
mode con: cols=91 lines=30
title 在线 KMS 完全卸载 %masver%

set "key=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\taskcache\tasks"

set "_C16R="
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" 2^>nul') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" set "_C16R=1"
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath /reg:32" 2^>nul') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" set "_C16R=1"
if %winbuild% GEQ 9200 if defined _C16R (
echo.
echo ## 注  意 ##
echo.
echo 为确保 Office 程序不显示非正版横幅，
echo 请运行一次激活选项，之后不要卸载。
echo __________________________________________________________________________________________
)

set error_=
echo.
call :Clear-KMS-Cache
call :clearstuff

if defined error_ (
if [%error_%]==[1] (
echo __________________________________________________________________________________________
%eline%
echo 请再试一次或重启系统
echo __________________________________________________________________________________________
)
) else (
echo __________________________________________________________________________________________
echo.
call :_color %Green% "在线 KMS 完整卸载已成功完成。"
echo __________________________________________________________________________________________
)

if defined _unattended timeout /t 2 & exit /b

echo.
call :_color %_Yellow% "请按任意键返回……"
pause >nul
exit /b

:clearstuff

reg query "%key%" /f Path /s | find /i "\Activation-Renewal" >nul && (
echo 正在删除 [任务] Activation-Renewal
schtasks /delete /tn Activation-Renewal /f %nul%
)

reg query "%key%" /f Path /s | find /i "\Activation-Run_Once" >nul && (
echo 正在删除 [任务] Activation-Run_Once
schtasks /delete /tn Activation-Run_Once /f %nul%
)

reg query "%key%" /f Path /s | find /i "\Online_KMS_Activation_Script-Renewal" >nul && (
echo 正在删除 [任务] Online_KMS_Activation_Script-Renewal
schtasks /delete /tn Online_KMS_Activation_Script-Renewal /f %nul%
)

reg query "%key%" /f Path /s | find /i "\Online_KMS_Activation_Script-Run_Once" >nul && (
echo 正在删除 [任务] Online_KMS_Activation_Script-Run_Once
schtasks /delete /tn Online_KMS_Activation_Script-Run_Once /f %nul%
)

If exist "%windir%\Online_KMS_Activation_Script\" (
echo 正在删除 [文件夹] %windir%\Online_KMS_Activation_Script\
rmdir /s /q "%windir%\Online_KMS_Activation_Script\" %nul%
)

if exist "%ProgramData%\Online_KMS_Activation.cmd" (
echo 正在删除 [文件] %ProgramData%\Online_KMS_Activation.cmd
del /f /q "%ProgramData%\Online_KMS_Activation.cmd" %nul%
)

If exist "%ProgramData%\Online_KMS_Activation\" (
echo 正在删除 [文件夹] %ProgramData%\Online_KMS_Activation\
rmdir /s /q "%ProgramData%\Online_KMS_Activation\" %nul%
)

If exist "%ProgramData%\Activation-Renewal\" (
echo 正在删除 [文件夹] %ProgramData%\Activation-Renewal\
rmdir /s /q "%ProgramData%\Activation-Renewal\" %nul%
)

If exist "%ProgramFiles%\Activation-Renewal\" (
echo 正在删除 [文件夹] %ProgramFiles%\Activation-Renewal\
rmdir /s /q "%ProgramFiles%\Activation-Renewal\" %nul%
)

reg query "HKCR\DesktopBackground\shell\Activate Windows - Office" %nul% && (
echo 正在删除 [注册表] HKCR\DesktopBackground\shell\Activate Windows - Office
Reg delete "HKCR\DesktopBackground\shell\Activate Windows - Office" /f %nul%
)

reg query "%key%" /f Path /s | find /i "\Activation-Renewal" >nul && (set error_=1)
reg query "%key%" /f Path /s | find /i "\Activation-Run_Once" >nul && (set error_=1)
reg query "%key%" /f Path /s | find /i "\Online_KMS_Activation_Script" >nul && (set error_=1)
If exist "%windir%\Online_KMS_Activation_Script\" (set error_=1)
reg query "HKCR\DesktopBackground\shell\Activate Windows - Office" %nul% && (set error_=1)
if exist "%ProgramData%\Online_KMS_Activation.cmd" (set error_=1)
if exist "%ProgramData%\Online_KMS_Activation\" (set error_=1)
if exist "%ProgramData%\Activation-Renewal\" (set error_=1)
if exist "%ProgramFiles%\Activation-Renewal\" (set error_=1)
exit /b

:=========================================================================================================================================

:RenTask

cls
mode con cols=91 lines=30
title 安装激活自动续期 %masver%

set error_=
set "_dest=%ProgramFiles%\Activation-Renewal"
set "key=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\taskcache\tasks"

call :clearstuff %nul%

if defined error_ (
%eline%
echo 清除 KMS 相关文件夹/任务失败。
echo 请运行卸载选项，然后再试一次。
goto :RenDone
)

if not exist "%_dest%\" md "%_dest%\" %nul%
set "_temp=%SystemRoot%\Temp\_taskwork_%Random%"

set nil=
if exist "%_temp%\.*" rmdir /s /q "%_temp%\" %nul%
md "%_temp%\" %nul%
call :RenExport renewal "%_temp%\Renewal.xml" Unicode
if defined ActTask (call :RenExport run_once "%_temp%\Run_Once.xml" Unicode)
s%nil%cht%nil%asks /cre%nil%ate /tn "Activation-Renewal" /ru "SYS%nil%TEM" /xml "%_temp%\Renewal.xml" %nul%
if defined ActTask (s%nil%cht%nil%asks /cre%nil%ate /tn "Activation-Run_Once" /ru "SYS%nil%TEM" /xml "%_temp%\Run_Once.xml" %nul%)
if exist "%_temp%\.*" rmdir /s /q "%_temp%\" %nul%

call :createInfo.txt
%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split \":_extracttask\:.*`r`n\"; [io.file]::WriteAllText('%_dest%\Activation_task.cmd', '@REM Dummy ' + '%random%' + [Environment]::NewLine + $f[1].Trim(), [System.Text.Encoding]::ASCII);"
title 安装激活自动续期 %masver%

::========================================================================================================================================

reg query "%key%" /f Path /s | find /i "\Activation-Renewal" >nul || (set error_=1)
if defined ActTask reg query "%key%" /f Path /s | find /i "\Activation-Run_Once" >nul || (set error_=1)

If not exist "%_dest%\Activation_task.cmd" (set error_=1)
If not exist "%_dest%\Info.txt" (set error_=1)

if defined error_ (

reg query "%key%" /f Path /s | find /i "\Activation-Renewal" >nul && (
schtasks /delete /tn Activation-Renewal /f %nul%
)
reg query "%key%" /f Path /s | find /i "\Activation-Run_Once" >nul && (
schtasks /delete /tn Activation-Run_Once /f %nul%
)

If exist "%_dest%\" (
rmdir /s /q "%_dest%\" %nul%
)

%eline%
echo 请运行卸载选项，然后再试一次。
goto :RenDone
)

echo __________________________________________________________________________________________
echo.
echo 文件已创建：
echo %_dest%\Activation_task.cmd
echo %_dest%\Info.txt
echo.
(if defined ActTask (echo 计划任务已创建：) else (echo 计划任务已创建：))
echo \Activation-Renewal    [续期 / 每周]
if defined ActTask (echo \Activation-Run_Once   [激活任务 - 激活后自动删除])
echo __________________________________________________________________________________________
echo.
echo 信息：
echo 如果找到 Internet 连接，将在每周续期激活。
echo 它只会续期已安装的 KMS 许可。它不会将任何许可证转换为 KMS。
echo __________________________________________________________________________________________
echo.
if defined ActTask (
call :_color %Green% "已成功创建续期和激活任务。"
) else (
call :_color %Green% "已成功创建续期任务。"
)
echo.
call :_color %Gray% "请确保至少运行一次“激活”选项。"
echo __________________________________________________________________________________________
)

::========================================================================================================================================

:RenDone

if defined _unattended exit /b

echo.
call :_color %_Yellow% "请按任意键返回……"
pause >nul
exit /b

::========================================================================================================================================

:createInfo.txt

(
echo   此脚本的用途是使用在线 KMS 激活/更新你的 Windows/Office 许可证。
echo:
echo   如果创建了续订/激活计划任务，将会存在以下内容，
echo:
echo   - 计划任务
echo     Activation-Renewal    [续期 / 每周]
echo     Activation-Run_Once   [激活任务 - 激活后自动删除]
echo     计划任务仅在系统连接到 Internet 时运行。
echo:
echo   - 文件
echo     C:\Program Files\Activation-Renewal\Activation_task.cmd
echo     C:\Program Files\Activation-Renewal\Info.txt
echo     C:\Program Files\Activation-Renewal\Logs.txt
echo ______________________________________________________________________________________________
echo:
echo   在线 KMS 脚本是“Microsoft 激活脚本”[MAS] 项目中的一部分
echo:   
echo   主    页：mass grave[.]dev
echo      Email：windowsaddict@protonmail.com
)>"%_dest%\Info.txt"
exit /b

::========================================================================================================================================

:renewal:
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Source>Microsoft Corporation</Source>
    <Date>1999-01-01T12:00:00.34375</Date>
    <Author>WindowsAddict</Author>
    <Version>1.0</Version>
    <Description>在线 K-M-S 自动续期 - 每周任务</Description>
    <URI>\Activation-Renewal</URI>
    <SecurityDescriptor>D:P(A;;FA;;;SY)(A;;FA;;;BA)(A;;FRFX;;;LS)(A;;FRFW;;;S-1-5-80-123231216-2592883651-3715271367-3753151631-4175906628)(A;;FR;;;S-1-5-4)</SecurityDescriptor>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>1999-01-01T12:00:00</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByWeek>
        <DaysOfWeek>
          <Sunday />
        </DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="LocalSystem">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>true</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT10M</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT2M</Interval>
      <Count>3</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="LocalSystem">
    <Exec>
      <Command>%ProgramFiles%\Activation-Renewal\Activation_task.cmd</Command>
    <Arguments>Task</Arguments>
    </Exec>
  </Actions>
</Task>
:renewal:

:run_once:
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Source>Microsoft Corporation</Source>
    <Date>1999-01-01T12:00:00.34375</Date>
    <Author>WindowsAddict</Author>
    <Version>1.0</Version>
    <Description>在线 K-M-S 激活运行一次 - 在首次联网时运行并删除自身</Description>
    <URI>\Activation-Run_Once</URI>
    <SecurityDescriptor>D:P(A;;FA;;;SY)(A;;FA;;;BA)(A;;FRFX;;;LS)(A;;FRFW;;;S-1-5-80-123231216-2592883651-3715271367-3753151631-4175906628)(A;;FR;;;S-1-5-4)</SecurityDescriptor>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="LocalSystem">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>true</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT10M</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT2M</Interval>
      <Count>3</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="LocalSystem">
    <Exec>
      <Command>%ProgramFiles%\Activation-Renewal\Activation_task.cmd</Command>
    <Arguments>Task</Arguments>
    </Exec>
  </Actions>
</Task>
:run_once:

::========================================================================================================================================

::  从批处理脚本中解压文本，无字符和文件编码问题

:RenExport

%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split \":%~1\:.*`r`n\"; [io.file]::WriteAllText('%~2',$f[1].Trim(),[System.Text.Encoding]::%~3);"
exit /b

::========================================================================================================================================

:_extracttask:
@echo off

::   通过计划任务使用在线服务器更新 K-M-S 激活

::============================================================================
::
::   此脚本是“Microsoft 激活脚本”（MAS）项目中的一部分。
::
::   主    页：mass grave[.]dev
::      Email：windowsaddict@protonmail.com
::
::============================================================================


if not "%~1"=="Task" (
echo.
echo ====== 错误 ======
echo.
echo 此文件仅应由计划任务运行。
echo.
echo 请按任意键退出脚本
pause >nul
exit /b
)

::  设置路径变量，如果在系统中配置错误时会有所帮助

set "PATH=%SystemRoot%\System32;%SystemRoot%\System32\wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\;%LocalAppData%\Microsoft\WindowsApps\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "PATH=%SystemRoot%\Sysnative;%SystemRoot%\Sysnative\wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%LocalAppData%\Microsoft\WindowsApps\;%PATH%"
)

>nul fltmc || exit /b

::========================================================================================================================================

set _tserror=
set winbuild=1
set "nul=>nul 2>&1"
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
set psc=powershell.exe

set run_once=
set t_name=续期任务
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\taskcache\tasks" /f Path /s | find /i "\Activation-Run_Once" >nul && (
set run_once=1
set t_name=单次运行任务
)

set _wmic=0
for %%# in (wmic.exe) do @if not "%%~$PATH:#"=="" (
wmic path Win32_ComputerSystem get CreationClassName /value 2>nul | find /i "computersystem" 1>nul && set _wmic=1
)

setlocal EnableDelayedExpansion
if exist "%ProgramFiles%\Activation-Renewal\" call :_taskstart>>"%ProgramFiles%\Activation-Renewal\Logs.txt" & exit

::========================================================================================================================================

:_taskstart

echo.
echo %date%, %time%

set /a loop=1
set /a max_loop=4

call :_tasksetserv

:_intrepeat

::  检查 Internet 连接。即使禁用 ICMP 回显也可以工作。

for %%a in (%srvlist%) do (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (
if not [%%#]==[] goto _taskIntConnected
)
)

nslookup dns.msftncsi.com 2>nul | find "131.107.255.255" 1>nul
if [%errorlevel%]==[0] goto _taskIntConnected

if %loop%==%max_loop% (
set _tserror=1
goto _taskend
)

echo.
echo 错误：Internet 未连接
echo 等待 30 秒

timeout /t 30 >nul
set /a loop=%loop%+1
goto _intrepeat

:_taskIntConnected

::========================================================================================================================================

::  检查非 x86 Windows

set notx86=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PROCESSOR_ARCHITECTURE') do set arch=%%b
if /i not "%arch%"=="x86" set notx86=1

::========================================================================================================================================

set "OPPk=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"

set "slp=SoftwareLicensingProduct"
set "ospp=OfficeSoftwareProtectionProduct"

set "_wApp=55c92734-d682-4d71-983e-d6ec3f16059f"
set "_oApp=0ff1ce15-a989-479d-af46-f275c6370663"
set "_oA14=59a52881-a989-479d-af46-f275c6370663"

::========================================================================================================================================

::  从注册表中清除现有的 KMS 缓存 / 将端口值设置为 1688

%nul% reg delete "HKLM\%SPPk%" /f /v KeyManagementServiceName
%nul% reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "1688"
%nul% reg delete "HKLM\%SPPk%\%_wApp%" /f
if %winbuild% GEQ 9200 (
if defined notx86 (
%nul% reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" /reg:32
%nul% reg delete "HKLM\%SPPk%\%_oApp%" /f /reg:32
%nul% reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" /reg:32
)
%nul% reg delete "HKLM\%SPPk%\%_oApp%" /f
%nul% reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "1688"
)
if %winbuild% GEQ 9600 (
%nul% reg delete "HKU\S-1-5-20\%SPPk%\%_wApp%" /f
%nul% reg delete "HKU\S-1-5-20\%SPPk%\%_oApp%" /f
)
%nul% reg add "HKLM\%OPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "1688"
%nul% reg delete "HKLM\%OPPk%\%_oA14%" /f
%nul% reg delete "HKLM\%OPPk%\%_oApp%" /f

::========================================================================================================================================

::  检查 WMI 和 sppsvc 错误

set applist=
net start sppsvc /y %nul%
if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path %slp% where (ApplicationID='%_wApp%') get ID /VALUE" 2^>nul')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM %slp% WHERE ApplicationID=''%_wApp%''').Get()).ID ^| %% {echo ('ID='+$_)}" 2^>nul')"
%chkapp% do (if defined applist (call set "applist=!applist! %%a") else (call set "applist=%%a"))

if not defined applist (
set _tserror=1
if %_wmic% EQU 1 wmic path Win32_ComputerSystem get CreationClassName /value 2>nul | find /i "computersystem" 1>nul
if %_wmic% EQU 0 %psc% "Get-CIMInstance -Class Win32_ComputerSystem | Select-Object -Property CreationClassName" 2>nul | find /i "computersystem" 1>nul
if !errorlevel! NEQ 0 (set e_wmispp=WMI, SPP) else (set e_wmispp=SPP)
echo.
echo 错误：没有响应—— !e_wmispp!
echo.
)

::========================================================================================================================================

::  检查已安装的批量产品激活 ID

call :_taskgetids sppwid %slp% windows
call :_taskgetids sppoid %slp% office
call :_taskgetids osppid %ospp% office

::========================================================================================================================================

echo.
echo 正在为所有已安装的批量产品更新 KMS 激活

if not defined sppwid if not defined sppoid if not defined osppid (
echo.
echo 未找到已安装的批量 Windows / Office 产品
echo.
echo 正在更新 KMS 服务器
call :_taskgetserv
call :_taskregserv
goto :_skipact
)

::========================================================================================================================================

::  检查 KMS38 激活

set gpr=0
set _kms38=0
if defined sppwid if %winbuild% GEQ 14393 (
set _path=%slp%
set _actid=%sppwid%
call :_taskgetgrace
)

if %gpr% NEQ 0 if %gpr% GTR 259200 (
set _kms38=1
call :_taskchkEnterpriseG _kms38
)

::  将特定 KMS 主机设置为本地主机，使全局 KMS IP 不能代替 KMS38 激活，但可以与 Office 和其他 Windows 版本一起使用。

if %_kms38% EQU 1 (
%nul% reg add "HKLM\%SPPk%\%_wApp%\%sppwid%" /f /v KeyManagementServiceName /t REG_SZ /d "127.0.0.2"
%nul% reg add "HKLM\%SPPk%\%_wApp%\%sppwid%" /f /v KeyManagementServicePort /t REG_SZ /d "1688"
)

::========================================================================================================================================

echo.
if defined sppwid (
set _path=%slp%
set _actid=%sppwid%
call :_actprod
call :_act act_win
call :_actinfo act_win
) else (
echo 正在检查：未安装批量版 Windows
)

if defined sppoid (
set _path=%slp%
for %%# in (%sppoid%) do (
echo.
set _actid=%%#
call :_actprod
call :_act
call :_actinfo
)
)

if defined osppid (
set _path=%ospp%
for %%# in (%osppid%) do (
echo.
set _actid=%%#
call :_actprod
call :_act
call :_actinfo
)
)

if not defined sppoid if not defined osppid (
echo.
echo 正在检查：未安装批量版 Office
)

:_skipact

::========================================================================================================================================

if defined run_once (
echo.
echo 正在删除计划任务 Activation-Run_Once
schtasks /delete /tn Activation-Run_Once /f %nul%
)

::========================================================================================================================================

:_taskend

echo.
echo 正在退出
echo ______________________________________________________________________

if defined _tserror (exit /b 123456789) else (exit /b 0)

::========================================================================================================================================

:_act

set errorcode=12345
set /a act_attempt=0

:_act2

if %act_attempt% GTR 4 exit /b

if not [%act_ok%]==[1] (
call :_taskgetserv
call :_taskregserv
)

if not !server_num! GTR %max_servers% (

if [%1]==[act_win] if %_kms38% EQU 1 (
set act_ok=1
exit /b
)

if %_wmic% EQU 1 wmic path !_path! where ID='!_actid!' call Activate %nul%
if %_wmic% EQU 0 %psc% "try {$null=(([WMISEARCHER]'SELECT ID FROM !_path! where ID=''!_actid!''').Get()).Activate(); exit 0} catch { exit $_.Exception.InnerException.HResult }"

call set errorcode=!errorlevel!

if !errorcode! EQU 0 (
set act_ok=1
exit /b
)
if [%1]==[act_win] if !errorcode! EQU -1073418187 if %winbuild% LSS 9200 (
set act_ok=1
exit /b
)

set act_ok=0
set /a act_attempt+=1
goto _act2
)
exit /b

:_actprod

if %_wmic% EQU 1 for /f "tokens=2 delims==" %%x in ('"wmic path !_path! where ID='!_actid!' get Name /VALUE" 2^>nul') do call echo 正在激活：%%x
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%x in ('%psc% "(([WMISEARCHER]'SELECT Name FROM !_path! WHERE ID=''!_actid!''').Get()).Name | %% {echo ('Name='+$_)}" 2^>nul') do call echo 正在激活：%%x
exit /b

::========================================================================================================================================

:_actinfo

if [%1]==[act_win] if %_kms38% EQU 1 (
echo Windows 已通过 KMS38 激活
exit /b
)

if %errorcode% EQU 12345 (
echo 产品激活失败
echo 由于受限或没有 Internet，无法测试 KMS 服务器
set _tserror=1
exit /b
)

if %errorcode% EQU -1073418187 (
echo 产品激活失败：0xC004F035
if [%1]==[act_win] if %winbuild% LSS 9200 echo 由于 OEM BIOS 不合格，无法在此计算机上激活 Windows 7
exit /b
)

if %errorcode% EQU -1073417728 (
echo 产品激活失败：0xC004F200
echo Windows 需要重建激活相关文件。
set _tserror=1
exit /b
)

set gpr=0
set gpr2=0
call :_taskgetgrace
set /a "gpr2=(%gpr%+1440-1)/1440"

if %errorcode% EQU 0 if %gpr% EQU 0 (
echo 产品激活成功，但剩余期限增加失败。
if [%1]==[act_win] if %winbuild% LSS 9200 echo 这可能与 KB4487266 中描述的错误有关
set _tserror=1
exit /b
)

set _actpass=1
if %gpr% EQU 43200  if [%1]==[act_win] if %winbuild% GEQ 9200 set _actpass=0
if %gpr% EQU 64800  set _actpass=0
if %gpr% GTR 259200 if [%1]==[act_win] call :_taskchkEnterpriseG _actpass
if %gpr% EQU 259200 set _actpass=0

if %errorcode% EQU 0 if %_actpass% EQU 0 (
echo 产品激活成功
echo 剩余期限：%gpr2% 天（%gpr% 分钟）
exit /b
)

cmd /c exit /b %errorcode%
if %errorcode% NEQ 0 (
echo 产品激活失败：0x!=ExitCode!
) else (
echo 产品激活失败
)
echo 剩余期限：%gpr2% 天（%gpr% 分钟）
set _tserror=1
exit /b

::========================================================================================================================================

:_taskgetids

set %1=
if %_wmic% EQU 1 set "chkapp=for /f "tokens=2 delims==" %%a in ('"wmic path %2 where (Name like '%%%3%%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get ID /VALUE" 2^>nul')"
if %_wmic% EQU 0 set "chkapp=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT ID FROM %2 WHERE Name like ''%%%3%%'' and Description like ''%%KMSCLIENT%%'' and PartialProductKey is not NULL').Get()).ID ^| %% {echo ('ID='+$_)}" 2^>nul')"
%chkapp% do (if defined %1 (call set "%1=!%1! %%a") else (call set "%1=%%a"))
exit /b

:_taskgetgrace

set gpr=0
if %_wmic% EQU 1 for /f "tokens=2 delims==" %%# in ('"wmic path !_path! where ID='!_actid!' get GracePeriodRemaining /VALUE" 2^>nul') do call set "gpr=%%#"
if %_wmic% EQU 0 for /f "tokens=2 delims==" %%# in ('%psc% "(([WMISEARCHER]'SELECT GracePeriodRemaining FROM !_path! where ID=''!_actid!''').Get()).GracePeriodRemaining | %% {echo ('GracePeriodRemaining='+$_)}" 2^>nul') do call set "gpr=%%#"
exit /b

:_taskchkEnterpriseG

for %%# in (e0b2d383-d112-413f-8a80-97f373a5820c e38454fb-41a4-4f59-a5dc-25080e354730) do (if %sppwid%==%%# set %1=0)
exit /b

::========================================================================================================================================

:_taskregserv

%nul% reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%"
%nul% reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%"

if %winbuild% GEQ 9200 (
%nul% reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%"
if defined notx86 (
%nul% reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" /reg:32
%nul% reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" /reg:32
)
)
exit /b

::========================================================================================================================================

:_tasksetserv

::  多 KMS 服务器集成和服务器随机化

set srvlist=
set -=

set "srvlist=kms.zhu%-%xiaole.org kms-default.cangs%-%hui.net kms.six%-%yin.com kms.moe%-%club.org kms.cgt%-%soft.com"
set "srvlist=%srvlist% kms.id%-%ina.cn kms.moe%-%yuuko.com xinch%-%eng213618.cn kms.wl%-%rxy.cn kms.ca%-%tqu.com"
set "srvlist=%srvlist% kms.0%-%t.net.cn kms.its%-%jzx.com kms.wx%-%lost.com kms.moe%-%yuuko.top kms.gh%-%pym.com"

set n=1
for %%a in (%srvlist%) do (set %%a=&set server!n!=%%a&set /a n+=1)
set max_servers=15
set /a server_num=0
exit /b

:_taskgetserv

if %server_num% geq %max_servers% (set /a server_num+=1&set KMS_IP=222.184.9.98&exit /b)
set /a rand=%Random%%%(15+1-1)+1
if defined !server%rand%! goto :_taskgetserv
set KMS_IP=!server%rand%!
set !server%rand%!=1

::  获取用于激活的 KMS 服务器的 IPv4 地址，即使禁用 ICMP 回显也可以工作。
::  如果直接使用公共 KMS 服务器主机名进行激活，Microsoft 和防病毒软件可能会标记此问题。

set /a server_num+=1
(for /f "delims=[] tokens=2" %%a in ('ping -4 -n 1 %KMS_IP% 2^>nul') do set "KMS_IP=%%a"
if [%KMS_IP%]==[!KMS_IP!] for /f "delims=[] tokens=2" %%# in ('pathping -4 -h 1 -n -p 1 -q 1 -w 1 %KMS_IP% 2^>nul') do set "KMS_IP=%%#"
if not [%KMS_IP%]==[!KMS_IP!] exit /b
goto :_taskgetserv
)

::  Ver:1.9
::========================================================================================================================================
:_extracttask:

:======================================================================================================================================================

:_color

if %_NCS% EQU 1 (
if defined _unattended (echo %~2) else (echo %esc%[%~1%~2%esc%[0m)
) else (
if defined _unattended (echo %~2) else (call :batcol %~1 "%~2")
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

::  https://gist.github.com/ave9858/9fff6af726ba3ddc646285d1bbf37e71
::  此代码用于清理 Office 许可

:cleanlicense:
function UninstallLicenses($DllPath) {
    $AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1)
    $ModuleBuilder = $AssemblyBuilder.DefineDynamicModule(2, $False)
    $TypeBuilder = $ModuleBuilder.DefineType(0)
    
    [void]$TypeBuilder.DefinePInvokeMethod('SLOpen', $DllPath, 'Public, Static', 1, [int], @([IntPtr].MakeByRefType()), 1, 3)
    [void]$TypeBuilder.DefinePInvokeMethod('SLGetSLIDList', $DllPath, 'Public, Static', 1, [int],
        @([IntPtr], [int], [Guid].MakeByRefType(), [int], [int].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
    [void]$TypeBuilder.DefinePInvokeMethod('SLUninstallLicense', $DllPath, 'Public, Static', 1, [int], @([IntPtr], [IntPtr]), 1, 3)

    $SPPC = $TypeBuilder.CreateType()
    $Handle = 0
    [void]$SPPC::SLOpen([ref]$Handle)
    $pnReturnIds = 0
    $ppReturnIds = 0

    if (!$SPPC::SLGetSLIDList($Handle, 0, [ref][Guid]"0ff1ce15-a989-479d-af46-f275c6370663", 6, [ref]$pnReturnIds, [ref]$ppReturnIds)) {
        foreach ($i in 0..($pnReturnIds - 1)) {
            [void]$SPPC::SLUninstallLicense($Handle, [Int64]$ppReturnIds + [Int64]16 * $i)
        }    
    }
}

$OSPP = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform" -ErrorAction SilentlyContinue).Path
if ($OSPP) {
    Write-Output "发现 Office 软件保护已安装，正在清理"
    UninstallLicenses($OSPP + "osppc.dll")
}
UninstallLicenses("sppc.dll")
:cleanlicense:

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:_Check_Status_vbs
@setlocal DisableDelayedExpansion
@echo off
@cls
color 07
title 检查激活状态 [vbs]
set "SysPath=%SystemRoot%\System32"
set "Path=%SystemRoot%\System32;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\;%LocalAppData%\Microsoft\WindowsApps\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "SysPath=%SystemRoot%\Sysnative"
set "Path=%SystemRoot%\Sysnative;%SystemRoot%\Sysnative\Wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%LocalAppData%\Microsoft\WindowsApps\;%Path%"
)

set ohook=
for %%# in (15 16) do (
for %%A in ("%ProgramFiles%" "%ProgramW6432%" "%ProgramFiles(x86)%") do (
if exist "%%~A\Microsoft Office\Office%%#\sppc*dll" set ohook=1
)
)

for %%# in (System SystemX86) do (
for %%G in ("Office 15" "Office") do (
for %%A in ("%ProgramFiles%" "%ProgramW6432%" "%ProgramFiles(x86)%") do (
if exist "%%~A\Microsoft %%~G\root\vfs\%%#\sppc*dll" set ohook=1
)
)
)

set "_bit=64"
set "_wow=1"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if "%PROCESSOR_ARCHITEW6432%"=="" set "_wow=0"&set "_bit=32"
set "_utemp=%TEMP%"
set "line2=************************************************************"
set "line3=____________________________________________________________"
set _sO16vbs=0
set _sO15vbs=0
if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" (
  set _sO15vbs=1
) else if exist "%ProgramW6432%\Microsoft Office\Office15\ospp.vbs" (
  set _sO15vbs=1
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" (
  set _sO15vbs=1
)
setlocal EnableDelayedExpansion
echo %line2%
echo ***                  Windows 激活状态                    ***
echo %line2%
pushd "!_utemp!"
copy /y %SystemRoot%\System32\slmgr.vbs . >nul 2>&1
net start sppsvc /y >nul 2>&1
cscript //nologo slmgr.vbs /dli || (echo 执行 slmgr.vbs 出错&del /f /q slmgr.vbs&popd&goto :casVend)
cscript //nologo slmgr.vbs /xpr
del /f /q slmgr.vbs >nul 2>&1
popd
echo %line3%

if defined ohook (
echo.
echo.
echo %line2%
echo ***                Office Ohook 激活状态                 ***
echo %line2%
echo.
powershell "write-host -back 'Black' -fore 'Yellow' '已使用 Ohook 永久激活 Office。'; write-host -back 'Black' -fore 'Yellow' '你可以忽略以下的 Office 激活状态。'"
echo.
)

:casVo16
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\ospp.vbs" (
set _sO16vbs=1
echo.
echo %line2%
if %_sO15vbs% EQU 0 (
echo ***              Office 2016 %_bit%-位激活状态               ***
) else (
echo ***              Office 2013/2016 激活状态               ***
)
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)
if %_wow%==0 goto :casVo13
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\ospp.vbs" (
set _sO16vbs=1
echo.
echo %line2%
if %_sO15vbs% EQU 0 (
echo ***              Office 2016 32-位激活状态               ***
) else (
echo ***              Office 2013/2016 激活状态               ***
)
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)

:casVo13
if %_sO16vbs% EQU 1 goto :casVo10
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\ospp.vbs" (
echo.
echo %line2%
echo ***              Office 2013 %_bit%-位激活状态               ***
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)
if %_wow%==0 goto :casVo10
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\ospp.vbs" (
echo.
echo %line2%
echo ***              Office 2013 32-位激活状态               ***
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)

:casVo10
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\ospp.vbs" (
echo.
echo %line2%
echo ***              Office 2010 %_bit%-位激活状态               ***
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)
if %_wow%==0 goto :casVc16
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\ospp.vbs" (
echo.
echo %line2%
echo ***              Office 2010 32-位激活状态               ***
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)

:casVc16
reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath >nul 2>&1 || (
reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v InstallPath >nul 2>&1 || goto :casVc13
)
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" 2^>nul') do (set "office=%%b\Office16")
if exist "!office!\ospp.vbs" (
set _sO16vbs=1
echo.
echo %line2%
if %_sO15vbs% EQU 0 (
echo ***             Office 2016-2021 C2R 激活状态            ***
) else (
echo ***               Office 2013-2021 激活状态              ***
)
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)
if %_wow%==0 goto :casVc13
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v InstallPath" 2^>nul') do (set "office=%%b\Office16")
if exist "!office!\ospp.vbs" (
set _sO16vbs=1
echo.
echo %line2%
if %_sO15vbs% EQU 0 (
echo ***             Office 2016-2021 C2R 激活状态            ***
) else (
echo ***               Office 2013-2021 激活状态              ***
)
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)

:casVc13
if %_sO16vbs% EQU 1 goto :casVc10
reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath >nul 2>&1 || (
reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun /v InstallPath >nul 2>&1 || goto :casVc10
)
set office=
if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" (
  set "office=%ProgramFiles%\Microsoft Office\Office15"
) else if exist "%ProgramW6432%\Microsoft Office\Office15\ospp.vbs" (
  set "office=%ProgramW6432%\Microsoft Office\Office15"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" (
  set "office=%ProgramFiles(x86)%\Microsoft Office\Office15"
)
if exist "!office!\ospp.vbs" (
echo.
echo %line2%
echo ***               Office 2013 C2R 激活状态               ***
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)

:casVc10
if %_wow%==0 reg query HKLM\SOFTWARE\Microsoft\Office\14.0\CVH /f Click2run /k >nul 2>&1 || goto :casVend
if %_wow%==1 reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\CVH /f Click2run /k >nul 2>&1 || goto :casVend
set office=
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" (
  set "office=%ProgramFiles%\Microsoft Office\Office14"
) else if exist "%ProgramW6432%\Microsoft Office\Office14\ospp.vbs" (
  set "office=%ProgramW6432%\Microsoft Office\Office14"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" (
  set "office=%ProgramFiles(x86)%\Microsoft Office\Office14"
)
if exist "!office!\ospp.vbs" (
echo.
echo %line2%
echo ***               Office 2010 C2R 激活状态               ***
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)

:casVend
echo.
call :_color %_Yellow% "请按任意键返回……"
pause >nul
exit /b

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:_Check_Status_wmi

@setlocal DisableDelayedExpansion
@echo off
color 07
title 检查激活状态 [wmi]

set WMI_VBS=0
@cls
set "_cmdf=%~f0"
set wspp=SoftwareLicensingProduct
set wsps=SoftwareLicensingService
set ospp=OfficeSoftwareProtectionProduct
set osps=OfficeSoftwareProtectionService
set winApp=55c92734-d682-4d71-983e-d6ec3f16059f
set o14App=59a52881-a989-479d-af46-f275c6370663
set o15App=0ff1ce15-a989-479d-af46-f275c6370663
for %%# in (spp_get,ospp_get,cW1nd0ws,sppw,c0ff1ce15,sppo,osppsvc,ospp14,ospp15) do set "%%#="
for /f "tokens=6 delims=[]. " %%# in ('ver') do set winbuild=%%#
set "spp_get=Description, DiscoveredKeyManagementServiceMachineName, DiscoveredKeyManagementServiceMachinePort, EvaluationEndDate, GracePeriodRemaining, ID, KeyManagementServiceMachine, KeyManagementServicePort, KeyManagementServiceProductKeyID, LicenseStatus, LicenseStatusReason, Name, PartialProductKey, ProductKeyID, VLActivationInterval, VLRenewalInterval"
set "ospp_get=%spp_get%"
if %winbuild% GEQ 9200 set "spp_get=%spp_get%, KeyManagementServiceLookupDomain, VLActivationTypeEnabled"
if %winbuild% GEQ 9600 set "spp_get=%spp_get%, DiscoveredKeyManagementServiceMachineIpAddress, ProductKeyChannel"
set "_work=%~dp0"
set "_batf=%~f0"
set "_batp=%_batf:'=''%"
set "_Local=%LocalAppData%"
set _Identity=0
setlocal EnableDelayedExpansion
dir /b /s /a:-d "!_Local!\Microsoft\Office\Licenses\*1*" 1>nul 2>nul && set _Identity=1
dir /b /s /a:-d "!ProgramData!\Microsoft\Office\Licenses\*1*" 1>nul 2>nul && set _Identity=1
pushd "!_work!"
setlocal DisableDelayedExpansion
if %winbuild% LSS 9200 if not exist "%SystemRoot%\servicing\Packages\Microsoft-Windows-PowerShell-WTR-Package~*.mum" set _Identity=0

set "SysPath=%SystemRoot%\System32"
set "Path=%SystemRoot%\System32;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\;%LocalAppData%\Microsoft\WindowsApps\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "SysPath=%SystemRoot%\Sysnative"
set "Path=%SystemRoot%\Sysnative;%SystemRoot%\Sysnative\Wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%LocalAppData%\Microsoft\WindowsApps\;%Path%"
)

set ohook=
for %%# in (15 16) do (
for %%A in ("%ProgramFiles%" "%ProgramW6432%" "%ProgramFiles(x86)%") do (
if exist "%%~A\Microsoft Office\Office%%#\sppc*dll" set ohook=1
)
)

for %%# in (System SystemX86) do (
for %%G in ("Office 15" "Office") do (
for %%A in ("%ProgramFiles%" "%ProgramW6432%" "%ProgramFiles(x86)%") do (
if exist "%%~A\Microsoft %%~G\root\vfs\%%#\sppc*dll" set ohook=1
)
)
)

set _cwmi=0
for %%# in (wmic.exe) do @if not "%%~$PATH:#"=="" (
wmic path Win32_ComputerSystem get CreationClassName /value 2>nul | find /i "ComputerSystem" 1>nul && set _cwmi=1
)

if %_cwmi% EQU 0 (
echo:
echo 错误：WMI 在系统中未响应。
echo:
echo 在 MAS 中，请转到疑难解答并运行修复 WMI 选项。
echo:
echo 请按任意键返回……
pause >nul
exit /b
)

set "line2=************************************************************"
set "line3=____________________________________________________________"
set "_psc=powershell"

set _prsh=1
for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" set _prsh=0
set "_csg=cscript.exe //NoLogo //Job:WmiMulti "%~nx0?.wsf""
set "_csq=cscript.exe //NoLogo //Job:WmiQuery "%~nx0?.wsf""
set "_csx=cscript.exe //NoLogo //Job:XPDT "%~nx0?.wsf""
if %_cwmi% EQU 0 set WMI_VBS=1
if %WMI_VBS% EQU 0 (
set "_zz1=wmic path"
set "_zz2=where"
set "_zz3=get"
set "_zz4=/value"
set "_zz5=("
set "_zz6=)"
set "_zz7="wmic path"
set "_zz8=/value""
) else (
set "_zz1=%_csq%"
set "_zz2="
set "_zz3="
set "_zz4="
set "_zz5=""
set "_zz6=""
set "_zz7=%_csq%"
set "_zz8="
)
set _WSH=0
set OsppHook=1
sc query osppsvc >nul 2>&1
if %errorlevel% EQU 1060 set OsppHook=0

net start sppsvc /y >nul 2>&1
call :casWpkey %wspp% %winApp% cW1nd0ws sppw
if %winbuild% GEQ 9200 call :casWpkey %wspp% %o15App% c0ff1ce15 sppo
if %OsppHook% NEQ 0 (
net start osppsvc /y >nul 2>&1
call :casWpkey %ospp% %o14App% osppsvc ospp14
if %winbuild% LSS 9200 call :casWpkey %ospp% %o15App% osppsvc ospp15
)

echo %line2%
echo ***                  Windows 激活状态                    ***
echo %line2%
if not defined cW1nd0ws (
echo.
echo 错误：未找到产品密钥。
goto :casWcon
)
set winID=1
set "_qr=%_zz7% %wspp% %_zz2% %_zz5%ApplicationID='%winApp%' and PartialProductKey is not null%_zz6% %_zz3% ID %_zz8%"
for /f "tokens=2 delims==" %%# in ('%_qr%') do (
  set "chkID=%%#"
  call :casWdet "%wspp%" "%wsps%" "%spp_get%"
  call :casWout
  echo %line3%
  echo.
)

if defined ohook (
echo.
echo.
echo %line2%
echo ***                Office Ohook 激活状态                 ***
echo %line2%
echo.
powershell "write-host -back 'Black' -fore 'Yellow' '已使用 Ohook 永久激活 Office。'; write-host -back 'Black' -fore 'Yellow' '你可以忽略以下的 Office 激活状态。'"
echo.
)

:casWcon
set winID=0
set verbose=1
if not defined c0ff1ce15 (
if defined osppsvc goto :casWospp
goto :casWend
)
echo %line2%
echo ***                  Office 激活状态                     ***
echo %line2%
set "_qr=%_zz7% %wspp% %_zz2% %_zz5%ApplicationID='%o15App%' and PartialProductKey is not null%_zz6% %_zz3% ID %_zz8%"
for /f "tokens=2 delims==" %%# in ('%_qr%') do (
  set "chkID=%%#"
  call :casWdet "%wspp%" "%wsps%" "%spp_get%"
  call :casWout
  echo %line3%
  echo.
)
set verbose=0
if defined osppsvc goto :casWospp
goto :casWend

:casWospp
if %verbose% EQU 1 (
echo %line2%
echo ***                  Office 激活状态                     ***
echo %line2%
)
set "_qr=%_zz7% %ospp% %_zz2% %_zz5%ApplicationID='%o15App%' and PartialProductKey is not null%_zz6% %_zz3% ID %_zz8%"
if defined ospp15 for /f "tokens=2 delims==" %%# in ('%_qr%') do (
  set "chkID=%%#"
  call :casWdet "%ospp%" "%osps%" "%ospp_get%"
  call :casWout
  echo %line3%
  echo.
)
set "_qr=%_zz7% %ospp% %_zz2% %_zz5%ApplicationID='%o14App%' and PartialProductKey is not null%_zz6% %_zz3% ID %_zz8%"
if defined ospp14 for /f "tokens=2 delims==" %%# in ('%_qr%') do (
  set "chkID=%%#"
  call :casWdet "%ospp%" "%osps%" "%ospp_get%"
  call :casWout
  echo %line3%
  echo.
)
goto :casWend

:casWpkey
set "_qr=%_zz1% %1 %_zz2% %_zz5%ApplicationID='%2' and PartialProductKey is not null%_zz6% %_zz3% ID %_zz4%"
%_qr% 2>nul | findstr /i ID 1>nul && (set %3=1&set %4=1)
exit /b

:casWdet
for %%# in (%~3) do set "%%#="
if /i %~1==%ospp% for %%# in (DiscoveredKeyManagementServiceMachineIpAddress, KeyManagementServiceLookupDomain, ProductKeyChannel, VLActivationTypeEnabled) do set "%%#="
set "cKmsClient="
set "cTblClient="
set "cAvmClient="
set "ExpireMsg="
set "_xpr="
set "_qr="wmic path %~1 where ID='%chkID%' get %~3 /value" ^| findstr ^="
if %WMI_VBS% NEQ 0 set "_qr=%_csg% %~1 "ID='%chkID%'" "%~3""
for /f "tokens=* delims=" %%# in ('%_qr%') do set "%%#"

set /a _gpr=(GracePeriodRemaining+1440-1)/1440
echo %Description%| findstr /i VOLUME_KMSCLIENT 1>nul && (set cKmsClient=1&set _mTag=批量)
echo %Description%| findstr /i TIMEBASED_ 1>nul && (set cTblClient=1&set _mTag=基于时间的)
echo %Description%| findstr /i VIRTUAL_MACHINE_ACTIVATION 1>nul && (set cAvmClient=1&set _mTag=自动虚拟机)
cmd /c exit /b %LicenseStatusReason%
set "LicenseReason=%=ExitCode%"
set "LicenseMsg=剩余时间：%GracePeriodRemaining% 分钟（%_gpr% 天）"
if %_gpr% GEQ 1 if %_WSH% EQU 1 (
for /f "tokens=* delims=" %%# in ('%_csx% %GracePeriodRemaining%') do set "_xpr=%%#"
)
if %_gpr% GEQ 1 if %_prsh% EQU 1 if not defined _xpr (
for /f "tokens=* delims=" %%# in ('%_psc% "$([DateTime]::Now.addMinutes(%GracePeriodRemaining%)).ToString('yyyy-MM-dd HH:mm:ss')" 2^>nul') do set "_xpr=%%#"
title 检查激活状态 [wmi]
)

if %LicenseStatus% EQU 0 (
set "License=未许可"
set "LicenseMsg="
)
if %LicenseStatus% EQU 1 (
set "License=已许可"
set "LicenseMsg="
if %GracePeriodRemaining% EQU 0 (
  if %winID% EQU 1 (set "ExpireMsg=计算机已永久激活。") else (set "ExpireMsg=产品已永久激活。")
  ) else (
  set "LicenseMsg=%_mTag%激活过期：%GracePeriodRemaining% 分钟（%_gpr% 天）"
  if defined _xpr set "ExpireMsg=%_mTag%激活将于 %_xpr% 过期"
  )
)
if %LicenseStatus% EQU 2 (
set "License=初始宽限期"
if defined _xpr set "ExpireMsg=初始宽限期将于 %_xpr% 结束"
)
if %LicenseStatus% EQU 3 (
set "License=附加宽限期（KMS 许可证已过期或硬件超出容差范围）"
if defined _xpr set "ExpireMsg=附加宽限期将于 %_xpr% 结束"
)
if %LicenseStatus% EQU 4 (
set "License=非正版宽限期。"
if defined _xpr set "ExpireMsg=非正版宽限期将于 %_xpr% 结束"
)
if %LicenseStatus% EQU 6 (
set "License=延长宽限期"
if defined _xpr set "ExpireMsg=延长宽限期将于 %_xpr% 结束"
)
if %LicenseStatus% EQU 5 (
set "License=通知"
  if "%LicenseReason%"=="C004F200" (set "LicenseMsg=通知原因：0xC004F200（非正版）。"
  ) else if "%LicenseReason%"=="C004F009" (set "LicenseMsg=通知原因：0xC004F009（宽限期到期）。"
  ) else (set "LicenseMsg=通知原因：0x%LicenseReason%"
  )
)
if %LicenseStatus% GTR 6 (
set "License=未知"
set "LicenseMsg="
)
if not defined cKmsClient exit /b

if %KeyManagementServicePort%==0 set KeyManagementServicePort=1688
set "KmsReg=已注册的 KMS 计算机名称：%KeyManagementServiceMachine%:%KeyManagementServicePort%"
if "%KeyManagementServiceMachine%"=="" set "KmsReg=已注册的 KMS 计算机名称：KMS 名称不可用"

if %DiscoveredKeyManagementServiceMachinePort%==0 set DiscoveredKeyManagementServiceMachinePort=1688
set "KmsDns=来自 DNS 的 KMS 计算机名称：%DiscoveredKeyManagementServiceMachineName%:%DiscoveredKeyManagementServiceMachinePort%"
if "%DiscoveredKeyManagementServiceMachineName%"=="" set "KmsDns=DNS 自动发现：KMS 名称不可用"

set "_qr="wmic path %~2 get ClientMachineID, KeyManagementServiceHostCaching /value" ^| findstr ^="
if %WMI_VBS% NEQ 0 set "_qr=%_csg% %~2 "ClientMachineID, KeyManagementServiceHostCaching""
for /f "tokens=* delims=" %%# in ('%_qr%') do set "%%#"
if /i %KeyManagementServiceHostCaching%==True (set KeyManagementServiceHostCaching=启用) else (set KeyManagementServiceHostCaching=禁用)

if %winbuild% LSS 9200 exit /b
if /i %~1==%ospp% exit /b

if "%KeyManagementServiceLookupDomain%"=="" set "KeyManagementServiceLookupDomain="

if %VLActivationTypeEnabled% EQU 3 (
set VLActivationType=令牌
) else if %VLActivationTypeEnabled% EQU 2 (
set VLActivationType=KMS
) else if %VLActivationTypeEnabled% EQU 1 (
set VLActivationType=AD
) else (
set VLActivationType=所有
)

if %winbuild% LSS 9600 exit /b
if "%DiscoveredKeyManagementServiceMachineIpAddress%"=="" set "DiscoveredKeyManagementServiceMachineIpAddress=不可用"
exit /b

:casWout
echo.
echo 名称：%Name%
echo 描述：%Description%
echo 激活 ID：%ID%
echo 扩展 PID：%ProductKeyID%
if defined ProductKeyChannel echo 产品密钥通道：%ProductKeyChannel%
echo 部分产品密钥：%PartialProductKey%
echo 许可状态：%License%
if defined LicenseMsg echo %LicenseMsg%
if not %LicenseStatus%==0 if not %EvaluationEndDate:~0,8%==16010101 echo 评估结束日期：%EvaluationEndDate:~0,4%-%EvaluationEndDate:~4,2%-%EvaluationEndDate:~6,2% %EvaluationEndDate:~8,2%:%EvaluationEndDate:~10,2% UTC
if not defined cKmsClient (
if defined ExpireMsg echo.&echo.    %ExpireMsg%
exit /b
)
if defined VLActivationTypeEnabled echo 已配置的激活类型：%VLActivationType%
echo.
if not %LicenseStatus%==1 (
echo 请激活产品以更新 KMS 客户端信息值。
exit /b
)
echo 最新的激活信息：
echo 密钥管理服务客户端信息
echo.    客户端计算机 ID（CMID）：%ClientMachineID%
echo.    %KmsDns%
echo.    %KmsReg%
if defined DiscoveredKeyManagementServiceMachineIpAddress echo.    KMS 计算机 IP 地址：%DiscoveredKeyManagementServiceMachineIpAddress%
echo.    KMS 计算机扩展 PID：%KeyManagementServiceProductKeyID%
echo.    激活间隔：%VLActivationInterval% 分钟
echo.    续订间隔：%VLRenewalInterval% 分钟
echo.    KMS 主机缓存：%KeyManagementServiceHostCaching%
if defined KeyManagementServiceLookupDomain echo.    KMS 服务器记录查找域：%KeyManagementServiceLookupDomain%
if defined ExpireMsg echo.&echo.    %ExpireMsg%
exit /b

:casWend
if %_Identity% EQU 1 if %_prsh% EQU 1 (
echo %line2%
echo ***                Office vNext 激活状态                 ***
echo %line2%
setlocal EnableDelayedExpansion
%_psc% "$f=[IO.File]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':vNextDiag\:.*';iex ($f[1])"
title 检查激活状态 [wmi]
echo %line3%
echo.
)
echo.
call :_color %_Yellow% "请按任意键返回……"
pause >nul
exit /b

:vNextDiag:
function PrintModePerPridFromRegistry
{
	$vNextRegkey = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\LicensingNext"
	$vNextPrids = Get-Item -Path $vNextRegkey -ErrorAction Ignore | Select-Object -ExpandProperty 'property' | Where-Object -FilterScript {$_.ToLower() -like "*retail" -or $_.ToLower() -like "*volume"}
	If ($vNextPrids -Eq $null)
	{
		Write-Host "没有找到注册表项。"
		Return
	}
	$vNextPrids | ForEach `
	{
		$mode = (Get-ItemProperty -Path $vNextRegkey -Name $_).$_
		Switch ($mode)
		{
			2 { $mode = "vNext"; Break }
			3 { $mode = "Device"; Break }
			Default { $mode = "Legacy"; Break }
		}
		Write-Host $_ = $mode
	}
}
function PrintSharedComputerLicensing
{
	$scaRegKey = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
	$scaValue = Get-ItemProperty -Path $scaRegKey -ErrorAction Ignore | Select-Object -ExpandProperty "SharedComputerLicensing" -ErrorAction Ignore
	$scaRegKey2 = "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing"
	$scaValue2 = Get-ItemProperty -Path $scaRegKey2 -ErrorAction Ignore | Select-Object -ExpandProperty "SharedComputerLicensing" -ErrorAction Ignore
	$scaPolicyKey = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Licensing"
	$scaPolicyValue = Get-ItemProperty -Path $scaPolicyKey -ErrorAction Ignore | Select-Object -ExpandProperty "SharedComputerLicensing" -ErrorAction Ignore
	If ($scaValue -Eq $null -And $scaValue2 -Eq $null -And $scaPolicyValue -Eq $null)
	{
		Write-Host "没有找到注册表项。"
		Return
	}
	$scaModeValue = $scaValue -Or $scaValue2 -Or $scaPolicyValue
	If ($scaModeValue -Eq 0)
	{
		$scaMode = "Disabled"
	}
	If ($scaModeValue -Eq 1)
	{
		$scaMode = "Enabled"
	}
	Write-Host "SharedComputerLicensing" = $scaMode
	Write-Host
	$tokenFiles = $null
	$tokenPath = "${env:LOCALAPPDATA}\Microsoft\Office\16.0\Licensing"
	If (Test-Path $tokenPath)
	{
		$tokenFiles = Get-ChildItem -Path $tokenPath -Recurse -File -Filter "*authString*"
	}
	If ($tokenFiles.length -Eq 0)
	{
		Write-Host "没有找到令牌。"
		Return
	}
	$tokenFiles | ForEach `
	{
		$tokenParts = (Get-Content -Encoding Unicode -Path $_.FullName).Split('_')
		$output = [PSCustomObject] `
			@{
				ACID = $tokenParts[0];
				User = $tokenParts[3]
				NotBefore = $tokenParts[4];
				NotAfter = $tokenParts[5];
			} | ConvertTo-Json
		Write-Host $output
	}
}
function PrintLicensesInformation
{
	Param(
		[ValidateSet("NUL", "Device")]
		[String]$mode
	)
	If ($mode -Eq "NUL")
	{
		$licensePath = "${env:LOCALAPPDATA}\Microsoft\Office\Licenses"
	}
	ElseIf ($mode -Eq "Device")
	{
		$licensePath = "${env:PROGRAMDATA}\Microsoft\Office\Licenses"
	}
	$licenseFiles = $null
	If (Test-Path $licensePath)
	{
		$licenseFiles = Get-ChildItem -Path $licensePath -Recurse -File
	}
	If ($licenseFiles.length -Eq 0)
	{
		Write-Host "没有找到许可证。"
		Return
	}
	$licenseFiles | ForEach `
	{
		$license = (Get-Content -Encoding Unicode $_.FullName | ConvertFrom-Json).License
		$decodedLicense = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($license)) | ConvertFrom-Json
		$licenseType = $decodedLicense.LicenseType
		If ($null -Ne $decodedLicense.ExpiresOn)
		{
			$expiry = [DateTime]::Parse($decodedLicense.ExpiresOn, $null, 48)
		}
		Else
		{
			$expiry = New-Object DateTime
		}
		$licenseState = $null
		If ((Get-Date) -Gt (Get-Date $decodedLicense.MetaData.NotAfter))
		{
			$licenseState = "RFM"
		}
		ElseIf ((Get-Date) -Lt (Get-Date $expiry))
		{
			$licenseState = "Licensed"
		}
		Else
		{
			$licenseState = "Grace"
		}
		if ($mode -Eq "NUL")
		{
			$output = [PSCustomObject] `
			@{
				Version = $_.Directory.Name
				Type = "User|${licenseType}";
				Product = $decodedLicense.ProductReleaseId;
				Acid = $decodedLicense.Acid;
				LicenseState = $licenseState;
				EntitlementStatus = $decodedLicense.Status;
				EntitlementExpiration = $decodedLicense.ExpiresOn;
				ReasonCode = $decodedLicense.ReasonCode;
				NotBefore = $decodedLicense.Metadata.NotBefore;
				NotAfter = $decodedLicense.Metadata.NotAfter;
				NextRenewal = $decodedLicense.Metadata.RenewAfter;
				TenantId = $decodedLicense.Metadata.TenantId;
			} | ConvertTo-Json
		}
		ElseIf ($mode -Eq "Device")
		{
			$output = [PSCustomObject] `
			@{
				Version = $_.Directory.Name
				Type = "Device|${licenseType}";
				Product = $decodedLicense.ProductReleaseId;
				Acid = $decodedLicense.Acid;
				DeviceId = $decodedLicense.Metadata.DeviceId;
				LicenseState = $licenseState;
				EntitlementStatus = $decodedLicense.Status;
				EntitlementExpiration = $decodedLicense.ExpiresOn;
				ReasonCode = $decodedLicense.ReasonCode;
				NotBefore = $decodedLicense.Metadata.NotBefore;
				NotAfter = $decodedLicense.Metadata.NotAfter;
				NextRenewal = $decodedLicense.Metadata.RenewAfter;
				TenantId = $decodedLicense.Metadata.TenantId;
			} | ConvertTo-Json
		}
		Write-Output $output
	}
}
	Write-Host
	Write-Host "============== 产品 ID 及许可方式 =============="
	Write-Host
PrintModePerPridFromRegistry
	Write-Host
	Write-Host "================ 共享计算机许可 ================"
	Write-Host
PrintSharedComputerLicensing
	Write-Host
	Write-Host "================= vNext 许可证 ================="
	Write-Host
PrintLicensesInformation -Mode "NUL"
	Write-Host
	Write-Host "================== 设备许可证 =================="
	Write-Host
PrintLicensesInformation -Mode "Device"
:vNextDiag:

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:troubleshoot
@setlocal DisableDelayedExpansion
@echo off

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

::  修复路径名称中的特殊字符限制

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

set _PSarg="""%~f0""" -el %_args%

set "_ttemp=%userprofile%\AppData\Local\Temp"

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
echo Stopping Winmgmt service
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

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:change_edition
@setlocal DisableDelayedExpansion
@echo off

::  要在使用 CBS 升级方法更改版本时标记当前版本，请在以下行中将 0 更改为 1
set _stg=0

::========================================================================================================================================

cls
color 07
title 更改 Windows 版本 %masver%

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

if %_NCS% EQU 1 (
for /F %%a in ('echo prompt $E ^| cmd') do set "esc=%%a"
set     "Red="41;97m""
set    "Gray="100;97m""
set   "Green="42;97m""
set    "Blue="44;97m""
set  "_White="40;37m""
set  "_Green="40;92m""
set "_Yellow="40;93m""
) else (
set     "Red="Red" "white""
set    "Gray="Darkgray" "white""
set   "Green="DarkGreen" "white""
set    "Blue="Blue" "white""
set  "_White="Black" "Gray""
set  "_Green="Black" "Green""
set "_Yellow="Black" "Yellow""
)

set "nceline=echo: &echo ==== 错误 ==== &echo:"
set "eline=echo: &call :dk_color %Red% "==== 错误 ====" &echo:"
set "line=echo ___________________________________________________________________________________________"
if %~z0 GEQ 200000 (set "_exitmsg=返回") else (set "_exitmsg=退出")

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

cls
mode 98, 30

echo:
echo 正在初始化……
echo:
call :dk_product
call :dk_ckeckwmic

::  显示潜在的脚本卡住情况的信息

sc start sppsvc %nul%
if %errorlevel% NEQ 1056 if %errorlevel% NEQ 0 (
echo:
echo 错误代码：%errorlevel%
call :dk_color %Red% "启动 [sppsvc] 服务失败，其余的进程可能需要很长时间……"
echo:
)

::========================================================================================================================================

::  检查激活 ID

call :dk_actids
if not defined applist (
net stop sppsvc /y %nul%
cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
if !errorlevel! NEQ 0 cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
call :dk_refresh
call :dk_actids
if not defined applist (
%eline%
echo 未找到激活 ID。正在中止……
echo:
echo 请查看此页面以获得帮助。 %mas%troubleshoot
goto ced_done
)
)

::========================================================================================================================================

call :dk_checksku

if not defined osSKU (
%eline%
echo 未正确检测到 SKU 值。正在中止……
goto ced_done
)

::========================================================================================================================================

::  检查 Windows 版本

set osedition=
set dismedition=
set dismnotworking=

set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('SLGetWindowsInformation', 'slc.dll', 22, 1, [int], @([String], [int], [int].MakeByRefType(), [String].MakeByRefType()), 1, 3);
set d1=%d1% $editionName = 0; [void]$TypeBuilder.CreateType()::SLGetWindowsInformation('Kernel-EditionName', 0, [ref]0, [ref]$editionName); $editionName
if %winbuild% GEQ 14393 for /f "delims=" %%s in ('"%psc% %d1%"') do if not errorlevel 1 (set osedition=%%s)
if "%osedition%"=="0" set osedition=

if not defined osedition (
for /f "skip=2 tokens=3" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul6%') do set "osedition=%%a"
)

::  解决了在1607至1709版本中将专业教育版显示为专业版的问题。

if %osSKU%==164 set osedition=ProfessionalEducation
if %osSKU%==165 set osedition=ProfessionalEducationN

for /f "tokens=3 delims=: " %%a in ('DISM /English /Online /Get-CurrentEdition %nul6% ^| find /i "Current Edition :"') do set "dismedition=%%a"
if not defined dismedition set dismnotworking=1

if defined dismedition if not defined osedition set osedition=%dismedition%

if not defined osedition (
%eline%
DISM /English /Online /Get-CurrentEdition %nul%
cmd /c exit /b !errorlevel!
echo DISM 命令失败 [错误代码 - 0x!=ExitCode!]
echo 操作系统版本未被正确检测到。正在中止……
echo:
echo 请查看此页面以获取帮助。%mas%troubleshoot
goto ced_done
)

::========================================================================================================================================

set branch=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v BuildBranch %nul6%') do set "branch=%%b"

::  检查 PowerShell

%psc% $ExecutionContext.SessionState.LanguageMode %nul2% | find /i "Full" %nul1% || (
%eline%
%psc% $ExecutionContext.SessionState.LanguageMode
echo:
echo PowerShell 不可用，正在中止……
echo 如果你对Powershell施加了限制，请撤销这些更改。
echo:
echo 请查看此页面以获得帮助。 %mas%troubleshoot
goto ced_done
)

::========================================================================================================================================

::  获取目标版本列表

set _target=
set _dtarget=
set _ptarget=
set _ntarget=
set _wtarget=

if %winbuild% GEQ 10240 for /f "tokens=4" %%a in ('dism /online /english /Get-TargetEditions ^| findstr /i /c:"Target Edition : "') do (if defined _dtarget (set "_dtarget= !_dtarget! %%a ") else (set "_dtarget= %%a "))
if %winbuild% LSS 10240 for /f "tokens=4" %%a in ('%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':cbsxml\:.*';& ([ScriptBlock]::Create($f[1])) -GetTargetEditions;" ^| findstr /i /c:"Target Edition : "') do (if defined _ptarget (set "_ptarget= !_ptarget! %%a ") else (set "_ptarget= %%a "))

if %winbuild% GEQ 10240 if not exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
call :ced_edilist
if /i "%osedition:~0,4%"=="Core" (set "_wtarget= Professional !_wtarget! ")
set "_dtarget= %_dtarget% !_wtarget! "
)

::========================================================================================================================================

::  阻止更改到 CloudEdition 版本或从 CloudEdition 升级

for %%# in (202 203) do if %osSKU%==%%# (
%eline%
echo [%winos% ^| SKU：%osSKU% ^| %winbuild%]
echo 不建议将此已安装版本更改到任何其他版本。
echo 正在中止……
goto ced_done
)

for %%# in ( %_dtarget% %_ptarget% ) do if /i not "%%#"=="%osedition%" (
echo "!_target!" | find /i " %%# " %nul1% || set "_target= !_target! %%# "
)

if defined _target (
for %%# in (%_target%) do (
echo %%# | findstr /i "CountrySpecific CloudEdition ServerRdsh" %nul% || (set "_ntarget=!_ntarget! %%#")
)
)

if not defined _ntarget (
%line%
echo:
if defined dismnotworking call :dk_color %Red% "DISM.exe 没有响应。"
call :dk_color %Gray% "目标版本未找到。"
echo 当前版本 [%osedition% ^| %winbuild%] 无法更改为任何其他版本。
%line%
goto ced_done
)

::========================================================================================================================================

:cedmenu2

cls
mode 98, 30
set inpt=
set counter=0
set verified=0
set targetedition=

%line%
echo:
call :dk_color %Gray% "你可以将当前版本 [%osedition%] [%winbuild%] 更改为下列之一。"
if defined dismnotworking (
call :dk_color %_Yellow% "注 - DISM.exe 没有响应。"
if /i "%osedition:~0,4%"=="Core" call :dk_color %_Yellow% "     - 在更改为 Pro 后，你将会看到更多版本选项可供选择。"
)
%line%
echo:

for %%A in (%_ntarget%) do (
set /a counter+=1
echo [!counter!]  %%A
set targetedition!counter!=%%A
)

%line%
echo:
echo [0]  %_exitmsg%
echo:
call :dk_color %_Green% "请输入选项编号，并按“Enter”键："
set /p inpt=
if "%inpt%"=="" goto cedmenu2
if "%inpt%"=="0" exit /b
for /l %%i in (1,1,%counter%) do (if "%inpt%"=="%%i" set verified=1)
set targetedition=!targetedition%inpt%!
if %verified%==0 goto cedmenu2

::========================================================================================================================================

if %winbuild% LSS 10240 goto :cbsmethod
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" goto :ced_change_server

cls
mode con cols=105 lines=32

set key=
set _chan=
set _dismapi=0

::  检查版本升级是否需要 DISM API 或 slmgr.vbs

if not exist "%SystemRoot%\System32\spp\tokens\skus\%targetedition%\" (
set _dismapi=1
)

set "keyflow=Retail OEM:NONSLP OEM:DM Volume:MAK Volume:GVLK"

call :ced_targetSKU %targetedition%
if defined targetSKU call :ced_windowskey
if defined key if defined pkeychannel set _chan=%pkeychannel%
if not defined key call :changeeditiondata

if not defined key (
%eline%
echo [%targetedition% ^| %winbuild%]
echo 无法从 pkeyhelper.dll 中获取产品密钥
echo:
echo 请查看此页面以获取帮助。%mas%troubleshoot
goto ced_done
)

::========================================================================================================================================

::  在早于 17134 的 Windows 版本中从 Core 更改为非 Core 和更改版本需要“changepk /productkey”方法并重新启动
::  在其他情况下，可以使用“slmgr /ipk”立即更改版本

if %_dismapi%==1 (
mode con cols=105 lines=40
%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':checkrebootflag\:.*';iex ($f[1]);" | find /i "True" %nul% && (
%eline%
echo 已找到挂起的重启标识。
echo:
echo 请重新启动系统，然后再试一次。
goto ced_done
)
)

cls
%line%
echo:
if defined dismnotworking call :dk_color %_Yellow% "DISM.exe 没有响应。"
echo 正在将当前版本 [%osedition%] %winbuild% 更改为 [%targetedition%]
echo:

if %_dismapi%==1 (
call :dk_color %Green% "注——"
echo:
echo  - 在继续之前保存你的工作，系统将会自动重新启动。
echo:
echo  - 在版本更改后，你将需要使用 HWID 选项激活。
%line%
echo:
choice /C:21 /N /M "[1] 继续 [2] %_exitmsg% ："
if !errorlevel!==1 exit /b
)

::========================================================================================================================================

if %_dismapi%==0 (
echo 正在安装 %_chan% 密钥 [%key%]
echo:
if %_wmic% EQU 1 wmic path SoftwareLicensingService where __CLASS='SoftwareLicensingService' call InstallProductKey ProductKey="%key%" %nul%
if %_wmic% EQU 0 %psc% "(([WMISEARCHER]'SELECT Version FROM SoftwareLicensingService').Get()).InstallProductKey('%key%')" %nul%
if not !errorlevel!==0 cscript //nologo %windir%\system32\slmgr.vbs /ipk %key% %nul%

set error_code=!errorlevel!
cmd /c exit /b !error_code!
if !error_code! NEQ 0 set "error_code=[0x!=ExitCode!]"

if !error_code! EQU 0 (
call :dk_refresh
call :dk_color %Green% "[成功]"
echo:
call :dk_color %Gray% "需要重新启动才能更改为正确的版本。"
) else (
call :dk_color %Red% "[不成功] [错误代码：0x!=ExitCode!]"
echo 请查看此页面以获取帮助。%mas%troubleshoot
)
)

if %_dismapi%==1 (
echo:
echo 正在使用 %_chan% 密钥 %key% 应用 DISM API 方法。请稍候……
echo:
%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':dismapi\:.*';& ([ScriptBlock]::Create($f[1])) %targetedition% %key%;"
timeout /t 3 %nul1%
echo:
call :dk_color %Blue% "如果错误，你必须重新启动系统，然后重试。"
echo 请查看此页面以获取帮助。%mas%troubleshoot
)
%line%

goto ced_done

::========================================================================================================================================

:cbsmethod

cls
mode con cols=105 lines=32
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=200;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}"

%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':checkrebootflag\:.*';iex ($f[1]);" | find /i "True" %nul% && (
%eline%
echo 已找到挂起的重启标识。
echo:
echo 请重新启动系统，然后再试一次。
goto ced_done
)

echo:
if defined dismnotworking call :dk_color %_Yellow% "注 - DISM.exe 没有响应。"
echo 正在将当前版本 [%osedition%] %winbuild% 更改为 [%targetedition%]
echo:
call :dk_color %Blue% "重要信息 —— 请在继续之前保存你的工作，系统将自动重新启动。"
echo:
choice /C:01 /N /M "[1] 继续 [0] %_exitmsg% ："
if %errorlevel%==1 exit /b

echo:
echo 正在初始化……
echo:

if %_stg%==0 (set stage=) else (set stage=-StageCurrent)
%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':cbsxml\:.*';& ([ScriptBlock]::Create($f[1])) -SetEdition %targetedition% %stage%;"
echo:
call :dk_color %Blue% "如果错误，你必须重新启动系统，然后重试。"
echo 请查看此页面以获取帮助。%mas%troubleshoot
%line%

goto ced_done

::========================================================================================================================================

:ced_change_server

cls
mode con cols=105 lines=32
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=200;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}"

set key=
set _chan=
set "keyflow=Volume:GVLK Retail Volume:MAK OEM:NONSLP OEM:DM"

call :ced_targetSKU %targetedition%
if defined targetSKU call :ced_windowskey
if defined key if defined pkeychannel set _chan=%pkeychannel%
if not defined key call :changeeditiondata

if not defined key (
%eline%
echo [%targetedition% ^| %winbuild%]
echo 无法从 pkeyhelper.dll 中获取产品密钥
echo:
echo 请查看此页面以获取帮助。%mas%troubleshoot
goto ced_done
)

::========================================================================================================================================

%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':checkrebootflag\:.*';iex ($f[1]);" | find /i "True" %nul% && (
%eline%
echo 已找到挂起的重启标识。
echo:
echo 请重新启动系统，然后再试一次。
goto ced_done
)

cls
echo:
if defined dismnotworking call :dk_color %_Yellow% "注 - DISM.exe 没有响应。"
echo 正在将当前版本 [%osedition%] %winbuild% 更改为 [%targetedition%]
echo:
echo 正在使用 %_chan% 密钥应用命令
echo DISM /online /Set-Edition:%targetedition% /ProductKey:%key% /AcceptEula
DISM /online /Set-Edition:%targetedition% /ProductKey:%key% /AcceptEula

call :dk_color %Blue% "你必须在此阶段重新启动系统。"
echo 帮助：%mas%troubleshoot

::========================================================================================================================================

:ced_done

echo:
call :dk_color %_Yellow% "请按任意键%_exitmsg%脚本……"
pause %nul1%
exit /b

::========================================================================================================================================

::  获取版本列表

:ced_edilist

if %_wmic% EQU 1 set "chkedi=for /f "tokens=2 delims==" %%a in ('"wmic path SoftwareLicensingProduct where (ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f') get LicenseFamily /VALUE" %nul6%')"
if %_wmic% EQU 0 set "chkedi=for /f "tokens=2 delims==" %%a in ('%psc% "(([WMISEARCHER]'SELECT LicenseFamily FROM SoftwareLicensingProduct WHERE ApplicationID=''55c92734-d682-4d71-983e-d6ec3f16059f''').Get()).LicenseFamily ^| %% {echo ('LicenseFamily='+$_)}" %nul6%')"
%chkedi% do (
call if exist %Systemdrive%\Windows\System32\spp\tokens\skus\%%a (
call set "_wtarget= !_wtarget! %%a "
)
)
exit /b

::========================================================================================================================================

::  检查正在挂起的重启标识

:checkrebootflag:
function Test-PendingReboot
{
 if (Test-Path -Path "$env:windir\WinSxS\pending.xml") { return $true }
 if (Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -EA Ignore) { return $true }
 if (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -EA Ignore) { return $true }
 try { 
   $util = [wmiclass]"\\.\root\ccm\clientsdk:CCM_ClientUtilities"
   $status = $util.DetermineIfRebootPending()
   if(($status -ne $null) -and $status.RebootPending){
     return $true
   }
 }catch{}
 
 return $false
}
Test-PendingReboot
:checkrebootflag:

::========================================================================================================================================

:ced_windowskey

for %%# in (pkeyhelper.dll) do @if "%%~$PATH:#"=="" exit /b
for %%# in (%keyflow%) do (
call :dk_pkey %targetSKU% '%%#'
if defined pkey call :dk_pkeychannel !pkey!
if /i [!pkeychannel!]==[%%#] (
set key=!pkey!
exit /b
)
)
exit /b

::========================================================================================================================================

:ced_targetSKU

set k=%1
set targetSKU=
for %%# in (pkeyhelper.dll) do @if "%%~$PATH:#"=="" exit /b

call :dk_reflection

set d1=%ref% [void]$TypeBuilder.DefinePInvokeMethod('GetEditionIdFromName', 'pkeyhelper.dll', 'Public, Static', 1, [int], @([String], [int].MakeByRefType()), 1, 3);
set d1=%d1% $out = 0; [void]$TypeBuilder.CreateType()::GetEditionIdFromName('%k%', [ref]$out); $out

for /f %%a in ('%psc% "%d1%"') do if not errorlevel 1 (set targetSKU=%%a)
if "%targetSKU%"=="0" set targetSKU=
exit /b

::========================================================================================================================================

::  https://github.com/Gamers-Against-Weed/Set-WindowsCbsEdition

:cbsxml:[
param (
    [Parameter()]
    [String]$SetEdition,

    [Parameter()]
    [Switch]$GetTargetEditions,

    [Parameter()]
    [Switch]$StageCurrent
)

function Get-AssemblyIdentity {
    param (
        [String]$PackageName
    )

    $PackageName = [String]$PackageName
    $packageData = ($PackageName -split '~')

    if($packageData[3] -eq '') {
        $packageData[3] = 'neutral'
    }

    return "<assemblyIdentity name=`"$($packageData[0])`" version=`"$($packageData[4])`" processorArchitecture=`"$($packageData[2])`" publicKeyToken=`"$($packageData[1])`" language=`"$($packageData[3])`" />"
}

function Get-SxsName {
    param (
        [String]$PackageName
    )

    $name = ($PackageName -replace '[^A-z0-9\-\._]', '')

    if($name.Length -gt 40) {
        $name = ($name[0..18] -join '') + '\.\.' + ($name[-19..-1] -join '')
    }

    return $name.ToLower()
}

function Find-EditionXmlInSxs {
    param (
        [String]$Edition
    )

    $candidates = @($Edition, 'Client', 'Server')
    $winSxs = $Env:SystemRoot + '\WinSxS'
    $allInSxs = Get-ChildItem -Path $winSxs | select Name

    foreach($candidate in $candidates) {
        $name = Get-SxsName -PackageName "Microsoft-Windows-Editions-$candidate"
        $packages = $allInSxs | where name -Match ('^.*_'+$name+'_31bf3856ad364e35')

        if($packages.Length -eq 0) {
            continue
        }

        $package = $packages[-1].Name
        $testPath = $winSxs + "\$package\" + $Edition + 'Edition.xml'

        if(Test-Path -Path $testPath -PathType Leaf) {
            return $testPath
        }
    }

    return $null
}

function Find-EditionXml {
    param (
        [String]$Edition
    )

    $servicingEditions = $Env:SystemRoot + '\servicing\Editions'
    $editionXml = $Edition + 'Edition.xml'

    $editionXmlInServicing = $servicingEditions + '\' + $editionXml

    if(Test-Path -Path $editionXmlInServicing -PathType Leaf) {
        return $editionXmlInServicing
    }

    return Find-EditionXmlInSxs -Edition $Edition
}

function Write-UpgradeCandidates {
    param (
        [HashTable]$InstallCandidates
    )

    $editionCount = 0
    Write-Host '可升级到的版本：'
    foreach($candidate in $InstallCandidates.Keys) {
        Write-Host "目标版本 ：$candidate"
        $editionCount++
    }

    if($editionCount -eq 0) {
        Write-Host '（无版本可用）'
    }
}

function Write-UpgradeXml {
    param (
        [Array]$RemovalCandidates,
        [Array]$InstallCandidates,
        [Boolean]$Stage
    )

    $removeAction = 'remove'
    if($Stage) {
        $removeAction = 'stage'
    }

    Write-Output '<?xml version="1.0"?>'
    Write-Output '<unattend xmlns="urn:schemas-microsoft-com:unattend">'
    Write-Output '<servicing>'

    foreach($package in $InstallCandidates) {
        Write-Output '<package action="install">'
        Write-Output (Get-AssemblyIdentity -PackageName $package)
        Write-Output '</package>'
    }

    foreach($package in $RemovalCandidates) {
        Write-Output "<package action=`"$removeAction`">"
        Write-Output (Get-AssemblyIdentity -PackageName $package)
        Write-Output '</package>'
    }

    Write-Output '</servicing>'
    Write-Output '</unattend>'
}

function Write-Usage {
    Get-Help $script:MyInvocation.MyCommand.Path -detailed
}

$version = '1.0'
$getTargetsParam = $GetTargetEditions.IsPresent
$stageCurrentParam = $StageCurrent.IsPresent

if($SetEdition -eq '' -and ($false -eq $getTargetsParam)) {
    Write-Usage
    Exit 1
}

$removalCandidates = @();
$installCandidates = @{};

$packages = Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages' | select Name | where { $_.name -match '^.*\\Microsoft-Windows-.*Edition~' }
foreach($package in $packages) {
    $state = (Get-ItemProperty -Path "Registry::$($package.Name)").CurrentState
    $packageName = ($package.Name -split '\\')[-1]
    $packageEdition = (($packageName -split 'Edition~')[0] -split 'Microsoft-Windows-')[-1]

    if($state -eq 0x40) {
        if($null -eq $installCandidates[$packageEdition]) {
            $installCandidates[$packageEdition] = @()
        }

        if($false -eq ($installCandidates[$packageEdition] -contains $packageName)) {
            $installCandidates[$packageEdition] = $installCandidates[$packageEdition] + @($packageName)
        }
    }

    if((($state -eq 0x50) -or ($state -eq 0x70)) -and ($false -eq ($removalCandidates -contains $packageName))) {
        $removalCandidates = $removalCandidates + @($packageName)
    }
}

if($getTargetsParam) {
    Write-UpgradeCandidates -InstallCandidates $installCandidates
    Exit
}

if($false -eq ($installCandidates.Keys -contains $SetEdition)) {
    Write-Error "系统无法升级到“$SetEdition”"
    Exit 1
}

$xmlPath = $Env:SystemRoot + '\Temp' + '\CbsUpgrade.xml'

Write-UpgradeXml -RemovalCandidates $removalCandidates `
    -InstallCandidates $installCandidates[$SetEdition] `
    -Stage $stageCurrentParam >$xmlPath

$editionXml = Find-EditionXml -Edition $SetEdition
if($null -eq $editionXml) {
    Write-Warning '无法找到特定于版本的设置 XML。不使用它继续……'
}

Write-Host '正在开始升级过程。这可能需要一段时间……'

DISM.EXE /English /NoRestart /Online /Apply-Unattend:$xmlPath
$dismError = $LASTEXITCODE

Remove-Item -Path $xmlPath -Force

if(($dismError -ne 0) -and ($dismError -ne 3010)) {
    Write-Error '升级到目标版本失败'
    Exit $dismError
}

if($null -ne $editionXml) {
    $destination = $Env:SystemRoot + '\' + $SetEdition + '.xml'
    Copy-Item -Path $editionXml -Destination $destination

    DISM.EXE /English /NoRestart /Online /Apply-Unattend:$editionXml
    $dismError = $LASTEXITCODE

    if(($dismError -ne 0) -and ($dismError -ne 3010)) {
        Write-Error '应用版本特定设置失败'
        Exit $dismError
    }
}

Restart-Computer
:cbsxml:]

::========================================================================================================================================

::  使用 DISM API 更改版本
::  感谢 Alex（感谢 may，ave9858）

:dismapi:[
param (
    [Parameter()]
    [String]$TargetEdition,

    [Parameter()]
    [String]$Key
)

$AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1)
$ModuleBuilder = $AssemblyBuilder.DefineDynamicModule(2, $False)
$TB = $ModuleBuilder.DefineType(0)

[void]$TB.DefinePInvokeMethod('DismInitialize', 'DismApi.dll', 22, 1, [int], @([int], [IntPtr], [IntPtr]), 1, 3)
[void]$TB.DefinePInvokeMethod('DismOpenSession', 'DismApi.dll', 22, 1, [int], @([String], [IntPtr], [IntPtr], [UInt32].MakeByRefType()), 1, 3)
[void]$TB.DefinePInvokeMethod('_DismSetEdition', 'DismApi.dll', 22, 1, [int], @([UInt32], [String], [String], [IntPtr], [IntPtr], [IntPtr]), 1, 3)
$Dism = $TB.CreateType()

[void]$Dism::DismInitialize(2, 0, 0)
$Session = 0
[void]$Dism::DismOpenSession('DISM_{53BFAE52-B167-4E2F-A258-0A37B57FF845}', 0, 0, [ref]$Session)
if (!$Dism::_DismSetEdition($Session, "$TargetEdition", "$Key", 0, 0, 0)) {
    Restart-Computer
}
:dismapi:]

::========================================================================================================================================

::  第 1 列 = 通用零售/OEM/MAK/GVLK 密钥
::  第 2 列 = 密钥类型
::  第 3 列 = WMI 版本 ID
::  第 4 列 = 版本名称，以防止相同的版本 ID 用于具有不同密钥的不同操作系统版本
::  分隔符  = _

::  对于 Windows 10/11 版本，尽可能列出 HWID 密钥，在服务器版本中，尽可能列出 KMS 密钥。
::  这里仅存储 RS3 和旧版本的通用密钥，以后的密钥是从 pkeyhelper.dll 中提取的本身

:changeeditiondata

if %winbuild% GTR 17763 exit /b
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*CorEdition~*.mum" (set Cor=Cor) else (set Cor=)

set h=
for %%# in (
XGV%h%PP-NM%h%H47-7TT%h%HJ-W%h%3FW7-8HV%h%2C__OEM:NONSLP_Enterprise
D6R%h%D9-D4%h%N8T-RT9%h%QX-Y%h%W6YT-FCW%h%WJ______Retail_Starter
3V6%h%Q6-NQ%h%XCX-V8Y%h%XR-9%h%QCYV-QPF%h%CT__Volume:MAK_EnterpriseN
3NF%h%XW-2T%h%27M-2BD%h%W6-4%h%GHRV-68X%h%RX______Retail_StarterN
VK7%h%JG-NP%h%HTM-C97%h%JM-9%h%MPGT-3V6%h%6T______Retail_Professional
2B8%h%7N-8K%h%FHP-DKV%h%6R-Y%h%2C8J-PKC%h%KT______Retail_ProfessionalN
4CP%h%RK-NM%h%3K3-X6X%h%XQ-R%h%XX86-WXC%h%HW______Retail_CoreN
N24%h%34-X9%h%D7W-8PF%h%6X-8%h%DV9T-8TY%h%MD______Retail_CoreCountrySpecific
BT7%h%9Q-G7%h%N6G-PGB%h%YW-4%h%YWX6-6F4%h%BT______Retail_CoreSingleLanguage
YTM%h%G3-N6%h%DKC-DKB%h%77-7%h%M9GH-8HV%h%X7______Retail_Core
XKC%h%NC-J2%h%6Q9-KFH%h%D2-F%h%KTHY-KD7%h%2Y__OEM:NONSLP_PPIPro
YNM%h%GQ-8R%h%YV3-4PG%h%Q3-C%h%8XTP-7CF%h%BY______Retail_Education
84N%h%GF-MH%h%BT6-FXB%h%X8-Q%h%WJK7-DRR%h%8H______Retail_EducationN
NK9%h%6Y-D9%h%CD8-W44%h%CQ-R%h%8YTK-DYJ%h%WX__OEM:NONSLP_EnterpriseS_RS1
FWN%h%7H-PF%h%93Q-4GG%h%P8-M%h%8RF3-MDW%h%WW__OEM:NONSLP_EnterpriseS_TH
2DB%h%W3-N2%h%PJG-MVH%h%W3-G%h%7TDK-9HK%h%R4__Volume:MAK_EnterpriseSN_RS1
NTX%h%6B-BR%h%YC2-K67%h%86-F%h%6MVQ-M7V%h%2X__Volume:MAK_EnterpriseSN_TH
G3K%h%NM-CH%h%G6T-R36%h%X3-9%h%QDG6-8M8%h%K9______Retail_ProfessionalSingleLanguage
HNG%h%CC-Y3%h%8KG-QVK%h%8D-W%h%MWRK-X86%h%VK______Retail_ProfessionalCountrySpecific
DXG%h%7C-N3%h%6C4-C4H%h%TG-X%h%4T3X-2YV%h%77______Retail_ProfessionalWorkstation
WYP%h%NQ-8C%h%467-V2W%h%6J-T%h%X4WX-WT2%h%RQ______Retail_ProfessionalWorkstationN
8PT%h%T6-RN%h%W4C-6V7%h%J2-C%h%2D3X-MHB%h%PB______Retail_ProfessionalEducation
GJT%h%YN-HD%h%MQY-FRR%h%76-H%h%VGC7-QPF%h%8P______Retail_ProfessionalEducationN
C4N%h%TJ-CX%h%6Q2-VXD%h%MR-X%h%VKGM-F9D%h%JC__Volume:MAK_EnterpriseG
46P%h%N6-R9%h%BK9-CVH%h%KB-H%h%WQ9V-MBJ%h%Y8__Volume:MAK_EnterpriseGN
NJC%h%F7-PW%h%8QT-332%h%4D-6%h%88JX-2YV%h%66______Retail_ServerRdsh
V3W%h%VW-N2%h%PV2-CGW%h%C3-3%h%4QGF-VMJ%h%2C______Retail_Cloud
NH9%h%J3-68%h%WK7-6FB%h%93-4%h%K3DF-DJ4%h%F6______Retail_CloudN
2HN%h%6V-HG%h%TM8-6C9%h%7C-R%h%K67V-JQP%h%FD______Retail_CloudE
WC2%h%BQ-8N%h%RM3-FDD%h%YY-2%h%BFGV-KHK%h%QY_Volume:GVLK_ServerStandard%Cor%_RS1
CB7%h%KF-BW%h%N84-R7R%h%2Y-7%h%93K2-8XD%h%DG_Volume:GVLK_ServerDatacenter%Cor%_RS1
JCK%h%RF-N3%h%7P4-C2D%h%82-9%h%YXRT-4M6%h%3B_Volume:GVLK_ServerSolution_RS1
QN4%h%C6-GB%h%JD2-FB4%h%22-G%h%HWJK-GJG%h%2R_Volume:GVLK_ServerCloudStorage_RS1
VP3%h%4G-4N%h%PPG-79J%h%TQ-8%h%64T4-R3M%h%QX_Volume:GVLK_ServerAzureCor_RS1
9JQ%h%NQ-V8%h%HQ6-PKB%h%8H-G%h%GHRY-R62%h%H6______Retail_ServerAzureNano_RS1
VN8%h%D3-PR%h%82H-DB6%h%BJ-J%h%9P4M-92F%h%6J______Retail_ServerStorageStandard_RS1
48T%h%QX-NV%h%K3R-D8Q%h%R3-G%h%THHM-8FH%h%XC______Retail_ServerStorageWorkgroup_RS1
2HX%h%DN-KR%h%XHB-GPY%h%C7-Y%h%CKFJ-7FV%h%DG_Volume:GVLK_ServerDatacenterACor_RS3
PTX%h%N8-JF%h%HJM-4WC%h%78-M%h%PCBR-9W4%h%KR_Volume:GVLK_ServerStandardACor_RS3
) do (
for /f "tokens=1-4 delims=_" %%A in ("%%#") do if /i %targetedition%==%%C (

if not defined key (
set 4th=%%D
if not defined 4th (
set "key=%%A" & set "_chan=%%B"
) else (
echo "%branch%" | find /i "%%D" %nul1% && (set "key=%%A" & set "_chan=%%B")
)
)
)
)
exit /b

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

:MASend
echo:
if defined _MASunattended timeout /t 2 & exit /b
echo 请按任意键退出脚本……
pause >nul
exit /b

::========================================================================================================================================
::  下方保留空行
