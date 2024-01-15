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
title 解压 $OEM$ 文件夹 %masver%

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
set "eline=echo: &call :ex_color %Red% "==== 错误 ====" &echo:"

::========================================================================================================================================

if %winbuild% LSS 7600 (
%nceline%
echo 检测到不受支持的操作系统版本 [%winbuild%]。
echo 此项目仅支持 Windows 7/8/8.1/10/11 和它们对应的服务器版本。
goto done2
)

for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" (
%nceline%
echo 在系统中未找到 powershell.exe。
goto done2
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
%eline%
echo 脚本是从 temp 文件夹启动的，
echo 最有可能的原因是，你是直接从压缩文件运行脚本。
echo:
echo 请解压压缩文件并从解压文件夹中启动脚本。
goto done2
)
)

::========================================================================================================================================

::  将脚本提升为管理员权限及传递参数并防止循环

%nul1% fltmc || (
if not defined _elev for %%# in (wt.exe) do @if "%%~$PATH:#"=="" %psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && exit /b
if not defined _elev %psc% "start wt.exe -arg 'new-tab cmd.exe /c \"!_PSarg:'=''!\"' -verb runas" && exit /b
%eline%
echo 此脚本需要管理员权限。
echo 为此，右键单击此脚本并选择“以管理员身份运行”。
goto done2
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
call :ex_color %_Green% "请输入一个菜单选项 [1,0] ："
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

set "_pdesk=%desktop:'=''%"
set "_dir=%desktop%\$OEM$\$$\Setup\Scripts"

if exist "!desktop!\" (
%eline%
echo 桌面位置未被检测到，正在中止……
goto done2
)

setlocal EnableDelayedExpansion

::========================================================================================================================================

mode con cols=78 lines=30

if exist "!desktop!\$OEM$\" (
echo _____________________________________________________
%eline%
echo $OEM$ 文件夹已经存在于桌面上。
echo _____________________________________________________
goto done2
)

set HWID_Activation.cmd=Activators\HWID_Activation.cmd
set KMS38_Activation.cmd=Activators\KMS38_Activation.cmd
set Online_KMS_Activation.cmd=Activators\Online_KMS_Activation.cmd
set Ohook_Activation_AIO.cmd=Activators\Ohook_Activation_AIO.cmd
pushd "!_work!"

set _nofile=
for %%# in (
%HWID_Activation.cmd%
%KMS38_Activation.cmd%
%Online_KMS_Activation.cmd%
%Ohook_Activation_AIO.cmd%
) do (
if not exist "%%#" set _nofile=1
)

popd

if defined _nofile (
echo _____________________________________________________
%eline%
echo Some files are missing in the 'Activators' folder.
echo _____________________________________________________
goto done2
)

::========================================================================================================================================

:Menu

cls
mode con cols=78 lines=30
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
call :ex_color2 %_White% "              [R] " %_Green% "自述文件"
echo:             [0] 返回
echo:           ________________________________________________________
echo:  
call :ex_color2 %_White% "             " %_Green% "请输入一个菜单选项 ："
choice /C:123456789R0 /N
set _erl=%errorlevel%

if %_erl%==11 exit /b
if %_erl%==10 start %mas%oem-folder.html &goto :Menu
if %_erl%==9 goto:kms_ohook
if %_erl%==8 goto:kms38_kms
if %_erl%==7 goto:kms38_ohook
if %_erl%==6 goto:hwid_kms
if %_erl%==5 goto:hwid_ohook
if %_erl%==4 goto:kms
if %_erl%==3 goto:kms38
if %_erl%==2 goto:ohook
if %_erl%==1 goto:hwid
goto :Menu

::========================================================================================================================================

:hwid

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%HWID_Activation.cmd%" "!_dir!\HWID_Activation.cmd" %nul%
popd
call :export hwid_setup

set _error=
if not exist "!_dir!\HWID_Activation.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=HWID
goto done

:hwid_setup:
@echo off

fltmc >nul || exit /b

call "%~dp0HWID_Activation.cmd" /HWID

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:hwid_setup:

::========================================================================================================================================

:ohook

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b %Ohook_Activation_AIO.cmd% "!_dir!\Ohook_Activation_AIO.cmd" %nul%
popd
call :export ohook_setup

set _error=
if not exist "!_dir!\Ohook_Activation_AIO.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=Ohook
goto done

:ohook_setup:
@echo off

fltmc >nul || exit /b

call "%~dp0Ohook_Activation_AIO.cmd" /Ohook

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:ohook_setup:

::========================================================================================================================================

:kms38

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%KMS38_Activation.cmd%" "!_dir!\KMS38_Activation.cmd" %nul%
popd
call :export kms38_setup

set _error=
if not exist "!_dir!\KMS38_Activation.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=KMS38
goto done

:kms38_setup:
@echo off

fltmc >nul || exit /b

call "%~dp0KMS38_Activation.cmd" /KMS38

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:kms38_setup:

::========================================================================================================================================

:kms

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%Online_KMS_Activation.cmd%" "!_dir!\Online_KMS_Activation.cmd" %nul%
popd
call :export kms_setup

set _error=
if not exist "!_dir!\Online_KMS_Activation.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=在线 KMS
goto done

:kms_setup:
@echo off

fltmc >nul || exit /b

call "%~dp0Online_KMS_Activation.cmd" /KMS-ActAndRenewalTask /KMS-WindowsOffice

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:kms_setup:

::========================================================================================================================================

:hwid_ohook

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%HWID_Activation.cmd%" "!_dir!\HWID_Activation.cmd" %nul%
copy /y /b "%Ohook_Activation_AIO.cmd%" "!_dir!\Ohook_Activation_AIO.cmd" %nul%
popd
call :export hwid_ohook_setup

set _error=
if not exist "!_dir!\HWID_Activation.cmd" set _error=1
if not exist "!_dir!\Ohook_Activation_AIO.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=HWID [Windows] + Ohook [Office]
goto done

:hwid_ohook_setup:
@echo off

fltmc >nul || exit /b

setlocal
call "%~dp0HWID_Activation.cmd" /HWID
endlocal

setlocal
call "%~dp0Ohook_Activation_AIO.cmd" /Ohook
endlocal

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:hwid_ohook_setup:

::========================================================================================================================================

:hwid_kms

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%HWID_Activation.cmd%" "!_dir!\HWID_Activation.cmd" %nul%
copy /y /b "%Online_KMS_Activation.cmd%" "!_dir!\Online_KMS_Activation.cmd" %nul%
popd
call :export hwid_kms_setup

set _error=
if not exist "!_dir!\HWID_Activation.cmd" set _error=1
if not exist "!_dir!\Online_KMS_Activation.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=HWID [Windows] + 在线 KMS [Office]
goto done

:hwid_kms_setup:
@echo off

fltmc >nul || exit /b

setlocal
call "%~dp0HWID_Activation.cmd" /HWID
endlocal

setlocal
call "%~dp0Online_KMS_Activation.cmd" /KMS-ActAndRenewalTask /KMS-Office
endlocal

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:hwid_kms_setup:

::========================================================================================================================================

:kms38_ohook

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%KMS38_Activation.cmd%" "!_dir!\KMS38_Activation.cmd" %nul%
copy /y /b "%Ohook_Activation_AIO.cmd%" "!_dir!\Ohook_Activation_AIO.cmd" %nul%
popd
call :export kms38_ohook_setup

set _error=
if not exist "!_dir!\KMS38_Activation.cmd" set _error=1
if not exist "!_dir!\Ohook_Activation_AIO.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=KMS38 [Windows] + Ohook [Office]
goto done

:kms38_ohook_setup:
@echo off

fltmc >nul || exit /b

setlocal
call "%~dp0KMS38_Activation.cmd" /KMS38
endlocal

setlocal
call "%~dp0Ohook_Activation_AIO.cmd" /Ohook
endlocal

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:kms38_ohook_setup:

::========================================================================================================================================

:kms38_kms

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%KMS38_Activation.cmd%" "!_dir!\KMS38_Activation.cmd" %nul%
copy /y /b "%Online_KMS_Activation.cmd%" "!_dir!\Online_KMS_Activation.cmd" %nul%
popd
call :export kms38_kms_setup

set _error=
if not exist "!_dir!\KMS38_Activation.cmd" set _error=1
if not exist "!_dir!\Online_KMS_Activation.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=KMS38 [Windows] + 在线 KMS [Office]
goto done

:kms38_kms_setup:
@echo off

fltmc >nul || exit /b

setlocal
call "%~dp0KMS38_Activation.cmd" /KMS38
endlocal

setlocal
call "%~dp0Online_KMS_Activation.cmd" /KMS-ActAndRenewalTask /KMS-Office
endlocal

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:kms38_kms_setup:

::========================================================================================================================================

:kms_ohook

cls
md "!desktop!\$OEM$\$$\Setup\Scripts"
pushd "!_work!"
copy /y /b "%Online_KMS_Activation.cmd%" "!_dir!\Online_KMS_Activation.cmd" %nul%
copy /y /b "%Ohook_Activation_AIO.cmd%" "!_dir!\Ohook_Activation_AIO.cmd" %nul%
popd
call :export kms_ohook_setup

set _error=
if not exist "!_dir!\Online_KMS_Activation.cmd" set _error=1
if not exist "!_dir!\Ohook_Activation_AIO.cmd" set _error=1
if not exist "!_dir!\SetupComplete.cmd" set _error=1
if defined _error goto errorfound

set oem=在线 KMS [Windows] + Ohook [Office]
goto done

:kms_ohook_setup:
@echo off

fltmc >nul || exit /b

setlocal
call "%~dp0Online_KMS_Activation.cmd" /KMS-ActAndRenewalTask /KMS-Windows
endlocal

setlocal
call "%~dp0Ohook_Activation_AIO.cmd" /Ohook
endlocal

cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")
:kms_ohook_setup:

::========================================================================================================================================

:errorfound

%eline%
echo $OEM$ 文件夹未能成功创建……
goto :done2

:done

echo ______________________________________________________________
echo:
call :ex_color %Blue% "%oem%"
call :ex_color %Green% "$OEM$ 文件夹已在桌面上成功创建。"
echo "%oem%" | find /i "38" %nul% && (
echo:
echo 对于 KMS38 激活服务器 Cor/Acor 版本（无 GUI 版本），
echo 查看此页面 %mas%oem-folder
)
echo ______________________________________________________________

:done2

echo:
call :ex_color %_Yellow% "请按任意键退出脚本……"
pause %nul1%
exit /b

::========================================================================================================================================

::  从批处理脚本中解压文本，无字符和文件编码问题

:export

%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split \":%~1\:.*`r`n\"; [io.file]::WriteAllText('!_pdesk!\$OEM$\$$\Setup\Scripts\SetupComplete.cmd',$f[1].Trim(),[System.Text.Encoding]::ASCII);"
exit /b

::========================================================================================================================================

:ex_color

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[0m
) else (
if not exist %psc% (echo %~3) else (%psc% write-host -back '%1' -fore '%2' '%3')
)
exit /b

:ex_color2

if %_NCS% EQU 1 (
echo %esc%[%~1%~2%esc%[%~3%~4%esc%[0m
) else (
if not exist %psc% (echo %~3%~6) else (%psc% write-host -back '%1' -fore '%2' '%3' -NoNewline; write-host -back '%4' -fore '%5' '%6')
)
exit /b

::========================================================================================================================================
::  下方保留空行
