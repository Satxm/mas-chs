@set masver=2.5
@setlocal DisableDelayedExpansion
@echo off



::=================================================================================================
::
::   �˽ű��ǡ�Microsoft ����ű�����MAS����Ŀ�е�һ���֡�
::
::   ��    ҳ��mass grave[.]dev
::      Email��windowsaddict@protonmail.com
::
::=================================================================================================



::========================================================================================================================================

::  ����·�������������ϵͳ�����ô���ʱ����������

set "PATH=%SystemRoot%\System32;%SystemRoot%\System32\wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\;%LocalAppData%\Microsoft\WindowsApps\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "PATH=%SystemRoot%\Sysnative;%SystemRoot%\Sysnative\wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%LocalAppData%\Microsoft\WindowsApps\;%PATH%"
)

::  ����ű����� x64 λ Windows �ϵ� x86 ���������ģ���ʹ�� x64 �������������ű�
::  ������������� ARM64 Windows �ϵ� x86/ARM32 ���������ģ���ʹ�� ARM64 ����

set "_cmdf=%~f0"
for %%# in (%*) do (
if /i "%%#"=="r1" set r1=1
if /i "%%#"=="r2" set r2=1
if /i "%%#"=="-qedit" (
reg add HKCU\Console /v QuickEdit /t REG_DWORD /d "1" /f %nul1%
rem �鿴�·��Ĺ���Ա���������˽���Ϊʲô������
)
)

if exist %SystemRoot%\Sysnative\cmd.exe if not defined r1 (
setlocal EnableDelayedExpansion
for %%# in (wt.exe) do @if "%%~$PATH:#"=="" start %SystemRoot%\Sysnative\cmd.exe /c ""!_cmdf!" %* r1" && exit /b
start wt.exe new-tab %SystemRoot%\Sysnative\cmd.exe /c ""!_cmdf!" %* r1" && exit /b
)

::  ʹ�� ARM32 �������������ű�������˽ű����� ARM64 Windows �ϵ� x64 ���������ģ�

if exist %SystemRoot%\SysArm32\cmd.exe if %PROCESSOR_ARCHITECTURE%==AMD64 if not defined r2 (
setlocal EnableDelayedExpansion
for %%# in (wt.exe) do @if "%%~$PATH:#"=="" start %SystemRoot%\SysArm32\cmd.exe /c ""!_cmdf!" %* r2" && exit /b
start wt.exe new-tab %SystemRoot%\SysArm32\cmd.exe /c ""!_cmdf!" %* r2" && exit /b
)

::========================================================================================================================================

set "blank="
set "mas=ht%blank%tps%blank%://mass%blank%grave.dev/"

::  ��� Null �����Ƿ��������������������ű�����Ҫ

sc query Null | find /i "RUNNING"
if %errorlevel% NEQ 0 (
echo:
echo Null ����δ���У��ű����ܻ��������
echo:
echo:
echo ���� - %mas%troubleshoot.html
echo:
echo:
ping 127.0.0.1 -n 10
)
cls

::  ��� LF ��β

pushd "%~dp0"
>nul findstr /v "$" "%~nx0" && (
echo:
echo ���󣺽ű������� LF �н��������⣬�����ڽű�ĩβȱ�ٿ��С�
echo:
ping 127.0.0.1 -n 6 >nul
popd
exit /b
)
popd

::========================================================================================================================================

cls
color 07
title ���ѽ�� %masver%

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

set "nceline=echo: &echo ==== ���� ==== &echo:"
set "eline=echo: &call :_color %Red% "==== ���� ====" &echo:"
set "line=_________________________________________________________________________________________________"
if %~z0 GEQ 200000 (set "_exitmsg=����") else (set "_exitmsg=�˳�")

::========================================================================================================================================

if %winbuild% LSS 7600 (
%nceline%
echo ��⵽����֧�ֵĲ���ϵͳ�汾 [%winbuild%]��
echo ����Ŀ��֧�� Windows 7/8/8.1/10/11 �����Ƕ�Ӧ�ķ������汾��
goto at_done
)

for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" (
%nceline%
echo ��ϵͳ��δ�ҵ� powershell.exe��
goto at_done
)

::========================================================================================================================================

::  �޸�·�������е������ַ�����

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
echo �ű��Ǵ� temp �ļ��������ģ�
echo ���п��ܵ�ԭ���ǣ�����ֱ�Ӵ�ѹ���ļ����нű���
echo:
echo ���ѹѹ���ļ����ӽ�ѹ�ļ����������ű���
goto at_done
)
)

::========================================================================================================================================

::  ���ű�����Ϊ����ԱȨ�޼����ݲ�������ֹѭ��

%nul1% fltmc || (
if not defined _elev for %%# in (wt.exe) do @if "%%~$PATH:#"=="" %psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && exit /b
if not defined _elev %psc% "start wt.exe -arg 'new-tab cmd.exe /c \"!_PSarg:'=''!\"' -verb runas" && exit /b
%nceline%
echo �˽ű���Ҫ����ԱȨ�ޡ�
echo Ϊ�ˣ��Ҽ������˽ű���ѡ���Թ���Ա������С���
goto at_done
)

::========================================================================================================================================

::  �˴�������ô� cmd.exe �Ự�Ŀ��ٱ༭��������ע���������ø���
::  �������ԭ���ǵ����ű����ڻ���ͣ���������½ű�������ֹͣ�Ļ���

for %%# in (%_args%) do (if /i "%%#"=="-qedit" set quedit=1)

reg query HKCU\Console /v QuickEdit %nul2% | find /i "0x0" %nul1% || if not defined quedit (
reg add HKCU\Console /v QuickEdit /t REG_DWORD /d "0" /f %nul1%
for %%# in (wt.exe) do @if "%%~$PATH:#"=="" start cmd.exe /c ""!_batf!" %_args% -qedit" && exit /b
start wt.exe new-tab cmd.exe /c ""!_batf!" %_args% -qedit" && exit /b
rem ���ٱ༭���ô���Ӧ����ڽű��Ŀ�ͷ�����Ǵ˴�����Ϊ��ĳЩ�������Ҫʱ������ӳ
exit /b
)

::========================================================================================================================================

::  ������

set -=
set old=

for /f "delims=[] tokens=2" %%# in ('ping -4 -n 1 updatecheck.mass%-%grave.dev') do (
if not [%%#]==[] (echo "%%#" | find "127.69" %nul1% && (echo "%%#" | find "127.69.%masver%" %nul1% || set old=1))
)

if defined old (
echo ________________________________________________
%eline%
echo echo ���������оɰ汾�� MAS �汾 %masver%
echo ________________________________________________
echo:
echo [1] �������°汾 MAS
echo [0] ��Ȼ����
echo:
call :_color %_Green% "������һ���˵�ѡ�� [1,0] ��"
choice /C:10 /N
if !errorlevel!==2 rem
if !errorlevel!==1 (start ht%-%tps://github.com/mass%-%gravel/Microsoft-Acti%-%vation-Scripts & start %mas% & exit /b)
)
cls

::========================================================================================================================================

setlocal DisableDelayedExpansion

::  �������λ��

set desktop=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "desktop=%%b"
if not defined desktop for /f "delims=" %%a in ('%psc% "& {write-host $([Environment]::GetFolderPath('Desktop'))}"') do call set "desktop=%%a"

if not defined desktop (
%eline%
echo ����λ��δ����⵽��������ֹ����
goto at_done
)

setlocal EnableDelayedExpansion

::========================================================================================================================================

:at_menu

cls
color 07
title ���ѽ�� %masver%
mode con cols=77 lines=30

echo:
echo:
echo:
echo:
echo:       _______________________________________________________________
echo:                                                   
call :_color2 %_White% "             [1] " %_Green% "����"
echo:             ___________________________________________________
echo:                                                                      
echo:             [2] Dism RestoreHealth
echo:             [3] SFC Scannow
echo:                                                                      
echo:             [4] �޸� WMI
echo:             [5] �޸����
echo:             [6] �޸� WPA ע���
echo:             ___________________________________________________
echo:
echo:             [0] %_exitmsg%
echo:       _______________________________________________________________
echo:          
call :_color2 %_White% "            " %_Green% "������һ���˵�ѡ�� ��"
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
echo ��⵽����֧�ֵĲ���ϵͳ�汾��
echo ���� Windows 8/8.1/10/11 �����Ӧ�� Server ֧�ִ����
goto :at_back
)

set _int=
for %%a in (l.root-servers.net resolver1.opendns.com download.windowsupdate.com google.com) do if not defined _int (
for /f "delims=[] tokens=2" %%# in ('ping -n 1 %%a') do (if not [%%#]==[] set _int=1)
)

echo:
if defined _int (
echo      ���ڼ�� Internet ����        [������]
) else (
call :_color2 %_White% "     " %Red% "���ڼ�� Internet ����        [δ����]"
)

echo %line%
echo:
echo      Dism ʹ�� Windows �������ṩ�޸���������ļ���
echo      �⽫��Ҫ 5-15 ���ӻ����ʱ�䡭��
echo %line%
echo:
echo      ע��
echo:
call :_color2 %_White% "     - " %Gray% "��ȷ�� Internet �����ӡ�"
call :_color2 %_White% "     - " %Gray% "��ȷ�� Windows ��������������"
echo:
echo %line%
echo:
choice /C:09 /N /M ">    [9] ���� [0] ���� ��"
if %errorlevel%==1 goto at_menu

cls
mode 110, 30
%psc% Stop-Service TrustedInstaller -force %nul%

set _time=
for /f %%a in ('%psc% "Get-Date -format HH_mm_ss"') do set _time=%%a
echo:
echo ����Ӧ�����
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
call :_color %Gray% "CBS �� DISM ��־�ѱ����Ƶ������ϵ� AT_Logs �ļ����С�"
goto :at_back

::========================================================================================================================================

:sfcscan

cls
mode 98, 30
title sfc /scannow

echo:
echo %line%
echo:    
echo      ϵͳ�ļ��������System File Checker�����޸���ʧ���𻵵�ϵͳ�ļ���
echo      �⽫��Ҫ 10-15 ���ӻ����ʱ�䡭��
echo:
echo      ��� SFC �޷��޸�ĳЩ���⣬���ٴ����и������Բ鿴��һ���Ƿ��ܹ��޸�����ʱ������
echo      Ҫ���� sfc /scannow ���� 3 �Σ�ÿ��֮���������� PC ������ȫ�޸����ܹ��޸�������
echo      ���⡣
echo:   
echo %line%
echo:
choice /C:09 /N /M ">    [9] ���� [0] ���� ��"
if %errorlevel%==1 goto at_menu

cls
%psc% Stop-Service TrustedInstaller -force %nul%

set _time=
for /f %%a in ('%psc% "Get-Date -format HH_mm_ss"') do set _time=%%a
echo:
echo ����Ӧ�����
echo sfc /scannow
sfc /scannow

%psc% Stop-Service TrustedInstaller -force %nul%

if not exist "!desktop!\AT_Logs\" md "!desktop!\AT_Logs\" %nul%

call :compresslog cbs\CBS.log SFC_CBS %nul%

if not exist "!desktop!\AT_Logs\SFC_CBS_%_time%.cab" (
copy /y /b "%SystemRoot%\logs\cbs\cbs.log" "!desktop!\AT_Logs\SFC_CBS_%_time%.log" %nul%
)

echo:
call :_color %Gray% "CBS ����Ҫ��ȡ����־�����Ƶ������ϵ� AT_Logs �ļ����С�"
goto :at_back

::========================================================================================================================================

:retokens

cls
mode con cols=125 lines=32
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=200;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}"
title �޸���ɣ�ClipSVC ^+ Office vNext ^+ SPP ^+ OSPP��

echo:
echo %line%
echo:   
echo      ע��
echo:
echo       - �������ڽ���������⡣
echo:
echo       - ��ѡ��ᣬ
echo            - ������ Windows �� Office���������Ҫ���¼���
echo              ��� Windows ��ʹ������ / OEM / �������֤����ģ��벻Ҫ����
echo:
echo            - ���� ClipSVC��Office vNext��SPP �� OSPP ���
echo            - �޸������ļ��к�ע���� SPP Ȩ��
echo            - ���� Office ���޸�ѡ�
echo:
call :_color2 %_White% "      - " %Red% "���ڱ�ҪʱӦ�á�"
echo:
echo %line%
echo:
choice /C:09 /N /M ">    [9] ���� [0] ���� ��"
if %errorlevel%==1 goto at_menu

::========================================================================================================================================

::  �ؽ� ClipSVC ���֤

cls
:cleanlicensing

echo:
echo %line%
echo:
call :_color %Blue% "�����ؽ� ClipSVC ���֤"
echo:

if %winbuild% LSS 10240 (
echo ClipSVC ���֤�ؽ����� Win 10/11 �ͷ�������Ӧ�汾����֧�֡�
echo ������������
goto :cleanvnext
)

%psc% "(([WMISEARCHER]'SELECT Name FROM SoftwareLicensingProduct WHERE LicenseStatus=1 AND GracePeriodRemaining=0 AND PartialProductKey IS NOT NULL').Get()).Name" %nul2% | findstr /i "Windows" %nul1% && (
echo Windows �����ü��
echo ���������ؽ� ClipSVC ���֤����
goto :cleanvnext
)

echo ����ֹͣ ClipSVC ���񡭡�
%psc% Stop-Service ClipSVC -force %nul%
timeout /t 2 %nul%

echo:
echo ���ڽ�����Ӧ�������� ClipSVC ���֤����
echo rundll32 clipc.dll,ClipCleanUpState

rundll32 clipc.dll,ClipCleanUpState

if %winbuild% LEQ 10240 (
echo [�ɹ�]
) else (
if exist "%ProgramData%\Microsoft\Windows\ClipSVC\tokens.dat" (
call :_color %Red% "[ʧ��]"
) else (
echo [�ɹ�]
)
)

::  �·���ע������ʧ�Ժ��ܱ������� ClipSVC ���֤��������֮�󴴽���������ϵͳ�����������Զ�ɾ����
::  ��Ҫɾ�������ܼ���ϵͳ����������������

set "RegKey=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ClipSVC\Volatile\PersistedSystemState"
set "_ident=HKU\S-1-5-19\SOFTWARE\Microsoft\IdentityCRL"

reg query "%RegKey%" %nul% && %nul% call :regownstart
reg delete "%RegKey%" /f %nul% 

echo:
echo ����ɾ����ʧ�Ժ��ܱ�����ע������
echo [%RegKey%]
reg query "%RegKey%" %nul% && (
call :_color %Red% "[ʧ��]"
echo ��������ϵͳ���⽫���Զ�ɾ����ע����
) || (
echo [�ɹ�]
)

::   ��� HWID �������ע������޸�����Է������κ���

echo:
echo ����ɾ�� IdentityCRL ע������
echo [%_ident%]
reg delete "%_ident%" /f %nul%
reg query "%_ident%" %nul% && (
call :_color %Red% "[ʧ��]"
) || (
echo [�ɹ�]
)

%psc% Stop-Service ClipSVC -force %nul%

::  �ؽ� ClipSVC �ļ������޸�Ȩ������

echo:
if %winbuild% GTR 10240 (
echo ����ɾ���ļ��� %ProgramData%\Microsoft\Windows\ClipSVC\
rmdir /s /q "C:\ProgramData\Microsoft\Windows\ClipSvc" %nul%

if exist "%ProgramData%\Microsoft\Windows\ClipSVC\" (
call :_color %Red% "[ʧ��]"
) else (
echo [�ɹ�]
)

echo:
echo Rebuilding Folder %ProgramData%\Microsoft\Windows\ClipSVC\
%psc% Start-Service ClipSVC %nul%
timeout /t 3 %nul%
if not exist "%ProgramData%\Microsoft\Windows\ClipSVC\" timeout /t 5 %nul%
if not exist "%ProgramData%\Microsoft\Windows\ClipSVC\" (
call :_color %Red% "[ʧ��]"
) else (
echo [�ɹ�]
)
)

echo:
echo �������� [wlidsvc LicenseManager] ���񡭡�
for %%# in (wlidsvc LicenseManager) do (%psc% Restart-Service %%# %nul%)

::========================================================================================================================================

::  ���� Office vNext ���֤��Ĳ��ಢ����ɾ������Ϊ������ֹ�� vNext ���֤����ʾ
::  https://learn.microsoft.com/en-us/office/troubleshoot/activation/reset-office-365-proplus-activation-state

:cleanvnext

echo:
echo %line%
echo:
call :_color %Blue% "�������� Office vNext ���֤"
echo:

setlocal DisableDelayedExpansion
set "_Local=%LocalAppData%"
setlocal EnableDelayedExpansion

attrib -R "!ProgramData!\Microsoft\Office\Licenses" %nul%
attrib -R "!_Local!\Microsoft\Office\Licenses" %nul%

if exist "!ProgramData!\Microsoft\Office\Licenses\" (
rd /s /q "!ProgramData!\Microsoft\Office\Licenses\" %nul%
if exist "!ProgramData!\Microsoft\Office\Licenses\" (
echo ɾ��ʧ�� - !ProgramData!\Microsoft\Office\Licenses\
) else (
echo ��ɾ���ļ��� - !ProgramData!\Microsoft\Office\Licenses\
)
) else (
echo δ�ҵ� - !ProgramData!\Microsoft\Office\Licenses\
)

if exist "!_Local!\Microsoft\Office\Licenses\" (
rd /s /q "!_Local!\Microsoft\Office\Licenses\" %nul%
if exist "!_Local!\Microsoft\Office\Licenses\" (
echo ɾ��ʧ�� - !_Local!\Microsoft\Office\Licenses\
) else (
echo ��ɾ���ļ��� - !_Local!\Microsoft\Office\Licenses\
)
) else (
echo δ�ҵ� - !_Local!\Microsoft\Office\Licenses\
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
echo ��ɾ��ע��� - %%A
) || (
echo ɾ��ʧ�� - %%A
)
)
)
)
if not defined regfound echo δ�ҵ� - Office vNext ע���

::========================================================================================================================================

::  �ؽ� SPP �������

echo:
echo %line%
echo:
call :_color %Blue% "�����ؽ� SPP �������"
echo:

call :scandat check

if not defined token (
call :_color %Red% "δ�ҵ� tokens.dat �ļ���"
) else (
echo tokens.dat �ļ���[%token%]
)

echo:
set wpainfo=
for /f "delims=" %%a in ('%psc% "$f=[io.file]::ReadAllText('!_batp!',[Text.Encoding]::Default) -split ':wpatest\:.*';iex ($f[1]);" %nul6%') do (set wpainfo=%%a)
echo "%wpainfo%" | find /i "Error Found" %nul% && (
call :_color %Red% "WPA ע�����%wpainfo%"
) || (
echo WPA ע�������%wpainfo%
)

set tokenstore=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v TokenStore %nul6%') do call set "tokenstore=%%b"

::  ��� sppsvc Ȩ�޲�Ӧ���޲�����

if %winbuild% GEQ 10240 (

echo:
echo ���ڼ�� SPP Ȩ��������⡭��
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
call :_color %Red% "[�޸�ʧ��]"
) else (
echo [�޸��ɹ�]
)
) else (
echo [δ���ִ���]
)
)

echo:
echo ����ֹͣ sppsvc ����   ...
%psc% Stop-Service sppsvc -force %nul%

echo:
call :scandat delete
call :scandat check

if defined token (
echo:
call :_color %Red% "ɾ�� .dat �ļ�ʧ�ܡ�"
echo:
)

echo:
echo �������°�װϵͳ��� [slmgr /rilc]����
cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
if %errorlevel% NEQ 0 cscript //nologo %windir%\system32\slmgr.vbs /rilc %nul%
if %errorlevel% EQU 0 (
echo [�ɹ�]
) else (
call :_color %Red% "[ʧ��]"
)

call :scandat check

echo:
if not defined token (
call :_color %Red% "�ؽ� tokens.dat �ļ�ʧ�ܡ�"
) else (
echo tokens.dat �ļ��ѳɹ��ؽ���
)

::========================================================================================================================================

::  �ؽ� OSPP ����

echo:
echo %line%
echo:
call :_color %Blue% "�����ؽ� OSPP �������"
echo:

sc qc osppsvc %nul% || (
echo δ��װ���� OSPP �� Office
echo ���������ؽ� OSPP ���ơ���
goto :repairoffice
)

call :scandatospp check

if not defined token (
call :_color %Red% "δ�ҵ� tokens.dat �ļ���"
) else (
echo tokens.dat �ļ���[%token%]
)

echo:
echo ����ֹͣ osppsvc ���񡭡�
%psc% Stop-Service osppsvc -force %nul%

echo:
call :scandatospp delete
call :scandatospp check

if defined token (
echo:
call :_color %Red% "ɾ�� .dat �ļ�ʧ�ܡ�"
echo:
)

echo:
echo �������� osppsvc ���������� tokens.dat
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
call :_color %Red% "�ؽ� tokens.dat �ļ�ʧ�ܡ�"
) else (
echo tokens.dat �ļ��ѳɹ��ؽ���
)

::========================================================================================================================================

:repairoffice

echo:
echo %line%
echo:
call :_color %Blue% "�����޸� Office ���֤"
echo:

for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PROCESSOR_ARCHITECTURE') do set arch=%%b

if /i "%arch%"=="ARM64" (
echo:
echo �ѷ��� ARM64 Windows��
echo ����Ҫ�� Office �� Windows ������ʹ���޸�ѡ�
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
echo ���ڼ���Ѱ�װ�� Office �汾����
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
echo ���ҵ���� office �汾��
echo ����ֻ��װ office ��һ���汾��
echo ________________________________________________________________
echo:
)

if %counter% EQU 0 (
echo:
echo û���ҵ��Ѱ�װ�� Office��
goto :repairend
echo:
) else (
echo:
call :_color %_Yellow% "������һ�����ڣ��ڴ˴���������Ҫѡ�� [����] �޸�ѡ���"
call :_color %_Yellow% "�밴���������ִ�С���"
echo:
pause %nul1%
)

if defined uwp16 (
echo:
echo ע������������� Office 16.0 UWP ���޸���
echo     ����Ҫ�� Windows ������ʹ������ѡ�
echo ________________________________________________________________
echo:
start ms-settings:appsfeatures
)

set c2r14=
if defined c2r14_68 set c2r14=1
if defined c2r14_86 set c2r14=1

if defined c2r14 (
echo:
echo ע���������� Office 14.0 C2R ���޸�
echo     ����Ҫ�� Windows ������ʹ���޸�ѡ�
echo ________________________________________________________________
echo:
start appwiz.cpl
)

if defined msi14_68 if exist "%msi14repair68%" echo �������� - "%msi14repair68%"                    & "%msi14repair68%"
if defined msi14_86 if exist "%msi14repair86%" echo �������� - "%msi14repair86%"                    & "%msi14repair86%"
if defined msi15_68 if exist "%msi15repair68%" echo �������� - "%msi15repair68%"                    & "%msi15repair68%"
if defined msi15_86 if exist "%msi15repair86%" echo �������� - "%msi15repair86%"                    & "%msi15repair86%"
if defined msi16_68 if exist "%msi16repair68%" echo �������� - "%msi16repair68%"                    & "%msi16repair68%"
if defined msi16_86 if exist "%msi16repair86%" echo �������� - "%msi16repair86%"                    & "%msi16repair86%"
if defined c2r15_68 if exist "%c2r15repair68%" echo �������� - "%c2r15repair68%" REPAIRUI RERUNMODE & "%c2r15repair68%" REPAIRUI RERUNMODE
if defined c2r15_86 if exist "%c2r15repair86%" echo �������� - "%c2r15repair86%" REPAIRUI RERUNMODE & "%c2r15repair86%" REPAIRUI RERUNMODE
if defined c2r16_68 if exist "%c2r16repair68%" echo �������� - "%c2r16repair68%" scenario=Repair    & "%c2r16repair68%" scenario=Repair
if defined c2r16_86 if exist "%c2r16repair86%" echo �������� - "%c2r16repair86%" scenario=Repair    & "%c2r16repair86%" scenario=Repair

:repairend

echo:
echo %line%
echo:
echo:
call :_color %Green% "�����"
goto :at_back

::========================================================================================================================================

:fixwmi

cls
mode 98, 34
title �޸� WMI

::  https://techcommunity.microsoft.com/t5/ask-the-performance-team/wmi-repository-corruption-or-not/ba-p/375484

if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
%eline%
echo �������� Windows Server ���ؽ� WMI��������ֹ����
goto :at_back
)

for %%# in (wmic.exe) do @if "%%~$PATH:#"=="" (
%eline%
echo ��ϵͳ���Ҳ��� WMIC.exe �ļ���������ֹ����
goto :at_back
)

echo:
echo ���ڼ�� WMI
call :checkwmi

::  ����Ӧ�û����޸������

if defined error (
%psc% Stop-Service Winmgmt -force %nul%
winmgmt /salvagerepository %nul%
call :checkwmi
)

if not defined error (
echo [���ڹ���]
echo ����Ӧ�ô�ѡ�������ֹ����
goto :at_back
)

call :_color %Red% "[δ��Ӧ]"

set _corrupt=
sc start Winmgmt %nul%
if %errorlevel% EQU 1060 set _corrupt=1
sc query Winmgmt %nul% || set _corrupt=1
for %%G in (DependOnService Description DisplayName ErrorControl ImagePath ObjectName Start Type) do if not defined _corrupt (reg query HKLM\SYSTEM\CurrentControlSet\Services\Winmgmt /v %%G %nul% || set _corrupt=1)

echo:
if defined _corrupt (
%eline%
echo Winmgmt ����δ��װ��������ֹ����
goto :at_back
)

echo ����ֹͣ Winmgmt ����
sc config Winmgmt start= disabled %nul%
if %errorlevel% EQU 0 (
echo [�ɹ�]
) else (
call :_color %Red% "[ʧ��] ������ֹ����"
sc config Winmgmt start= auto %nul%
goto :at_back
)

echo:
echo ����ֹͣ Winmgmt ����
%psc% Stop-Service Winmgmt -force %nul%
%psc% Stop-Service Winmgmt -force %nul%
%psc% Stop-Service Winmgmt -force %nul%
sc query Winmgmt | find /i "STOPPED" %nul% && (
echo [�ɹ�]
) || (
call :_color %Red% "[ʧ��]"
echo:
call :_color %Blue% "����ѡ�� [��������] ѡ�Ȼ���ٴ�Ӧ���޸� WMI ѡ�"
echo %line%
echo:
choice /C:21 /N /M "> [1] ��������  [2] �ָ����� ��"
if !errorlevel!==1 (sc config Winmgmt start= auto %nul%&goto :at_back)
echo:
echo ������������
shutdown -t 5 -r
exit
)

echo:
echo ����ɾ�� WMI �洢��
rmdir /s /q "%windir%\System32\wbem\repository\" %nul%
if exist "%windir%\System32\wbem\repository\" (
call :_color %Red% "[ʧ��]"
) else (
echo [�ɹ�]
)

echo:
echo �������� Winmgmt ����
sc config Winmgmt start= auto %nul%
if %errorlevel% EQU 0 (
echo [�ɹ�]
) else (
call :_color %Red% "[ʧ��]"
)

call :checkwmi
if not defined error (
echo:
echo ���ڼ�� WMI
call :_color %Green% "[���ڹ���]"
goto :at_back
)

echo:
echo ����ע�� .dll ������ .mof �� .mfl
call :registerobj %nul%

echo:
echo ���ڼ�� WMI
call :checkwmi
if defined error (
call :_color %Red% "[δ��Ӧ]"
echo:
echo ���� [Dism RestoreHealth] �� [SFC Scannow] ѡ���ȷ��û�д���
) else (
call :_color %Green% "[���ڹ���]"
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
call :_color %_Yellow% "�밴��������ء���"
pause %nul1%
goto :at_menu

::========================================================================================================================================

:at_done

echo:
echo �밴�����%_exitmsg%�ű�����
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

::  �˴����� HKLM\SYSTEM\WPA ���Ƿ������Ч��ע������ʹ���������е�ϵͳ��Ҳ���ܻ���ִ����⡣

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

::  �����ǻ�ȡ��ʧ��ע����������Ȩ������ɾ���Ĵ���
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

::  �����������Ĳ�ɫ�ı�
::  ��л @dbenham �� @jeb
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
::  �·���������
