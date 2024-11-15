@echo off
title Activate Microsoft Office 2021
echo =====================================================================================
echo Activating Microsoft Office 2021
echo =====================================================================================

:: Navigate to the appropriate Office installation directory
(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16") 
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")

:: Install licenses
(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)

:: Begin activation process
echo Activating your product...
cscript //nologo slmgr.vbs /ckms >nul
cscript //nologo ospp.vbs /setprt:1688 >nul
cscript //nologo ospp.vbs /unpkey:6F7TH >nul
set i=1
cscript //nologo ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH >nul || goto notsupported

:skms
if %i% GTR 10 goto busy
if %i% EQU 1 set KMS=107.175.77.7
if %i% GTR 1 goto ato
cscript //nologo ospp.vbs /sethst:%KMS% >nul

:ato
cscript //nologo ospp.vbs /act | find /i "successful" && (
    echo =====================================================================================
    echo Activation was successful!
    echo =====================================================================================
    exit
) || (
    echo The connection to the KMS server failed. Trying another one...
    set /a i+=1
    goto skms
)

:notsupported
echo =====================================================================================
echo Sorry, your version is not supported.
echo =====================================================================================
goto halt

:busy
echo =====================================================================================
echo Sorry, the server is busy and cannot respond to your request. Please try again later.
echo =====================================================================================
goto halt

:halt
pause >nul
