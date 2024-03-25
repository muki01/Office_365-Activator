mode con: cols=100 lines=20
title Office 365 Activator by Muki
@echo off
chcp 65001 > nul
cls
color 3

net session >nul 2>&1
if %errorLevel% == 0 (
    goto menu
) else (
    goto administatorError
)


:menu
cls
echo.
echo      ____   __  __ _             ____    __ _____                 _   _            _             
echo "   / __ \ / _|/ _(_)           |___ \  / /| ____|      /\       | | (_)          | |            
echo "  | |  | | |_| |_ _  ___ ___     __) |/ /_| |__       /  \   ___| |_ ___   ____ _| |_ ___  _ __ 
echo "  | |  | |  _|  _| |/ __/ _ \   |__ <| '_ \___ \     / /\ \ / __| __| \ \ / / _` | __/ _ \| '__|
echo "  | |__| | | | | | | (_|  __/   ___) | (_) |__) |   / ____ \ (__| |_| |\ V / (_| | || (_) | |   
echo "   \____/|_| |_| |_|\___\___|  |____/ \___/____/   /_/    \_\___|\__|_| \_/ \__,_|\__\___/|_|   
echo      Created by Muki                                                         github.com/muki01
echo.  
echo                                      1. Activate Office 365
echo.
set /p select=">>> "
if %select%==1 goto applyKey



:applyKey
    cls
    echo.
    echo ------------------------
    echo   Activation Started !
    echo   Please wait...
    echo ------------------------
    echo.
    if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" (
        cd /d "%ProgramFiles%\Microsoft Office\Office16"
    ) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" (
        cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
    )
    for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do (
        cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
    )
    cscript //nologo ospp.vbs /setprt:1688 >nul
    cscript //nologo ospp.vbs /unpkey:6F7TH >nul
    cscript //nologo ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH >nul
    set server=1
    goto skms


:skms
    if %server% == 1 set KMS=107.175.77.7
    if %server% == 2 set KMS=kms7.MSGuides.com
    if %server% == 3 goto busyError
    cscript //nologo ospp.vbs /sethst:%KMS% >nul
    goto activate


:activate
    cscript //nologo ospp.vbs /act | find /i "successful" && (
        goto activatedSuccessfully
    ) || (
        goto serverError
    )



:activatedSuccessfully
    echo.
    echo ---------------------------------
    echo   Office Activated Successfully
    echo ---------------------------------
    echo.
    ping localhost -n 5 >nul
    goto menu


:serverError
    echo.
    echo ------------------------------------------
    echo   The connection to the server failed! 
    echo   Trying to connect to another server...
    echo   Please wait...
    echo ------------------------------------------
    echo.
    set /a server+=1
    ping localhost -n 5 >nul
    goto skms

:busyError
    cls
    echo.
    echo ----------------------------------------------------------------
    echo   Sorry, the server is busy and can't respond to your request. 
    echo   Please try again.
    echo ----------------------------------------------------------------
    echo.
    ping localhost -n 5 >nul
    goto menu


:administatorError
    cls
    echo.
    echo ---------------------------------------
    echo   Run this Program As Administrator !
    echo ---------------------------------------
    echo.
    pause
    exit

