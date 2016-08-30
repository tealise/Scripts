@echo off
:: Script by Joshua Nasiatka
:: Office 2013 Activator

:: ===== MODIFY HERE ===== ::
set licenseKeySTD=XXXXX-12345-ABCDE-67890-00000
set licenseKeyPRO=XXXXX-23456-BCDEF-78910-11111


:: ===== DO NOT MODIFY BELOW ===== ::

TITLE Office Activator for 2013
color 1f
cls
:menu
echo ==== MENU ====
echo [1] Standard
echo [2] Pro Plus
echo.
set /p version="Which edition of office? 1/2: "
cls
Echo Activating Office 2013
echo.
echo Please wait . . .
Echo Activating Office 2013									     > C:\office.log
Echo.												    >> C:\office.log
:: Activate Office
if %version% EQU 1 ( 
cscript "C:\Program Files\Microsoft Office\Office15\ospp.vbs" /inpkey:%licenseKeySTD% >> C:\office.log
set versionname="Standard"
) else if %version% EQU 2 (
cscript "C:\Program Files\Microsoft Office\Office15\ospp.vbs" /inpkey:%licenseKeyPRO% >> C:\office.log
set versionname="Pro Plus"
) else (
cls
echo Invalid selection.
goto menu
)
pause
cscript "C:\Program Files\Microsoft Office\Office15\ospp.vbs" /act				    >> C:\office.log
cscript "C:\Program Files\Microsoft Office\Office15\ospp.vbs" /dstatus				    >> C:\office.log
echo.												    >> C:\office.log
echo Activation of Microsoft Office 2013 %versionname% Complete.									    >> C:\office.log
cls
echo **** Activation of Microsoft Office 2013 %versionname% Complete. ****
echo.
echo Activation information located in C:\office.log
echo.
pause