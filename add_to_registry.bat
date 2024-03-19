@echo off
set /p AppPath=Enter the full path to the application: 

echo Windows Registry Editor Version 5.00 > AddToContextMenu.reg
echo. >> AddToContextMenu.reg
echo [HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsx\shell\ProcessWithYourApp] >> AddToContextMenu.reg
echo @="Process with YourApp" >> AddToContextMenu.reg
echo [HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsx\shell\ProcessWithYourApp\command] >> AddToContextMenu.reg
echo @="\"%AppPath%\" \"%%1\"" >> AddToContextMenu.reg

echo Registry file created.
pause
