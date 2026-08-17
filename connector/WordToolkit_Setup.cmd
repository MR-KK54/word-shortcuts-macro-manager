@echo off
title Microsoft Word Toolkit - Clean Setup (.cmd)
color 0A
echo ====================================================
echo  Microsoft Word Toolkit - Clean Setup
echo ====================================================
echo.

echo 1. Terminating open Word instances to unlock Normal.dotm...
taskkill /F /IM winword.exe /T 2>nul
timeout /t 1 /nobreak >nul

echo 2. Enabling Word VBA Trust Access in Windows Registry...
reg add "HKCU\Software\Microsoft\Office\16.0\Word\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\15.0\Word\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\14.0\Word\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\12.0\Word\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul 2>&1

reg add "HKCU\Software\Microsoft\Office\16.0\Word\Security" /v Level /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\Word\Security" /v VBAWarnings /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\Word\Security" /v DisableAllMacros /t REG_DWORD /d 0 /f >nul 2>&1

echo 3. Registering wordtoolkit:// 1-click browser sync protocol...
reg add "HKCU\Software\Classes\wordtoolkit" /d "URL:WordToolkit Protocol" /f >nul 2>&1
reg add "HKCU\Software\Classes\wordtoolkit" /v "URL Protocol" /t REG_SZ /d "" /f >nul 2>&1
reg add "HKCU\Software\Classes\wordtoolkit\shell\open\command" /t REG_SZ /d "\"%SystemRoot%\System32\wscript.exe\" \"%APPDATA%\WordToolkit\sync-handler.vbs\" \"%%1\"" /f >nul 2>&1

echo 4. Creating %APPDATA%\WordToolkit folder...
if not exist "%APPDATA%\WordToolkit" mkdir "%APPDATA%\WordToolkit"

echo 5. Downloading connector files...
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object System.Net.WebClient).DownloadFile('http://localhost:3000/api/connector/Toolkit_Helpers.bas', '$env:APPDATA\WordToolkit\Toolkit_Helpers.bas')" >nul 2>&1
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object System.Net.WebClient).DownloadFile('http://localhost:3000/api/connector/Toolkit_Macros.bas', '$env:APPDATA\WordToolkit\Toolkit_Macros.bas')" >nul 2>&1
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object System.Net.WebClient).DownloadFile('http://localhost:3000/api/connector/Toolkit_Menu.bas', '$env:APPDATA\WordToolkit\Toolkit_Menu.bas')" >nul 2>&1
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object System.Net.WebClient).DownloadFile('http://localhost:3000/api/connector/Toolkit_RibbonQAT.bas', '$env:APPDATA\WordToolkit\Toolkit_RibbonQAT.bas')" >nul 2>&1
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object System.Net.WebClient).DownloadFile('http://localhost:3000/api/connector/Toolkit_Shortcuts.bas', '$env:APPDATA\WordToolkit\Toolkit_Shortcuts.bas')" >nul 2>&1
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object System.Net.WebClient).DownloadFile('http://localhost:3000/api/connector/Toolkit_Sync.bas', '$env:APPDATA\WordToolkit\Toolkit_Sync.bas')" >nul 2>&1
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object System.Net.WebClient).DownloadFile('http://localhost:3000/api/connector/sync-handler.vbs', '$env:APPDATA\WordToolkit\sync-handler.vbs')" >nul 2>&1

echo 6. Importing connector modules into Word Normal template...
if exist "%APPDATA%\WordToolkit\sync-handler.vbs" (
    cscript //nologo "%APPDATA%\WordToolkit\sync-handler.vbs" "act=enable" >nul 2>&1
)

echo.
echo ====================================================
echo  🎉 Setup Complete! Microsoft Word is ready.
echo ====================================================
echo.
pause
