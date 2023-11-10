@echo off

rem ---------------------------------------
rem Windows 10 WinHlp32.exe installation batch file
rem Ver 3.00
rem ---------------------------------------

rem ---------------------------------------
rem MuiFile(winhlp32.exe.mui) location directory name
rem It's a language file directory name inside the Windows installation directory. (e.g. C:\Windows\ja-JP\ for Japanese OS and C:\Windows\zh-CN for Simplified Chinese OS)
rem ---------------------------------------
set ILanguages=zh-CN

rem ---------------------------------------
rem Muifile directory name in the Microsoft update pack
rem Modified it based on your system language. You may need to manually unpack the msi and cab to know the correct names.
rem ---------------------------------------
set MuiFolder_x86=x86_microsoft-windows-winhstb.resources_31bf3856ad364e35_6.3.9600.20470_zh-cn_c711722449b29d71
set MuiFolder_x64=amd64_microsoft-windows-winhstb.resources_31bf3856ad364e35_6.3.9600.20470_zh-cn_23300da802100ea7

rem ---------------------------------------
rem Common files directory in the Microsoft update pack
rem ---------------------------------------
set ExeDllFolder_x86=x86_microsoft-windows-winhstb_31bf3856ad364e35_6.3.9600.20470_none_be363e6f3e19858c
set ExeDllFolder_x64=amd64_microsoft-windows-winhstb_31bf3856ad364e35_6.3.9600.20470_none_1a54d9f2f676f6c2

rem ---------------------------------------
rem 64bit and 32bit OS msu and cab filename
rem ---------------------------------------
set MsuFileName=Windows8.1-KB917607-x86.msu
set CabFileName=Windows8.1-KB917607-x86.cab
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  set MsuFileName=Windows8.1-KB917607-x64.msu
  set CabFileName=Windows8.1-KB917607-x64.cab
)

rem ---------------------------------------
rem Administrator privilege check
rem ---------------------------------------
for /f "tokens=1 delims=," %%i in ('whoami /groups /FO CSV /NH') do (
  if "%%~i"=="BUILTIN\Administrators" set ADMIN=yes
  if "%%~i"=="Mandatory Label\High Mandatory Level" set ELEVATED=yes
)

if "%ADMIN%" neq "yes" (
  echo Please run as Administrator.
  echo Installation aborted.
  pause
  goto :eof
)
if "%ELEVATED%" neq "yes" (
  echo Please run as Administrator.
  echo Installation aborted.
  pause
  goto :eof
)

rem ---------------------------------------
rem Windows 10 version check
rem ---------------------------------------
ver | findstr /il "10\.0\." > nul
if %errorlevel% equ 0 (
  goto :InstallStart
)
echo Microsoft Windows 10 [Version 10.0] is required
echo Installation aborted.
pause
goto :eof

rem ---------------------------------------
rem Installation start
rem ---------------------------------------
:InstallStart
echo Start WinHlp32 installation
echo.

rem ---------------------------------------
rem Set current directory
rem ---------------------------------------
cd /d %~dp0

rem ---------------------------------------
rem Chec if msu file exist
rem ---------------------------------------
IF NOT EXIST %MsuFileName% (
  echo Can not locate %MsuFileName%
  echo Installation aborted.
  pause
  goto :eof
)

rem ---------------------------------------
rem Create msu extraction directory
rem ---------------------------------------
cd /d %~dp0
mkdir .\msutemp

rem ---------------------------------------
rem msu extraction
rem ---------------------------------------
echo Extracting %MsuFileName%
echo.
expand.exe -f:*%CabFileName% %MsuFileName% .\msutemp >nul

rem ---------------------------------------
rem Check if cab file exist
rem ---------------------------------------
IF NOT EXIST .\msutemp\%CabFileName% (
  echo Can not locate %CabFileName%
  echo Installation aborted.
  pause
  goto :eof
)

rem ---------------------------------------
rem Create cab extraction directory
rem ---------------------------------------
cd /d %~dp0
mkdir .\cabtemp

rem ---------------------------------------
rem cab extraction
rem ---------------------------------------
echo Extracting %CabFileName%
echo.
expand.exe -f:* .\msutemp\%CabFileName% .\cabtemp >nul

rem ---------------------------------------
rem Verify extracted files
rem ---------------------------------------
IF NOT EXIST .\cabtemp\%MuiFolder_x86%\winhlp32.exe.mui (
  echo winhlp32.exe.mui not found
  echo Installation aborted.
  pause
  goto :eof
)
IF NOT EXIST .\cabtemp\%MuiFolder_x86%\ftsrch.dll.mui (
  echo ftsrch.dll.mui not found
  echo Installation aborted.
  pause
  goto :eof
)
IF NOT EXIST .\cabtemp\%ExeDllFolder_x86%\ftlx041e.dll (
  echo ftlx041e.dll not found
  echo Installation aborted.
  pause
  goto :eof
)
IF NOT EXIST .\cabtemp\%ExeDllFolder_x86%\ftlx0411.dll (
  echo ftlx0411.dll not found
  echo Installation aborted.
  pause
  goto :eof
)
IF NOT EXIST .\cabtemp\%ExeDllFolder_x86%\ftsrch.dll (
  echo ftsrch.dll not found
  echo Installation aborted.
  pause
  goto :eof
)
IF NOT EXIST .\cabtemp\%ExeDllFolder_x86%\winhlp32.exe (
  echo winhlp32.exe not found
  echo Installation aborted.
  pause
  goto :eof
)
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  IF NOT EXIST .\cabtemp\%MuiFolder_x64%\winhlp32.exe.mui (
    echo winhlp32.exe.mui not found
    echo Installation aborted.
    pause
    goto :eof
  )
  IF NOT EXIST .\cabtemp\%MuiFolder_x64%\ftsrch.dll.mui (
    echo ftsrch.dll.mui not found
    echo Installation aborted.
    pause
    goto :eof
  )
  IF NOT EXIST .\cabtemp\%ExeDllFolder_x64%\ftlx041e.dll (
    echo ftlx041e.dll not found
    echo Installation aborted.
    pause
    goto :eof
  )
  IF NOT EXIST .\cabtemp\%ExeDllFolder_x64%\ftlx0411.dll (
    echo ftlx0411.dll not found
    echo Installation aborted.
    pause
    goto :eof
  )
  IF NOT EXIST .\cabtemp\%ExeDllFolder_x64%\ftsrch.dll (
    echo ftsrch.dll not found
    echo Installation aborted.
    pause
    goto :eof
  )
  IF NOT EXIST .\cabtemp\%ExeDllFolder_x64%\winhlp32.exe (
    echo winhlp32.exe not found
    echo Installation aborted.
    pause
    goto :eof
  )
)

rem ---------------------------------------
rem Kill WinHlp32.exe in running
rem ---------------------------------------
taskkill /f /im %ExeFileName% /t 2>nul

rem ---------------------------------------
rem Copy winhlp32.exe.mui to the MUI directory
rem ---------------------------------------
echo Copying winhlp32.exe.mui file
set CopyFile1=.\cabtemp\%MuiFolder_x86%\winhlp32.exe.mui
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  set CopyFile1=.\cabtemp\%MuiFolder_x64%\winhlp32.exe.mui
)

rem Take the ownership of winhlp32.exe.mui
takeown /f "%SystemRoot%\%ILanguages%\winhlp32.exe.mui"
echo.

rem Grant permission of winhlp32.exe.mui to current user
icacls "%SystemRoot%\%ILanguages%\winhlp32.exe.mui" /grant "%UserName%":F
echo.

rem Copy and overwrite winhlp32.exe.mui
xcopy /r /y /h /q "%CopyFile1%" "%SystemRoot%\%ILanguages%"
echo.
if errorlevel 1 goto :Error

rem Set the owner of winhlp32.exe.mui to TrustedInstaller
icacls "%SystemRoot%\%ILanguages%\winhlp32.exe.mui" /setowner "NT Service\TrustedInstaller"
echo.

rem Remove current user permission from winhlp32.exe.mui
icacls "%SystemRoot%\%ILanguages%\winhlp32.exe.mui" /remove "%UserName%"
echo.

rem ---------------------------------------
rem Overwrite the winhlp32.exe in C:\Windows\
rem ---------------------------------------
echo ●winhlp32.exe ファイルコピー処理中
set CopyFile2=.\cabtemp\%ExeDllFolder_x86%\winhlp32.exe
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  set CopyFile2=.\cabtemp\%ExeDllFolder_x64%\winhlp32.exe
)

takeown /f "%SystemRoot%\winhlp32.exe"
echo.

icacls "%SystemRoot%\winhlp32.exe" /grant "%UserName%":F
echo.

xcopy /r /y /h /q "%CopyFile2%" "%SystemRoot%"
echo.
if errorlevel 1 goto :Error

icacls "%SystemRoot%\winhlp32.exe" /setowner "NT Service\TrustedInstaller"
echo.

icacls "%SystemRoot%\winhlp32.exe" /remove "%UserName%"
echo.

rem ---------------------------------------
rem Copy other files
rem These files should not exist in default Windows10 installation
rem ---------------------------------------
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  IF NOT EXIST %SystemRoot%\SysWOW64\%ILanguages%\ftsrch.dll.mui (
    echo Copying ftsrch.dll.mui
  
    xcopy /r /y /h /q ".\cabtemp\%MuiFolder_x86%\ftsrch.dll.mui" "%SystemRoot%\SysWOW64\%ILanguages%"
    echo.
  
    icacls "%SystemRoot%\SysWOW64\%ILanguages%\ftsrch.dll.mui" /setowner "NT Service\TrustedInstaller"
    echo.
  
    icacls "%SystemRoot%\SysWOW64\%ILanguages%\ftsrch.dll.mui" /remove "%UserName%"
    echo.
  )
  
  IF NOT EXIST %SystemRoot%\SysWOW64\ftlx041e.dll (
    echo Copying ftlx041e.dll
  
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftlx041e.dll" "%SystemRoot%\SysWOW64"
    echo.
  
    icacls "%SystemRoot%\SysWOW64\ftlx041e.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    icacls "%SystemRoot%\SysWOW64\ftlx041e.dll" /remove "%UserName%"
    echo.
  
  )
  
  IF NOT EXIST %SystemRoot%\SysWOW64\ftlx0411.dll (
    echo Copying ftlx0411.dll
  
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftlx0411.dll" "%SystemRoot%\SysWOW64"
    echo.
  
    icacls "%SystemRoot%\SysWOW64\ftlx0411.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    icacls "%SystemRoot%\SysWOW64\ftlx0411.dll" /remove "%UserName%"
    echo.
  )
  
  IF NOT EXIST %SystemRoot%\SysWOW64\ftsrch.dll (
    echo Copying ftsrch.dll
  
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftsrch.dll" "%SystemRoot%\SysWOW64"
    echo.
  
    icacls "%SystemRoot%\SysWOW64\ftsrch.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    icacls "%SystemRoot%\SysWOW64\ftsrch.dll" /remove "%UserName%"
    echo.
  )


  IF NOT EXIST %SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui (
    echo Copying ftsrch.dll.mui

    xcopy /r /y /h /q ".\cabtemp\%MuiFolder_x64%\ftsrch.dll.mui" "%SystemRoot%\System32\%ILanguages%"
    echo.
  
    icacls "%SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui" /setowner "NT Service\TrustedInstaller"
    echo.

    icacls "%SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui" /remove "%UserName%"
    echo.
  )

  IF NOT EXIST %SystemRoot%\System32\ftlx041e.dll (
    echo Copying ftlx041e.dll

    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x64%\ftlx041e.dll" "%SystemRoot%\System32"
    echo.

    icacls "%SystemRoot%\System32\ftlx041e.dll" /setowner "NT Service\TrustedInstaller"
    echo.

    icacls "%SystemRoot%\System32\ftlx041e.dll" /remove "%UserName%"
    echo.
  )

  IF NOT EXIST %SystemRoot%\System32\ftlx0411.dll (
    echo Copying ftlx0411.dll

    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x64%\ftlx0411.dll" "%SystemRoot%\System32"
    echo.

    icacls "%SystemRoot%\System32\ftlx0411.dll" /setowner "NT Service\TrustedInstaller"
    echo.

    icacls "%SystemRoot%\System32\ftlx0411.dll" /remove "%UserName%"
    echo.
  )

  IF NOT EXIST %SystemRoot%\System32\ftsrch.dll (
    echo Copying ftsrch.dll

    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x64%\ftsrch.dll" "%SystemRoot%\System32"
    echo.

    icacls "%SystemRoot%\System32\ftsrch.dll" /setowner "NT Service\TrustedInstaller"
    echo.

    icacls "%SystemRoot%\System32\ftsrch.dll" /remove "%UserName%"
    echo.
  )
) else (
  IF NOT EXIST %SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui (
    echo Copying ftsrch.dll.mui
  
    xcopy /r /y /h /q ".\cabtemp\%MuiFolder_x86%\ftsrch.dll.mui" "%SystemRoot%\System32\%ILanguages%"
    echo.
  
    icacls "%SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui" /setowner "NT Service\TrustedInstaller"
    echo.
  
    icacls "%SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui" /remove "%UserName%"
    echo.
  )
  
  IF NOT EXIST %SystemRoot%\System32\ftlx041e.dll (
    echo Copying ftlx041e.dll
  
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftlx041e.dll" "%SystemRoot%\System32"
    echo.
  
    icacls "%SystemRoot%\System32\ftlx041e.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    icacls "%SystemRoot%\System32\ftlx041e.dll" /remove "%UserName%"
    echo.
  
  )
  
  IF NOT EXIST %SystemRoot%\System32\ftlx0411.dll (
    echo Copying ftlx0411.dll
  
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftlx0411.dll" "%SystemRoot%\System32"
    echo.
  
    icacls "%SystemRoot%\System32\ftlx0411.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    icacls "%SystemRoot%\System32\ftlx0411.dll" /remove "%UserName%"
    echo.
  )
  
  IF NOT EXIST %SystemRoot%\System32\ftsrch.dll (
    echo Copying ftsrch.dll
  
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftsrch.dll" "%SystemRoot%\System32"
    echo.
  
    icacls "%SystemRoot%\System32\ftsrch.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    icacls "%SystemRoot%\System32\ftsrch.dll" /remove "%UserName%"
    echo.
  )
)

rem ---------------------------------------
rem Delete temporary files
rem ---------------------------------------
cd /d %~dp0
rd /s /q .\cabtemp >nul
rd /s /q .\msutemp >nul

rem ---------------------------------------
rem Fix the registry
rem ---------------------------------------
echo Fix the registry
reg delete "HKLM\SOFTWARE\Microsoft\WinHelp" /f 2>nul
echo.
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\WinHelp" /f 2>nul
  echo.
)
rem ---------------------------------------
rem Add registry key
rem AllowProgrammaticMacros        Enable Macro
rem AllowIntranetAccess        Unblocking .hlp files stored on the intranet
rem These are not added by default in this batch file, for it will reduce the security
rem ---------------------------------------
rem reg add "HKLM\SOFTWARE\Microsoft\WinHelp" /v "AllowProgrammaticMacros" /t REG_DWORD /d "0x00000001" /f
rem echo.
rem if errorlevel 1 goto :Error
rem reg add "HKLM\SOFTWARE\Microsoft\WinHelp" /v "AllowIntranetAccess" /t REG_DWORD /d "0x00000001" /f
rem echo.
rem if errorlevel 1 goto :Error
rem if %PROCESSOR_ARCHITECTURE%==AMD64 (
rem   reg add "HKLM\SOFTWARE\Wow6432Node\Microsoft\WinHelp" /v "AllowProgrammaticMacros" /t REG_DWORD /d "0x00000001" /f
rem   echo.
rem   if errorlevel 1 goto :Error
rem   reg add "HKLM\SOFTWARE\Wow6432Node\Microsoft\WinHelp" /v "AllowIntranetAccess" /t REG_DWORD /d "0x00000001" /f
rem   echo.
rem   if errorlevel 1 goto :Error
rem )
rem ---------------------------------------

echo Installation completed successfully.
pause
goto :eof

:Error
echo Error occurred, installation aborted.
pause
goto :eof
