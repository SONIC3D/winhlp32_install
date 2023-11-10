@echo off

rem ---------------------------------------
rem Windows 10 用 WinHlp32.exe インストールバッチファイル
rem Ver 3.00
rem ---------------------------------------

rem ---------------------------------------
rem MuiFile(winhlp32.exe.mui)のコピー先フォルダ名(C:\Windows\ja-JP\)の一部を設定
rem ※※※日本語のWindowsではない場合は変更する必要がある※※※
rem ※※※When it isn't Japanese Windows, it's necessary to change it.※※※
rem ---------------------------------------
set ILanguages=ja-JP

rem ---------------------------------------
rem 展開されたフォルダ名1
rem ※※※日本語のWindowsではない場合はここを変更する必要がある※※※
rem ※※※When it isn't Japanese Windows, it's necessary to change it.※※※
rem ---------------------------------------
set MuiFolder_x86=x86_microsoft-windows-winhstb.resources_31bf3856ad364e35_6.3.9600.20470_ja-jp_965b4fed1fd6c2bf
set MuiFolder_x64=amd64_microsoft-windows-winhstb.resources_31bf3856ad364e35_6.3.9600.20470_ja-jp_f279eb70d83433f5

rem ---------------------------------------
rem 展開されたフォルダ名2
rem ---------------------------------------
set ExeDllFolder_x86=x86_microsoft-windows-winhstb_31bf3856ad364e35_6.3.9600.20470_none_be363e6f3e19858c
set ExeDllFolder_x64=amd64_microsoft-windows-winhstb_31bf3856ad364e35_6.3.9600.20470_none_1a54d9f2f676f6c2

rem ---------------------------------------
rem 64bitか32bitかで展開するmsu・cabファイル名を変える
rem ---------------------------------------
set MsuFileName=Windows8.1-KB917607-x86.msu
set CabFileName=Windows8.1-KB917607-x86.cab
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  set MsuFileName=Windows8.1-KB917607-x64.msu
  set CabFileName=Windows8.1-KB917607-x64.cab
)

rem ---------------------------------------
rem 管理者権限チェック
rem ---------------------------------------
for /f "tokens=1 delims=," %%i in ('whoami /groups /FO CSV /NH') do (
  if "%%~i"=="BUILTIN\Administrators" set ADMIN=yes
  if "%%~i"=="Mandatory Label\High Mandatory Level" set ELEVATED=yes
)

if "%ADMIN%" neq "yes" (
  echo 管理者権限で実行して下さい
  echo インストールを中止します
  pause
  goto :eof
)
if "%ELEVATED%" neq "yes" (
  echo 管理者権限で実行して下さい
  echo インストールを中止します
  pause
  goto :eof
)

rem ---------------------------------------
rem Windows 10 かどうかチェック
rem ---------------------------------------
ver | findstr /il "10\.0\." > nul
if %errorlevel% equ 0 (
  goto :InstallStart
)
echo Microsoft Windows 10 [Version 10.0] ではありません
echo インストールを中止します
pause
goto :eof

rem ---------------------------------------
rem インストール開始
rem ---------------------------------------
:InstallStart
echo ●winhlp32 インストール開始
echo.

rem ---------------------------------------
rem カレントディレクトリをこのファイルのあるフォルダに設定
rem ---------------------------------------
cd /d %~dp0

rem ---------------------------------------
rem 展開するmsuファイルがあるかどうか調べる
rem ---------------------------------------
IF NOT EXIST %MsuFileName% (
  echo ファイル %MsuFileName% がありません
  echo インストールを中止します
  pause
  goto :eof
)

rem ---------------------------------------
rem msuファイルの展開先テンポラリフォルダ作成
rem ---------------------------------------
cd /d %~dp0
mkdir .\msutemp

rem ---------------------------------------
rem msuファイルの展開
rem ---------------------------------------
echo ●%MsuFileName% を展開中
echo.
expand.exe -f:*%CabFileName% %MsuFileName% .\msutemp >nul

rem ---------------------------------------
rem 展開するcabファイルがあるかどうか調べる
rem ---------------------------------------
IF NOT EXIST .\msutemp\%CabFileName% (
  echo ファイル %CabFileName% が見つかりません
  echo インストールを中止します
  pause
  goto :eof
)

rem ---------------------------------------
rem cabファイルの展開先テンポラリフォルダ作成
rem ---------------------------------------
cd /d %~dp0
mkdir .\cabtemp

rem ---------------------------------------
rem cabファイルの展開
rem ---------------------------------------
echo ●%CabFileName% を展開中
echo.
expand.exe -f:* .\msutemp\%CabFileName% .\cabtemp >nul

rem ---------------------------------------
rem 必要なファイルが展開されてるか調べる
rem ---------------------------------------
IF NOT EXIST .\cabtemp\%MuiFolder_x86%\winhlp32.exe.mui (
  echo ファイル winhlp32.exe.mui が見つかりません
  echo インストールを中止します
  pause
  goto :eof
)
IF NOT EXIST .\cabtemp\%MuiFolder_x86%\ftsrch.dll.mui (
  echo ファイル ftsrch.dll.mui が見つかりません
  echo インストールを中止します
  pause
  goto :eof
)
IF NOT EXIST .\cabtemp\%ExeDllFolder_x86%\ftlx041e.dll (
  echo ファイル ftlx041e.dll が見つかりません
  echo インストールを中止します
  pause
  goto :eof
)
IF NOT EXIST .\cabtemp\%ExeDllFolder_x86%\ftlx0411.dll (
  echo ファイル ftlx0411.dll が見つかりません
  echo インストールを中止します
  pause
  goto :eof
)
IF NOT EXIST .\cabtemp\%ExeDllFolder_x86%\ftsrch.dll (
  echo ファイル ftsrch.dll が見つかりません
  echo インストールを中止します
  pause
  goto :eof
)
IF NOT EXIST .\cabtemp\%ExeDllFolder_x86%\winhlp32.exe (
  echo ファイル winhlp32.exe が見つかりません
  echo インストールを中止します
  pause
  goto :eof
)
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  IF NOT EXIST .\cabtemp\%MuiFolder_x64%\winhlp32.exe.mui (
    echo ファイル winhlp32.exe.mui が見つかりません
    echo インストールを中止します
    pause
    goto :eof
  )
  IF NOT EXIST .\cabtemp\%MuiFolder_x64%\ftsrch.dll.mui (
    echo ファイル ftsrch.dll.mui が見つかりません
    echo インストールを中止します
    pause
    goto :eof
  )
  IF NOT EXIST .\cabtemp\%ExeDllFolder_x64%\ftlx041e.dll (
    echo ファイル ftlx041e.dll が見つかりません
    echo インストールを中止します
    pause
    goto :eof
  )
  IF NOT EXIST .\cabtemp\%ExeDllFolder_x64%\ftlx0411.dll (
    echo ファイル ftlx0411.dll が見つかりません
    echo インストールを中止します
    pause
    goto :eof
  )
  IF NOT EXIST .\cabtemp\%ExeDllFolder_x64%\ftsrch.dll (
    echo ファイル ftsrch.dll が見つかりません
    echo インストールを中止します
    pause
    goto :eof
  )
  IF NOT EXIST .\cabtemp\%ExeDllFolder_x64%\winhlp32.exe (
    echo ファイル winhlp32.exe が見つかりません
    echo インストールを中止します
    pause
    goto :eof
  )
)

rem ---------------------------------------
rem WinHlp32.exe のタスクが起動していたら強制終了させる
rem ---------------------------------------
taskkill /f /im %ExeFileName% /t 2>nul

rem ---------------------------------------
rem winhlp32.exe.mui を C:\Windows\ja-JP\ へコピー
rem Windowsフォルダ名が環境により変わる事を考慮する
rem ---------------------------------------
echo ●winhlp32.exe.mui ファイルコピー処理中
set CopyFile1=.\cabtemp\%MuiFolder_x86%\winhlp32.exe.mui
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  set CopyFile1=.\cabtemp\%MuiFolder_x64%\winhlp32.exe.mui
)

rem OS側にある winhlp32.exe.mui の所有権を自分に設定
takeown /f "%SystemRoot%\%ILanguages%\winhlp32.exe.mui"
echo.

rem OS側にある winhlp32.exe.mui のフルアクセス権を自分に設定
icacls "%SystemRoot%\%ILanguages%\winhlp32.exe.mui" /grant "%UserName%":F
echo.

rem winhlp32.exe.mui をOS側へ上書きコピー
xcopy /r /y /h /q "%CopyFile1%" "%SystemRoot%\%ILanguages%"
echo.
if errorlevel 1 goto :Error

rem winhlp32.exe.mui の所有権を TrustedInstaller に設定
icacls "%SystemRoot%\%ILanguages%\winhlp32.exe.mui" /setowner "NT Service\TrustedInstaller"
echo.

rem winhlp32.exe.mui のアクセス権から自分を削除
icacls "%SystemRoot%\%ILanguages%\winhlp32.exe.mui" /remove "%UserName%"
echo.

rem ---------------------------------------
rem winhlp32.exe を C:\Windows\ へコピー
rem Windowsフォルダ名が環境により変わる事を考慮する
rem ---------------------------------------
echo ●winhlp32.exe ファイルコピー処理中
set CopyFile2=.\cabtemp\%ExeDllFolder_x86%\winhlp32.exe
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  set CopyFile2=.\cabtemp\%ExeDllFolder_x64%\winhlp32.exe
)

rem OS側にある winhlp32.exe の所有権を自分に設定
takeown /f "%SystemRoot%\winhlp32.exe"
echo.

rem OS側にある winhlp32.exe のフルアクセス権を自分に設定
icacls "%SystemRoot%\winhlp32.exe" /grant "%UserName%":F
echo.

rem winhlp32.exe をOS側へ上書きコピー
xcopy /r /y /h /q "%CopyFile2%" "%SystemRoot%"
echo.
if errorlevel 1 goto :Error

rem winhlp32.exe の所有権を TrustedInstaller に設定
icacls "%SystemRoot%\winhlp32.exe" /setowner "NT Service\TrustedInstaller"
echo.

rem winhlp32.exe のアクセス権から自分を削除
icacls "%SystemRoot%\winhlp32.exe" /remove "%UserName%"
echo.

rem ---------------------------------------
rem その他ファイルをコピー
rem Windows10標準では存在しないはずなので
rem コピー先に既に既にファイルがある場合は念のためコピーしない
rem ---------------------------------------
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  IF NOT EXIST %SystemRoot%\SysWOW64\%ILanguages%\ftsrch.dll.mui (
    echo ●ftsrch.dll.mui ファイルコピー処理中
  
    rem ftsrch.dll.mui をOS側へコピー
    xcopy /r /y /h /q ".\cabtemp\%MuiFolder_x86%\ftsrch.dll.mui" "%SystemRoot%\SysWOW64\%ILanguages%"
    echo.
  
    rem ftsrch.dll.mui の所有権を TrustedInstaller に設定
    icacls "%SystemRoot%\SysWOW64\%ILanguages%\ftsrch.dll.mui" /setowner "NT Service\TrustedInstaller"
    echo.
  
    rem ftsrch.dll.mui のアクセス権から自分を削除
    icacls "%SystemRoot%\SysWOW64\%ILanguages%\ftsrch.dll.mui" /remove "%UserName%"
    echo.
  )
  
  IF NOT EXIST %SystemRoot%\SysWOW64\ftlx041e.dll (
    echo ●ftlx041e.dll ファイルコピー処理中
  
    rem ftlx041e.dll をOS側へコピー
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftlx041e.dll" "%SystemRoot%\SysWOW64"
    echo.
  
    rem ftlx041e.dll の所有権を TrustedInstaller に設定
    icacls "%SystemRoot%\SysWOW64\ftlx041e.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    rem ftlx041e.dll のアクセス権から自分を削除
    icacls "%SystemRoot%\SysWOW64\ftlx041e.dll" /remove "%UserName%"
    echo.
  
  )
  
  IF NOT EXIST %SystemRoot%\SysWOW64\ftlx0411.dll (
    echo ●ftlx0411.dll ファイルコピー処理中
  
    rem ftlx0411.dll をOS側へコピー
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftlx0411.dll" "%SystemRoot%\SysWOW64"
    echo.
  
    rem ftlx0411.dll の所有権を TrustedInstaller に設定
    icacls "%SystemRoot%\SysWOW64\ftlx0411.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    rem ftlx0411.dll のアクセス権から自分を削除
    icacls "%SystemRoot%\SysWOW64\ftlx0411.dll" /remove "%UserName%"
    echo.
  )
  
  IF NOT EXIST %SystemRoot%\SysWOW64\ftsrch.dll (
    echo ●ftsrch.dll ファイルコピー処理中
  
    rem ftsrch.dll をOS側へコピー
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftsrch.dll" "%SystemRoot%\SysWOW64"
    echo.
  
    rem ftsrch.dll の所有権を TrustedInstaller に設定
    icacls "%SystemRoot%\SysWOW64\ftsrch.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    rem ftsrch.dll のアクセス権から自分を削除
    icacls "%SystemRoot%\SysWOW64\ftsrch.dll" /remove "%UserName%"
    echo.
  )


  IF NOT EXIST %SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui (
    echo ●ftsrch.dll.mui ファイルコピー処理中

    rem ftsrch.dll.mui をOS側へコピー
    xcopy /r /y /h /q ".\cabtemp\%MuiFolder_x64%\ftsrch.dll.mui" "%SystemRoot%\System32\%ILanguages%"
    echo.
  
    rem ftsrch.dll.mui の所有権を TrustedInstaller に設定
    icacls "%SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui" /setowner "NT Service\TrustedInstaller"
    echo.

    rem ftsrch.dll.mui のアクセス権から自分を削除
    icacls "%SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui" /remove "%UserName%"
    echo.
  )

  IF NOT EXIST %SystemRoot%\System32\ftlx041e.dll (
    echo ●ftlx041e.dll ファイルコピー処理中

    rem ftlx041e.dll をOS側へコピー
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x64%\ftlx041e.dll" "%SystemRoot%\System32"
    echo.

    rem ftlx041e.dll の所有権を TrustedInstaller に設定
    icacls "%SystemRoot%\System32\ftlx041e.dll" /setowner "NT Service\TrustedInstaller"
    echo.

    rem ftlx041e.dll のアクセス権から自分を削除
    icacls "%SystemRoot%\System32\ftlx041e.dll" /remove "%UserName%"
    echo.
  )

  IF NOT EXIST %SystemRoot%\System32\ftlx0411.dll (
    echo ●ftlx0411.dll ファイルコピー処理中

    rem ftlx0411.dll をOS側へコピー
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x64%\ftlx0411.dll" "%SystemRoot%\System32"
    echo.

    rem ftlx0411.dll の所有権を TrustedInstaller に設定
    icacls "%SystemRoot%\System32\ftlx0411.dll" /setowner "NT Service\TrustedInstaller"
    echo.

    rem ftlx0411.dll のアクセス権から自分を削除
    icacls "%SystemRoot%\System32\ftlx0411.dll" /remove "%UserName%"
    echo.
  )

  IF NOT EXIST %SystemRoot%\System32\ftsrch.dll (
    echo ●ftsrch.dll ファイルコピー処理中

    rem ftsrch.dll をOS側へコピー
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x64%\ftsrch.dll" "%SystemRoot%\System32"
    echo.

    rem ftsrch.dll の所有権を TrustedInstaller に設定
    icacls "%SystemRoot%\System32\ftsrch.dll" /setowner "NT Service\TrustedInstaller"
    echo.

    rem ftsrch.dll のアクセス権から自分を削除
    icacls "%SystemRoot%\System32\ftsrch.dll" /remove "%UserName%"
    echo.
  )
) else (
  IF NOT EXIST %SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui (
    echo ●ftsrch.dll.mui ファイルコピー処理中
  
    rem ftsrch.dll.mui をOS側へコピー
    xcopy /r /y /h /q ".\cabtemp\%MuiFolder_x86%\ftsrch.dll.mui" "%SystemRoot%\System32\%ILanguages%"
    echo.
  
    rem ftsrch.dll.mui の所有権を TrustedInstaller に設定
    icacls "%SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui" /setowner "NT Service\TrustedInstaller"
    echo.
  
    rem ftsrch.dll.mui のアクセス権から自分を削除
    icacls "%SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui" /remove "%UserName%"
    echo.
  )
  
  IF NOT EXIST %SystemRoot%\System32\ftlx041e.dll (
    echo ●ftlx041e.dll ファイルコピー処理中
  
    rem ftlx041e.dll をOS側へコピー
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftlx041e.dll" "%SystemRoot%\System32"
    echo.
  
    rem ftlx041e.dll の所有権を TrustedInstaller に設定
    icacls "%SystemRoot%\System32\ftlx041e.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    rem ftlx041e.dll のアクセス権から自分を削除
    icacls "%SystemRoot%\System32\ftlx041e.dll" /remove "%UserName%"
    echo.
  
  )
  
  IF NOT EXIST %SystemRoot%\System32\ftlx0411.dll (
    echo ●ftlx0411.dll ファイルコピー処理中
  
    rem ftlx0411.dll をOS側へコピー
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftlx0411.dll" "%SystemRoot%\System32"
    echo.
  
    rem ftlx0411.dll の所有権を TrustedInstaller に設定
    icacls "%SystemRoot%\System32\ftlx0411.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    rem ftlx0411.dll のアクセス権から自分を削除
    icacls "%SystemRoot%\System32\ftlx0411.dll" /remove "%UserName%"
    echo.
  )
  
  IF NOT EXIST %SystemRoot%\System32\ftsrch.dll (
    echo ●ftsrch.dll ファイルコピー処理中
  
    rem ftsrch.dll をOS側へコピー
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftsrch.dll" "%SystemRoot%\System32"
    echo.
  
    rem ftsrch.dll の所有権を TrustedInstaller に設定
    icacls "%SystemRoot%\System32\ftsrch.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    rem ftsrch.dll のアクセス権から自分を削除
    icacls "%SystemRoot%\System32\ftsrch.dll" /remove "%UserName%"
    echo.
  )
)

rem ---------------------------------------
rem テンポラリフォルダ削除
rem ---------------------------------------
cd /d %~dp0
rd /s /q .\cabtemp >nul
rd /s /q .\msutemp >nul

rem ---------------------------------------
rem レジストリを修正
rem ---------------------------------------
echo ●レジストリ修正中
reg delete "HKLM\SOFTWARE\Microsoft\WinHelp" /f 2>nul
echo.
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\WinHelp" /f 2>nul
  echo.
)
rem ---------------------------------------
rem レジストリ追加
rem AllowProgrammaticMacros…マクロを有効にする
rem AllowIntranetAccess…イントラネット上に格納されている.hlpファイルのブロック解除
rem セキュリティー低下を招く可能性があるので標準で追加するのはやめる事にする
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

echo ●インストールが正常に完了しました
pause
goto :eof

:Error
echo ●エラーが発生しました、インストールを中止します
pause
goto :eof
