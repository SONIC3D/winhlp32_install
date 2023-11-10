@echo off

rem ---------------------------------------
rem Windows 10 �p WinHlp32.exe �C���X�g�[���o�b�`�t�@�C��
rem Ver 3.00
rem ---------------------------------------

rem ---------------------------------------
rem MuiFile(winhlp32.exe.mui)�̃R�s�[��t�H���_��(C:\Windows\ja-JP\)�̈ꕔ��ݒ�
rem ���������{���Windows�ł͂Ȃ��ꍇ�͕ύX����K�v�����遦����
rem ������When it isn't Japanese Windows, it's necessary to change it.������
rem ---------------------------------------
set ILanguages=ja-JP

rem ---------------------------------------
rem �W�J���ꂽ�t�H���_��1
rem ���������{���Windows�ł͂Ȃ��ꍇ�͂�����ύX����K�v�����遦����
rem ������When it isn't Japanese Windows, it's necessary to change it.������
rem ---------------------------------------
set MuiFolder_x86=x86_microsoft-windows-winhstb.resources_31bf3856ad364e35_6.3.9600.20470_ja-jp_965b4fed1fd6c2bf
set MuiFolder_x64=amd64_microsoft-windows-winhstb.resources_31bf3856ad364e35_6.3.9600.20470_ja-jp_f279eb70d83433f5

rem ---------------------------------------
rem �W�J���ꂽ�t�H���_��2
rem ---------------------------------------
set ExeDllFolder_x86=x86_microsoft-windows-winhstb_31bf3856ad364e35_6.3.9600.20470_none_be363e6f3e19858c
set ExeDllFolder_x64=amd64_microsoft-windows-winhstb_31bf3856ad364e35_6.3.9600.20470_none_1a54d9f2f676f6c2

rem ---------------------------------------
rem 64bit��32bit���œW�J����msu�Ecab�t�@�C������ς���
rem ---------------------------------------
set MsuFileName=Windows8.1-KB917607-x86.msu
set CabFileName=Windows8.1-KB917607-x86.cab
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  set MsuFileName=Windows8.1-KB917607-x64.msu
  set CabFileName=Windows8.1-KB917607-x64.cab
)

rem ---------------------------------------
rem �Ǘ��Ҍ����`�F�b�N
rem ---------------------------------------
for /f "tokens=1 delims=," %%i in ('whoami /groups /FO CSV /NH') do (
  if "%%~i"=="BUILTIN\Administrators" set ADMIN=yes
  if "%%~i"=="Mandatory Label\High Mandatory Level" set ELEVATED=yes
)

if "%ADMIN%" neq "yes" (
  echo �Ǘ��Ҍ����Ŏ��s���ĉ�����
  echo �C���X�g�[���𒆎~���܂�
  pause
  goto :eof
)
if "%ELEVATED%" neq "yes" (
  echo �Ǘ��Ҍ����Ŏ��s���ĉ�����
  echo �C���X�g�[���𒆎~���܂�
  pause
  goto :eof
)

rem ---------------------------------------
rem Windows 10 ���ǂ����`�F�b�N
rem ---------------------------------------
ver | findstr /il "10\.0\." > nul
if %errorlevel% equ 0 (
  goto :InstallStart
)
echo Microsoft Windows 10 [Version 10.0] �ł͂���܂���
echo �C���X�g�[���𒆎~���܂�
pause
goto :eof

rem ---------------------------------------
rem �C���X�g�[���J�n
rem ---------------------------------------
:InstallStart
echo ��winhlp32 �C���X�g�[���J�n
echo.

rem ---------------------------------------
rem �J�����g�f�B���N�g�������̃t�@�C���̂���t�H���_�ɐݒ�
rem ---------------------------------------
cd /d %~dp0

rem ---------------------------------------
rem �W�J����msu�t�@�C�������邩�ǂ������ׂ�
rem ---------------------------------------
IF NOT EXIST %MsuFileName% (
  echo �t�@�C�� %MsuFileName% ������܂���
  echo �C���X�g�[���𒆎~���܂�
  pause
  goto :eof
)

rem ---------------------------------------
rem msu�t�@�C���̓W�J��e���|�����t�H���_�쐬
rem ---------------------------------------
cd /d %~dp0
mkdir .\msutemp

rem ---------------------------------------
rem msu�t�@�C���̓W�J
rem ---------------------------------------
echo ��%MsuFileName% ��W�J��
echo.
expand.exe -f:*%CabFileName% %MsuFileName% .\msutemp >nul

rem ---------------------------------------
rem �W�J����cab�t�@�C�������邩�ǂ������ׂ�
rem ---------------------------------------
IF NOT EXIST .\msutemp\%CabFileName% (
  echo �t�@�C�� %CabFileName% ��������܂���
  echo �C���X�g�[���𒆎~���܂�
  pause
  goto :eof
)

rem ---------------------------------------
rem cab�t�@�C���̓W�J��e���|�����t�H���_�쐬
rem ---------------------------------------
cd /d %~dp0
mkdir .\cabtemp

rem ---------------------------------------
rem cab�t�@�C���̓W�J
rem ---------------------------------------
echo ��%CabFileName% ��W�J��
echo.
expand.exe -f:* .\msutemp\%CabFileName% .\cabtemp >nul

rem ---------------------------------------
rem �K�v�ȃt�@�C�����W�J����Ă邩���ׂ�
rem ---------------------------------------
IF NOT EXIST .\cabtemp\%MuiFolder_x86%\winhlp32.exe.mui (
  echo �t�@�C�� winhlp32.exe.mui ��������܂���
  echo �C���X�g�[���𒆎~���܂�
  pause
  goto :eof
)
IF NOT EXIST .\cabtemp\%MuiFolder_x86%\ftsrch.dll.mui (
  echo �t�@�C�� ftsrch.dll.mui ��������܂���
  echo �C���X�g�[���𒆎~���܂�
  pause
  goto :eof
)
IF NOT EXIST .\cabtemp\%ExeDllFolder_x86%\ftlx041e.dll (
  echo �t�@�C�� ftlx041e.dll ��������܂���
  echo �C���X�g�[���𒆎~���܂�
  pause
  goto :eof
)
IF NOT EXIST .\cabtemp\%ExeDllFolder_x86%\ftlx0411.dll (
  echo �t�@�C�� ftlx0411.dll ��������܂���
  echo �C���X�g�[���𒆎~���܂�
  pause
  goto :eof
)
IF NOT EXIST .\cabtemp\%ExeDllFolder_x86%\ftsrch.dll (
  echo �t�@�C�� ftsrch.dll ��������܂���
  echo �C���X�g�[���𒆎~���܂�
  pause
  goto :eof
)
IF NOT EXIST .\cabtemp\%ExeDllFolder_x86%\winhlp32.exe (
  echo �t�@�C�� winhlp32.exe ��������܂���
  echo �C���X�g�[���𒆎~���܂�
  pause
  goto :eof
)
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  IF NOT EXIST .\cabtemp\%MuiFolder_x64%\winhlp32.exe.mui (
    echo �t�@�C�� winhlp32.exe.mui ��������܂���
    echo �C���X�g�[���𒆎~���܂�
    pause
    goto :eof
  )
  IF NOT EXIST .\cabtemp\%MuiFolder_x64%\ftsrch.dll.mui (
    echo �t�@�C�� ftsrch.dll.mui ��������܂���
    echo �C���X�g�[���𒆎~���܂�
    pause
    goto :eof
  )
  IF NOT EXIST .\cabtemp\%ExeDllFolder_x64%\ftlx041e.dll (
    echo �t�@�C�� ftlx041e.dll ��������܂���
    echo �C���X�g�[���𒆎~���܂�
    pause
    goto :eof
  )
  IF NOT EXIST .\cabtemp\%ExeDllFolder_x64%\ftlx0411.dll (
    echo �t�@�C�� ftlx0411.dll ��������܂���
    echo �C���X�g�[���𒆎~���܂�
    pause
    goto :eof
  )
  IF NOT EXIST .\cabtemp\%ExeDllFolder_x64%\ftsrch.dll (
    echo �t�@�C�� ftsrch.dll ��������܂���
    echo �C���X�g�[���𒆎~���܂�
    pause
    goto :eof
  )
  IF NOT EXIST .\cabtemp\%ExeDllFolder_x64%\winhlp32.exe (
    echo �t�@�C�� winhlp32.exe ��������܂���
    echo �C���X�g�[���𒆎~���܂�
    pause
    goto :eof
  )
)

rem ---------------------------------------
rem WinHlp32.exe �̃^�X�N���N�����Ă����狭���I��������
rem ---------------------------------------
taskkill /f /im %ExeFileName% /t 2>nul

rem ---------------------------------------
rem winhlp32.exe.mui �� C:\Windows\ja-JP\ �փR�s�[
rem Windows�t�H���_�������ɂ��ς�鎖���l������
rem ---------------------------------------
echo ��winhlp32.exe.mui �t�@�C���R�s�[������
set CopyFile1=.\cabtemp\%MuiFolder_x86%\winhlp32.exe.mui
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  set CopyFile1=.\cabtemp\%MuiFolder_x64%\winhlp32.exe.mui
)

rem OS���ɂ��� winhlp32.exe.mui �̏��L���������ɐݒ�
takeown /f "%SystemRoot%\%ILanguages%\winhlp32.exe.mui"
echo.

rem OS���ɂ��� winhlp32.exe.mui �̃t���A�N�Z�X���������ɐݒ�
icacls "%SystemRoot%\%ILanguages%\winhlp32.exe.mui" /grant "%UserName%":F
echo.

rem winhlp32.exe.mui ��OS���֏㏑���R�s�[
xcopy /r /y /h /q "%CopyFile1%" "%SystemRoot%\%ILanguages%"
echo.
if errorlevel 1 goto :Error

rem winhlp32.exe.mui �̏��L���� TrustedInstaller �ɐݒ�
icacls "%SystemRoot%\%ILanguages%\winhlp32.exe.mui" /setowner "NT Service\TrustedInstaller"
echo.

rem winhlp32.exe.mui �̃A�N�Z�X�����玩�����폜
icacls "%SystemRoot%\%ILanguages%\winhlp32.exe.mui" /remove "%UserName%"
echo.

rem ---------------------------------------
rem winhlp32.exe �� C:\Windows\ �փR�s�[
rem Windows�t�H���_�������ɂ��ς�鎖���l������
rem ---------------------------------------
echo ��winhlp32.exe �t�@�C���R�s�[������
set CopyFile2=.\cabtemp\%ExeDllFolder_x86%\winhlp32.exe
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  set CopyFile2=.\cabtemp\%ExeDllFolder_x64%\winhlp32.exe
)

rem OS���ɂ��� winhlp32.exe �̏��L���������ɐݒ�
takeown /f "%SystemRoot%\winhlp32.exe"
echo.

rem OS���ɂ��� winhlp32.exe �̃t���A�N�Z�X���������ɐݒ�
icacls "%SystemRoot%\winhlp32.exe" /grant "%UserName%":F
echo.

rem winhlp32.exe ��OS���֏㏑���R�s�[
xcopy /r /y /h /q "%CopyFile2%" "%SystemRoot%"
echo.
if errorlevel 1 goto :Error

rem winhlp32.exe �̏��L���� TrustedInstaller �ɐݒ�
icacls "%SystemRoot%\winhlp32.exe" /setowner "NT Service\TrustedInstaller"
echo.

rem winhlp32.exe �̃A�N�Z�X�����玩�����폜
icacls "%SystemRoot%\winhlp32.exe" /remove "%UserName%"
echo.

rem ---------------------------------------
rem ���̑��t�@�C�����R�s�[
rem Windows10�W���ł͑��݂��Ȃ��͂��Ȃ̂�
rem �R�s�[��Ɋ��Ɋ��Ƀt�@�C��������ꍇ�͔O�̂��߃R�s�[���Ȃ�
rem ---------------------------------------
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  IF NOT EXIST %SystemRoot%\SysWOW64\%ILanguages%\ftsrch.dll.mui (
    echo ��ftsrch.dll.mui �t�@�C���R�s�[������
  
    rem ftsrch.dll.mui ��OS���փR�s�[
    xcopy /r /y /h /q ".\cabtemp\%MuiFolder_x86%\ftsrch.dll.mui" "%SystemRoot%\SysWOW64\%ILanguages%"
    echo.
  
    rem ftsrch.dll.mui �̏��L���� TrustedInstaller �ɐݒ�
    icacls "%SystemRoot%\SysWOW64\%ILanguages%\ftsrch.dll.mui" /setowner "NT Service\TrustedInstaller"
    echo.
  
    rem ftsrch.dll.mui �̃A�N�Z�X�����玩�����폜
    icacls "%SystemRoot%\SysWOW64\%ILanguages%\ftsrch.dll.mui" /remove "%UserName%"
    echo.
  )
  
  IF NOT EXIST %SystemRoot%\SysWOW64\ftlx041e.dll (
    echo ��ftlx041e.dll �t�@�C���R�s�[������
  
    rem ftlx041e.dll ��OS���փR�s�[
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftlx041e.dll" "%SystemRoot%\SysWOW64"
    echo.
  
    rem ftlx041e.dll �̏��L���� TrustedInstaller �ɐݒ�
    icacls "%SystemRoot%\SysWOW64\ftlx041e.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    rem ftlx041e.dll �̃A�N�Z�X�����玩�����폜
    icacls "%SystemRoot%\SysWOW64\ftlx041e.dll" /remove "%UserName%"
    echo.
  
  )
  
  IF NOT EXIST %SystemRoot%\SysWOW64\ftlx0411.dll (
    echo ��ftlx0411.dll �t�@�C���R�s�[������
  
    rem ftlx0411.dll ��OS���փR�s�[
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftlx0411.dll" "%SystemRoot%\SysWOW64"
    echo.
  
    rem ftlx0411.dll �̏��L���� TrustedInstaller �ɐݒ�
    icacls "%SystemRoot%\SysWOW64\ftlx0411.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    rem ftlx0411.dll �̃A�N�Z�X�����玩�����폜
    icacls "%SystemRoot%\SysWOW64\ftlx0411.dll" /remove "%UserName%"
    echo.
  )
  
  IF NOT EXIST %SystemRoot%\SysWOW64\ftsrch.dll (
    echo ��ftsrch.dll �t�@�C���R�s�[������
  
    rem ftsrch.dll ��OS���փR�s�[
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftsrch.dll" "%SystemRoot%\SysWOW64"
    echo.
  
    rem ftsrch.dll �̏��L���� TrustedInstaller �ɐݒ�
    icacls "%SystemRoot%\SysWOW64\ftsrch.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    rem ftsrch.dll �̃A�N�Z�X�����玩�����폜
    icacls "%SystemRoot%\SysWOW64\ftsrch.dll" /remove "%UserName%"
    echo.
  )


  IF NOT EXIST %SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui (
    echo ��ftsrch.dll.mui �t�@�C���R�s�[������

    rem ftsrch.dll.mui ��OS���փR�s�[
    xcopy /r /y /h /q ".\cabtemp\%MuiFolder_x64%\ftsrch.dll.mui" "%SystemRoot%\System32\%ILanguages%"
    echo.
  
    rem ftsrch.dll.mui �̏��L���� TrustedInstaller �ɐݒ�
    icacls "%SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui" /setowner "NT Service\TrustedInstaller"
    echo.

    rem ftsrch.dll.mui �̃A�N�Z�X�����玩�����폜
    icacls "%SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui" /remove "%UserName%"
    echo.
  )

  IF NOT EXIST %SystemRoot%\System32\ftlx041e.dll (
    echo ��ftlx041e.dll �t�@�C���R�s�[������

    rem ftlx041e.dll ��OS���փR�s�[
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x64%\ftlx041e.dll" "%SystemRoot%\System32"
    echo.

    rem ftlx041e.dll �̏��L���� TrustedInstaller �ɐݒ�
    icacls "%SystemRoot%\System32\ftlx041e.dll" /setowner "NT Service\TrustedInstaller"
    echo.

    rem ftlx041e.dll �̃A�N�Z�X�����玩�����폜
    icacls "%SystemRoot%\System32\ftlx041e.dll" /remove "%UserName%"
    echo.
  )

  IF NOT EXIST %SystemRoot%\System32\ftlx0411.dll (
    echo ��ftlx0411.dll �t�@�C���R�s�[������

    rem ftlx0411.dll ��OS���փR�s�[
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x64%\ftlx0411.dll" "%SystemRoot%\System32"
    echo.

    rem ftlx0411.dll �̏��L���� TrustedInstaller �ɐݒ�
    icacls "%SystemRoot%\System32\ftlx0411.dll" /setowner "NT Service\TrustedInstaller"
    echo.

    rem ftlx0411.dll �̃A�N�Z�X�����玩�����폜
    icacls "%SystemRoot%\System32\ftlx0411.dll" /remove "%UserName%"
    echo.
  )

  IF NOT EXIST %SystemRoot%\System32\ftsrch.dll (
    echo ��ftsrch.dll �t�@�C���R�s�[������

    rem ftsrch.dll ��OS���փR�s�[
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x64%\ftsrch.dll" "%SystemRoot%\System32"
    echo.

    rem ftsrch.dll �̏��L���� TrustedInstaller �ɐݒ�
    icacls "%SystemRoot%\System32\ftsrch.dll" /setowner "NT Service\TrustedInstaller"
    echo.

    rem ftsrch.dll �̃A�N�Z�X�����玩�����폜
    icacls "%SystemRoot%\System32\ftsrch.dll" /remove "%UserName%"
    echo.
  )
) else (
  IF NOT EXIST %SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui (
    echo ��ftsrch.dll.mui �t�@�C���R�s�[������
  
    rem ftsrch.dll.mui ��OS���փR�s�[
    xcopy /r /y /h /q ".\cabtemp\%MuiFolder_x86%\ftsrch.dll.mui" "%SystemRoot%\System32\%ILanguages%"
    echo.
  
    rem ftsrch.dll.mui �̏��L���� TrustedInstaller �ɐݒ�
    icacls "%SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui" /setowner "NT Service\TrustedInstaller"
    echo.
  
    rem ftsrch.dll.mui �̃A�N�Z�X�����玩�����폜
    icacls "%SystemRoot%\System32\%ILanguages%\ftsrch.dll.mui" /remove "%UserName%"
    echo.
  )
  
  IF NOT EXIST %SystemRoot%\System32\ftlx041e.dll (
    echo ��ftlx041e.dll �t�@�C���R�s�[������
  
    rem ftlx041e.dll ��OS���փR�s�[
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftlx041e.dll" "%SystemRoot%\System32"
    echo.
  
    rem ftlx041e.dll �̏��L���� TrustedInstaller �ɐݒ�
    icacls "%SystemRoot%\System32\ftlx041e.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    rem ftlx041e.dll �̃A�N�Z�X�����玩�����폜
    icacls "%SystemRoot%\System32\ftlx041e.dll" /remove "%UserName%"
    echo.
  
  )
  
  IF NOT EXIST %SystemRoot%\System32\ftlx0411.dll (
    echo ��ftlx0411.dll �t�@�C���R�s�[������
  
    rem ftlx0411.dll ��OS���փR�s�[
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftlx0411.dll" "%SystemRoot%\System32"
    echo.
  
    rem ftlx0411.dll �̏��L���� TrustedInstaller �ɐݒ�
    icacls "%SystemRoot%\System32\ftlx0411.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    rem ftlx0411.dll �̃A�N�Z�X�����玩�����폜
    icacls "%SystemRoot%\System32\ftlx0411.dll" /remove "%UserName%"
    echo.
  )
  
  IF NOT EXIST %SystemRoot%\System32\ftsrch.dll (
    echo ��ftsrch.dll �t�@�C���R�s�[������
  
    rem ftsrch.dll ��OS���փR�s�[
    xcopy /r /y /h /q ".\cabtemp\%ExeDllFolder_x86%\ftsrch.dll" "%SystemRoot%\System32"
    echo.
  
    rem ftsrch.dll �̏��L���� TrustedInstaller �ɐݒ�
    icacls "%SystemRoot%\System32\ftsrch.dll" /setowner "NT Service\TrustedInstaller"
    echo.
  
    rem ftsrch.dll �̃A�N�Z�X�����玩�����폜
    icacls "%SystemRoot%\System32\ftsrch.dll" /remove "%UserName%"
    echo.
  )
)

rem ---------------------------------------
rem �e���|�����t�H���_�폜
rem ---------------------------------------
cd /d %~dp0
rd /s /q .\cabtemp >nul
rd /s /q .\msutemp >nul

rem ---------------------------------------
rem ���W�X�g�����C��
rem ---------------------------------------
echo �����W�X�g���C����
reg delete "HKLM\SOFTWARE\Microsoft\WinHelp" /f 2>nul
echo.
if %PROCESSOR_ARCHITECTURE%==AMD64 (
  reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\WinHelp" /f 2>nul
  echo.
)
rem ---------------------------------------
rem ���W�X�g���ǉ�
rem AllowProgrammaticMacros�c�}�N����L���ɂ���
rem AllowIntranetAccess�c�C���g���l�b�g��Ɋi�[����Ă���.hlp�t�@�C���̃u���b�N����
rem �Z�L�����e�B�[�ቺ�������\��������̂ŕW���Œǉ�����̂͂�߂鎖�ɂ���
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

echo ���C���X�g�[��������Ɋ������܂���
pause
goto :eof

:Error
echo ���G���[���������܂����A�C���X�g�[���𒆎~���܂�
pause
goto :eof
