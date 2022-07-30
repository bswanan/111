@ECHO OFF
CLS


ECHO	正在复制文件....
copy webchkdll.dll %windir%\system32
ECHO.

ECHO.
ECHO	正在注册组件....
ECHO	注册组件文件:webchkdll.dll
regsvr32 /s %windir%\system32\webchkdll.dll
ECHO.

ECHO.
ECHO	恭喜你！组件注册成功！
ECHO.

pause
exit

