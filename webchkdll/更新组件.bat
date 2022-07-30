@ECHO OFF
CLS



ECHO	正在停止IIS服务....
net stop iisadmin /y
ECHO.
ECHO.
ECHO	正在注销组件....
regsvr32 /u /s %windir%\system32\webchkdll.dll
ECHO.
ECHO.
ECHO	正在复制文件....
copy webchkdll.dll %windir%\system32
ECHO.
ECHO.
ECHO	正在注册组件....
ECHO	注册组件文件:webchkdll.dll
regsvr32 /s %windir%\system32\webchkdll.dll

ECHO.
ECHO.
ECHO	正在启动IIS服务....
net start w3svc
ECHO.
ECHO.
ECHO	恭喜你！组件更新成功！
ECHO.

pause
exit