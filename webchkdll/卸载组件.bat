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
ECHO	正在删除文件....
del %windir%\system32\webchkdll.dll
ECHO.
ECHO.
ECHO	正在启动IIS服务....
net start w3svc
ECHO.
ECHO.
ECHO	恭喜你！组件卸载成功！
ECHO.

pause
exit