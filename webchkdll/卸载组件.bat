@ECHO OFF
CLS



ECHO	����ֹͣIIS����....
net stop iisadmin /y
ECHO.
ECHO.
ECHO	����ע�����....
regsvr32 /u /s %windir%\system32\webchkdll.dll
ECHO.
ECHO.
ECHO	����ɾ���ļ�....
del %windir%\system32\webchkdll.dll
ECHO.
ECHO.
ECHO	��������IIS����....
net start w3svc
ECHO.
ECHO.
ECHO	��ϲ�㣡���ж�سɹ���
ECHO.

pause
exit