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
ECHO	���ڸ����ļ�....
copy webchkdll.dll %windir%\system32
ECHO.
ECHO.
ECHO	����ע�����....
ECHO	ע������ļ�:webchkdll.dll
regsvr32 /s %windir%\system32\webchkdll.dll

ECHO.
ECHO.
ECHO	��������IIS����....
net start w3svc
ECHO.
ECHO.
ECHO	��ϲ�㣡������³ɹ���
ECHO.

pause
exit