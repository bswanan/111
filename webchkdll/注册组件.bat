@ECHO OFF
CLS


ECHO	���ڸ����ļ�....
copy webchkdll.dll %windir%\system32
ECHO.

ECHO.
ECHO	����ע�����....
ECHO	ע������ļ�:webchkdll.dll
regsvr32 /s %windir%\system32\webchkdll.dll
ECHO.

ECHO.
ECHO	��ϲ�㣡���ע��ɹ���
ECHO.

pause
exit

