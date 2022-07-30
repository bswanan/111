<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->
<%
Call OpenConn '建立数据库连接

Dim SysInfo
If not IsArray(Application(CacheName&"_sysconfig")) then
Set rs=ConnExecute("select * from [enablesoft_config] ")
If Not rs.eof Then
	Dim AppStr
	AppStr=rs("sys_name")&"<$$>"
	Application.Lock()	
	Application(CacheName&"_sysconfig")=Split(AppStr,"<$$>")
	Application.UnLock()
End If
	rs.close:set rs=nothing
End If
SysInfo=Application(CacheName&"_sysconfig")
%>