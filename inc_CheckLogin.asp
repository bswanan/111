<!--#include file="const.asp"-->
<%
Dim Rs
Dim A_UserID	'����ԱID
Dim A_UserName	'����Ա�˻�
Dim A_TrueName	'����Ա����


If InStr(Request.ServerVariables("Query_String"),"'")<>0 then
	Netlog A_UserName,"�Ƿ���ַ���ʡ�"
End If

Call CheckAdminLogin()

Sub CheckAdminLogin()
	Dim AdminName,AdminPwd
	AdminName=GetCookies("Admin","AdminName")
	AdminPwd=GetCookies("Admin","AdminPassword")
	IF AdminName="" Then
			alert "��½��ʱ������û�е�½�������µ�½","./main_login.asp"
	End If
	If not CheckName(AdminName) or not CheckPassword(AdminPwd) then
		Response.redirect"main_login.asp"
		Response.end
	End If
	Set rs=conn.execute("Select ID,username,TrueName From [enablesoft_manager] where username='"& AdminName &"' and [password]='"& AdminPwd &"'")
	If rs.eof then
		Response.redirect"main_login.asp"
		Response.End
	Else
		A_UserID=rs("ID")
		A_UserName=rs("username")
		A_TrueName=rs("TrueName")
		End If
	rs.close : Set rs=Nothing
End Sub

Sub CheckString(Flag)

End Sub

%>