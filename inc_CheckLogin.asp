<!--#include file="const.asp"-->
<%
Dim Rs
Dim A_UserID	'管理员ID
Dim A_UserName	'管理员账户
Dim A_TrueName	'管理员姓名


If InStr(Request.ServerVariables("Query_String"),"'")<>0 then
	Netlog A_UserName,"非法地址访问。"
End If

Call CheckAdminLogin()

Sub CheckAdminLogin()
	Dim AdminName,AdminPwd
	AdminName=GetCookies("Admin","AdminName")
	AdminPwd=GetCookies("Admin","AdminPassword")
	IF AdminName="" Then
			alert "登陆超时或您还没有登陆，请重新登陆","./main_login.asp"
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