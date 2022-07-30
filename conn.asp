<%@language=vbscript codepage=936 %>
<%
'On Error Resume Next
'Response.Buffer = True
Dim Conn,StartTime,PageUrl,CacheName,AccessDb,MainPath,NowSysTime
StartTime = Timer()			'初始化开始运行时间
NowSysTime = FormatDateTime(Now()+Timeset/24,0)			'服务器时间
Const IsSqlDataBase = 0

If IsSqlDataBase = 1 Then

	Const SqlDatabaseName = "enablesoft#@#emaindata"
	Const SqlUsername = "sa"
	Const SqlPassword = "123"
	Const SqlLocalName = "(local)"

	SqlNowString = "GetDate()"
Else
	AccessDb = "yxyzdata#adi@.asp"
	SqlNowString = "Now()"
End If


PageURL=Lcase(Request.ServerVariables("URL"))
CacheName="mains"	 '系统的的缓存等名称，多个系统请区分开

Sub OpenConn
	Dim ConnStr
	If IsSqlDataBase = 1 Then
		ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
	Else
		ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(MainPath&AccessDb)
	End If
	on error resume next
	Set conn=Server.CreateObject("ADODB.Connection")
	Conn.Open ConnStr
	If Err Then
		Err.Clear
		Set Conn = Nothing
		Response.Write "数据库连接出错，请检查连接字串。"
		Response.End
	End If
End Sub

Function ConnExecute(Command)
		If Not IsObject(Conn) Then OpenConn		
		If IsDeBug = 0 Then 
			On Error Resume Next
			Set ConnExecute = Conn.Execute(Command)
			If Err Then
				err.Clear
				Set Conn = Nothing
				Response.Write "查询数据的时候发现错误，请检查您的查询代码是否正确。"
				Response.End
			End If
		Else
			If IsShowSQL=1 Then
				Response.Write command & "<br>"
			End If
			Set ConnExecute = Conn.Execute(Command)
		End If	
		SqlQueryNum = SqlQueryNum+1
End Function

Sub CloseConn
	If IsObject(Conn) Then
		Conn.close
		Set Conn=Nothing
	End if
End Sub



Fy_Cl=2				'处理方式：1=提示信息，2=转向页面，3=先提示再转向指定页面
Fy_Zx=""&WebURL&""	'出错时转向的页面，现在设置的是提取网站网址

On Error Resume Next
Fy_Url=Request.ServerVariables("QUERY_STRING")
Fy_a=split(Fy_Url,"&")
redim Fy_Cs(ubound(Fy_a))
On Error Resume Next
for Fy_x=0 to ubound(Fy_a)
Fy_Cs(Fy_x) = left(Fy_a(Fy_x),instr(Fy_a(Fy_x),"=")-1)
Next
For Fy_x=0 to ubound(Fy_Cs)
If Fy_Cs(Fy_x)<>"" Then
If Instr(LCase(Request(Fy_Cs(Fy_x))),"'")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"and")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"=")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"(")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),")")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),">")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"select")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"update")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"chr")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"delete%20from")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),";")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"insert")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"mid")<>0 Or Instr(LCase(Request(Fy_Cs(Fy_x))),"master.")<>0 Then
Select Case Fy_Cl
  Case "1"
Response.Write "<Script Language=JavaScript>alert('操作失败...');window.close();</Script>"
  Case "2"
Response.Write "<Script Language=JavaScript>location.href='"&Fy_Zx&"'</Script>"
  Case "3"
Response.Write "<Script Language=JavaScript>alert('参数错误..');location.href='"&Fy_Zx&"';</Script>"
End Select
Response.End
End If
End If
Next

%>                                                                                              
