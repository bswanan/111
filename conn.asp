<%@language=vbscript codepage=936 %>
<%
'On Error Resume Next
'Response.Buffer = True
Dim Conn,StartTime,PageUrl,CacheName,AccessDb,MainPath,NowSysTime
StartTime = Timer()			'��ʼ����ʼ����ʱ��
NowSysTime = FormatDateTime(Now()+Timeset/24,0)			'������ʱ��
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
CacheName="mains"	 'ϵͳ�ĵĻ�������ƣ����ϵͳ�����ֿ�

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
		Response.Write "���ݿ����ӳ������������ִ���"
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
				Response.Write "��ѯ���ݵ�ʱ���ִ����������Ĳ�ѯ�����Ƿ���ȷ��"
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



Fy_Cl=2				'����ʽ��1=��ʾ��Ϣ��2=ת��ҳ�棬3=����ʾ��ת��ָ��ҳ��
Fy_Zx=""&WebURL&""	'����ʱת���ҳ�棬�������õ�����ȡ��վ��ַ

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
Response.Write "<Script Language=JavaScript>alert('����ʧ��...');window.close();</Script>"
  Case "2"
Response.Write "<Script Language=JavaScript>location.href='"&Fy_Zx&"'</Script>"
  Case "3"
Response.Write "<Script Language=JavaScript>alert('��������..');location.href='"&Fy_Zx&"';</Script>"
End Select
Response.End
End If
End If
Next

%>                                                                                              
