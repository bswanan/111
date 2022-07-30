<%

'***************************************************
'��������IsObjInstalled
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True  ----�Ѿ���װ  False ----û�а�װ
'***************************************************
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function

'***************************************************
'��������GetIP
'��  �ã���ȡ�û�IP
'����ֵ��IP��ַ
'***************************************************
	Function GetIP()
		Dim Temp
		Temp = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		If Temp = "" or isnull(Temp) or isEmpty(Temp) Then Temp = Request.ServerVariables("REMOTE_ADDR")
		If Instr(Temp,"'")>0 Then Temp="0.0.0.0"
		GetIP = Temp
	End Function

'***************************************************
'��������netlog
'��  �ã���վ��־
'***************************************************
	Sub Netlog(addname,str) '��¼��־
		Dim Temp
		If addname="" Then addname="-"
		Temp=Left(Request.ServerVariables("script_name")&"<br>"&Replace(Request.ServerVariables("Query_String"),"'","''"),255)
		conn.Execute("insert into [enablesoft_log] (UserName,UserIP,Remark,LogTime,Geturl) values ('"& addname &"','"& GetIP &"','"&str&"','"& now &"','"& Temp &"')")
	End Sub

'***************************************************
'��������CheckMake
'��  �ã���ֹ�ⲿ�ύ����
'����ֵ��True  ----�ⲿ  False ----����
'***************************************************
	function CheckMake()
		Dim Come,Here
		Come=Cstr(Request.ServerVariables("HTTP_REFERER"))
		Here=Cstr(Request.ServerVariables("SERVER_NAME"))
		If Come<>"" And Mid(Come,8,Len(Here)) <> Here Then CheckMake=False Else CheckMake=True 
	End function

'***************************************************
'��������Alert
'��  �ã�javascript��ʾ��ת�򡢹ر�
'��  ����msg ----��ʾ�ַ�  goUrl ----"back"����,"close"�ر�,����Ϊת���ַ
'***************************************************
	Sub Alert(msg,goUrl) '
		msg = replace(msg,"'","\'")
	  	If goUrl="back" Then
		Response.Write ("<script LANGUAGE='javascript'>alert('" & msg & "');window.history.go(-1);</script>")
		Response.End
		ElseIf goUrl="this" Then
		Response.Write ("<script LANGUAGE='javascript'>alert('" & msg & "');</script>")
		ElseIf goUrl="close" Then
		Response.Write ("<script LANGUAGE='javascript'>alert('" & msg & "');window.close();</script>")
		Response.End
		Else
		Response.Write ("<script LANGUAGE='javascript'>alert('" & msg & "');window.location.href='"&goUrl&"'</script>")
		Response.End
		End IF
		
	End Sub


'***************************************************
'��������HTMLEncode
'��  �ã�ת����Html��ǩ
'***************************************************
function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR /> ")
	fString = Replace(fString, "script", "&#115cript")
    HTMLEncode = fString
end if
end Function

'***************************************************
'��������GetJsStr
'��  �ã�ת��JS�ַ�
'***************************************************
	Function GetJsStr(Str)
		If IsNull(Str) Then
			Str = ""
		Else
			Str = Replace(Str,"'","\'")
			Str = replace(Str,chr(34),"\"&chr(34))
		End If
		GetJsStr = Str
	End Function


'***************************************************
'��������GetForm
'��  �ã���ȡPost���ݲ�����
'***************************************************
Function GetForm(Str)
		Str = Trim(Request.Form(Str))
		If IsEmpty(Str) Then
			GetForm = ""
			Exit Function 
		End If
		Str = Replace(Str,Chr(0),"")
		GetForm = Replace(Str,"'","''")
End Function

'***************************************************
'��������CheckStr
'��  �ã������ַ�
'***************************************************
Function CheckStr(Str)
	If Trim(Str)="" Or IsNull(str) Then 
	CheckStr=""
	Exit Function
	End If
	Checkstr=Replace(Str,"'","''")
End Function

'***************************************************
'��������GetEncode
'��  �ã�ת��htmlencode����
'***************************************************
Function GetEncode(Str)
		If IsEmpty(Str) Or IsNull(Str) Then
			GetEncode = ""
			Exit Function
		End If
		Str = Replace(Str,Chr(0),"")
		GetEncode = server.htmlencode(Str)
End Function

'********************************************
'��������IsValidEmain
'��  �ã����Emain��ַ�Ϸ���
'��  ����emain ----Ҫ����Emain��ַ
'����ֵ��True  ----Emain��ַ�Ϸ�
'       False ----Emain��ַ���Ϸ�
'********************************************
function IsValidEmain(emain)
	dim names, name, i, c
	IsValidEmain = true
	names = Split(emain, "@")
	if UBound(names) <> 1 then
	   IsValidEmain = false
	   exit function
	end if
	for each name in names
		if Len(name) <= 0 then
			IsValidEmain = false
    		exit function
		end if
		for i = 1 to Len(name)
		    c = Lcase(Mid(name, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
		       IsValidEmain = false
		       exit function
		     end if
	   next
	   if Left(name, 1) = "." or Right(name, 1) = "." then
    	  IsValidEmain = false
	      exit function
	   end if
	next
	if InStr(names(1), ".") <= 0 then
		IsValidEmain = false
	   exit function
	end if
	i = Len(names(1)) - InStrRev(names(1), ".")
	if i <> 2 and i <> 3 then
	   IsValidEmain = false
	   exit function
	end if
	if InStr(emain, "..") > 0 then
	   IsValidEmain = false
	end if
end function


'********************************************
'��������gotTopic
'��  �ã���ȡ�ַ������֡������������ַ���Ӣ��������һ���ַ���
'********************************************
	Public Function gotTopic(str,strlen)
	   Dim l,t,i,c
	   If str="" Or isnull(str) Then
	      gotTopic=""
		  Exit Function
	   End If
	   str=Replace(Replace(Replace(Replace(Replace(str,"&nbsp;"," "),"&quot;",Chr(34)),"&gt;",">"),"&lt;","<"),"&#124;","|")
	   l=Len(str)
	   t=0
	   For i=1 To l
		  c=Abs(Asc(Mid(str,i,1)))
		  If c>255 Then
		    t=t+2
		  Else
		    t=t+1
		  End If
		  If t>=strlen Then
		    gotTopic=Left(str,i) & ".."
		    Exit For
		  Else
		    gotTopic=str
		  End If
	   Next
	   gotTopic=Replace(Replace(Replace(Replace(replace(gotTopic," ","&nbsp;"),Chr(34),"&quot;"),">","&gt;"),"<","&lt;"),"|","&#124;")
	End Function

	'����html
'********************************************
'��������Replacehtml
'��  �ã�����html����
'********************************************
	Public Function Replacehtml(tstr)
		Dim Str,re
		Str=Tstr
		If isNUll(Str) then 
			  Replacehtml=""
			  exit function
		End if
		Set re=new RegExp
			re.IgnoreCase =True
			re.Global=True
			re.Pattern="<(p|\/p|br)>"
			Str=re.Replace(Str,vbNewLine)
			re.Pattern="<img.[^>]*src(=| )(.[^>]*)>"
			str=re.replace(str,"[img]$2[/img]")
			re.Pattern="<(.[^>]*)>"
			Str=re.Replace(Str,"")
			Set Re=Nothing
			Replacehtml=Str
	End Function	


'********************************************
'��������CheckName
'��  �ã������ַ����� - ���� -_+-/*
'********************************************
	Public Function CheckName(Str)
		Checkname=True
		Dim Rep,pass
		Set Rep=New RegExp
		Rep.Global=True
		Rep.IgnoreCase=True
		Rep.Pattern="[\u0009\u0020\u0022-\u0029\u002C\u002E\u003A-\u003F\u005B\u005C\u0060\u007C\u007E\u00FF\uE5E5]"
		Set pass=Rep.Execute(Str)
		If pass.count<>0 Then CheckName=False
		Set Rep=Nothing
	End Function
	
'********************************************
'��������CheckPassword
'��  �ã��������
'********************************************
	Public Function CheckPassword(Str)
		Dim pass
		CheckPassword=false
		If Str <> "" Then
		Dim Rep
		Set Rep = New RegExp
		Rep.Global = True
		Rep.IgnoreCase = True
		Rep.Pattern="[\u4E00-\u9FA5\uF900-\uFA2D\u0027\u007C]"
		Pass=rep.Test(Str)
		Set Rep=nothing
		If Not Pass Then CheckPassword=True
		End If
	End Function


'********************************************
'��������isInteger
'��  �ã���������
'********************************************
	Public function isInteger(para)
		   on error resume Next
		   Dim str
		   Dim l,i
		   If isNUll(para) then 
			  isInteger=false
			  exit function
		   End if
		   str=cstr(para)
		   If trim(str)="" then
			  isInteger=false
			  exit function
		   End if
		   l=len(str)
		   For i=1 to l
			   If mid(str,i,1)>"9" or mid(str,i,1)<"0" then
				  isInteger=false 
				  exit function
			   End if
		   Next
		   isInteger=true
		   If err.number<>0 then err.clear
	End function


'********************************************
'��������IsValidNumber
'��  �ã������ͣ���С������
'********************************************
	Public function IsValidNumber(Num)
		If Num="" Or IsNull(Num) Then
			IsValidNumber=false
			Exit Function
		End if
		Dim Rep
		Set Rep = new RegExp
		rep.pattern = "^[-0-9]*[\.]?[0-9]+$"
		IsValidNumber=rep.Test(Num)
		Set Rep=Nothing
		If len(Num)>30 Then IsValidNumber=false
	End function

'**************************************************
'��������strLength
'��  �ã����ַ������ȡ������������ַ���Ӣ����һ���ַ���
'��  ����findstr  ----Ҫ�󳤶ȵ��ַ���
'����ֵ���ַ�������
'**************************************************
	Public function strLength(findstr)
		Dim Rep,lens,i
		If findstr="" Or IsNull(findstr) Then
			strLength=0
			Exit Function
		End if
		Set rep=new regexp
		rep.Global=true
		rep.IgnoreCase=true
		rep.Pattern="[\u4E00-\u9FA5\uF900-\uFA2D]"
		For each i in rep.Execute(findstr)
			lens=lens+1
		Next
		Set Rep=Nothing
		lens=lens + len(findstr)
		strLength=lens
	End Function




'**************************************************
'��������SetSelected
'��  �ã�������Ĭ��ֵ
'**************************************************
Public Function SetSelected(val1,val2)
	If IsNull(val1) Or IsNull(val2) Then Exit Function
	if cstr(val1)=cstr(val2) then
		SetSelected=" selected=""selected"""
	else
		SetSelected=""
	end if
end Function

'**************************************************
'��������SetChecked
'��  �ã���ѡĬ��ֵ
'**************************************************
Public Function SetChecked(val1,val2)
	If IsNull(val1) Or IsNull(val2) Then Exit Function
	if cstr(val1)=cstr(val2) then
		SetChecked=" checked=""checked"""
	else
		SetChecked=""
	end if
end Function

'**************************************************
'��������SetChecked2
'��  �ã���ѡĬ��ֵ
'**************************************************
Public Function SetChecked2(val1,val2)
	If IsNull(val1) Or IsNull(val2) Then Exit Function
	if instr(cstr(val1),cstr(val2))>0 then
		SetChecked2=" checked=""checked"""
	else
		SetChecked2=""
	end if
end Function


	'����session��cookies     ,session=ϵͳ������+session��,Cookies=(ϵͳ������+����)+��
'**************************************************
'��������SetCookies
'��  �ã�����session��cookies
'��  ����
'**************************************************
	Sub SetCookies(root,name,value)
		Session(CacheName & name)=value
		Response.Cookies(CacheName & root)(name)=value
	End Sub
	'��ȡĳһ���Ƶ�ֵ
	Function GetCookies(root,name)
		GetCookies=Session(CacheName & name)
		If GetCookies="" Then GetCookies=Request.Cookies(CacheName & root)(name)
	End Function




Function SetErrMsg(MsgStr)
	Dim tempmsg
	If Trim(MsgStr)="" Then 
		SetErrMsg=""
		Exit Function
	End if
	tempmsg=Split(MsgStr,"��")
	If (UBound(tempmsg)-1)>2 Then
		SetErrMsg="<script>alert("""&"��"& tempmsg(1)&"��"&tempmsg(2)&"��"&tempmsg(3) &""");</script>"
	Else
		SetErrMsg="<script>alert("""& MsgStr &""");</script>"
	End if
End Function

Function SetErrMsg_Back(MsgStr)
	Dim tempmsg
	If Trim(MsgStr)="" Then 
		SetErrMsg=""
		Exit Function
	End if
	tempmsg=Split(MsgStr,"��")
	If (UBound(tempmsg)-1)>2 Then
		SetErrMsg="<script>alert("""&"��"& tempmsg(1)&"��"&tempmsg(2)&"��"&tempmsg(3) &""");</script>"
	Else
		SetErrMsg="<script>alert("""& MsgStr &""");window.history.go(-1);</script>"
	End if
End Function
%>