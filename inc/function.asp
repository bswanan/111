<%

'***************************************************
'函数名：IsObjInstalled
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：True  ----已经安装  False ----没有安装
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
'函数名：GetIP
'作  用：获取用户IP
'返回值：IP地址
'***************************************************
	Function GetIP()
		Dim Temp
		Temp = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		If Temp = "" or isnull(Temp) or isEmpty(Temp) Then Temp = Request.ServerVariables("REMOTE_ADDR")
		If Instr(Temp,"'")>0 Then Temp="0.0.0.0"
		GetIP = Temp
	End Function

'***************************************************
'函数名：netlog
'作  用：网站日志
'***************************************************
	Sub Netlog(addname,str) '记录日志
		Dim Temp
		If addname="" Then addname="-"
		Temp=Left(Request.ServerVariables("script_name")&"<br>"&Replace(Request.ServerVariables("Query_String"),"'","''"),255)
		conn.Execute("insert into [enablesoft_log] (UserName,UserIP,Remark,LogTime,Geturl) values ('"& addname &"','"& GetIP &"','"&str&"','"& now &"','"& Temp &"')")
	End Sub

'***************************************************
'函数名：CheckMake
'作  用：禁止外部提交数据
'返回值：True  ----外部  False ----正常
'***************************************************
	function CheckMake()
		Dim Come,Here
		Come=Cstr(Request.ServerVariables("HTTP_REFERER"))
		Here=Cstr(Request.ServerVariables("SERVER_NAME"))
		If Come<>"" And Mid(Come,8,Len(Here)) <> Here Then CheckMake=False Else CheckMake=True 
	End function

'***************************************************
'函数名：Alert
'作  用：javascript提示、转向、关闭
'参  数：msg ----提示字符  goUrl ----"back"后退,"close"关闭,其他为转向地址
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
'函数名：HTMLEncode
'作  用：转换成Html标签
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
'函数名：GetJsStr
'作  用：转换JS字符
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
'函数名：GetForm
'作  用：获取Post数据并过滤
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
'函数名：CheckStr
'作  用：过滤字符
'***************************************************
Function CheckStr(Str)
	If Trim(Str)="" Or IsNull(str) Then 
	CheckStr=""
	Exit Function
	End If
	Checkstr=Replace(Str,"'","''")
End Function

'***************************************************
'函数名：GetEncode
'作  用：转换htmlencode编码
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
'函数名：IsValidEmain
'作  用：检查Emain地址合法性
'参  数：emain ----要检查的Emain地址
'返回值：True  ----Emain地址合法
'       False ----Emain地址不合法
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
'函数名：gotTopic
'作  用：截取字符串部分。汉字算两个字符，英文数字算一个字符。
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

	'清理html
'********************************************
'函数名：Replacehtml
'作  用：清理html代码
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
'函数名：CheckName
'作  用：名字字符检验 - 中文 -_+-/*
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
'函数名：CheckPassword
'作  用：密码检验
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
'函数名：isInteger
'作  用：整数检验
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
'函数名：IsValidNumber
'作  用：数字型，带小数检验
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
'函数名：strLength
'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
'参  数：findstr  ----要求长度的字符串
'返回值：字符串长度
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
'函数名：SetSelected
'作  用：下拉框默认值
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
'函数名：SetChecked
'作  用：单选默认值
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
'函数名：SetChecked2
'作  用：多选默认值
'**************************************************
Public Function SetChecked2(val1,val2)
	If IsNull(val1) Or IsNull(val2) Then Exit Function
	if instr(cstr(val1),cstr(val2))>0 then
		SetChecked2=" checked=""checked"""
	else
		SetChecked2=""
	end if
end Function


	'增加session和cookies     ,session=系统缓存名+session名,Cookies=(系统缓存名+名称)+名
'**************************************************
'函数名：SetCookies
'作  用：增加session和cookies
'参  数：
'**************************************************
	Sub SetCookies(root,name,value)
		Session(CacheName & name)=value
		Response.Cookies(CacheName & root)(name)=value
	End Sub
	'获取某一名称的值
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
	tempmsg=Split(MsgStr,"●")
	If (UBound(tempmsg)-1)>2 Then
		SetErrMsg="<script>alert("""&"●"& tempmsg(1)&"●"&tempmsg(2)&"●"&tempmsg(3) &""");</script>"
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
	tempmsg=Split(MsgStr,"●")
	If (UBound(tempmsg)-1)>2 Then
		SetErrMsg="<script>alert("""&"●"& tempmsg(1)&"●"&tempmsg(2)&"●"&tempmsg(3) &""");</script>"
	Else
		SetErrMsg="<script>alert("""& MsgStr &""");window.history.go(-1);</script>"
	End if
End Function
%>