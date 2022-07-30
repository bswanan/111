<!--#include file="inc_CheckLogin.asp"-->
<!--#include file="inc/md5.asp"-->
<html>
<head>
<title>管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="javascript" src="inc/js.js"></script>
<link rel="stylesheet" href="inc/style.css">
</head>
<body>
<%
action=request.QueryString("action")
select case action
	case "edit" :CheckString("04"): call edit
	case "del" : call del
	case else
		call main
end select

sub main
CheckString("02")
if GetForm("act")="addsave" Then
	CheckString("03")
	UserName=GetForm("UserName")
	PassWord=GetForm("PassWord")
	Truename=GetForm("Truename")
	Admin_State=GetForm("Admin_State")
	
	if strLength(UserName)<4 then ErrMsg = ErrMsg & "● 为了账户安全登陆账户不得小于4个字符；\n"
	if strLength(PassWord)<4 then ErrMsg = ErrMsg & "● 为了账户安全登陆密码不得小于4个字符；\n"
	if ErrMsg="" then
		if not CheckName(UserName) then ErrMsg = ErrMsg & "● 登陆账户包含有非法字符；\n"
		if not CheckPassword(PassWord) then ErrMsg = ErrMsg & "● 登陆密码包含有非法字符；\n"
	end if
	if ErrMsg="" then
		if not conn.execute("select id from [enablesoft_manager] where UserName='"& UserName &"'").eof then
			ErrMsg = ErrMsg & "此登陆账户已经被使用，请用其他用户名重试；"
			FoundErr=true
		end if
	end if
	
	if ErrMsg="" then
		PassWord=md5(PassWord,32)
		conn.execute("insert into[enablesoft_manager](UserName,[PassWord],LoginTimes,Truename,Joindate,Admin_State)values('"& UserName &"','"& PassWord &"',0,'"& Truename &"',"& SqlNowString &","& Admin_State &")")
		Netlog A_UserName,"添加新管理员"& UserName &""
		alert "管理员账户添加成功；","?"
	end if
	if ErrMsg<>"" then response.Write(SetErrMsg(ErrMsg))
end if
%>
<table width="99%" border="0" align="center" cellpadding="5" cellspacing="1" class="tablebk" style="border-collapse: collapse">
<form name="Form1" action="?" method="post">
<input name="act" value="addsave" type="hidden">
  <tr>
    <td colspan="2" class="table_titlebg">管理人员添加</td>
  </tr>
  <tr>
    <td width="41%" align="right" class="table_trbg02"><strong>登陆账户：</strong></td>
    <td width="59%" class="table_trbg02"><span class="table_trbg02">
      <input name="UserName" type="text" class="INPUT" id="UserName" size="30" value="<%=UserName%>">
    </span></td>
  </tr>
  <tr>
    <td align="right" class="table_trbg02"><strong>登陆密码：</strong></td>
    <td class="table_trbg02"><input name="PassWord" type="password" class="INPUT" id="PassWord" size="30"></td>
  </tr>
  <tr>
    <td align="right" class="table_trbg02"><strong>管理姓名：</strong></td>
    <td class="table_trbg02"><span class="table_trbg02">
      <input name="Truename" type="text" class="INPUT" id="Truename" size="30" value="<%=Truename%>">
    </span></td>
  </tr>
  <tr>
    <td align="right" class="table_trbg02"><strong>账户状态：</strong></td>
    <td class="table_trbg02"><input type="radio" name="Admin_State" value="0"<%if Admin_State="" then response.Write(" checked=""checked""") else response.Write(SetChecked(Admin_State,"0"))%>>正常 &nbsp; 
      <input type="radio" name="Admin_State" value="1"<%=SetChecked(Admin_State,"1")%>>锁定</td>
  </tr>
  <tr>
    <td height="40" colspan="2" align="center" class="table_trbg02"><input type="submit" name="Submit" value="提交"> 
      &nbsp; 
      <input type="reset" name="Submit" value="重置"></td>
  </tr>
  </form>
</table>

<br class="table_br" />
<table width="99%" border="0" align="center" cellpadding="5" cellspacing="1" class="tablebk" style="border-collapse: collapse">
  <tr>
    <td colspan="7" class="table_titlebg">管理人员列表</td>
  </tr>
  <tr>
    <td width="8%" align="center" class="table_trbg01"><strong>ID</strong></td>
    <td width="18%" align="center" class="table_trbg01"><strong>登陆账户</strong></td>
    <td width="17%" align="center" class="table_trbg01"><strong>管理姓名</strong></td>
    <td width="12%" align="center" class="table_trbg01"><strong>登陆次数</strong></td>
    <td width="12%" align="center" class="table_trbg01"><strong>账户状态</strong></td>
    <td width="20%" align="center" class="table_trbg01"><strong>加入时间</strong></td>
    <td align="center" class="table_trbg01"><strong>操作</strong></td>
  </tr>
<%set rs=conn.execute("select * from [enablesoft_manager] order by id desc")
i=0
dim classname
do while not rs.eof
if i mod 2 =0 then classname=" class=""table_trbg03""" else classname=" class=""table_trbg02"""
%>
  <tr>
    <td align="center" class="table_trbg02"<%=classname%>><%=rs("id")%></td>
    <td align="center" class="table_trbg02"<%=classname%>><%=rs("username")%></td>
    <td align="center" class="table_trbg02"<%=classname%>><%=rs("truename")%></td>
    <td align="center" class="table_trbg02"<%=classname%>><%=rs("LoginTimes")%></td>
    <td align="center" class="table_trbg02"<%=classname%>><%if rs("Admin_State")="0" then response.Write("正常") else response.Write("<span class=""red"">锁定</span>")%></td>
    <td align="center" class="table_trbg02"<%=classname%>><%=rs("joindate")%></td>
    <td align="center" class="table_trbg02"<%=classname%>><A href="?action=edit&id=<%=rs("id")%>">编辑</A> <A href="?action=del&id=<%=rs("id")%>">删除</A></td>
  </tr>
 <%rs.movenext
 loop
 rs.close:set rs=nothing%>
</table>
<%end sub

sub edit
	CheckString("04")
id=checkstr(request.QueryString("id"))
if not isInteger(id) then alert "参数传递出错，请重试；","back"
set rs=conn.execute("select * from [enablesoft_manager] where id="& id &"")
if rs.eof then
	alert "没有找到此管理人员，请重试；","back"
else
	UserName=rs("UserName")
	Truename=rs("Truename")
	Admin_State=rs("Admin_State")
end if
	rs.close:set rs=nothing


if GetForm("act")="editsave" then

	UserName=GetForm("UserName")
	PassWord=GetForm("PassWord")
	Truename=GetForm("Truename")
	Admin_State=GetForm("Admin_State")

	if strLength(UserName)<4 then ErrMsg = ErrMsg & "● 为了账户安全登陆账户不得小于4个字符；\n"
	if strLength(PassWord)<4 and PassWord<>"" then ErrMsg = ErrMsg & "● 为了账户安全登陆密码不得小于4个字符；\n"
	if ErrMsg="" then
		if not CheckName(UserName) then ErrMsg = ErrMsg & "● 登陆账户包含有非法字符；\n"
		if (not CheckPassword(PassWord)) and PassWord<>"" then ErrMsg = ErrMsg & "● 登陆密码包含有非法字符；\n"
	end if

	if not conn.execute("select id from [enablesoft_manager] where UserName='"& UserName &"' and id<>"& id &"").eof then
		ErrMsg = ErrMsg & "● 此登陆账户已经被使用，请用其他用户名重试；\n"
		FoundErr=true
	end if
	
	
	if ErrMsg="" then
		if PassWord<>"" then
			PassWord=md5(PassWord,32)
			sql2=" ,[PassWord]='"& PassWord &"' "
		end if
		conn.execute("update [enablesoft_manager] set UserName='"& UserName &"',Truename='"& Truename &"',Admin_State="& Admin_State &" "& sql2 &" where id="& id &"")
		
		if cstr(A_UserID)=cstr(id) then
		SetCookies "Admin","AdminName",UserName '建立session cookies
			if PassWord<>"" then SetCookies "Admin","AdminPassword",PassWord
		end if
		Netlog A_UserName,"编辑保存管理员"& UserName &""
		alert "编辑保存成功；","?action=edit&id="& id &""
	end if
	if ErrMsg<>"" then response.Write(SetErrMsg(ErrMsg))
end if
%>
<table width="99%" border="0" align="center" cellpadding="5" cellspacing="1" class="tablebk" style="border-collapse: collapse">
<form name="Form1" action="?action=edit&id=<%=id%>" method="post">
<input name="act" value="editsave" type="hidden">
  <tr>
    <td colspan="2" class="table_titlebg">管理人员编辑</td>
  </tr>
  <tr>
    <td width="18%" align="right" class="table_trbg02"><strong>登陆账户：</strong></td>
    <td width="82%" class="table_trbg02"><span class="table_trbg02">
      <input name="UserName" type="text" class="INPUT" id="UserName" size="30" value="<%=UserName%>">
    </span></td>
  </tr>
  <tr>
    <td align="right" class="table_trbg02"><strong>登陆密码：</strong></td>
    <td class="table_trbg02"><input name="PassWord" type="password" class="INPUT" id="PassWord" size="30"> 
    不修改请留空</td>
  </tr>
  <tr>
    <td align="right" class="table_trbg02"><strong>管理姓名：</strong></td>
    <td class="table_trbg02"><span class="table_trbg02">
      <input name="Truename" type="text" class="INPUT" id="Truename" size="30" value="<%=Truename%>">
    </span></td>
  </tr>
  <tr>
    <td align="right" class="table_trbg02"><strong>账户状态：</strong></td>
    <td class="table_trbg02"><input type="radio" name="Admin_State" value="0"<%=SetChecked(Admin_State,"0")%>>正常 &nbsp; 
      <input type="radio" name="Admin_State" value="1"<%=SetChecked(Admin_State,"1")%>>锁定</td>
  </tr>
  <tr>
    <td height="40" colspan="2" align="center" class="table_trbg02"><input type="submit" name="Submit" value="提交"> 
      &nbsp; 
      <input type="button" name="Submit" value="返回" onClick="window.location='?';"></td>
  </tr>
  </form>
</table>
<%end sub

sub del
	CheckString("05")
id=checkstr(request.QueryString("id"))
if not isInteger(id) then alert "参数传递出错，请重试；","back"
set rs=conn.execute("select * from [enablesoft_manager] where id="& id &"")
if rs.eof then
	alert "没有找到此管理人员，请重试；","back"
end If
	UserName=rs("UserName")
	rs.close:set rs=nothing
	conn.execute("delete from [enablesoft_manager] where id="& id &"")
	alert "管理人员 "& UserName &" 删除成功；","?"

End sub
%>
</body>
</html>