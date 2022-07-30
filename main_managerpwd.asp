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
<%CheckString("06")
if GetForm("act")="editsave" then

	PassWord=GetForm("PassWord")
	PassWord2=GetForm("PassWord2")
	PassWord3=GetForm("PassWord3")

	if PassWord="" then
		ErrMsg = ErrMsg & "● 请输入旧登陆密码；\n"
	end If
	if ErrMsg="" then
		if strLength(PassWord2)<4 or strLength(PassWord2)>20  then
			ErrMsg = ErrMsg & "● 为了账户安全登陆密码应该在4-20个字符；\n"
		end if
		if not CheckPassword(PassWord2) then
			ErrMsg = ErrMsg & "● 新登陆密码包含有非法字符；\n"
		end If
	End If
	
	if ErrMsg="" then
		if PassWord2<>PassWord3  then
			ErrMsg = ErrMsg & "● 两次输入的密码不一致；\n"
		end If
	end if

	if ErrMsg="" then
		if conn.execute("select id from [enablesoft_manager] where [PassWord]='"& md5(PassWord,32) &"' and id="& A_UserID &"").eof then
			ErrMsg = ErrMsg & "● 旧密码输入错误，请重新输入；\n"
		end if
	end if
	
	
	if ErrMsg="" then
			PassWord2=md5(PassWord2,32)
		conn.execute("update [enablesoft_manager] set [PassWord]='"& PassWord2 &"' where id="& A_UserID &"")
		SetCookies "Admin","AdminPassword",PassWord2
		Netlog A_UserName,"编辑保存登陆密码"
		alert "编辑保存成功；","?"
	end if
	if ErrMsg<>"" then response.Write(SetErrMsg(ErrMsg))
end if
%>
<table width="99%" border="0" align="center" cellpadding="5" cellspacing="1" class="tablebk" style="border-collapse: collapse">
<form name="Form1" action="?" method="post">
<input name="act" value="editsave" type="hidden">
   <tr>
    <td colspan="2" class="table_titlebg">修改登陆密码</td>
  </tr>
  <tr>
    <td width="40%" align="right" class="table_trbg02"><strong>当前用户名：</strong></td>
    <td width="60%" class="table_trbg02"><strong><%=A_UserName%></strong></td>
  </tr>
  <tr>
    <td align="right" class="table_trbg02"><strong>输入旧密码：</strong></td>
    <td class="table_trbg02"><input name="PassWord" type="password" class="INPUT" id="PassWord" size="30"> 
    </td>
  </tr>
  <tr>
    <td align="right" class="table_trbg02"><strong>输入新密码：</strong></td>
    <td class="table_trbg02"><input name="PassWord2" type="password" class="INPUT" id="PassWord2" size="30"> 
    </td>
  </tr>
  <tr>
    <td align="right" class="table_trbg02"><strong>确认新密码：</strong></td>
    <td class="table_trbg02"><input name="PassWord3" type="password" class="INPUT" id="PassWord3" size="30"></td>
  </tr>
  <tr>
    <td height="40" colspan="2" align="center" class="table_trbg02"><input type="submit" name="Submit" value="提交"> 
      &nbsp; 
      <input type="button" name="Submit" value="返回" onClick="window.location='?';"></td>
  </tr>
  </form>
</table>
</body>
</html>