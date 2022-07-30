<!--#include file="const.asp" -->
<!--#include file="inc/md5.asp" -->
<%
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 

action=trim(request("action"))
if action="logout" then
	Session(CacheName &"AdminName") = Empty
	Response.Cookies(CacheName &"Admin")("AdminName")= Empty
	Session(CacheName &"AdminName") = Empty
	Response.Cookies(CacheName &"Admin")("AdminPassword")= Empty
	Response.Redirect "main_login.asp"
end if

if action="loginchk" then '登陆验证
OpenConn
if CheckMake=False then Alert "你提交的路径有误，禁止从站点外部提交数据请不要乱加参数！","./"
adminname=GetForm("adminname")
adminpwd=GetForm("adminpwd")
checkcode=GetForm("checkcode")

if adminname="" then Alert "请输入登陆用户名！","back"
if adminpwd="" then Alert "请输入登陆密码！","back"
if checkcode="" then Alert "请输入安全码！","back"

if session("CheckCode")="" then Alert "验证码已经过期，请重新登录。","?"
if Lcase(checkcode)<>LCase(CStr(session("CheckCode"))) then Alert "您输入的附加码和系统产生的不一致，请重新输入。","?"
session("CheckCode")=""
if GetCookies("Admin","Login_ErrNum")="" then call SetCookies("Admin","Login_ErrNum",0)
if GetCookies("Admin","Login_ErrNum")>5 then 
	Alert "登陆多次失败，系统禁止你提交登陆并记录下您的登陆信息!","?"
	Netlog adminname,"用户登陆尝试5次失败,尝试密码："&password
End if
	adminpwd=md5(adminpwd,32)
	set rs=conn.execute("select id,Admin_State from [enablesoft_manager] where password='"& adminpwd &"' and username='"& adminname &"'")
	if rs.eof then
    	rs.close
    	set rs=nothing
		SetCookies "Admin","Login_ErrNum", GetCookies("Admin","Login_ErrNum")+1
		Alert "您输入的用户名或密码不正确!","?"
	else	
		If rs("Admin_State")="1" Then Alert "你的登陆账户处于锁定状态，系统不允许登陆!","?"
			conn.execute("update [enablesoft_manager] set LoginTimes=LoginTimes+1 where id="& rs("id") &"")
			SetCookies "Admin","AdminName",adminname '建立session cookies
			SetCookies "Admin","AdminPassword",adminpwd
			rs.close
			set rs=Nothing
			
			conn.execute("delete From [enablesoft_log] where LogTime<"& SqlNowString &"-7")
			Netlog adminname,"成功登陆系统"
			CloseConn
			Response.Redirect "main_index.asp"
	end if
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<LINK href="inc/style.css" type=text/css rel=stylesheet>
<script language="javascript" src="inc/js.js"></script>
<title><%=SysInfo(0)%></title>
<script language="javascript">
function FormCheck(theForm){
if(!checkEmpty(theForm.adminname,"请输入登陆用户名！")) return false;
if(!checkEmpty(theForm.adminpwd,"请输入登陆密码！")) return false;
if(!checkEmpty(theForm.checkcode,"请输入安全码！")) return false;
return true;
}
</script>
<style type="text/css">
	body{background-color:#002779;}
	.input {
	BORDER: #000000 1px solid;FONT-FAMILY: "宋体"; BACKGROUND-COLOR: #ffffff;
}
</style>
</head>
<body onLoad="loginform.adminname.focus();">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" style="padding-bottom:100px;"><table width="460" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/login_1.jpg" width="190" height="23"></td>
      </tr>
    </table>
      <table width="460" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="images/login_2.jpg" width="460" height="142"></td>
        </tr>
      </table>
      <table width="460" border="0" cellspacing="0" cellpadding="0">
        <TR>
          <TD bgColor="#eeeeee" height="6"> </TD>
	    </TR>
	  </TABLE>
      <table width="460" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr>
          <td align="center" style="padding-top:15px;"><table width="300" border="0" cellspacing="0" cellpadding="0">
            <form name="loginform" method="post" action="?action=loginchk" onSubmit="return FormCheck(this);">
              <tr>
                <td width="35%" height="30" align="right"><strong>用户名：</strong></td>
                <td><INPUT style="width:150px;" class="input" maxLength="20" type="text" name="adminname" autocomplete="off"></td>
              </tr>
              <tr>
                <td height="30" align="right"><strong>密　码：</strong></td>
                <td><INPUT style="width:150px;" class="input" maxLength="20" type="password" name="adminpwd" autocomplete="off"></td>
              </tr>
              <tr>
                <td height="30" align="right"><strong>安全码：</strong></td>
                <td><INPUT style="width:40px;" class="input" maxLength="20" type="text" name="checkcode" autocomplete="off">
                    <script type="text/javascript">document.write("<img align='absmiddle' src='Inc/VerifyCode.asp?",Math.random(),"' style='cursor:pointer;' alt='请输入图片验证码' onClick=\"this.src='inc/VerifyCode.asp'\">")</script></td>
              </tr>
              <tr>
                <td colspan="2" align="center" height="40">
                  <input type="submit" name="Submit" value=" 提 交 " class="input">
                  &nbsp; <input type="reset" name="Submit" value=" 重 置 " class="input"></td>
              </tr>
            </form>
          </table></td>
        </tr>
      </table>
      <TABLE cellSpacing="0" cellPadding="0" width="460" bgColor="#ffffff" border="0">
          <TR>
            <TD><IMG height="10" src="images/login_3.gif" width="10"></TD>
            <TD align=right><IMG height="10" src="IMAGES/login_4.gif" width="10"></TD>
          </TR>
      </TABLE>
      </td>
  </tr>
</table>
</body>
</html>
