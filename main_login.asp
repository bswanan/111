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

if action="loginchk" then '��½��֤
OpenConn
if CheckMake=False then Alert "���ύ��·�����󣬽�ֹ��վ���ⲿ�ύ�����벻Ҫ�ҼӲ�����","./"
adminname=GetForm("adminname")
adminpwd=GetForm("adminpwd")
checkcode=GetForm("checkcode")

if adminname="" then Alert "�������½�û�����","back"
if adminpwd="" then Alert "�������½���룡","back"
if checkcode="" then Alert "�����밲ȫ�룡","back"

if session("CheckCode")="" then Alert "��֤���Ѿ����ڣ������µ�¼��","?"
if Lcase(checkcode)<>LCase(CStr(session("CheckCode"))) then Alert "������ĸ������ϵͳ�����Ĳ�һ�£����������롣","?"
session("CheckCode")=""
if GetCookies("Admin","Login_ErrNum")="" then call SetCookies("Admin","Login_ErrNum",0)
if GetCookies("Admin","Login_ErrNum")>5 then 
	Alert "��½���ʧ�ܣ�ϵͳ��ֹ���ύ��½����¼�����ĵ�½��Ϣ!","?"
	Netlog adminname,"�û���½����5��ʧ��,�������룺"&password
End if
	adminpwd=md5(adminpwd,32)
	set rs=conn.execute("select id,Admin_State from [enablesoft_manager] where password='"& adminpwd &"' and username='"& adminname &"'")
	if rs.eof then
    	rs.close
    	set rs=nothing
		SetCookies "Admin","Login_ErrNum", GetCookies("Admin","Login_ErrNum")+1
		Alert "��������û��������벻��ȷ!","?"
	else	
		If rs("Admin_State")="1" Then Alert "��ĵ�½�˻���������״̬��ϵͳ�������½!","?"
			conn.execute("update [enablesoft_manager] set LoginTimes=LoginTimes+1 where id="& rs("id") &"")
			SetCookies "Admin","AdminName",adminname '����session cookies
			SetCookies "Admin","AdminPassword",adminpwd
			rs.close
			set rs=Nothing
			
			conn.execute("delete From [enablesoft_log] where LogTime<"& SqlNowString &"-7")
			Netlog adminname,"�ɹ���½ϵͳ"
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
if(!checkEmpty(theForm.adminname,"�������½�û�����")) return false;
if(!checkEmpty(theForm.adminpwd,"�������½���룡")) return false;
if(!checkEmpty(theForm.checkcode,"�����밲ȫ�룡")) return false;
return true;
}
</script>
<style type="text/css">
	body{background-color:#002779;}
	.input {
	BORDER: #000000 1px solid;FONT-FAMILY: "����"; BACKGROUND-COLOR: #ffffff;
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
                <td width="35%" height="30" align="right"><strong>�û�����</strong></td>
                <td><INPUT style="width:150px;" class="input" maxLength="20" type="text" name="adminname" autocomplete="off"></td>
              </tr>
              <tr>
                <td height="30" align="right"><strong>�ܡ��룺</strong></td>
                <td><INPUT style="width:150px;" class="input" maxLength="20" type="password" name="adminpwd" autocomplete="off"></td>
              </tr>
              <tr>
                <td height="30" align="right"><strong>��ȫ�룺</strong></td>
                <td><INPUT style="width:40px;" class="input" maxLength="20" type="text" name="checkcode" autocomplete="off">
                    <script type="text/javascript">document.write("<img align='absmiddle' src='Inc/VerifyCode.asp?",Math.random(),"' style='cursor:pointer;' alt='������ͼƬ��֤��' onClick=\"this.src='inc/VerifyCode.asp'\">")</script></td>
              </tr>
              <tr>
                <td colspan="2" align="center" height="40">
                  <input type="submit" name="Submit" value=" �� �� " class="input">
                  &nbsp; <input type="reset" name="Submit" value=" �� �� " class="input"></td>
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
