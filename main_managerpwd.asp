<!--#include file="inc_CheckLogin.asp"-->
<!--#include file="inc/md5.asp"-->
<html>
<head>
<title>����</title>
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
		ErrMsg = ErrMsg & "�� ������ɵ�½���룻\n"
	end If
	if ErrMsg="" then
		if strLength(PassWord2)<4 or strLength(PassWord2)>20  then
			ErrMsg = ErrMsg & "�� Ϊ���˻���ȫ��½����Ӧ����4-20���ַ���\n"
		end if
		if not CheckPassword(PassWord2) then
			ErrMsg = ErrMsg & "�� �µ�½��������зǷ��ַ���\n"
		end If
	End If
	
	if ErrMsg="" then
		if PassWord2<>PassWord3  then
			ErrMsg = ErrMsg & "�� ������������벻һ�£�\n"
		end If
	end if

	if ErrMsg="" then
		if conn.execute("select id from [enablesoft_manager] where [PassWord]='"& md5(PassWord,32) &"' and id="& A_UserID &"").eof then
			ErrMsg = ErrMsg & "�� ����������������������룻\n"
		end if
	end if
	
	
	if ErrMsg="" then
			PassWord2=md5(PassWord2,32)
		conn.execute("update [enablesoft_manager] set [PassWord]='"& PassWord2 &"' where id="& A_UserID &"")
		SetCookies "Admin","AdminPassword",PassWord2
		Netlog A_UserName,"�༭�����½����"
		alert "�༭����ɹ���","?"
	end if
	if ErrMsg<>"" then response.Write(SetErrMsg(ErrMsg))
end if
%>
<table width="99%" border="0" align="center" cellpadding="5" cellspacing="1" class="tablebk" style="border-collapse: collapse">
<form name="Form1" action="?" method="post">
<input name="act" value="editsave" type="hidden">
   <tr>
    <td colspan="2" class="table_titlebg">�޸ĵ�½����</td>
  </tr>
  <tr>
    <td width="40%" align="right" class="table_trbg02"><strong>��ǰ�û�����</strong></td>
    <td width="60%" class="table_trbg02"><strong><%=A_UserName%></strong></td>
  </tr>
  <tr>
    <td align="right" class="table_trbg02"><strong>��������룺</strong></td>
    <td class="table_trbg02"><input name="PassWord" type="password" class="INPUT" id="PassWord" size="30"> 
    </td>
  </tr>
  <tr>
    <td align="right" class="table_trbg02"><strong>���������룺</strong></td>
    <td class="table_trbg02"><input name="PassWord2" type="password" class="INPUT" id="PassWord2" size="30"> 
    </td>
  </tr>
  <tr>
    <td align="right" class="table_trbg02"><strong>ȷ�������룺</strong></td>
    <td class="table_trbg02"><input name="PassWord3" type="password" class="INPUT" id="PassWord3" size="30"></td>
  </tr>
  <tr>
    <td height="40" colspan="2" align="center" class="table_trbg02"><input type="submit" name="Submit" value="�ύ"> 
      &nbsp; 
      <input type="button" name="Submit" value="����" onClick="window.location='?';"></td>
  </tr>
  </form>
</table>
</body>
</html>