<!--#include file="inc_CheckLogin.asp"-->
<html>
<head>
<title>����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/style.css">
</head>
<body>
<%CheckString("01")%>
<%
dim action
action=request.QueryString("action")

set rs=conn.execute("select * from [enablesoft_config]")
if not rs.eof then
	sys_name=rs("sys_name")
	sys_trytimer=rs("sys_trytimer")

end if
	rs.close:set rs=nothing


if action="editsave" then
	sys_name=GetForm("sys_name")
	sys_trytimer=GetForm("sys_trytimer")


	if not isInteger(sys_trytimer) then ErrMsg = ErrMsg & "�� ����ʱ�����ʹ���������룻\n"	
	
	if ErrMsg="" then
		conn.execute("update [enablesoft_config] set sys_name='"& sys_name &"',sys_trytimer='"& sys_trytimer &"'")
		
	Application.Lock()	
	Application(CacheName&"_sysconfig")=""
	Application.UnLock()
		alert "�ύ����ɹ���","?"
	end if
	if ErrMsg<>"" then response.Write(SetErrMsg(ErrMsg))
end if
%>
<table width="99%" border="0" align="center" cellpadding="4" cellspacing="1" class="tablebk" style="border-collapse: collapse">
<form name="Form1" action="?action=editsave" method="post">
  <tr>
    <td colspan="2" class="table_titlebg">ϵͳ��Ϣ����</td>
  </tr>
  <tr>
    <td width="34%" align="right" class="table_trbg02"><strong>ϵͳ���ƣ�</strong></td>
    <td width="66%" class="table_trbg02"><input name="sys_name" type="text" class="input" size="40" value="<%=sys_name%>"></td>
  </tr>
  <tr>
    <td align="right" class="table_trbg02"><strong>Ĭ������ʱ�䣺</strong></td>
    <td class="table_trbg02"><input name="sys_trytimer" type="text" class="input" size="40" value="<%=sys_trytimer%>">(����)
      Ĭ��������õ�ʱ��</td>
  </tr>   
 
  <tr>
    <td height="40" colspan="2" align="center" class="table_trbg02"><input type="submit" name="Submit" value="�ύ"> 
      &nbsp; 
      <input type="reset" name="Submit" value="����"></td>
  </tr>
  </form>
</table>
</body>
</html>