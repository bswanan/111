<!--#include file="inc_CheckLogin.asp"-->
<!--#include file="inc/page_cls.asp"-->
<html>
<head>
<title>����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="javascript" src="inc/js.js"></script>
<link rel="stylesheet" href="inc/style.css"></head>
<body>
<%CheckString("07")
call main

sub main
%>
<br class="table_br" />
<table width="99%" border="0" align="center" cellpadding="5" cellspacing="1" class="tablebk" style="border-collapse: collapse">
  <tr>
    <td colspan="6" class="table_titlebg">��վ��־</td>
  </tr>
  <tr>
    <td align="center" class="table_trbg01"><strong>ID</strong></td>
    <td align="center" class="table_trbg01"><strong>�����û�</strong></td>
    <td align="center" class="table_trbg01"><strong>�����¼�</strong></td>
    <td align="center" class="table_trbg01"><strong>��Դ��ַ</strong></td>
    <td align="center" class="table_trbg01"><strong>ʱ��</strong></td>
    <td align="center" class="table_trbg01"><strong>IP��ַ</strong></td>
  </tr>
<%
dim PageMaxSize 
PageMaxSize=15	'ÿҳ����
Set MyPage = New XdownPage	'��������

MyPage.GetSQL ="select * from [enablesoft_log] order by id desc"
MyPage.PageSize = PageMaxSize  '����ÿһҳ�ļ�¼������Ϊ10��
Set rs = MyPage.GetRS()  '����Recordset

If rs.eof Then
	response.Write("<tr><td height=""30"" align=""center"" colspan=""6"" class=""table_trbg02"">û���κ���־��Ϣ��</td></tr>")
else
For i=1 To MyPage.PageSize  '��ʾ����
	If Not rs.eof Then

%>
  <tr>
    <td align="center" class="table_trbg02"><%=rs("id")%></td>
    <td align="center" class="table_trbg02"><%=rs("username")%></td>
    <td align="center" class="table_trbg02"><%=rs("Remark")%></td>
    <td align="center" class="table_trbg02"><%=rs("Geturl")%></td>
    <td align="center" class="table_trbg02"><%=rs("Logtime")%></td>
    <td align="center" class="table_trbg02"><%=rs("UserIP")%></td>
  </tr>
 <%
	rs.movenext
	Else
    	Exit For
	End If
next
End If
%>
   <tr>
    <td align="center" class="table_trbg02" colspan="6"><%if MyPage.ShowTotalRecord>0 then MyPage.ShowPage()%></td>
  </tr>
</table>
 <%end sub
%>
</body>
</html>