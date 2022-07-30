<!--#include file="inc_CheckLogin.asp"-->
<!--#include file="inc/page_cls.asp"-->
<html>
<head>
<title>管理</title>
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
    <td colspan="6" class="table_titlebg">网站日志</td>
  </tr>
  <tr>
    <td align="center" class="table_trbg01"><strong>ID</strong></td>
    <td align="center" class="table_trbg01"><strong>操作用户</strong></td>
    <td align="center" class="table_trbg01"><strong>操作事件</strong></td>
    <td align="center" class="table_trbg01"><strong>来源地址</strong></td>
    <td align="center" class="table_trbg01"><strong>时间</strong></td>
    <td align="center" class="table_trbg01"><strong>IP地址</strong></td>
  </tr>
<%
dim PageMaxSize 
PageMaxSize=15	'每页几条
Set MyPage = New XdownPage	'创建对象

MyPage.GetSQL ="select * from [enablesoft_log] order by id desc"
MyPage.PageSize = PageMaxSize  '设置每一页的记录条数据为10条
Set rs = MyPage.GetRS()  '返回Recordset

If rs.eof Then
	response.Write("<tr><td height=""30"" align=""center"" colspan=""6"" class=""table_trbg02"">没有任何日志信息！</td></tr>")
else
For i=1 To MyPage.PageSize  '显示数据
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