<!--#include file="inc_CheckLogin.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/page_cls.asp"-->
<html>
<head>
<title>����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="javascript" src="inc/js.js"></script>
<link rel="stylesheet" href="inc/style.css">
</head>
<body>
<%
action=request.QueryString("action")
Admin_State=checkstr(request.QueryString("Admin_State"))
select case action
	case "edit" :CheckString("04"): call edit
	case "del" : call del
	case "list" : CheckString("110"):call main

	case else
		call main
end select

sub main
CheckString("02")
if GetForm("act")="addsave" Then
	CheckString("03")
	UserName=GetForm("UserName")
	lx=GetForm("lx")
	if strLength(UserName)<1 then ErrMsg = ErrMsg & "�� ע���벻��Ϊ�գ�\n"
		if ErrMsg="" then
		if not CheckName(UserName) then ErrMsg = ErrMsg & "�� ��½�˻������зǷ��ַ���\n"
		
	end if
	if ErrMsg="" then
		if not conn.execute("select id from [enablesoft_user] where UserName='"& UserName &"'").eof then
			ErrMsg = ErrMsg & "��ע�����Ѿ���ʹ�ã���������ע�������ԣ�"
			FoundErr=true
		end if
	end if
	
	if ErrMsg="" then
		PassWord=md5(UserName,32)
		lxp=md5(lx,16)
		conn.execute("insert into[enablesoft_user](UserName,[PassWord],LoginTimes,Truename,Joindate,Admin_State,lx,lxp)values('"& UserName &"','"& PassWord &"',0,'"& Truename &"',"& SqlNowString &",0,'"& lx &"','"& lxp &"')")
		Netlog A_UserName,"�����ע����"& UserName &""
		alert "ע������ӳɹ���","?"
	end if
	if ErrMsg<>"" then response.Write(SetErrMsg(ErrMsg))
end if
%>


<table width="99%" border="0" align="center" cellpadding="5" cellspacing="1" class="tablebk" style="border-collapse: collapse">
  <tr>
    <td colspan="4" align="center" class="table_trbg01"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="40%">��</td>
		<form name="Form1" method="get" action="?">
		<input name="act" value="list" type="hidden">
        <td width="60%" align="right">
ע��״̬��    
  <select name="Admin_State">
          <option value="0">����</option>
          <option value="1">����</option>
        </select>
        <input type="submit" name="Submit" value="����"></td>
		</form>
      </tr>
    </table>	</td>
  </tr>
  </table>
  
<table width="99%" border="0" align="center" cellpadding="5" cellspacing="1" class="tablebk" style="border-collapse: collapse">
<form name="Form1" action="?" method="post">
<input name="act" value="addsave" type="hidden">
  <tr>
    <td colspan="2" class="table_titlebg">ע�������</td>
  </tr>
  <tr>
    <td width="41%" align="right" class="table_trbg02"><strong>ע���룺</strong></td>
    <td width="59%" class="table_trbg02"><span class="table_trbg02">
      <input name="UserName" type="text" class="INPUT" id="UserName" size="32" value="<%=UserName%>">
    </span></td>
  </tr>
 
 
  <tr>
    <td width="41%" align="right" class="table_trbg02"><strong>���</strong></td>
    <td width="59%" class="table_trbg02"><span class="table_trbg02">
     
 <select name="lx">
<%
set rs1=server.createobject("adodb.recordset")
	sql="select * from [enablesoft_mainclass] order by orders asc "
	rs1.open sql,conn,1,1
	do while not rs1.eof%>
	 <option value="<%=rs1("classcode")%>"><%=rs1("classname")%></option>

	<%
	rs1.movenext
	loop
	rs1.close
	set rs1=nothing
%>
 </select>
    </span></td>
  </tr>


       
        
   <tr>
    <td height="40" colspan="2" align="center" class="table_trbg02"><input type="submit" name="Submit" value="�ύ"> 
      &nbsp; 
      <input type="reset" name="Submit" value="����"></td>
  </tr>
  </form>
</table>

<br class="table_br" />
<table width="99%" border="0" align="center" cellpadding="5" cellspacing="1" class="tablebk" style="border-collapse: collapse">
  <tr>
    <td colspan="9" class="table_titlebg">ע�����б�</td>
  </tr>
  <tr>
    <td width="5%" align="center" class="table_trbg01"><strong>ID</strong></td>
    <td width="27%" align="center" class="table_trbg01"><strong>ע����</strong></td>
    <td width="9%" align="center" class="table_trbg01"><strong>���</strong></td>
    <td width="8%" align="center" class="table_trbg01"><strong>��½����</strong></td>
    <td width="5%" align="center" class="table_trbg01"><strong>״̬</strong></td>
    <td width="10%" align="center" class="table_trbg01"><strong>����ʱ��</strong></td>
    <td width="10%" align="center" class="table_trbg01"><strong>�״ε�½ʱ��</strong></td>
    <td width="9%" align="center" class="table_trbg01"><strong>����ʱ��</strong></td>
<td align="center" class="table_trbg01" width="6%"><strong>����</strong></td>
  </tr>
  
  
  <%
dim PageMaxSize 
PageMaxSize=18	'ÿҳ����
Set MyPage = New XdownPage	'��������

if Admin_State<>"" then
	sql2 = sql2 & " and Admin_State="&Admin_State&""
end if


MyPage.GetSQL ="SELECT * FROM [enablesoft_user]  where 1=1 "& sql2 &" order by id desc"
MyPage.PageSize = PageMaxSize  '����ÿһҳ�ļ�¼������Ϊ10��
Set rs = MyPage.GetRS()  '����Recordset

If rs.eof Then
	response.Write("<tr><td height=""30"" align=""center"" colspan=""4"" class=""table_trbg02"">û���κ���Ϣ��</td></tr>")
else
For i=1 To MyPage.PageSize  '��ʾ����
	If Not rs.eof Then
	
	set rs1=server.createobject("adodb.recordset")
	sql="select classname from [enablesoft_mainclass] where classcode='"& rs("lx") &"'"
	rs1.open sql,conn,1,1
	if not rs1.eof then
	lxs=rs1("classname")
	end if
	rs1.close
	set rs1=nothing

%>
  <tr>
    <td align="center" class="table_trbg02"<%=classname%>><%=rs("id")%></td>
    <td align="center" class="table_trbg02"<%=classname%>><%=rs("username")%></td>
    
    <td align="center" class="table_trbg02"<%=classname%>><%=lxs%></td>
    
    <td align="center" class="table_trbg02"<%=classname%>><%=rs("LoginTimes")%></td>
    <td align="center" class="table_trbg02"<%=classname%>><%if rs("Admin_State")="0" then response.Write("����") else response.Write("<span class=""red"">����</span>")%></td>
    <td align="center" class="table_trbg02"<%=classname%>><%=rs("joindate")%></td>
    <td align="center" class="table_trbg02"<%=classname%>><%=rs("begindate")%></td>
    <td align="center" class="table_trbg02"<%=classname%>><%=rs("enddate")%></td>
      <td align="center" class="table_trbg02"<%=classname%> width="6%"><A href="?action=edit&id=<%=rs("id")%>">�༭</A></td>
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
     <td colspan="9" align="center" class="table_trbg02"><%if MyPage.ShowTotalRecord>0 then MyPage.ShowPage()%></td>
   </tr>

</table>
<%end sub



sub edit
	CheckString("04")
id=checkstr(request.QueryString("id"))
if not isInteger(id) then alert "�������ݳ��������ԣ�","back"
set rs=conn.execute("select * from [enablesoft_user] where id="& id &"")
if rs.eof then
	alert "û���ҵ���ע���룬�����ԣ�","back"
else
	username=rs("username")
	enddate2=rs("enddate")
end if
	rs.close:set rs=nothing


if GetForm("act")="editsave" then


	enddate=GetForm("enddate")
			
	if ErrMsg="" then
		conn.execute("update [enablesoft_user] set enddate='"& enddate &"'  where id="& id &"")
		
		if cstr(A_UserID)=cstr(id) then
		SetCookies "Admin","AdminName",UserName '����session cookies
			if PassWord<>"" then SetCookies "Admin","AdminPassword",PassWord
		end if
		Netlog A_UserName,"�༭ע����"& UserName &""
		alert "�༭����ɹ���","?action=edit&id="& id &""
	end if
	if ErrMsg<>"" then response.Write(SetErrMsg(ErrMsg))
end if
%>
<table width="99%" border="0" align="center" cellpadding="5" cellspacing="1" class="tablebk" style="border-collapse: collapse">
<form name="Form1" action="?action=edit&id=<%=id%>" method="post">
<input name="act" value="editsave" type="hidden">
  <tr>
    <td colspan="2" class="table_titlebg">ע����༭</td>
  </tr>
  <tr>
    <td width="18%" align="right" class="table_trbg02"><strong>ע���룺</strong></td>
    <td width="82%" class="table_trbg02"><span class="table_trbg02">
      <%=UserName%>
    </span></td>
  </tr>
  
  <tr>
    <td width="18%" align="right" class="table_trbg02"><strong>����ʱ�䣺</strong></td>
    <td width="82%" class="table_trbg02"><span class="table_trbg02">
      <input name="enddate" type="text" class="INPUT" id="enddate" size="30" value="<%=enddate2%>">
    </span></td>
  </tr>

      <tr>
    <td height="40" colspan="2" align="center" class="table_trbg02"><input type="submit" name="Submit" value="�ύ"> 
      &nbsp; 
      <input type="button" name="Submit" value="����" onClick="window.location='?';"></td>
  </tr>
  </form>
</table>
<%end sub


%>


</body>
</html>