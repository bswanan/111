<!--#include file="inc_CheckLogin.asp"-->
<!--#include file="inc/page_cls.asp"-->
<html>
<head>
<title>管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/style.css">
<script language="javascript" src="inc/js.js"></script>
</head>
<body>
<%dim page
page=request.QueryString("page")
classid=checkstr(request.QueryString("classid"))
keys=checkstr(request.QueryString("keys"))

urlQuery="&classid="& classid &"&keys="& keys

action=request.QueryString("action")
select case action
	case "add" : CheckString("16"):call add
	case "addsave" : CheckString("16"):call addsave
	case "list" : CheckString("110"):call list
	case "upsave" : call upsave
end select

sub list
%>
<table width="99%" border="0" align="center" cellpadding="3" cellspacing="1" class="tablebk" style="border-collapse: collapse">
  <tr>
    <td colspan="4" class="table_titlebg">邮件地址列表</td>
  </tr>
  <tr>
    <td colspan="4" align="center" class="table_trbg01"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="40%">　</td>
		<form name="Form2" method="get" action="?">
		<input type="hidden" name="action" value="list">
        <td width="60%" align="right">
查找：    
  <select name="classid">
          <option value="">所有分类</option>
	  <%set rs=conn.execute("select * from [enablesoft_mainclass] order by orders, id desc")
	  if rs.eof then rs.close:set rs=nothing:alert "● 请选添加邮址分类；","back"
	  do while not rs.eof%>
	  <option value="<%=rs("id")%>"<%=SetSelected(rs("id"),classid)%>><%=rs("classname")%></option>
	  <%
	  rs.movenext
	  loop
	  rs.close:set rs=nothing
	  %>		  
        </select> &nbsp; <input name="keys" type="text" id="keys" size="15" maxlength="30" class="INPUT" value="<%=keys%>">
        <input type="submit" name="" value="搜索"></td>
		</form>
      </tr>
    </table>	</td>
  </tr>
<form name="Form1" method="post" action="?action=upsave&page=<%=page%><%=urlQuery%>" onSubmit="return confirm('确定要执行此操作吗？\n\n注意：执行删除将不可恢复数据');">
  <tr>
    <td align="center" class="table_trbg02"><strong>ID</strong></td>
    <td align="center" class="table_trbg02"><strong>所属分类</strong></td>
    <td align="center" class="table_trbg02"><strong>邮件地址</strong></td>
    <td align="center" class="table_trbg02"><strong>选择</strong></td>
  </tr>
<%
dim PageMaxSize 
PageMaxSize=12	'每页几条
Set MyPage = New XdownPage	'创建对象

if keys<>"" then
	sql2 = sql2 & " and A.mainaddress like '%"& keys &"%' "
end if
if isInteger(classid) then sql2 = sql2 & " and A.classid = "& classid &" "

MyPage.GetSQL ="SELECT A.id,A.mainaddress,B.classname FROM [enablesoft_mainaddress] AS A LEFT JOIN [enablesoft_mainclass] AS B ON A.classid = B.id where 1=1 "& sql2 &" order by A.id desc"
MyPage.PageSize = PageMaxSize  '设置每一页的记录条数据为10条
Set rs = MyPage.GetRS()  '返回Recordset

If rs.eof Then
	response.Write("<tr><td height=""30"" align=""center"" colspan=""4"" class=""table_trbg02"">没有任何信息！</td></tr>")
else
For i=1 To MyPage.PageSize  '显示数据
	If Not rs.eof Then
%>
  <tr>
    <td align="center" class="table_trbg02"><%=rs("ID")%></td>
    <td align="center" class="table_trbg02"><%=rs("classname")%></td>
    <td align="center" class="table_trbg02"><%=rs("mainaddress")%></td>
    <td align="center" class="table_trbg02"><input type="checkbox" name="id" value="<%=rs("ID")%>"><input type="hidden" name="hideid" value="<%=rs("ID")%>"></td>
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
   		<td colspan="4" align="right" class="table_trbg02">
		<input type="checkbox" name="chkall" value="on" onClick="CheckAll(this.form,'id')" />
      全选
        <select name="point">
          <option value="">操作方式</option>
		  <option value="1">删除</option>
        </select>
        <input type="submit" name="Submit" value="执 行">
&nbsp;
<input type="reset" name="Submit2" value="重 写">		</td>
    </tr>
  </form>
   <tr>
     <td colspan="4" align="center" class="table_trbg02"><%if MyPage.ShowTotalRecord>0 then MyPage.ShowPage()%></td>
   </tr>
</table>
<%end sub

sub add
%>

<table width="99%"  border="0" align="center" cellpadding="4" cellspacing="1" bordercolordark="#F1F3F5" class="tablebk">
  <form name="Form1" method="post" action="?action=addsave">
    <tr>
      <td height="15" colspan="2" class="table_titlebg">邮件地址添加</td>
    </tr>
    <tr>
      <td width="32%" height="15" align="right" class="table_trbg02"><strong>归属分类：</strong></td>
      <td class="table_trbg02"><select name="classid">
	  <%set rs=conn.execute("select * from [enablesoft_mainclass] order by orders, id desc")
	  if rs.eof then rs.close:set rs=nothing:alert "● 请选添加邮址分类；","back"
	  do while not rs.eof%>
	  <option value="<%=rs("id")%>"><%=rs("classname")%></option>
	  <%
	  rs.movenext
	  loop
	  rs.close:set rs=nothing
	  %>
	  </select></td>
    </tr>
    <tr>
      <td height="15" align="right" class="table_trbg02"><strong>邮址列表：</strong><br>回车或逗号分隔</td>
      <td class="table_trbg02"><textarea name="maincontent" cols="64" rows="24" class="textarea"></textarea></td>
    </tr>
    <tr>
      <td height="35" colspan="2" align="center" class="table_trbg02"><input type="submit" name="Submit" value=" 添加 ">
        &nbsp;&nbsp;&nbsp;
      <input type="reset" name="Submit" value=" 重置 "></td>
    </tr>
  </form>
</table>

<%
end sub 
sub addsave
	classid=GetForm("classid")
	maincontent=GetForm("maincontent")

	if not isInteger(classid) then  alert "● 请选择归属分类；","back"
	if maincontent="" then alert "● 请输入邮址列表；","back"

	maincontent=replace(maincontent," ","")
	if maincontent="" then alert "● 请输入邮址列表；","back"

	maincontent=lcase(maincontent)
	maincontent=replace(maincontent,chr(13)&chr(10),",")
	maincontent_arr=split(maincontent,",")

		e=0
		for i=0 to ubound(maincontent_arr)
			if IsValidEmain(maincontent_arr(i)) then
				'检测重复
				set rs=conn.execute("select id from [enablesoft_mainaddress] where mainaddress='"& maincontent_arr(i) &"' and classid="& classid &"")			
				if not rs.eof then
					repeat_maincontent = repeat_maincontent & "," & maincontent_arr(i)			'重复就相加重复的emain地址
				Else
					e=e+1
					succeed_maincontent = succeed_maincontent & "," & maincontent_arr(i)			'重复就相加重复的emain地址
					conn.execute("insert into [enablesoft_mainaddress](classid,mainaddress)values("& classid &",'"& maincontent_arr(i) &"')")			'不重复就入库
				end if
				rs.close:set rs=nothing
			else
				err_maincontent = err_maincontent & "," & maincontent_arr(i)
			end if
		next
		if repeat_maincontent<>"" then repeat_maincontent = mid(repeat_maincontent,2)
		if err_maincontent<>"" then err_maincontent = mid(err_maincontent,2)
		if succeed_maincontent<>"" then succeed_maincontent = mid(succeed_maincontent,2)
%>
<table width="99%"  border="0" align="center" cellpadding="4" cellspacing="1" bordercolordark="#F1F3F5" class="tablebk">
  <form name="Form1" method="post" action="?action=edit&act=editsave&id=<%=id%>&page=<%=page%><%=urlQuery%>">
    <tr>
      <td height="15" colspan="3" class="table_titlebg">邮件地址编辑</td>
    </tr>
    <tr>
      <td width="33%" height="15" align="center" class="table_trbg02">成功的邮址(已入库)</td>
      <td width="33%" align="center" class="table_trbg02">失败的邮址</td>
      <td width="33%" align="center" class="table_trbg02">重复的邮址</td>
    </tr>
    <tr>
      <td height="15" align="center" class="table_trbg02"><textarea name="textarea1" cols="26" rows="28" style="width:85%"><%=replace(succeed_maincontent,",",vbcrlf)%></textarea></td>
      <td align="center" class="table_trbg02"><textarea name="textarea2" cols="26" rows="28" style="width:85%"><%=replace(err_maincontent,",",vbcrlf)%></textarea></td>
      <td align="center" class="table_trbg02"><textarea name="textarea3" cols="26" rows="28" style="width:85%"><%=replace(repeat_maincontent,",",vbcrlf)%></textarea></td>
    </tr>
    <tr>
      <td height="35" colspan="3" align="center" class="table_trbg02"><input type="button" name="Submit22" onClick="javascript:window.location='?action=add';" value="返回继续添加"></td>
    </tr>
  </form>
</table>

<%end sub
sub upsave
	point=GetForm("point")
	if point="" then alert "您没有指定操作方式!","back"

	if point="1" Then
		CheckString("111")
		id=GetForm("id")
		if id="" then alert "请选择要删除的信息!","back"
		conn.execute("delete From [enablesoft_mainaddress] where id in ("& id &")")
		alert "所选信息删除成功！","?action=list&page="& page & urlQuery
	end if
end sub
%>
</body>
</html>