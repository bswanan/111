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

stype=checkstr(request.QueryString("stype"))
keys=checkstr(request.QueryString("keys"))

urlQuery="&stype="& stype &"&keys="& keys

action=request.QueryString("action")
select case action
	case "add" : CheckString("11"):call add
	case "list" : CheckString("12"):call list
	case "edit" : CheckString("13"):call edit
	case "upsave" :call upsave
end select

sub list
%>
<table width="99%" border="0" align="center" cellpadding="4" cellspacing="1" class="tablebk" style="border-collapse: collapse">
  <tr>
    <td colspan="9" class="table_titlebg">分类列表</td>
  </tr>
  <tr>
    <td colspan="9" align="center" class="table_trbg01"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
		<form name="Form2" method="get" action="?">
        <td align="right">  
查找：    <select name="stype">
          <option value="1"<%=SetSelected(stype,"1")%>>分类名称</option>
        </select> &nbsp; <input name="keys" type="text" id="keys" size="15" maxlength="30" class="INPUT" value="<%=keys%>">
        <input type="submit" name="" value="搜索"><input type="hidden" name="action" value="list"></td></form>
      </tr>
    </table>	</td>
  </tr>
<form name="Form1" method="post" action="?action=upsave&page=<%=page%>&stype=<%=stype%>&keys=<%=keys%>" onSubmit="return confirm('确定要执行此操作吗？\n\n注意：执行删除时将删除所属的分类');">
  <tr>
    <td align="center" class="table_trbg02"><strong>ID</strong></td>
    <td align="center" class="table_trbg02"><strong>分类名称</strong></td>
    <td align="center" class="table_trbg02"><strong>分类拼音</strong></td>
    <td align="center" class="table_trbg02"><strong>试用时间(分钟)</strong></td>
    <td align="center" class="table_trbg02"><strong>试用数</strong></td>
    <td align="center" class="table_trbg02"><strong>信息排序</strong></td>
    <td align="center" class="table_trbg02">注册码数量</td>
    <td align="center" class="table_trbg02"><strong>操作</strong></td>
    <td align="center" class="table_trbg02"><strong>选择</strong></td>
  </tr>
<%
dim PageMaxSize 
PageMaxSize=12	'每页几条
Set MyPage = New XdownPage	'创建对象

if keys<>"" then
if stype="1" then sql2 = sql2 & " and classname like '%"& keys &"%' "
end if

MyPage.GetSQL ="SELECT * from [enablesoft_mainclass] where 1=1 "& sql2 &" order by orders asc"
MyPage.PageSize = PageMaxSize  '设置每一页的记录条数据为10条
Set rs = MyPage.GetRS()  '返回Recordset

If rs.eof Then
	response.Write("<tr><td height=""30"" align=""center"" colspan=""6"" class=""table_trbg02"">没有任何信息！</td></tr>")
else
For i=1 To MyPage.PageSize  '显示数据
	If Not rs.eof Then
%>
  <tr>
    <td align="center" class="table_trbg02"><%=rs("ID")%></td>
    <td align="center" class="table_trbg02"><a href="?action=edit&id=<%=rs("ID")%>&page=<%=page%><%=urlQuery%>"><%=rs("classname")%></a></td>
    
     <td align="center" class="table_trbg02"><%=rs("classcode")%></td>
    <td align="center" class="table_trbg02"><%=rs("trytimer")%></td>
    <td align="center" class="table_trbg02"><%=rs("tryhits")%></td>
    <td align="center" class="table_trbg02"><input name="orders" type="text" class="input" id="orders" value="<%=rs("orders")%>" size="6" maxlength="5"></td>
    <td align="center" class="table_trbg02"><%=conn.execute("select count(id) from [enablesoft_user] where lx='"& rs("classcode") &"'")(0)%></td>
    <td align="center" class="table_trbg02"><a href="?action=edit&id=<%=rs("ID")%>&page=<%=page%><%=urlQuery%>">编辑</a></td>
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
   		<td colspan="9" align="right" class="table_trbg02">
		<input type="checkbox" name="chkall" value="on" onClick="CheckAll(this.form,'id')" />
      全选
        <select name="point">
          <option value="">操作方式</option>
          <option value="1">更新</option>
       		  <option value="2">删除</option>
        </select>
        <input type="submit" name="Submit" value="执 行">
&nbsp;
<input type="reset" name="Submit2" value="重 写">		</td>
    </tr>
  </form>
   <tr>
     <td colspan="9" align="center" class="table_trbg02"><%if MyPage.ShowTotalRecord>0 then MyPage.ShowPage()%></td>
   </tr>
</table>
<%end sub

sub add
if request.QueryString("act")="addsave" then
	
	classname=GetForm("classname")
	classcode=GetForm("classcode")
	orders=GetForm("orders")
trytimer=GetForm("trytimer")
	if classname="" then ErrMsg = ErrMsg & "● 请输入分类名称；"
	if classcode="" then ErrMsg = ErrMsg & "● 请输入分类拼音；"
	if not isInteger(Orders) then ErrMsg = ErrMsg & "● 排序必须使用整数输入；\n"
if not isInteger(trytimer) then ErrMsg = ErrMsg & "● 试用时间必须使用整数输入；\n"

if ErrMsg="" then
		if not conn.execute("select id from [enablesoft_mainclass] where classcode='"& classcode &"'").eof then
			ErrMsg = ErrMsg & "此类别已经被使用，请用其他类别重试；"
			FoundErr=true
		end if
	end if



	if ErrMsg="" then
		conn.execute("insert into [enablesoft_mainclass](classname,classcode,orders,trytimer)values('"& classname &"','"& classcode &"',"& orders &","& trytimer &")")
		alert "分类 "& classname &" 添加成功；","?action=add"
	end if
	if ErrMsg<>"" then response.Write(SetErrMsg(ErrMsg))
end if
%>

<table width="99%"  border="0" align="center" cellpadding="4" cellspacing="1" bordercolordark="#F1F3F5" class="tablebk">
  <form name="Form1" method="post" action="?action=add&act=addsave">
    <tr>
      <td height="15" colspan="2" class="table_titlebg">分类添加</td>
    </tr>
    <tr>
      <td width="40%" height="15" align="right" class="table_trbg02"><strong>分类名称：</strong></td>
      <td class="table_trbg02"><input name="classname" type="text" class="input" id="classname" size="40" value="<%=classname%>"></td>
    </tr>
    
    <tr>
      <td width="40%" height="15" align="right" class="table_trbg02"><strong>分类拼音：</strong></td>
      <td class="table_trbg02"><input name="classcode" type="text" class="input" id="classcode" size="40" value="<%=classcode%>"></td>
    </tr>
    <%
    set rs=conn.execute("select sys_trytimer from [enablesoft_config]")
if not rs.eof then
	
	trytimer=rs("sys_trytimer")

end if
	rs.close:set rs=nothing%>
       <tr>
      <td width="40%" height="15" align="right" class="table_trbg02"><strong>试用时间：</strong></td>
      <td class="table_trbg02"><input name="trytimer" type="text" class="input" id="trytimer" size="8" value="<%=trytimer%>">(分钟)0为不提供试用</td>
    </tr>

    <tr>
      <td height="15" align="right" class="table_trbg02"><p><strong>信息排序：</strong></p>      </td>
      <td class="table_trbg02"><input name="orders" type="text" class="input" id="orders" value="<%if orders="" then response.Write("0") else response.Write(orders)%>" size="14" />
整型，从小到大排序</td>
    </tr>
    <tr>
      <td height="35" colspan="2" align="center" class="table_trbg02"><input type="submit" name="Submit" value=" 完成 ">
        &nbsp;&nbsp;&nbsp;
        <input type="reset" name="Submit" value=" 重置 "></td>
    </tr>
  </form>
</table>

<%
end sub 
sub edit
id=request.QueryString("id")
if not isInteger(id) then alert "参数出错；","back"
set rs=conn.execute("select * from [enablesoft_mainclass] where id="& id &"")
if rs.eof then
	alert "参数出错，找不到此信息；","back"
else
classname=Rs("classname")
classcode=Rs("classcode")
orders=Rs("orders")
trytimer=Rs("trytimer")
end if
rs.close : set rs=nothing


if request.QueryString("act")="editsave" then
	classname=GetForm("classname")
		orders=GetForm("orders")
trytimer=GetForm("trytimer")
	if classname="" then ErrMsg = ErrMsg & "● 请输入分类名称；"
	if not isInteger(Orders) then ErrMsg = ErrMsg & "● 排序必须使用整数输入；\n"
if not isInteger(trytimer) then ErrMsg = ErrMsg & "● 试用时间必须使用整数输入；\n"
	if ErrMsg="" then
		conn.execute("update [enablesoft_mainclass] set classname='"& classname &"',orders="& orders &",trytimer="& trytimer &" where id="& id &"")
		alert "分类编辑保存成功；","?action=edit&page="& page &"&id="& id & urlQuery
	end if
	if ErrMsg<>"" then response.Write(SetErrMsg(ErrMsg))
end if

%>
<table width="99%"  border="0" align="center" cellpadding="4" cellspacing="1" bordercolordark="#F1F3F5" class="tablebk">
  <form name="Form1" method="post" action="?action=edit&act=editsave&id=<%=id%>&page=<%=page%><%=urlQuery%>">
    <tr>
      <td height="15" colspan="2" class="table_titlebg">分类编辑</td>
    </tr>
    <tr>
      <td width="40%" height="15" align="right" class="table_trbg02"><strong>分类名称：</strong></td>
      <td class="table_trbg02"><input name="classname" type="text" class="input" id="classname" size="40" value="<%=classname%>"></td>
    </tr>
    
    
           <tr>
      <td width="40%" height="15" align="right" class="table_trbg02"><strong>试用时间：</strong></td>
      <td class="table_trbg02"><input name="trytimer" type="text" class="input" id="trytimer" size="8" value="<%=trytimer%>">(分钟)0为不提供试用</td>
    </tr>
    
       <tr>
      <td height="15" align="right" class="table_trbg02"><p><strong>信息排序：</strong></p></td>
      <td class="table_trbg02"><input name="orders" type="text" class="input" id="orders" value="<%if orders="" then response.Write("0") else response.Write(orders)%>" size="14" />
        整型，从小到大排序</td>
    </tr>
    <tr>
      <td height="35" colspan="2" align="center" class="table_trbg02"><input type="submit" name="Submit3" value=" 保存 ">
        &nbsp;&nbsp;&nbsp;
        <input type="button" name="Submit22" onClick="javascript:window.location='?action=list&page=<%=page%><%=urlQuery%>';" value=" 返回 "></td>
    </tr>
  </form>
</table>
<%
end sub
sub upsave
	point=GetForm("point")
	if point="" then alert "您没有指定操作方式!","back"

if point="1" Then
	CheckString("14")
	For i=1 to request.form("hideid").count
		orders = trim(request.form("orders")(i))
		IF Not isInteger(orders) Then alert "输入的排序格式不正确!","back"
	Next
	For i=1 to request.form("hideid").count
		hideid = trim(request.form("hideid")(i))
		orders = trim(request.form("orders")(i))
		conn.Execute("update [enablesoft_mainclass] Set orders="& orders &" where ID="& hideid &"")
	Next
	alert "信息更新成功！","?page="& page & urlQuery
end if

if point="2" Then
	CheckString("15")
	id=GetForm("id")
	if id="" then alert "请选择要删除的信息!","back"
	conn.execute("delete From [enablesoft_mainclass] where id in ("& id &")")
	conn.execute("delete From [enablesoft_mainaddress] where classid in ("& id &")")
	alert "所选删除成功！","?page="& page & urlQuery
end if 

end sub
%>
</body>
</html>