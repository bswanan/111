<!--#include file="inc_CheckLogin.asp"-->
<%
action=Request("action")

if action="left" then 
	Call mainleft()
ElseIf action="top" Then
	Call maintop()
ElseIf action="hidelist" Then
	Call hidelist()
Else 

End If


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<LINK href="inc/style.css" type=text/css rel=stylesheet>
<script language="javascript" src="inc/js.js"></script>
<title><%=SysInfo(0)%> - ϵͳ����</title>
</head>
<frameset rows="32,*" border=0 frameborder="YES" name="top_frame"> 

  <frame src="?action=top" frameborder="NO" name="ads" scrolling="NO"  marginwidth="0" marginheight="0" noresize="noresize">
<frameset rows="675" border=0 name="bbs" framespacing="1"> 
        <frameset cols="175,9,*" border=0 name="bbs" framespacing="1"> 
        <frame src="?action=left"  name="list" marginwidth="0" marginheight="0">
        <frame src="?action=hidelist" SCROLLING="NO" name="hidelist" marginwidth="0" marginheight="0" noresize="noresize">
        <frame src="main_main.asp" name="main" marginwidth="0" marginheight="0">
		</frameset>
</frameset>
</frameset>
<noframes><body>�Բ���,�����������֧�ֿ��!</body></noframes>
</html>

<%
sub maintop		'***************************	����

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>����</title>
<style type="text/css">
<!--
body {margin:0px;FONT-SIZE: 12px;COLOR: #000000; FONT-FAMILY: "����";}
td{font-family:����; font-size: 12px; line-height: 20px;}

A:visited{TEXT-DECORATION: none;color: #000000;}
A:active{TEXT-DECORATION: none;color: #000000;}
A:link{text-decoration: none;color: #000000;}
A:hover {BORDER-BOTTOM: 1px dotted; BORDER-LEFT-WIDTH: 1px; BORDER-RIGHT-WIDTH: 1px; BORDER-TOP-WIDTH: 1px; COLOR: #666666; TEXT-DECORATION: none}
-->
</style>
</head>
<body background="images/main_top_bg.gif">
<table width="100%" height="32"  border="0" cellpadding="0" cellspacing="0" class="logobg">
  <tr> 
	<td width="145">��</td>
	<td width="100%">��</td>
	<td width="320" style="padding-bottom:3px"> 
	  <table width="320"  border="0" cellspacing="0" cellpadding="0">
		<tr> 
		 
		  <td width="20%" align="center"><a href="main_main.asp" target="main">ϵͳ��ҳ</a></td>
		  <td width="20%" align="center"><a href="javascript:checkclick('ȷ��Ҫ�˳�ϵͳ��','main_login.asp?action=logout');" target="_top">�˳���¼</a></td>
		  <td width="20%" align="center">&nbsp; </td>
		</tr>
	  </table>	</td>
  </tr>
</table>
</body>
</html>

<%end sub
sub hidelist	'***************************	���ز˵���ť
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>���ز˵���ť</title>
<style type="text/css">
<!--
body {margin:0;padding:0; background-color:#D0D0D0}
.hiddentable{width:9px;}
.hiddentop{height:6px;}
.hiddenl{background:url(images/hidelist02.gif) no-repeat left center;}
.hiddenr{background:url(images/hidelist01.gif) no-repeat left center;}
.hiddenbottom{height:9px;}
-->
</style>
<script language=javascript>
function HideList(ss)
{
	if (frmHide.liststatus.value==0)
	{
		eval("document.getElementById('arrow').className='hiddenr'");
		top.bbs.cols="0,9,*";
	}
	else
	{
		eval("document.getElementById('arrow').className='hiddenl'");
		top.bbs.cols="175,9,*";
	}
	frmHide.liststatus.value = 1 - frmHide.liststatus.value;

}
</script>
</head>
<body>
<table height="100%"  border="0" cellpadding="0" cellspacing="0" class="hiddentable">
  <tr>
    <td class="hiddentop"></td>
  </tr>
  <tr>
    <td class="hiddenl" onClick="HideList(arrow)" style="cursor:hand" id="arrow">
<form name=frmHide>
<input type=hidden name="liststatus" value=0>
</form>	
</td>
  </tr>
  <tr>
    <td class="hiddenbottom"></td>
  </tr>
</table>
</body>
</html>
<%end sub
sub mainleft		'***************************	��߲˵�
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��߲˵�</title>
<style type="text/css"> 
body{ margin:0px;FONT-SIZE: 12px;COLOR: #000000; FONT-FAMILY: "����";background-color: #F6F6F6;scrollbar-face-color:#D9D9D9;scrollbar-highlight-color:#FfFFFF;scrollbar-3dlight-color:#FfFFFF;scrollbar-darkshadow-color:#999999;scrollbar-shadow-color:#000000;scrollbar-arrow-color:#000000;scrollbar-track-color:#E4E4E4;}

TD{ font-family:����; font-size: 12px; line-height: 15px;}
a  { font:normal 12px ����; color:#000000; text-decoration:none; }
a:hover  { color:#124981;text-decoration:underline; }

.dtree {
	font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
	font-size: 12px;
	color: #666;
	white-space: nowrap;
	padding-left:3px;
}
.dtree img {
	border: 0px;
	vertical-align: middle;
}
.dtree a {
	color: #333;
	text-decoration: none;
}
.dtree a.node, .dtree a.nodeSel {
	white-space: nowrap;
	padding: 1px 2px 1px 2px;
}
.dtree a.node:hover, .dtree a.nodeSel:hover {
	color: #333;
	text-decoration: underline;
}
.dtree a.nodeSel {
	background-color: #c0d2ec;
}
.dtree .clip {
	overflow: hidden;
}
</style>
<SCRIPT LANGUAGE="JavaScript" Src="inc/dtree.js"></SCRIPT>
</head>
<body>
<TABLE width="158" cellSpacing="0" cellPadding="0" border="0" background="img/menu_1.gif" height=60>
<tr>
 <td colspan='2' align="center"><strong>��ӭ��,<%=a_truename%></strong></td>
</tr>
<tr>
  <td align="center"><img src="images/left_fold0.gif" border="0"> <a href="main_main.asp" target="main">ϵͳ��ҳ</a></td>
  <td align="center"><img src="images/left_fold0.gif" border="0"> <a href="javascript:checkclick('ȷ��Ҫ�˳�ϵͳ��','main_login.asp?action=logout');" target="_top">��ȫ�˳�</a></td>
</tr>
</table>
<div class="dtree">
	<script type="text/javascript">
		<!--
		d = new dTree('d');
		d.add(0,-1,'ϵͳ����˵�');
		d.add(1,0,'ϵͳ����');
		d.add(2,1,'ϵͳ������Ϣ','main_setting.asp','','main');
		d.add(3,1,'ϵͳ������Ա','main_manager.asp','','main');
		d.add(4,1,'�޸ĵ�½����','main_managerpwd.asp','','main');
		d.add(5,1,'ϵͳ��ȫ��־','main_loglist.asp','','main');
		
		
		d.add(8,0,'ע�������');
		d.add(9,8,'�������','main_mainclass.asp?action=add','','main');
		d.add(10,8,'�����б�','main_mainclass.asp?action=list','','main');

		d.add(11,8,'ע�������','main_user.asp','','main');
		
		document.write(d);

		//-->
	</script>
</div>
</body>
</html>
<%
Call CloseConn
End Sub
%>