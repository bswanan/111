<!--#include file="inc_CheckLogin.asp"-->
<html>
<head>
<title>������ҳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/style.css">
</head>
<body>


<br class="table_br" />
<table width="99%" border="0" align="center" cellpadding="5" cellspacing="1" class="tablebk" style="border-collapse: collapse">
  <tr>
    <td colspan="2" class="table_titlebg">����������</td>
  </tr>
  <tr>
    <td width="50%" class="table_trbg02">�������������еĶ˿ڣ�<%=Request.ServerVariables("server_port")%></td>
    <td class="table_trbg02">�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr>
    <td class="table_trbg02">���������ƣ�<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
    <td class="table_trbg02">������IP��<%=Request.ServerVariables("LOCAL_ADDR")%></td>
  </tr>
  <tr>
    <td class="table_trbg02">վ������·����<%=Request.ServerVariables("path_translated")%></td>
    <td class="table_trbg02">����·����<%=Request.ServerVariables("server_name")%></td>
  </tr>
  <tr>
    <td class="table_trbg02">������Application������<%=Application.Contents.Count%> ��</td>
    <td class="table_trbg02">������Session������<%=Session.Contents.Count%> ��</td>
  </tr>
  <tr>
    <td class="table_trbg02">��������ǰʱ�䣺<%=now()%></td>
    <td class="table_trbg02">�ű����ӳ�ʱʱ�䣺<%=Server.ScriptTimeout%> ��</td>
  </tr>
  <tr>
    <td colspan="2" class="table_trbg02"><strong>IIS�Դ���ASP���</strong></td>
  </tr>
  <tr>
    <td class="table_trbg02">MSWC.AdRotator��<%If IsObjInstalled("MSWC.AdRotator") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
    <td class="table_trbg02">MSWC.BrowserType��
    <%If IsObjInstalled("MSWC.BrowserType") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
  </tr>
   <tr>
    <td class="table_trbg02">MSWC.NextLink��<%If IsObjInstalled("MSWC.NextLink") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
    <td class="table_trbg02">aMSWC.Tools��
    <%If IsObjInstalled("MSWC.Tools") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
  </tr>
   <tr>
    <td class="table_trbg02">MSWC.Status��<%If IsObjInstalled("MSWC.Status") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
    <td class="table_trbg02">MSWC.Counters��
    <%If IsObjInstalled("MSWC.Counters") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
  </tr>
  <tr>
    <td class="table_trbg02">IISSample.ContentRotator��<%If IsObjInstalled("IISSample.ContentRotator") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
    <td class="table_trbg02">IISSample.PageCounter��
    <%If IsObjInstalled("IISSample.PageCounter") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
  </tr>
   <tr>
    <td class="table_trbg02">MSWC.PermissionChecker��<%If IsObjInstalled("MSWC.PermissionChecker") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
    <td class="table_trbg02">Msxml2.FreeThreadedDOMDocument.3.0��
    <%If IsObjInstalled("Msxml2.FreeThreadedDOMDocument.3.0") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
  </tr>
  <tr>
    <td class="table_trbg02">Scripting.FileSystemObject(FSO �ı��ļ���д)��<%If IsObjInstalled("Scripting.FileSystemObject") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
    <td class="table_trbg02">adodb.connection(ADO ���ݶ���)��
    <%If IsObjInstalled("adodb.connection") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
  </tr>
  <tr>
    <td colspan="2" class="table_trbg02"><strong>�������ļ��ϴ��͹������</strong></td>
  </tr>
   <tr>
    <td class="table_trbg02">SoftArtisans.FileUp(SA-FileUp �ļ��ϴ�)��<%If IsObjInstalled("SoftArtisans.FileUp") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
    <td class="table_trbg02">SoftArtisans.FileManager(SoftArtisans �ļ�����)��
    <%If IsObjInstalled("SoftArtisans.FileManager") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
  </tr>
   <tr>
    <td class="table_trbg02">LyfUpload.UploadFile(���Ʒ���ļ��ϴ����)��<%If IsObjInstalled("LyfUpload.UploadFile") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
    <td class="table_trbg02">Persits.Upload.1(ASPUpload �ļ��ϴ�)��
    <%If IsObjInstalled("Persits.Upload.1") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
  </tr>
   <tr>
    <td class="table_trbg02">w3.upload(Dimac �ļ��ϴ�)��<%If IsObjInstalled("w3.upload") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
    <td class="table_trbg02">��</td>
  </tr>
  <tr>
    <td colspan="2" class="table_trbg02"><strong>�������շ��ʼ����</strong></td>
  </tr>
  <tr>
    <td class="table_trbg02"><%If IsObjInstalled("Jmain.Message") Then%>
Jmain4.3�������֧�֣�
  <%else%>
Jmain4.2���֧�֣�
<%end if%>
<%If IsObjInstalled("Jmain.Message") or IsObjInstalled("Jmain.SMTPmain") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
    <td class="table_trbg02">CDONTS.Newmain(���� SMTP ����)��<%If IsObjInstalled("CDONTS.Newmain") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
  </tr>
  <tr>
    <td class="table_trbg02">Persits.mainSender(ASPemain ����)��<%If IsObjInstalled("Persits.mainSender") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
    <td class="table_trbg02">SMTPsvg.mainer(ASPmain ����)��
    <%If IsObjInstalled("SMTPsvg.mainer") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
  </tr>
  <tr>
    <td class="table_trbg02">Smtpmain.Smtpmain.1(Smtpmain ����)��<%If IsObjInstalled("Smtpmain.Smtpmain.1") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
    <td class="table_trbg02">IISmain.Iismain.1(IISmain ����)��
    <%If IsObjInstalled("IISmain.Iismain.1") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
  </tr>
  <tr>
    <td class="table_trbg02">DkQmain.Qmain(dkQmain ����)��<%If IsObjInstalled("DkQmain.Qmain") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
    <td class="table_trbg02">Geocel.mainer(Geocel ����)��
    <%If IsObjInstalled("Geocel.mainer") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
  </tr>  
  <tr>
    <td colspan="2" class="table_trbg02"><strong>ͼ�������</strong></td>
  </tr>
  <tr>
    <td class="table_trbg02">Persits.Jpeg(AspJpegͼ��)��
    <%If IsObjInstalled("Persits.Jpeg") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
    <td class="table_trbg02">SoftArtisans.ImageGen(SA ��ͼ���д���)��<%If IsObjInstalled("SoftArtisans.ImageGen") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
  </tr>
  <tr>
    <td class="table_trbg02">W3Image.Image(Dimac ��ͼ���д���)��
    <%If IsObjInstalled("W3Image.Image") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
    <td class="table_trbg02">wsImage.Resize�� <%If IsObjInstalled("wsImage.Resize") Then response.Write("<span class='red'>��</span>") else response.Write("��")%></td>
  </tr>  
</table>

</body>
</html>