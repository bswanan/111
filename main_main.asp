<!--#include file="inc_CheckLogin.asp"-->
<html>
<head>
<title>管理首页</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/style.css">
</head>
<body>


<br class="table_br" />
<table width="99%" border="0" align="center" cellpadding="5" cellspacing="1" class="tablebk" style="border-collapse: collapse">
  <tr>
    <td colspan="2" class="table_titlebg">服务器参数</td>
  </tr>
  <tr>
    <td width="50%" class="table_trbg02">服务器正在运行的端口：<%=Request.ServerVariables("server_port")%></td>
    <td class="table_trbg02">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr>
    <td class="table_trbg02">服务器名称：<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
    <td class="table_trbg02">服务器IP：<%=Request.ServerVariables("LOCAL_ADDR")%></td>
  </tr>
  <tr>
    <td class="table_trbg02">站点物理路径：<%=Request.ServerVariables("path_translated")%></td>
    <td class="table_trbg02">虚拟路径：<%=Request.ServerVariables("server_name")%></td>
  </tr>
  <tr>
    <td class="table_trbg02">服务器Application数量：<%=Application.Contents.Count%> 个</td>
    <td class="table_trbg02">服务器Session数量：<%=Session.Contents.Count%> 个</td>
  </tr>
  <tr>
    <td class="table_trbg02">服务器当前时间：<%=now()%></td>
    <td class="table_trbg02">脚本连接超时时间：<%=Server.ScriptTimeout%> 秒</td>
  </tr>
  <tr>
    <td colspan="2" class="table_trbg02"><strong>IIS自带的ASP组件</strong></td>
  </tr>
  <tr>
    <td class="table_trbg02">MSWC.AdRotator：<%If IsObjInstalled("MSWC.AdRotator") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
    <td class="table_trbg02">MSWC.BrowserType：
    <%If IsObjInstalled("MSWC.BrowserType") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
  </tr>
   <tr>
    <td class="table_trbg02">MSWC.NextLink：<%If IsObjInstalled("MSWC.NextLink") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
    <td class="table_trbg02">aMSWC.Tools：
    <%If IsObjInstalled("MSWC.Tools") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
  </tr>
   <tr>
    <td class="table_trbg02">MSWC.Status：<%If IsObjInstalled("MSWC.Status") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
    <td class="table_trbg02">MSWC.Counters：
    <%If IsObjInstalled("MSWC.Counters") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
  </tr>
  <tr>
    <td class="table_trbg02">IISSample.ContentRotator：<%If IsObjInstalled("IISSample.ContentRotator") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
    <td class="table_trbg02">IISSample.PageCounter：
    <%If IsObjInstalled("IISSample.PageCounter") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
  </tr>
   <tr>
    <td class="table_trbg02">MSWC.PermissionChecker：<%If IsObjInstalled("MSWC.PermissionChecker") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
    <td class="table_trbg02">Msxml2.FreeThreadedDOMDocument.3.0：
    <%If IsObjInstalled("Msxml2.FreeThreadedDOMDocument.3.0") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
  </tr>
  <tr>
    <td class="table_trbg02">Scripting.FileSystemObject(FSO 文本文件读写)：<%If IsObjInstalled("Scripting.FileSystemObject") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
    <td class="table_trbg02">adodb.connection(ADO 数据对象)：
    <%If IsObjInstalled("adodb.connection") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
  </tr>
  <tr>
    <td colspan="2" class="table_trbg02"><strong>常见的文件上传和管理组件</strong></td>
  </tr>
   <tr>
    <td class="table_trbg02">SoftArtisans.FileUp(SA-FileUp 文件上传)：<%If IsObjInstalled("SoftArtisans.FileUp") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
    <td class="table_trbg02">SoftArtisans.FileManager(SoftArtisans 文件管理)：
    <%If IsObjInstalled("SoftArtisans.FileManager") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
  </tr>
   <tr>
    <td class="table_trbg02">LyfUpload.UploadFile(刘云峰的文件上传组件)：<%If IsObjInstalled("LyfUpload.UploadFile") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
    <td class="table_trbg02">Persits.Upload.1(ASPUpload 文件上传)：
    <%If IsObjInstalled("Persits.Upload.1") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
  </tr>
   <tr>
    <td class="table_trbg02">w3.upload(Dimac 文件上传)：<%If IsObjInstalled("w3.upload") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
    <td class="table_trbg02">　</td>
  </tr>
  <tr>
    <td colspan="2" class="table_trbg02"><strong>常见的收发邮件组件</strong></td>
  </tr>
  <tr>
    <td class="table_trbg02"><%If IsObjInstalled("Jmain.Message") Then%>
Jmain4.3邮箱组件支持：
  <%else%>
Jmain4.2组件支持：
<%end if%>
<%If IsObjInstalled("Jmain.Message") or IsObjInstalled("Jmain.SMTPmain") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
    <td class="table_trbg02">CDONTS.Newmain(虚拟 SMTP 发信)：<%If IsObjInstalled("CDONTS.Newmain") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
  </tr>
  <tr>
    <td class="table_trbg02">Persits.mainSender(ASPemain 发信)：<%If IsObjInstalled("Persits.mainSender") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
    <td class="table_trbg02">SMTPsvg.mainer(ASPmain 发信)：
    <%If IsObjInstalled("SMTPsvg.mainer") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
  </tr>
  <tr>
    <td class="table_trbg02">Smtpmain.Smtpmain.1(Smtpmain 发信)：<%If IsObjInstalled("Smtpmain.Smtpmain.1") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
    <td class="table_trbg02">IISmain.Iismain.1(IISmain 发信)：
    <%If IsObjInstalled("IISmain.Iismain.1") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
  </tr>
  <tr>
    <td class="table_trbg02">DkQmain.Qmain(dkQmain 发信)：<%If IsObjInstalled("DkQmain.Qmain") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
    <td class="table_trbg02">Geocel.mainer(Geocel 发信)：
    <%If IsObjInstalled("Geocel.mainer") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
  </tr>  
  <tr>
    <td colspan="2" class="table_trbg02"><strong>图像处理组件</strong></td>
  </tr>
  <tr>
    <td class="table_trbg02">Persits.Jpeg(AspJpeg图像)：
    <%If IsObjInstalled("Persits.Jpeg") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
    <td class="table_trbg02">SoftArtisans.ImageGen(SA 的图像读写组件)：<%If IsObjInstalled("SoftArtisans.ImageGen") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
  </tr>
  <tr>
    <td class="table_trbg02">W3Image.Image(Dimac 的图像读写组件)：
    <%If IsObjInstalled("W3Image.Image") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
    <td class="table_trbg02">wsImage.Resize： <%If IsObjInstalled("wsImage.Resize") Then response.Write("<span class='red'>√</span>") else response.Write("×")%></td>
  </tr>  
</table>

</body>
</html>