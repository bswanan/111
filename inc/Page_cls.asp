<%'΢�����磨www.vwen.com����Ȩ���� ASP��������QQȺ:19535106
Dim Btn_First,Btn_Prev,Btn_Next,Btn_Last
Btn_First="<font face=""webdings""><img src='"& MainPath &"images/Page_First.gif' border='0' /></font>"  '�����һҳ��ť��ʾ��ʽ
Btn_Prev="<font face=""webdings""><img src='"& MainPath &"images/Page_Previous.gif' border='0' /></font>"  '����ǰһҳ��ť��ʾ��ʽ
Btn_Next="<font face=""webdings""><img src='"& MainPath &"images/Page_Next.gif' border='0' /></font>"  '������һҳ��ť��ʾ��ʽ
Btn_Last="<font face=""webdings""><img src='"& MainPath &"images/Page_Last.gif' border='0' /></font>"  '�������һҳ��ť��ʾ��ʽ
Const XD_Align="Center" '�����ҳ��Ϣ���뷽ʽ
Const XD_Width="100%" '�����ҳ��Ϣ���С

Class XdownPage
Private XD_PageCount,XD_Conn,XD_Rs,XD_SQL,XD_PageSize,Str_errors,int_curpage,str_URL,int_totalPage,int_totalRecord,XD_sURL

'=================================================================
'PageSize ����
'����ÿһҳ�ķ�ҳ��С
'=================================================================
Public Property Let PageSize(int_PageSize)
  If IsNumeric(Int_PageSize) Then
    XD_PageSize=CLng(int_PageSize)
  Else
    str_error=str_error & "PageSize�Ĳ�������ȷ"
    ShowError()
  End If
End Property

Public Property Get PageSize
  If XD_PageSize="" or (not(IsNumeric(XD_PageSize))) Then
    PageSize=10 
  Else
    PageSize=XD_PageSize
  End If
End Property

'=================================================================
'GetRS ����
'���ط�ҳ��ļ�¼��
'=================================================================
Public Property Get GetRS()
  Set XD_Rs=Server.createobject("adodb.recordset")
  XD_Rs.PageSize=PageSize
  XD_Rs.Open XD_SQL,XD_Conn,1,1
  If not(XD_Rs.eof and XD_RS.BOF) Then
    If int_curpage>XD_RS.PageCount Then
      int_curpage=XD_RS.PageCount
    End If
    XD_Rs.AbsolutePage=int_curpage
  End If
  Set GetRS=XD_RS
End Property

'================================================================
'GetConn �õ����ݿ�����
'================================================================ 
'Public Property Let GetConn(obj_Conn)
'  Set XD_Conn=obj_Conn
'End Property

'================================================================
'GetSQL �õ���ѯ���
'================================================================
Public Property Let GetSQL(str_sql)
  XD_SQL=str_sql
End Property

'==================================================================
'Class_Initialize ��ĳ�ʼ��
'��ʼ����ǰҳ��ֵ
'================================================================== 
Private Sub Class_Initialize
  If Not IsObject(Conn) Then ConnectionDatabase
  Set XD_Conn=Conn
'========================
'�趨��ҳ��Ĭ��ֵ
'========================
  XD_PageSize=10  '�趨��ҳ��Ĭ��ֵΪ10
'========================
'��ȡ��ǰ���ֵ
'========================
  If Request("page") = "" Then
    int_curpage=1
  ElseIf Not(IsNumeric(request("page"))) Then
    int_curpage=1
  ElseIf CInt(Trim(Request("page"))) < 1 Then
    int_curpage=1
  Else
    Int_curpage = CInt(Trim(Request("page")))
  End If
End Sub

'========================
'��ȡ��ǰ���м�¼��
'========================
Public function ShowTotalRecord()
int_totalRecord=XD_RS.RecordCount
ShowTotalRecord=int_totalrecord
End function 

'====================================================================
'ShowPage ������ҳ������
'����ҳ��ǰһҳ����һҳ��ĩҳ���������ֵ���
'====================================================================
Public Sub ShowPage()
  Dim str_tmp
  XD_sURL = GetUrl()
  int_totalRecord=XD_RS.RecordCount
  If int_totalRecord<=0 Then
    str_error=str_error & "�ܼ�¼��Ϊ�㣬����������"
    Call ShowError()
  End If
  If int_totalRecord="" then
    int_TotalPage=1
  Else
    If int_totalRecord mod PageSize =0 Then
      int_TotalPage = int_TotalRecord \ XD_PageSize
    Else
      int_TotalPage = int_TotalRecord \ XD_PageSize+1
    End If
  End If
 
  If Int_curpage>int_Totalpage Then
    int_curpage=int_TotalPage
  End If
'==================================================================
'��ʾ��ҳ��Ϣ������ģ������Լ�Ҫ���������λ��
'==================================================================
  Response.Write("")
  str_tmp = ShowFirstPrv
  Response.Write(str_tmp)
  str_tmp=showNumBtn
  Response.Write(str_tmp)
  str_tmp = ShowNextLast
  Response.Write str_tmp
  str_tmp = ShowNextBtn
  Response.Write str_tmp  
  str_tmp = ShowPageInfo
  Response.Write(str_tmp)
  str_tmp = ShowGotoBtn
  Response.Write(str_tmp)  
  response.write("")
End Sub
'int_curpage��ǰҳ�룬int_totalRecord��¼������int_TotalPage ��ҳ��
'====================================================================
'ShowFirstPrv ��ʾ��ҳ��ǰ5ҳ
'====================================================================

Private Function ShowFirstPrv()
  Dim Str_tmp,int_prvpage
    str_tmp = "��ҳ��"
  If int_curpage = 1 Then
    str_tmp = str_tmp & Btn_First & " "
  Else
  	str_tmp = str_tmp & "<a href=""" & XD_sURL & "1" & """ title=""��ҳ"">" & Btn_First&"</a> "
    if int_curpage>5 then
    str_tmp = str_tmp & "<a href=""" & XD_sURL & CStr(int_curpage-5) & """ title=""����ҳ"">" & Btn_Prev & "</a> "
	end if
  End If
  ShowFirstPrv = str_tmp
End Function

'====================================================================
'ShowNextLast ��6ҳ��ĩҳ
'====================================================================
Private Function ShowNextLast()
  Dim str_tmp,int_Nextpage
  If Int_curpage >= int_totalpage Then
    str_tmp = Btn_Last &" "
  Else
  	
  	If (int_TotalPage-int_curpage)>5  Then
    str_tmp = "<a href=""" & XD_sURL & CStr(int_curpage+5) & """ title=""����ҳ"">" & Btn_Next & "</a> "
	end if
	str_tmp = str_tmp & "<a href="""& XD_sURL & CStr(int_totalpage) & """ title=""βҳ"">" & Btn_Last & "</a>"
  End If
  ShowNextLast = str_tmp
End Function

'====================================================================
'ShowNumBtn ���ֵ���
'====================================================================
Private Function showNumBtn()
  Dim i,str_tmp
  PageNMin=int_curpage-4
  if PageNMin<1 then PageNMin=1' ������Сҳ��
  PageNMax=int_curpage+4
  if PageNMax>int_TotalPage then PageNMax=int_TotalPage' �������ҳ��
  For i=PageNMin To PageNMax
  if i=int_curpage then
  		str_tmp = str_tmp &"<span style='color:#F60'><b>"& i &"</b></span> "
	else
		str_tmp = str_tmp &"<b><a title='ת����"& i &"ҳ' href='" & XD_sURL & CStr(i) &"'>"& i &"</a></b> "
  	end if
  Next
  showNumBtn = str_tmp
End Function

'====================================================================
'ShowNextBtn ��ҳ ��ҳ
'====================================================================
Private Function ShowNextBtn()
  Dim i,str_tmp
  previous=int_curpage-1 '��ǰҳ-1 ��int_TotalPage��ҳ��
  if previous<1 then previous=1
  NextPage=int_curpage+1
  if NextPage> int_TotalPage then NextPage=int_TotalPage '�����һҳ������ҳ��
  response.Write("&nbsp;")
  if int_TotalPage<>1 then
  	if previous<>1 and int_curpage=int_TotalPage then
  		str_tmp = str_tmp &" <a title='ת����"& previous &"ҳ' href='" & XD_sURL & CStr(previous) &"'>[��ҳ]</a> [��ҳ] "
	elseif int_curpage=1 and NextPage<>int_TotalPage then
		str_tmp = str_tmp &" [��ҳ] <a title='ת����"& NextPage &"ҳ' href='" & XD_sURL & CStr(NextPage) &"'>[��ҳ]</a> "
	else
		str_tmp = str_tmp &" <a title='ת����"& previous &"ҳ' href='" & XD_sURL & CStr(previous) &"'>[��ҳ]</a> <a title='ת����"& NextPage &"ҳ' href='" & XD_sURL & CStr(NextPage) &"'>[��ҳ]</a> "
	end if
  end if

  ShowNextBtn = str_tmp
End Function

'====================================================================
'ShowGotoBtn ֱ����תҳ��
'====================================================================
Private Function ShowGotoBtn()
  Dim i,str_tmp
  		str_tmp = str_tmp &"&nbsp; ����<input name=""goto"" type=""text"" size=""2"" maxlength=""10"" style=""width:26px; border:1px solid #999999"" onkeydown=""javascript:if(event.keyCode==13)window.location='"& XD_sURL &"'+this.value;"" />ҳ "
  ShowGotoBtn = str_tmp
End Function

'====================================================================
'ShowPageInfo ��ҳ��Ϣ
'����Ҫ�������޸�
'====================================================================
Private Function ShowPageInfo()
  Dim str_tmp
  str_tmp = " &nbsp; ҳ��:<strong>" & int_curpage & "</strong>/<strong>" & int_totalpage & "</strong>ҳ&nbsp;��<strong>" & int_totalrecord & "</strong>����¼ <strong>" & XD_PageSize & "</strong>��/ÿҳ"
  ShowPageInfo = str_tmp
End Function

'==================================================================
'GetURL �õ���ǰ��URL
'����URL������ͬ����ȡ��ͬ�Ľ��
'==================================================================
Private Function GetURL()
  Dim strurl,str_url,i,j,search_str,result_url
  search_str = "page="
 
  strurl = Request.ServerVariables("URL")
  Strurl = split(strurl,"/")
  i = UBound(strurl,1)
  str_url = strurl(i)  '�õ���ǰҳ�ļ���
 
  str_params=Trim(Request.ServerVariables("QUERY_STRING"))

  If str_params = "" Then
    result_url = str_url & "?page="
  Else
    If InstrRev(str_params,search_str)=0 Then
      result_url = str_url & "?" & str_params & "&page="
    Else
      j = InstrRev(str_params,search_str)-2
      If j=-1 Then
        result_url=str_url & "?page="
      Else
        str_params = Left(str_params,j)
        result_url = str_url & "?" & str_params & "&page="
      End If
    End If
  End If
  GetURL = result_url
End Function

'====================================================================
' ���� Terminate �¼���
'====================================================================
Private Sub Class_Terminate 
  XD_RS.Close
  Set XD_RS = nothing
End Sub

'====================================================================
'ShowError ������ʾ
'====================================================================
Private Sub ShowError()
  If str_Error <> "" Then
    Response.Write("" & str_Error & "")
    Response.End
  End If
End Sub
End Class
%>



<%
'���������
'Set MyPage = New XdownPage  '��������
'MyPage.GetConn = conn  '�õ����ݿ�����
'MyPage.GetSQL = "Select * From [���ݿ������] Order By ID Desc"  'sql���
'MyPage.PageSize = 10  '����ÿһҳ�ļ�¼������Ϊ10��
'Set rs = MyPage.GetRS()  '����Recordset

'For i=1 To MyPage.PageSize  '��ʾ����
'  If Not rs.EOF Then 
'    Response.Write(rs("�����ֶ�") & "<br/>")
'    rs.Movenext
'  Else
'     Exit For
'  End If
'Next

'MyPage.ShowPage()  '��ʾ��ҳ��Ϣ������������ԣ���set rs = mypage.getrs()�Ժ�,��������λ�õ��ã����Ե��ö��
%>