<%'微网网络（www.vwen.com）版权所有 ASP技术交流QQ群:19535106
Dim Btn_First,Btn_Prev,Btn_Next,Btn_Last
Btn_First="<font face=""webdings""><img src='"& MainPath &"images/Page_First.gif' border='0' /></font>"  '定义第一页按钮显示样式
Btn_Prev="<font face=""webdings""><img src='"& MainPath &"images/Page_Previous.gif' border='0' /></font>"  '定义前一页按钮显示样式
Btn_Next="<font face=""webdings""><img src='"& MainPath &"images/Page_Next.gif' border='0' /></font>"  '定义下一页按钮显示样式
Btn_Last="<font face=""webdings""><img src='"& MainPath &"images/Page_Last.gif' border='0' /></font>"  '定义最后一页按钮显示样式
Const XD_Align="Center" '定义分页信息对齐方式
Const XD_Width="100%" '定义分页信息框大小

Class XdownPage
Private XD_PageCount,XD_Conn,XD_Rs,XD_SQL,XD_PageSize,Str_errors,int_curpage,str_URL,int_totalPage,int_totalRecord,XD_sURL

'=================================================================
'PageSize 属性
'设置每一页的分页大小
'=================================================================
Public Property Let PageSize(int_PageSize)
  If IsNumeric(Int_PageSize) Then
    XD_PageSize=CLng(int_PageSize)
  Else
    str_error=str_error & "PageSize的参数不正确"
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
'GetRS 属性
'返回分页后的记录集
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
'GetConn 得到数据库连接
'================================================================ 
'Public Property Let GetConn(obj_Conn)
'  Set XD_Conn=obj_Conn
'End Property

'================================================================
'GetSQL 得到查询语句
'================================================================
Public Property Let GetSQL(str_sql)
  XD_SQL=str_sql
End Property

'==================================================================
'Class_Initialize 类的初始化
'初始化当前页的值
'================================================================== 
Private Sub Class_Initialize
  If Not IsObject(Conn) Then ConnectionDatabase
  Set XD_Conn=Conn
'========================
'设定分页的默认值
'========================
  XD_PageSize=10  '设定分页的默认值为10
'========================
'获取当前面的值
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
'获取当前所有纪录数
'========================
Public function ShowTotalRecord()
int_totalRecord=XD_RS.RecordCount
ShowTotalRecord=int_totalrecord
End function 

'====================================================================
'ShowPage 创建分页导航条
'有首页、前一页、下一页、末页、还有数字导航
'====================================================================
Public Sub ShowPage()
  Dim str_tmp
  XD_sURL = GetUrl()
  int_totalRecord=XD_RS.RecordCount
  If int_totalRecord<=0 Then
    str_error=str_error & "总记录数为零，请输入数据"
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
'显示分页信息，各个模块根据自己要求更改显求位置
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
'int_curpage当前页码，int_totalRecord纪录总数，int_TotalPage 总页数
'====================================================================
'ShowFirstPrv 显示首页、前5页
'====================================================================

Private Function ShowFirstPrv()
  Dim Str_tmp,int_prvpage
    str_tmp = "分页："
  If int_curpage = 1 Then
    str_tmp = str_tmp & Btn_First & " "
  Else
  	str_tmp = str_tmp & "<a href=""" & XD_sURL & "1" & """ title=""首页"">" & Btn_First&"</a> "
    if int_curpage>5 then
    str_tmp = str_tmp & "<a href=""" & XD_sURL & CStr(int_curpage-5) & """ title=""上五页"">" & Btn_Prev & "</a> "
	end if
  End If
  ShowFirstPrv = str_tmp
End Function

'====================================================================
'ShowNextLast 下6页、末页
'====================================================================
Private Function ShowNextLast()
  Dim str_tmp,int_Nextpage
  If Int_curpage >= int_totalpage Then
    str_tmp = Btn_Last &" "
  Else
  	
  	If (int_TotalPage-int_curpage)>5  Then
    str_tmp = "<a href=""" & XD_sURL & CStr(int_curpage+5) & """ title=""下五页"">" & Btn_Next & "</a> "
	end if
	str_tmp = str_tmp & "<a href="""& XD_sURL & CStr(int_totalpage) & """ title=""尾页"">" & Btn_Last & "</a>"
  End If
  ShowNextLast = str_tmp
End Function

'====================================================================
'ShowNumBtn 数字导航
'====================================================================
Private Function showNumBtn()
  Dim i,str_tmp
  PageNMin=int_curpage-4
  if PageNMin<1 then PageNMin=1' 计算最小页数
  PageNMax=int_curpage+4
  if PageNMax>int_TotalPage then PageNMax=int_TotalPage' 计算最大页数
  For i=PageNMin To PageNMax
  if i=int_curpage then
  		str_tmp = str_tmp &"<span style='color:#F60'><b>"& i &"</b></span> "
	else
		str_tmp = str_tmp &"<b><a title='转到第"& i &"页' href='" & XD_sURL & CStr(i) &"'>"& i &"</a></b> "
  	end if
  Next
  showNumBtn = str_tmp
End Function

'====================================================================
'ShowNextBtn 上页 下页
'====================================================================
Private Function ShowNextBtn()
  Dim i,str_tmp
  previous=int_curpage-1 '当前页-1 ‘int_TotalPage总页数
  if previous<1 then previous=1
  NextPage=int_curpage+1
  if NextPage> int_TotalPage then NextPage=int_TotalPage '如果下一页大于总页数
  response.Write("&nbsp;")
  if int_TotalPage<>1 then
  	if previous<>1 and int_curpage=int_TotalPage then
  		str_tmp = str_tmp &" <a title='转到第"& previous &"页' href='" & XD_sURL & CStr(previous) &"'>[上页]</a> [下页] "
	elseif int_curpage=1 and NextPage<>int_TotalPage then
		str_tmp = str_tmp &" [上页] <a title='转到第"& NextPage &"页' href='" & XD_sURL & CStr(NextPage) &"'>[下页]</a> "
	else
		str_tmp = str_tmp &" <a title='转到第"& previous &"页' href='" & XD_sURL & CStr(previous) &"'>[上页]</a> <a title='转到第"& NextPage &"页' href='" & XD_sURL & CStr(NextPage) &"'>[下页]</a> "
	end if
  end if

  ShowNextBtn = str_tmp
End Function

'====================================================================
'ShowGotoBtn 直接跳转页数
'====================================================================
Private Function ShowGotoBtn()
  Dim i,str_tmp
  		str_tmp = str_tmp &"&nbsp; 跳到<input name=""goto"" type=""text"" size=""2"" maxlength=""10"" style=""width:26px; border:1px solid #999999"" onkeydown=""javascript:if(event.keyCode==13)window.location='"& XD_sURL &"'+this.value;"" />页 "
  ShowGotoBtn = str_tmp
End Function

'====================================================================
'ShowPageInfo 分页信息
'更据要求自行修改
'====================================================================
Private Function ShowPageInfo()
  Dim str_tmp
  str_tmp = " &nbsp; 页次:<strong>" & int_curpage & "</strong>/<strong>" & int_totalpage & "</strong>页&nbsp;共<strong>" & int_totalrecord & "</strong>条记录 <strong>" & XD_PageSize & "</strong>条/每页"
  ShowPageInfo = str_tmp
End Function

'==================================================================
'GetURL 得到当前的URL
'更据URL参数不同，获取不同的结果
'==================================================================
Private Function GetURL()
  Dim strurl,str_url,i,j,search_str,result_url
  search_str = "page="
 
  strurl = Request.ServerVariables("URL")
  Strurl = split(strurl,"/")
  i = UBound(strurl,1)
  str_url = strurl(i)  '得到当前页文件名
 
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
' 设置 Terminate 事件。
'====================================================================
Private Sub Class_Terminate 
  XD_RS.Close
  Set XD_RS = nothing
End Sub

'====================================================================
'ShowError 错误提示
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
'类调用样例
'Set MyPage = New XdownPage  '创建对象
'MyPage.GetConn = conn  '得到数据库连接
'MyPage.GetSQL = "Select * From [数据库表名称] Order By ID Desc"  'sql语句
'MyPage.PageSize = 10  '设置每一页的记录条数据为10条
'Set rs = MyPage.GetRS()  '返回Recordset

'For i=1 To MyPage.PageSize  '显示数据
'  If Not rs.EOF Then 
'    Response.Write(rs("数据字段") & "<br/>")
'    rs.Movenext
'  Else
'     Exit For
'  End If
'Next

'MyPage.ShowPage()  '显示分页信息，这个方法可以，在set rs = mypage.getrs()以后,可在任意位置调用，可以调用多次
%>