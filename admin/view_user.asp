<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/md5.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<html>
<head>
<title>网址导航管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<%
if request("action")="del" then
conn.execute "delete from [$user] where id in ("&request.form("selectdel")&")"
response.Redirect"view_user.asp"
end if%>

<BODY>

<table cellpadding="2" cellspacing="1" border="0" width="95%" align="center" class="table_southidc">
  <tr class="tr_southidc">
    <td height="20">      <%
				Const MaxPerPage=15
   				dim totalPut   
   				dim CurrentPage
   				dim TotalPages
   				dim j
   				dim sql
    				if Not isempty(Request("page")) then
      				currentPage=Cint(Request("page"))
   				else
      				currentPage=1
   				end if 
		 set rs=server.createobject("adodb.recordset")
		 
	           rs.open "select * from [$user] order by id desc",conn,1,1
			
		  
				if err.number<>0 then
				response.write "数据库中暂时无数据"
				end if
				
  				if rs.eof And rs.bof then
       				Response.Write "<p align='center' class='contents'> 暂时没有数据！</p>"
   				else
	  				totalPut=rs.recordcount

      				if currentpage<1 then
          				currentpage=1
      				end if

      				if (currentpage-1)*MaxPerPage>totalput then
	   					if (totalPut mod MaxPerPage)=0 then
	     					currentpage= totalPut \ MaxPerPage
	   					else
	      					currentpage= totalPut \ MaxPerPage + 1
	   					end if
      				end if

       				if currentPage=1 then
            			showContent
            			showpage totalput,MaxPerPage,"view_user.asp"
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"view_user.asp"
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"view_user.asp"
	      				end if
	   				end if
   				end if

   				sub showContent
       			dim i
	   			i=0
				j=1
%>

  <form name="form1" method="post" action="?action=del">
    <TABLE width=100% height="0" border=0 cellPadding=1 cellSpacing=1 bordercolor="#BBE2FF" bgcolor="#BBE2FF">
      <TBODY>
        <TR bgcolor="#006AB8">
          <TD width="125" align="center" vAlign=middle 
  bgColor=#E8F5FF class="tr_southidc">选择</TD>
          <TD width="123" height="30" align="center" vAlign=middle 
  bgColor=#E8F5FF class="tr_southidc"><span class="style13"> ID</span></TD>
          <TD width="217" align="center" vAlign=middle 
  bgColor=#E8F5FF class="tr_southidc">用户名</TD>
<!--          <TD width="28" align="center" vAlign=middle 
  bgColor=#E8F5FF class="tr_southidc"><span class="style14"> 性别</span></TD>-->
          <TD width="346" align="center" vAlign=middle 
  bgColor=#E8F5FF class="tr_southidc"><span class="style14">密码</span></TD>
          <TD width="386" align="center" vAlign=middle 
  bgColor=#E8F5FF class="tr_southidc"><span class="style14"> 注册时间 </span></TD>
          <TD width="174" align="center" vAlign=middle 
  bgColor=#E8F5FF class="tr_southidc">操作</TD>
        </TR>
        <%

i=0
do while not rs.eof
%>
        <TR>
          <TD align="center" vAlign=middle 
  bgColor=#FFFFFF><input name="selectdel" type="checkbox" id="selectdel" value=<%=rs("id")%>></TD>
          <TD height="30" align="center" vAlign=middle 
  bgColor=#FFFFFF><%=rs("id")%>&nbsp;</TD>
          <TD width="217" align="center" vAlign=middle 
  bgColor=#FFFFFF><%=rs("username")%>&nbsp;</TD>
<!--          <TD width="28" align="center" vAlign=middle 
  bgColor=#FFFFFF ><%if rs("sex")=1 then response.write "男" else response.write "女" end if %>
            &nbsp;</TD>-->
          <TD width="346" align="left" vAlign=middle 
  bgColor=#FFFFFF style="word-break:break-all"><%=rs("password")%></TD>
          <TD width="386" align="right" vAlign=middle 
  bgColor=#FFFFFF><%=rs("regtime")%></TD>
          <TD width="174" align="center" vAlign=middle 
  bgColor=#FFFFFF><a href="edit_user.asp?id=<%=rs("id")%>">修改</a></TD>
        </TR>
        <%
	 i=i+1
		rs.movenext
		  loop
		  rs.close
		  set rs=nothing%>
      </TBODY>
    </TABLE>
    <TABLE width="100%" 
border=0 align=center cellPadding=3 cellSpacing=1 bgColor=#6a1c0f class="table_southidc">
        <TBODY>
      <TR>
        <TD width="153" align=middle bgcolor="#FFFFFF"><div align="center">
          <INPUT class="button1" type="submit" value="删除所选" name="Submit1">
        </div></TD>
      </TR>
    </TBODY>

</TABLE>
  </form>
<TABLE width="100%" border=0 cellPadding=5 cellSpacing=1 bgColor=#6a1c0f class="table_southidc">
  <TBODY>
    <TR bgColor=#ffffff>
      <TD width=639 bgColor=#ffffff height=34><%  
				End Sub   
  
				Function showpage(totalnumber,maxperpage,filename)  
  				Dim n
  				
				If totalnumber Mod maxperpage=0 Then  
					n= totalnumber \ maxperpage  
				Else
					n= totalnumber \ maxperpage+1  
				End If
				
				Response.Write "<form method=Post action="&filename&">"  
				Response.Write "<p align='center' class='contents'> "  
				If CurrentPage<2 Then  
					Response.Write "<font class='contents'>首 页 上一页</font> "  
				Else  
					Response.Write "<a href="&filename&"?page=1 class='contents'>首 页</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&" class='contents'>上一页</a> "  
				End If
				
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>下一页 末 页</font>"  
				Else  
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&" class='contents'>"  
					Response.Write "下一页</a> <a href="&filename&"?page="&n&" class='contents'>末 页</a>"  
				End If  
					Response.Write "<font class='contents'> 页次：</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"页</font> "  
					Response.Write "<font class='contents'> 共有"&totalnumber&"个推荐 " 
					Response.Write "<font class='contents'>转到：</font><input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">"  
					Response.Write "&nbsp;<input type='submit'  class='contents' value='跳转' name='cndok'></form>"  
				End Function  
			%>
      </TD>
  <%htmlend%>  
    </TR>
  </TBODY>
</TABLE></td>
  </tr>

</BODY></HTML>
