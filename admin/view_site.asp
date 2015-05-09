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
if Request("action")="del" then
conn.execute "delete from [kj] where id in ("&request.form("selectdel")&")"
response.Redirect"view_site.asp"
end if%>

<BODY><%
		 set rs=server.createobject("adodb.recordset")
		 
	           rs.open "select * from [kj] order by id desc",conn,1,1
			
		  
				if err.number<>0 then
				response.write "数据库中暂时无数据"
				end if
				
				If Request.QueryString("Page") = "" or Request.QueryString("Page") <= 0 then	
				Page = 1
				Else
				Page = CINT(Request.QueryString("Page"))
				End If
				rs.pagesize =120
				total  = rs.RecordCount
				rs.absolutepage=page
				mypagesize = rs.pagesize 

%>
<span class="tdbg"><a href="add_site.asp">添加</a></span>
<tr class="tr_southidc"><td height="23"><form name="form1" method="post" action="?action=del">
      <TABLE width=100% height="0" border=0 cellPadding=1 cellSpacing=1 bordercolor="#BBE2FF" bgcolor="#BBE2FF">
        <TBODY>
          <TR bgcolor="#006AB8">
            <TD width="85" height="30" align="center" vAlign=middle 
  bgColor=#E8F5FF class="tr_southidc">选择</TD>
            <TD width="154" align="center" vAlign=middle 
  bgColor=#E8F5FF class="style14"> 期数</TD>
            <TD width="284" align="center" vAlign=middle 
  bgColor=#E8F5FF class="style14">开奖号码</TD>
            <TD width="155" align="center" vAlign=middle 
  bgColor=#E8F5FF class="style14">开奖日期</TD>
            <TD width="88" align="center" vAlign=middle 
  bgColor=#E8F5FF class="tr_southidc">操作</TD>
          </TR>

    

<%    	i=1
	do while ((not rs.eof) and i<=120)%>
	  
	            <TR>
            <TD height="30" align="center" vAlign=middle 
  bgColor=#FFFFFF><input name="selectdel" type="checkbox" id="selectdel" value=<%=rs("id")%>></TD>
            <TD width="154" align="center" vAlign=middle 
  bgColor=#FFFFFF style="word-break:break-all"><%=rs("qs")%>&nbsp;</TD>
            <TD width="284" align="left" vAlign=middle 
  bgColor=#FFFFFF style="word-break:break-all"><%=rs("hm")%></TD>
            <TD width="155" align="left" vAlign=middle 
  bgColor=#FFFFFF style="word-break:break-all"><%=rs("rq")%></TD>
            <TD width="88" align="center" vAlign=middle 
  bgColor=#FFFFFF><a href="edit_site.asp?id=<%=rs("id")%>">修改</a></TD>
          </TR>

      <%
	 i=i+1
	 j=j+1

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
        <TD width="252" align=middle bgcolor="#FFFFFF"><div align="center">
          <INPUT class="button1" type="submit" value="删除所选" name="Submit1">
        </div></TD>
        <TD width="1102" align=middle bgcolor="#FFFFFF">
		<%if page>1 then%>   
<a href='?page=<%=page-1%>'>上页</a>
<%end if%>
<%if page<rs.pagecount   then%>
<a href='?page=<%=page+1%>'>下页</a>
<%end if%>
</TD>
      </TR>
    </TBODY>
</TABLE>
  </form>
</td>
</tr>

</BODY>
</HTML>
