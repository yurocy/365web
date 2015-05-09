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
</head>
<BODY>
<%
place=request("id")
if request("action")="save" then


				Set rs = Server.CreateObject("ADODB.Recordset")
				sql="select * from [siteinfo] where place='"&place&"' "
				rs.Open sql,conn,1,3
				rs("content")=Request("content1")
				rs("open")=Request("open")
				rs.Update
				rs.Close
				set rs=nothing
				conn.Close
				set conn=nothing
					response.write"<SCRIPT language=JavaScript>alert('成功修改');history.back();</SCRIPT>" 
	response.end
else	
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from [siteinfo] where place='"&place&"' "
rs.Open sql,conn,1,3
if not rs.eof then 
ad_src=rs("content")
open=rs("open")
rs.Close
else
conn.execute "insert into [siteinfo](place) values('"&place&"') "
open=1
end if%>
 <FORM NAME="form1" METHOD="post" ACTION="editinfo.asp?id=<%=place%>&action=save">
  <table cellpadding="2" cellspacing="1" border="0" width="100%" align="center" class="border">
    <tr class="tr_southidc">
      <td class="topbg" height="25" align="center">网站内容修改-->><%=place%></td>
    </tr>
		    <tr class="tdbg">
      <td height="25" colspan="2" align="center" class="bgcolor">显示<input type="radio" name="open" value="1"<%if open=1 then%> checked="checked"<%end if%>>
        &nbsp;隐藏<input type="radio" name="open" value="0"<%if open=0 then%> checked="checked"<%end if%>> </td>
    </tr>
    <tr class="tr_southidc">
	 <INPUT TYPE="hidden" NAME="content1" VALUE='<%=ad_src%>'>
     
      <td width="100%" height="600">
      <IFRAME ID="eWebEditor1" SRC="/webedit/ewebeditor.htm?id=content1&style=mini" FRAMEBORDER="0" SCROLLING="no" WIDTH="100%" HEIGHT="100%"></IFRAME>
      </td>
    </tr>
    <tr class="tr_southidc">
      <td height="23" align="center" valign="middle"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="278">&nbsp;</td>
            <td width="97"><input type="submit"  name="Submit" value="确定修改"></td>
            <td width="66">&nbsp;</td>
            <td width="215">&nbsp;</td>
          </tr>
      </table></td>
    </tr>
  </table>
</form>
<%end if%>
<%htmlend%>
</BODY></HTML>

