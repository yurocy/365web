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
dim id
id=trim(request("id"))
if trim(request("action"))="save" then

set rs1=server.CreateObject("adodb.recordset")
rs1.open"select * from [$user] where id="&id,conn,1,3
rs1("password")=trim(request("password"))

rs1.update
rs1.close
set rs1=nothing
response.write "<script language=javascript>alert('更改成功！');history.go(-1);</script>"
response.End
end if
%>
<BODY>
<DIV ALIGN="center">

<%set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [$user] where id="&id,conn,1,1
%>

  <form name="form1" method="post" action="?action=save&id=<%=id%>">

    <table cellpadding="2" cellspacing="1" border="0" width="95%" align="center" class="table_southidc">
      <tr class="tr_southidc">
        <td class="back_southidc" colspan="2" height="25" align="center">用户修改</td>
      </tr>
      <tr class="tr_southidc">
        <td width="15%" height="23">用户名</td>
        <td width="80%"><%=rs("username")%></td>
      </tr>
      <tr class="tr_southidc">
        <td height="23">密码</td>
        <td><input name="password" type="text" id="password" size="20" value="<%=rs("password")%>"></td>
      </tr>

      <tr class="tr_southidc">
        <td height="23" colspan="2" align="center" valign="middle"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="234">&nbsp;</td>
            <td width="69"><input type="submit"  name="Submit" value="确定修改"></td>
            <td width="103">&nbsp;</td>
            <td width="499"><input type="reset"  name="Submit2" value="取消重写"></td>
          </tr>
        </table></td>
      </tr>
    </table>
  </form>
  <p>&nbsp;</p>
</DIV>
<%htmlend%>
</BODY>
</HTML>
