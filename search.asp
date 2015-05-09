<!--#include file="conn.asp"--><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>搜索结果</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk">
<META http-equiv=Pragma content=no-cache>
<META http-equiv=Cache-Control content=no-cache>
<META http-equiv=Expires content=0>
<META content="MSHTML 6.00.2800.1561" name=GENERATOR>
</HEAD>

<BODY >
<table width="539" height="100" border="0" cellpadding="6" cellspacing="0">
  <tr>
    <td width="527"><table width="144" height="80" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="144"><img src="images/search.gif" width="114" height="25"></td>
        </tr>
    </table>
      <table width="515" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
        <tr>
          <td width="513" align="center" bgcolor="#FFFFFF">
<table width="93%" border="0" cellpadding="3" cellspacing="0">

<%
search=request("keyword")
typeid=int(request("typeid"))

set rs=Server.CreateObject("ADODB.Recordset")
rsstr="select * from WebGoodSite_Url where (Site_Title like '%" & search & "%')"  'Class_ID="&typeid&" and 
rs.open rsstr,conn,1,2
do while not rs.EOF
%>
<tr>
<td align='left'>
<font size=2><a href='<%=replace(Rs("Site_Url"),"测速","")%>' target=_blank ><%=Rs("Site_Title")%></a>&nbsp;&nbsp;</font>
</td>
</tr>
 <%Rs.movenext
              loop
              Rs.close
	      set Rs=nothing%>







          </table></td>
        </tr>
      </table>
      <table width="41%" border="0" cellpadding="8" cellspacing="0">
        <tr>
          <td align="center">
            <input type="button" onClick="window.close()" value="关闭">          </td>
        </tr>
    </table></td>
  </tr>
</table>
</BODY>
</HTML>
