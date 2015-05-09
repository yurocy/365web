<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<html>
<head>
<title>网址导航管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>
<body>
<%	  Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from [web_ad] where web_adname='Side' "
rs.Open sql,conn,1,3
if not rs.eof then 
ad_src=rs("web_info")
open=rs("open")
ad_fg=split(ad_src,"$@$@$@$")

rs.Close
end if%>
 <FORM NAME="form1" METHOD="post" ACTION="ad_enter.asp?id=side">
  <table cellpadding="2" cellspacing="1" border="0" width="100%" align="center" class="border">
    <tr class="tdbg">
      <td height="25" colspan="2" align="center" class="topbg">对联广告修改</td>
    </tr>
	    <tr class="tdbg">
      <td height="25" colspan="2" align="center" class="bgcolor">显示<input type="radio" name="open" value="1"<%if open=1 then%> checked="checked"<%end if%>>
        &nbsp;隐藏<input type="radio" name="open" value="0"<%if open=0 then%> checked="checked"<%end if%>> </td>
    </tr>
    <tr class="tr_southidc">
	 <INPUT TYPE="hidden" NAME="content1" VALUE='<%=ad_fg(0)%>'>
      <INPUT TYPE="hidden" NAME="content2" VALUE='<%=ad_fg(1)%>'>
      <td width="50%" height="494" class="bgcolor"><iframe id="eWebEditor1" src="/webedit/ewebeditor.htm?id=content1&style=coolblue" frameborder="0" scrolling="no" width="100%" height="100%"></iframe></td>
      <td width="50%" height="494" class="bgcolor"><iframe id="eWebEditor2" src="/webedit/ewebeditor.htm?id=content2&style=coolblue" frameborder="0" scrolling="no" width="100%" height="100%"></iframe></td>
    </tr>
    <tr class="tr_southidc">
      <td height="23" colspan="2" align="center" valign="middle" class="bgcolor"><table width="100%" border="0" cellpadding="0" cellspacing="0">
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
<%htmlend%>
</BODY></HTML>

