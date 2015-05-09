<%
Option Explicit
response.buffer=True
Const PurviewLevel=2
Const CheckChannelID=0
Const PurviewLevel_Others="Area"
%>
<!--#include file="../inc/md5.asp"-->
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/config.Asp"-->
<%
dim rs
%>
<html>
<head>
<title>专题管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<%
if request("Success")="True" then
 response.write "<script language='JavaScript'>alert('网站配置保存成功！');</script>"
end if
%>
<style>
.s01{border:1px solid #80c0ff;width:700px;}
</style>

<form method="POST" action="Admin_ConfigSave.asp">
<table border="0" align="center" class="border">
<tr class="title">
  <td height="22" colspan="2" align="center"><b>基本信息</b></td>
</tr>
<tr class="tdbg"> 
<td width="200" height="25" align="right">网站名称：</td>
<td width="1483"><input name="WebName" class="s01" type="text" id="WebName" value="<%=WebName%>"> 
</td>
</tr>
<tr class="tdbg"> 
<td height="25" align="right">网站标题：</td>
<td><input name="WebTitle" class="s01" type="text" id="WebTitle" value="<%=WebTitle%>"> 
</td>
<tr class="tdbg"> 
<td height="25" align="right">网址：</td>
<td><input name="WebUrl" class="s01" type="text" id="WebUrl" value="<%=WebUrl%>"> 
</td>
</tr>
<tr class="tdbg"> 
<td height="25" align="right">网站LOGO：</td>
<td><input name="WebLogo" class="s01" type="text" id="WebLogo" value="<%=WebLogo%>"> 
</td>
</tr>
<tr class="tdbg"> 
<td height="25" align="right">服务QQ：</td>
<td><input name="WebQQ" class="s01" type="text" id="WebQQ" value="<%=WebQQ%>"> 
</td>
</tr><tr class="tdbg"> 
<td height="25" align="right">联系电话：</td>
<td><input name="Webmobi" class="s01" type="text" id="Webmobi" value="<%=Webmobi%>"> 
</td>
</tr>
<tr class="tdbg"> 
<td height="25" align="right">关键字：</td>
<td><input name="keywords" class="s01" type="text" id="keywords" value="<%=keywords%>"> 
</td>
</tr>
<tr class="tdbg"> 
<td height="25" align="right">描述：</td>
<td><input name="description" class="s01" type="text" id="description" value="<%=description%>"> 
</td>
</tr>
<tr class="tdbg"> 
<td height="25" align="right">简介：</td>
<td><input name="COPYRIGHT" class="s01" type="text" id="COPYRIGHT" value="<%=COPYRIGHT%>"> 
</td>
</tr>
<tr align="center" class="tdbg">
<td colspan="3"><input name="Action" type="hidden" id="Action" value="Save">
<input name="cmdSave" type="submit" class="button" id="cmdSave" value=" 保存设置 "></td>
</tr>
</table>
<br>
</form>
</body>
</html>