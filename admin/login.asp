<%@language=vbscript codepage=936 %>
<%
option explicit
'Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
'Response.CacheControl = "no-cache" 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>助赢建站  后台登陆</title>
<!--0-->
<meta name="MSSmartTagsPreventParsing" content="TRUE" />
<meta http-equiv="MSThemeCompatible" content="Yes" />

<script language=javascript>
<!--
function SetFocus()
{
if (document.Login.username.value=="")
	document.Login.username.focus();
else
	document.Login.username.select();
}
function CheckForm()
{
	if(document.Login.username.value=="")
	{
		alert("请输入用户名！");
		document.Login.username.focus();
		return false;
	}
	if(document.Login.password.value == "")
	{
		alert("请输入密码！");
		document.Login.password.focus();
		return false;
	}
	if (document.Login.CheckCode.value==""){
       alert ("请输入您的验证码！");
       document.Login.CheckCode.focus();
       return(false);
    }
}
function CheckBrowser() 
{
  var app=navigator.appName;
  var verStr=navigator.appVersion;
  if (app.indexOf('Netscape') != -1) {
    alert("定做王提示：\n    你使用的是Netscape浏览器，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。");
  } 
  else if (app.indexOf('Microsoft') != -1) {
    if (verStr.indexOf("MSIE 3.0")!=-1 || verStr.indexOf("MSIE 4.0") != -1 || verStr.indexOf("MSIE 5.0") != -1 || verStr.indexOf("MSIE 5.1") != -1)
      alert("定做王提示：\n    您的浏览器版本太低，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。");
  }
}
//-->
</script>
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<style type="text/css">
body { background:#C3E8F8;}
td { font-size:12px;}
</style>
</head>
<body bgcolor="#C3E8F8" topmargin="0">
<div align="center">
<table border="0" cellpadding="0" cellspacing="0" width="0" background="images/login_back.gif" height="36">
<form name="Login" action="Admin_ChkLogin.asp" method="post" target="_parent" onSubmit="return CheckForm();">

  <tr>
    <td width="990" height="228" colspan="2"></td>
  </tr>
  <tr>
    <td width="423" height="36">
      <p align="right"><b><font color="#FFFFFF">用户名:</font></b></p>
    </td>
    <td width="570" height="36"><input type="text" class="Input" name="username" size="20"></td>
  </tr>
  <tr>
    <td width="423" height="34">
      <p align="right"><b><font color="#FFFFFF">密　码:</font></b></td>
    <td width="570" height="34"><input type="password" class="Input" name="password" size="22"></td>
  </tr>
  <tr>
    <td width="423" height="34">
      <p align="right"><b><font color="#FFFFFF">验证码:</font></b></td>
    <td width="570" height="34"><input type="text" class="Input" name="CheckCode" size="6" maxlength="4">&nbsp;&nbsp;<img src="../getcode.asp" alt="验证码,看不清楚?请点击刷新验证码" style="cursor:pointer; vertical-align:middle;height:18px;" onClick="this.src='../getcode.asp?t='+Math.random()"/></td>
  </tr>
  <tr>
    <td width="423" height="31"></td>
    <td width="570" height="31"><input width="50" class="button" type="submit" name="submit" value="登 录"/></td>
  </tr>
  <tr>
    <td width="423" height="106">  </td>
    <td width="570" height="106"><a href="/index.asp"><font color="#65A2B8"><b>返回首页</b></font></a></td>
  </tr>
  </form>
</table>
</div>
</body>

</html>