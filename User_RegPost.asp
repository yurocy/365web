<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>会员注册</title>
<!--#include file="conn.asp" -->
<%
UserName=replace(trim(request("username")),"'","")
Password=replace(trim(Request("password")),"'","")

PwdConfirm=request.Form("repassword")

sex=request.Form("sex")

mobile=request.Form("telephone")

result="由于以下的原因不能注册用户！"
if len(UserName)=0 then error="请输入用户名"
if Password="" then  error=error&"请输入密码"
if Password<>PwdConfirm then error=error&"两次输入密码不一致"
if error="" then

		set rs=server.CreateObject("adodb.recordset")
		sql="select * from [$user] where username='"&username&"'"
		rs.open sql,conn,1,1
			if rs.recordcount<1 then
				set rs1=server.createobject("adodb.recordset")
				sql="select * from [$user]" 
				rs1.open sql,conn,1,3
				rs1.addnew
				rs1("UserName")=UserName
				rs1("Password")=Password
				rs1.update
				rs1.addnew


			
			error=error&"注册成功!请牢记您的用户名密码!"
			result="恭喜你!"
			
			
			else
			
			error=error&"用户已经存在"
			
			end if
		
end if
%>

<script language="javascript">
alert("<%=error%>");
window.close();
</script>
</body>
</html>
