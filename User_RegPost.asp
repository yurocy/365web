<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Աע��</title>
<!--#include file="conn.asp" -->
<%
UserName=replace(trim(request("username")),"'","")
Password=replace(trim(Request("password")),"'","")

PwdConfirm=request.Form("repassword")

sex=request.Form("sex")

mobile=request.Form("telephone")

result="�������µ�ԭ����ע���û���"
if len(UserName)=0 then error="�������û���"
if Password="" then  error=error&"����������"
if Password<>PwdConfirm then error=error&"�����������벻һ��"
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


			
			error=error&"ע��ɹ�!���μ������û�������!"
			result="��ϲ��!"
			
			
			else
			
			error=error&"�û��Ѿ�����"
			
			end if
		
end if
%>

<script language="javascript">
alert("<%=error%>");
window.close();
</script>
</body>
</html>
