<!--#include file="conn.asp" -->
<%
UserName=replace(trim(request("username")),"'","")
Password=replace(trim(Request("password")),"'","")

if UserName="" then
	FoundErr=True
	ErrMsg=ErrMsg & "�û�������Ϊ�գ�\n"
	result="history.back(-1);"
end if
if Password="" then
	FoundErr=True
	ErrMsg=ErrMsg & "���벻��Ϊ�գ�\n"
	result="history.back(-1);"
end if


if FoundErr<>True then
	set rs=server.createobject("adodb.recordset")
	sql="select * from [$user] where [password]='"&password&"' and [username]='"&username&"'"
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		
		ErrMsg=ErrMsg & "�û������������\n"
		result="history.back(-1);"
	else
		if password<>rs("password") then
			FoundErr=True
			ErrMsg=ErrMsg & "�û������������\n"
			result="history.back(-1);"
		else
			if rs("pass")=1 then
			session.Timeout=60*24
			session("username")=rs("username")
			session("isVIP")=rs("username")
			ErrMsg=ErrMsg & "��½�ɹ���\n"
			result="location.href='vip.asp';"
			else
			ErrMsg=ErrMsg & "��½ʧ�ܣ��û�δͨ�����\n"
			result="location.href='index.asp';"
			end if
			
		end if
	end if
	rs.close
	set rs=nothing
end if


%>

<html>
<head>
<title></title>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
</head>
<body>
<script>alert('<%=ErrMsg%>');<%=result%></script>
</body>
</html>
