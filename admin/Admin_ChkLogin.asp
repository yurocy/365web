<%@language=vbscript codepage=936 %>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../inc/md5.asp"-->

<%
dim sql,rs
dim username,password,CheckCode
username=replace(trim(request("username")),"'","")
password=replace(trim(Request("password")),"'","")
CheckCode=replace(trim(Request("CheckCode")),"'","")
if UserName="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�û�������Ϊ�գ�</li>"
end if
if Password="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>���벻��Ϊ�գ�</li>"
end if
if CheckCode="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>��֤�벻��Ϊ�գ�</li>"
end if
if Session("GetCode")="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>���¼ʱ������������·��ص�¼ҳ����е�¼��</li>"
end if
if CheckCode<>CStr(Session("GetCode")) then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�������ȷ�����ϵͳ�����Ĳ�һ�£����������롣</li>"
end if
if FoundErr<>True then
	password=md5(password)
	set rs=server.createobject("adodb.recordset")
	sql="select * from admin where password='"&password&"' and username='"&username&"'"
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�û�����������󣡣���</li>"
	else
		if password<>rs("password") then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>�û�����������󣡣���</li>"
		else
			rs("LastLoginIP")=Request.ServerVariables("REMOTE_ADDR")
			rs("LastLoginTime")=now()
			rs("LoginTimes")=rs("LoginTimes")+1
			rs.update
			session.Timeout=1440
			session("AdminName")=rs("username")
			rs.close
			set rs=nothing
			call CloseConn()
			Response.Redirect "Index.asp"
		end if
	end if
	rs.close
	set rs=nothing
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()

'****************************************************
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'****************************************************
sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td height='22' class='title'><strong>������Ϣ</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr><td height='100' class='tdbg' valign='top'><b>��������Ŀ���ԭ��</b><br>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td class='tdbg'><a href='Login.asp'>&lt;&lt; ���ص�¼ҳ��</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub
%>