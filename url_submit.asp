<%
option explicit
response.buffer=true
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/config.Asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��վ�ύ-&gt;-<%=webname%></title>
</head>
<body>
<%
	dim SiteUrl,SiteName,SenderEmail,SiteContact,CreateTime,AlexaValue,Comments,chazhang,fl
        SiteName=trim(request("Text1"))
        SiteUrl=trim(request("Text2"))
        SiteContact=trim(request("Text3"))
		chazhang=trim(request("Text4"))
		fl=trim(request("fl"))
if SiteUrl="" or SiteName="" then 
response.write "<script>history.go(-1);</script>"
response.End()
end if
		dim sqlLink,rsLink
		sqlLink="select top 1 * from WebSiteSubmit_URL where SiteUrl='" & SiteUrl & "'"
		set rsLink=Server.CreateObject("Adodb.RecordSet")
		rsLink.open sqlLink,conn,1,3
		if not (rsLink.bof and rsLink.eof) and 1=2 then
%>
       <table border="0" width="100%" id="table1" cellpadding="0" height="106">
	<tr>
		<td height="30">
		<p align="center"><b><font color="#FF0000">�� ʾ</font></b></td>
	</tr>
	<tr>
		<td height="31">
		<p align="center">���ύ����ַ�Ѿ����ڣ��벻Ҫ�ظ��ύ��</td>
	</tr>
	<tr>
		<td>
		<p align="center"><a href="#" onClick="history.back();" target="_self">�����ύ��վ</a></td>
	</tr>
        </table>
<%
		else
			rsLink.Addnew
			rsLink("SiteUrl")=SiteUrl
			rsLink("SiteName")=SiteName
			rsLink("SenderEmail")=SenderEmail
			rsLink("SiteContact")=SiteContact
			rsLink("CreateTime")=CreateTime
			rsLink("AlexaValue")=AlexaValue
			rsLink("Comments")=Comments
			rsLink("chazhang")=chazhang
			rsLink("fl")=fl
			rsLink.update
			rsLink.close
			set rsLink=nothing
			call CloseConn()
%>
<script language="JavaScript" type="text/javascript">
alert("�ύ�ɹ�");
location="<%=request.servervariables("HTTP_REFERER")%>";
</script>

<%			
		end if
%>
</body>
</html>
<%
call CloseConn()
%>
