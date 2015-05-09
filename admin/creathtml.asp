<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
</head>

<body><!--#include file="../Conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../pubinc/$function.asp"-->
<%=now()%>
<%
url="http://"&request.ServerVariables("HTTP_HOST")&"/i.asp?time="&now()

content=gethttppage(url)
if len(content)>200 then		
				
	
		On Error Resume Next
		JsFileName = Server.MapPath("/index.html")
	set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
		.Open
		.Type = 2
		.Charset = "GB2312"
		.WriteText content
		.SaveToFile JsFileName, 2
		.Close
	End With
	Set objStream = Nothing	
				
				
	response.write"<SCRIPT language=JavaScript>alert('成功生成静态文件');</SCRIPT>" 
else
	response.write"<SCRIPT language=JavaScript>alert('生成失败!');</SCRIPT>" 
end if

%>


</body>

</html>