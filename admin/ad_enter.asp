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
<%
select case request("id")
case "side"

				Set rs = Server.CreateObject("ADODB.Recordset")
				sql="select * from [web_ad] where web_adname='Side' "
				rs.Open sql,conn,1,3
				rs("web_info")=Request("content1")&"$@$@$@$"&Request("content2")
				rs("open")=Request("open")
				rs.Update
				rs.Close
				set rs=nothing
				conn.Close
				set conn=nothing
				
				
				

				
				
	response.write"<SCRIPT language=JavaScript>alert('成功修改对联');history.back();</SCRIPT>" 
	response.end



case "top"

				Set rs = Server.CreateObject("ADODB.Recordset")
				sql="select * from [web_ad] where web_adname='top' "
				rs.Open sql,conn,1,3
				rs("web_info")=Request("content1")
				rs.Update
				rs.Close
				set rs=nothing
				conn.Close
				set conn=nothing
					response.write"<SCRIPT language=JavaScript>alert('成功修改');history.back();</SCRIPT>" 
	response.end
	
	
	
case else

				Set rs = Server.CreateObject("ADODB.Recordset")
				sql="select * from [web_ad] where web_adname='"&request("id")&"' "
				rs.Open sql,conn,1,3
				rs("web_info")=Request("content1")
				rs.Update
				rs.Close
				set rs=nothing
				conn.Close
				set conn=nothing
					response.write"<SCRIPT language=JavaScript>alert('成功修改');history.back();</SCRIPT>" 
	response.end
	
	
	
end select


%>


