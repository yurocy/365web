<% Response.Charset="gb2312"%>
<!--#include file="Conn.asp" -->

<%
	
	dim RsMore
	sql="select * from WebGoodSite_Url order by id asc"
	set RsMore=server.createobject("adodb.recordset")
	RsMore.open sql,conn,1,3
	
	do while not RsMore.eof

	RsMore("firstword")=getpychar(BigToGb(left(RsMore("Site_Title"),1)))
	'response.write getpychar(BigToGb(left(RsMore("Site_Title"),1)))
	'response.write BigToGb(left(RsMore("Site_Title"),1))
	
	RsMore.Update
	RsMore.movenext
	loop

%>
<%

%> 
