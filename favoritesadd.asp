<% Response.Charset="gb2312"%><!--#include file="Conn.asp" -->

<%

memberuserid=request.Cookies("memberuserid")
sessionidstr=request.Cookies("sessionidstr")
if memberuserid<>"" and sessionidstr<>"" then
	set rs=server.createobject("adodb.recordset")
	sql="select * from [$user] where [id]="&memberuserid&" and [sessionidstr]='"&sessionidstr&"'"
	rs.open sql,conn,1,3
	if not (rs.bof and rs.eof) then
	if isnull(Rs("fav")) then
	Rs("fav")=request("id")&","
	else
	Rs("fav")=Rs("fav")&request("id")&","
	end if
	rs.update
	
	end if
	rs.close
	set rs=nothing
end if

%>