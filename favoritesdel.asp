<% Response.Charset="gb2312"%><!--#include file="Conn.asp" -->

<%

memberuserid=request.Cookies("memberuserid")
sessionidstr=request.Cookies("sessionidstr")
if memberuserid<>"" and sessionidstr<>"" then
	set rs=server.createobject("adodb.recordset")
	sql="select * from [$user] where [id]="&memberuserid&" and [sessionidstr]='"&sessionidstr&"'"
	rs.open sql,conn,1,3
	if not (rs.bof and rs.eof) then

		arrya=split(Rs("fav"),",")
		stra=""
		delstr=request("id")
		for i=0 to ubound(arrya)
			if arrya(i)<>delstr then
				if stra<>"" then
					stra=stra & "," & arrya(i)
				else
				stra=arrya(i)
				end if
			end if
		next
	Rs("fav")=stra
	rs.update
	
	end if
	rs.close
	set rs=nothing
end if

%>