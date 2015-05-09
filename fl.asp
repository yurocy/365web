<!--#include file="conn.asp" -->
<%
        dim rsLinka
	sqlLink="select * from WebGoodSiteType_Main"
	set rsLinka=Server.CreateObject("Adodb.RecordSet")
	rsLinka.open sqlLink,conn,1,3
        do while not rsLinka.eof
           response.write "document.write(""<OPTION value='"&rsLinka("ID")&"'>"&rsLinka("Type_Name")&"</OPTION>"");"
           rsLinka.movenext
        loop
	rsLinka.close
	set rsLinka=nothing
        %>
