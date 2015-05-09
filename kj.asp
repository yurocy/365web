<%response.ContentType="text/xml"%><!--#include file="conn.asp"--><!--#include file="inc/kj.asp"--><?xml version="1.0" encoding="utf-8"?>
<root>
  <R M="<%
		 set rs=server.createobject("adodb.recordset")
		 
	           rs.open "select top 1 * from [kj] order by rq desc",conn,1,1
			   qs=rs("qs")
			   hm=rs("hm")
			   hm=replace(hm,",","^")
			   xq=rs("xq")
			   rq=rs("rq")
			   xqqs=rs("xqqs")
			   folder=rs("folder")

if hm="^^^^^^" and status<>1 and status<>2 and status<>3 then status=0

haoma=split(hm,"^")

if haoma(6)<>"" then status=4

if haoma(0)<>"" and haoma(6)="" then status=1

	      '期数^下期日期^下期时间^下期星期^号码x7^状态码^下期期数^年份

response.write qs & "^" & xq &"^^^" & hm & "^" & status & "^" &xqqs & "^" & folder
%>" /> 
  </root>