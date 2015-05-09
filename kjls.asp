<!--#include file="conn.asp"--><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE>历史记录</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312"><LINK 
href="css/ls.css" type=text/css rel=stylesheet><TD align="center" 
nowrap>
<META content="MSHTML 6.00.6000.16809" name=GENERATOR></HEAD>
<BODY>
<TABLE width=650 border=0 align="center" cellPadding=0 cellSpacing=1 bgColor=#cccccc>
  <TBODY>
  <TR style="TEXT-ALIGN: center">
    <TD style="FONT-SIZE: 15px" bgColor=#ffffff colSpan=15 
      height=20>历史数据 <%
	  
	  y=request("y")
	  

	  
	  			set rs=server.createobject("adodb.recordset")
		 
	           rs.open "select distinct year(rq) from kj order by year(rq) desc",conn,1,1
			   	  if y="" then
				  y=rs(0)
				  end if
			   do while not rs.eof%><a href="?y=<%=rs(0)%>"><%=rs(0)%></a> &nbsp;
			   <%rs.movenext
		  		loop
		 		rs.close
		  		set rs=nothing%></TD>
  </TR>
  <TR style="TEXT-ALIGN: center">
    <TD width=60 background="" bgColor=#ffffff height=23>期次</TD>
    <TD width=65 background="" bgColor=#ffffff>开奖日期</TD>
    <TD width=31 background="" bgColor=#ffffff>正1</TD>
    <TD width=31 background="" bgColor=#ffffff>正2</TD>
    <TD width=31 background="" bgColor=#ffffff>正3</TD>
    <TD width=31 background="" bgColor=#ffffff>正4</TD>
    <TD width=31 background="" bgColor=#ffffff>正5</TD>
    <TD width=31 background="" bgColor=#ffffff>正6</TD>
    <TD width=35 background="" bgColor=#ffffff>特码</TD>
    <TD width=37 background="" bgColor=#ffffff>总和</TD>
    <TD width=42 background="" bgColor=#ffffff>特单双</TD>
    <TD width=42 background="" bgColor=#ffffff>特大小</TD>
    <TD width=42 background="" bgColor=#ffffff>合单双</TD>
    <TD width=43 background="" bgColor=#ffffff>总单双</TD>
    <TD width=42 background="" bgColor=#ffffff>总大小</TD></TR>
<%

				set rs=server.createobject("adodb.recordset")
		 
	           rs.open "select * from [kj] where year(rq)="&y&" order by rq desc",conn,1,1
			   
			   do while not rs.eof
				qs=rs("qs")
				hm=rs("hm")
				haoma=split(hm,",")
				xq=rs("xq")
				rq=rs("rq")
				folder=rs("folder")


%>
  <TR style="TEXT-ALIGN: center" bgColor=#ffffff>
    <TD height=31><%=rs("qs")%></TD>
    <TD><%=rs("rq")%></TD>
           <TD> <img src="<%=folder%>/<%=int(haoma(0))%>.gif" width="35" height="20" /></TD>
            <TD> <img src="<%=folder%>/<%=int(haoma(1))%>.gif" width="35" height="20" /></TD>
            <TD>  <img src="<%=folder%>/<%=int(haoma(2))%>.gif" width="35" height="20" /></TD>
      <TD>  <img src="<%=folder%>/<%=int(haoma(3))%>.gif" width="35" height="20" /></TD>
         <TD>   <img src="<%=folder%>/<%=int(haoma(4))%>.gif" width="35" height="20" /></TD>    <TD><img src="<%=folder%>/<%=int(haoma(5))%>.gif" width="35" height="20" /></TD>
	<TD>    <img src="<%=folder%>/<%=int(haoma(6))%>.gif" width="35" height="20" /></TD>
    <TD><%=(int(haoma(0))+int(haoma(1))+int(haoma(2))+int(haoma(3))+int(haoma(4))+int(haoma(5))+int(haoma(6)))%></TD>
    <TD><%
   tmp_hao6=int(haoma(6))
  if tmp_hao6=49 then
   response.write "和"
   else
	ds=tmp_hao6 mod 2
	select case ds
	case 1
	response.write "单"
	case 0
	response.write "双"
	end select
	
	end if
	%></TD>
    <TD><%
  
  if tmp_hao6=49 then
   response.write "和"
   else
  
  
  if tmp_hao6 >= (49/2) then
	response.write "大"
	else
	response.write "小"
	end if
	
	end if
	%></TD>
    <TD><%
he=int(haoma(6))
dim tmp_he
tmp_he=0
  if tmp_hao6=49 then
   response.write "和"
   else
  for kk=1 to len(he)

	tmp_he=tmp_he+int(left(he,kk) mod 10)
	'response.write kk&"*"&int(left(he,kk) mod 10)&"*"&tmp_he&"~"
	next
	select case tmp_he mod 2
	case 1
	response.write "单"'"#"&tmp_he & "单" & he
	case 0
	response.write "双"'"#"&tmp_he & "双" & he
	end select
	end if
	
	%>
</TD>
    <TD><%zong=int(haoma(0))+int(haoma(1))+int(haoma(2))+int(haoma(3))+int(haoma(4))+int(haoma(5))+int(haoma(6))
	select case zong mod 2
	case 1
	response.write "单"
	case 0
	response.write "双"
	end select%></TD>
    <TD><%if zong>=49*7/2 then response.write "大" else response.write "小" end if %></TD></TR>
				<%rs.movenext
		  loop
		  rs.close
		  set rs=nothing%>
</TBODY></TABLE></BODY></HTML>
