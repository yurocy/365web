<!--#include file="Conn.asp" -->
<table width="100%"  border="0" cellpadding="3" cellspacing="1" class="table" bgcolor="#cccccc">
<%
'favorites.asp?idstr=
search=request.QueryString("idstr")
if len(search)=0 then 
search=0
%>
<tr bgcolor="#F5F5F5" class="tr"><td class="tablebg" height="50" align="center" width="100%">非注册用户,也可以添加自定义收藏网站，但是不能永久保存。请免费注册成本站会员，帐户登陆后可永久保存收藏网址！</td></tr>
<%
end if
set rs=Server.CreateObject("ADODB.Recordset")
rsstr="select * from WebGoodSite_Url where id in ("&search&") order by Is_Recommend desc,id asc"
rs.open rsstr,conn,1,2

total=(rs.recordcount+7-(rs.recordcount mod 7))

if search<>0 then
for i=1 to total
%>

<%if i mod 7 = 1 or i=1 then%><tr bgcolor="#F5F5F5" class="tr">
<%
if rs.eof and i=1 then
%>
<td>注册用户<%=search%></td>
<%
end if

end if%>
<%if not rs.eof then%>
		     <td height="26" align="center" width="13%" onMouseOver="this.className='t2'; title='<%=Rs("Site_Title")%>'" onMouseOut="this.className='tablebg';" class="tablebg">
		     	<div class="menu2" onMouseOver="this.className='menu1'" onMouseOut="this.className='menu2'">
		     		<a  href='<%=Rs("Site_Url")%>' target="_blank" style="font-size:10pt;">
		     			<font style='color:<%=Rs("Font_Color")%>;font-weight:<% if Rs("bold")=1 then response.write "bold" else response.write "normal"  end if%>'><%=Rs("Site_Title")%></font></a>
			  <%
			  select case rs("img")
			  				case 1
			  response.write "<img src=""/images/1.gif"" />"
			  			  case 2 
			  response.write "<img src=""/images/2.gif"" />"
			  			  case 3
			  response.write "<img src=""/images/3.gif"" />"
			  			  case 4
			  response.write "<img src=""/images/4.gif"" />"
			  			  case 5
			  response.write "<img src=""/images/5.gif"" />"
			  			  case else
			  response.write ""
			  end select
			  %>
        <ul style="left:-20px; top:15px">
		    <% sql="select * from WebGoodSite_Url_lite where Class_ID="&rs("id")&" order by id asc"
	      			set RsLite=server.createobject("adodb.recordset")
	    			 RsLite.open sql,conn,1,1
		 			
					    if not rslite.eof then%>


                  <%do while not RsLite.eof%>
				  <li><a href="<%=Rslite("Site_Url")%>" title=<%=Rslite("Site_Url")%> target="_blank"><%=Rslite("Site_Title")%></a></li>
                  <%Rslite.movenext
              		loop
              Rslite.close
	     	set Rslite=nothing%>
              <%end if%>
          <li><a href="#" onclick="removeFavorites(<%=Rs("ID")%>,'<%=Rs("Site_Title")%>')">取消收藏</a></li>
        </ul>
      </div></td>
	 
<%
rs.movenext
else
%>
    <td  onMouseOver="this.className='t2'" onMouseOut="this.className='tablebg';" class="tablebg"  width="13%"></td>
<%end if%>
<%if i mod 7 = 0 then%></tr><%end if%>
<!--<%=i%>-->
  
  <%

  next
  
              Rs.close
	      set Rs=nothing
	  
end if%>
</table>

