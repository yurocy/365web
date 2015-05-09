<% Response.Charset="gb2312"%>
<!--#include file="Conn.asp" -->

<%
'list.asp?id=6&char=D

ID=int(request("id"))
charA=left(request("char"),3)




	sql="select * from WebGoodSiteType_Main where ID="&ID
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if not (rs.eof and rs.bof) then
	Type_Name=rs("Type_Name")
	Type_Color=rs("Type_Color")
	Bg_Color=rs("Bg_Color")
	Bg_img=rs("Bg_img")
	end if



    sql="select * from WebGoodSiteType_Class where Main_ID="&ID
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
        if rs.eof and rs.bof then
	else
        i=1
        do while not rs.eof
		
			dim RsMore
			if len(charA)=3 then
			sql="select * from WebGoodSite_Url where Class_ID="&rs("id")&" order by Is_Recommend desc,id asc"
			else
			sql="select * from WebGoodSite_Url where Class_ID="&rs("id")&" and firstWord='"&charA&"' order by Is_Recommend desc,id asc"
			end if
			set RsMore=server.createobject("adodb.recordset")
			RsMore.open sql,conn,1,1
			total=RsMore.recordcount
			
			if total>0 then
			
			
			
		%>
<table width="845" border="0" align="center" cellpadding="0" cellspacing="0" id="ch<%=rs("id")%>">
  <tr>
    <td width="30" valign="top">
	<table width="30" height="59"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td background="images/leftbar.gif" style="cursor:hand;" onClick="coloseChl(ch<%=rs("id")%>)"><div align="center" class="leftbar"><%=rs("Type_Class")%></div></td>
        </tr>
      </table>
	  </td>
    <td width="815">
	<table width="100%"  border="0" cellpadding="3" cellspacing="1" class="table" bgcolor="<%=Bg_Color%>">
          <%
		 	 tmp_num=1
            'do while not RsMore.eof
			if total mod 7 <>0 then
			total=(total+7-(total mod 7))
      end if
			for tmp_num=1 to total
			  
			  
			  'response.write "<!--"&RsMore("Site_Title")&"-->"

	       %>
		   <%if tmp_num mod 7 = 1 or tmp_num=1 then%><tr bgcolor="<%=Type_Color%>" class="tr"><%end if%>
		   <%if not RsMore.eof then%>
		     <td height="26" align="center" width="13%" onMouseOver="this.className='t2'; title='<%=RsMore("Site_Title")%>'" onMouseOut="this.className='tablebg';" class="tablebg"><div class="menu2" onMouseOver="this.className='menu1'" onMouseOut="this.className='menu2'"><a <%if RsMore("Site_Url")="测速" then %>onclick="openFlyBarS(<%=rsmore("id")%>)" <%else%> href="<%=RsMore("Site_Url")%>" target=_blank <%end if%> title="<%=RsMore("Site_Title")%>" style="font-size:10pt;"><font style='color:<%=RsMore("Font_Color")%>;font-weight:<% if RsMore("bold")=1 then response.write "bold" else response.write "normal"  end if%>'><%=RsMore("Site_Title")%></font></a>
              <%
			  select case rsmore("img")
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

		   <ul style="left:23px; top:15px">
              <% sql="select * from WebGoodSite_Url_lite where Class_ID="&rsmore("id")&" order by id asc"
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
			 <li><a href="#" onclick="addFavorites(<%=rsmore("id")%>,'<%=RsMore("Site_Title")%>')">加入收藏</a></li>
              </ul>
            </div></td>
			
			<%
RsMore.movenext
else
%>
    <td  onMouseOver="this.className='t2'" onMouseOut="this.className='tablebg';" class="tablebg"  width="13%"></td>
<%end if%>

<%if tmp_num mod 7 = 0 then%></tr><%end if%>
       <%
		next
'		RsMore.movenext
'              loop
'              RsMore.close
'	      set RsMore=nothing
		  
		  %>

        </table></td>
    </tr>
  </table>
  <%end if%>
<table width="845" height="20" style='padding-top:8px;' border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width='845' height='8' border='0' align='center' cellpadding='0' cellspacing='0'>
        <tr>
          <td width='50%'><div class='ad'>
              <div id='tad2'></div>
            </div></td>
          <td width='50%'><div class='ad'>
            <div id='tad3'></div></td>
        </tr>
      </table></td>
  </tr>
</table>
<%
		  	  rs.movenext
        loop
        end if
        rs.close
	set rs=nothing
		  %>
