<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2
Const CheckChannelID=0
Const PurviewLevel_Others="Admin_Web_Type_Manage"
%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/md5.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<%
dim strFileName
const MaxPerPage=15
dim totalPut,CurrentPage,TotalPages
dim sql,rs,ID,str_select,str_text
dim Action,FoundErr,ErrMsg,TypeID
Action=trim(request("Action"))
ID=Trim(Request("ID"))
TypeID=CLng(Trim(Request("TypeID")))
str_select=trim(request("select_Query"))
str_text=trim(request("Query_text"))
strFileName="Admin_Web_Type_Manage.asp?select_Query="&str_select&"&Query_text="&str_text

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

if ID<>"" then
	if Action="Check" then
		conn.execute "Update Words set IsOK=1 where ID=" & CLng(ID)
	elseif Action="CancelCheck" then
		conn.execute "Update Words set IsOK=0 Where ID=" & CLng(ID)
	elseif Action="Del" then
		conn.execute "Delete From WebSiteType_Main Where ID=" & CLng(ID)
		conn.execute "Delete From WebSite_Url Where Class_ID=(select ID from WebSiteType_Class where Main_ID=" & CLng(ID)&")"
		conn.execute "Delete From WebSiteType_Class Where Main_ID=" & CLng(ID)
	elseif Action="Del1" then
		conn.execute "Delete From WebSiteType_Class Where ID=" & CLng(ID)
		conn.execute "Delete From WebSite_Url Where Class_ID=" & CLng(ID)
                Action="list2"
	end if
end if
%>
<html>
<head>
<title>��ַ��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<script LANGUAGE="javascript">
function foreColor()
{
  var arr = showModalDialog("editor_selcolor.asp", "", "dialogWidth:18.5em; dialogHeight:17.5em; help: no; scroll: no; status: no");
  if (arr != null)
  {
   document.AddLink.select_color.value=arr;
   document.all("color_img").style.backgroundColor =arr;
  }
}

function Check() {
if (document.AddLink.SiteName.value=="")
	{
	  alert("�������������!")
	  document.AddLink.SiteName.focus()
	  return false
	 }
}
function Check1() {
if (document.AddLink.SelectType.value=="0")
	{
	  alert("��ѡ��һ�����!")
	  document.AddLink.SelectType.focus()
	  return false
	 }
if (document.AddLink.SiteName.value=="")
	{
	  alert("�������������!")
	  document.AddLink.SiteName.focus()
	  return false
	 }
}
function ConfirmDel()
{
   if(confirm("ɾ�������,��ص���վҲ��ɾ��,ȷ��Ҫɾ����?"))
     return true;
   else
     return false;
}
function InitDocument(){
if(typeof(document.all.select_color)=="object")
{
   document.all("color_img").style.backgroundColor=document.all("select_color").value;
}
}
</script>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" onload="InitDocument()">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg"> 
    <td height="22" colspan="2" align="center"><strong>�� ַ �� �� �� ��</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>��������</strong></td>
    <td height="30"> <a href="Admin_Web_Type_Manage.asp?Action=list1">һ��������</a> | <a href="Admin_Web_Type_Manage.asp?Action=Add1">���һ�����</a> | <a href="Admin_Web_Type_Manage.asp?Action=list2">����������</a> | <a href="Admin_Web_Type_Manage.asp?Action=Add2">��Ӷ������</a></td>
  </tr>
</table>
<br>
<%
if Action="Add1" then
	call Add1()
elseif Action="Add2" then
	call Add2()
elseif Action="list2" then
	call main1()
elseif Action="SaveAdd1" then
	call SaveAdd1()
elseif Action="SaveAdd2" then
	call SaveAdd2()
elseif Action="Modify1" then
	call Modify1()
elseif Action="SaveModify1" then
	call SaveModify1()
elseif Action="Modify2" then
	call Modify2()
elseif Action="SaveModify2" then
	call SaveModify2()
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()

sub main()
	sql="select * from WebSiteType_Main"
        if str_select="��ѡ��" or str_select="" then
	sql=sql
        else
	sql=sql&" where ("&str_select&" like '%"&str_text&"%')"
        end if
	sql=sql&" order by id desc"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	
 	if rs.eof and rs.bof then
		response.write "�� 0 ����¼"
	else
    	totalPut=rs.recordcount
	if currentpage<1 then
       		currentpage=1
    	end if
    	if (currentpage-1)*MaxPerPage>totalput then
	   		if (totalPut mod MaxPerPage)=0 then
	     		currentpage= totalPut \ MaxPerPage
		  	else
		      	currentpage= totalPut \ MaxPerPage + 1
	   		end if

    	end if
	    if currentPage=1 then
        	List
        	showpage strFileName,totalput,MaxPerPage,true,true,"����¼"
   	    else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	List
            	showpage strFileName,totalput,MaxPerPage,true,true,"����¼"
        	else
	        	currentPage=1
           		List
           		showpage strFileName,totalput,MaxPerPage,true,true,"����¼"
	    	end if
	     end if
	end if
	rs.close
	set rs=nothing
end sub
sub main1()
	sql="select b.ID,a.Type_Name,b.Type_Class,b.Type_Color from WebSiteType_Main a,WebSiteType_Class b where a.id=b.main_id "
        if str_select="��ѡ��" or str_select="" then
	sql=sql
        else
	sql=sql&" and ("&str_select&" like '%"&str_text&"%')"
        end if
	sql=sql&" order by a.type_order,b.type_order desc"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	
 	if rs.eof and rs.bof then
		response.write "�� 0 ����¼"
	else
    	totalPut=rs.recordcount
	if currentpage<1 then
       		currentpage=1
    	end if
    	if (currentpage-1)*MaxPerPage>totalput then
	   if (totalPut mod MaxPerPage)=0 then
	   	currentpage= totalPut \ MaxPerPage
	   else
	      	currentpage= totalPut \ MaxPerPage + 1
	   end if

    	end if
	    if currentPage=1 then
        	List1
        	showpage strFileName&"&Action=list2",totalput,MaxPerPage,true,true,"����¼"
   	    else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	List1
            	showpage strFileName&"&Action=list2",totalput,MaxPerPage,true,true,"����¼"
        	else
	        	currentPage=1
           		List1
           		showpage strFileName&"&Action=list2",totalput,MaxPerPage,true,true,"����¼"
	    	end if
	     end if
	end if
	rs.close
	set rs=nothing
end sub

sub List
   	dim i
    i=0
%>
<table border="0" cellpadding="0" cellspacing="0">
<form method="post" name="AddLink" action="Admin_Web_Type_Manage.asp">
  <tr>
    <td width="76" align="right"><b>��ѯ:</b>&nbsp;&nbsp;</td>
    <td><select size="1" name="select_Query">
        <option selected value="��ѡ��">��ѡ��</option>
        <option <%if str_select="type_name" then response.write "selected" end if%> value="type_name">���</option>
        </select>&nbsp;<input name="Query_text" size="30" class="Input" maxlength="100" type="text" value="<%=str_text%>">&nbsp;<input class="Button" type="submit" value="��ѯ" name="B3"></td>
  </tr>
  <tr>
    <td colspan="2" height="5"></td>
  </tr>
</form>
</table>
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr class="title">
      <td width="200" height="22" align="center">�������</td>
      <td width="50" height="22" align="center">������ɫ</td>
      <td height="22" align="center">����</td>
    </tr>
<%
do while not rs.eof
%>
    <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td width="200" align="center"><%=rs("Type_Name")%></td>
      <td width="50" align="center"><%=rs("Type_Color")%></td>
      <td align="center"> 
<%
      response.write "<a href='Admin_Web_Type_Manage.asp?Action=Modify1&ID=" & rs("ID") & "'>�޸�</a> | "
      response.write "<a href='Admin_Web_Type_Manage.asp?Action=Del&ID=" & rs("ID") & "' onclick='return ConfirmDel();'>ɾ��</a>"
	  %>

 </td>
    </tr>
   <%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
  </table>
<%
end sub

sub List1
   	dim i
    i=0
%>
<table border="0" cellpadding="0" cellspacing="0">
<form method="post" name="AddLink" action="Admin_Web_Type_Manage.asp">
  <tr>
    <td width="76" align="right"><b>��ѯ:</b>&nbsp;&nbsp;</td>
    <td><select size="1" name="select_Query">
        <option selected value="��ѡ��">��ѡ��</option>
        <option <%if str_select="a.type_name" then response.write "selected" end if%> value="a.type_name">һ�����</option>
        <option <%if str_select="b.type_class" then response.write "selected" end if%> value="b.type_class">�������</option>
        </select>&nbsp;<input name="Query_text" size="30" class="Input" maxlength="100" type="text" value="<%=str_text%>">&nbsp;<input class="Button" type="submit" value="��ѯ" name="B3"></td>
  </tr>
  <tr>
    <td colspan="2" height="5"></td>
  </tr>
</form>
</table>
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr class="title">
      <td width="200" height="22" align="center">һ���������</td>
      <td width="200" height="22" align="center">�����������</td>
      <td width="50" height="22" align="center">������ɫ</td>
      <td height="22" align="center">����</td>
    </tr>
<%
do while not rs.eof
%>
    <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td width="200" align="center"><%=rs("Type_Name")%></td>
      <td width="200" align="center"><%=rs("Type_Class")%></td>
      <td width="50" align="center"><%=rs("Type_Color")%></td>
      <td align="center"> 
<%
      response.write "<a href='Admin_Web_Type_Manage.asp?Action=Modify2&ID=" & rs("ID") & "'>�޸�</a> | "
      response.write "<a href='Admin_Web_Type_Manage.asp?Action=Del1&ID=" & rs("ID") & "' onclick='return ConfirmDel();'>ɾ��</a>"
	  %>

 </td>
    </tr>
   <%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
  </table>
<%
end sub

sub Add1()
%>
<form method="post" name="AddLink" onsubmit="return Check()" action="Admin_Web_Type_Manage.asp">
  <table border="0" cellpadding="2" cellspacing="1" align="center" width="100%" class="border">
    <tr class="title">
      <td height="22" colspan="2" align="center"><strong>���һ�����</strong></td>
    </tr>
    <tr class="tdbg">
      <td width="300" height="25" valign="middle"  align="right"><strong>�������:</strong></td>
      <td height="25"> <input name="SiteName" class="Input" size="30"  maxlength="20">
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr class="tdbg">
      <td width="300"  align="right"><strong>������ɫ:</strong></td>
      <td valign="middle">
          <INPUT TYPE="text" NAME="select_color" class="Input" ID="select_color" size="7" maxlength="7">&nbsp;<img LANGUAGE="javascript" onclick="foreColor()"  WIDTH="18" HEIGHT="18" id="color_img" border=0 src="images/rect.gif" style="cursor:pointer;background-color:#ffffff">
      </td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd1"> 
        <input type="submit" class="Button" value=" ȷ �� " name="cmdOk"> &nbsp; <input type="reset" class="Button" value=" �� �� " name="cmdReset"> 
      </td>
    </tr>
  </table>
</form>
<%
end sub
sub Modify1()
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ�����ID</li>"
		exit sub
	else
		ID=Clng(ID)
	end if
	dim sqlLink,rsLink
	sqlLink="select * from WebSiteType_Main where ID=" & ID
	set rsLink=Server.CreateObject("Adodb.RecordSet")
	rsLink.open sqlLink,conn,1,3
	if rsLink.bof and rsLink.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ�������!</li>"
		rsLink.close
		set rsLink=nothing
		exit sub
	end if

%>
<form method="post" name="AddLink" onsubmit="return Check()" action="Admin_Web_Type_Manage.asp">
  <table border="0" cellpadding="2" cellspacing="1" align="center" width="100%" class="border">
    <tr class="title"> 
      <td height="22" colspan="2" align="center"><strong>�޸�һ�����</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="300" height="25" valign="middle"  align="right"><strong>�������:</strong></td>
      <td height="25"> <input name="SiteName" class="Input" size="30"  maxlength="20"  value="<%=rsLink("Type_Name")%>"> 
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr class="tdbg">
      <td width="300"  align="right"><strong>������ɫ:</strong></td>
      <td valign="middle">
          <INPUT TYPE="text" NAME="select_color" class="Input1" ID="select_color" size="7" maxlength="7" value="<%=rsLink("Type_Color")%>">&nbsp;<img LANGUAGE="javascript" onclick="foreColor()"  WIDTH="18" HEIGHT="18" id="color_img" border=0 src="images/rect.gif" style="cursor:pointer;background-color:#ffffff">
      </td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="ID" type="hidden" id="ID" value="<%=rsLink("ID")%>"> 
        <input name="Action" type="hidden" id="Action" value="SaveModify1"> <input type="submit" class="Button" value=" ȷ �� " name="cmdOk"> 
      </td>
    </tr>
  </table>
</form>
<%
	rsLink.close
	set rsLink=nothing
end sub

sub Add2()
%>
<form method="post" name="AddLink" onsubmit="return Check1()" action="Admin_Web_Type_Manage.asp">
  <table border="0" cellpadding="2" cellspacing="1" align="center" width="100%" class="border">
    <tr class="title">
      <td height="22" colspan="2" align="center"><strong>��Ӷ������</strong></td>
    </tr>
    <tr class="tdbg">
      <td width="300"  align="right"><strong>һ�����:</strong></td>
      <td valign="middle">
      	<SELECT name="SelectType">
	   <OPTION value="0">��ѡ��</OPTION>
	<%
        dim sqlLink,rsLink
	sqlLink="select * from WebSiteType_Main"
	set rsLink=Server.CreateObject("Adodb.RecordSet")
	rsLink.open sqlLink,conn,1,3
        do while not rsLink.eof
           if TypeID=rsLink("ID") then
           response.write "<OPTION value='"&rsLink("ID")&"' selected>"&rsLink("Type_Name")&"</OPTION>"
           else
           response.write "<OPTION value='"&rsLink("ID")&"'>"&rsLink("Type_Name")&"</OPTION>"
           end if
           rsLink.movenext
        loop
	rsLink.close
	set rsLink=nothing        
        %>
	</SELECT>
      </td>
    </tr>
    <tr class="tdbg">
      <td width="300" height="25" valign="middle"  align="right"><strong>�������:</strong></td>
      <td height="25"> <input name="SiteName" class="Input" size="30"  maxlength="20">
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr class="tdbg">
      <td width="300"  align="right"><strong>������ɫ:</strong></td>
      <td valign="middle">
          <INPUT TYPE="text" NAME="select_color" class="Input" ID="select_color" size="7" maxlength="7">&nbsp;<img LANGUAGE="javascript" onclick="foreColor()"  WIDTH="18" HEIGHT="18" id="color_img" border=0 src="images/rect.gif" style="cursor:pointer;background-color:#ffffff">
      </td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd2"> 
        <input type="submit" class="Button" value=" ȷ �� " name="cmdOk"> &nbsp; <input type="reset" class="Button" value=" �� �� " name="cmdReset"> 
      </td>
    </tr>
  </table>
</form>
<%
end sub

sub Modify2()
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ�����ID</li>"
		exit sub
	else
		ID=Clng(ID)
	end if
	dim sqlLink,rsLink
	sqlLink="select * from WebSiteType_Class where ID=" & ID
	set rsLink=Server.CreateObject("Adodb.RecordSet")
	rsLink.open sqlLink,conn,1,3
	if rsLink.bof and rsLink.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ�������!</li>"
		rsLink.close
		set rsLink=nothing
		exit sub
	end if
%>
<form method="post" name="AddLink" onsubmit="return Check1()" action="Admin_Web_Type_Manage.asp">
  <table border="0" cellpadding="2" cellspacing="1" align="center" width="100%" class="border">
    <tr class="title">
      <td height="22" colspan="2" align="center"><strong>�޸Ķ������</strong></td>
    </tr>
    <tr class="tdbg">
      <td width="300"  align="right"><strong>һ�����:</strong></td>
      <td valign="middle">
      	<SELECT name="SelectType">
	   <OPTION value="0">��ѡ��</OPTION>
	<%
        dim rsLinka
	sqlLink="select * from WebSiteType_Main"
	set rsLinka=Server.CreateObject("Adodb.RecordSet")
	rsLinka.open sqlLink,conn,1,3
        do while not rsLinka.eof
           response.write "<OPTION value='"&rsLinka("ID")&"' "
           if rsLink("Main_ID")=rsLinka("ID") then
              response.write "selected" 
           end if
           response.write ">"&rsLinka("Type_Name")&"</OPTION>"
           rsLinka.movenext
        loop
	rsLinka.close
	set rsLinka=nothing
        %>
	</SELECT>
      </td>
    </tr>
    <tr class="tdbg">
      <td width="300" height="25" valign="middle"  align="right"><strong>�������:</strong></td>
      <td height="25"> <input name="SiteName" class="Input" size="30"  maxlength="20" value="<%=rsLink("Type_Class")%>">
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr class="tdbg">
      <td width="300"  align="right"><strong>������ɫ:</strong></td>
      <td valign="middle">
          <INPUT TYPE="text" NAME="select_color" class="Input" ID="select_color" size="7" maxlength="7" value="<%=rsLink("Type_Color")%>">&nbsp;<img LANGUAGE="javascript" onclick="foreColor()"  WIDTH="18" HEIGHT="18" id="color_img" border=0 src="images/rect.gif" style="cursor:pointer;background-color:#ffffff">
      </td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify2"><input name="ID" type="hidden" id="ID" value="<%=rsLink("ID")%>">  
        <input type="submit" class="Button" value=" ȷ �� " name="cmdOk"> &nbsp; <input type="reset" class="Button" value=" �� �� " name="cmdReset"> 
      </td>
    </tr>
  </table>
</form>
<%
     rsLink.close
     set rsLink=nothing
end sub

%></body>
</html>
<%

sub SaveAdd1()
	dim SiteName,LinkColor
        SiteName=trim(request("SiteName"))
        LinkColor=trim(request("select_color"))
	if SiteName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������������!</li>"
	end if
	if LinkColor="#NaNNaNNaN" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>������ɫ�Ƿ���ɫ�ַ���</li>"
	end if
	if LinkColor="" then
                LinkColor="#07519A"
	end if
	if FoundErr=True then
		exit sub
	end if
	if FoundErr<>True then
		dim sqlLink,rsLink
		sqlLink="select top 1 * from WebSiteType_Main where Type_Name='" & dvHtmlEncode(SiteName) & "'"
		set rsLink=Server.CreateObject("Adodb.RecordSet")
		rsLink.open sqlLink,conn,1,3
		if not (rsLink.bof and rsLink.eof) then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��Ҫ��ӵ���������Ѿ����ڣ�</li>"
		else
			rsLink.Addnew
			rsLink("Type_Name")=dvHtmlEncode(SiteName)
			rsLink("Type_Color")=dvHtmlEncode(LinkColor)
			rsLink("Type_Order")=1
			rsLink.update
			rsLink.close
			set rsLink=nothing
			call CloseConn()
			Response.Redirect "Admin_Web_Type_Manage.asp?Action=Add1"
		end if
		rsLink.close
		set rsLink=nothing
	end if
end sub

sub SaveAdd2()
	dim SiteName,LinkColor,SelectType
        SelectType=trim(request("SelectType"))
        SiteName=trim(request("SiteName"))
        LinkColor=trim(request("select_color"))
	if SelectType="0" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ѡ��һ�����!</li>"
	end if
	if SiteName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>����������������!</li>"
	end if
	if LinkColor="#NaNNaNNaN" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>������ɫ�Ƿ���ɫ�ַ���</li>"
	end if
	if LinkColor="" then
                LinkColor="#07519A"
	end if
	if FoundErr=True then
		exit sub
	end if
	if FoundErr<>True then
		dim sqlLink,rsLink
		sqlLink="select top 1 * from WebSiteType_Class where Main_ID="&CLng(SelectType)&" and Type_Class='" & dvHtmlEncode(SiteName) & "'"
		set rsLink=Server.CreateObject("Adodb.RecordSet")
		rsLink.open sqlLink,conn,1,3
		if not (rsLink.bof and rsLink.eof) then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��Ҫ��ӵ���������Ѿ����ڣ�</li>"
		else
			rsLink.Addnew
			rsLink("Main_ID")=SelectType
			rsLink("Type_Class")=dvHtmlEncode(SiteName)
			rsLink("Type_Color")=dvHtmlEncode(LinkColor)
			rsLink("Type_Order")=1
			rsLink.update
			rsLink.close
			set rsLink=nothing
			call CloseConn()
			Response.Redirect "Admin_Web_Type_Manage.asp?Action=Add2&TypeID="&SelectType
		end if
		rsLink.close
		set rsLink=nothing
	end if
end sub

sub SaveModify1()
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ�����ID</li>"
		exit sub
	else
		ID=Clng(ID)
	end if
	dim LinkSiteName,LinkColor
        LinkColor=trim(request("select_color"))
	LinkSiteName=trim(request("SiteName"))
	if LinkSiteName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>���������!</li>"
	end if
	if LinkColor="#NaNNaNNaN" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>������ɫ�Ƿ���ɫ�ַ���</li>"
	end if
	if LinkColor="" then
           LinkColor="#07519A"
	end if
	if FoundErr=True then
		exit sub
	end if
	dim sqlLink,rsLink
	sqlLink="select * from WebSiteType_Main where ID=" & ID
	set rsLink=Server.CreateObject("Adodb.RecordSet")
	rsLink.open sqlLink,conn,1,3
	if rsLink.bof and rsLink.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ�������!</li>"
	else
		rsLink("Type_Name")=dvHtmlEncode(LinkSiteName)
		rsLink("Type_Color")=dvHtmlEncode(LinkColor)
		rsLink.update
		rsLink.close
		set rsLink=nothing
		call CloseConn()
		Response.Redirect "Admin_Web_Type_Manage.asp"
	end if
	rsLink.close
	set rsLink=nothing
end sub
sub SaveModify2()
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ�����ID</li>"
		exit sub
	else
		ID=Clng(ID)
	end if
	dim LinkSiteName,LinkColor,SelectType
        SelectType=trim(request("SelectType"))
        LinkColor=trim(request("select_color"))
	LinkSiteName=trim(request("SiteName"))
	if LinkSiteName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>���������!</li>"
	end if
	if LinkColor="#NaNNaNNaN" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>������ɫ�Ƿ���ɫ�ַ���</li>"
	end if
	if LinkColor="" then
           LinkColor="#07519A"
	end if
	if FoundErr=True then
		exit sub
	end if
	dim sqlLink,rsLink
	sqlLink="select * from WebSiteType_Class where ID=" & ID
	set rsLink=Server.CreateObject("Adodb.RecordSet")
	rsLink.open sqlLink,conn,1,3
	if rsLink.bof and rsLink.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ID!</li>"
	else
		rsLink("Main_ID")=SelectType
		rsLink("Type_Class")=dvHtmlEncode(LinkSiteName)
		rsLink("Type_Color")=dvHtmlEncode(LinkColor)
		rsLink.update
		rsLink.close
		set rsLink=nothing
		call CloseConn()
		Response.Redirect "Admin_Web_Type_Manage.asp?Action=list2"
	end if
	rsLink.close
	set rsLink=nothing
end sub
%>
