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
dim Action,FoundErr,ErrMsg,Select1,Select2
Action=trim(request("Action"))
ID=Trim(Request("ID"))
Select1=Trim(Request("Select1"))
Select2=Trim(Request("Select2"))
if Select1="" then
   Select1=0
else
Select1=CLng(Select1)
end if
if Select2="" then
   Select2=0
else
Select2=CLng(Select2)
end if
str_select=trim(request("select_Query"))
str_text=trim(request("Query_text"))
strFileName="Admin_Web_AdvertisementURL_Manage.asp?select_Query="&str_select&"&Query_text="&str_text

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

if ID<>"" then
	if Action="Del" then
		conn.execute "Delete From WebSiteAdvertisementWords_URL Where ID=" & CLng(ID)
	end if
end if
%>
<html>
<head>
<title>网址导航管理</title>
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
	  alert("请输入广告词名称!")
	  document.AddLink.SiteName.focus()
	  return false
	 }
if (document.AddLink.WebUrl.value=="")
	{
	  alert("请输入广告词链接!")
	  document.AddLink.WebUrl.focus()
	  return false
	 }
}
function ConfirmDel()
{
   if(confirm("确定要删除吗?"))
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
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" onLoad="InitDocument()">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg">
    <td height="22" colspan="2" align="center"><strong>广 告 词 管 理</strong></td>
  </tr>
  <tr class="tdbg">
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"> <a href="Admin_Web_AdvertisementURL_Manage.asp?Action=list">广告词管理</a> | <a href="Admin_Web_AdvertisementURL_Manage.asp?Action=Add">添加广告词</a></td>
  </tr>
</table>
<br>
<%
if Action="Add" then
	call Add()
elseif Action="SaveAdd" then
	call SaveAdd()
elseif Action="Modify" then
	call Modify()
elseif Action="SaveModify" then
	call SaveModify()
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()

sub main()
	sql="select * from WebSiteAdvertisementWords_URL "
        if str_select="请选择" or str_select="" then
	sql=sql
        else
	sql=sql&" where ("&str_select&" like '%"&str_text&"%')"
        end if
	sql=sql&" order by id desc"

	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1

 	if rs.eof and rs.bof then
		response.write "有 0 条记录"
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
        	showpage strFileName,totalput,MaxPerPage,true,true,"条记录"
   	    else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	List
            	showpage strFileName,totalput,MaxPerPage,true,true,"条记录"
        	else
	        	currentPage=1
           		List
           		showpage strFileName,totalput,MaxPerPage,true,true,"条记录"
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
<form method="post" name="AddLink" action="Admin_Web_AdvertisementURL_Manage.asp">
  <tr>
    <td width="76" align="right"><b>查询:</b>&nbsp;&nbsp;</td>
    <td><select size="1" name="select_Query">
        <option selected value="请选择">请选择</option>
        <option <%if str_select="Site_Title" then response.write "selected" end if%> value="Site_Title">广告词名称</option>
        </select>&nbsp;<input name="Query_text" size="30" class="Input" maxlength="100" type="text" value="<%=str_text%>">&nbsp;<input class="Button" type="submit" value="查询" name="B3"></td>
  </tr>
  <tr>
    <td colspan="2" height="5"></td>
  </tr>
</form>
</table>
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr class="title">
      <td width="10%" height="22" align="center">ID</td>
      <td width="75%" height="22" align="center">广告词名称</td>
      <td height="22" align="center">操作</td>
    </tr>
<%
do while not rs.eof
%>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
    	<td align="center"><%=rs("ID")%></td>
      <td><a href="<%=rs("Site_Url")%>" target="_blank" style="color: <%=rs("Font_Color")%>"><%=rs("Site_Title")%></a></td>
      <td align="center">
<%
      response.write "<a href='Admin_Web_AdvertisementURL_Manage.asp?Action=Modify&ID=" & rs("ID") & "'>修改</a> | "
      response.write "<a href='Admin_Web_AdvertisementURL_Manage.asp?Action=Del&ID=" & rs("ID") & "' onclick='return ConfirmDel();'>删除</a>"
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

sub Add()
%>
<form method="post" name="AddLink" onSubmit="return Check()" action="Admin_Web_AdvertisementURL_Manage.asp">
  <table border="0" cellpadding="2" cellspacing="1" align="center" width="100%" class="border">
    <tr class="title">
      <td height="22" colspan="2" align="center"><strong>添加广告词</strong></td>
    </tr>
    <tr class="tdbg">
      <td width="85" height="25" valign="middle"  align="right"><strong>广告词名称:</strong></td>
      <td height="25"> <input name="SiteName" class="Input" size="30"  maxlength="50">
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr class="tdbg">
      <td width="85" height="25" valign="middle"  align="right"><strong>广告词链接:</strong></td>
      <td height="25"> <input name="WebUrl" class="Input" size="30"  maxlength="100">
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr class="tdbg">
      <td width="85"  align="right"><strong>字体颜色:</strong></td>
      <td valign="middle">
          <INPUT TYPE="text" NAME="select_color" class="Input" ID="select_color" size="7" maxlength="7">&nbsp;<img LANGUAGE="javascript" onClick="foreColor()"  WIDTH="18" HEIGHT="18" id="color_img" border=0 src="images/rect.gif" style="cursor:pointer;background-color:#ffffff">
      </td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd"> 
        <input type="submit" class="Button" value=" 确 定 " name="cmdOk"> &nbsp; <input type="reset" class="Button" value=" 重 填 " name="cmdReset"> 
      </td>
    </tr>
  </table>
</form>
<%
end sub

sub Modify()
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定ID</li>"
		exit sub
	else
		ID=Clng(ID)
	end if
	dim sqlLink,rsLink,ClassID
	sqlLink="select * from WebSiteAdvertisementWords_URL where ID=" & ID
	set rsLink=Server.CreateObject("Adodb.RecordSet")
	rsLink.open sqlLink,conn,1,3
	if rsLink.bof and rsLink.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到ID!</li>"
		rsLink.close
		set rsLink=nothing
		exit sub
	end if
%>
<form method="post" name="AddLink" onSubmit="return Check()" action="Admin_Web_AdvertisementURL_Manage.asp">
  <table border="0" cellpadding="2" cellspacing="1" align="center" width="100%" class="border">
    <tr class="title">
      <td height="22" colspan="2" align="center"><strong>修改广告词</strong></td>
    </tr>
    <tr class="tdbg">
      <td width="300" height="25" valign="middle"  align="right"><strong>广告词名称:</strong></td>
      <td height="25"> <input name="SiteName" class="Input" size="30"  maxlength="50" value="<%=rsLink("Site_Title")%>">
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr class="tdbg">
      <td width="300" height="25" valign="middle"  align="right"><strong>广告词链接:</strong></td>
      <td height="25"> <input name="WebUrl" class="Input" size="30"  maxlength="100" value="<%=rsLink("Site_Url")%>">
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr class="tdbg">
      <td width="300"  align="right"><strong>字体颜色:</strong></td>
      <td valign="middle">
          <INPUT TYPE="text" NAME="select_color" class="Input" ID="select_color" size="7" maxlength="7" value="<%=rsLink("Font_Color")%>">&nbsp;<img LANGUAGE="javascript" onClick="foreColor()"  WIDTH="18" HEIGHT="18" id="color_img" border=0 src="images/rect.gif" style="cursor:pointer;background-color:#ffffff">
      </td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"><input name="ID" type="hidden" id="ID" value="<%=rsLink("ID")%>">  
        <input type="submit" class="Button" value=" 确 定 " name="cmdOk"> &nbsp; <input type="reset" class="Button" value=" 重 填 " name="cmdReset"> 
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
sub SaveAdd()
	dim WebName,WebUrl,WebColor
        WebName=trim(request("SiteName"))
        WebUrl=trim(request("WebUrl"))
        WebColor=trim(request("select_color"))
	if WebName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入广告词名称!</li>"
	end if
	if WebUrl="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入广告词链接!</li>"
        else
           if InStr(WebUrl,"http://")=0 then
              WebUrl="http://"&WebUrl
           end if
	end if
	if WebColor="#NaNNaNNaN" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>字体颜色是非颜色字符！</li>"
	end if
	if WebColor="" then
                WebColor="#07519A"
	end if
	if FoundErr=True then
		exit sub
	end if
	if FoundErr<>True then
		dim sqlLink,rsLink
		sqlLink="select top 1 * from WebSiteAdvertisementWords_URL where Site_Title='" & dvHtmlEncode(WebName) & "'"
		set rsLink=Server.CreateObject("Adodb.RecordSet")
		rsLink.open sqlLink,conn,1,3
		if not (rsLink.bof and rsLink.eof) then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>你要添加的广告词已经存在！</li>"
		else
			rsLink.Addnew
			rsLink("Site_Url")=dvHtmlEncode(WebUrl)
			rsLink("Site_Title")=dvHtmlEncode(WebName)
			rsLink("Font_Color")=dvHtmlEncode(WebColor)
			rsLink.update
			rsLink.close
			set rsLink=nothing
			call CloseConn()
			Response.Redirect "Admin_Web_AdvertisementURL_Manage.asp?Action=Add"
		end if
		rsLink.close
		set rsLink=nothing
	end if
end sub

sub SaveModify()
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定广告词ID</li>"
		exit sub
	else
		ID=Clng(ID)
	end if
	dim WebName,WebUrl,WebColor
        WebName=trim(request("SiteName"))
        WebUrl=trim(request("WebUrl"))
        WebColor=trim(request("select_color"))
	if WebName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入广告词名称!</li>"
	end if
	if WebUrl="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入广告词链接!</li>"
        else
           if InStr(WebUrl,"http://")=0 then
              WebUrl="http://"&WebUrl
           end if
	end if
	if WebColor="#NaNNaNNaN" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>字体颜色是非颜色字符！</li>"
	end if
	if WebColor="" then
                WebColor="#07519A"
	end if
	if FoundErr=True then
		exit sub
	end if
	dim sqlLink,rsLink
	sqlLink="select * from WebSiteAdvertisementWords_URL where ID=" & ID
	set rsLink=Server.CreateObject("Adodb.RecordSet")
	rsLink.open sqlLink,conn,1,3
	if rsLink.bof and rsLink.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到广告词!</li>"
	else
		rsLink("Site_Url")=dvHtmlEncode(WebUrl)
		rsLink("Site_Title")=dvHtmlEncode(WebName)
		rsLink("Font_Color")=dvHtmlEncode(WebColor)
		rsLink.update
		rsLink.close
		set rsLink=nothing
		call CloseConn()
		Response.Redirect "Admin_Web_AdvertisementURL_Manage.asp"
	end if
	rsLink.close
	set rsLink=nothing
end sub
%>