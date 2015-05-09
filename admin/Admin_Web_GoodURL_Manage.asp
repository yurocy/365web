<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
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
const MaxPerPage=30
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
strFileName="Admin_Web_GoodURL_Manage.asp?select_Query="&str_select&"&Query_text="&str_text

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

if ID<>"" then
	if Action="Del" then
		conn.execute "Delete From WebGoodSite_Url Where ID=" & CLng(ID)
	end if
	if Action="Del_lite" then
		conn.execute "Delete From WebGoodSite_Url_lite Where ID=" & CLng(ID)
		response.write "<script>location='"&Request.ServerVariables("HTTP_REFERER")&"';</script>"
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
function tsp(e,id){

if(e.checked==true)
{
document.getElementById("xsWebUrl"+id).value='/out.asp?turl='+document.getElementById("xsWebUrl"+id).value;
}
else
{
var aaa=document.getElementById("xsWebUrl"+id).value;
document.getElementById("xsWebUrl"+id).value=aaa.replace('/out.asp?turl=','');
}
}


function Check() {
if (document.AddLink.SelectType.value=="0")
	{
	  alert("请选择一级类别!")
	  document.AddLink.SelectType.focus()
	  return false
	 }
if (document.AddLink.SelectTypea.value=="0")
	{
	  alert("请选择二级类别!")
	  document.AddLink.SelectTypea.focus()
	  return false
	 }
if (document.AddLink.SiteName.value=="")
	{
	  alert("请输入网站名称!")
	  document.AddLink.SiteName.focus()
	  return false
	 }
if (document.AddLink.WebUrl.value=="")
	{
	  alert("请输入网站链接!")
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
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" onLoad="InitDocument()">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg">
    <td height="22" colspan="2" align="center"><strong> 网 址 管 理</strong></td>
  </tr>
  <tr class="tdbg">
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"> <a href="Admin_Web_GoodURL_Manage.asp?Action=list">网址管理</a> | <a href="Admin_Web_GoodURL_Manage.asp?Action=Add">添加网站</a>| <a href="Admin_Web_GoodURL_Manage.asp?Action=AddMuch">批量添加</a></td>
  </tr>
</table>
<br>
<%
if Action="Add" then
	call Add()
elseif Action="AddMuch" then
	call AddMuch()
elseif Action="SaveAddMuch" then
	call SaveAddMuch()	
elseif Action="SaveAdd" then
	call SaveAdd()
elseif Action="Modify" then
	call Modify()
elseif Action="SaveModify" then
	call SaveModify()
elseif Action="SaveModify_lite" then
	call SaveModify_lite()
elseif Action="add_lite" then
call add_lite()
elseif Action="SaveAdd_lite" then
	call SaveAdd_lite()
elseif Action="list_lite" then
call list_lite()
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()

sub main()
	sql="select a.Class_ID,a.Is_Recommend,a.bold,a.ID,a.Site_Url,a.Site_Title,a.Font_Color,b.Type_Class,c.Type_Name,c.type_order from WebGoodSite_Url a,WebGoodSiteType_Class b,WebGoodSiteType_Main c where a.Class_ID=b.ID and b.Main_ID=c.ID"
        if str_select="请选择" or str_select="" then
	sql=sql
        else
	sql=sql&" and ("&str_select&" like '%"&str_text&"%')"
        end if
	sql=sql&" order by c.type_order desc,a.Class_ID asc,a.Is_Recommend desc,a.id asc"

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
<form method="post" name="AddLink" action="Admin_Web_GoodURL_Manage.asp">
  <tr>
    <td width="76" align="right"><b>查询:</b>&nbsp;&nbsp;</td>
    <td><select size="1" name="select_Query">
        
		 <option selected value="a.Site_Title">网站名称</option>
		  <option<%if str_select="a.Site_URL" then response.write "selected" end if%> value="a.Site_Title">网站地址</option>
        <option <%if str_select="c.Type_Name" then response.write "selected" end if%> value="c.Type_Name">一级类别</option>
        <option <%if str_select="b.Type_Class" then response.write "selected" end if%> value="b.Type_Class">二级类别</option>
       
        </select>&nbsp;
        <input name="Query_text" size="50" class="Input" maxlength="100" type="text" value="<%=str_text%>">
        &nbsp;
        <input class="Button" type="submit" value="查询" name="B3"></td>
  </tr>
  <tr>
    <td colspan="2" height="5"></td>
  </tr>
</form>
</table>
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr class="title">
      <td width="100" height="22" align="center">一级类别</td>
      <td width="100" height="22" align="center">二级类别</td>
      <td width="200" height="22" align="center">网站名称</td>
      <td width="50" height="22" align="center">粗体</td>
	  <td width="50" height="22" align="center">推荐</td>
      <td width="100" height="22" align="center">操作</td>
    </tr>
<%
do while not rs.eof
%>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="100" align="center"><%=rs("Type_Name")%></td>
      <td width="100" align="center"><%=rs("Type_Class")%></td>
      <td width="200" align="center"><a href="<%=rs("Site_Url")%>" target="_blank" style="color: <%=rs("Font_Color")%>"><%=rs("Site_Title")%></a></td>
	        <td width="50" align="center">
      <% 
         if rs("bold")=1 then
            response.write "是"
         else
            response.write "否"
         end if
       %>
      </td>
      <td width="50" align="center">
      <% 
         if rs("Is_Recommend")=1 then
            response.write "是"
         else
            response.write "否"
         end if
       %>
      </td>
      <td width="100" align="center">
<% 	
      response.write "<a href='Admin_Web_GoodURL_Manage.asp?Action=Modify&ID=" & rs("ID") & "'>查看修改</a> | "
      response.write "<a href='Admin_Web_GoodURL_Manage.asp?Action=Del&ID=" & rs("ID") & "' onclick='return ConfirmDel();'>删除</a>"
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
<form method="post" name="AddLink" onSubmit="return Check()" action="Admin_Web_GoodURL_Manage.asp">
  <table border="0" cellpadding="2" cellspacing="1" align="center" width="100%" class="border">
    <tr class="title">
      <td height="22" colspan="2" align="center"><strong>添加网站</strong></td>
    </tr>
    <tr class="tdbg">
      <td width="85"  align="right"><strong>一级类别:</strong></td>
      <td valign="middle">
<%
dim count
sql="select * from WebGoodSiteType_Class"
set rs=server.CreateObject ("Adodb.recordset")
rs.open sql,conn,1,3
%>
<SCRIPT language="JavaScript">
var onecount;
onecount=0;
subcat = new Array();

<%
count = 0
do while not rs.eof 
%>
subcat[<%=count%>] = new Array("<%= trim(rs("Type_Class"))%>","<%= trim(rs("Main_ID"))%>","<%= trim(rs("ID"))%>");
<%
count = count + 1
rs.movenext
loop
rs.close

%>
onecount=<%=count%>;
function changelocation(locationid)
{
document.AddLink.SelectTypea.length = 0; 
var locationid=locationid;
var i;
document.AddLink.SelectTypea.options[document.AddLink.SelectTypea.length] = new Option('请选择','0','请选择');
for (i=0;i < onecount; i++)
{
if (subcat[i][1] == locationid)
{
document.AddLink.SelectTypea.options[document.AddLink.SelectTypea.length] = new Option(subcat[i][0], subcat[i][2]);
} 
}
} 

</SCRIPT>
      	<SELECT name="SelectType" onChange="changelocation(document.AddLink.SelectType.options[document.AddLink.SelectType.selectedIndex].value)">
	   <OPTION value="0">请选择</OPTION>
	<%
        dim sqlLink,rsLink
	sqlLink="select * from WebGoodSiteType_Main"
	set rsLink=Server.CreateObject("Adodb.RecordSet")
	rsLink.open sqlLink,conn,1,3
        do while not rsLink.eof
           if Select1=rsLink("ID") then
              response.write "<OPTION value='"&rsLink("ID")&"' selected>"&rsLink("Type_Name")&"</OPTION>"
           else
              response.write "<OPTION value='"&rsLink("ID")&"'>"&rsLink("Type_Name")&"</OPTION>"
           end if
           rsLink.movenext
        loop
	rsLink.close
	set rsLink=nothing        
        %>
	</SELECT>      </td>
    </tr>
    <tr class="tdbg">
      <td width="85"  align="right"><strong>二级类别:</strong></td>
      <td valign="middle">
      	<SELECT name="SelectTypea">
	   <OPTION value="0">请选择</OPTION>
        <%
        dim rsLinka
	sqlLink="select * from WebGoodSiteType_Class where main_id="&Select1
	set rsLinka=Server.CreateObject("Adodb.RecordSet")
	rsLinka.open sqlLink,conn,1,3
        do while not rsLinka.eof
           response.write "<OPTION value='"&rsLinka("ID")&"' "
           if Select2=rsLinka("ID") then
              response.write "selected"
           end if
           response.write ">"&rsLinka("Type_Class")&"</OPTION>"
           rsLinka.movenext
        loop
	rsLinka.close
	set rsLinka=nothing
           %>
	</SELECT>      </td>
    </tr>
    <tr class="tdbg">
      <td width="85" height="25" valign="middle"  align="right"><strong>网站名称:</strong></td>
      <td height="25"><input name="SiteName" class="Input" size="50"  maxlength="50">
          <font color="#FF0000"> *</font>粗体:<input type="checkbox" name="bold" value="1" ></td>
    </tr>
    <tr class="tdbg">
      <td width="85" height="25" valign="middle"  align="right"><strong>网站链接:</strong></td>
      <td height="25"><input id="WebUrl" name="WebUrl" class="Input" size="50"  maxlength="100">
          <font color="#FF0000">  中转跳转:
      <input type="checkbox" name="out" value="1" onClick="if(this.checked==true){document.AddLink.WebUrl.value='/out.asp?turl='+document.AddLink.WebUrl.value}else{var aaa=document.AddLink.WebUrl.value;document.AddLink.WebUrl.value=aaa.replace('/out.asp?turl=','');}"> </font></td>
    </tr>
    <tr class="tdbg">
      <td width="85"  align="right"><strong>字体颜色:</strong></td>
      <td valign="middle"><input type="text" name="select_color" class="Input" id="select_color" size="7" maxlength="7">
        &nbsp;<img language="javascript" onClick="foreColor()"  width="18" height="18" id="color_img" border=0 src="images/rect.gif" style="cursor:pointer;background-color:#ffffff"> </td>
    </tr>

    <tr class="tdbg">
      <td width="85" height="25"  align="right" valign="middle"><strong>推荐:</strong></td>
      <td height="25"><p>
          <input type="checkbox" name="Recommend" value="1" >
      </p></td>
    </tr>
    <tr class="tdbg">
      <td width="85" height="25"  align="right" valign="middle"><strong>样式:</strong></td>
      <td height="25"><input id="radio" type="radio" name="num1" value="1" />
          <img src="/Images/1.gif" />
          <input id="radio2" type="radio" name="num1" value="2" />
          <img src="/Images/2.gif" />
          <input id="radio3" type="radio" name="num1" value="3" />
          <img src="/Images/3.gif" />
          <input id="radio4" type="radio" name="num1" value="4" />
          <img src="/Images/4.gif" />
          <input id="radio5" type="radio" name="num1" value="5" />
          <img src="/Images/5.gif" />
          <input id="radio6" type="radio" name="num1" value="0" checked="checked" />
          <label for="radio6">无</label>
        &nbsp;</td>
    </tr>
	<%for i=1 to 10%>
    <tr class="tdbg">
      <td width="85" height="25"  align="right" valign="middle"><strong>下属名称<%=i%>:</strong></td>
      <td height="25"><input name="xsSiteName" class="Input" id="xsSiteName" size="50"  maxlength="50"></td>
    </tr>
    <tr class="tdbg">
      <td width="85" height="25"  align="right" valign="middle"><strong>链接<%=i%>:</strong></td>
      <td height="25"><input name="xsWebUrl" class="Input" id="xsWebUrl<%=i%>" size="50"  maxlength="100">  中转跳转:<input type="checkbox" value="1" onClick="tsp(this,'<%=i%>');"> </td>
    </tr>
<%next%>
    <tr class="tdbg">
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd">
          <input type="submit" class="Button" value=" 确 定 " name="cmdOk2">
        &nbsp;
        <input type="reset" class="Button" value=" 重 填 " name="cmdReset2">      </td>
    </tr>
  </table>
</form>
<%
end sub

sub AddMuch()
%>
<form method="post" name="AddLink" onSubmit="return Check()" action="Admin_Web_GoodURL_Manage.asp">
  <table border="0" cellpadding="2" cellspacing="1" align="center" width="100%" class="border">
    <tr class="title">
      <td height="22" colspan="2" align="center"><strong>批量添加</strong></td>
    </tr>
    <tr class="tdbg">
      <td width="490" align="right"><strong>一级类别:</strong></td>
      <td width="595" valign="middle">
<%
dim count
sql="select * from WebGoodSiteType_Class"
set rs=server.CreateObject ("Adodb.recordset")
rs.open sql,conn,1,3
%>
<SCRIPT language="JavaScript">
var onecount;
onecount=0;
subcat = new Array();

<%
count = 0
do while not rs.eof 
%>
subcat[<%=count%>] = new Array("<%= trim(rs("Type_Class"))%>","<%= trim(rs("Main_ID"))%>","<%= trim(rs("ID"))%>");
<%
count = count + 1
rs.movenext
loop
rs.close

%>
onecount=<%=count%>;
function changelocation(locationid)
{
document.AddLink.SelectTypea.length = 0; 
var locationid=locationid;
var i;
document.AddLink.SelectTypea.options[document.AddLink.SelectTypea.length] = new Option('请选择','0','请选择');
for (i=0;i < onecount; i++)
{
if (subcat[i][1] == locationid)
{
document.AddLink.SelectTypea.options[document.AddLink.SelectTypea.length] = new Option(subcat[i][0], subcat[i][2]);
} 
}
} 

</SCRIPT>
      	<SELECT name="SelectType" onChange="changelocation(document.AddLink.SelectType.options[document.AddLink.SelectType.selectedIndex].value)">
	   <OPTION value="0">请选择</OPTION>
	<%
        dim sqlLink,rsLink
	sqlLink="select * from WebGoodSiteType_Main"
	set rsLink=Server.CreateObject("Adodb.RecordSet")
	rsLink.open sqlLink,conn,1,3
        do while not rsLink.eof
           if Select1=rsLink("ID") then
              response.write "<OPTION value='"&rsLink("ID")&"' selected>"&rsLink("Type_Name")&"</OPTION>"
           else
              response.write "<OPTION value='"&rsLink("ID")&"'>"&rsLink("Type_Name")&"</OPTION>"
           end if
           rsLink.movenext
        loop
	rsLink.close
	set rsLink=nothing        
        %>
	</SELECT>      </td>
    </tr>
    <tr class="tdbg">
      <td width="490" align="right"><strong>二级类别:</strong></td>
      <td valign="middle">
      	<SELECT name="SelectTypea">
	   <OPTION value="0">请选择</OPTION>
        <%
        dim rsLinka
	sqlLink="select * from WebGoodSiteType_Class where main_id="&Select1
	set rsLinka=Server.CreateObject("Adodb.RecordSet")
	rsLinka.open sqlLink,conn,1,3
        do while not rsLinka.eof
           response.write "<OPTION value='"&rsLinka("ID")&"' "
           if Select2=rsLinka("ID") then
              response.write "selected"
           end if
           response.write ">"&rsLinka("Type_Class")&"</OPTION>"
           rsLinka.movenext
        loop
	rsLinka.close
	set rsLinka=nothing
           %>
	</SELECT>      </td>
    </tr>
    <tr class="tdbg">
      <td width="490"  height="26"  align="right" valign="middle"><strong>网站名称:
        <input name="SiteName" class="Input" size="50"  maxlength="50">
      </strong></td>
      <td height="26"><strong>网站链接:
        <input name="WebUrl" class="Input" size="50"  maxlength="100">
      </strong></td>
    </tr>
    <tr class="tdbg">
      <td width="490" height="25"  align="right" valign="middle"><strong>网站名称:
        <input name="SiteName" class="Input" size="50"  maxlength="50">
      </strong></td>
      <td height="25"><strong>网站链接:
        <input name="WebUrl" class="Input" size="50"  maxlength="100">
      </strong></td>
    </tr>
    <tr class="tdbg">
      <td width="490" height="25"  align="right" valign="middle"><strong>网站名称:
        <input name="SiteName" class="Input" size="50"  maxlength="50">
      </strong></td>
      <td height="25"><strong>网站链接:
        <input name="WebUrl" class="Input" size="50"  maxlength="100">
      </strong></td>
    </tr>
    <tr class="tdbg">
      <td width="490" height="25"  align="right" valign="middle"><strong>网站名称:
        <input name="SiteName" class="Input" size="50"  maxlength="50">
      </strong></td>
      <td height="25"><strong>网站链接:
        <input name="WebUrl" class="Input" size="50"  maxlength="100">
      </strong></td>
    </tr>
    <tr class="tdbg">
      <td width="490" height="25"  align="right" valign="middle"><strong>网站名称:
        <input name="SiteName" class="Input" size="50"  maxlength="50">
      </strong></td>
      <td height="25"><strong>网站链接:
        <input name="WebUrl" class="Input" size="50"  maxlength="100">
      </strong></td>
    </tr>
    <tr class="tdbg">
      <td width="490" height="25"  align="right" valign="middle"><strong>网站名称:
        <input name="SiteName" class="Input" size="50"  maxlength="50">
      </strong></td>
      <td height="25"><strong>网站链接:
        <input name="WebUrl" class="Input" size="50"  maxlength="100">
      </strong></td>
    </tr>
    <tr class="tdbg">
      <td width="490" height="25"  align="right" valign="middle"><strong>网站名称:
        <input name="SiteName" class="Input" size="50"  maxlength="50">
      </strong></td>
      <td height="25"><strong>网站链接:
        <input name="WebUrl" class="Input" size="50"  maxlength="100">
      </strong></td>
    </tr>
    <tr class="tdbg">
      <td width="490" height="25"  align="right" valign="middle"><strong>网站名称:
        <input name="SiteName" class="Input" size="50"  maxlength="50">
      </strong></td>
      <td height="25"><strong>网站链接:
        <input name="WebUrl" class="Input" size="50"  maxlength="100">
      </strong></td>
    </tr>
    <tr class="tdbg">
      <td width="490" height="25"  align="right" valign="middle"><strong>网站名称:
        <input name="SiteName" class="Input" size="50"  maxlength="50">
      </strong></td>
      <td height="25"><strong>网站链接:
        <input name="WebUrl" class="Input" size="50"  maxlength="100">
      </strong></td>
    </tr>
    <tr class="tdbg">
      <td width="490" height="25"  align="right" valign="middle"><strong>网站名称:
        <input name="SiteName" class="Input" size="50"  maxlength="50">
      </strong></td>
      <td height="25"><strong>网站链接:
        <input name="WebUrl" class="Input" size="50"  maxlength="100">
      </strong></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAddMuch"> 
        <input type="submit" class="Button" value=" 确 定 " name="cmdOk"> &nbsp; <input type="reset" class="Button" value=" 重 填 " name="cmdReset">      </td>
    </tr>
  </table>
</form>
<%
end sub
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
sub Add_lite()
%>
<form method="post" name="AddLink" onSubmit="return Check()" action="Admin_Web_GoodURL_Manage.asp">
  <table border="0" cellpadding="2" cellspacing="1" align="center" width="100%" class="border">
    <tr class="title">
      <td height="22" colspan="2" align="center"><strong>添加下属网站</strong></td>
    </tr>
    <tr class="tdbg">
      <td width="85" height="25" valign="middle"  align="right"><strong>下属网站名称:</strong></td>
      <td height="25"> <input name="SiteName" class="Input" size="50"  maxlength="50">
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr class="tdbg">
      <td width="85" height="25" valign="middle"  align="right"><strong>下属网站链接:</strong></td>
      <td height="25"> <input name="WebUrl" class="Input" size="50"  maxlength="100">
        <font color="#FF0000"> * 中转跳转:<input type="checkbox" name="out" value="1" onClick="if(this.checked==true){document.AddLink.WebUrl.value='/out.asp?turl='+document.AddLink.WebUrl.value}else{var aaa=document.AddLink.WebUrl.value;document.AddLink.WebUrl.value=aaa.replace('/out.asp?turl=','');}"></font></td>
    </tr>
    <tr class="tdbg">
      <td width="85"  align="right"><strong>字体颜色:</strong></td>
      <td valign="middle">
          <INPUT TYPE="text" NAME="select_color" class="Input" ID="select_color" size="7" maxlength="7">&nbsp;<img LANGUAGE="javascript" onClick="foreColor()"  WIDTH="18" HEIGHT="18" id="color_img" border=0 src="images/rect.gif" style="cursor:pointer;background-color:#ffffff">      </td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd_lite"> <input name="class_id" type="hidden" id="class_id" value="<%=request("id")%>">
        <input type="submit" class="Button" value=" 确 定 " name="cmdOk"> &nbsp; <input type="reset" class="Button" value=" 重 填 " name="cmdReset">      </td>
    </tr>
  </table>
</form>
<%
end sub
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
sub Modify()
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定类别ID</li>"
		exit sub
	else
		ID=Clng(ID)
	end if
	dim sqlLink,rsLink,ClassID
	sqlLink="select * from WebGoodSite_Url where ID=" & ID
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

  <table border="0" cellpadding="2" cellspacing="1" align="center" width="100%" class="border">
  <form method="post" name="AddLink" onSubmit="return Check()" action="Admin_Web_GoodURL_Manage.asp">
    <tr class="title">
      <td height="22" colspan="2" align="center"><strong>修改网站</strong></td>
    </tr>
    <tr class="tdbg">
      <td width="160"  align="right"><strong>一级类别:</strong></td>
      <td width="1233" valign="middle">
<%
dim count
sql="select * from WebGoodSiteType_Class"
set rs=server.CreateObject ("Adodb.recordset")
rs.open sql,conn,1,3
%>
<SCRIPT language="JavaScript">
var onecount;
onecount=0;
subcat = new Array();

<%
count = 0
do while not rs.eof 
%>
subcat[<%=count%>] = new Array("<%= trim(rs("Type_Class"))%>","<%= trim(rs("Main_ID"))%>","<%= trim(rs("ID"))%>");
<%
count = count + 1
rs.movenext
loop
rs.close

%>
onecount=<%=count%>;
function changelocation(locationid)
{
document.AddLink.SelectTypea.length = 0; 
var locationid=locationid;
var i;
document.AddLink.SelectTypea.options[document.AddLink.SelectTypea.length] = new Option('请选择','0','请选择');
for (i=0;i < onecount; i++)
{
if (subcat[i][1] == locationid)
{
document.AddLink.SelectTypea.options[document.AddLink.SelectTypea.length] = new Option(subcat[i][0], subcat[i][2]);
} 
}
} 

</SCRIPT>
      	<SELECT name="SelectType" onChange="changelocation(document.AddLink.SelectType.options[document.AddLink.SelectType.selectedIndex].value)">
	   <OPTION value="0">请选择</OPTION>
	<%
        dim rsclass
	sqlLink="select * from WebGoodSiteType_Class where ID=" & rsLink("Class_ID")
	set rsclass=Server.CreateObject("Adodb.RecordSet")
	rsclass.open sqlLink,conn,1,3
	if rsclass.bof and rsclass.eof then
           ClassID=0
        else
           ClassID=rsclass("Main_ID")
	end if
	rsclass.close
	set rsclass=nothing

        dim rsLinka
	sqlLink="select * from WebGoodSiteType_Main"
	set rsLinka=Server.CreateObject("Adodb.RecordSet")
	rsLinka.open sqlLink,conn,1,3
        do while not rsLinka.eof
           response.write "<OPTION value='"&rsLinka("ID")&"' "
           if ClassID=rsLinka("ID") then
              response.write "selected"
           end if
           response.write ">"&rsLinka("Type_Name")&"</OPTION>"
           rsLinka.movenext
        loop
	rsLinka.close
	set rsLinka=nothing
        %>
	</SELECT>      </td>
    </tr>
    <tr class="tdbg">
      <td width="160"  align="right"><strong>二级类别:</strong></td>
      <td valign="middle">
      	<SELECT name="SelectTypea">
	   <OPTION value="0" selected>请选择</OPTION>
           <%
	sqlLink="select * from WebGoodSiteType_Class where main_id="&ClassID
	set rsLinka=Server.CreateObject("Adodb.RecordSet")
	rsLinka.open sqlLink,conn,1,3
        do while not rsLinka.eof
           response.write "<OPTION value='"&rsLinka("ID")&"' "
           if rsLink("Class_ID")=rsLinka("ID") then
              response.write "selected"
           end if
           response.write ">"&rsLinka("Type_Class")&"</OPTION>"
           rsLinka.movenext
        loop
	rsLinka.close
	set rsLinka=nothing
           %>
	</SELECT>      </td>
    </tr>
    <tr class="tdbg">
      <td width="160" height="25"  align="right" valign="middle"><strong>网站名称:</strong></td>
      <td height="25"> <input name="SiteName" class="Input" size="50"  maxlength="50" value="<%=rsLink("Site_Title")%>">
        <font color="#FF0000"> *</font>粗体:<input type="checkbox" name="bold" value="1" <% if rsLink("bold")=1 then response.write "checked" end if%>></td>
    </tr>
    <tr class="tdbg">
      <td width="160" height="25"  align="right" valign="middle"><strong>网站链接:</strong></td>
      <td height="25"> <input name="WebUrl" class="Input" size="50"  maxlength="100" value="<%=rsLink("Site_Url")%>">
        <font color="#FF0000"> * 测速:<input type="checkbox" name="testspeed" value="1" onClick="if(this.checked==true){document.AddLink.WebUrl.value='测速'}else{document.AddLink.WebUrl.value=''};" > 中转跳转:<input type="checkbox" name="out" value="1" onClick="if(this.checked==true){document.AddLink.WebUrl.value='/out.asp?turl='+document.AddLink.WebUrl.value}else{var aaa=document.AddLink.WebUrl.value;document.AddLink.WebUrl.value=aaa.replace('/out.asp?turl=','');}"> </font></td>
    </tr>
    <tr class="tdbg">
      <td width="160"  align="right"><strong>字体颜色:</strong></td>
      <td valign="middle">
          <INPUT TYPE="text" NAME="select_color" class="Input" ID="select_color" size="7" maxlength="7" value="<%=rsLink("Font_Color")%>">&nbsp;<img LANGUAGE="javascript" onClick="foreColor()"  WIDTH="18" HEIGHT="18" id="color_img" border=0 src="images/rect.gif" style="cursor:pointer;background-color:#ffffff">      </td>
    </tr>
    <tr class="tdbg">
      <td width="160" height="25"  align="right" valign="middle"><strong>推荐:</strong></td>
      <td height="25"><p>
          <input type="checkbox" name="Recommend"  value="1" <% if rsLink("Is_Recommend")=1 then response.write "checked" end if%>>
      </p></td>
    </tr>
    <tr class="tdbg">
      <td width="160" height="25"  align="right" valign="middle"><strong>样式:</strong></td>
      <td height="25"><input id="RadioButton1" type="radio" name="num1" value="1" />
          <img src="/Images/1.gif" />
          <input id="RadioButton2" type="radio" name="num1" value="2" />
          <img src="/Images/2.gif" />
          <input id="RadioButton3" type="radio" name="num1" value="3" />
          <img src="/Images/3.gif" />
          <input id="RadioButton4" type="radio" name="num1" value="4" />
          <img src="/Images/4.gif" />
          <input id="RadioButton5" type="radio" name="num1" value="5" />
          <img src="/Images/5.gif" />
          <input id="RadioButton6" type="radio" name="num1" value="0" checked="checked" />
        <label for="RadioButton6">无</label>
        &nbsp;</td>
    </tr>
 
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"><input name="ID" type="hidden" id="ID" value="<%=rsLink("ID")%>">  
        <input type="submit" class="Button" value=" 确 定 " name="cmdOk"> &nbsp; <input type="reset" class="Button" value=" 重 填 " name="cmdReset">      </td>
    </tr></form>
	   <tr class="tdbg">
      <td height="25" colspan="2"  align="center" valign="middle">下属网站	<a href='Admin_Web_GoodURL_Manage.asp?Action=add_lite&ID=<%=rslink("ID")%>' class="STYLE1">增加下属</a> </td>
    </tr>
		<%

        sql="select * from WebGoodSite_Url_lite where class_ID="&rsLink("ID")
	set rsxx=server.createobject("adodb.recordset")
	rsxx.open sql,conn,1,1
        if rsxx.eof and rsxx.bof then
	else
        i=1
        do while not rsxx.eof
		%>
<form action="Admin_Web_GoodURL_Manage.asp" method="post">
    <tr class="tdbg">
	
      <td width="160" height="25"  align="right" valign="middle"><input name="SiteName" class="Input" type="text" value="<%=rsxx("Site_Title")%>"></td>
      <td height="25"><input name="WebUrl" class="Input" size="60" id="xsWebUrl<%=i%>" value="<%=rsxx("Site_Url")%>">  中转跳转:<input type="checkbox" value="1" onClick="tsp(this,'<%=i%>');">
<input name="Action" type="hidden" id="Action" value="SaveModify_lite"><input name="ID" type="hidden" id="ID" value="<%=rsxx("ID")%>">  
   <input type="submit" class="Button" value=" 确 定 " name="cmdOk">
<%response.write "<a href='Admin_Web_GoodURL_Manage.asp?Action=Del_lite&ID=" & rsxx("ID") & "' onclick='return ConfirmDel();'>删除</a>"%></td>
    </tr>
</form>
	
	              <%
		Rsxx.movenext
		i=i+1
              loop
              Rsxx.close
	      set Rsxx=nothing
		  end if%>
</table>

<%
     rsLink.close
     set rsLink=nothing
end sub

%></body>
</html>
<%
sub SaveAdd()
	dim SelectType,SelectTypea,WebName,WebUrl,WebColor,Recommend,bold
        SelectType=trim(request("SelectType"))
        SelectTypea=trim(request("SelectTypea"))
        WebName=trim(request("SiteName"))
        WebUrl=trim(request("WebUrl"))
        WebColor=trim(request("select_color"))
        Recommend=int(request("Recommend"))
		bold=int(request("bold"))
		img=trim(request("num1"))

	if SelectType="0" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请选择一级类别!</li>"
	end if
	if SelectTypea="0" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请选择二级类别!</li>"
	end if
	if WebName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入网站名称!</li>"
	end if
	if WebUrl="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入网站链接!</li>"
	end if
	if WebColor="#NaNNaNNaN" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>字体颜色是非颜色字符！</li>"
	end if
	if WebColor="" then
                WebColor="#000000"
	end if
        if Recommend="1" then
           Recommend=1
        else
           Recommend=0
        end if 
		if bold="1" then
           bold=1
        else
           bold=0
        end if
	if FoundErr=True then
		response.write ErrMsg
		exit sub
	end if

	if FoundErr<>True then
		dim sqlLink,rsLink
		sqlLink="select top 1 * from WebGoodSite_Url where Site_Title='"&WebName&"'"
		set rsLink=Server.CreateObject("Adodb.RecordSet")
		rsLink.open sqlLink,conn,1,3
'		if rsLink.bof or rsLink.eof then
			rsLink.Addnew
			rsLink("Class_ID")=SelectTypea
			rsLink("Site_Url")=dvHtmlEncode(WebUrl)
			rsLink("Site_Title")=dvHtmlEncode(WebName)
			rsLink("Font_Color")=dvHtmlEncode(WebColor)
			rsLink("Is_Recommend")=Recommend
			rsLink("bold")=bold
			rsLink("img")=img
			rsLink("firstWord")=getpychar(BigToGb(left(dvHtmlEncode(WebName),1)))
			rsLink.update
			myid=rsLink("id")
			rsLink.close
			set rsLink=nothing

			'Response.Redirect "Admin_Web_GoodURL_Manage.asp?Action=Add&Select1="&SelectType&"&Select2="&SelectTypea
			'
			
			tmp_xsweburl=split(replace(request("xsweburl")," ",""),",")
			tmp_xssitename=split(replace(request("xssitename")," ",""),",")
			
			For tmp_i = LBound(tmp_xssitename) To UBound(tmp_xssitename)
			
			if tmp_xssitename(tmp_i)<>"" then
			
				sqlLink="select * from WebGoodSite_Url_lite"
				set rsLink=Server.CreateObject("Adodb.RecordSet")
				rsLink.open sqlLink,conn,1,3
					rsLink.Addnew
					rsLink("Class_ID")=myid
					rsLink("Site_Url")=dvHtmlEncode(tmp_xsweburl(tmp_i))
					rsLink("Site_Title")=dvHtmlEncode(tmp_xssitename(tmp_i))
					rsLink("Font_Color")="#000000"
					rsLink.update
					rsLink.close
					set rsLink=nothing
					
			'response.write dvHtmlEncode(tmp_xsweburl(tmp_i))&"*+*"&dvHtmlEncode(tmp_xssitename(tmp_i))&"<br>"	
			end if
			next
			
			
			
			
			
			
			
			call CloseConn()
			
			
			response.Redirect "Admin_Web_GoodURL_Manage.asp?Action=Modify&ID="&myid
	end if
end sub


sub SaveAddMuch()
	dim SelectType,SelectTypea,WebName,WebUrl,WebColor,Recommend
        SelectType=trim(request("SelectType"))
        SelectTypea=trim(request("SelectTypea"))
        WebName=trim(request("SiteName"))
        WebUrl=trim(request("WebUrl"))
        WebColor=trim(request("select_color"))
        Recommend=trim(request("Recommend"))
		
	if SelectType="0" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请选择一级类别!</li>"
	end if
	if SelectTypea="0" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请选择二级类别!</li>"
	end if

	if WebColor="" then
                WebColor="#000000"
	end if
        if Recommend="1" then
           Recommend=1
        else
           Recommend=0
        end if
	if FoundErr=True then
		exit sub
	end if

	if FoundErr<>True then
		dim sqlLink,rsLink
		sqlLink="select top 1 * from WebGoodSite_Url"
		set rsLink=Server.CreateObject("Adodb.RecordSet")
		rsLink.open sqlLink,conn,1,3
		tmp_WebName=split(WebName,",")
		tmp_WebUrl=split(WebUrl,",")
		for tmp_i=0 to UBound(tmp_WebName)
			if trim(tmp_WebName(tmp_i))<>"" then
			rsLink.Addnew
			rsLink("Class_ID")=SelectTypea
			rsLink("Site_Url")=dvHtmlEncode(trim(tmp_WebUrl(tmp_i)))
			rsLink("Site_Title")=dvHtmlEncode(trim(tmp_WebName(tmp_i)))
			rsLink("Font_Color")=dvHtmlEncode(WebColor)
			rsLink("Is_Recommend")=Recommend
			rsLink("img")=img
			rsLink.update
			end if
		next
			rsLink.close
			set rsLink=nothing

			
			
			call CloseConn()
			
			Response.Redirect "Admin_Web_GoodURL_Manage.asp?Action=AddMuch&Select1="&SelectType&"&Select2="&SelectTypea

	end if
end sub


sub SaveAdd_lite()


        class_id=trim(request("class_id"))
        WebName=trim(request("SiteName"))
        WebUrl=trim(request("WebUrl"))
        WebColor=trim(request("select_color"))
        Recommend=trim(request("Recommend"))

	if WebName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入网站名称!</li>"
	end if
	if WebUrl="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入网站链接!</li>"
        else
           if InStr(WebUrl,"http://")=0 and InStr(WebUrl,"https://")=0 then
              WebUrl="http://"&WebUrl
           end if
	end if
	if WebColor="#NaNNaNNaN" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>字体颜色是非颜色字符！</li>"
	end if
	if WebColor="" then
                WebColor="#000000"
	end if
        if Recommend="1" then
           Recommend=1
        else
           Recommend=0
        end if
	if FoundErr=True then
		exit sub
	end if

	if FoundErr<>True then
		dim sqlLink,rsLink
		sqlLink="select * from WebGoodSite_Url_lite"
		set rsLink=Server.CreateObject("Adodb.RecordSet")
		rsLink.open sqlLink,conn,1,3
'		if rsLink.bof or rsLink.eof then
			rsLink.Addnew
			rsLink("Class_ID")=class_id
			rsLink("Site_Url")=dvHtmlEncode(WebUrl)
			rsLink("Site_Title")=dvHtmlEncode(WebName)
			rsLink("Font_Color")=dvHtmlEncode(WebColor)
			rsLink("Is_Recommend")=Recommend
			rsLink.update
			rsLink.close
			set rsLink=nothing
			
		response.write "<script>history.go(-1);</script>"
		'			Response.Redirect "Admin_Web_GoodURL_Manage.asp?Action=Modify&ID="&class_id
'		end if
	end if
end sub

sub SaveModify()
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定网站ID</li>"
		exit sub
	else
		ID=Clng(ID)
	end if
	dim SelectType,SelectTypea,WebName,WebUrl,WebColor,Recommend,bold
        SelectType=trim(request("SelectType"))
        SelectTypea=trim(request("SelectTypea"))
        WebName=trim(request("SiteName"))
        WebUrl=trim(request("WebUrl"))
        WebColor=trim(request("select_color"))
        Recommend=int(request("Recommend"))
		bold=int(request("bold"))
		img=trim(request("num1"))
	if SelectType="0" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请选择一级类别!</li>"
	end if
	if SelectTypea="0" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请选择二级类别!</li>"
	end if
	if WebName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入网站名称!</li>"
	end if
	if WebUrl="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入网站链接!</li>"
	end if
	if WebColor="#NaNNaNNaN" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>字体颜色是非颜色字符！</li>"
	end if
	if WebColor="" then
                WebColor="#000000"
	end if
        if Recommend=1 then
           Recommend=1
        else
           Recommend=0
        end if
		if bold=1 then
           bold=1
        else
          bold=0
        end if
	if FoundErr=True then
		exit sub
	end if
	dim sqlLink,rsLink
	sqlLink="select * from WebGoodSite_Url where ID=" & ID
	set rsLink=Server.CreateObject("Adodb.RecordSet")
	rsLink.open sqlLink,conn,1,3
	if rsLink.bof and rsLink.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到网站!</li>"
	else
		rsLink("Class_ID")=SelectTypea
		rsLink("Site_Url")=dvHtmlEncode(WebUrl)
		rsLink("Site_Title")=dvHtmlEncode(WebName)
		rsLink("Font_Color")=dvHtmlEncode(WebColor)
		rsLink("Is_Recommend")=Recommend
		rsLink("bold")=bold
		rsLink("img")=img
		rsLink("firstWord")=getpychar(BigToGb(left(dvHtmlEncode(WebName),1)))
		rsLink.update
		rsLink.close
		set rsLink=nothing
		call CloseConn()
		Response.Redirect "Admin_Web_GoodURL_Manage.asp"
	end if
	rsLink.close
	set rsLink=nothing
end sub
sub SaveModify_lite()
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定网站ID</li>"
		exit sub
	else
		ID=Clng(ID)
	end if
        WebName=trim(request("SiteName"))
        WebUrl=trim(request("WebUrl"))
	if WebName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入网站名称!</li>"
	end if
	if WebUrl="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入网站链接!</li>"
        else
           if InStr(WebUrl,"http://")=0 and InStr(WebUrl,"https://")=0 then
              WebUrl="http://"&WebUrl
           end if
	end if

	if FoundErr=True then
		exit sub
	end if
	dim sqlLink,rsLink
	sqlLink="select * from WebGoodSite_Url_lite where ID=" & ID
	set rsLink=Server.CreateObject("Adodb.RecordSet")
	rsLink.open sqlLink,conn,1,3
	if rsLink.bof and rsLink.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到网站!</li>"
	else
		rsLink("Site_Url")=dvHtmlEncode(WebUrl)
		rsLink("Site_Title")=dvHtmlEncode(WebName)
		rsLink.update
		rsLink.close
		set rsLink=nothing
		call CloseConn()
		response.write "<script>history.go(-1);</script>"
	end if
	rsLink.close
	set rsLink=nothing
end sub
%>