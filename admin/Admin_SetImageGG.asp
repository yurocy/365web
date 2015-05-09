<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2
Const CheckChannelID=0
Const PurviewLevel_Others="Admin_GG"
%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/md5.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<%
dim strFileName
const MaxPerPage=15
dim totalPut,CurrentPage,TotalPages
dim sql,rs,ID,str_select,str_text
dim Action,FoundErr,ErrMsg
Action=trim(request("Action"))
ID=Trim(Request("ID"))
str_select=trim(request("select_Query"))
str_text=trim(request("Query_text"))
strFileName="Admin_SetImageGG.asp?select_Query="&str_select&"&Query_text="&str_text

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if


if ID<>"" then
        if Action="Del" then
		conn.execute "Delete From admin_GG Where ID=" & CLng(ID)
	end if
end if
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<script LANGUAGE="javascript">
function ConfirmDel()
{
   if(confirm("确定要删除吗?"))
     return true;
   else
     return false;
}
function OpenUpFile(){
  if (document.AddLink.GG_Type.value=="图片")
  {
  var arr = showModalDialog("editor_InsertPic.asp?type=pic", "", "dialogWidth:31em; dialogHeight:10em; help: no; scroll: no; status: no");
  }
  else
  {
  var arr = showModalDialog("editor_InsertPic.asp?type=flash", "", "dialogWidth:31em; dialogHeight:10em; help: no; scroll: no; status: no");
  }
  if (arr != null){
    var ss=arr.split("$$$");
	if (ss[1]!="None")
	{
	 document.AddLink.TX_Image.value=ss[1];
	}
  }
  document.AddLink.TX_Image.focus();
}
</script>

</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg"> 
    <td height="22" colspan="2" align="center"><strong>广 告 管 理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航:</strong></td>
    <td height="30"> <a href="Admin_SetImageGG.asp?Action=list">管理首页</a></td>
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
	sql="select * from Admin_GG "
        if str_select="请选择" or str_select="" then
	sql=sql
        else
	sql=sql&" where ("&str_select&" like '%"&str_text&"%')"
        end if
	sql=sql&" order by id"
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
<form method="post" name="AddLink" action="Admin_SetImageGG.asp">
  <tr>
    <td width="76" align="right"><b>查询:</b>&nbsp;&nbsp;</td>
    <td><select size="1" name="select_Query">
        <option selected value="请选择">请选择</option>
        <option <%if str_select="LP1" then response.write "selected" end if%> value="LP1">地块名称</option>
        <option <%if str_select="LP2" then response.write "selected" end if%> value="LP2">地块编号</option>
        <option <%if str_select="LP3" then response.write "selected" end if%> value="LP3">用地位置</option>
        </select>&nbsp;<input name="Query_text" size="30" class="Input" maxlength="100" type="text" value="<%=str_text%>">&nbsp;<input class="Button" type="submit" value="查询" name="B3"></td>
  </tr>
  <tr>
    <td colspan="2" height="5"></td>
  </tr>
</form>
</table>

  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr class="title">
      <td width="100" height="22" align="center">栏目名称</td>
      <td width="100" height="22" align="center">广告位置</td>
      <td width="100" height="22" align="center">广告类别</td>
      <td width="100" height="22" align="center">宽</td>
      <td width="100" height="22" align="center">高</td>
      <td height="22" align="center">操作</td>
    </tr>
<%
do while not rs.eof
%>
    <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td width="100" align="center"><%=rs("ClassName")%></td>
      <td width="100" align="center"><%=rs("ClassSeat")%></td>
      <td width="100" align="center"><%=rs("GG_Type")%></td>
      <td width="100" align="center"><%=rs("Width")%></td>
      <td width="100" align="center"><%=rs("Height")%></td>
      <td align="center">
<%
      response.write "<a href='Admin_SetImageGG.asp?Action=Modify&ID=" & rs("ID") & "'>修改</a>"
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

sub Modify()
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定ID</li>"
		exit sub
	else
		ID=Clng(ID)
	end if
	dim sqlLink,rsLink
	sqlLink="select * from Admin_GG where ID=" & ID
	set rsLink=Server.CreateObject("Adodb.RecordSet")
	rsLink.open sqlLink,conn,1,3
	if rsLink.bof and rsLink.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到相关信息！</li>"
		rsLink.close
		set rsLink=nothing
		exit sub
	end if

%>
<form method="post" name="AddLink" onsubmit="return Check()" action="Admin_SetImageGG.asp">
  <table border="0" cellpadding="2" cellspacing="1" align="center" width="100%" class="border">
    <tr class="title"> 
      <td height="22" colspan="2" align="center"><strong>修改广告</strong></td>
    </tr>
    <tr class="tdbg">
      <td width="300" height="25" align="right"><strong>栏目名称:</strong></td>
      <td height="25"><%=rsLink("ClassName")%></td>
    </tr>
    <tr class="tdbg">
      <td width="300" height="25" align="right"><strong>广告位置:</strong></td>
      <td height="25"><%=rsLink("ClassSeat")%></td>
    </tr>
    <tr class="tdbg">
      <td width="300" height="25" align="right"><strong>宽:</strong></td>
      <td height="25"><%=rsLink("width")%></td>
    </tr>
    <tr class="tdbg">
      <td width="300" height="25" align="right"><strong>高:</strong></td>
      <td height="25"><%=rsLink("height")%></td>
    </tr>
    <tr class="tdbg">
      <td width="300" height="25" align="right"><strong>广告类别:</strong></td>
      <td height="25"><select size="1" name="GG_Type">
	<option value="图片" <%if rsLink("gg_type")="图片" then response.write "selected" end if%>>图片</option>
	<option value="FLASH" <%if rsLink("gg_type")="FLASH" then response.write "selected" end if%>>FLASH</option>
        <option value="外部代码" <%if rsLink("gg_type")="外部代码" then response.write "selected" end if%>>外部代码</option>
	</select></td>
    </tr>
    <tr class="tdbg">
      <td width="300" height="25" align="right"><strong>广告:</strong></td>
      <td height="25"><input name="TX_Image" type="text"  class="Input" id="TX_Image" size="40" value="<%if rsLink("gg_type")<>"外部代码" then response.write rsLink("ClassImages") end if%>" maxlength="255"><font color="#FF0000">*</font><input name="Cancel" type="button" id="Cancel" value="上传广告" class="button" onclick="OpenUpFile();" style="cursor:hand;"> 
      </td>
    </tr>
    <tr class="tdbg">
      <td width="300" height="25" align="right"><strong>链接:</strong></td>
      <td height="25"><input name="GG_URL" type="text"  class="Input" id="GG_URL" size="40" maxlength="255" value="<%=rsLink("GG_URL")%>"><font color="#FF0000">请加上http://</font>
      </td>
    </tr>
    <tr class="tdbg">
       <td width="300" align="right"><strong>外部代码:</strong></td>
      <td height="25"><textarea name="TX_Text" class="Input1" cols="55" rows="4" id="TX_Text"><%if rsLink("gg_type")="外部代码" then response.write rsLink("ClassImages") end if%></textarea>
      </td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="ID" type="hidden" id="ID" value="<%=rsLink("ID")%>"><input name="width" type="hidden" id="width" value="<%=rsLink("width")%>"><input name="height" type="hidden" id="height" value="<%=rsLink("height")%>">
        <input name="Action" class="button" type="hidden" id="Action" value="SaveModify"> <input class="button" type="submit" value=" 确 定 " name="cmdOk"> 
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
sub SaveModify()
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定ID</li>"
		exit sub
	else
		ID=Clng(ID)
	end if
	dim TX_Image,GG_Type,GG_URL
        GG_Type=trim(request("GG_Type"))
        if GG_Type="外部代码" then
        TX_Image=trim(request("TX_Text"))
        else
	TX_Image=trim(request("TX_Image"))
        end if
        GG_URL=trim(request("GG_URL"))
	if TX_Image="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>必须输入广告</li>"
	end if
	if FoundErr=True then
		exit sub
	end if

	dim sqlLink,rsLink
	sqlLink="select * from Admin_GG where ID=" & ID
	set rsLink=Server.CreateObject("Adodb.RecordSet")
	rsLink.open sqlLink,conn,1,3

	if rsLink.bof and rsLink.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到广告！</li>"
	else
		rsLink("GG_Type")=GG_Type
		rsLink("ClassImages")=TX_Image
		rsLink("GG_URL")=GG_URL
		rsLink.update
		rsLink.close
		set rsLink=nothing
		call CloseConn()
		Response.Redirect "Admin_SetImageGG.asp"
	end if
	rsLink.close
	set rsLink=nothing
end sub
%>
