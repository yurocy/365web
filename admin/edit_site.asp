<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/md5.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<html>
<head>
<title>网址导航管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<%
dim id
id=trim(request("id"))

if trim(request("status"))<>"" then 

		On Error Resume Next
		JsFileName = Server.MapPath("/inc/kj.asp")
	set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
		.Open
		.Type = 2
		.Charset = "UTF-8"
		.WriteText "<" & "%" & "status=" & trim(request("status")) & "%" & ">"
		.SaveToFile JsFileName, 2
		.Close
	End With
	Set objStream = Nothing	



response.Write "<script language='javascript'>history.go(-1);</script>"




end if 



if trim(request("action"))="save" then
if trim(request("name"))="" or trim(request("url"))="" then
conn.close
set conn = nothing
response.Write "<script language='javascript'>alert('参数不能为空！');history.go(-1);</script>"
response.End
end if
set rs1=server.CreateObject("adodb.recordset")
rs1.open"select * from [kj] where id="&id,conn,1,3
rs1("qs")=request("name")
rs1("hm")=replace(trim(request("url"))," ","")
rs1("xq")=trim(request("logo"))
rs1("rq")=request("rq")
rs1("folder")=request("folder")
rs1("xqqs")=request("xqqs")
rs1.update
rs1.close
set rs1=nothing
response.write "<script language=javascript>history.go(-1);</script>"
response.End
end if
%>
<BODY>
<DIV ALIGN="center">

<%set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [kj] where id="&id,conn,1,1
%>

  <form name="form1" method="post" action="?action=save&id=<%=id%>">

    <table cellpadding="2" cellspacing="1" border="0" width="95%" align="center" class="border">
      <tr class="topbg">
        <td height="25"  colspan="2" align="center" class="topbg">修改开奖信息</td>
      </tr>
      <tr class="tr_southidc">
        <td class="tdbg" colspan="2" height="25" align="center">修改状态: <a href="?status=2">节目广告</a> <a href="?status=3">主持人解说</a>  <a href="?status=1">搅珠中</a> </td>
      </tr>
      <tr class="tr_southidc">
        <td height="23" class="tdbg"><span class="tdbg">开奖日期</span></td>
        <td class="tdbg"><input name="rq" type="text" id="logo" value="<%=rs("rq")%>" size="40"></td>
      </tr>
      <tr class="tr_southidc">
        <td width="15%" height="23" class="tdbg">期数<span style="color: #FF0000">*</span></td>
        <td width="80%" class="tdbg"><input name="name" type="text" id="name" value="<%=rs("qs")%>" size="40"></td>
      </tr>
<%haoma=split(rs("hm"),",")%>
      <tr class="tr_southidc">
        <td height="23" class="tdbg"><span class="style6">号码</span></td>
        <td class="tdbg"><input name="url" type="text" id="url" size="4" maxlength="2" value="<%=haoma(0)%>">
        <input name="url" type="text" id="url2" size="4" maxlength="2" value="<%=haoma(1)%>">
        <input name="url" type="text" id="url22" size="4" maxlength="2" value="<%=haoma(2)%>">
        <input name="url" type="text" id="url23" size="4" maxlength="2" value="<%=haoma(3)%>">
        <input name="url" type="text" id="url24" size="4" maxlength="2" value="<%=haoma(4)%>">
        <input name="url" type="text" id="url25" size="4" maxlength="2" value="<%=haoma(5)%>">
        +
        <input name="url" type="text" id="url26" size="4" maxlength="2" value="<%=haoma(6)%>"></td>
      </tr>
	  	      <tr>
        <td height="23" class="tdbg"><span class="tdbg">下期期数</span></td>
        <td class="tdbg"><input name="xqqs" type="text" value="<%=rs("xqqs")%>" size="20"></td>
      </tr>
      <tr class="tr_southidc">
        <td height="23"><span class="style6">下期时间</span></td>
        <td><input name="logo" type="text" id="logo" value="<%=rs("xq")%>" size="60"></td>
      </tr>     
	   <tr class="tr_southidc">
        <td height="23" class="tdbg"><span class="style6">图片目录</span></td>
        <td class="tdbg"><select name="folder">
		<%
filepath="/images/"
Set fso = Server.CreateObject("Scripting.FileSystemObject")
Set fileobj = fso.GetFolder(server.mappath(filepath))
Set fsofolders = fileobj.SubFolders
Set fsofile = fileobj.Files

response.Write "<option value='"&rs("folder")&"'>"&rs("folder")&"</option>"

For Each folder in fsofolders
  'response.Write("<div class='filelist'><ul><li>"&folder.name&"</li><li>"&folder.size&"</li><li>"&folder.datelastmodified&"</li></ul></div>")
   
  response.Write "<option value='/images/"&folder.name&"'>/images/"&folder.name&"</option>"
Next 
%>
          
        </select></td>
      </tr>
      <tr class="tr_southidc">
        <td height="23" colspan="2" align="center" valign="middle" class="tdbg"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="234">&nbsp;</td>
            <td width="69"><input type="submit"  name="Submit" value="确定修改"></td>
            <td width="103">&nbsp;</td>
            <td width="499"><input type="reset"  name="Submit2" value="取消重写"></td>
          </tr>
        </table></td>
      </tr>
    </table>
  </form>
  <p>&nbsp;</p>
</DIV>
<%htmlend%>
</BODY>
</HTML>
