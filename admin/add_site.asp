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
if trim(request("action"))="save" then

if trim(request("name"))="" or trim(request("url"))="" then
conn.close
set conn = nothing
response.Write "<script language='javascript'>alert('参数不能为空！');history.go(-1);</script>"
response.End
end if
set rs1=server.CreateObject("adodb.recordset")
rs1.open"select * from [kj]",conn,1,3
rs1.addnew
rs1("qs")=request("name")
rs1("hm")=replace(trim(request("url"))," ","")
rs1("xq")=trim(request("logo"))
rs1("rq")=request("rq")
rs1("folder")=request("folder")
rs1("xqqs")=request("xqqs")
myid=rs1("id")
rs1.update
rs1.close
set rs1=nothing
	
		On Error Resume Next
		JsFileName = Server.MapPath("/inc/kj.asp")
	set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
		.Open
		.Type = 2
		.Charset = "UTF-8"
		.WriteText "<" & "%" & "status=0" & "%" & ">"
		.SaveToFile JsFileName, 2
		.Close
	End With
	Set objStream = Nothing	
				
				

'response.write trim(request("logo"))&"++"&request.Form
response.Redirect("edit_site.asp?id="&myid)
response.End
end if%>

<BODY>
<DIV ALIGN="center">
  <form name="form1" method="post" action="?action=save">
    <table cellpadding="2" cellspacing="1" border="0" width="95%" align="center" class="border">
      <tr class="topbg">
        <td  height="22" colspan="2" align="center"><strong>添加</strong></td>
      </tr>
      <tr>
        <td height="23" class="tdbg">开奖日期<span style="color: #FF0000"></span></td>
        <td class="tdbg"><input name="rq" type="text" id="rq" value="<%=date()%>" size="40"></td>
      </tr>
      <tr>
        <td height="23" class="tdbg">期数</td>
        <td class="tdbg"><input name="name" type="text" id="name" size="40"></td>
      </tr>
      <tr>
        <td height="23" class="tdbg"><span class="style6">号码</span></td>
        <td class="tdbg"><input name="url" type="text" id="url" size="4" maxlength="2">
        <input name="url" type="text" id="url2" size="4" maxlength="2">
        <input name="url" type="text" id="url22" size="4" maxlength="2">
        <input name="url" type="text" id="url23" size="4" maxlength="2">
        <input name="url" type="text" id="url24" size="4" maxlength="2">
        <input name="url" type="text" id="url25" size="4" maxlength="2">
        +
        <input name="url" type="text" id="url26" size="4" maxlength="2"></td>
      </tr>
	        <tr>
        <td height="23" class="tdbg"><span class="tdbg">图片目录</span></td>
        <td class="tdbg"><select name="folder">
		<%
filepath="/images/"
Set fso = Server.CreateObject("Scripting.FileSystemObject")
Set fileobj = fso.GetFolder(server.mappath(filepath))
Set fsofolders = fileobj.SubFolders
Set fsofile = fileobj.Files


For Each folder in fsofolders
  'response.Write("<div class='filelist'><ul><li>"&folder.name&"</li><li>"&folder.size&"</li><li>"&folder.datelastmodified&"</li></ul></div>")
   
  response.Write "<option value='/images/"&folder.name&"'>/images/"&folder.name&"</option>"
Next 
%>
          
        </select>
        </td>
	        </tr>
	      <tr>
        <td height="23" class="tdbg"><span class="tdbg">下期期数</span></td>
        <td class="tdbg"><input name="xqqs" type="text" value="" size="20"></td>
      </tr>
      <tr>
        <td height="23" class="tdbg"><span class="tdbg">下期时间</span></td>
        <td class="tdbg"><input name="logo" type="text" id="logo" value="月日时分星期" size="60"></td>
      </tr>
      <tr>
        <td height="23" colspan="2" align="center" valign="middle" class="tdbg"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="177">&nbsp;</td>
            <td width="95"><input type="submit"  name="Submit" value="确定"></td>
            <td width="55">&nbsp;</td>
            <td width="957"><input type="reset"  name="Submit2" value="重写"></td>
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
