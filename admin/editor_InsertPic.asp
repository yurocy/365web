<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
dim u_type,u_showtype
u_type=trim(request("type"))
if u_type<>"flash" then
u_type="pic"
end if
if u_type="pic" then
u_showtype="图片"
else
u_showtype="Flash"
end if
%>
<HTML>
<HEAD>
<TITLE>上传<%=u_showtype%>文件</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<script language="JavaScript">
function OK(){
  var str1="";
  var strurl=document.form1.url.value;
  if (strurl==""||strurl=="http://")
  {
  	alert("请先输入<%=u_showtype%>地址，或者上传<%=u_showtype%>！");
	document.form1.url.focus();
	return false;
  }
  else
  {
    window.returnValue = str1+"$$$"+document.form1.UpFileName.value;
    window.close();
  }
}
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
</head>
<BODY bgColor=menu topmargin=15 leftmargin=15 >
<form name="form1" method="post" action="">
<table width=100% border="0" align="center" cellpadding="0" cellspacing="2">
  <tr><td>
<FIELDSET align=left>
<LEGEND align=left>输入<%=u_showtype%>参数</LEGEND>
<table border="0" align=center cellpadding="0" cellspacing="3">
<tr>
  <td colspan="2"><%=u_showtype%>地址：
    <input name="url" id="url" value='http://' size=38 maxlength="200" class="Input">
</td>
  </tr><tr>
          <td> 说明文字：
            <input name="alttext" class="Input" id=alttext size=38 maxlength="100">
            </td>
          <td></td>
  </tr>
</table>
</fieldset>
</td><td width=80 align="center"><input name="cmdOK" class="button" type="button" id="cmdOK" value="  确定  " onClick="OK();">
  <br>
  <br>
<input name="cmdCancel" type=button id="cmdCancel" class="button" onclick="window.close();" value='  取消  '></td></tr>
  <tr>
    <td>
<FIELDSET align=left>
<LEGEND align=left>上传本地<%=u_showtype%></LEGEND>
<iframe class="TBGen" style="top:2px" ID="UploadFiles" src="upload_dialog.asp?DialogType=<%=u_type%>" frameborder=0 scrolling=no width="350" height="25"></iframe>
</fieldset>
	</td>
    <td align="center"><input name="UpFileName" type="hidden" id="UpFileName" value="None"></td>
  </tr>
</table>
</form>
</body>
</html>

