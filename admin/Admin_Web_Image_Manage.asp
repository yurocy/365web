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
strFileName="Admin_Web_Image_Manage.asp?select_Query="&str_select&"&Query_text="&str_text

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

if ID<>"" then
	if Action="Del" then
		conn.execute "Delete From WebSite_Image Where ID=" & CLng(ID)
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
	  alert("������ͼƬ����!")
	  document.AddLink.SiteName.focus()
	  return false
	 }
if (document.AddLink.WebUrl.value=="")
	{
	  alert("������ͼƬ����!")
	  document.AddLink.WebUrl.focus()
	  return false
	 }
}
function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ����?"))
     return true;
   else
     return false;
}
</script>
<script language="JavaScript">

// �ļ���Դ��ѡ����¼�
function RadioClick(what){
	if (what=="url"){
		d_checkfromfile.checked=false;
		d_fromurl.disabled=false;
		d_checkfromurl.checked=true;
		d_file.myform.uploadfile.disabled=true;
	}else{
		d_checkfromurl.checked=false;
		d_file.myform.uploadfile.disabled=false;
		d_checkfromfile.checked=true;
		d_fromurl.disabled=true;
	}
}

// �ϴ�֡�������ʱִ��
function UploadLoaded(){
	// ��ʼradio
	RadioClick('file');
}

// �ϴ�����
function UploadError(sErrDesc){
	AbleItems();
	RadioClick('file');
	divProcessing.style.display="none";
	try {
		BaseAlert(d_file.myform.uploadfile,sErrDesc);
	}
	catch(e){}
}

// �ļ��ϴ����ʱִ��,�����ϴ��ļ���
function UploadSaved(sPathFileName){
	d_fromurl.value = sPathFileName;
	ReturnValue();
}

// �����ڷ���ֵ
function ReturnValue(){
	var url = d_fromurl.value;
	var sFilePic = getFilePic(url);
	var sPicSrc = relativePath2setPath("sysimage/file/"+sFilePic);
	dialogArguments.insertHTML("<img id=eWebEditor_TempElement_Img border=0 src='"+sPicSrc+"'><a id=eWebEditor_TempElement_Href href='"+url+"' target=_blank>"+d_filename.value+"</a>");

	var oTempElement = dialogArguments.eWebEditor.document.getElementById("eWebEditor_TempElement_Img");
	oTempElement.src = sPicSrc;
	oTempElement.removeAttribute("id");
	
	oTempElement = dialogArguments.eWebEditor.document.getElementById("eWebEditor_TempElement_Href");
	oTempElement.href = url;
	oTempElement.removeAttribute("id");
	
	window.returnValue = null;
	window.close();
}

// ��ȷ��ʱִ��
function ok(){
	if (d_checkfromurl.checked){
		// ȡ�ļ���
		GetFileName(d_fromurl.value, "/");
		// ����ֵ
		ReturnValue();
	}else{
		// �ϴ��ļ��ж�
		if (!d_file.CheckUploadForm()) return false;
		// ȡ�ļ���
		GetFileName(d_file.myform.uploadfile.value, "\\");
		// ʹ���������Ч
		DisableItems();
		// ��ʾ�����ϴ��ļ�
		divProcessing.style.display="";
		// �ϴ����ύ
		d_file.myform.submit();
	}
}

// ʹ�����������Ч
function DisableItems(){
	d_checkfromfile.disabled=true;
	d_checkfromurl.disabled=true;
	d_fromurl.disabled=true;
	Ok.disabled=true;
}

// ʹ�����������Ч
function AbleItems(){
	d_checkfromfile.disabled=false;
	d_checkfromurl.disabled=false;
	d_fromurl.disabled=false;
	Ok.disabled=false;
}

// ���ļ���չ��ȡͼ������������
function getFilePic(url){
	var sExt;
	sExt=url.substr(url.lastIndexOf(".")+1);
	sExt=sExt.toUpperCase();
	var sPicName;
	switch(sExt){
	case "TXT":
		sPicName = "txt.gif";
		break;
	case "CHM":
	case "HLP":
		sPicName = "hlp.gif";
		break;
	case "DOC":
		sPicName = "doc.gif";
		break;
	case "PDF":
		sPicName = "pdf.gif";
		break;
	case "MDB":
		sPicName = "mdb.gif";
		break;
	case "GIF":
		sPicName = "gif.gif";
		break;
	case "JPG":
		sPicName = "jpg.gif";
		break;
	case "BMP":
		sPicName = "bmp.gif";
		break;
	case "PNG":
		sPicName = "pic.gif";
		break;
	case "ASP":
	case "JSP":
	case "JS":
	case "PHP":
	case "PHP3":
	case "ASPX":
		sPicName = "code.gif";
		break;
	case "HTM":
	case "HTML":
	case "SHTML":
		sPicName = "htm.gif";
		break;
	case "ZIP":
		sPicName = "zip.gif";
		break;
	case "RAR":
		sPicName = "rar.gif";
		break;
	case "EXE":
		sPicName = "exe.gif";
		break;
	case "AVI":
		sPicName = "avi.gif";
		break;
	case "MPG":
	case "MPEG":
	case "ASF":
		sPicName = "mp.gif";
		break;
	case "RA":
	case "RM":
		sPicName = "rm.gif";
		break;
	case "MP3":
		sPicName = "mp3.gif";
		break;
	case "MID":
	case "MIDI":
		sPicName = "mid.gif";
		break;
	case "WAV":
		sPicName = "audio.gif";
		break;
	case "XLS":
		sPicName = "xls.gif";
		break;
	case "PPT":
	case "PPS":
		sPicName = "ppt.gif";
		break;
	case "SWF":
		sPicName = "swf.gif";
		break;
	default:
		sPicName = "unknow.gif";
		break;
	}
	return sPicName;

}

// ȡ�ļ��������ص������
function GetFileName(url, opt){
	d_filename.value=url.substr(url.lastIndexOf(opt)+1);
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
<div id=divProcessing style="width:200px;height:30px;position:absolute;left:70px;top:30px;display:none">
<table border=0 cellpadding=0 cellspacing=1 bgcolor="#000000" width="100%" height="100%"><tr><td bgcolor=#3A6EA5><marquee align="middle" behavior="alternate" scrollamount="5"><font color=#FFFFFF>...�ļ��ϴ���...��ȴ�...</font></marquee></td></tr></table>
</div>
<input type=hidden id=d_filename value="">

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg">
    <td height="22" colspan="2" align="center"><strong>�� �� ͼ Ƭ �� ��</strong></td>
  </tr>
  <tr class="tdbg">
    <td width="70" height="30"><strong>��������</strong></td>
    <td height="30"> <a href="Admin_Web_Image_Manage.asp?Action=list">ͼƬ����</a> | <a href="Admin_Web_Image_Manage.asp?Action=Add">���ͼƬ</a></td>
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
	sql="select * from WebSite_Image "
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
sub List
   	dim i
    i=0
%>
<table border="0" cellpadding="0" cellspacing="0">
<form method="post" name="AddLink" action="Admin_Web_Image_Manage.asp">
  <tr>
    <td width="76" align="right"><b>��ѯ:</b>&nbsp;&nbsp;</td>
    <td><select size="1" name="select_Query">
        <option selected value="��ѡ��">��ѡ��</option>
        <option <%if str_select="Site_Title" then response.write "selected" end if%> value="Site_Title">ͼƬ����</option>
        </select>&nbsp;<input name="Query_text" size="30" class="Input" maxlength="100" type="text" value="<%=str_text%>">&nbsp;<input class="Button" type="submit" value="��ѯ" name="B3"></td>
  </tr>
  <tr>
    <td colspan="2" height="5"></td>
  </tr>
</form>
</table>
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr class="title">
      <td width="200" height="22" align="center">ͼƬ����</td>
      <td height="22" align="center">����</td>
    </tr>
<%
do while not rs.eof
%>
    <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td width="200" align="center"><a href="<%=rs("Site_Url")%>" target="_blank" style="color: <%=rs("Font_Color")%>"><%=rs("Site_Title")%></a></td>
      <td align="center">
<%
      response.write "<a href='Admin_Web_Image_Manage.asp?Action=Modify&ID=" & rs("ID") & "'>�޸�</a> | "
      response.write "<a href='Admin_Web_Image_Manage.asp?Action=Del&ID=" & rs("ID") & "' onclick='return ConfirmDel();'>ɾ��</a>"
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
<form method="post" name="AddLink" onsubmit="return Check()" action="Admin_Web_Image_Manage.asp">
  <table border="0" cellpadding="2" cellspacing="1" align="center" width="100%" class="border">
    <tr class="title">
      <td height="22" colspan="2" align="center"><strong>���ͼƬ</strong></td>
    </tr>
    <tr class="tdbg">
      <td width="300" height="25" valign="middle"  align="right"><strong>������Ŀ:</strong></td>
      <td height="25"><select size="1" name="SelectClass">
    <option value="��ѡ��" selected>��ѡ��</option>
    <option value="��ҳ����">��ҳ����</option>
    <option value="��ҳ����">��ҳ����</option>
  </select>
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr class="tdbg">
      <td width="300" height="25" valign="middle"  align="right"><strong>ͼƬ�ļ�:</strong></td>
      <td height="25">
       <iframe id=d_file frameborder=0 src="editor/upload.asp?type=file" width="100%" height="22" scrolling=no></iframe>
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr class="tdbg">
      <td width="300" height="25" valign="middle"  align="right"><strong>ͼƬ����:</strong></td>
      <td height="25"> <input name="WebUrl" class="Input" size="30"  maxlength="100">
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr class="tdbg">
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd"> 
        <input type="submit" class="Button" value=" ȷ �� " name="cmdOk" onclick="ok()"> &nbsp; <input type="reset" class="Button" value=" �� �� " name="cmdReset"> 
      </td>
    </tr>
  </table>
</form>
<%
end sub

sub Modify()
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��ID</li>"
		exit sub
	else
		ID=Clng(ID)
	end if
	dim sqlLink,rsLink,ClassID
	sqlLink="select * from WebSite_Image where ID=" & ID
	set rsLink=Server.CreateObject("Adodb.RecordSet")
	rsLink.open sqlLink,conn,1,3
	if rsLink.bof and rsLink.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ID!</li>"
		rsLink.close
		set rsLink=nothing
		exit sub
	end if
%>
<form method="post" name="AddLink" onsubmit="return Check()" action="Admin_Web_Image_Manage.asp">
  <table border="0" cellpadding="2" cellspacing="1" align="center" width="100%" class="border">
    <tr class="title">
      <td height="22" colspan="2" align="center"><strong>�޸�ͼƬ</strong></td>
    </tr>
    <tr class="tdbg">
      <td width="300" height="25" valign="middle"  align="right"><strong>ͼƬ����:</strong></td>
      <td height="25"> <input name="SiteName" class="Input" size="30"  maxlength="50" value="<%=rsLink("Site_Title")%>">
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr class="tdbg">
      <td width="300" height="25" valign="middle"  align="right"><strong>ͼƬ����:</strong></td>
      <td height="25"> <input name="WebUrl" class="Input" size="30"  maxlength="100" value="<%=rsLink("Site_Url")%>">
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"><input name="ID" type="hidden" id="ID" value="<%=rsLink("ID")%>">  
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
sub SaveAdd()
	dim WebName,WebUrl
        WebName=trim(request("SiteName"))
        WebUrl=trim(request("WebUrl"))
	if WebName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>������ͼƬ����!</li>"
	end if
	if WebUrl="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>������ͼƬ����!</li>"
        else
           if InStr(WebUrl,"http://")=0 then
              WebUrl="http://"&WebUrl
           end if
	end if
	if FoundErr=True then
		exit sub
	end if
	if FoundErr<>True then
		dim sqlLink,rsLink
		sqlLink="select top 1 * from WebSite_Image where Site_Title='" & dvHtmlEncode(WebName) & "'"
		set rsLink=Server.CreateObject("Adodb.RecordSet")
		rsLink.open sqlLink,conn,1,3
		if not (rsLink.bof and rsLink.eof) then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��Ҫ��ӵ�ͼƬ�Ѿ����ڣ�</li>"
		else
			rsLink.Addnew
			rsLink("Site_Url")=dvHtmlEncode(WebUrl)
			rsLink("Site_Title")=dvHtmlEncode(WebName)
			rsLink("Font_Color")=dvHtmlEncode(WebColor)
			rsLink.update
			rsLink.close
			set rsLink=nothing
			call CloseConn()
			Response.Redirect "Admin_Web_Image_Manage.asp?Action=Add"
		end if
		rsLink.close
		set rsLink=nothing
	end if
end sub

sub SaveModify()
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��ͼƬID</li>"
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
		ErrMsg=ErrMsg & "<br><li>������ͼƬ����!</li>"
	end if
	if WebUrl="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>������ͼƬ����!</li>"
        else
           if InStr(WebUrl,"http://")=0 then
              WebUrl="http://"&WebUrl
           end if
	end if
	if FoundErr=True then
		exit sub
	end if
	dim sqlLink,rsLink
	sqlLink="select * from WebSite_Image where ID=" & ID
	set rsLink=Server.CreateObject("Adodb.RecordSet")
	rsLink.open sqlLink,conn,1,3
	if rsLink.bof and rsLink.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ�������!</li>"
	else
		rsLink("Site_Url")=dvHtmlEncode(WebUrl)
		rsLink("Site_Title")=dvHtmlEncode(WebName)
		rsLink.update
		rsLink.close
		set rsLink=nothing
		call CloseConn()
		Response.Redirect "Admin_Web_Image_Manage.asp"
	end if
	rsLink.close
	set rsLink=nothing
end sub
%>