<!--#include file="../Inc/manageconfig.asp"-->
<html>
<head>
<title>�ϴ��ļ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<SCRIPT language=javascript>
function check() 
{
	var strFileName=document.form1.FileName.value;
	if (strFileName=="")
	{
    	alert("��ѡ��Ҫ�ϴ����ļ�");
		document.form1.FileName.focus();
    	return false;
  	}
}
</SCRIPT>
</head>
<body bgColor=menu leftmargin="5" topmargin="0">
<%
if EnableUploadFile="Yes" and session("AdminName")<>"" then
%>
<form action="Upfile_Dialog.asp" method="post" name="form1" onSubmit="return check()" enctype="multipart/form-data">
  <input name="FileName" type="FILE" class="Input" size="35">
  <input type="submit" name="Submit" class="button" value="�ϴ�">
  <input name="DialogType" type="hidden" id="DialogType" value="<%=trim(request("DialogType"))%>">
</form>
<%
else
	response.write "�Բ��𣬱���վ�������ϴ��ļ���"
end if
%>
</body>
</html>
