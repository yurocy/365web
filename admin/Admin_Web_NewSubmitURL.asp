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
strFileName="Admin_Web_NewSubmitURL.asp?select_Query="&str_select&"&Query_text="&str_text

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

if ID<>"" then
	if Action="Del" then
		conn.execute "Delete From WebSiteSubmit_URL Where ID=" & CLng(ID)
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
	  alert("��ѡ��һ�����!")
	  document.AddLink.SelectType.focus()
	  return false
	 }
if (document.AddLink.SelectTypea.value=="0")
	{
	  alert("��ѡ��������!")
	  document.AddLink.SelectTypea.focus()
	  return false
	 }
if (document.AddLink.SiteName.value=="")
	{
	  alert("��������վ����!")
	  document.AddLink.SiteName.focus()
	  return false
	 }
if (document.AddLink.WebUrl.value=="")
	{
	  alert("��������վ����!")
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
    <td height="22" colspan="2" align="center"><strong>�� �� �� վ �� ��</strong></td>
  </tr>

</table>

<%
if Action="Modify" then
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

if Request("action")="del" then
conn.execute "delete from [WebSiteSubmit_URL] where id in ("&request.form("selectdel")&")"
end if


	sql="select * from WebSiteSubmit_URL where is_ok=0 "
        if str_select="��ѡ��" or str_select="" then
	sql=sql
        else
	sql=sql&" and ("&str_select&" like '%"&str_text&"%')"
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
<form name="form1" method="post" action="?action=del">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border" style="WORD-WRAP: break-word;TABLE-LAYOUT: fixed;word-break:break-all">
    <tr class="title">
      <td width="28" height="22" align="center">ѡ��</td>
      <td width="10%" height="22" align="center">��վ����</td>
      <td width="20%" height="22" align="center">��վ��ַ</td>
      <!--    <td width="100" height="22" align="center">��ϵ</td>-->
      <td width="20%" height="22" align="center">����</td>
      <td width="20%" height="22" align="center">����</td>
      <td height="22" align="center">���ʱ��</td>
      <td width="10%" height="20" align="center">����</td>
    </tr>
    <%
do while not rs.eof
%>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="28" align="center"><input name="selectdel" type="checkbox" id="selectdel" value=<%=rs("id")%>></td>
      <td align="center"><%=rs("Sitename")%></td>
      <td><%=rs("SiteUrl")%></td>
      <td align="center"><%=rs("SiteContact")%></td>
      <td width="20%" align="center"><%=rs("chazhang")%></td>
      <td align="center"><%=rs("AddDateTime")%></td>
      <td align="center"><%
      response.write "<a href='Admin_Web_NewSubmitURL.asp?Action=Modify&ID=" & rs("ID") & "'>���</a>|<a href='Admin_Web_NewSubmitURL.asp?Action=Del&ID=" & rs("ID") & "' onclick='return ConfirmDel();'>ɾ��</a>"
	  %>
      </td>
    </tr>
    <%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td colspan="2" align="left"><input class="button1" type="submit" value="ɾ����ѡ" name="Submit1"></td>
      <td>&nbsp;</td>
      <td align="center">&nbsp;</td>
      <td width="20%" align="center">&nbsp;</td>
      <td align="center">&nbsp;</td>
      <td align="center">&nbsp;</td>
    </tr>
  </table>
</form>
<%
end sub

sub Modify()
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ�����ID</li>"
		exit sub
	else
		ID=Clng(ID)
	end if
	dim sqlLink,rsLink,ClassID
	sqlLink="select * from WebSiteSubmit_URL where ID=" & ID
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
<form method="post" name="AddLink" onSubmit="return Check()" action="Admin_Web_NewSubmitURL.asp">
  <table border="0" cellpadding="2" cellspacing="1" align="center" width="100%" class="border">
    <tr class="title">
      <td height="22" colspan="2" align="center"><strong>�����վ</strong></td>
    </tr>
    <tr class="tdbg">
      <td width="85"  align="right"><strong>һ�����:</strong></td>
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
document.AddLink.SelectTypea.options[document.AddLink.SelectTypea.length] = new Option('��ѡ��','0','��ѡ��');
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
	   	<%
     dim rsLinka
	sqlLink="select * from WebGoodSiteType_Main where id="&rsLink("fl")
	set rsLinka=Server.CreateObject("Adodb.RecordSet")
	rsLinka.open sqlLink,conn,1,3
        if not rsLinka.eof then
           response.write "<OPTION selected value='"&rsLinka("ID")&"' "
           response.write ">"&rsLinka("Type_Name")&"</OPTION>"
		end if
	rsLinka.close

	sqlLink="select * from WebGoodSiteType_Main"
	set rsLinka=Server.CreateObject("Adodb.RecordSet")
	rsLinka.open sqlLink,conn,1,3
        do while not rsLinka.eof
           response.write "<OPTION value='"&rsLinka("ID")&"' "
           response.write ">"&rsLinka("Type_Name")&"</OPTION>"
        rsLinka.movenext
        loop
	rsLinka.close
	set rsLinka=nothing
        %>
	</SELECT>      </td>
    </tr>
    <tr class="tdbg">
      <td width="85"  align="right"><strong>�������:</strong></td>
      <td valign="middle">
      	<SELECT name="SelectTypea">
	   <OPTION value="0" selected>��ѡ��</OPTION>
	</SELECT>      </td>
    </tr>
	    <tr class="tdbg">
      <td width="85" height="25" valign="middle"  align="right"><strong>��վ����:</strong></td>
      <td height="25"><input name="SiteName" class="Input" value="<%=rsLink("Sitename")%>" size="50"  maxlength="50">
          <font color="#FF0000"> *</font>����:<input type="checkbox" name="bold" value="1" ></td>
    </tr>
    <tr class="tdbg">
      <td width="85" height="25" valign="middle"  align="right"><strong>��վ����:</strong></td>
      <td height="25"><input name="WebUrl" class="Input" id="WebUrl" value="<%if instr(rsLink("SiteUrl"),";")>0 then response.write "����" else response.write rsLink("SiteUrl") end if%>" size="50"  maxlength="100">
          <font color="#FF0000"> * ����:<input type="checkbox" name="testspeed" value="1" onClick="if(this.checked==true){document.AddLink.WebUrl.value='����'}else{document.AddLink.WebUrl.value=''};" > ��ת��ת:<input type="checkbox" name="out" value="1" onClick="if(this.checked==true){document.AddLink.WebUrl.value='/out.asp?turl='+document.AddLink.WebUrl.value}else{var aaa=document.AddLink.WebUrl.value;document.AddLink.WebUrl.value=aaa.replace('/out.asp?turl=','');}"> </font></td>
    </tr>
    <tr class="tdbg">
      <td width="85"  align="right"><strong>������ɫ:</strong></td>
      <td valign="middle"><input type="text" name="select_color" class="Input" id="select_color" size="7" maxlength="7">
        &nbsp;<img language="javascript" onClick="foreColor()"  width="18" height="18" id="color_img" border=0 src="images/rect.gif" style="cursor:pointer;background-color:#ffffff"> </td>
    </tr>

    <tr class="tdbg">
      <td width="85" height="25"  align="right" valign="middle"><strong>�Ƽ�:</strong></td>
      <td height="25"><p>
          <input type="checkbox" name="Recommend" value="1" >
      </p></td>
    </tr>
    <tr class="tdbg">
      <td width="85" height="25"  align="right" valign="middle"><strong>��ʽ:</strong></td>
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
          <label for="radio6">��</label>
        &nbsp;</td>
    </tr>
<%
dim arrya,i,k
arrya=split(rsLink("SiteUrl"),";")
for i=0 to ubound(arrya)
%>
    <tr class="tdbg">
      <td width="85" height="25"  align="right" valign="middle"><strong>��������<%=i+1%>:</strong></td>
      <td height="25"><input name="xsSiteName" class="Input" id="xsSiteName" value="��Ա��·<%=i+1%>" size="50"  maxlength="50"></td>
    </tr>
    <tr class="tdbg">
      <td width="85" height="25"  align="right" valign="middle"><strong>����<%=i+1%>:</strong></td>
      <td height="25"><input name="xsWebUrl" class="Input" id="xsWebUrl<%=i+1%>" value="<%=arrya(i)%>" size="50"  maxlength="100">  ��ת��ת:<input type="checkbox" value="1" onClick="tsp(this,'<%=i+1%>');"> </td>
    </tr>
<%next%>

<%
k=i
arrya=split(rsLink("SiteContact"),";")
for i=0 to ubound(arrya)
%>
    <tr class="tdbg">
      <td width="85" height="25"  align="right" valign="middle"><strong>��������<%=k+i+1%>:</strong></td>
      <td height="25"><input name="xsSiteName" class="Input" id="xsSiteName" value="�����½<%=i+1%>" size="50"  maxlength="50"></td>
    </tr>
    <tr class="tdbg">
      <td width="85" height="25"  align="right" valign="middle"><strong>����<%=k+i+1%>:</strong></td>
      <td height="25"><input name="xsWebUrl" class="Input" id="xsWebUrl<%=k+i+1%>" value="<%=arrya(i)%>" size="50"  maxlength="100">  ��ת��ת:<input type="checkbox" value="1" onClick="tsp(this,'<%=k+i+1%>');"> </td>
    </tr>
<%next%>
<%
k=k+i
arrya=split(rsLink("chazhang"),";")
for i=0 to ubound(arrya)
%>
    <tr class="tdbg">
      <td width="85" height="25"  align="right" valign="middle"><strong>��������<%=k+i+1%>:</strong></td>
      <td height="25"><input name="xsSiteName" class="Input" id="xsSiteName" value="����<%=i+1%>" size="50"  maxlength="50"></td>
    </tr>
    <tr class="tdbg">
      <td width="85" height="25"  align="right" valign="middle"><strong>����<%=k+i+1%>:</strong></td>
      <td height="25"><input name="xsWebUrl" class="Input" id="xsWebUrl<%=k+i+1%>" value="<%=arrya(i)%>" size="50"  maxlength="100">  ��ת��ת:<input type="checkbox" value="1" onClick="tsp(this,'<%=k+i+1%>');"> </td>
    </tr>
<%next%>


    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"><input name="ID" type="hidden" id="ID" value="<%=ID%>">  
        <input type="submit" class="Button" value=" ȷ �� " name="cmdOk"> &nbsp; <input type="reset" class="Button" value=" �� �� " name="cmdReset">      </td>
    </tr>
  </table>
</form>
<script language="JavaScript" type="text/javascript">
changelocation(document.AddLink.SelectType.options[document.AddLink.SelectType.selectedIndex].value)
</script>

<%
     rsLink.close
     set rsLink=nothing
end sub

%></body>
</html>
<%
sub SaveModify()
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
		ErrMsg=ErrMsg & "<br><li>��ѡ��һ�����!</li>"
	end if
	if SelectTypea="0" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ѡ��������!</li>"
	end if
	if WebName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��������վ����!</li>"
	end if
	if WebUrl="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��������վ����!</li>"
	end if
	if WebColor="#NaNNaNNaN" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>������ɫ�Ƿ���ɫ�ַ���</li>"
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
			
			
			
			
			conn.execute "Delete From WebSiteSubmit_URL Where ID=" & CLng(ID)
			
			
			call CloseConn()
			
			
			response.write  "<script>alert('��˳ɹ�!');location='Admin_Web_NewSubmitURL.asp';</script>"
	end if
end sub
%>
