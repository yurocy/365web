<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2
Const CheckChannelID=0
Const PurviewLevel_Others="EditPwd"
%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/md5.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<%
dim Action,FoundErr,ErrMsg
dim rs,sql
Action=trim(request("Action"))
sql="Select * from Admin order by id desc"
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open sql,conn,1,3
if rs.Bof and rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�����ڴ��û���</li>"
else
	if Action="Modify" then
		call ModifyPwd()
	else
		call main()
	end if
end if
rs.close
set rs=nothing
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()
%>
<html>
<head>
<title>�޸Ĺ���Ա��Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
<script language=javascript>
function check()
{
  if(document.form1.Old_Password.value=="")
    {
      alert("ԭ���벻��Ϊ�գ�");
	  document.form1.Old_Password.focus();
      return false;
    }
  if(document.form1.Password.value=="")
    {
      alert("���벻��Ϊ�գ�");
	  document.form1.Password.focus();
      return false;
    }
    
  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("��ʼ������ȷ�����벻ͬ��");
	  document.form1.PwdConfirm.select();
	  document.form1.PwdConfirm.focus();	  
      return false;
    }
}
</script>
</head>
<body>
<%
sub main()
%>
<form method="post" name="form1" onsubmit="javascript:return check();" action="Admin_EditPwd.asp">
<br>
<br>
  <table width="300" border="0" align="center" cellpadding="2" cellspacing="1" class="border" >
    <tr class="title"> 
      <td height="22" colspan="2"> <div align="center"><strong>�� �� �� �� Ա �� ��</strong></div></td>
    </tr>
    <tr> 
      <td width="100" align="right" class="tdbg"><strong>�� �� ����</strong></td>
      <td class="tdbg"><%=rs("UserName")%></td>
    </tr>
    <tr>
      <td width="100" align="right" class="tdbg"><strong>�û�Ȩ�ޣ�</strong></td>
      <td class="tdbg">
        <%
		  select case rs("purview")
		  	case 1
				response.write "��������Ա"
			case 2
				response.write "��ͨ����Ա"
		  end select
		  %>
      </td>
    </tr>

   <tr> 
      <td width="100" align="right" class="tdbg"><strong>���û�����</strong></td>
      <td class="tdbg"><input type="text" class="Input" name="newusername"></td>
    </tr>
    <tr> 
      <td width="100" align="right" class="tdbg"><strong>ԭ �� �룺</strong></td>
      <td class="tdbg"><input type="password" class="Input" name="Old_Password"></td>
    </tr>
    <tr> 
      <td width="100" align="right" class="tdbg"><strong>�� �� �룺</strong></td>
      <td class="tdbg"><input type="password" class="Input" name="Password"> </td>
    </tr>
    <tr> 
      <td width="100" align="right" class="tdbg"><strong>ȷ�����룺</strong></td>
      <td class="tdbg"><input type="password" class="Input" name="PwdConfirm"> </td>
    </tr>
    <tr> 
      <td height="40" colspan="2" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="Modify"> 
        <input  type="submit" class="button" name="Submit" value=" ȷ �� " style="cursor:hand;"> 
        &nbsp; <input name="Cancel" type="button" class="button" id="Cancel" value=" ȡ �� " onClick="window.location.href='main_index.asp'" style="cursor:hand;"></td>
    </tr>
  </table>
  </form>
<%
end sub
sub ModifyPwd()
	dim password,PwdConfirm,OldPassWord,newusername
	newusername=trim(Request("newusername"))
password=trim(Request("Password"))
	PwdConfirm=trim(request("PwdConfirm"))
        OldPassWord=trim(request("Old_Password"))
	if OldPassWord="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>ԭ���벻��Ϊ�գ�</li>"
        else
                OldPassWord=md5(OldPassWord)
	end if
	if password="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�����벻��Ϊ�գ�</li>"
	end if
	if PwdConfirm<>Password then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>ȷ�������������������ͬ��</li>"
		exit sub
	end if
        if FoundErr<>True then
		dim sqlLink,rsLink
		sqlLink="select top 1 * from Admin where Password='" & OldPassWord & "'"

		set rsLink=Server.CreateObject("Adodb.RecordSet")
		rsLink.open sqlLink,conn,1,3
		if rsLink.bof and rsLink.eof then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>ԭ�������,��������������!</li>"
                        rsLink.close
		        set rsLink=nothing
		else
	           UserName=rs("UserName")
	           if Password<>"" then
		      rs("password")=md5(password)
			rs("UserName")=newusername

	           end if
                   rs.update
       		   rs.close
		   set rs=nothing
   		   rsLink.close
		   set rsLink=nothing
	           call WriteSuccessMsg("<p align='center'>�޸�����ɹ�,�´ε�¼ʱ����������,лл!</p><br><div align='right'><a href='main_index.asp'>�����ع�����ҳ��</a></div>")
		end if
                call CloseConn()
	end if
end sub
%>
</body>
</html>

