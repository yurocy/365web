
<%
dim UserLogined,UserName,UserLevel,ChargeType,UserPoint,ValidDays

Function ReturnGG(strClassName,strClassSeat)
     Dim sql,rs
     sql="select top 1 * from Admin_GG where ClassName='" & strClassName & "' and ClassSeat='"&strClassSeat&"'"
     set rs=server.createobject("adodb.recordset")
     rs.open sql,conn,1,1
     if rs.eof and rs.bof then
        ReturnGG="�޹��"
     else
        if rs("gg_type")="ͼƬ" then
           ReturnGG="<a href='"&rs("GG_URL")&"' target='_blank'><img src='"&rs("ClassImages")&"' width='"&rs("width")&"' height='"&rs("height")&"' border='0'></a>"
        elseif rs("gg_type")="�ⲿ����" then
           ReturnGG=rs("ClassImages")
        else
           ReturnGG="<OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0 height="&rs("height")&" width="&rs("width")&" classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000><PARAM NAME='_cx' VALUE='13229'><PARAM NAME='_cy' VALUE='15875'><PARAM NAME='FlashVars' VALUE=''><PARAM NAME='Movie' VALUE='"&TX_Image&"'><PARAM NAME='Src' VALUE='"&rs("ClassImages")&"'><PARAM NAME='WMode' VALUE='Window'><PARAM NAME='Play' VALUE='-1'><PARAM NAME='Loop' VALUE='-1'><PARAM NAME='Quality' VALUE='High'><PARAM NAME='SAlign' VALUE=''><PARAM NAME='Menu' VALUE='-1'><PARAM NAME='Base' VALUE=''><PARAM NAME='AllowScriptAccess' VALUE=''><PARAM NAME='Scale' VALUE='ShowAll'><PARAM NAME='DeviceFont' VALUE='0'><PARAM NAME='EmbedMovie' VALUE='0'><PARAM NAME='BGColor' VALUE=''><PARAM NAME='SWRemote' VALUE=''><PARAM NAME='MovieData' VALUE=''><PARAM NAME='SeamlessTabbing' VALUE='1'><PARAM NAME='Profile' VALUE='0'><PARAM NAME='ProfileAddress' VALUE=''><PARAM NAME='ProfilePort' VALUE='0'><PARAM NAME='AllowNetworking' VALUE='all'><PARAM NAME='AllowFullScreen' VALUE='false'><embed src='"&rs("ClassImages")&"' quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width="&rs("width")&" height="&rs("height")&">"&rs("ClassImages")&"</embed></OBJECT>"
        end if
     end if
     rs.close
     set rs=nothing
End Function

Function htmlencode(string)
htmlencode=replace(string,"'","")
'htmlencode=replace(htmlencode,"=","��")
'htmlencode=replace(htmlencode,chr(13)&chr(10),"<br>")
htmlencode=replace(htmlencode," ","")
htmlencode=replace(htmlencode,"<","&lt;")
htmlencode=replace(htmlencode,">","&gt;")
End Function

Function htmldecode(string)
htmldecode=replace(string,chr(13),"<br>")
End Function

Function GetRndPassword(PasswordLen)
 Dim Ran, i, strPassword
 strPassword = ""
 For i = 1 To PasswordLen
 Randomize
 Ran = CInt(Rnd * 2)
 Randomize
 If Ran = 0 Then
 Ran = CInt(Rnd * 25) + 97
 strPassword = strPassword & UCase(Chr(Ran))
 ElseIf Ran = 1 Then
 Ran = CInt(Rnd * 9)
 strPassword = strPassword & Ran
 ElseIf Ran = 2 Then
 Ran = CInt(Rnd * 25) + 97
 strPassword = strPassword & Chr(Ran)
 End If
 Next
 GetRndPassword = strPassword
End Function

'**************************************************
'��������ReplaceBadChar
'�� �ã����˷Ƿ���SQL�ַ�
'�� ����strChar-----Ҫ���˵��ַ�
'����ֵ�����˺���ַ�
'**************************************************
Public Function ReplaceBadChar(strChar)
 If strChar = "" Or IsNull(strChar) Then
 ReplaceBadChar = ""
 Exit Function
 End If
 Dim strBadChar, arrBadChar, tempChar, i
 strBadChar = "',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ""
 arrBadChar = Split(strBadChar, ",")
 tempChar = strChar
 For i = 0 To UBound(arrBadChar)
 tempChar = Replace(tempChar, arrBadChar(i), "")
 Next
 ReplaceBadChar = tempChar
End Function

'***************************************************
'�������Ƿ��Ѿ���װ
'***************************************************
Function IsObjInstalled(strClassString)
 On Error Resume Next
 IsObjInstalled = False
 Err = 0
 Dim xTestObj
 Set xTestObj = Server.CreateObject(strClassString)
 If 0 = Err Then IsObjInstalled = True
 Set xTestObj = Nothing
 Err = 0
End Function

'**************************************************
'��������gotTopic
'��  �ã����ַ���������һ���������ַ���Ӣ����һ���ַ�
'��  ����str   ----ԭ�ַ���
'       strlen ----��ȡ����
'����ֵ����ȡ����ַ���
'**************************************************
function gotTopic(str,strlen)
	if str="" then
		gotTopic=""
		exit function
	end if
	dim l,t,c, i
	str=replace(replace(replace(replace(str,"&nbsp;"," "),"&quot;",chr(34)),"&gt;",">"),"&lt;","<")
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			gotTopic=left(str,i) & "��"
			exit for
		else
			gotTopic=str
		end if
	next
	gotTopic=replace(replace(replace(replace(gotTopic," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function

'**************************************************
'��������JoinChar
'��  �ã����ַ�м��� ? �� &
'��  ����strUrl  ----��ַ
'����ֵ������ ? �� & ����ַ
'**************************************************
function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			else
				JoinChar=strUrl
			end if
		else
			JoinChar=strUrl & "?"
		end if
	else
		JoinChar=strUrl
	end if
end function

'**************************************************
'��������showpage
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����sfilename  ----���ӵ�ַ
'       totalnumber ----������
'       maxperpage  ----ÿҳ����
'       ShowTotal   ----�Ƿ���ʾ������
'       ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת����ĳЩҳ�治��ʹ�ã���������JS����
'       strUnit     ----������λ
'**************************************************
sub showpage(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<table align='center'><tr><td>"
	if ShowTotal=true then 
		strTemp=strTemp & "�� <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "��ҳ ��һҳ&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>��ҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>��һҳ</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "��һҳ βҳ"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>��һҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>βҳ</a>"
  	end if
   	strTemp=strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/ҳ"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;ת����<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">��" & i & "ҳ</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></table>"
	response.write strTemp
end sub

'**************************************************
'��������IsValidEmail
'��  �ã����Email��ַ�Ϸ���
'��  ����email ----Ҫ����Email��ַ
'����ֵ��True  ----Email��ַ�Ϸ�
'       False ----Email��ַ���Ϸ�
'**************************************************
function IsValidEmail(email)
	dim names, name, i, c
	IsValidEmail = true
	names = Split(email, "@")
	if UBound(names) <> 1 then
	   IsValidEmail = false
	   exit function
	end if
	for each name in names
		if Len(name) <= 0 then
			IsValidEmail = false
    		exit function
		end if
		for i = 1 to Len(name)
		    c = Lcase(Mid(name, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
		       IsValidEmail = false
		       exit function
		     end if
	   next
	   if Left(name, 1) = "." or Right(name, 1) = "." then
    	  IsValidEmail = false
	      exit function
	   end if
	next
	if InStr(names(1), ".") <= 0 then
		IsValidEmail = false
	   exit function
	end if
	i = Len(names(1)) - InStrRev(names(1), ".")
	if i <> 2 and i <> 3 then
	   IsValidEmail = false
	   exit function
	end if
	if InStr(email, "..") > 0 then
	   IsValidEmail = false
	end if
end function

'**************************************************
'��������IsObjInstalled
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True  ----�Ѿ���װ
'       False ----û�а�װ
'**************************************************
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

'**************************************************
'��������strLength
'��  �ã����ַ������ȡ������������ַ���Ӣ����һ���ַ���
'��  ����str  ----Ҫ�󳤶ȵ��ַ���
'����ֵ���ַ�������
'**************************************************
function strLength(str)
	ON ERROR RESUME NEXT
	dim WINNT_CHINESE
	WINNT_CHINESE    = (len("�й�")=2)
	if WINNT_CHINESE then
        dim l,t,c
        dim i
        l=len(str)
        t=l
        for i=1 to l
        	c=asc(mid(str,i,1))
            if c<0 then c=c+65536
            if c>255 then
                t=t+1
            end if
        next
        strLength=t
    else 
        strLength=len(str)
    end if
    if err.number<>0 then err.clear
end function

'------------------���ĳһĿ¼�Ƿ����-------------------
Function CheckDir(FolderPath)
	dim fso
	folderpath=Server.MapPath(".")&"\"&folderpath
	Set fso1 = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(FolderPath) then
	'����
		CheckDir = True
	Else
	'������
		CheckDir = False
	End if
	Set fso = nothing
End Function

'-------------����ָ����������Ŀ¼---------
Function MakeNewsDir(foldername)
	dim fso,f
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateFolder(foldername)
    MakeNewsDir = True
	Set fso = nothing
End Function


'**************************************************
'��������SendMail
'��  �ã���Jmail��������ʼ�
'��  ����MailtoAddress  ----�����˵�ַ
'        MailtoName    -----����������
'        Subject       -----����
'        MailBody      -----�ż�����
'        FromName      -----����������
'        MailFrom      -----�����˵�ַ
'        Priority      -----�ż����ȼ�
'**************************************************
function SendMail(MailtoAddress,MailtoName,Subject,MailBody,FromName,MailFrom,Priority)
	on error resume next
	Dim JMail
	Set JMail=Server.CreateObject("JMail.Message")
	if err then
		SendMail= "<br><li>û�а�װJMail���</li>"
		err.clear
		exit function
	end if
	JMail.Charset="gb2312"          '�ʼ�����
	JMail.silent=true
	JMail.ContentType = "text/html"     '�ʼ����ĸ�ʽ
	'JMail.ServerAddress=MailServer     '���������ʼ���SMTP������
   	'�����������ҪSMTP�����֤����ָ�����²���
	JMail.MailServerUserName = MailServerUserName    '��¼�û���
   	JMail.MailServerPassWord = MailServerPassword        '��¼����
  	JMail.MailDomain = MailDomain       '����������á�name@domain.com���������û�����¼ʱ����ָ��domain.com
	JMail.AddRecipient MailtoAddress,MailtoName     '������
	JMail.Subject=Subject         '����
	JMail.HMTLBody=MailBody       '�ʼ����ģ�HTML��ʽ��
	JMail.Body=MailBody          '�ʼ����ģ����ı���ʽ��
	JMail.FromName=FromName         '����������
	JMail.From = MailFrom         '������Email
	JMail.Priority=Priority              '�ʼ��ȼ���1Ϊ�Ӽ���3Ϊ��ͨ��5Ϊ�ͼ�
	JMail.Send(MailServer)
	SendMail =JMail.ErrorMessage
	JMail.Close
	Set JMail=nothing
end function

'**************************************************
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'**************************************************
sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>������Ϣ</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>��������Ŀ���ԭ��</b>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; ������һҳ</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'**************************************************
'��������WriteSuccessMsg
'��  �ã���ʾ�ɹ���ʾ��Ϣ
'��  ������
'**************************************************
sub WriteSuccessMsg(SuccessMsg)
	dim strSuccess
	strSuccess=strSuccess & "<html><head><title>�ɹ���Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strSuccess=strSuccess & "<link href='style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strSuccess=strSuccess & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strSuccess=strSuccess & "  <tr align='center' class='title'><td height='22'><strong>��ϲ�㣡</strong></td></tr>" & vbcrlf
	strSuccess=strSuccess & "  <tr class='tdbg'><td height='100' valign='top'><br>" & SuccessMsg &"</td></tr>" & vbcrlf
	strSuccess=strSuccess & "  <tr align='center' class='tdbg'><td>&nbsp;</td></tr>" & vbcrlf
	strSuccess=strSuccess & "</table>" & vbcrlf
	strSuccess=strSuccess & "</body></html>" & vbcrlf
	response.write strSuccess
end sub

'**************************************************
'��������WriteSuccessMsg
'��  �ã���ʾ�ɹ���ʾ��Ϣ
'��  ������
'**************************************************
sub WriteSuccessMsga(Title,SuccessMsg,URL1,URL2)
	dim strSuccess
	strSuccess=strSuccess & "<html><head><title>�ɹ���Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strSuccess=strSuccess & "<link href='style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strSuccess=strSuccess & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strSuccess=strSuccess & "  <tr align='center' class='title'><td height='22'><strong>"&Title&"</strong></td></tr>" & vbcrlf
	strSuccess=strSuccess & "  <tr class='tdbg'><td height='100' valign='top'><br>" & SuccessMsg &"</td></tr>" & vbcrlf
	strSuccess=strSuccess & "  <tr align='center' class='tdbg'><td>"&URL1&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&URL2&"</td></tr>" & vbcrlf
	strSuccess=strSuccess & "</table>" & vbcrlf
	strSuccess=strSuccess & "</body></html>" & vbcrlf
	response.write strSuccess
end sub

'**************************************************
'��������CheckUserLogined
'��  �ã�����û��Ƿ��¼
'��  ������
'����ֵ��True ----�Ѿ���¼
'        False ---û�е�¼
'**************************************************
function CheckUserLogined()
	dim Logined,Password,rsLogin,sqlLogin
	Logined=True
	UserName=Request.Cookies("asp163")("UserName")
	Password=Request.Cookies("asp163")("Password")
	UserLevel=Request.Cookies("asp163")("UserLevel")
	if UserName="" then
		Logined=False
	end if
	if Password="" then
		Logined=False
	end if
	if UserLevel="" then
		Logined=False
		UserLevel=9999
	end if
	if Logined=True then
		username=replace(trim(username),"'","")
		password=replace(trim(password),"'","")
		UserLevel=Cint(trim(UserLevel))
		set rsLogin=server.createobject("adodb.recordset")
		sqlLogin="select * from " & db_User_Table & " where " & db_User_LockUser & "=False and " & db_User_Name & "='" & username & "' and " & db_User_Password & "='" & password &"'"
		rsLogin.open sqlLogin,Conn_User,1,1
		if rsLogin.bof and rsLogin.eof then
			Logined=False
		else
			if password<>rsLogin(db_User_Password) or UserLevel<rsLogin(db_User_UserLevel) then
				Logined=False
			end if
			UserName=rsLogin(db_User_Name)
			UserLevel=rsLogin(db_User_UserLevel)
			ChargeType=rsLogin(db_User_ChargeType)
			UserPoint=rsLogin(db_User_UserPoint)
		  	if rsLogin(db_User_Valid_Unit)=1 then
				ValidDays=rsLogin(db_User_Valid_Num)
		  	elseif rsLogin(db_User_Valid_Unit)=2 then
				ValidDays=rsLogin(db_User_Valid_Num)*30
		  	elseif rsLogin(db_User_Valid_Unit)=3 then
				ValidDays=rsLogin(db_User_Valid_Num)*365
		  	end if
		  	ValidDays=ValidDays-DateDiff("D",rsLogin(db_User_BeginDate),now())
		end if
		rsLogin.close
		set rsLogin=nothing
	end if
	CheckUserLogined=Logined
end function

'**************************************************
'��������ReplaceBadChar
'��  �ã����˷Ƿ���SQL�ַ�
'��  ����strChar-----Ҫ���˵��ַ�
'����ֵ�����˺���ַ�
'**************************************************
function ReplaceBadChar(strChar)
	if strChar="" then
		ReplaceBadChar=""
	else
		ReplaceBadChar=replace(replace(replace(replace(replace(replace(replace(strChar,"'",""),"*",""),"?",""),"(",""),")",""),"<",""),".","")
	end if
end function

'**************************************************
'��������CheckLevel
'��  �ã�����û�����
'��  ����LevelNum-----Ҫ���ļ���ֵ
'����ֵ����������
'**************************************************
function CheckLevel(LevelNum)
	select case LevelNum
	case 9999
		CheckLevel="�ο�"
	case 999
		CheckLevel="ע���û�"
	case 99
		CheckLevel="�շ��û�"
	case 9
		CheckLevel="VIP�û�"
	case 5
		CheckLevel="����Ա"
	end select
end function

'==================================================
'��������ShowLogo
'��  �ã���ʾ��վLOGO
'��  ������
'==================================================
sub ShowLogo()
	if LogoUrl<>"" then
		response.write "<a href='" & SiteUrl & "' title='" & SiteName & "'>"
		if lcase(right(LogoUrl,3))<>"swf" then
			response.write "<img src='" & LogoUrl & "' width='180' height='60' border='0'>"
		else
			Response.Write "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='180' height='60'><param name='movie' value='" & LogoUrl & "'><param name='quality' value='high'><embed src='" & LogoUrl & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='480' height='60'></embed></object>"
		end if
		response.write "</a>"
	else
		response.write "<a href='http://www.asp163.net' title='�����ռ�'><img src='http://www.asp163.net/Photo/images/logo.gif' width='180' height='60' border='0'></a>"
	end if
end sub

'==================================================
'��������ShowBanner
'��  �ã���ʾ��վBanner
'��  ������
'==================================================
sub ShowBanner()
	if BannerUrl<>"" then
		if lcase(right(BannerUrl,3))="swf" then
			Response.Write "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='480' height='60'><param name='movie' value='" & BannerUrl & "'><param name='quality' value='high'><embed src='" & BannerUrl & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='480' height='60'></embed></object>"
		else
			response.Write "<a href='" & SiteUrl & "' title='" & SiteName & "'><img src='" & BannerUrl & "' width='480' height='60' border='0'></a>"
		end if
	else
		call ShowAD(1)
	end if
end sub

'==================================================
'��������ShowVote
'��  �ã���ʾ��վ����
'��  ������
'==================================================
sub ShowVote()
	dim sqlVote,rsVote,i
	sqlVote="select top 1 * from Vote where IsSelected=True"
	sqlVote=sqlVote& " and (ChannelID=0 or ChannelID=" & ChannelID & ") order by ID Desc"
	Set rsVote= Server.CreateObject("ADODB.Recordset")
	rsVote.open sqlVote,conn,1,1
	if rsVote.bof and rsVote.eof then 
		response.Write "&nbsp;û���κε���"
	else
		response.write "<form name='VoteForm' method='post' action='vote.asp' target='_blank'>"
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;" & rsVote("Title") & "<br>"
		if rsVote("VoteType")="Single" then
			for i=1 to 8
				if trim(rsVote("Select" & i) & "")="" then exit for
				response.Write "<input type='radio' name='VoteOption' value='" & i & "' style='border:0'>" & rsVote("Select" & i) & "<br>"
			next
		else
			for i=1 to 8
				if trim(rsVote("Select" & i) & "")="" then exit for
				response.Write "<input type='checkbox' name='VoteOption' value='" & i & "' style='border:0'>" & rsVote("Select" & i) & "<br>"
			next
		end if
		response.write "<br><input name='VoteType' type='hidden'value='" & rsVote("VoteType") & "'>"
		response.write "<input name='Action' type='hidden' value='Vote'>"
		response.write "<input name='ID' type='hidden' value='" & rsVote("ID") & "'>"
		response.write "<div align='center'>"
		response.write "<a href='javascript:VoteForm.submit();'><img src='images/voteSubmit.gif' width='52' height='18' border='0'></a>&nbsp;&nbsp;"
        response.write "<a href='Vote.asp?ID=" & rsVote("ID") & "&Action=Show' target='_blank'><img src='images/voteView.gif' width='52' height='18' border='0'></a>"
		response.write "</div></form>"
	end if
	rsVote.close
	set rsVote=nothing
end sub

'==================================================
'��������ShowAnnounce
'��  �ã���ʾ��վ������Ϣ
'��  ����ShowType ------��ʾ��ʽ��1Ϊ����2Ϊ����
'        AnnounceNum  ----�����ʾ����������
'==================================================
sub ShowAnnounce(ShowType,AnnounceNum)
	dim sqlAnnounce,rsAnnounce,i
	if AnnounceNum>0 and AnnounceNum<=10 then
		sqlAnnounce="select top " & AnnounceNum
	else
		sqlAnnounce="select top 10"
	end if
	sqlAnnounce=sqlAnnounce & " * from Announce where IsSelected=True"
	sqlAnnounce=sqlAnnounce & " and (ChannelID=0 or ChannelID=" & ChannelID & ")"
	sqlAnnounce=sqlAnnounce & " and (ShowType=0 or ShowType=1) order by ID Desc"
	Set rsAnnounce= Server.CreateObject("ADODB.Recordset")
	rsAnnounce.open sqlAnnounce,conn,1,1
	if rsAnnounce.bof and rsAnnounce.eof then 
		AnnounceCount=0
		response.write "<p>&nbsp;&nbsp;û��ͨ��</p>" 
	else 
		AnnounceCount=rsAnnounce.recordcount
		if ShowType=1 then
			do while not rsAnnounce.eof   
				response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' onclick=""javascript:window.open('Announce.asp?ChannelID=" & ChannelID & "&ID=" & rsAnnounce("id") &"', 'newwindow', 'height=300, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"" title='" & rsAnnounce("Content") & "'>" & rsAnnounce("title") & "</div><br><div align='right'>" & rsAnnounce("Author") & "&nbsp;&nbsp;<br>" & FormatDateTime(rsAnnounce("DateAndTime"),1) & "</a>"
				rsAnnounce.movenext
				i=i+1
				if i<AnnounceCount then response.write "<hr>"   
			loop
		else
			do while not rsAnnounce.eof   
				response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' onclick=""javascript:window.open('Announce.asp?ChannelID=" & ChannelID & "&ID=" & rsAnnounce("id") &"', 'newwindow', 'height=300, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"" title='" & rsAnnounce("Content") & "' >" & rsAnnounce("title") & "&nbsp;&nbsp;[" & rsAnnounce("Author") & "&nbsp;&nbsp;" & FormatDateTime(rsAnnounce("DateAndTime"),1) & "]</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				rsAnnounce.movenext
			loop
       	end if	
	end if  
	rsAnnounce.close
	set rsAnnounce=nothing
end sub


'==================================================
'��������ShowFriendSite
'��  �ã���ʾ��������վ��
'��  ����LinkType  ----���ӷ�ʽ��1ΪLOGO���ӣ�2Ϊ��������
'       SiteNum   ----�����ʾ���ٸ�վ��
'       Cols      ----�ּ�����ʾ
'       ShowType  ----��ʾ��ʽ��1Ϊ���Ϲ�����2Ϊ�����б�3Ϊ�����б��
'==================================================
sub ShowFriendSite(LinkType,SiteNum,Cols,ShowType)
	dim sqlLink,rsLink,SiteCount,i,strLink
	if LinkType<>1 and LinkType<>2 then
		LinkType=1
	else
		LinkType=Cint(LinkType)
	end if
	if SiteNum<=0 or SiteNum>100 then
		SiteNum=10
	end if
	if Cols<=0 or Cols>20 then
		Cols=10
	end if
	if ShowType=1 then
'		strLink=strLink & "<marquee id='LinkScrollArea' direction='up' scrolldelay='50' scrollamount='1' width='100' height='100' onmouseover='this.stop();' onmouseout='this.start();'>"
        strLink=strLink & "<div id=rolllink style=overflow:hidden;height:100;width:100><div id=rolllink1>"    '�����ӵĴ���
	elseif ShowType=3 then
		strLink=strLink & "<select name='FriendSite' onchange=""if(this.options[this.selectedIndex].value!=''){window.open(this.options[this.selectedIndex].value,'_blank');}""><option value=''>������������վ��</option>"
	end if
	if ShowType=1 or ShowType=2 then
		strLink=strLink & "<table width='100%' cellSpacing='5'><tr align='center' class='tdbg'>"
	end if
	
	sqlLink="select top " & SiteNum & " * from FriendSite where IsOK=True and LinkType=" & LinkType & " order by IsGood,id desc"
	set rsLink=server.createobject("adodb.recordset")
	rsLink.open sqlLink,conn,1,1
	if rsLink.bof and rsLink.eof then
		if ShowType=1 or ShowType=2 then
	  		for i=1 to SiteNum
				strLink=strLink & "<td><a href='FriendSiteReg.asp' target='_blank'>"
				if LinkType=1 then
					strLink=strLink & "<img src='images/nologo.jpg' width='88' height='31' border='0' alt='�������'>"
				else
					strLink=strLink & "�������"
				end if
				strLink=strLink & "</a></td>"
				if i mod Cols=0 and i<SiteNum then
					strLink=strLink & "</tr><tr align='center' class='tdbg'>"
				end if
			next
		end if
	else
		SiteCount=rsLink.recordcount
		for i=1 to SiteCount
			if ShowType=1 or ShowType=2 then
			  if LinkType=1 then
				strLink=strLink & "<td width='88'><a href='" & rsLink("SiteUrl") & "' target='_blank' title='��վ���ƣ�" & rsLink("SiteName") & vbcrlf & "��վ��ַ��" & rsLink("SiteUrl") & vbcrlf & "��վ��飺" & rsLink("SiteIntro") & "'>"
				if rsLink("LogoUrl")="" or rsLink("LogoUrl")="http://" then
					strLink=strLink & "<img src='images/nologo.gif' width='88' height='31' border='0'>"
				else
					strLink=strLink & "<img src='" & rsLink("LogoUrl") & "' width='88' height='31' border='0'>"
				end if
				strLink=strLink & "</a></td>"
			  else
				strLink=strLink & "<td width='88'><a href='" & rsLink("SiteUrl") & "' target='_blank' title='��վ���ƣ�" & rsLink("SiteName") & vbcrlf & "��վ��ַ��" & rsLink("SiteUrl") & vbcrlf & "��վ��飺" & rsLink("SiteIntro") & "'>" & rsLink("SiteName") & "</a></td>"
			  end if
			  if i mod Cols=0 and i<SiteNum then
				strLink=strLink & "</tr><tr align='center' class='tdbg'>"
			  end if
			else
				strLink=strLink & "<option value='" & rsLink("SiteUrl") & "'>" & rsLink("SiteName") & "</option>"
			end if
			rsLink.moveNext
		next
		if SiteCount<SiteNum and (ShowType=1 or ShowType=2) then
			for i=SiteCount+1 to SiteNum
				if LinkType=1 then
					strLink=strLink & "<td width='88'><a href='FriendSiteReg.asp' target='_blank'><img src='images/nologo.jpg' width='88' height='31' border='0' alt='�������'></a></td>"
				else
					strLink=strLink & "<td width='88'><a href='FriendSiteReg.asp' target='_blank'>�������</a></td>"
				end if
				if i mod Cols=0 and i<SiteNum then
					strLink=strLink & "</tr><tr align='center' class='tdbg'>"
				end if
			next
		end if
	end if
	if ShowType=1 or ShowType=2 then
		strLink=strLink & "</tr></table>"
	end if
	if ShowType=1 then
'		strLink=strLink & "</marquee>"
        strLink=strLink & "</div><div id=rolllink2></div></div>"   '��������
	elseif ShowType=3 then
		strLink=strLink & "</select>"
	end if
	response.write strLink
	if ShowType=1 then call RollFriendSite()    '��������
	rsLink.close
	set rsLink=nothing
end sub

'==================================================
'��������RollFriendSite
'��  �ã�������ʾ��������վ��
'��  ������
'==================================================
sub RollFriendSite()
%>
<script>
   var rollspeed=30
   rolllink2.innerHTML=rolllink1.innerHTML //��¡rolllink1Ϊrolllink2
   function Marquee(){
   if(rolllink2.offsetTop-rolllink.scrollTop<=0) //��������rolllink1��rolllink2����ʱ
   rolllink.scrollTop-=rolllink1.offsetHeight  //rolllink�������
   else{
   rolllink.scrollTop++
   }
   }
   var MyMar=setInterval(Marquee,rollspeed) //���ö�ʱ��
   rolllink.onmouseover=function() {clearInterval(MyMar)}//�������ʱ�����ʱ���ﵽ����ֹͣ��Ŀ��
   rolllink.onmouseout=function() {MyMar=setInterval(Marquee,rollspeed)}//����ƿ�ʱ���趨ʱ��
</script>
<%
end sub

sub ShowGoodSite(SiteNum)
	dim sqlLink,rsLink,SiteCount,i,strLink
	if SiteNum<=0 or SiteNum>100 then
		SiteNum=10
	end if
	strLink=strLink & "<table width='100%' cellSpacing='5'>"
	
	sqlLink="select top " & SiteNum & " * from FriendSite where IsOK=True and LinkType=1 and IsGood=True order by id desc"
	set rsLink=server.createobject("adodb.recordset")
	rsLink.open sqlLink,conn,1,1
	if rsLink.bof and rsLink.eof then
	 	for i=1 to SiteNum
			strLink=strLink & "<tr align='center'><td><a href='FriendSiteReg.asp' target='_blank'><img src='images/nologo.jpg' width='88' height='31' border='0' alt='�������'></a></td></tr>"
		next
	else
		SiteCount=rsLink.recordcount
		for i=1 to SiteCount
			strLink=strLink & "<tr align='center'><td><a href='" & rsLink("SiteUrl") & "' target='_blank' title='��վ���ƣ�" & rsLink("SiteName") & vbcrlf & "��վ��ַ��" & rsLink("SiteUrl") & vbcrlf & "��վ��飺" & rsLink("SiteIntro") & "'>"
			if rsLink("LogoUrl")="" or rsLink("LogoUrl")="http://" then
				strLink=strLink & "<img src='images/nologo.gif' width='88' height='31' border='0'>"
			else
				strLink=strLink & "<img src='" & rsLink("LogoUrl") & "' width='88' height='31' border='0'>"
			end if
			strLink=strLink & "</a></td></tr>"
			rsLink.moveNext
		next
		for i=SiteCount+1 to SiteNum
			strLink=strLink & "<tr align='center'><td><a href='FriendSiteReg.asp' target='_blank'><img src='images/nologo.jpg' width='88' height='31' border='0' alt='�������'></a></td></tr>"
		next
	end if
	strLink=strLink & "</table>"
	response.write strLink
	rsLink.close
	set rsLink=nothing

end sub

sub Bottom()
	dim strTemp
	strTemp="<table width='760' align='center' border='0' class='topborder' cellpadding='0' cellspacing='0'><tr height='22' align='center'><td class='title_maintxt'>"
	strTemp= strTemp & "|&nbsp;<a href='#' onClick=this.style.behavior='url(#default#homepage)';this.setHomePage('"& SiteUrl & "');>��"&"Ϊ"&"��"&"ҳ</a>&nbsp;|&nbsp;"
	strTemp= strTemp & "<a href=javascript:window.external.addFavorite('" & SiteUrl & "','" & SiteName & "')>��"&"��"&"��"&"��</a>&nbsp;|&nbsp;"
	strTemp= strTemp & "<a href='mailto:" & WebmasterEmail & "'>��"&"ϵ"&"վ"&"��</a>&nbsp;|&nbsp;"
	strTemp= strTemp & "<a href='FriendSite.asp' target='_blank'>��"&"��"&"��"&"��</a>&nbsp;|&nbsp;"
	strTemp= strTemp & "<a href='Copyright.asp' target='_blank'>��"&"Ȩ"&"��"&"��</a>&nbsp;|&nbsp;"
	strTemp= strTemp & "<a href='login.asp' target='_blank'>��"&"��"&"��"&"¼</a>&nbsp;|&nbsp;"
	strTemp= strTemp & "</td></tr><tr align='center' height='20' valign='bottom'><td>"
	strTemp= strTemp & Copyright
	strTemp= strTemp & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;վ����<a href='mailto:" & WebmasterEmail & "'>" & WebmasterName & "</a>"
	if ShowRunTime="Yes" then
		strTemp= strTemp & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ҳ"&"��"&"ִ"&"��"&"ʱ"&"�䣺" & CStr(FormatNumber((Timer-BeginTime)*1000,2)) & "����"
	end if
	strTemp= strTemp & "<br>Powe"&"red by��<a href='http://www.asp1"&"63.net'>MyPo"&"wer Ve"&"r3.51</a>"
	strTemp= strTemp & "</td></tr></table>"
	response.write strTemp
end sub

'==================================================
'��������ShowUserLogin
'��  �ã���ʾ�û���¼��
'��  ������
'==================================================
sub ShowUserLogin()
	dim strLogin
	if CheckUserLogined()=False then
    	strLogin="<table align='center' width='100%' border='0' cellspacing='0' cellpadding='0'>" & vbcrlf
		strLogin=strLogin &  "<form action='User_ChkLogin.asp' method='post' name='UserLogin' onSubmit='return CheckForm();'>" & vbcrlf
        strLogin=strLogin & "<tr><td height='25' align='right'>�û�����</td><td height='25'><input name='UserName' type='text' id='UserName' size='10' maxlength='20'></td></tr>" & vbcrlf
        strLogin=strLogin & "<tr><td height='25' align='right'>��&nbsp;&nbsp;�룺</td><td height='25'><input name='Password' type='password' id='Password' size='10' maxlength='20'></td></tr>" & vbcrlf
        strLogin=strLogin & "<tr><td height='25' align='right'>Cookie��</td><td height='25'><select name=CookieDate><option selected value=0>������</option><option value=1>����һ��</option>" & vbcrlf
		strLogin=strLogin & "<option value=2>����һ��</option><option value=3>����һ��</option></select></td></tr>" & vbcrlf
		strLogin=strLogin & "<tr align='center'><td height='30' colspan='2'><input name='Login' type='submit' id='Login' value=' ��¼ '> <input name='Reset' type='reset' id='Reset' value=' ��� '>" & vbcrlf
        strLogin=strLogin & "<br><br><a href='User_Reg.asp' target='_blank'>���û�ע��</a>&nbsp;&nbsp;<a href='User_GetPassword.asp'>�������룿</a><br></td>" & vbcrlf      
        strLogin=strLogin & "</tr></form></table>" & vbcrlf
		response.write strLogin
%>
<script language=javascript>
	function CheckForm()
	{
		if(document.UserLogin.UserName.value=="")
		{
			alert("�������û�����");
			document.UserLogin.UserName.focus();
			return false;
		}
		if(document.UserLogin.Password.value == "")
		{
			alert("���������룡");
			document.UserLogin.Password.focus();
			return false;
		}
	}
	function openScript(url, width, height)
	{
		var Win = window.open(url,"UserControlPad",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=yes,status=yes' );
	}
</script>
<%
	Else 
		response.write "��ӭ����<font color=green><b>" & UserName & "</b></font>���þò�����"
		response.write "<br>������ݣ�"
		if UserLevel=999 then
			response.write "ע���û�"
		elseif UserLevel=99 then
			response.write "�շ��û�"
		elseif UserLevel=9 then
			response.write "VIP�û�"
		end if
		response.write "<br>�Ʒѷ�ʽ��"
		if ChargeType=1 then
			if UserPoint>0 then
				response.write "�۵���<br>���õ����� <b><font color=blue>" & UserPoint & "</font></b> ��"
				if UserPoint<=10 then
					response.write "<br><font color=red>��Ŀ��õ����Ѳ��࣬�뼰ʱ��ϵ���ǽ��г�ֵ��</font>"
				end if
			else
				response.write "�۵���<br>���õ����� <b><font color=red>" & UserPoint & "</font></b> ��"
				response.write "<br><font color=red>��Ŀ��õ����Ѿ����꣬����ϵ���ǽ��г�ֵ�������㽫�����Ķ��շ����ݡ�</font>"
			end if
		else
			if ValidDays>0 then
				response.write "��Ч��<br>��Ч������ <b><font color=blue>" & ValidDays & "</font></b> ��"
				if ValidDays<=10 then
					response.write "<br><font color=red>�����Ч��ʱ���Ѳ������뼰ʱ��ϵ���ǽ��г�ֵ��</font>"
				end if
			else
				response.write "��Ч��<br>��Ч������ <b><font color=red>" & ValidDays & "</font></b> ��"
				response.write "<br><font color=red>�����Ч���Ѿ����ڣ�����ϵ���ǽ��г�ֵ�������㽫�����Ķ��շ����ݡ�</font>"
			end if
		end if
		response.write "<br><b>�û�������壺</b><br>" & vbcrlf
		response.write "&nbsp;&nbsp;&nbsp;<a href=""JavaScript:openScript('User_ControlPad.asp?Action=ArticleAdd')"">��������</a>" & vbcrlf
		response.write "&nbsp;&nbsp;<a href=""JavaScript:openScript('User_ControlPad.asp?Action=ArticleManage')"">���¹���</a><br>" & vbcrlf
		response.write "&nbsp;&nbsp;&nbsp;<a href=""JavaScript:openScript('User_ControlPad.asp?Action=ModifyPwd')"">�޸�����</a>" & vbcrlf
		response.write "&nbsp;&nbsp;<a href=""JavaScript:openScript('User_ControlPad.asp?Action=ModifyInfo')"">������Ϣ</a><br>" & vbcrlf
		response.write "<div align='center'><a href='User_Logout.asp'>��ע����¼��</a></div>" & vbcrlf
	end if
%>
<script language=javascript>
	function openScript(url)
	{
		var Win = window.open(url,"UserControlPad");
	}
	function openScript2(url, width, height)
	{
		var Win = window.open(url,"UserControlPad",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=yes,status=yes' );
	}
</script>
<%
end sub

'==================================================
'��������ShowTopUser
'��  �ã���ʾ�û����У����ѷ������������������ȣ��ٰ�ע���Ⱥ�˳������
'��  ����UserNum-------��ʾ���û�����
'==================================================
sub ShowTopUser(UserNum)
	if UserNum<=0 or UserNum>100 then UserNum=10
	dim sqlTopUser,rsTopUser,i
	sqlTopUser="select top " & UserNum & " * from " & db_User_Table & " order by " & db_User_ArticleChecked & " desc," & db_User_ID & " asc"
	set rsTopUser=server.createobject("adodb.recordset")
	rsTopUser.open sqlTopUser,Conn_User,1,1
	if rsTopUser.bof and rsTopUser.eof then
		response.write "û���κ��û�"
	else
		response.write "<table width='100%' border='0' cellspacing='0' cellpadding='0'><tr><td align='left'>����</td><td align='left'>�û���</td><td align='right'>������</td></tr>"
		for i=1 to rsTopUser.recordcount
			response.write "<tr><td align='center'>" & cstr(i) & "</td><td align='left'><a href='UserInfo.asp?UserID=" & rsTopUser(db_User_ID) & "'>" & rsTopUser(db_User_Name) & "</a></td><td align='right'>" & rsTopUser(db_User_ArticleChecked) & "</td></tr>"
			rsTopUser.movenext
		next
		response.write "</table><div align='right'><a href='UserList.asp'>more...</a></div>"
	end if
	set rsTopUser=nothing
end sub

'==================================================
'��������ShowAllUser
'��  �ã���ҳ��ʾ�����û�
'��  ������
'==================================================
sub ShowAllUser()
	select case OrderType
	case 1
		sqlUser="select * from " & db_User_Table & " order by " & db_User_ArticleChecked & " desc"
	case 2
		sqlUser="select * from " & db_User_Table & " order by " & db_User_RegDate & " desc"
	case 3
		sqlUser="select * from " & db_User_Table & " order by " & db_User_ID & " desc"
	end select
	set rsUser=server.createobject("adodb.recordset")
	rsUser.open sqlUser,Conn_User,1,1
	if rsUser.bof and rsUser.eof then
		totalput=0
		response.write "<br><li>û���κ��û�</li>"
	else
		totalput=rsUser.recordcount
		if currentPage=1 then
			call ShowUserList()
		else
			if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsUser.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsUser.bookmark
            	call ShowUserList()
        	else
	        	currentPage=1
           		call ShowUserList()
	    	end if
		end if
	end if
	rsUser.close
	set rsUser=nothing
end sub

sub ShowUserList()
	dim i
	i=0
	response.write "<div align='center'><br><a href='UserList.asp?OrderType=1'>����������������</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='UserList.asp?OrderType=2'>��ע����������</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='UserList.asp?OrderType=3'>���û�ID����</a><br></div>"
	response.write "<table width='100%' border='0' cellspacing='1' cellpadding='3' bgcolor='#f9f9f9'><tr align='center'><td bgcolor='#f0f0f0'>�û���</td><td bgcolor='#f0f0f0'>�Ա�</td><td bgcolor='#f0f0f0'>Email</td><td bgcolor='#f0f0f0'>QQ����</td><td bgcolor='#f0f0f0'>MSN</td><td bgcolor='#f0f0f0'>��ҳ</td><td bgcolor='#f0f0f0'>ע������</td><td bgcolor='#f0f0f0'>������</td><tr>"
	do while not rsUser.eof
		response.write "<tr onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
		response.write "<td><a href='UserInfo.asp?UserID=" & rsUser(db_User_ID) & "'>" & rsUser(db_User_Name) & "</a></td><td align='center'>"
		if rsUser(db_User_Sex)=1 then
			response.write "��"
		else
			response.write "Ů"
		end if
		response.write "</td><td><a href='mailto:" & rsUser(db_User_Email) & "'>" & rsUser(db_User_Email) & "</a><td align='center'>"
		if rsUser(db_User_QQ)<>"" then
			response.write rsUser(db_User_QQ)
		else
			response.write "δ��"
		end if
		response.write "</td><td align='center'>"
		if rsUser(db_User_Msn)<>"" then
			response.write rsUser(db_User_Msn)
		else
			response.write "δ��"
		end if
		response.write "</td><td align='center'>"
		if rsUser(db_User_Homepage)<>"" and rsUser(db_User_Homepage)<>"http://" then
			response.write "<a href='" & rsUser(db_User_Homepage) & "' title='" & rsUser(db_User_Homepage) & "'>��˷���</a>"
		else
			response.write "δ��"
		end if
		response.write "</td><td align='center'>" & FormatDateTime(rsUser(db_User_RegDate),2) & "</td><td align='right'>" & rsUser(db_User_ArticleChecked) & "</td></tr>"
		
		rsUser.movenext
		i=i+1
		if i>=MaxPerPage then exit do
	loop
	response.write "</table>"
end sub

'==================================================
'��������PopAnnouceWindow
'��  �ã��������洰��
'��  ����Width-------�������ڿ��
'		 Height------�������ڸ߶�
'==================================================
sub PopAnnouceWindow(Width,Height)
	dim popCount,rsAnnounce
	set rsAnnounce=conn.execute("select count(*) from Announce where IsSelected=True and (ChannelID=0 or ChannelID=" & ChannelID & ") and (ShowType=0 or ShowType=2)")
	popCount=rsAnnounce(0)
	if popCount>0 then
		if  PopAnnounce="Yes" and session("Poped")<>ChannelID then
			response.write "<script LANGUAGE='JavaScript'>"
			response.write "window.open ('Announce.asp?ChannelID=" & ChannelID & "', 'newwindow', 'height=" & Height & ", width=" & Width & ", toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"
			response.write "</script>"
			session("Poped")=ChannelID
		end if
	end if
end sub

'==================================================
'��������ShowPath
'��  �ã���ʾ������������λ�á�������Ϣ
'��  ������
'==================================================
sub ShowPath()
	if PageTitle<>"" and ChannelID<>1 then
		strPath=strPath & "&nbsp;&gt;&gt;&nbsp;" & PageTitle
	end if
	response.write strPath
end sub

'==================================================
'��������MenuJS
'��  �ã����������˵���ص�JS����
'��  ������
'==================================================
sub MenuJS()
	dim strMenu
	if ShowMyStyle="Yes" then
%>
<script language="JavaScript" type="text/JavaScript">
//�����˵���ش���
 var h;
 var w;
 var l;
 var t;
 var topMar = 1;
 var leftMar = -2;
 var space = 1;
 var isvisible;
 var MENU_SHADOW_COLOR='#999999';//���������˵���Ӱɫ
 var global = window.document
 global.fo_currentMenu = null
 global.fo_shadows = new Array

function HideMenu() 
{
 var mX;
 var mY;
 var vDiv;
 var mDiv;
	if (isvisible == true)
{
		vDiv = document.all("menuDiv");
		mX = window.event.clientX + document.body.scrollLeft;
		mY = window.event.clientY + document.body.scrollTop;
		if ((mX < parseInt(vDiv.style.left)) || (mX > parseInt(vDiv.style.left)+vDiv.offsetWidth) || (mY < parseInt(vDiv.style.top)-h) || (mY > parseInt(vDiv.style.top)+vDiv.offsetHeight)){
			vDiv.style.visibility = "hidden";
			isvisible = false;
		}
}
}

function ShowMenu(vMnuCode,tWidth) {
	vSrc = window.event.srcElement;
	vMnuCode = "<table id='submenu' cellspacing=1 cellpadding=3 style='width:"+tWidth+"' class=menu onmouseout='HideMenu()'><tr height=23><td nowrap align=left class=MenuBody>" + vMnuCode + "</td></tr></table>";

	h = vSrc.offsetHeight;
	w = vSrc.offsetWidth;
	l = vSrc.offsetLeft + leftMar+4;
	t = vSrc.offsetTop + topMar + h + space-2;
	vParent = vSrc.offsetParent;
	while (vParent.tagName.toUpperCase() != "BODY")
	{
		l += vParent.offsetLeft;
		t += vParent.offsetTop;
		vParent = vParent.offsetParent;
	}

	menuDiv.innerHTML = vMnuCode;
	menuDiv.style.top = t;
	menuDiv.style.left = l;
	menuDiv.style.visibility = "visible";
	isvisible = true;
    makeRectangularDropShadow(submenu, MENU_SHADOW_COLOR, 4)
}

function makeRectangularDropShadow(el, color, size)
{
	var i;
	for (i=size; i>0; i--)
	{
		var rect = document.createElement('div');
		var rs = rect.style
		rs.position = 'absolute';
		rs.left = (el.style.posLeft + i) + 'px';
		rs.top = (el.style.posTop + i) + 'px';
		rs.width = el.offsetWidth + 'px';
		rs.height = el.offsetHeight + 'px';
		rs.zIndex = el.style.zIndex - i;
		rs.backgroundColor = color;
		var opacity = 1 - i / (i + 1);
		rs.filter = 'alpha(opacity=' + (100 * opacity) + ')';
		el.insertAdjacentElement('afterEnd', rect);
		global.fo_shadows[global.fo_shadows.length] = rect;
	}
}
</script>
<%
		response.write "<script language='JavaScript' type='text/JavaScript'>" & vbcrlf
		response.write "//�˵��б�" & vbcrlf
	
		'��ѡ���Ĳ˵�����
		strMenu="var menu_skin=" & chr(34)
		dim rsSkin
		set rsSkin=conn.execute("select SkinID,SkinName from Skin")
		do while not rsSkin.eof
			strMenu=strMenu & "&nbsp;<a style=font-size:9pt;line-height:14pt; href='SetCookie.asp?Action=SetSkin&ClassID=" & ClassID & "&SkinID=" & rsSkin(0) & "'>" & rsSkin(1) & "</a><br>"
			rsSkin.movenext
		loop
		rsSkin.close
		set rsSkin=nothing
		response.write strMenu & chr(34) & ";" & vbcrlf
		response.write "</script>" & vbcrlf
	else
	%>
	<script language="JavaScript" type="text/JavaScript">
	function HideMenu() 
	{
	}
	</script>
	<%
	end if
	
	if ChannelID>=2 and ChannelID<=4 then
		'���޼������˵���JS�����ļ�
		response.write "<script type='text/javascript' language='JavaScript1.2' src='stm31.js'></script>"
		if ShowClassTreeGuide="Yes" then
%>
<script language="JavaScript" type="text/JavaScript">
//���ε�����JS����
document.write("<style type=text/css>#master {LEFT: -200px; POSITION: absolute; TOP: 25px; VISIBILITY: visible; Z-INDEX: 999}</style>")
document.write("<table id=master width='218' border='0' cellspacing='0' cellpadding='0'><tr><td><img border=0 height=6 src=images/menutop.gif  width=200></td><td rowspan='2' valign='top'><img id=menu onMouseOver=javascript:expand() border=0 height=70 name=menutop src=images/menuo.gif width=18></td></tr>");
document.write("<tr><td valign='top'><table width='100%' border='0' cellspacing='5' cellpadding='0'><tr><td height='400' valign='top'><table width=100% height='100%' border=1 cellpadding=0 cellspacing=5 bordercolor='#666666' bgcolor=#ecf6f5 style=FILTER: alpha(opacity=90)><tr>");
document.write("<td height='10' align='center' bordercolor='#ecf6f5'><font color=999900><strong>�� Ŀ �� �� �� ��</strong></font></td></tr><tr><td valign='top' bordercolor='#ecf6f5'>");
document.write("<iframe width=100% height=100% src='classtree.asp?ChannelID=<%=ChannelID%>' frameborder=0></iframe></td></tr></table></td></tr></table></td></tr></table>");

var ie = document.all ? 1 : 0
var ns = document.layers ? 1 : 0
var master = new Object("element")
master.curLeft = -200;	master.curTop = 10;
master.gapLeft = 0;		master.gapTop = 0;
master.timer = null;

if(ie){var sidemenu = document.all.master;}
if(ns){var sidemenu = document.master;}
setInterval("FixY()",100);

function moveAlong(layerName, paceLeft, paceTop, fromLeft, fromTop){
	clearTimeout(eval(layerName).timer)
	if(eval(layerName).curLeft != fromLeft){
		if((Math.max(eval(layerName).curLeft, fromLeft) - Math.min(eval(layerName).curLeft, fromLeft)) < paceLeft){eval(layerName).curLeft = fromLeft}
		else if(eval(layerName).curLeft < fromLeft){eval(layerName).curLeft = eval(layerName).curLeft + paceLeft}
			else if(eval(layerName).curLeft > fromLeft){eval(layerName).curLeft = eval(layerName).curLeft - paceLeft}
		if(ie){document.all[layerName].style.left = eval(layerName).curLeft}
		if(ns){document[layerName].left = eval(layerName).curLeft}
	}
	if(eval(layerName).curTop != fromTop){
   if((Math.max(eval(layerName).curTop, fromTop) - Math.min(eval(layerName).curTop, fromTop)) < paceTop){eval(layerName).curTop = fromTop}
		else if(eval(layerName).curTop < fromTop){eval(layerName).curTop = eval(layerName).curTop + paceTop}
			else if(eval(layerName).curTop > fromTop){eval(layerName).curTop = eval(layerName).curTop - paceTop}
		if(ie){document.all[layerName].style.top = eval(layerName).curTop}
		if(ns){document[layerName].top = eval(layerName).curTop}
	}
	eval(layerName).timer=setTimeout('moveAlong("'+layerName+'",'+paceLeft+','+paceTop+','+fromLeft+','+fromTop+')',30)
}

function setPace(layerName, fromLeft, fromTop, motionSpeed){
	eval(layerName).gapLeft = (Math.max(eval(layerName).curLeft, fromLeft) - Math.min(eval(layerName).curLeft, fromLeft))/motionSpeed
	eval(layerName).gapTop = (Math.max(eval(layerName).curTop, fromTop) - Math.min(eval(layerName).curTop, fromTop))/motionSpeed
	moveAlong(layerName, eval(layerName).gapLeft, eval(layerName).gapTop, fromLeft, fromTop)
}

var expandState = 0

function expand(){
	if(expandState == 0){setPace("master", 0, 10, 10); if(ie){document.menutop.src = "images/menui.gif"}; expandState = 1;}
	else{setPace("master", -200, 10, 10); if(ie){document.menutop.src = "images/menuo.gif"}; expandState = 0;}
}

function FixY(){
	if(ie){sidemenu.style.top = document.body.scrollTop+10}
	if(ns){sidemenu.top = window.pageYOffset+10}
}
</script>
<%
		end if
	end if
end sub

'==================================================
'��������ShowSearchForm
'��  �ã���ʾ����������
'��  ����ShowType ----��ʾ��ʽ��1Ϊ���ģʽ��2Ϊ��׼ģʽ��3Ϊ�߼�ģʽ
'==================================================
sub ShowSearchForm(Action,ShowType)
	if ShowType<>1 and ShowType<>2 and ShowType<>3 then
		ShowType=1
	end if
	response.write "<table border='0' cellpadding='0' cellspacing='0'>"
	response.write "<form method='Get' name='SearchForm' action='" & Action & "'>"
	response.write "<tr><td height='28' align='center'>"
	if ShowType=1 then
		response.write "<input type='text' name='keyword'  size='15' value='�ؼ���' maxlength='50' onFocus='this.select();'>&nbsp;"
		response.write "<input type='hidden' name='field' value='Title'>"
		response.write "<input type='submit' name='Submit'  value='����'>"
		'response.write "<br><br>�߼�����"
	elseif Showtype=2 then
		response.write "<select name='Field' size='1'>"
    	if ChannelID=2 then
			response.write "<option value='Title' selected>���±���</option>"
			response.write "<option value='Content'>��������</option>"
			response.write "<option value='Author'>��������</option>"
			response.write "<option value='Editor'>�༭����</option>"
		elseif ChannelID=3 then	
			response.write "<option value='SoftName' selected>�������</option>"
			response.write "<option value='SoftIntro'>������</option>"
			response.write "<option value='Author'>�������</option>"
			response.write "<option value='Editor'>�༭����</option>"
		elseif ChannelID=4 then	
			response.write "<option value='PhotoName' selected>ͼƬ����</option>"
			response.write "<option value='PhotoIntro'>ͼƬ���</option>"
			response.write "<option value='Author'>ͼƬ����</option>"
			response.write "<option value='Editor'>�༭����</option>"
		else
			response.write "<option value='Title' selected>���±���</option>"
			response.write "<option value='Content'>��������</option>"
			response.write "<option value='Author'>��������</option>"
			response.write "<option value='Editor'>�༭����</option>"
		end if
		response.write "</select>&nbsp;"
		response.write "<select name='ClassID'><option value=''>������Ŀ</option>"
		call Admin_ShowClass_Option(5,0)
		response.write "</select>&nbsp;<input type='text' name='keyword'  size='20' value='�ؼ���' maxlength='50' onFocus='this.select();'>&nbsp;"
		response.write "<input type='submit' name='Submit'  value=' ���� '>"
	elseif Showtype=3 then
	
	end if
	response.write "</td></tr></form></table>"
end sub

'==================================================
'��������ShowGuest
'��  �ã���ʾ��վ����
'��  ����GuestTitleLen ---��ʾ���Ա��ⳤ��
'		 GuestItemNum  ---��ʾ��������
'==================================================
sub ShowGuest(GuestTitleLen,GuestItemNum)
 	dim sqlGuest,rsGuest
 	if GuestItemNum<=0 or GuestItemNum>50 then
 		GuestItemNum=10
	end if
 	sqlGuest="select top " & GuestItemNum & " * from Guest where GuestIsPassed=True order by GuestMaxId desc"
 	Set rsGuest= Server.CreateObject("ADODB.Recordset")
 	rsGuest.open sqlGuest,conn,1,1
 	if rsGuest.bof and rsGuest.eof then 
  		response.Write " û���κ�����"
 	else
		do while Not rsGuest.eof
			response.write "<font color=#b70000><b>��</b></font><a href='guestbook.asp' "
			response.write " title='���⣺" & rsGuest("GuestTitle") & vbcrlf & "������" & rsGuest("GuestName") & vbcrlf & "ʱ�䣺" & rsGuest("GuestDatetime") &"'"
			response.write " target='_blank'>"
			response.write gotTopic(rsGuest("GuestTitle"),GuestTitleLen)
			response.write "</a><br>"
			rsGuest.movenext
  		Loop
 	end if
 	rsGuest.close
 	set rsGuest=nothing
end sub

'==================================================
'��������ShowAD
'��  �ã���ʾ���
'��  ����ADType ---�������
'==================================================
sub ShowAD(ADType)
	dim sqlAD,rsAD,AD,arrSetting,popleft,poptop,floatleft,floattop,fixedleft,fixedtop
	sqlAD="select * from Advertisement where IsSelected=True"
	sqlAD=sqlAD & " and (ChannelID=0 or ChannelID=" & ChannelID & ")"
	sqlAD=sqlAD & " and ADType=" & ADtype & " order by ID Desc"
	set rsAD=server.createobject("adodb.recordset")
	rsAD.open sqlAD,conn,1,1
	if not rsAd.bof and not rsAD.eof then
		do while not rsAD.eof
			if rsAD("isflash")=true then
				AD= "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0'"
				if rsAD("ImgWidth")>0 then AD = AD & " width='" & rsAD("ImgWidth") & "'"
				if rsAD("ImgHeight")>0 then AD = AD & " height='" & rsAD("ImgHeight") & "'"
				AD = AD & "><param name='movie' value='" & rsAD("ImgUrl") & "'><param name='quality' value='high'><embed src='" & rsAD("ImgUrl") & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'"
				if rsAD("ImgWidth")>0 then AD = AD & " width='" & rsAD("ImgWidth") & "'"
				if rsAD("ImgHeight")>0 then AD = AD & " height='" & rsAD("ImgHeight") & "'"
				AD = AD & "></embed></object>"
			else
				AD ="<a href='" & rsAD("SiteUrl") & "' target='_blank' title='" & rsAD("SiteName") & "��" & rsAD("SiteUrl") & "'><img src='" & rsAD("ImgUrl") & "'"
				if rsAD("ImgWidth")>0 then AD = AD & " width='" & rsAD("ImgWidth") & "'"
				if rsAD("ImgHeight")>0 then AD = AD & " height='" & rsAD("ImgHeight") & "'"
				AD = AD & " border='0'></a>"
			end if
			if ADtype=0 then
				if  session("PopAD"&rsAD("ID")&ChannelID)<>True then
					if instr(rsAD("ADSetting"),"|")>0 then
						arrSetting=split(rsAD("ADSetting"),"|")
						popleft=arrsetting(0)
						poptop=arrsetting(1)
					end if
					response.write "<SCRIPT language=javascript>"
					response.write "window.open(""PopAD.asp?Id="& rsAD("ID")&""",""popad"&rsAD("ID")&""",""toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,width="&rsAD("ImgWidth")&",height="&rsAD("ImgHeight")&",top="&poptop&",left="&popleft&""");"
					response.write "</SCRIPT>"
					session("PopAD"&rsAD("ID")&ChannelID)=True
				end if
			elseif ADtype=1 then
				response.write AD
				exit do
			elseif ADtype=2 then
				response.write AD
				exit do
			elseif ADtype=3 then
				response.write AD
				exit do
			elseif ADtype=4 then
				if instr(rsAD("ADSetting"),"|")>0 then
					arrSetting=split(rsAD("ADSetting"),"|")
					floatleft=arrsetting(0)
					floattop=arrsetting(1)
				end if
				response.write "<div id='FlAD' style='position:absolute; z-index:10;left: "&floatleft&"; top: "&floattop&"'>" & AD & "</div>"
				call FloatAD()
				exit do
			elseif ADtype=5 then
				if instr(rsAD("ADSetting"),"|")>0 then
					arrSetting=split(rsAD("ADSetting"),"|")
					fixedleft=arrsetting(0)
					fixedtop=arrsetting(1)
				end if
				response.write "<div id='FixAD' style='position:absolute; z-index:10;left: "&fixedleft&"; top: "&fixedtop&"'>" & AD & "</div>"
				call FixedAD()
				exit do
			end if
			rsAD.movenext
		loop
	end if
	rsAD.close
	set rsAD=nothing
end sub

'==================================================
'��������FloatAD
'��  �ã��������
'��  ������
'==================================================
sub FloatAD()
%>
<SCRIPT language=javascript>
<!--moving logo-->
window.onload=FlAD;
var brOK=false;
var mie=false;
var aver=parseInt(navigator.appVersion.substring(0,1));
var aname=navigator.appName;
var mystop=0;

function checkbrOK()
{if(aname.indexOf("Internet Explorer")!=-1)
{if(aver>=4) brOK=navigator.javaEnabled();
mie=true;
}
if(aname.indexOf("Netscape")!=-1)  
{if(aver>=4) brOK=navigator.javaEnabled();}
}
var vmin=2;
var vmax=5;
var vr=2;
var timer1;

function Chip(chipname,width,height)
{this.named=chipname;
this.vx=vmin+vmax*Math.random();
this.vy=vmin+vmax*Math.random();
this.w=width;
this.h=height;
this.xx=0;
this.yy=0;
this.timer1=null;
}

function movechip(chipname)
{
if(brOK && mystop==0)
{eval("chip="+chipname);
if(!mie)
{pageX=window.pageXOffset;
pageW=window.innerWidth;
pageY=window.pageYOffset;
pageH=window.innerHeight;
}
else
{pageX=window.document.body.scrollLeft;
pageW=window.document.body.offsetWidth-8;
pageY=window.document.body.scrollTop;
pageH=window.document.body.offsetHeight;
} 
chip.xx=chip.xx+chip.vx;
chip.yy=chip.yy+chip.vy;
chip.vx+=vr*(Math.random()-0.5);
chip.vy+=vr*(Math.random()-0.5);
if(chip.vx>(vmax+vmin))  chip.vx=(vmax+vmin)*2-chip.vx;
if(chip.vx<(-vmax-vmin)) chip.vx=(-vmax-vmin)*2-chip.vx;
if(chip.vy>(vmax+vmin))  chip.vy=(vmax+vmin)*2-chip.vy;
if(chip.vy<(-vmax-vmin)) chip.vy=(-vmax-vmin)*2-chip.vy;
if(chip.xx<=pageX)
{chip.xx=pageX;
chip.vx=vmin+vmax*Math.random();
}
if(chip.xx>=pageX+pageW-chip.w)
{chip.xx=pageX+pageW-chip.w;
chip.vx=-vmin-vmax*Math.random();
}
if(chip.yy<=pageY)
{chip.yy=pageY;
chip.vy=vmin+vmax*Math.random();
}
if(chip.yy>=pageY+pageH-chip.h)
{chip.yy=pageY+pageH-chip.h;
chip.vy=-vmin-vmax*Math.random();
}
if(!mie)
{eval('document.'+chip.named+'.top ='+chip.yy);
eval('document.'+chip.named+'.left='+chip.xx);
} 
else
{eval('document.all.'+chip.named+'.style.pixelLeft='+chip.xx);
eval('document.all.'+chip.named+'.style.pixelTop ='+chip.yy); 
}
	chip.timer1=setTimeout("movechip('"+chip.named+"')",100);
}
}
function stopme(x)
{
brOk=true;
mystop=x;
movechip("FlAD");
}
var FlAD;
var chip;
function FlAD()
{checkbrOK(); 
FlAD=new Chip("FlAD",80,80);
if(brOK) 
{ movechip("FlAD");
}
}
ns4=(document.layers)?true:false;
ie4=(document.all)?true:false;

function cncover()
{
if(ns4){
	//document.cnc.left=window.innerWidth/2-400;
	document.FlAD.visibility="hide";
	eval('document.cnc.left=document.'+chip.named+'.left');
	eval('document.cnc.top=document.'+chip.named+'.top-15');	
	document.cnc.visibility="show";
	}else if(ie4) 
	{
	document.all.FlAD.style.visibility="hidden";
	//document.all.cnc.style.left=window.document.body.offsetWidth/2-400;
	document.all.cnc.style.left=parseInt(document.all.FlAD.style.left)-0;
	document.all.cnc.style.top=parseInt(document.all.FlAD.style.top)-0;	
	document.all.cnc.style.visibility="visible";
	stopme(1);
	}
}

function cncout()
{
if(ns4){
	document.cnc.visibility="hide";
	document.FlAD.visibility="show";
	}else if(ie4) 
	{
	document.all.cnc.style.visibility="hidden";
	document.all.FlAD.style.visibility="visible";
	stopme(0);
	}
}
</script>
<%
end sub


'==================================================
'��������FixedAD
'��  �ã��̶�λ�ù��
'��  ������
'==================================================
sub FixedAD()
%>
<script LANGUAGE="JavaScript">
<!-- Begin
var imgheight
var imgleft
document.ns = navigator.appName == "Netscape"
if (navigator.appName == "Netscape")
{
imgheight=document.FixAD.pageY
imgleft=document.FixAD.pageX
}
else
{
imgheight=600-parseInt(FixAD.style.top)
imgleft=parseInt(FixAD.style.left)
}
myload()
function myload()
{
if (navigator.appName == "Netscape")
{document.FixAD.pageY=pageYOffset+window.innerHeight-imgheight;
document.FixAD.pageX=imgleft;
leftmove();
}
else
{
FixAD.style.top=document.body.scrollTop+document.body.offsetHeight-imgheight;
FixAD.style.left=imgleft;
leftmove();
}
}
function leftmove()
 {
 if(document.ns)
 {
 document.FixAD.top=pageYOffset+window.innerHeight-imgheight
 document.FixAD.left=imgleft;
 setTimeout("leftmove();",50)
 }
 else
 {
 FixAD.style.top=document.body.scrollTop+document.body.offsetHeight-imgheight;
 FixAD.style.left=imgleft;
 setTimeout("leftmove();",50)
 }
 }

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true)
//  End -->
</script>
<%
end sub

'==================================================
'��������FixedAD1
'��  �ã��̶�λ�ù�棨ͼƬλ�ó�������ʱ��ʱ�����⣩
'��  ������
'==================================================
sub FixedAD1()
%>
<script LANGUAGE="JavaScript">
<!-- Begin
        self.onError=null;
        currentX = currentY = 0;
        whichIt = null;
        lastScrollX = 0; lastScrollY = 0;
        NS = (document.layers) ? 1 : 0;
        IE = (document.all) ? 1: 0;
        function heartBeat() {
                if(IE) { diffY = document.body.scrollTop; diffX = document.body.scrollLeft; }
            if(NS) { diffY = self.pageYOffset; diffX = self.pageXOffset; }
                if(diffY != lastScrollY) {
                        percent = .1 * (diffY - lastScrollY);
                        if(percent > 0) percent = Math.ceil(percent);
                        else percent = Math.floor(percent);
                                        if(IE) document.all.FixAD.style.pixelTop += percent;
                                        if(NS) document.FixAD.top += percent;
                        lastScrollY = lastScrollY + percent;
            }
                if(diffX != lastScrollX) {
                        percent = .1 * (diffX - lastScrollX);
                        if(percent > 0) percent = Math.ceil(percent);
                        else percent = Math.floor(percent);
                        if(IE) document.all.FixAD.style.pixelLeft += percent;
                        if(NS) document.FixAD.left += percent;
                        lastScrollX = lastScrollX + percent;
                }
        }
        function checkFocus(x,y) {
                stalkerx = document.FixAD.pageX;
                stalkery = document.FixAD.pageY;
                stalkerwidth = document.FixAD.clip.width;
                stalkerheight = document.FixAD.clip.height;
                if( (x > stalkerx && x < (stalkerx+stalkerwidth)) && (y > stalkery && y < (stalkery+stalkerheight))) return true;
                else return false;
        }
        function grabIt(e) {
                if(IE) {
                        whichIt = event.srcElement;
                        while (whichIt.id.indexOf("FixAD") == -1) {
                                whichIt = whichIt.parentElement;
                                if (whichIt == null) { return true; }
                    }
                        whichIt.style.pixelLeft = whichIt.offsetLeft;
                    whichIt.style.pixelTop = whichIt.offsetTop;
                        currentX = (event.clientX + document.body.scrollLeft);
                           currentY = (event.clientY + document.body.scrollTop);
                } else {
                window.captureEvents(Event.MOUSEMOVE);
                if(checkFocus (e.pageX,e.pageY)) {
                        whichIt = document.FixAD;
                        StalkerTouchedX = e.pageX-document.FixAD.pageX;
                        StalkerTouchedY = e.pageY-document.FixAD.pageY;
                }
                }
            return true;
        }
        function moveIt(e) {
                if (whichIt == null) { return false; }
                if(IE) {
                    newX = (event.clientX + document.body.scrollLeft);
                    newY = (event.clientY + document.body.scrollTop);
                    distanceX = (newX - currentX);    distanceY = (newY - currentY);
                    currentX = newX;    currentY = newY;
                    whichIt.style.pixelLeft += distanceX;
                    whichIt.style.pixelTop += distanceY;
                        if(whichIt.style.pixelTop < document.body.scrollTop) whichIt.style.pixelTop = document.body.scrollTop;
                        if(whichIt.style.pixelLeft < document.body.scrollLeft) whichIt.style.pixelLeft = document.body.scrollLeft;
                        if(whichIt.style.pixelLeft > document.body.offsetWidth - document.body.scrollLeft - whichIt.style.pixelWidth - 20) whichIt.style.pixelLeft = document.body.offsetWidth - whichIt.style.pixelWidth - 20;
                        if(whichIt.style.pixelTop > document.body.offsetHeight + document.body.scrollTop - whichIt.style.pixelHeight - 5) whichIt.style.pixelTop = document.body.offsetHeight + document.body.scrollTop - whichIt.style.pixelHeight - 5;
                        event.returnValue = false;
                } else {
                        whichIt.moveTo(e.pageX-StalkerTouchedX,e.pageY-StalkerTouchedY);
                if(whichIt.left < 0+self.pageXOffset) whichIt.left = 0+self.pageXOffset;
                if(whichIt.top < 0+self.pageYOffset) whichIt.top = 0+self.pageYOffset;
                if( (whichIt.left + whichIt.clip.width) >= (window.innerWidth+self.pageXOffset-17)) whichIt.left = ((window.innerWidth+self.pageXOffset)-whichIt.clip.width)-17;
                if( (whichIt.top + whichIt.clip.height) >= (window.innerHeight+self.pageYOffset-17)) whichIt.top = ((window.innerHeight+self.pageYOffset)-whichIt.clip.height)-17;
                return false;
                }
            return false;
        }
        function dropIt() {
                whichIt = null;
            if(NS) window.releaseEvents (Event.MOUSEMOVE);
            return true;
        }
        if(NS) {
                window.captureEvents(Event.MOUSEUP|Event.MOUSEDOWN);
                window.onmousedown = grabIt;
                 window.onmousemove = moveIt;
                window.onmouseup = dropIt;
        }
        if(IE) {
                document.onmousedown = grabIt;
                 document.onmousemove = moveIt;
                document.onmouseup = dropIt;
        }
        if(NS || IE) action = window.setInterval("heartBeat()",1);
//  End -->
</script>
<%
end sub
%>