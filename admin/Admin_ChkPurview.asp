<%
dim AdminName,AdminPurview,PurviewPassed
dim AdminPurview_Article,AdminPurview_Soft,AdminPurview_Photo,AdminPurview_Guest,AdminPurview_Others
dim rsGetAdmin,sqlGetAdmin
dim ComeUrl,cUrl
ComeUrl=lcase(trim(request.ServerVariables("HTTP_REFERER")))
if ComeUrl="" then
        response.redirect "login.asp"
	'response.write "<br><p align=center><font color='red'>对不起，为了系统安全，不允许直接输入地址访问本系统的后台管理页面。</font></p>"
	response.end
else
	cUrl=trim("http://" & Request.ServerVariables("SERVER_NAME"))
	if mid(ComeUrl,len(cUrl)+1,1)=":" then
		cUrl=cUrl & ":" & Request.ServerVariables("SERVER_PORT")
	end if
	cUrl=lcase(cUrl & request.ServerVariables("SCRIPT_NAME"))
	if lcase(left(ComeUrl,instrrev(ComeUrl,"/")))<>lcase(left(cUrl,instrrev(cUrl,"/"))) then
                response.redirect "login.asp"
		'response.write "<br><p align=center><font color='red'>对不起，为了系统安全，不允许从外部链接地址访问本系统的后台管理页面。</font></p>"
		response.end
	end if
end if

AdminName=replace(session("AdminName"),"'","")
if AdminName="" then
	call CloseConn()
	response.redirect "login.asp"
end if
sqlGetAdmin="select * from Admin where UserName='" & AdminName & "'"
set rsGetAdmin=server.CreateObject("adodb.recordset")
rsGetAdmin.open sqlGetAdmin,conn,1,1
if rsGetAdmin.bof and rsGetAdmin.eof then
	rsGetAdmin.close
	set rsGetAdmin=nothing
	call CloseConn()
	response.redirect "login.asp"
end if
AdminPurview=rsGetAdmin("Purview")
AdminPurview_Article=rsGetAdmin("AdminPurview_Article")
AdminPurview_Soft=rsGetAdmin("AdminPurview_Soft")
AdminPurview_Photo=rsGetAdmin("AdminPurview_Photo")
AdminPurview_Guest=rsGetAdmin("AdminPurview_Guest")
AdminPurview_Others=rsGetAdmin("AdminPurview_Others")
rsGetAdmin.close
set rsGetAdmin=nothing
PurviewPassed=True
if PurviewLevel>0 then
	if AdminPurview>PurviewLevel then
		PurviewPassed=False
	else
		if AdminPurview=2 then
			select case CheckChannelID
				case 0        '其他管理操作
					PurviewPassed=CheckPurview(AdminPurview_Others,PurviewLevel_Others)
				case 2        '文章频道
					if AdminPurview_Article>PurviewLevel_Article then
						PurviewPassed=False
					end if
				case 3       '下载频道
					if AdminPurview_Soft>PurviewLevel_Soft then
						PurviewPassed=False
					end if
				case 4       '图片频道
					if AdminPurview_Photo>PurviewLevel_Photo then
						PurviewPassed=False
					end if
				case 5       '留言板
					if AdminType=True then
						PurviewPassed=CheckPurview(AdminPurview_Guest,PurviewLevel_Guest)
					else
						PurviewPassed=True
					end if
			end select
		end if
	end if
end if
if PurviewPassed=False then
	response.write "<br><p align=center><font color='red'>对不起，你没有此项操作的权限。</font></p>"
	response.end
end if

function CheckPurview(AllPurviews,strPurview)
	if isNull(AllPurviews) or AllPurviews="" or strPurview="" then
		CheckPurview=False
		exit function
	end if
	CheckPurview=False
	if instr(AllPurviews,",")>0 then
		dim arrPurviews,i
		arrPurviews=split(AllPurviews,",")
		for i=0 to ubound(arrPurviews)
			if trim(arrPurviews(i))=strPurview then
				CheckPurview=True
				exit for
			end if
		next
	else
		if AllPurviews=strPurview then
			CheckPurview=True
		end if
	end if
end function

function CheckClassMaster(AllMaster,MasterName)
	if isNull(AllMaster) or AllMaster="" or MasterName="" then
		CheckClassMaster=False
		exit function
	end if
	CheckClassMaster=False
	if instr(AllMaster,"|")>0 then
		dim arrMaster,i
		arrMaster=split(AllMaster,"|")
		for i=0 to ubound(arrMaster)
			if trim(arrMaster(i))=MasterName then
				CheckClassMaster=True
				exit for
			end if
		next
	else
		if AllMaster=MasterName then
			CheckClassMaster=True
		end if
	end if
end function

function GetURL(strURL)
dim node,i,nodecount,str_text
set xml = CreateObject("Microsoft.XMLDOM")
xml.async = false
xml.load(strURL)
set root = xml.documentElement
set nodeLis = root.childNodes
nodeCount = nodeLis.length
str_text=""
For i=1 to nodeCount
set node = nodeLis.nextNode()
if node.selectSingleNode("urlname").text<>"" then
str_text=str_text&"<a herf="&node.selectSingleNode("urlname").text&"  target=_blank>"&node.selectSingleNode("name").text&"</a>"
else
str_text=str_text&node.selectSingleNode("name").text
end if
Next
response.write str_text
end function
%>
