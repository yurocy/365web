<%
sub Admin_ShowRootClass()
	dim sqlRoot,rsRoot
	sqlRoot="select ClassID,ClassName,RootID,Child From MenuClass where ParentID=0 order by RootID"
	Set rsRoot= Server.CreateObject("ADODB.Recordset")
	rsRoot.open sqlRoot,conn,1,1
	if rsRoot.bof and rsRoot.eof then 
		response.Write("��û���κ���Ŀ�������������Ŀ��")
	else
		response.write "|&nbsp;"
		do while not rsRoot.eof
			if rsRoot(2)=RootID then
				response.Write("<a href='" & FileName & "?ClassID=" & rsRoot(0) & "'><font color=red>" & rsRoot(1) & "</font></a> | ")
				tID=rsRoot(0)
				tChild=rsRoot(3)
			else
				response.Write("<a href='" & FileName & "?ClassID=" & rsRoot(0) & "'>" & rsRoot(1) & "</a> | ")
			end if
			rsRoot.movenext
		loop
	end if
	rsRoot.close
	set rsRoot=nothing
end sub

sub Admin_ShowClass_Option(ShowType,CurrentID)
	if ShowType=0 then
	    response.write "<option value='0'"
		if CurrentID=0 then response.write " selected"
		response.write ">��(һ����Ŀ)</option>"
	    response.write "<option value='-1'"
		response.write ">��(������Ŀ)</option>"

	end if
	dim rsClass,sqlClass,strTemp,tmpDepth,i
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	sqlClass="Select * From MenuClass where parentid>=0 order by RootID,OrderID"
	set rsClass=server.CreateObject("adodb.recordset")
	rsClass.open sqlClass,conn,1,1
	if rsClass.bof and rsClass.bof then
		response.write "<option value=''>���������Ŀ</option>"
	else
		dim UserLevel
		UserLevel=request.Cookies("asp163")("UserLevel")
		if UserLevel="" then
			UserLevel=5000
		else
			UserLevel=Cint(UserLevel)
		end if
		do while not rsClass.eof
			tmpDepth=rsClass("Depth")
			if rsClass("NextID")>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
			if ShowType=1 then
				if rsClass("LinkUrl")<>"" then
					strTemp="<option value=''"
				else
					strTemp="<option value='" & rsClass("ClassID") & "'"
				end if
				if AdminPurview=2 and AdminPurview_Article=3 then
					if CheckClassMaster(rsClass("ClassChecker"),AdminName)=True then
						strTemp=strTemp & "style='background-color:#ff0000'"
					end if
				end if
			elseif ShowType=2 then
				if rsClass("LinkUrl")<>"" then
					strTemp="<option value=''"
				else
					strTemp="<option value='" & rsClass("ClassID") & "'"
				end if
				if AdminPurview=2 and AdminPurview_Article=3 then
					if CheckClassMaster(rsClass("ClassMaster"),AdminName)=True then
						strTemp=strTemp & "style='background-color:#ff0000'"
					end if
				end if
			elseif ShowType=3 then
				if rsClass("Child")>0 then
					strTemp="<option value=''"
				elseif rsClass("LinkUrl")<>"" then
					strTemp="<option value='0'"
				else
					strTemp="<option value='" & rsClass("ClassID") & "'"
				end if
				if AdminPurview=2 and AdminPurview_Article=3 then
					if CheckClassMaster(rsClass("ClassInputer"),AdminName)=True then
						strTemp=strTemp & "style='background-color:#ff0000'"
					end if
				end if
			elseif ShowType=4 then
				if rsClass("Child")>0 then
					strTemp="<option value=''"
				elseif rsClass("LinkUrl")<>"" then
					strTemp="<option value='0'"
				elseif rsClass("AddPurview")<UserLevel then
					strTemp="<option value='-1'"
				else
					strTemp="<option value='" & rsClass("ClassID") & "'"
				end if
			else
				strTemp="<option value='" & rsClass("ClassID") & "'"
			end if
			if CurrentID>0 and rsClass("ClassID")=CurrentID then
				 strTemp=strTemp & " selected"
				 SkinID=rsClass("SkinID")
				 LayoutID=rsClass("LayoutID")
				 BrowsePurview=rsClass("BrowsePurview")
				 AddPurview=rsClass("AddPurview")
			end if
			strTemp=strTemp & ">"
			
			if tmpDepth>0 then
				for i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					if i=tmpDepth then
						if rsClass("NextID")>0 then
							strTemp=strTemp & "��&nbsp;"
						else
							strTemp=strTemp & "��&nbsp;"
						end if
					else
						if arrShowLine(i)=True then
							strTemp=strTemp & "��"
						else
							strTemp=strTemp & "&nbsp;"
						end if
					end if
				next
			end if
			strTemp=strTemp & rsClass("ClassName")
			if ShowType=4 and rsClass("AddPurview")<UserLevel then
				strTemp=strTemp & " *"
			end if
			strTemp=strTemp & "</option>"
			response.write strTemp
			rsClass.movenext
		loop
	end if
	rsClass.close
	set rsClass=nothing
end sub

sub Dzwang_ShowClass_Option(ShowType,CurrentID)
dim sqlClass,strTemp
if ShowType=3 then
    dim rsa
    strTemp=""
    sqlClass="Select * From MenuClass where ClassType='�ĵ�' order by RootID,OrderID"
    set rsa=server.CreateObject("adodb.recordset")
    rsa.open sqlClass,conn,1,1
    do while not rsa.eof
       strTemp=strTemp&"<option value='" & rsa("ClassID") & "'>"&rsa("ClassName")&"</option>"       
       rsa.movenext
    loop
    rsa.close
    set rsa=nothing
end if
response.write strTemp
end sub


sub Admin_ShowPath(RootName)
	response.write "�����ڵ�λ�ã�&nbsp;<a href='" & FileName & "'>" & RootName & "</a>&nbsp;&gt;&gt;&nbsp;"
	if ClassID>0 then
		if ParentID>0 then
			dim sqlPath,rsPath
			sqlPath="select ClassID,ClassName From MenuClass where ClassID in (" & ParentPath & ") order by Depth"
			set rsPath=server.createobject("adodb.recordset")
			rsPath.open sqlPath,conn,1,1
			do while not rsPath.eof
				response.Write "<a href='" & FileName & "?ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
				rsPath.movenext
			loop
			rsPath.close
			set rsPath=nothing
		end if
		response.write "<a href='" & FileName & "?ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
	end if
	if ManageType="MyArticle" then
		response.write "<font color=red>" & AdminName & "</font> ���ĵ�"
	else
		if keyword="" then
			response.write "�����ĵ�"
		else
			select case strField
				case "Title"
					response.Write "�����к��� <font color=red>"&keyword&"</font> ���ĵ�"
				case "Content"
					response.Write "���ݺ��� <font color=red>"&keyword&"</font> ���ĵ�"
				case "Author"
					response.Write "���������к��� <font color=red>"&keyword&"</font> ���ĵ�"
				case "Editor"
					response.write "�༭�����к��� <font color=red>" & keyword & "</font> ���ĵ�"
				case else
					response.Write "�����к��� <font color=red>"&keyword&"</font> ���ĵ�"
			end select
		end if
	end if
end sub

sub Admin_ShowSpecial_Option(ShowType,SpecialID)
	dim UserLevel
	UserLevel=request.Cookies("asp163")("UserLevel")
	if UserLevel="" then
		UserLevel=5000
	else
		UserLevel=Cint(UserLevel)
	end if
	response.write "<select name='SpecialID' id='SpecialID'><option value=''"
	if SpecialID=0 then
		response.write " selected"
	end if
	response.write ">�������κ�ר��</option>"
	                
	dim sqlSpecial,rsSpecial
    if ShowType=1 then
		sqlSpecial = "select * from Special"
	else
		sqlSpecial="select * from Special where AddPurview>=" & UserLevel
	end if	
	set rsSpecial=server.CreateObject("adodb.recordset")
	rsSpecial.open sqlSpecial,conn,1,1
	do while not rsSpecial.eof
		if rsSpecial("SpecialID")=SpecialID then
			response.write "<option value='" & rsSpecial("SpecialID") & "' selected>" & rsSpecial("SpecialName") & "</option>"
		else
			response.write "<option value='" & rsSpecial("SpecialID") & "'>" & rsSpecial("SpecialName") & "</option>"
		end if
		rsSpecial.movenext
	loop
	rsSpecial.close
    set rsSpecial = nothing
end sub

sub Admin_ShowChild()
	dim sqlChild,rsChild
	sqlChild="select ClassID,ClassName,Child From MenuClass where ParentID=" & ClassID & " order by OrderID"
	Set rsChild= Server.CreateObject("ADODB.Recordset")
	rsChild.open sqlChild,conn,1,1
	i=0
	do while not rsChild.eof
		response.Write "&nbsp;&nbsp;<a href='" & FileName & "?ClassID=" & rsChild(0) & "'>" & rsChild(1) & "</a>"
		if rsChild(2)>0 then
			response.write "(" & rsChild(2) & ")"
		else
			if ChildID="" then
				ChildID=Cstr(rsChild(0))
			else
				ChildID=ChildID & "," & Cstr(rsChild(0))
			end if
		end if		
		rsChild.movenext
		i=i+1
		if i mod 8=0 then
			response.write "<br>"
		else
			response.write "&nbsp;&nbsp;"
		end if
	loop
	rsChild.close
	set rsChild=nothing
end sub


sub Admin_ShowChild2()
	response.write "<table width='100%' border='0' align='center' cellpadding='5' cellspacing='1'>"
	response.write "  <tr class='tdbg'>"
	dim sqlChild,rsChild,rsChild2
	sqlChild="select ClassID,ClassName,Child From MenuClass where ParentID=" & ClassID
	Set rsChild= Server.CreateObject("ADODB.Recordset")
	set rsChild2= Server.CreateObject("ADODB.Recordset")
	rsChild.open sqlChild,conn,1,1
	i=0
	do while not rsChild.eof
		response.Write "<td width='50%' valign='top'><a href='" & FileName & "?ClassID=" & rsChild(0) & "'><b><font color=red>" & rsChild(1) & "</b></font></a>"
		if rsChild(2)>0 then
			response.write "(" & rsChild(2) & ")<br>"
			sqlChild="select ClassID,ClassName,Child From MenuClass where ParentID=" & rsChild(0)
			rsChild2.open sqlChild,conn,1,1
			j=0
			do while not rsChild2.eof
				response.Write "<a href='" & FileName & "?ClassID=" & rsChild2(0) & "'>" & rsChild2(1) & "</a>"
				if rsChild2(2)>0 then
					response.write "(" & rsChild2(2) & ")"
				end if
				response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
				rsChild2.movenext
				j=j+1
				if j mod 5=0 then response.write "<br>"
			loop
			rsChild2.close
		else
			if ChildID="" then
				ChildID=Cstr(rsChild(0))
			else
				ChildID=ChildID & "," & Cstr(rsChild(0))
			end if
		end if		
		rsChild.movenext
		i=i+1
		response.write "</td>"
		if i mod 2=0 then
			response.write "</tr><tr class='tdbg'>"
		end if
	loop
	if i mod 2<>0 then response.write "<td>&nbsp;</td>"
	
	rsChild.close
	set rsChild=nothing
	set rsChild2=nothing
  	response.write "</tr></table>"	
end sub

sub Admin_ShowPath2(strParentPath,strClassName,iDepth)
	if iDepth<=0 then
		response.write strClassName
		exit sub
	end if
	dim sqlPath,rsPath,i
	sqlPath="select * From MenuClass where ClassID in (" & strParentPath & ") order by Depth"
	set rsPath=server.createobject("adodb.recordset")
	rsPath.open sqlPath,conn,1,1
	do while not rsPath.eof
		for i=1 to rsPath("Depth")
			response.write "&nbsp;&nbsp;&nbsp;"
		next
		if rsPath("Depth")>0 then
			response.write "��"
		end if
		response.Write rsPath("ClassName") & "<br>"
		rsPath.movenext
	loop
	rsPath.close
	set rsPath=nothing
	if iDepth>0 and strClassName<>"" then
		for i=1 to iDepth
			response.write "&nbsp;&nbsp;&nbsp;"
		next
		response.write "��" & strClassName
	end if
end sub

sub AddMaster(ClassMaster)
	dim arrClassMaster,rsAdmin
	if instr(ClassMaster,"|")>0 then
		arrClassMaster=split(ClassMaster,"|")
		ClassMaster=""
		for i=0 to ubound(arrClassMaster)
			set rsAdmin=conn.execute("select * from Admin where UserName='" & arrClassMaster(i) & "'")
			if rsAdmin.bof and rsAdmin.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>����Ա��" & arrClassMaster(i) & "�������ڣ��Ƿ�������ˣ�</li>"
			else
				if rsAdmin("Purview")>4 then
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>��" & arrClassMaster(i) & "��Ȩ�޲�����������Ϊ��Ŀ�༭��"
				else
					if rsAdmin("Purview")=4 then
						if ClassMaster="" then
							ClassMaster=arrClassMaster(i)
						else
							ClassMaster=ClassMaster & "|" & arrClassMaster(i)
						end if
					end if
				end if
			end if
		next
	else
		set rsAdmin=conn.execute("select * from Admin where UserName='" & ClassMaster & "'")
		if rsAdmin.bof and rsAdmin.eof then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>����Ա��" & ClassMaster & "�������ڣ��Ƿ�������ˣ�</li>"
		else
			if rsAdmin("Purview")>4 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>��" & ClassMaster & "��Ȩ�޲�����������Ϊ��Ŀ�༭��"
			else
				if rsAdmin("Purview")<4 then
					ClassMaster=""
				end if
			end if
		end if
	end if
end sub


sub Admin_ShowSearchForm(Action,ShowType)
	if ShowType<>1 and ShowType<>2 and ShowType<>3 then
		ShowType=1
	end if
	response.write "<table border='0' cellpadding='0' cellspacing='0'>"
	response.write "<form method='Get' name='SearchForm' action='" & Action & "'>"
	response.write "<tr><td height='28' align='center'>"
	if ShowType=1 then
		response.write "<input type='hidden' name='field' value='Title'>"
		response.write "<input type='text' name='keyword'  size='20' value='�ؼ���' maxlength='50' onFocus='this.select();'>"
		response.write "<input type='submit' name='Submit' class='button' value='����' style='border-style: solid; border-width: 1;background-color: #419010;'>"
		response.write "<br>�߼�����"
	elseif Showtype=2 then
		response.write "<select name='Field' size='1'>"
    	response.write "<option value='Title' selected>�ĵ�����</option>"
	    response.write "<option value='Content'>�ĵ�����</option>"
    	response.write "<option value='Author'>�ĵ�����</option>"
    	if AdminPurview=1 or AdminPurview_Article<=2 then	
			response.write "<option value='Editor'>�༭����</option>"
		end if
		response.write "</select>"
		response.write "<select name='ClassID'><option value=''>������Ŀ</option>"
		call Admin_ShowClass_Option(1,0)
		response.write "<input type='text' name='keyword' class='Input'  size='20' value='�ؼ���' maxlength='50' onFocus='this.select();'>"
		response.write "&nbsp;&nbsp;<input type='submit' name='Submit' class='button' value='����'>"
	end if
	response.write "</td></tr></form></table>"
end sub

sub DelFiles(strUploadFiles)
	if strUploadFiles="" then exit sub
	if DelUpFiles="Yes" and ObjInstalled=True then
		dim fso,arrUploadFiles,i
		Set fso = CreateObject("Scripting.FileSystemObject")
		if instr(strUploadFiles,"|")>1 then
			arrUploadFiles=split(strUploadFiles,"|")
			for i=0 to ubound(arrUploadFiles)
				if fso.FileExists(server.MapPath("" & arrUploadfiles(i))) then
					fso.DeleteFile(server.MapPath("" & arrUploadfiles(i)))
				end if
			next
		else
			if fso.FileExists(server.MapPath("" & strUploadfiles)) then
				fso.DeleteFile(server.MapPath("" & strUploadfiles))
			end if
		end if
		Set fso = nothing
	end if
end sub

sub Admin_ShowSkin_Option(SkinID)
	response.write "<select name='SkinID' id='SkinID'>"
	dim sqlSkin,rsSkin
	sqlSkin="select * from Skin"
	set rsSkin=server.CreateObject("adodb.recordset")
	rsSkin.open sqlSkin,conn,1,1
	if rsSkin.bof and rsSkin.eof then
	 	SkinCount=0
	  	response.write "<option value=''>���������Ŀ��ɫģ��</option>"
	else
	  	SkinCount=rsSkin.recordcount
	  	do while not rsSkin.eof
	  		if rsSkin("SkinID")=SkinID then
				response.write "<option value='" & rsSkin("SkinID") & "' selected>" & rsSkin("SkinName") & "</option>"
			else		
				response.write "<option value='" & rsSkin("SkinID") & "'>" & rsSkin("SkinName") & "</option>"
	  		end if		
			rsSkin.movenext
	  	loop
	end if
	rsSkin.close
	set rsSkin=nothing
    response.write "</select> <input name='PreviewSkin' type='button' id='PreviewSkin' value='�鿴Ч��ͼ' onclick=""alert('�˹�����������У������ڴ���');"">"
end sub

sub Admin_ShowLayout_Option(LayoutType,LayoutID)
	response.write "<select name='LayoutID' id='LayoutID'>"
	dim sqlLayout,rsLayout
	sqlLayout="select * from Layout where LayoutType=" & LayoutType
	set rsLayout=server.CreateObject("adodb.recordset")
	rsLayout.open sqlLayout,conn,1,1
	if rsLayout.bof and rsLayout.eof then
	  	LayoutCount=0
	  	response.write "<option value=''>���������Ŀ�������ģ��</option>"
	else
	  	LayoutCount=rsLayout.recordcount
	  	do while not rsLayout.eof
	  		if rsLayout("LayoutID")=LayoutID then
				response.write "<option value='" & rsLayout("LayoutID") & "' selected>" & rsLayout("LayoutName") & "</option>"
			else		
				response.write "<option value='" & rsLayout("LayoutID") & "'>" & rsLayout("LayoutName") & "</option>"
	  		end if		
			rsLayout.movenext
	  	loop
	end if
	rsLayout.close
	set rsLayout=nothing
    response.write "</select> <input name='PreviewLayout' type='button' id='PreviewLayout' value='�鿴Ч��ͼ' onclick=""alert('�˹�����������У������ڴ���');"">"
end sub
%>