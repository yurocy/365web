<!--#Include File="Dv_ClsMain.asp"-->
<%
Set MyBoardOnline=new Cls_UserOnlne 
Dvbbs.GetForum_Setting
Dvbbs.CheckUserLogin

Function checkXHTML(XMLstr)
	Dim XML,node
	Set xml=Server.Createobject("msxml2.DOMDocument"& MsxmlVersion)
	If xml.loadxml("<div>" & replace(XMLstr,"&","&amp;") &"</div>") Then
		checkXHTML=""
		If xml.documentElement.getElementsByTagName("link").length >0 Then
			checkXHTML="���ݺ�ǰ̨��ֹ�ύ�ı�ǩ""link"""
			Exit Function
		End If
		If xml.documentElement.getElementsByTagName("iframe").length >0 Then
			checkXHTML="���ݺ�ǰ̨��ֹ�ύ�ı�ǩ""iframe"""
			Exit Function
		End If
		If xml.documentElement.getElementsByTagName("meta").length >0 Then
			checkXHTML="���ݺ�ǰ̨��ֹ�ύ�ı�ǩ""meta"""
			Exit Function
		End If
		If xml.documentElement.getElementsByTagName("script").length >0 Then
			checkXHTML="���ݺ�ǰ̨��ֹ�ύ�ı�ǩ""script"""
			Exit Function
		End If
		If xml.documentElement.getElementsByTagName("object").length >0 Then
			checkXHTML="���ݺ�ǰ̨��ֹ�ύ�ı�ǩ""object"""
			Exit Function
		End If
		If xml.documentElement.getElementsByTagName("embed").length >0 Then
			checkXHTML="���ݺ�ǰ̨��ֹ�ύ�ı�ǩ""embed"""
			Exit Function
		End If
		'href��Ľű�
		For Each Node in xml.documentElement.selectNodes("//a[@href]")
			If InStr(LCase(Node.selectSingleNode("@href").text),"script:")>0  Then
				checkXHTML="���������к��Ƿ��Ľű�����"
				Exit For
			End If
		Next
		If checkXHTML<>"" Then Exit Function
		'����src��Ľű�
		For Each Node in xml.documentElement.selectNodes("//*[@src]")
			If InStr(LCase(Node.selectSingleNode("@src").text),"script:")>0  Then
				checkXHTML="ͼƬ��ַ�а����ű�����"
				Exit For
			End If
		Next
		If checkXHTML<>"" Then Exit Function
		'���е��¼�����
		For Each Node in xml.documentElement.selectNodes("//@*")
			If Left(Node.nodeName,2)="on" Then
				checkXHTML="���ݺ�ǰ̨��ֹ�ύ������"
				Exit For
			End If
		Next
	Else
		checkXHTML="�����޷�У�飬���ݲ��Ϸ�"
	End If
	Set xml=nothing
End Function
Private Function entity2Str(strText)
		Dim s,match,po,i,re
		s=replace(strText,"&amp;","&")
		If InStr(s,"\")=0 And InStr(s,"&#")=0 And InStr(s,"%")=0 and InStr(s,Chr(13))=0 and InStr(s,Chr(10))=0  and InStr(s,Chr(9))=0 and InStr(s,"	")=0  and InStr(s,"/*")=0 and InStr(s,"*/")=0 Then
			entity2Str=LCase(strText)
			Exit Function
		End If
		Set re=new RegExp
		re.IgnoreCase =true
		re.Global=true
		re.Pattern="(&#x)([0-9|a-z]{1,2})"
		Set match = re.Execute(s)
		For i= 0 to  match.count -1
			po=re.Replace(match.item(i),"$2")
			po="&H"+po
			If IsNumeric(po) Then
				s=Replace(s,match.item(i),Chr(po))
			End If
		Next
		re.Pattern="(&#0*)"
		s=re.Replace(s,"&#")
		re.Pattern="&#([0-9]{1,3})"
		Set match = re.Execute(s)
		For i= 0 to  match.count -1
			po=re.Replace(match.item(i),"$1")
			s=Replace(s,"&#"&po&";",Chr(po))
			s=Replace(s,"&#"&po&"",Chr(po))
		Next
		re.Pattern="(\\0*)"
		s=re.Replace(s,"\")
		re.Pattern="(\\)([0-9|a-z]{1,2})"
		Set match = re.Execute(s)
		For i= 0 to  match.count -1
			po=re.Replace(match.item(i),"$2")
			po="&H"+po
			If IsNumeric(po) Then
				s=Replace(s,match.item(i),Chr(po))
			End If
		Next
		Rem url����ת��
		re.Pattern="(%)([0-9|a-z]{1,2})"
		Set match = re.Execute(s)
		For i= 0 to  match.count -1
			po=re.Replace(match.item(i),"$2")
			po="&H"+po
			If IsNumeric(po) Then
				s=Replace(s,match.item(i),Chr(po))
			End If
		Next
		s=replace(s,Chr(13),"")
		s=replace(s,Chr(10),"")
		s=replace(s,Chr(9),"")
		s=replace(s,"	","")
		s=replace(s,"/*","")
		s=replace(s,"*/","")
		entity2Str=LCase(s)
		Set Re =nothing
	End Function
Function Dv_FilterJS(v)
	Dim userface
	If  Not Isnull(V) Then
		 userface=entity2Str(v)
		If InStr(userface,"script:")>0 or InStr(userface,"document.")>0 Or InStr(userface,"xss:") > 0 Or InStr(userface,"expression") > 0 Then
				userface="http://bbs.cndw.com/images/zhutou.jpg"
		End If
		Dv_FilterJS=userface
	End If 
End Function
%>