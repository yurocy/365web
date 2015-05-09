<%
On Error Resume Next 
Server.ScriptTimeOut=9999999 
'Function getHTTPPage(Path)
' t = GetBody(Path)
'End function

function getHTTPPage(url)
dim Http
set Http=server.createobject("MSXML2.XMLHTTP")
Http.open "GET",url,false
Http.send()
if Http.readystate<>4 then 
exit function
end if
r_data=bytesToBSTR(Http.responseBody,"GB2312")
if instr(r_data,"该页面暂无数据")>0 then NO_data
getHTTPPage=r_data
set http=nothing
if err.number<>0 then err.Clear 
end function

Function bytestobstr(body,CSET)
DIM OBJSTREAM
SET OBJSTREAM = Server.CreateObject("adodb.stream")
objstream.type = 1
objstream.Mode = 3
objstream.Open 
objstream.Write body
objstream.Position = 0
objstream.type = 2
objstream.Charset=Cset
bytesTobstr = objstream.readText
objstream.close
set objstream=nothing
end function





Public Function CheckDir(ByVal ckDirname)

Dim M_fso

CheckDir=False

Set M_fso = CreateObject("Scripting.FileSystemObject")

If (M_fso.FolderExists(server.mappath(ckDirname))) Then

CheckDir=True

End If

Set M_fso = Nothing

End Function

'作 用：创建目录


Public Function CreateDir(ByVal crDirname)

Dim M_fso

CreateDir=False

Set M_fso = CreateObject("Scripting.FileSystemObject")

If (M_fso.FolderExists(server.mappath(crDirname))) Then

CreateDir=False

Else

M_fso.CreateFolder(server.mappath(crDirname))

CreateDir=True

End If

Set M_fso = Nothing

End Function








%>
<%sub NO_data()%>

<%response.end
end sub%>