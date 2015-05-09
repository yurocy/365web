<%
Dim Startime,SqlNowString,ConnStr
Dim conn,Db,MyDbPath,pos
Startime = Timer()
MyDbPath = ""
MyDbPath=request.ServerVariables("url")
pos=instr(MyDbPath,"/")+1
MyDbPath=left(MyDbPath,instr(pos,MyDbPath,"/"))


Const IsSqlDataBase = 0


If IsSqlDataBase = 1 Then
	Const SqlDatabaseName = "dzwang"
	Const SqlPassword = ""
	Const SqlUsername = "sa"
	Const SqlLocalName = "(local)"
	SqlNowString = "GetDate()"
Else
	Db = "../../database/#Website_8168123_data.mdb"
	SqlNowString = "Now()"
End If

	If IsSqlDataBase = 1 Then
		ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
	Else
		ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(MyDbPath & Db)
	End If
       ' On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open ConnStr
	If Err Then
		err.Clear
		Set Conn = Nothing
		Response.Write "数据库连接出错,请检查连接字串。"
		Response.End
	End If

sub CloseConn()
	conn.close
	set conn=nothing
end sub
%>

