<%@language=vbscript codepage=936 %>
<%
Dim Startime
Dim SqlNowString
Dim conn,Db,MyDbPath
Startime = Timer()
MyDbPath = ""

Const IsSqlDataBase = 0

If IsSqlDataBase = 1 Then
	Const SqlDatabaseName = "dzwang"
	Const SqlPassword = ""
	Const SqlUsername = "sa"
	Const SqlLocalName = "(local)"
	SqlNowString = "GetDate()"
Else
	Db = "database/#Website_8168123_data.mdb"
	SqlNowString = "Now()"
End If


	Dim ConnStr
	If IsSqlDataBase = 1 Then
		ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
	Else
		ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(MyDbPath & db)
	End If

	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open ConnStr
	

sub CloseConn()
	conn.close
	set conn=nothing
end sub
%>

