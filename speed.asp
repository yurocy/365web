<!--#include file="Conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>时间测试</title>
<style type="text/css">
BODY {
	font-size: 9pt;
	color: #000000;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	font-family: Geneva, Arial, Helvetica, sans-serif;
	TEXT-DECORATION: none;
   
}
.input_04 {
	width: 150px;
	font-size: 12px;
	background-color: #FFFFE1;
	color: #FF0000;
	height: 22px;
	line-height: 15px;
	BORDER-RIGHT: #FFFFE1;
	BORDER-LEFT: #FFFFE1;
	BORDER-TOP: #FFFFE1;
	BORDER-BOTTOM: #FFFFE1;
}
</style>
</head>

<body style="background-color: #FFFFE1;">
	<table width="100%" align="center" style="background-color: #FFFFE1;font-size:9pt;">
		<form name="myform" action="" method="post">
		
		<tr><td>
		<%
'favorites.asp?idstr=
search=request.QueryString("id")
if len(search)=0 then search=0
sql="select * from WebGoodSite_Url_lite where Class_ID="&search&" order by id asc"
set RsLite=server.createobject("adodb.recordset")
RsLite.open sql,conn,1,1
total=RsLite.recordcount
i=1
do while not RsLite.eof%>
	测速网址： <a href="<%=Rslite("Site_Url")%>" target="_blank" style="color:#000000;"><%=Rslite("Site_Title")%>  <%=replace(Rslite("Site_Url"),"/out.asp?turl=","")%></a> <br>  <input style="border:hidden" readonly type=text name="status<%=i%>" size=20 class="input_04"><br>
<%
Rslite.movenext
i=i+1
loop
Rslite.close
set Rslite=nothing
%>
	
		</td></tr>
		<tr><td><font color="blue">温馨提示:反应时间越小,打开就越快</font></td></tr>
		</form>
	</table>

<SCRIPT LANGUAGE=JavaScript>

var timecount=1;
var timerstart0;
var bTimer = false;
var timer;

function autotime(h){
	if(timecount>150){
		for(b=0;b<=h;b++){
                     
		if(document.myform["status"+b].value=='测速中...'){eval('document.myform.status'+b+'.value="站点的连接超时"')};
		}
	}else{
		timecount++;
	}
}


function CountTime(i,timespace){
if (timespace>100)
{eval('document.myform.status'+i+'.value="站点的连接超时"');}
else
{if(timespace<1)  
  {eval('document.myform.status'+i+'.value="反应极快"');}  
  else
  {var timestr="反应时间："+timespace/100*1000+"ms"
 eval('document.myform.status'+i+'.value=timestr');}
 }
}

function testspeed(url){
timerstart0=timecount;


		<%
'favorites.asp?idstr=
search=request.QueryString("id")
if len(search)=0 then search=0
sql="select * from WebGoodSite_Url_lite where Class_ID="&search&" order by id asc"
set RsLite=server.createobject("adodb.recordset")
RsLite.open sql,conn,1,1
total=RsLite.recordcount
i=1
do while not RsLite.eof%>
eval('document.myform.status<%=i%>.value="测速中..."');
document.write("<img src='<%=replace(Rslite("Site_Url"),"/out.asp?turl=","")%>' width=1 height=1 onerror='CountTime(<%=i%>,timecount-timerstart0);'>");

<%
Rslite.movenext
i=i+1
loop
Rslite.close
set Rslite=nothing
%>

 
}
timer=setInterval("autotime(<%=total%>)",100);
testspeed();
</script>



</body>
</html>







