<!--#include file="Conn.asp" --><!--#include file="inc/config.Asp"--><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<head>
<title><%=webname%></title>
<link rel="shortcut icon" href="favicon.ico" />
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="keywords" content="<%=keywords%>" />
<meta name="description" content="<%=description%>" />
<META NAME="COPYRIGHT" CONTENT="<%=COPYRIGHT%>" />
<link href="/collection-pop.css" type="text/css" rel="stylesheet" />
<style type="text/css">
<!--
.pf {
	position:absolute;
}
a:link {
	color:#000000;
	text-decoration:none;
	font-size: 13px;
}
a:visited {
	color:#000000;
	text-decoration:none;
	font-size: 13px;
}
a:hover {
	color:#ff0000;
	text-decoration:underline;
	font-size: 13px;
}
a:active {
	color:#0000FF;
	text-decoration:none;
	font-size: 13px;
}
.ad {
	padding-left:31px;
	text-align:left;
}
#info {
	width:693px;
}
#info ul {
	margin:0;
}
#info li {
	width:77px;
	height:28px;
	float:left;
}
#info li a {
	line-height:28px;
	display:block;
	background:url("/images/bx039.gif") no-repeat;
}
#info li a:link, #info li a:visited {
	font-size:12px;
	text-decoration:none;
	color:#ffffff;
	background-position:center top
}
#info li a:hover, #info li a:active {
	font-size:12px;
	text-decoration:none;
	color:#000000;
	font-weight :bolder;
	background-position:center bottom
}
.w01 {
	font-family: "宋体";
	font-size: 12px;
	color: #333333;
	text-decoration: none;
}
.w02 {
	font-family: "宋体";
	font-size: 14px;
	font-weight: bold;
	color: #FFFFFF;
	text-decoration: none;
}
.w03 {
	font-family: "宋体";
	font-size: 14px;
	color: #333333;
	text-decoration: none;
	line-height: 24px;
	height:59px;
	writing-mode:tb-rl;
}
.w04 {
	font-family: "宋体";
	font-size: 14px;
	font-weight: bold;
	color: #CC3333;
	text-decoration: none;
}
.w05 {
	font-family: "宋体";
	font-size: 12px;
	color: #FF6600;
}
.w06 {
	font-family: "宋体";
	font-size: 12px;
	font-weight: bold;
	color: #FFFFFF;
	text-decoration: none;
}
.w07 {
	font-family: "宋体";
	font-size: 12px;
	color: #FFCC00;
	text-decoration: none;
}
.w08 {
	font-family: "宋体";
	font-size: 12px;
	color: #FFFFFF;
}
.w09 {
	font-family: "宋体";
	font-size: 13px;
	color: #FF0000;
	text-decoration: none;
}
.w10 {
	font-family: "宋体";
	font-size: 12px;
	color: #333333;
	text-decoration: none;
}
.dd01 {
	font-family: "宋体";
	font-size: 12px;
	color: #FFFFFF;
	text-decoration: none;
	background-color: #000000;
	height: 16px;
	border: 1px solid #999999;
}
/*页面下拉框*/  
.menu1 {
	position:relative;
	cursor:hand;
	z-index:2;
}
.menu2 ul {
	display:none;
}
.menu3 {
	position:relative;
	cursor:hand;
	z-index:2;
}
.menu4 ul {
	display:none;
}
.menu1 ul {
	list-style:none;
	display:block;
	position:absolute;
	left:32px;
	top:-5px;
	width:80px;
	background:#42C4DB;
	padding-top:5px;
	padding-bottom:5px;
}
.menu1 ul li a {
	color:#ffffff;
	display:block;
	height:22px;
	line-height:22px;
	font-size:9pt;
	margin-top:1px;
	margin-left:5px;
	margin-right:5px;
	background:#6AD2E5 url(/images/m_arrow.gif) no-repeat 5px center;
}
.menu1 ul li a:hover {
	color:#000000;
	text-decoration:none;
	background:#ffffff url(/images/m_arrow1.gif) no-repeat 5px center;
}
.menu3 ul {
	list-style:none;
	display:block;
	position:absolute;
	left:32px;
	top:-5px;
	width:80px;
	background:#42C4DB;
	padding-top:5px;
	padding-bottom:5px;
}
.menu3 ul a {
	color:#ffffff;
	display:block;
	height:22px;
	line-height:22px;
	font-size:9pt;
	margin-top:1px;
	margin-left:5px;
	margin-right:5px;
	background:#6AD2E5 url(/images/m_arrow.gif) no-repeat 5px center;
	padding-left:8px;
}
.menu3 ul a:hover {
	color:#000000;
	text-decoration:none;
	background:#ffffff url(/images/m_arrow1.gif) no-repeat 5px center;
}
/*单元格*/ 
#tablebg {
	border-collapse: collapse;
	font-size:12px;
}
.t2 {
	background:#8EC2EE;
}
.paixu {
	background-color:#ffd;
	margin:0;
}
.paixu1 {
	height:22px;
	line-height:22px;
	border:1px solid #ecdc8f;
	background-color:#ffd;
	text-align:center;
}
.paixu2 {
	float:left;
	width:92px;
	height:22px;
	border-right:1px solid #ecdc8f;
	background-color:#ffd;
}
.paixu3 {
	float:left;
	font-family:Verdana;
}
.paixu3 ul {
	list-style:none;
	margin:0;
}
.paixu3 ul li {
	float:left;
	border-collapse:collapse;
	border:1px solid #ecdc8f;
	margin-left:-1px;
	margin-top:-1px;
	margin-bottom:-1px;
	padding-left:5px;
	padding-right:5px;
}
-->
</style>
<link href="css/main.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="/js/funb.js"></script>
<script language="javascript">	
	function coloseChl(id){id.style.display = 'none';}
	function hideallmenu(){
		var counts = 0;
		for(var i=0;i<document.all.tags("div").length;i++){
			if(document.all.tags("div")[i].className.substring(0,10) == 'menu menu_'){
				counts ++ ;
				document.all.tags("div")[i].className = 'menu menu_nosel';
			}
			if(counts > 10){
				break;
			}
		}
	}
	function selectChl(id)
	{	
		
		if(id=='pID1'||id=='pID2'||id=='pID3'||id=='pID4'||id=='pID5'||id=='pID6'||id=='pID7'||id=='pID8'||id=='pID9')
		{
		document.getElementById("showall").className=""
		document.getElementById("spID1").className=""
		//document.getElementById("spID2").className=""
		document.getElementById("spID3").className=""
		document.getElementById("spID4").className=""
		document.getElementById("spID5").className=""
		document.getElementById("spID6").className=""
		document.getElementById("spID7").className=""
		document.getElementById("spID8").className=""
		document.getElementById("spID9").className=""
		document.getElementById("s"+id).className="sel"
		
		var counts = 0;
		for(var i=0;i<document.all.tags("table").length;i++){
			if(document.all.tags("table")[i].className=="classtable"){
				document.all.tags("table")[i].style.display = 'none';
				counts++;
			}
			if(document.all.tags("table")[i].id.substring(0,2) == 'ch')
				document.all.tags("table")[i].style.display = 'block';
		}
		
		if(counts <=1){
			location.href = './.?' + id;
			return ;
		}
		
		//hideallmenu();
		try{
			document.all(id).style.display = 'block';
			document.all("titleID" + id.substring(3,id.length)).style.display = '';
		} catch(e){}
		try{
			document.all("menusel" + id.substring(3,id.length)).className = 'sel'
		} catch(e){}
		}
		else
		{
		showall();
		}
		
	}
	function showall()
	{
		document.getElementById("spID1").className=""
		//document.getElementById("spID2").className=""
		document.getElementById("spID3").className=""
		document.getElementById("spID4").className=""
		document.getElementById("spID5").className=""
		document.getElementById("spID6").className=""
		document.getElementById("spID7").className=""
		document.getElementById("spID8").className=""
		document.getElementById("spID9").className=""
		document.getElementById("showall").className="sel"
		var counts = 0;
		for(var i=0;i<document.all.tags("table").length;i++){
			if(document.all.tags("table")[i].className=="classtable"){
				document.all.tags("table")[i].style.display = 'block';
				counts ++ ;
			}
			if(document.all.tags("table")[i].id.substring(0,2) == 'ch')
				document.all.tags("table")[i].style.display = 'block';
			if(document.all.tags("table")[i].id.substring(0,7) == 'titleID')
				document.all.tags("table")[i].style.display = 'block';
		}
		hideallmenu();
		if(counts <= 1){
			location.href = './.';
		}
	}
</script>
<script language="javascript" src="/js/main.js"></script>
<script>
var HomepageFavorite = {
//设为首页
Homepage: function () {
if (document.all) {
document.body.style.behavior = 'url(#default#homepage)';
document.body.setHomePage(window.location.href);

}
else if (window.sidebar) {
if (window.netscape) {
try {
netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");
}
catch (e) {
alert("该操作被浏览器拒绝，如果想启用该功能，请在地址栏内输入 about:config,然后将项 signed.applets.codebase_principal_support 值该为true");
history.go(-1);
}
}
var prefs = Components.classes['@mozilla.org/preferences-service;1'].getService(Components.interfaces.nsIPrefBranch);
prefs.setCharPref('browser.startup.homepage', window.location.href);
}
}
,

//加入收藏
Favorite: function Favorite(sURL, sTitle) {
try {
window.external.addFavorite(sURL, sTitle);
}
catch (e) {
try {
window.sidebar.addPanel(sTitle, sURL, "");
}
catch (e) {
alert("加入收藏失败,请手动添加.");
}
}
}
}
</script>
</head>
<body>
<div class="main">


    
    <div id="alldata">
      <%	

	sql="select id,Type_Name from WebGoodSiteType_Main order by type_order desc"
	set tmp_rs=server.createobject("adodb.recordset")
	tmp_rs.open sql,conn,1,1
	tmp_i=1
	do while not tmp_rs.eof
	
	show_site(tmp_rs("i"))
	
	response.write "<script>document.getElementById('tad"&tmp_rs("id")&"').innerHTML="""&show_ad(tmp_i)&""";</script>"
	response.write "<script>document.getElementById('tbad"&tmp_rs("id")&"').innerHTML="""&htmltojs(showinfo(tmp_rs("Type_Name")&"下广告"))&""";</script>"
	response.write "<!--"&tmp_rs("Type_Name")&"-->"
	
	tmp_rs.movenext
	
	tmp_i=tmp_i+1
	
	loop

	response.write "<script>document.getElementById('tbad0').innerHTML="""&htmltojs(showinfo("收藏下广告"))&""";</script>"
	
	%>

  <script>
Fun_GetPostData();
</script>
    </div>
  <div style="font-size:12px; background:#FFFFFF"> <%=showinfo("友情链接")%> </div>


  <%function show_site(id)%>
  <%

	sql="select * from WebGoodSiteType_Main where ID="&ID
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if not (rs.eof and rs.bof) then
	Type_Name=rs("Type_Name")
	Type_Color=rs("Type_Color")
	Bg_Color=rs("Bg_Color")
	Bg_img=rs("Bg_img")
	end if
%>
  <table width="845" height="27" border="0" align="center" class="classtable" cellpadding="0" cellspacing="0"  id="pID<%=id%>" >
    <tr>
      <td>
        <div id="tableid<%=id%>">
          <%
	sql="select * from WebGoodSiteType_Main where ID="&ID
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if not (rs.eof and rs.bof) then
	Type_Name=rs("Type_Name")
	Type_Color=rs("Type_Color")
	Bg_Color=rs("Bg_Color")
	Bg_img=rs("Bg_img")
	end if

    sql="select * from WebGoodSiteType_Class where Main_ID="&ID
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
        if rs.eof and rs.bof then
	else
        i=1
        do while not rs.eof
		
			dim RsMore
			sql="select * from WebGoodSite_Url where Class_ID="&rs("id")&" order by Is_Recommend desc,id asc"
			set RsMore=server.createobject("adodb.recordset")
			RsMore.open sql,conn,1,1
			total=RsMore.recordcount
			
			if total>0 then
		%>
          <table width="845" border="0" align="center" cellpadding="0" cellspacing="0" id="ch<%=rs("id")%>">
            <tr>
              <td width="30" valign="top">
              	<table width="30" height="59"  border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td background="images/leftbar.gif" style="cursor:hand;" onClick="coloseChl(ch<%=rs("id")%>)"><div align="center" class="w03"><%=rs("Type_Class")%></div></td>
                  </tr>
                </table></td>
              <td width="815">
              <table width="100%"  border="0" cellpadding="3" cellspacing="1" class="table" bgcolor="<%=Bg_Color%>">
                  <%
		 	 tmp_num=1
            'do while not RsMore.eof
			if total mod 7<>0 then
			total=(total+7-(total mod 7))
      end if
      
			for tmp_num=1 to total
			  
			  'response.write "<!--"&RsMore("Site_Title")&total&"-->"

	       %>
                  <%if tmp_num mod 7 = 1 or tmp_num=1 then%>
                  <tr bgcolor="<%=Type_Color%>" class="tr">
                    <%end if%>
                    <%if not RsMore.eof then%>
                    <td height="26" align="center" width="13%" onMouseOver="this.className='t2'; title='<%=RsMore("Site_Title")%>'" onMouseOut="this.className='tablebg';" class="tablebg"><div class="menu2" onMouseOver="this.className='menu1'" onMouseOut="this.className='menu2'"><a <%if RsMore("Site_Url")="测速" then %>onclick="openFlyBarS(<%=rsmore("id")%>)" <%else%> href="<%=RsMore("Site_Url")%>" target=_blank rel=nofollow <%end if%> title="<%=RsMore("Site_Title")%>" style="font-size:10pt;"><font style='color:<%=RsMore("Font_Color")%>;font-weight:<% if RsMore("bold")=1 then response.write "bold" else response.write "normal"  end if%>'><%=RsMore("Site_Title")%></font></a>
                       
                       
                      </div></td>
                    <%
RsMore.movenext
else
response.write "<!--"&total&"-->"
%>
                    <td  onMouseOver="this.className='t2'" onMouseOut="this.className='tablebg';" class="tablebg"  width="13%">　</td>
                    <%end if%>
                    <%if tmp_num mod 7 = 0 then%>
                  </tr>
                  <%end if%>
                  <%
		next
'		RsMore.movenext
'              loop
'              RsMore.close
'	      set RsMore=nothing
		  
		  %>
                </table></td>
            </tr>
          </table>
          <%end if%>
 
          <%
		  	  rs.movenext
        loop
        end if
        rs.close
	set rs=nothing
		  %>
        </div></td>
    </tr>
  </table>
</div>
<%end function%>

<%function show_ad(id)

        sql="select * from WebSiteAdvertisementWords_URL"
	set adrs=server.createobject("adodb.recordset")
	adrs.open sql,conn,1,1
	if id<adrs.recordcount then
	'response.write adrs.recordcount
       adrs.move id
          show_ad="<a href='"&adrs("Site_Url")&"' target='_blank' rel='nofollow'  style='color: "&adrs("Font_Color")&"'>"&adrs("Site_Title")&"</a>&nbsp;&nbsp;&nbsp;"
      else
	  show_ad="&nbsp;"
        end if
        adrs.close
	set adrs=nothing
     end function   
%>
<script type="text/javascript" src="/collection-pop-min.js"></script>

</body>
</html>