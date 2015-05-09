<!--#include file="Conn.asp" --><!--#include file="inc/config.Asp"--><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
<title><%=webname%></title>
<link rel="shortcut icon" href="favicon.ico" />
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />

<meta name="keywords" content="<%=keywords%>" />
<meta name="description" content="<%=description%>" />
<META NAME="COPYRIGHT" CONTENT="<%=COPYRIGHT%>" />

<style type="text/css">
<!--
.pf{position:absolute;}
a:link { color:#000000;text-decoration:none;font-size: 13px;}
a:visited { color:#000000;text-decoration:none;font-size: 13px;}
a:hover { color:#ff0000;text-decoration:underline;font-size: 13px;}
a:active { color:#0000FF;text-decoration:none;font-size: 13px;}
.ad{
	padding-left:31px;
	text-align:left;
}
#info{width:693px;}
#info ul{ margin:0;}
#info li{ width:77px; height:28px; float:left;}
#info li a{line-height:28px; display:block;background:url("/images/bx039.gif") no-repeat;}
#info li a:link,#info li a:visited{ font-size:12px; text-decoration:none; color:#ffffff; background-position:center top}
#info li a:hover,#info li a:active{ font-size:12px; text-decoration:none; color:#000000; font-weight :bolder; background-position:center bottom}

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
.menu1{ position:relative;cursor:hand;z-index:2;  }
.menu2 ul { display:none; }
.menu3{ position:relative;cursor:hand;z-index:2;  }
.menu4 ul { display:none; }
.menu1 ul{ list-style:none;display:block;position:absolute;left:32px;top:-5px;width:80px;background:#42C4DB;padding-top:5px;padding-bottom:5px;}
.menu1 ul li a{ color:#ffffff;display:block;height:22px;line-height:22px;font-size:9pt;margin-top:1px;margin-left:5px;margin-right:5px;background:#6AD2E5 url(/images/m_arrow.gif) no-repeat 5px center;}
.menu1 ul li a:hover{ color:#000000;text-decoration:none;background:#ffffff url(/images/m_arrow1.gif) no-repeat 5px center;}
.menu3 ul{ list-style:none;display:block;position:absolute;left:32px;top:-5px;width:80px;background:#42C4DB;padding-top:5px;padding-bottom:5px;}
.menu3 ul a{ color:#ffffff;display:block;height:22px;line-height:22px;font-size:9pt;margin-top:1px;margin-left:5px;margin-right:5px;background:#6AD2E5 url(/images/m_arrow.gif) no-repeat 5px center;padding-left:8px;}
.menu3 ul a:hover{ color:#000000;text-decoration:none;background:#ffffff url(/images/m_arrow1.gif) no-repeat 5px center;}

/*单元格*/ 
#tablebg{ border-collapse: collapse;font-size:12px;}
.t2 { background:#8EC2EE;}

.paixu{background-color:#ffd; margin:0;}
.paixu1{height:22px;line-height:22px;border:1px solid #ecdc8f;background-color:#ffd;text-align:center;}
.paixu2{float:left;width:92px;height:22px;border-right:1px solid #ecdc8f;background-color:#ffd;}
.paixu3{float:left;font-family:Verdana;}
.paixu3 ul{ list-style:none; margin:0;}
.paixu3 ul li{float:left;border-collapse:collapse;border:1px solid #ecdc8f;margin-left:-1px;margin-top:-1px;margin-bottom:-1px;padding-left:5px;padding-right:5px;}
-->
</style>
<link href="css/main.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
body {
	margin-bottom: 10px;
}
-->
</style></head>
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


<body  align="center">
<div class="main">

<table width="845" height="60" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td valign="top" background="/images/bx002.gif"><table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0">
	<form action="url_submit.asp" method="post" style="margin:0px" target="_blank">
      <tr>
        <td width="90" bgcolor="CC0000" class="w06"><div align="center">我要提交网址</div></td>
      <td width="57" class="w08"><div align="right">类型：</div></td>
        <td width="109"><select name="fl" class="dd01">
<script language="JavaScript" type="text/javascript" src="fl.asp"></script>

        </select></td>
        <td width="45" class="w08"><div align="right">网名：</div></td>
        <td width="75"><input name="Text1" type="text" class="dd01" id="Text1" size="10" /></td>
        <td width="45" class="w08"><div align="right">网址：</div></td>
        <td width="150"><input name="Text2" type="text" class="dd01" id="Text2" size="25"></td>
        <td width="45" class="w08"><div align="right">代理：</div></td>
        <td width="150"><input name="Text3" type="text" class="dd01" id="Text3" size="25"></td>
        <td width="52"><div align="center">
          <input name="Submit" type="submit" class="dd01" value="提交"  onClick="return chkform(this.form)" >
        </div></td>
      </tr>
	  </form>	
    </table>
	
      <table width="98%" height="28" style="line-height:28px;"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr><td><div align="center"><span class="w05">公告栏：</span><span class="w01" style="color:#FFFFFF; "><marquee direction="left" scrollamount="2" onmousemove="this.stop()" onmouseout="this.start()" scrolldelay="1" width="480"><%=showinfo("公告栏")%></marquee></span></div></td>
          <td valign="bottom" class="w05"><div align="right" class="w07"><a href="#" onclick="window.open('changeUrl.html','','width=516,height=390','','scrollbars=yes,resizable=yes')" style="color: #FFCC00;">我要更换网址</a> | <a href="#" onclick="window.open('fk.html','','width=516,height=390','','scrollbars=yes,resizable=yes')" style="color: #FFCC00;"></a><a href="ad.asp" target="_blank" style="color: #FFCC00;">我要投放广告</a></div></td>
        </tr>
      </table></td>
  </tr>
</table>
<table width="845" height="94" border="0" align="center" cellpadding="1" cellspacing="0">
  <tr>
    <td width="172" bgcolor="#FDDF13"><div align="center"><a href="/"><img src="images/logo.gif" width="157" height="83" /></a></div></td>
  <td width="574" background="/images/bx004.gif"><div id="ad_top" align="right">
     <%=showinfo("横幅广告")%>
    </div></td>
    <td width="95" background="/images/bx004.gif"><table width="73" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="22" bgcolor="FFCC33" class="w01"><div align="center">
<a href="javascript:HomepageFavorite.Favorite(window.location.href, document.title)" _fcksavedurl="javascript:HomepageFavorite.Favorite(window.location.href, document.title)" >加入收藏</a> </div></td>
      </tr>
      <tr>
        <td height="10"> </td>
      </tr>
      <tr>
        <td height="22" bgcolor="FFCC33" class="w01"><div align="center"><a href="javascript:HomepageFavorite.Homepage()" _fcksavedurl="javascript:HomepageFavorite.Homepage()" >设为首页</a>
</div></td>
      </tr>
    </table></td>
  </tr>
  <tr><td bgcolor="#636463" colspan="3" style="height: 5px;"></td></tr>
</table>
<div id="indexhfad" style="background:#FFFFFF; overflow:hidden; margin-top:5px; margin-bottom:5px"><%=showinfo("菜单上面广告")%></div>
<table width="845" border="0" align="center" cellpadding="0" cellspacing="0">
<form target="_blank" method="post" action="search.asp" name="serarch1">
  <tbody><tr>
  <td width="693"><div id="info">
<ul>
<li><a onclick="showall()" id="showall" href="#">全部显示</a></li>
<li><a onclick="selectChl('pID1')" id="spID1" href="#">热门网站</a></li>
<li><a href="#" onclick="selectChl('pID3')" id="spID3">足球登陆</a></li>
<li><a href="#" onclick="selectChl('pID4')" id="spID4">彩票网站</a></li>
<li><a href="#" onclick="selectChl('pID5')" id="spID5">赛马网站</a></li>
<li><a href="#" onclick="selectChl('pID6')" id="spID6">娱乐网站</a></li>
<li><a href="#" onclick="selectChl('pID7')" id="spID7">银行网站</a></li>
<li><a href="#" onclick="selectChl('pID8')" id="spID8">金融网站</a></li>
<li><a href="#" onclick="selectChl('pID9')" id="spID9">实用网站</a></li>
</ul>
</div></td>
    <td align="right" bgcolor="#636463"><table width="95%" cellspacing="0" cellpadding="0" border="0">
      <tr background="/images/bx036.gif">
      <td width="104" align="left" class="w01"><input style=" border:0px; background:url(/images/bx036.gif) no-repeat left top;" name="keyword" type="text" size="12" value="站内搜索" onclick="this.focus(3);this.value=''"></td>
        <td width="40" align="left"><input type="image" src="/images/bx037.gif"  width="33" height="20"></td>
      </tr>
	</table></td>
  </tr>
 <tr><td style="height:5px"></td> </tr>
</tbody></form></table>
  <!--头部-->
  <div class="ad_6" style="background:#E1E1E1">
   <%=showinfo("公告栏下部广告")%>
  </div>

<table width="845" height="27" border="0" align="center"  cellpadding="0" cellspacing="0"  ><tr><td><table style=" background-repeat:no-repeat;background:url(images/title1.gif)" width="845" height="27" border="0" align="center" cellpadding="0" cellspacing="0" >
<tr>
<td width="300""><div align="center" class="w02" style="margin-top:10px">我的收藏</div></td>
<td width="500" class="w09" align="right">提示:登陆帐户可永久收藏网址</td>
<td width="1420">
<div align="right" class="w01">
<iframe src="login.asp" width="100%" height="25" frameborder="0" marginwidth="0" marginheight="0" scrolling="no" id="ifr"></iframe>

</div>
 
  </td>

<td width="25"></td>
</tr>
</table>


<table width="845" height="4" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<td></td>
</tr>
</table>
<table width="845" height="8" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<td><div id="tbad0"></div></td>
</tr>
</table>
<table width="845" height="4" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<td></td>
</tr>
</table>

</td></tr></table>



<table width="845" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<td width="30" valign="top"><table width="30" height="59"  border="0" cellpadding="0" cellspacing="0">
<tr>
<td background="/images/bx031.gif" style="cursor:hand;" onClick="coloseChl(chl5)"><div align="center" class="w03">收 藏</div></td>
</tr>
</table></td>
<td width="815">

<div id="favorites_body"></div>


</td>
</tr>
</table>
<script language="javascript" type="text/javascript">
// bigin fly bar
var bIsCatchFlyBar = false;
var dragClickX = 0;
var dragClickY = 0;
function HideFlyBar(){
document.getElementById("divFlyBar").style.visibility = "hidden";
//myFlyBarRestorButton.style.display = '';
}

function openFlyBarS(id){
try
{ 
 document.frames["runid"].StopAll(); 
}
catch (e)
{}  
document.getElementById("runid").src="";
document.getElementById("runid").src='speed.asp?id='+id+'&date='+new Date().getTime();
document.getElementById("divFlyBar").style.display = "";
myload_flybar();
document.getElementById("divFlyBar").style.visibility = "visible";
}




function myload_flybar(){
document.getElementById("divFlyBar").style.top=document.documentElement.scrollTop;
//document.getElementById("divFlyBar").style.left=document.body.offsetWidth-document.getElementById("divFlyBar").clientWidth-670+document.body.scrollLeft;
//document.getElementById("divFlyBar").style.left=document.body.offsetWidth-document.body.scrollLeft;
}
window.onresize = myload_flybar;
//window.onscroll = myload_flybar;


function catchFlyBar(e){
bIsCatchFlyBar = true;
var x=event.x+document.body.scrollLeft;
var y=event.y+document.documentElement.scrollTop;
dragClickX=x-document.getElementById("divFlyBar").style.pixelLeft;
dragClickY=y-document.getElementById("divFlyBar").style.pixelTop;
document.getElementById("divFlyBar").setCapture();
document.onmousemove = moveFlyBar;
}
function releaseFlyBar(e){
bIsCatchFlyBar = false;
document.getElementById("divFlyBar").releaseCapture();
document.onmousemove = null;
}
function moveFlyBar(e){
if(bIsCatchFlyBar){
document.getElementById("divFlyBar").style.left = event.x+document.body.scrollLeft-dragClickX;
document.getElementById("divFlyBar").style.top = event.y+document.documentElement.scrollTop-dragClickY;
}}
</script>
<div id="divFlyBar" style="position:absolute;top:0;left:1;visibility:hidden;cursor:move;z-index:1000;display:none;"  align="center">
  <table style="filter:alpha(opacity=95);background-color:#9A9A9A;margin-top:253px;margin-right:588px;" border="0" cellspacing="1" cellpadding="0">
    <tr>
      <td><table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
          <tr valign="middle" style=" background:url(images/bx001.gif) repeat-x left bottom;">
            <td style="font-size:12px;color:#000000;font-weight:bold" width="90%" align="left">&nbsp; &nbsp;<span id="ftitle">站点测速</span></td>
            <td width="60" align="center"><a href='javascript:HideFlyBar();'><img src="images/close.gif" alt="点击关闭窗口"  border="0" width="48" height="15"/></a></td>
          </tr>
        </table></td>
    </tr>
    <tr id="flyTailerTr" name="flyTailerTr">
      <td id="flyTailerHolder" name="flyTailerHolder" style="background-color:#ff0000;color:black;font-weight:bold;font-size:12px;font-family:Courier New;" align="center"><iframe id="runid" scrolling="no" frameborder="0" marginheight="0" marginwidth="0" width="400" height="600" src="" ></iframe></td>
    </tr>
  </table>
</div>
<div id="alldata_s"></div>
<div id="alldata">

<%	

	sql="select id,Type_Name from WebGoodSiteType_Main order by type_order desc"
	set tmp_rs=server.createobject("adodb.recordset")
	tmp_rs.open sql,conn,1,1
	tmp_i=1
	do while not tmp_rs.eof
	
	show_site(tmp_rs("id"))
	
	response.write "<script>document.getElementById('tad"&tmp_rs("id")&"').innerHTML="""&show_ad(tmp_i)&""";</script>"
	response.write "<script>document.getElementById('tbad"&tmp_rs("id")&"').innerHTML="""&htmltojs(showinfo(tmp_rs("Type_Name")&"下广告"))&""";</script>"
	response.write "<!--"&tmp_rs("Type_Name")&"-->"
	
	tmp_rs.movenext
	
	tmp_i=tmp_i+1
	
	loop

	response.write "<script>document.getElementById('tbad0').innerHTML="""&htmltojs(showinfo("收藏下广告"))&""";</script>"
	
	%>
  
  
  
  
</div>


<script>
Fun_GetPostData();
</script>
<div class="main" style="font-size:12px; background:#FFFFFF">
<%=showinfo("友情链接")%>
</div>

    <table width=845 align=center cellpadding="0" cellspacing="0" border=0>
      <tr>
        <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="1" bgcolor="#CCCCCC"></td>
          </tr>
          <tr>
            <td bgcolor="#ECECEC"><table width="700" border="0" align="center" cellpadding="10" cellspacing="0">
              <tr>
                <td width="150"><div align="right"><img src="/images/bx032.gif" width="106" height="53" /></div></td>
                <td width="600" class="w01" style="text-align:center; line-height:normal"> 广告联系QQ:2624632368<font color="#FF0000">
				</font>
				
				<font color="#0000FF"></font>客服时间:7×24小时<%=weburl%> 技术支持 <br /> 声明:网上有诈骗.所有在本站刊登广告的网站和内容，一概与本站无关，请各位网友密切注意!<span style="display:none">

</span><br />
                  <a href="http://www.miibeian.gov.cn/" target="_blank" style="color:#ECECEC; ">备案号:粤ICP备08107990号 
</a> &nbsp; </td>

                <td width="1"></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
</table>
	<div align="center" style="width:100%; padding-top:15px;">
&nbsp;&nbsp;<a href="" target="_blank"><font style="text-decoration:none; color:#000000"></font></a><br>
<a href="" target="_blank"><font style="text-decoration:none;"></font></a>
</div>

<%
  				Set rs = Server.CreateObject("ADODB.Recordset")
				sql="select * from [web_ad] where web_adname='Side' "
				rs.Open sql,conn,1,3
				ad_src=rs("web_info")
				open=rs("open")
				rs.Close
				set rs=nothing
				
				ad_fg=split(ad_src,"$@$@$@$")
				if open=1 then
				%>
  <!--左边5个-->
<div id="Javascript.LeftDiv" style="position: absolute;z-index:999;top:-1000px;word-break:break-all;display:none;">
<%=ad_fg(0)%>
  
</div>
  

  <!--右边5个-->
<div id="Javascript.RightDiv" style="position: absolute;z-index:999;top:-1000px;word-break:break-all;display:none;">
  <%=ad_fg(1)%>
</div>
<center>
</center>
<script language="javascript" type="text/javascript">
<!--
var showad = true;        //是否显示广告
var Toppx = 30;            //上端位置
var AdDivW = 100;        //宽度
var AdDivH = 360;        //高度
var PageWidth = 800;    //页面多少宽度象素下正好不出现左右滚动条
var MinScreenW = 1024;    //显示广告的最小屏幕宽度象素

function scall(){
    if(!showad){return;}
    if (window.screen.width<MinScreenW){
        alert("临时提示：\n\n显示器分辨率宽度小于"+MinScreenW+",不显示广告");
        showad = false;
        document.getElementById("Javascript.LeftDiv").style.display="none";
        document.getElementById("Javascript.RightDiv").style.display="none";
        return;
    }
    var Borderpx = ((window.screen.width-PageWidth)/2-AdDivW)/2;

    document.getElementById("Javascript.LeftDiv").style.display="";
    document.getElementById("Javascript.LeftDiv").style.top=(document.documentElement.scrollTop+Toppx)+"px";
    document.getElementById("Javascript.LeftDiv").style.left="1px";
    document.getElementById("Javascript.RightDiv").style.display="";
    document.getElementById("Javascript.RightDiv").style.top=(document.documentElement.scrollTop+Toppx)+"px";
    document.getElementById("Javascript.RightDiv").style.right="1px";
}

function hidead()
{
    showad = false;
    document.getElementById("Javascript.LeftDiv").style.display="none";
    document.getElementById("Javascript.RightDiv").style.display="none";
}
window.onscroll=scall;
window.onresize=scall;
window.onload=scall;
//-->
</script>
<%end if%>










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
      <td><table width="845" height="27" border="0" align="center" cellpadding="0" cellspacing="0"  id="titleID<%=id%>" >
          <form  action="search.asp?typeid=<%=id%>" method="post" target="_blank" name="myform<%=id%>">
            <tr>
              <td><div class="title_1" style="background:url(<%=Bg_img%>) no-repeat left top;;">
                  <div class='reach'>精确查找：
                    <input name="keyword" type="text" class="input_2" value="关键字" onclick="this.focus(3);this.value=''" onkeydown="if(event.keyCode==13){toFind();return;}">
                    <input type="image" src="/images/bx023.gif" width="18" height="18" align="absmiddle">
                  </div>
                  <h2><%=Type_Name%></h2>
                  <div class="ad_text" id='tad<%=id%>'></div>
              </div></td>
            </tr>
          </form>
        </table>
        <table width='845' height='8' border='0' align='center' cellpadding='0' cellspacing='0'>
          <tr>
            <td><div id="tbad<%=id%>"></div></td>
          </tr>
        </table>
        <table width='845' height='8' border='0' align='center' cellpadding='0' cellspacing='0'>
          <tr>
            <td><div id="goldreach"><img class='left' align='absmiddle' src='/images/goldreach.gif' /> <span class="left" style="cursor:hand;" onclick="selectchar_s('ALL',<%=id%>,'tableid<%=id%>')"><img align='absmiddle' src='/images/dis_all.gif' width='57' /></span> <a href="javascript:selectchar_s('A',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">A</a> <a href="javascript:selectchar_s('B',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">B</a> <a href="javascript:selectchar_s('C',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">C</a> <a href="javascript:selectchar_s('D',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">D</a> <a href="javascript:selectchar_s('E',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">E</a> <a href="javascript:selectchar_s('F',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">F</a> <a href="javascript:selectchar_s('G',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">G</a> <a href="javascript:selectchar_s('H',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">H</a> <a href="javascript:selectchar_s('I',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">I</a> <a href="javascript:selectchar_s('J',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">J</a> <a href="javascript:selectchar_s('K',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">K</a> <a href="javascript:selectchar_s('L',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">L</a> <a href="javascript:selectchar_s('M',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">M</a> <a href="javascript:selectchar_s('N',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">N</a> <a href="javascript:selectchar_s('O',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">O</a> <a href="javascript:selectchar_s('P',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">P</a> <a href="javascript:selectchar_s('Q',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">Q</a> <a href="javascript:selectchar_s('R',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">R</a> <a href="javascript:selectchar_s('S',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">S</a> <a href="javascript:selectchar_s('T',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">T</a> <a href="javascript:selectchar_s('U',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">U</a> <a href="javascript:selectchar_s('V',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">V</a> <a href="javascript:selectchar_s('W',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">W</a> <a href="javascript:selectchar_s('X',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">X</a> <a href="javascript:selectchar_s('Y',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">Y</a> <a href="javascript:selectchar_s('Z',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">Z</a> <a href="javascript:selectchar_s('0',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">0</a> <a href="javascript:selectchar_s('1',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">1</a> <a href="javascript:selectchar_s('2',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">2</a> <a href="javascript:selectchar_s('3',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">3</a> <a href="javascript:selectchar_s('4',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">4</a> <a href="javascript:selectchar_s('5',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">5</a> <a href="javascript:selectchar_s('6',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">6</a> <a href="javascript:selectchar_s('7',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">7</a> <a href="javascript:selectchar_s('8',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">8</a> <a href="javascript:selectchar_s('9',<%=id%>,'tableid<%=id%>')" style="cursor:hand;">9</a> </div></td>
          </tr>
        </table>
        <table width="845" height="8" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td></td>
          </tr>
        </table>
        <div id="tableid<%=id%>" align="center">
		
        
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
      </table>
	  </td>
    <td width="815" align="center">
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
		   <%if tmp_num mod 7 = 1 or tmp_num=1 then%><tr bgcolor="<%=Type_Color%>" class="tr"><%end if%>
		   <%if not RsMore.eof then%>
		     <td height="26" align="center" width="13%" onMouseOver="this.className='t2'; title='<%=RsMore("Site_Title")%>'" onMouseOut="this.className='tablebg';" class="tablebg"><div class="menu2" onMouseOver="this.className='menu1'" onMouseOut="this.className='menu2'"><a <%if RsMore("Site_Url")="测速" then %>onclick="openFlyBarS(<%=rsmore("id")%>)" <%else%> href="<%=RsMore("Site_Url")%>" target=_blank rel=nofollow <%end if%> title="<%=RsMore("Site_Title")%>" style="font-size:10pt;"><font style='color:<%=RsMore("Font_Color")%>;font-weight:<% if RsMore("bold")=1 then response.write "bold" else response.write "normal"  end if%>'><%=RsMore("Site_Title")%></font></a>
              <%
			  select case rsmore("img")
			  				case 1
			  response.write "<img src=""/images/1.gif"" />"
			  			  case 2 
			  response.write "<img src=""/images/2.gif"" />"
			  			  case 3
			  response.write "<img src=""/images/3.gif"" />"
			  			  case 4
			  response.write "<img src=""/images/4.gif"" />"
			  			  case 5
			  response.write "<img src=""/images/5.gif"" />"
			  			  case else
			  response.write ""
			  end select
			  %>

		   <ul style="left:23px; top:15px">
              <% sql="select * from WebGoodSite_Url_lite where Class_ID="&rsmore("id")&" order by id asc"
	      			set RsLite=server.createobject("adodb.recordset")
	    			 RsLite.open sql,conn,1,1
		 			
					    if not rslite.eof then%>


                  <%do while not RsLite.eof%>
				  <li><a href="<%=Rslite("Site_Url")%>"  title=<%=Rslite("Site_Url")%> rel="nofollow" target="_blank"><%=Rslite("Site_Title")%></a></li>
                  <%Rslite.movenext
              		loop
              Rslite.close
	     	set Rslite=nothing%>
              <%end if%>
			 <li><a href="#" onclick="addFavorites(<%=rsmore("id")%>,'<%=RsMore("Site_Title")%>')">加入收藏</a></li>
              </ul>
            </div></td>
			
			<%
RsMore.movenext
else
response.write "<!--"&total&"-->"
%>
    <td  onMouseOver="this.className='t2'" onMouseOut="this.className='tablebg';" class="tablebg"  width="13%">　</td>
<%end if%>

<%if tmp_num mod 7 = 0 then%></tr><%end if%>
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
<table width="845" height="20" style='padding-top:8px;' border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width='845' height='8' border='0' align='center' cellpadding='0' cellspacing='0'>
        <tr>
          <td width='50%'><div class='ad'>
              <div></div>
            </div></td>
          <td width='50%'><div class='ad'>
            <div></div></td>
        </tr>
      </table></td>
  </tr>
</table>
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
</div>
</body>
</html>