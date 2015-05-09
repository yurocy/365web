<!--#include file="../Conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<%If Session("flag")<>"" Then Response.redirect "login.asp"%>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312" />
<TITLE>后台管理</TITLE>
<link href="skins/css/style.css" rel="stylesheet" type="text/css" />
<style>
.main_left {table-layout:auto; background:#C2D5E3;}
.main_left_top{ background:url(skins/images/left_menu_bg.gif); padding-top:2px !important; padding-top:5px;}
.main_left_title{text-align:left; padding-left:15px; font-size:14px; font-weight:bold; color:#000;}
.left_iframe{HEIGHT: 100%; VISIBILITY: inherit;WIDTH: 180px; background:transparent;}
.main_iframe{HEIGHT: 100%; ;
	VISIBILITY: inherit;
	WIDTH:100%;
}
table { font-size:12px; font-family : tahoma, 宋体, fantasy; }
td { font-size:12px; font-family : tahoma, 宋体, fantasy;}
</style>
<style type="text/css">
SCROLLBAR-FACE-COLOR: #2A024D;
 SCROLLBAR-HIGHLIGHT-COLOR: #0042FF;
 SCROLLBAR-SHADOW-COLOR: #000000;
 SCROLLBAR-3DLIGHT-COLOR:#FDFDFD;
 SCROLLBAR-TRACK-COLOR: #000000;
 SCROLLBAR-ARROW-COLOR: #0042FF;
 SCROLLBAR-DARKSHADOW-COLOR: #000000;
		</style>
<script language = "javaScript" src = "admin.js" type="text/javascript"></script>
<SCRIPT>
var status = 1;
var Menus = new DvMenuCls;
document.onclick=Menus.Clear;
function switchSysBar(){
     if (1 == window.status){
		  window.status = 0;
          switchPoint.innerHTML = '<img src="skins/images/left.gif">';
          document.all("frmTitle").style.display="none"
     }
     else{
		  window.status = 1;
          switchPoint.innerHTML = '<img src="skins/images/right.gif">';
          document.all("frmTitle").style.display=""
     }
}
</SCRIPT>
<script language="JavaScript">
 javascript:window.history.forward(1);
</script>
<BODY style="MARGIN: 0px">
<TABLE border=0 cellPadding=0 cellSpacing=0 height="100%" width="100%" style="background:#C2D5E3;">
<TBODY>
<TR>
  <TD align=middle id=frmTitle vAlign=top name="fmTitle" class="main_left">
	<table width="100%" border="0" cellspacing="0" cellpadding="0" class="main_left_top">
	  <tr height="32">
	    <td valign="top"></td>
	    <td class="main_left_title" id="leftmenu_title">常用快捷功能</td>
	    <td valign="top" align="right"></td>
	  </tr>
	</table>
	<IFRAME frameBorder=0 id=frmleft name=frmleft src="left.htm" class="left_iframe" allowTransparency="true"></IFRAME>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr height="32">
	    <td valign="top"></td>
	    <td valign="bottom" align="center"></td>
	    <td valign="top" align="right"></td>
	  </tr>
	</table>
  </td>
  <TD bgColor=#C2D5E3 style="WIDTH: 10px">
	   <TABLE border=0 cellPadding=0 cellSpacing=0 height="100%">
	    <TBODY>
	    <TR>
	     <TD onclick=switchSysBar() style="HEIGHT: 100%">
	     <SPAN class=navPoint id=switchPoint title="关闭/打开左栏"><img src="skins/images/right.gif"></SPAN>
	     </TD>
	    </TR>
	    </TBODY>
	   </TABLE>
     </TD>
  <TD bgcolor="#C2D5E3" width="100%" vAlign=top>
	<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#C4D8ED">
	  <tr height="32">
	    <td valign="top" width="10" background="skins/images/bg2.gif"><img src="skins/images/teble_top_left.gif" alt="" /></td>
	    <td background="skins/images/bg2.gif" width="28"><img src="skins/images/arrow.gif" alt="" align="absmiddle" /></td>
	    <td background="skins/images/bg2.gif" width="80"><span style="float:left">官方公告：</span></td>
		<td background="skins/images/bg2.gif" style="text-align:right; color: #135294; "><marquee ONMOUSEOUT="start();" ONMOUSEOVER="stop();" style="width:expression(document.body.clientWidth-650)"></marquee>  <a href="index.asp" target='_top'>后台首页</a> | <a href="../index.html" target="_blank">网站首页</a> | <a href="Admin_EditPwd.asp" target='frmright'>修改密码</a> | <a href="logout.asp" target="_top">退出</a> </td>
	    <td align="right" valign="top" background="skins/images/bg2.gif" width="28" ><img src="skins/images/teble_top_right.gif" alt="" /></td>
	    <td align="right" width="16" bgcolor="#C2D5E3"></td>
	  </tr>
	</table>
	<IFRAME frameBorder=0 id=frmright name=frmright scrolling=yes src="main_index.asp" class="main_iframe"></IFRAME>
	</TD>
	
</TR>
</TBODY>
</TABLE>
</body>
</html>