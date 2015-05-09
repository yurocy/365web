<!--#include file="Conn.asp" --><!--#include file="inc/config.Asp"--><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>广告联系</title>
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
.menu1 ul{ list-style:none;display:block;position:absolute;left:32px;top:-5px;width:80px;background:#42C4DB;padding-top:5px;padding-bottom:5px;}
.menu1 ul li a{ color:#ffffff;display:block;height:22px;line-height:22px;font-size:9pt;margin-top:1px;margin-left:5px;margin-right:5px;background:#6AD2E5 url(/images/m_arrow.gif) no-repeat 5px center;}
.menu1 ul li a:hover{ color:#000000;text-decoration:none;background:#ffffff url(/images/m_arrow1.gif) no-repeat 5px center;}

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
</head>


<body>

<div class="main">
<table width="845" height="30" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="100" background="/images/bx001.gif" class="w01"><div align="right"><iframe src="time.html" width="150" height="15" frameborder="0" marginwidth="0" marginheight="0" scrolling="no" name="I2"></iframe> </div></td>
   		 <td width="690" background="/images/bx001.gif"><iframe src="live.html" width="690" height="28" frameborder="0" marginwidth="0" marginheight="0" scrolling="no"></iframe></td>
  </tr>
</table>
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
<td width="167" bgcolor="#FFCCEE" background="/images/bx004.gif"><div align="center"><a href="/"><img src="images/logo.gif" width="157" height="83" /></a></div></td>
  <td width="574" background="/images/bx004.gif"><div id="ad_top" align="right">
     <%=showinfo("横幅广告")%>
    </div></td>
    <td width="95" background="/images/bx004.gif"><table width="73" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="22" bgcolor="FFCC33" class="w01"><div align="center"><a href="#" onclick="window.external.AddFavorite(document.location.href,'<%=webtitle%>');return false;" style="color:#000000;">收藏本站</a></div></td>
      </tr>
      <tr>
        <td height="10"> </td>
      </tr>
      <tr>
        <td height="22" bgcolor="FFCC33" class="w01"><div align="center"><a href="#" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('<%=weburl%>');return false;"  style="color:#000000;">设为首页</a></div></td>
      </tr>
    </table></td>
  </tr>
  <tr><td bgcolor="#636463" colspan="3" style="height: 5px;"></td></tr>
</table>
<!--头部-->

<table width="845" height="8" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px solid #999999;margin-bottom:10px;">
  <tr>
    <td align="left"><%=showinfo("广告联系")%></td>
  </tr>
</table>

<div class="main" style="font-size:12px;">
<%=showinfo("友情链接")%>
</div>
<div class="main">
  <div id="footer">
    <p>客服热线:<%=webmobi%> QQ:<%=webqq%> 客服时间:7×24小时 Copyright  2009 <%=weburl%> Inc. All Rights Reserved.技术支持 <br />
    </p>
    <p class="red">注意:网上有诈骗.所有在本站刊登广告的网站和内容，一概与本站无关，请各位网友密切注意!<br />
    </p>
    
  </div>
</div>
</body>
</html>


