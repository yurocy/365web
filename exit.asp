<%
			Response.Cookies("memberuserid")=""
			Response.Cookies("sessionidstr")=""
			Response.Cookies("urlid_cookies")=""
					%>
<script language="javascript" src="js/main.js"></script>
<script language="javascript">writeCookie("urlid_cookies",'');parent.ifr.location.href = "login.asp";parent.Fun_GetPostData();</script>