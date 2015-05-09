<%
Response.Cookies("urlid_cookies")="1,2,3,4,5"
'response.write request.Cookies("urlid_cookies")
%>
<script>
function getCookie(name)
{
	var cname = name + "=";
	var dc = document.cookie;
	if (dc.length > 0) 
	{
		begin = dc.indexOf(cname);
		if (begin != -1) 
		{
			begin += cname.length;
			end = dc.indexOf(";", begin);
			if (end == -1) end = dc.length;
			return dc.substring(begin, end);
		}
	}
	return null;
}


document.write(getCookie("urlidcookies"));</script>