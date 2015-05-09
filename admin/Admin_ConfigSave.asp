<!--#include file="../Conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="../inc/function.asp"-->
<%

if trim(request("Action"))="Save" then


 
 
 
 
 
str=str& "<" & "%" & vbcrlf
str=str&"Const webname=" & chr(34) & trim(request("webname")) & chr(34) & vbcrlf
str=str&"Const webtitle=" & chr(34) & trim(request("webtitle")) & chr(34) & vbcrlf
str=str&"Const weburl=" & chr(34) & trim(request("weburl")) & chr(34) & vbcrlf
str=str&"Const weblogo=" & chr(34) & trim(request("weblogo")) & chr(34) & vbcrlf 
str=str&"Const webqq=" & chr(34) & trim(request("webqq")) & chr(34) & vbcrlf
str=str&"Const Webmobi=" & chr(34) & trim(request("Webmobi")) & chr(34) & vbcrlf
str=str&"Const keywords=" & chr(34) & trim(request("keywords")) & chr(34) & vbcrlf
str=str&"Const description=" & chr(34) & trim(request("description")) & chr(34) & vbcrlf
str=str&"Const COPYRIGHT=" & chr(34) & trim(request("COPYRIGHT")) & chr(34) & vbcrlf
str=str&"%" & ">" & vbcrlf

 
 
      set stm=server.CreateObject("adodb.stream")
    stm.Type=2 '以本模式读取
    stm.mode=3
    stm.charset="GB2312"
    stm.open
    stm.WriteText str
    stm.SaveToFile Server.mappath("../inc/config.asp"),2    
    stm.flush
    stm.Close
    set stm=nothing
 response.Redirect "Admin_SetSite.asp?Success=True"
end if
%>