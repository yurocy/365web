function tips_pop(){var a=document.getElementById("winpop");var b=parseInt(a.style.height);if(b==0){a.style.display="block";show=setInterval("changeH('up')",2)}else{hide=setInterval("changeH('down')",2)}}function changeH(c){var a=document.getElementById("winpop");var b=parseInt(a.style.height);if(c=="up"){if(b<=100){a.style.height=(b+4).toString()+"px"}else{clearInterval(show)}}if(c=="down"){if(b>=4){a.style.height=(b-4).toString()+"px"}else{clearInterval(hide);a.style.display="none"}}}function setCookie(b,h){var c=new Date();var g=setCookie.arguments;var e=setCookie.arguments.length;var d=(e>2)?g[2]:null;var j=(e>3)?g[3]:null;var f=(e>4)?g[4]:null;var a=(e>5)?g[5]:false;if(d!=null){c.setTime(c.getTime()+(d*1000))}document.cookie=b+"="+escape(h)+((d==null)?"":("; expires="+c.toGMTString()))+((j==null)?"":("; path="+j))+((f==null)?"":("; domain="+f))+((a==true)?"; secure":"")}function $G(){var e=top.window.location.href;var b,c,a="";if(arguments[arguments.length-1]=="#"){b=e.split("#")}else{b=e.split("?")}if(b.length==1){c=""}else{c=b[1]}if(c!=""){gg=c.split("&");var d=gg.length;str=arguments[0]+"=";for(i=0;i<d;i++){if(gg[i].indexOf(str)==0){a=gg[i].replace(str,"");break}}}return a}function writeCookie(c,d,a){var b="";if(a!=null){b=new Date((new Date()).getTime()+a*3600000);b="; expires="+b.toGMTString()}document.cookie=c+"="+escape(d)+b}function PopStringBuffer(){this.cPopdata=new Array()}PopStringBuffer.prototype.append=function(a){this.cPopdata.push(a);return this};PopStringBuffer.prototype.toString=function(){return this.cPopdata.join("")};function getTipsPopHtml(){var b=document.createElement("div");b.style.display="none";b.setAttribute("id","winpop");var a=new PopStringBuffer();a.append('<div id=eMeng style="BORDER-RIGHT: #455690 1px solid; BORDER-TOP: #a6b4cf 1px solid; Z-INDEX:99999; LEFT: 0px; BORDER-LEFT: #a6b4cf 1px solid; BORDER-BOTTOM: #455690 1px solid; POSITION: absolute; TOP: 0px; BACKGROUND-COLOR: #c9d3f3">');a.append('<TABLE style="BORDER-TOP: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid" cellSpacing=0 cellPadding=0 width="100%" bgColor=#ffffff border=0>');a.append("<TBODY>");a.append("<TR bgColor=#ff6600>");a.append('<TD style="font-size: 12px; color: #0f2c8c" ></TD>');a.append('<TD style="font-weight: normal; font-size: 12px; color: #ffffff; padding-left: 4px; padding-top: 4px" vAlign=center width="100%">&nbsp;&nbsp;\u6e29\u99a8\u63d0\u793a</TD>');a.append('<TD style="padding-right: 2px; padding-top: 2px" vAlign=center align=right width=19><span title=\u5173\u95ed style="cursor:pointer;color:white;font-size:12px;font-weight:bold;margin-right:4px;" onclick=tips_pop() >\u00d7</span></TD>');a.append("</TR>");a.append("<TR>");a.append('<TD style="padding-right: 1px; padding-bottom: 1px" colSpan=3 >');a.append('<div style="BORDER-RIGHT: #b9c9ef 1px solid; PADDING-RIGHT: 13px; BORDER-TOP: #728eb8 1px solid; PADDING-LEFT: 13px; FONT-SIZE: 12px; PADDING-BOTTOM: 13px; BORDER-LEFT: #728eb8 1px solid; COLOR: #000000; PADDING-TOP: 18px; BORDER-BOTTOM: #b9c9ef 1px solid; HEIGHT: 100%">\u4e3a\u65b9\u4fbf\u60a8\u4e0b\u6b21\u8bbf\u95ee\uff0c\u8bf7\u6536\u85cf\u6211\u4eec\u7684\u7f51\u5740<br />');a.append('<div align=center style="word-break:break-all">');a.append('<span style="border:#a6b4cf 1px solid; padding:5px; background-color:#ff6600">');a.append('<a href="http://www.8168123.com/全讯网.url"><font color=#ffffff>\u6536\u85cf\u5230\u684c\u9762</font></a>');a.append("</span>");a.append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;");a.append('<span style="border:#a6b4cf 1px solid; padding:5px; background-color:#ff6600"">');a.append('<a href="http://www.8168123.com/" onclick="window.external.addFavorite(this.href,this.title);return false;" title="全讯网官方网站 www.8168123.com" rel="sidebar"><font color=#ffffff>\u52a0\u5165\u6536\u85cf\u5939</font></a>');a.append("</span>");a.append("</div>");a.append("</div>");a.append("</TD>");a.append("</TR>");a.append("</TBODY>");a.append("</TABLE>");a.append("</div>");b.innerHTML=a.toString();document.body.appendChild(b)}window.onload=function(){getTipsPopHtml();document.getElementById("winpop").style.height="0px";setTimeout("tips_pop()",200);setCookie("shutUp",1)};
