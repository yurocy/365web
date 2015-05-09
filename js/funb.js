function g_cookie(name)
{var arg = name + "=";var alen = arg.length;var c= document.cookie;var clen =c.length;var i = 0
while (i < clen){var j = i + alen
if (c.substring(i, j) == arg)
return _GetCookieVal (j)
i = c.indexOf(" ", i) + 1
if (i == 0) break}
return null}
function s_cookie(name, value)
{var expdate = new Date();var argv = s_cookie.arguments;var argc = s_cookie.arguments.length
var expires = (argc > 2) ? argv[2] : null;var path = (argc > 3) ? argv[3] : null
var domain = (argc > 4) ? argv[4] : null;var secure = (argc > 5) ? argv[5] : false
if(expires!=null) expdate.setTime(expdate.getTime() + ( expires * 1000 ))
document.cookie = name + "=" + escape (value) +((expires == null) ? "" : ("; expires="+ expdate.toGMTString()))
+((path == null) ? "" : ("; path=" + path)) +((domain == null) ? "" : ("; domain=" + domain))
+((secure == true) ? "; secure" : "")}
function _GetCookieVal(offset)
{var c= document.cookie;var endstr = c.indexOf (";", offset)
if (endstr == -1) endstr = c.length
return unescape(c.substring(offset, endstr))}

var lunarInfo=new Array(
0x04bd8,0x04ae0,0x0a570,0x054d5,0x0d260,0x0d950,0x16554,0x056a0,0x09ad0,0x055d2,0x04ae0,0x0a5b6,0x0a4d0,0x0d250,0x1d255,0x0b540,0x0d6a0,0x0ada2,0x095b0,0x14977,0x04970,0x0a4b0,0x0b4b5,0x06a50,0x06d40,0x1ab54,0x02b60,0x09570,0x052f2,0x04970,0x06566,0x0d4a0,0x0ea50,0x06e95,0x05ad0,0x02b60,0x186e3,0x092e0,0x1c8d7,0x0c950,0x0d4a0,0x1d8a6,0x0b550,0x056a0,0x1a5b4,0x025d0,0x092d0,0x0d2b2,0x0a950,0x0b557,0x06ca0,0x0b550,0x15355,0x04da0,0x0a5d0,0x14573,0x052d0,0x0a9a8,0x0e950,0x06aa0,0x0aea6,0x0ab50,0x04b60,0x0aae4,0x0a570,0x05260,0x0f263,0x0d950,0x05b57,0x056a0,0x096d0,0x04dd5,0x04ad0,0x0a4d0,0x0d4d4,0x0d250,0x0d558,0x0b540,0x0b5a0,0x195a6,0x095b0,0x049b0,0x0a974,0x0a4b0,0x0b27a,0x06a50,0x06d40,0x0af46,0x0ab60,0x09570,0x04af5,0x04970,0x064b0,0x074a3,0x0ea50,0x06b58,0x055c0,0x0ab60,0x096d5,0x092e0,0x0c960,0x0d954,0x0d4a0,0x0da50,0x07552,0x056a0,0x0abb7,0x025d0,0x092d0,0x0cab5,0x0a950,0x0b4a0,0x0baa4,0x0ad50,0x055d9,0x04ba0,0x0a5b0,0x15176,0x052b0,0x0a930,0x07954,0x06aa0,0x0ad50,0x05b52,0x04b60,0x0a6e6,0x0a4e0,0x0d260,0x0ea65,0x0d530,0x05aa0,0x076a3,0x096d0,0x04bd7,0x04ad0,0x0a4d0,0x1d0b6,0x0d250,0x0d520,0x0dd45,0x0b5a0,0x056d0,0x055b2,0x049b0,0x0a577,0x0a4b0,0x0aa50,0x1b255,0x06d20,0x0ada0)
var Gan=new Array("甲","乙","丙","丁","戊","己","庚","辛","壬","癸")
var Zhi=new Array("子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥")
var cmStr = new Array('日','正','二','三','四','五','六','七','八','九','十','冬','腊')
var nStr1 = new Array('日','一','二','三','四','五','六','七','八','九','十')
var now;var SY;var SM;var SD
function cyclical(num) { return(Gan[num%10]+Zhi[num%12]) }
function lYearDays(y) {
var i, sum = 348
for(i=0x8000; i>0x8; i>>=1) sum += (lunarInfo[y-1900] & i)? 1: 0
return(sum+leapDays(y))}
function leapDays(y) {
   if(leapMonth(y))  return((lunarInfo[y-1900] & 0x10000)? 30: 29)
   else return(0)}
function leapMonth(y) { return(lunarInfo[y-1900] & 0xf)}
function monthDays(y,m) { return( (lunarInfo[y-1900] & (0x10000>>m))? 30: 29 )}
function Lunar(objDate) {
var i, leap=0, temp=0
var baseDate = new Date(1900,0,31)
var offset   = (objDate - baseDate)/86400000
this.dayCyl = offset + 40
this.monCyl = 14
for(i=1900; i<2050 && offset>0; i++) {
temp = lYearDays(i)
offset -= temp
this.monCyl += 12}
if(offset<0) {
offset += temp;
i--;
this.monCyl -= 12}
this.year = i
this.yearCyl = i-1864
leap = leapMonth(i)
this.isLeap = false
for(i=1; i<13 && offset>0; i++) {
if(leap>0 && i==(leap+1) && this.isLeap==false)
{ --i; this.isLeap = true; temp = leapDays(this.year); }
else
{ temp = monthDays(this.year, i); }
if(this.isLeap==true && i==(leap+1)) this.isLeap = false
offset -= temp
if(this.isLeap == false) this.monCyl ++}
if(offset==0 && leap>0 && i==leap+1)
if(this.isLeap)
{ this.isLeap = false; }
else
{ this.isLeap = true; --i; --this.monCyl;}
if(offset<0){ offset += temp; --i; --this.monCyl; }
this.month = i
this.day = offset + 1}
function YYMMDD() {    return('<font style="font-size:14px">'+SY+'年'+(SM+1)+'月'+SD+'日</font>')}
function weekday(){
    var cl = '<font style="font-size:12px"';
    if (now.getDay() == 0) cl += ' color=#DF0A10';
    if (now.getDay() == 6) cl += ' color=#DF0A10';
    return(cl+'>星期'+nStr1[now.getDay()]+'</font>');
}
function cDay(m,d){
var nStr2 = new Array('初','十','廿','卅','　');var s
s= cmStr[m]+'月'
switch (d) {
  case 10:s += '初十'; break;
  case 20:s += '二十'; break;
  case 30:s += '三十'; break;
  default:s += nStr2[Math.floor(d/10)]; s += nStr1[Math.round(d%10)];
}return(s)}
function solarDay(){
var sTermInfo = new Array(0,21208,42467,63836,85337,107014,128867,150921,173149,195551,218072,240693,263343,285989,308563,331033,353350,375494,397447,419210,440795,462224,483532,504758)
var solarTerm = new Array("小寒","大寒","立春","雨水","惊蛰","春分","清明","谷雨","立夏","小满","芒种","夏至","小暑","大暑","立秋","处暑","白露","秋分","寒露","霜降","立冬","小雪","大雪","冬至")
var lFtv = new Array("0101*春节","0115 元宵节","0505 端午节","0707 七夕","0715 中元节","0815 中秋节","0909 重阳节","1208 腊八节","1224 小年","0100*除夕")
var sFtv = new Array("0101*元旦","0214 情人节","0308 妇女节","0312 植树节","0401 愚人节","0501 劳动节","0504 青年节","0512 护士节","0601 儿童节","0701 建党节","0801 建军节","0910 教师节","1001*国庆节","1101 万圣节","1108 记者日","1225 圣诞节","0617 父亲节")
  var sDObj = new Date(SY,SM,SD);
  var lDObj = new Lunar(sDObj);
  var lDPOS = new Array(3)
  var festival='',solarTerms='',solarFestival='',lunarFestival='',solarTerms='',tmp1,tmp2;

  for(i in lFtv)
  if(lFtv[i].match(/^(\d{2})(.{2})([\s\*])(.+)$/)) {
   tmp1=Number(RegExp.$1)-lDObj.month
   tmp2=Number(RegExp.$2)-lDObj.day
   if(tmp1==0 && tmp2==0) lunarFestival=RegExp.$4}
  if(lunarFestival=='') {
  for(i in sFtv)
  if(sFtv[i].match(/^(\d{2})(\d{2})([\s\*])(.+)$/)){
   tmp1=Number(RegExp.$1)-(SM+1)
   tmp2=Number(RegExp.$2)-SD
   if(tmp1==0 && tmp2==0) solarFestival = RegExp.$4}
  if(solarFestival =='') {
	  tmp1 = new Date((31556925974.7*(SY-1900)+sTermInfo[SM*2+1]*60000)+Date.UTC(1900,0,6,2,5))
  tmp2 = tmp1.getUTCDate()
  if (tmp2==SD) solarTerms = solarTerm[SM*2+1]
  tmp1 = new Date((31556925974.7*(SY-1900)+sTermInfo[SM*2]*60000)+Date.UTC(1900,0,6,2,5))
  tmp2= tmp1.getUTCDate()

  if (tmp2==SD) solarTerms = solarTerm[SM*2]
	if(solarTerms=='') sFtv='';else sFtv=solarTerms
  } else sFtv=solarFestival
  } else sFtv=lunarFestival
  if(sFtv=='')
	sTermInfo=cyclical(lDObj.year-1900+36)+'年 '+cDay(lDObj.month,lDObj.day)
  else sTermInfo=cDay(lDObj.month,lDObj.day)+' <font color=#DF0A10>'+sFtv+'</font>'
  return(sTermInfo)
}
function CurentTime()
{var now = new Date();var hh = now.getHours();var mm = now.getMinutes();var ss = now.getTime() % 60000;
ss = (ss - (ss % 1000)) / 1000;
if(hh==0&&mm==0&&ss==0) showcal(1)
var clock = hh+':';
if (mm < 10) clock += '0';
clock += mm;return(clock)}
function refreshCalendarClock() {document.getElementById('ClockTime').innerHTML = CurentTime()}
function showcal(t) {
now = new Date();SY = now.getFullYear();SM = now.getMonth();SD = now.getDate();
var hh = now.getHours();var mm = now.getMinutes();var ss = now.getTime() % 60000;
ss = (ss - (ss % 1000)) / 1000;hh+=':'
if (mm < 10) hh+='0'
hh += mm
var str='<a href="javascript:clock()" target=_self  title=点击设置闹钟>闹钟</a>'
if(t==1) document.getElementById('rili').innerHTML=str
else document.write(str)}
function clock() {window.open('/click.htm','time','left='+(screen.width-320)+',height=310,width=305,status=no,toolbar=no,menubar=no')}


 var frmact = new Array();
 frmact[1]="http://www.baidu.com/s";
 frmact[2]="http://news.baidu.com/ns"
 frmact[3]="http://post.baidu.com/f";
 frmact[4]="http://mp3.baidu.com/m";
 frmact[6]="http://image.baidu.com/i";
 var frmtn = new Array();
 frmtn[1]="tupianguan_pg";
 frmtn[2]="news";
 frmtn[3]="baiduPostSearch";
 frmtn[4]="baidump3";
 frmtn[6]="baiduimage";
 var frmct = new Array();
 frmct[1] = "";
 frmct[2] = "";
 frmct[3] = "352321536";
 frmct[4] = "134217728";
 frmct[6] = "201326592";
 var frmlm = new Array();
 frmlm[1] = "";
 frmlm[2] = "";
 frmlm[3] = "65536";
 frmlm[4] = "-1";
 frmlm[6] = "-1";
 var frmz = new Array();
 frmz[1] = "";
 frmz[2] = "";
 frmz[3] = "";
 frmz[4] = "";
 frmz[6] = "0";
 var frmrn = new Array();
 frmrn[1] = "";
 frmrn[2] = "";
 frmrn[3] = "10";
 frmrn[4] = "";
 frmrn[6] = "";

 var bd_idx="1";

function bd_chg_idx(idx)
{
 bd_idx=idx;
}

function gowhere(formname)
{
 var url;
 var idx = bd_idx;
 if (frmact[idx] == null || frmact[idx] == "")  idx = "1";
 url = frmact[idx];
 if (frmtn[idx] != "") document.search_form.tn.value = frmtn[idx];
 if (frmct[idx] != "") document.search_form.ct.value = frmct[idx];
 if (frmlm[idx] != "") document.search_form.lm.value = frmlm[idx];
 if (frmz[idx] != "") document.search_form.z.value = frmz[idx];
 if (frmrn[idx] != "") document.search_form.rn.value = frmrn[idx];
 formname.method = "get";
 formname.action = url;
 return true;
} 



var stab=1,srh_htm;
function sh(i) {
if(i!=stab){
em('lt'+i).className='active'
em('lt'+stab).className=''}
stab=i;
document.search2_form.by.value = i;
bd_chg_idx(i);
document.all.srh_i.src='i/search/srh_'+i+'.gif';
document.all.gl_i.src='i/search/gl_'+i+'.gif';

}
function index_search()
{
if(search2_form.by.value=="")window.open("http://www.google.cn/search?hl=zh-CN&meta=&aq=f&oq=&q="+search2_form.in_word.value);
if(search2_form.by.value==1)window.open("http://www.google.cn/search?hl=zh-CN&meta=&aq=f&oq=&q="+search2_form.in_word.value);
if(search2_form.by.value==4)window.open("http://one.cn.yahoo.com/s?v=music&ei=gbk&pid=ysearch&x=&source=ysearch_music_result_topsearch&p="+search2_form.in_word.value);
if(search2_form.by.value==6)window.open("http://images.google.cn/images?prog=aff&oe=GB2312&IE=GB2312&hl=zh-CN&q="+search2_form.in_word.value);
if(search2_form.by.value==2)window.open("http://news.google.cn/news?prog=aff&oe=GB2312&IE=GB2312&hl=zh-CN&q="+search2_form.in_word.value);
if(search2_form.by.value==3)window.open("http://cn.news.yahoo.com/search1.html?p="+search2_form.in_word.value);
return false;
}
function ShowPannel(index){
	var i=index
	if(i==1){
		document.getElementById("indexmain").style.display="";
		document.getElementById("favmain").style.display="none";
	}
	if(i==2){
		document.getElementById("favmain").style.display="";
		document.getElementById("indexmain").style.display="none";
	}
	
}
function ChkAdDiv(divid,value){
	var  i=divid; 
	var  v=value;
	document.getElementById(i).innerHTML=document.getElementById(v).innerHTML;
}  
document.writeln("")

//快速导航

function ks(n){

 document.scripts[1].src = "ks/?key=" + n;

}
