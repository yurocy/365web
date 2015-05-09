var zXml = { useActiveX: (typeof ActiveXObject != "undefined"), useXmlHttp: (typeof XMLHttpRequest != "undefined") }; zXml.ARR_XMLHTTP_VERS = ["MSXML2.XmlHttp.6.0", "MSXML2.XmlHttp.3.0"]; function zXmlHttp() { }
zXmlHttp.createRequest = function () {
    if (zXml.useXmlHttp) return new XMLHttpRequest(); if (zXml.useActiveX) {
        if (!zXml.XMLHTTP_VER) { for (var i = 0; i < zXml.ARR_XMLHTTP_VERS.length; i++) { try { new ActiveXObject(zXml.ARR_XMLHTTP_VERS[i]); zXml.XMLHTTP_VER = zXml.ARR_XMLHTTP_VERS[i]; break; } catch (oError) { } } }
        if (zXml.XMLHTTP_VER) return new ActiveXObject(zXml.XMLHTTP_VER);
    }
    alert("对不起，您的电脑不支持 XML 插件，请安装好或升级浏览器。");
};


function MM_reloadPage(init) { //reloads the window if Nav4 resized 
    if (init == true) with (navigator) {
        if ((appName == "Netscape") && (parseInt(appVersion) == 4)) {
            document.MM_pgW = innerWidth; document.MM_pgH = innerHeight; onresize = MM_reloadPage;
        }
    }
    else if (innerWidth != document.MM_pgW || innerHeight != document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_findObj(n, d) { //v4.0 
    var p, i, x; if (!d) d = document; if ((p = n.indexOf("?")) > 0 && parent.frames.length) {
        d = parent.frames[n.substring(p + 1)].document; n = n.substring(0, p);
    }
    if (!(x = d[n]) && d.all) x = d.all[n]; for (i = 0; !x && i < d.forms.length; i++) x = d.forms[i][n];
    for (i = 0; !x && d.layers && i < d.layers.length; i++) x = MM_findObj(n, d.layers[i].document);
    if (!x && document.getElementById) x = document.getElementById(n); return x;
}

function MM_showHideLayers() { //v3.0 
    var i, p, v, obj, args = MM_showHideLayers.arguments;
    for (i = 0; i < (args.length - 2); i += 3) if ((obj = MM_findObj(args[i])) != null) {
        v = args[i + 2];
        if (obj.style) { obj = obj.style; v = (v == 'show') ? 'visible' : (v = 'hide') ? 'hidden' : v; }
        obj.visibility = v;
    }
}

//取元素相对于body的坐标
function ZB() {
    this.left = 0,
 this.top = 0,
 this.width = 0,
 this.height = 0
};

function getCurrentStyle(o, style) {
    var number = parseInt(o.currentStyle[style]);
    return isNaN(number) ? 0 : number;
}
function getComputedStyle(o, style) {
    return parseInt(document.defaultView.getComputedStyle(o, null).getPropertyValue(style));
}

function GetZB(obj) {
    var o = obj;
    var oLTWH = new ZB();
    oLTWH.width = o.offsetWidth;
    oLTWH.height = o.offsetHeight;
    oLTWH.left = o.offsetLeft;
    oLTWH.top = o.offsetTop;
    while (true) {
        o = o.offsetParent;
        if (o == (document.body && null)) break;
        oLTWH.left += o.offsetLeft;
        oLTWH.top += o.offsetTop;
        /*
        if(document.all)
        {
        oLTWH.left+=getCurrentStyle(o,"borderLeftWidth");
        oLTWH.top+=getCurrentStyle(o,"borderTopWidth");
        }
        else
        {
        oLTWH.left+=getComputedStyle(o,"border-left-width");
        oLTWH.top+=getComputedStyle(o,"border-top-width");
        }
        */
    }
    return oLTWH;
}


var obj2;
var color_id;
var color_c = ""
function showit1(obj, color_id1, myid, f) {

    color_id = color_id1;
    obj2 = obj

    color_c = stylecolor.split(",")[color_id];

    obj2.className = color_c + "_td2";

    var zb = GetZB(obj2);

    var myjs1;
    if (myjs[myid] == null) {
        myjs1 = "";
    }
    else {
        myjs1 = myjs[myid];
    }

    var ststst = "";

    if (document.getElementById("a_" + myid).innerText == "全讯网全部网址" || document.getElementById("a_" + myid).innerText == "全訊網全部網址") {
        ststst += "<tr><td class='href1' onmouseover='javascript:color1(this)' onmouseout='javascript:color2(this)'><b><a href='javascript:;' onclick='javascript:urltest1(" + myid + ");return false;' title='点击打开'>点击打开</a></b></td></tr>";
        ststst += "<tr><td class='href1' onmouseover='javascript:color1(this)' onmouseout='javascript:color2(this)'><a href='javascript:;' onclick='javascript:addf(this," + myid + ");return false;' title='加入收藏'>加入收藏</a></td></tr>";
        ststst = "<table align='center' bgcolor='#42c4db' cellspacing='1' cellpadding='1' width='81'>" + ststst + "</table>";
    }
    else { 
    if (myjs1 != "") {
        if (myjs1.split("|").length > 15) {
            ststst += "<tr>";
            ststst += "<td class='href1' onmouseover='javascript:color1(this)' onmouseout='javascript:color2(this)'><b><a href='javascript:;' onclick='javascript:urltest(" + myid + ");return false;' title='点击测速'>点击测速</a></b></td>";
            if (f == 1) {
                ststst += "<td class='href1' onmouseover='javascript:color1(this)' onmouseout='javascript:color2(this)'><a href='javascript:;' onclick='javascript:delf(this," + myid + ");return false;' title='删除收藏'>删除收藏</a></td>";
            }
            else {
                ststst += "<td class='href1' onmouseover='javascript:color1(this)' onmouseout='javascript:color2(this)'><a href='javascript:;' onclick='javascript:addf(this," + myid + ");return false;' title='加入收藏'>加入收藏</a></td>";
            }
            ststst += "</tr><tr>";
            for (var i = 0; i < myjs1.split("|").length; i++) {
                ststst += "<td class='href1' onmouseover='javascript:color1(this)' onmouseout='javascript:color2(this)'><a href='" + myjs1.split("|")[i].split(",")[1] + "' target='_blank' title='" + myjs1.split("|")[i].split(",")[0] + "'>" + myjs1.split("|")[i].split(",")[0] + "</a></td>";
                if (i % 2 == 1) {
                    ststst += "</tr><tr>";
                }
            }
            ststst = "<table align='center' bgcolor='#42c4db' cellspacing='1' cellpadding='1' width='162'>" + ststst + "</table>";
         }
        else {
            ststst += "<tr><td class='href1' onmouseover='javascript:color1(this)' onmouseout='javascript:color2(this)'><b><a href='javascript:;' onclick='javascript:urltest(" + myid + ");return false;' title='点击测速'>点击测速</a></b></td></tr>";
            for (var i = 0; i < myjs1.split("|").length; i++) {
                ststst += "<tr><td class='href1' onmouseover='javascript:color1(this)' onmouseout='javascript:color2(this)'><a href='" + myjs1.split("|")[i].split(",")[1] + "' target='_blank' title='" + myjs1.split("|")[i].split(",")[0] + "'>" + myjs1.split("|")[i].split(",")[0] + "</a></td></tr>";
            }
            if (f == 1) {
                ststst += "<tr><td class='href1' onmouseover='javascript:color1(this)' onmouseout='javascript:color2(this)'><a href='javascript:;' onclick='javascript:delf(this," + myid + ");return false;' title='删除收藏'>删除收藏</a></td></tr>";
            }
            else {
                ststst += "<tr><td class='href1' onmouseover='javascript:color1(this)' onmouseout='javascript:color2(this)'><a href='javascript:;' onclick='javascript:addf(this," + myid + ");return false;' title='加入收藏'>加入收藏</a></td></tr>";
            }
            ststst = "<table align='center' bgcolor='#42c4db' cellspacing='1' cellpadding='1' width='81'>" + ststst + "</table>";
        }
    }
    else {
        if (f == 1) {
            ststst += "<tr><td class='href1' onmouseover='javascript:color1(this)' onmouseout='javascript:color2(this)'><a href='javascript:;' onclick='javascript:delf(this," + myid + ");return false;' title='删除收藏'>删除收藏</a></td></tr>";
        } else {
            ststst += "<tr><td class='href1' onmouseover='javascript:color1(this)' onmouseout='javascript:color2(this)'><a href='javascript:;' onclick='javascript:addf(this," + myid + ");return false;' title='加入收藏'>加入收藏</a></td></tr>";
        }
        ststst = "<table align='center' bgcolor='#42c4db' cellspacing='1' cellpadding='1' width='81'>" + ststst + "</table>";
    }    
    }
    
   

    document.getElementById("dddddd").innerHTML = ststst;
    document.getElementById("hk1532888_menu1").style.left = zb.left + (zb.width - document.getElementById("hk1532888_menu1").offsetWidth) / 2;
    document.getElementById("hk1532888_menu1").style.top = zb.top + obj2.offsetHeight;
    MM_showHideLayers("hk1532888_menu1", "", "show");
}



function showit1_(obj, color_id1, myid) {
    color_id = color_id1;
    obj2 = obj

    color_c = stylecolor.split(",")[color_id];

    obj2.className = color_c + "_td2";

    var zb = GetZB(obj2);

    var myjs1;
    if (myjs[myid] == null) {
        myjs1 = "";
    }
    else {
        myjs1 = myjs[myid];
    }

    var ststst = "";

    ststst += "<tr><td class='href1' onmouseover='javascript:color1(this)' onmouseout='javascript:color2(this)'><a href='javascript:;' onclick='javascript:delf(" + myid + ");return false;' title='删除收藏'>删除收藏</a></td></tr>";
    ststst = "<table align='center' bgcolor='#42c4db' cellspacing='1' cellpadding='1' width='81'>" + ststst + "</table>";

    document.getElementById("dddddd").innerHTML = ststst;
    document.getElementById("hk1532888_menu1").style.left = zb.left + (zb.width - document.getElementById("hk1532888_menu1").offsetWidth) / 2;
    document.getElementById("hk1532888_menu1").style.top = zb.top + obj2.offsetHeight;
    MM_showHideLayers("hk1532888_menu1", "", "show");
}



function hideit1() {
    obj2.className = color_c + "_td1";
    MM_showHideLayers("hk1532888_menu1", "", "hide");
}

function showit2() {
    obj2.className = color_c + "_td2";
    MM_showHideLayers("hk1532888_menu1", "", "show");
}

function hideit2() {
    obj2.className = color_c + "_td1";
    MM_showHideLayers("hk1532888_menu1", "", "hide");
}

function color1(obj) {
    obj.className = "href2";
}

function color2(obj) {
    obj.className = "href1";
}

function urltest(myid) {
    if (document.getElementById("a_" + myid).innerText == "全讯网全部网址" || document.getElementById("a_" + myid).innerText == "全訊網全部網址") {
        urltest1(myid);
        return;
    }
    document.getElementById("test_title").innerText = document.getElementById("a_" + myid).innerText;
    document.getElementById("myiframe").src = "frame_speed2010.aspx?id=" + myid;
    var zb = GetZB(document.getElementById("a_" + myid).parentElement);
    document.getElementById("hk1532888_menu2").style.left = zb.left-2;
    document.getElementById("hk1532888_menu2").style.top = zb.top + document.getElementById("a_" + myid).parentElement.offsetHeight;
}

function urltest1(myid) {
    document.getElementById("test_title").innerText = document.getElementById("a_" + myid).innerText;
    document.getElementById("myiframe").src = "frame_speed_index.htm?id=" + myid;
    var zb = GetZB(document.getElementById("a_" + myid).parentElement);
    document.getElementById("hk1532888_menu2").style.left = zb.left - 2;
    document.getElementById("hk1532888_menu2").style.top = zb.top + document.getElementById("a_" + myid).parentElement.offsetHeight;
}

function hidetest() {
    MM_showHideLayers("hk1532888_menu2", "", "hide");
}

function addf(obj,myid) {
    var oXmlHttp = zXmlHttp.createRequest();
    var url1 = "back_fav2010.aspx?addid=" + myid + "&t=" + Date.parse(new Date());
    oXmlHttp.open("get", url1, true);
    oXmlHttp.onreadystatechange = function () {
        if (oXmlHttp.readyState == 4 && oXmlHttp.status == 200) {
            var responseText1 = oXmlHttp.responseText;
            if (responseText1 == "") {
                parent.showf0();
            }
            else {
                document.getElementById("tdffff").innerHTML = responseText1;

                var zb = GetZB(obj.parentElement.parentElement.parentElement);
                document.getElementById("dmsg").innerHTML = "收藏 <b>『" + document.getElementById("a_" + myid).innerText + "』</b> 成功！";
                document.getElementById("hk1532888_menu4").style.left = zb.left - 18;
                document.getElementById("hk1532888_menu4").style.top = zb.top - 80;
                MM_showHideLayers("hk1532888_menu4", "", "show");
                clearTimeout(vhide4);
                vhide4 = setTimeout("hide4()", 3000);
            }
        }
    };
    oXmlHttp.setRequestHeader("If-Modified-Since", "0");
    oXmlHttp.send(null);
    //document.getElementById("iframeU").src = "index_2010_chr.aspx?addid=" + myid;
}

function delf(obj,myid) {
    var oXmlHttp = zXmlHttp.createRequest();
    var url1 = "back_fav2010.aspx?delid=" + myid + "&t=" + Date.parse(new Date());
    oXmlHttp.open("get", url1, true);
    oXmlHttp.onreadystatechange = function () {
        if (oXmlHttp.readyState == 4 && oXmlHttp.status == 200) {
            var responseText1 = oXmlHttp.responseText;
            if (responseText1 == "") {
                parent.showf0();
            }
            else {
                var zb = GetZB(obj.parentElement.parentElement.parentElement);
                var zbleft = zb.left;
                var zbtop = zb.top;
                document.getElementById("tdffff").innerHTML = responseText1;
                document.getElementById("dmsg").innerHTML = "取消 <b>『" + document.getElementById("a_" + myid).innerText + "』</b> 收藏成功！";
                document.getElementById("hk1532888_menu4").style.left = zbleft - 18;
                document.getElementById("hk1532888_menu4").style.top = zbtop - 80;
                MM_showHideLayers("hk1532888_menu4", "", "show");
                clearTimeout(vhide4);
                vhide4 = setTimeout("hide4()", 3000);
            }
        }
    };
    oXmlHttp.setRequestHeader("If-Modified-Since", "0");
    oXmlHttp.send(null);
    //document.getElementById("iframeU").src="index_2010_chr.aspx?delid=" + myid;
}

function showf0() {
    document.getElementById("tdffff").innerHTML = "<table cellspacing='1' cellpadding='1' bgcolor='#bbbbbb'><tr><td style='color:#999999;text-align:center;background-color:#f5f5f5;height:63px;width:810px;font-size:13px;'>注册收藏夹用户,可以添加自定义网站,用户可以多处重复登陆,手机用户请登录：wap.1532777.com</td></tr></table>"
}

function reinitIframe(h1) {
    document.getElementById("myiframe").style.width = 440;
    document.getElementById("myiframe").style.height = h1 + 50;
    MM_showHideLayers("hk1532888_menu2", "", "show");
}

function reinitIframe1() {
    document.getElementById("myiframe").style.height = 270;
    document.getElementById("myiframe").style.width = 756;
    MM_showHideLayers("hk1532888_menu2", "", "show");
}


function showt(tid1,tdtd_sel1) {
    tdtd_sel = tdtd_sel1;
    var tid;
    if (tid1 == "0") {
        for (var i = 0; i < t_c; i++) {
            document.getElementById("tc_" + i).style.display = "";
        }
        for (var i = 0; i < 8; i++) {
            document.getElementById("tchr_" + i).style.display = "";
            document.getElementById("tc0_" + i).style.display = "";
            document.getElementById("ad01_" + i).style.display = "";
            document.getElementById("ad02_" + i).style.display = "";
        }
    }
    else {
        for (var i = 0; i < t_c; i++) {
            tid = parseInt(document.getElementById("tc_" + i).tid);
            if (tid == tid1) {
                document.getElementById("tc_" + i).style.display = "";
            }
            else {
                document.getElementById("tc_" + i).style.display = "none";
            }
        }
        for (var i = 0; i < 8; i++) {
            tid = parseInt(document.getElementById("tchr_" + i).tid);
            if (tid == tid1) {
                document.getElementById("tchr_" + i).style.display = "";
                document.getElementById("tc0_" + i).style.display = "";
                document.getElementById("ad01_" + i).style.display = "";
                document.getElementById("ad02_" + i).style.display = "";
            }
            else {
                document.getElementById("tchr_" + i).style.display = "none";
                document.getElementById("tc0_" + i).style.display = "none";
                document.getElementById("ad01_" + i).style.display = "none";
                document.getElementById("ad02_" + i).style.display = "none";
            }
        }
    }
    
}

var tdtd_sel = 0;

function overtd(tid) {
    for (var i = 0; i < 9; i++) {
        if (i == tid) {
            document.getElementById("tdtd_" + i).className = "tdtd2";
        }
        else {
            document.getElementById("tdtd_" + i).className = "tdtd1";
        }
    }
}

function outtd() {
    for (var i = 0; i < 9; i++) {
        if (i == tdtd_sel) {
            document.getElementById("tdtd_" + i).className = "tdtd2";
        }
        else {
            document.getElementById("tdtd_" + i).className = "tdtd1";
        }
    }
}


function over_(obj) {
    obj.className = "chr1_";
}
function out_(obj) {
    obj.className = "chr1";
}
function over_1(obj) {
    obj.className = "chr0_";
}
function out_1(obj) {
    obj.className = "chr0";
}

function showchrsearch(tid,index_) {
    var st;
    st = "<table id='tchr_" + index_ + "' cellpadding='1' cellspacing='1' style='width:845px;background-color:#bbbbbb;' align='center' tid=" + tid + "><tr><td style='background-color:White;'>";
    st += "<table cellpadding='0' cellspacing='0' style='width:100%'><tr><td><img src='images/goldreach.gif' style='margin:3px;height:17px;' /></td>";
    st += "<td><div class='chr0' onmouseover='javascript:over_1(this)' onmouseout='javascript:out_1(this)' onclick='javascript:loadx(" + tid + "," + 0 + ")'>全部显示</div></td>";
    for (var i = 0; i < 26; i++) {
        st += "<td><div class='chr1' onmouseover='javascript:over_(this)' onmouseout='javascript:out_(this)' onclick='javascript:loadx(" + tid + "," + (i + 65) + ")'>" + String.fromCharCode(i + 65) + "</div></td>";
    }
    for (var i = 0; i < 10; i++) {
        st += "<td><div class='chr1' onmouseover='javascript:over_(this)' onmouseout='javascript:out_(this)' onclick='javascript:loadx(" + tid + "," + (i + 48) + ")'>" + String.fromCharCode(i + 48) + "</div></td>";
    }
    st += "</tr></table>";
    st += "</td></tr></table>";
    document.write(st);
}

var tid_js = new Array();

function loadx(tid1, chr1) {
    var k;
    for (var i = 0; i < t_c; i++) {
        tid = parseInt(document.getElementById("tc_" + i).tid);
        if (tid == tid1) {
            if (tid_js[i] == null) {
                tid_js[i] = new Array();
                k = 0;
                for (var j = 0; j < document.getElementById("tc1_" + i).getElementsByTagName("td").length; j++) {
                    className1 = document.getElementById("tc1_" + i).getElementsByTagName("td")[j].className;
                    if (document.getElementById("tc1_" + i).getElementsByTagName("td")[j].innerHTML != "&nbsp;") {
                        tid_js[i][k] = document.getElementById("tc1_" + i).getElementsByTagName("td")[j].outerHTML;
                        k += 1;
                    }
                }
            }
            buildtable(i, chr1, document.getElementById("tc1_" + i).className.replace("_table1","") + "_td1");
        }
    }
}

function buildtable(i, chr1, className1) {

    var bbbb = new Array();
    var k = 0;
    if (chr1 == 0) {
        for (var j = 0; j < tid_js[i].length; j++) {
             bbbb[k] = tid_js[i][j];
             k += 1;
        }
    }
    else {
        for (var j = 0; j < tid_js[i].length; j++) {
            if (tid_js[i][j].indexOf('p="' + String.fromCharCode(chr1) + '"') != -1) {
                bbbb[k] = tid_js[i][j];
                k += 1;
            }
        }
    }
    var st = "<table align='center' class='" + className1.replace("_td1","") + "_table1' cellspacing='1' cellpadding='1' id='tc1_" + i + "'><tr>"
    if (bbbb.length == 0) {
        for (var j = 0; j < 7; j++) {
            st += "<td class='" + className1 + "'>&nbsp;</td>";
        }
        st += "</tr><tr>";
        for (var j = 0; j < 7; j++) {
            st += "<td class='" + className1 + "'>&nbsp;</td>";
        }
        st += "</tr>";
    }
    else {
        for (var j = 0; j < bbbb.length; j++) {
            st += bbbb[j];
            if ((j + 1) % 7 == 0) {
                if (j == bbbb.length - 1) {
                    if (bbbb.length % 7 != 0) {
                        for (var f = 0; f < 7 - (bbbb.length % 7); f++) {
                            st += "<td class='" + className1 + "'>&nbsp;</td>";
                        }
                    }
                    st += "</tr>";
                }
                else {
                    st += "</tr><tr>";
                }
            }
            else {
                if (j == bbbb.length - 1) {
                    if (bbbb.length % 7 != 0) {
                        for (var f = 0; f < 7 - (bbbb.length % 7); f++) {
                            st += "<td class='" + className1 + "'>&nbsp;</td>";
                        }
                    }
                    st += "</tr>";
                }
            }
        }
        if (bbbb.length < 8) {
            st += "<tr>";
            for (var j = 0; j < 7; j++) {
                st += "<td class='" + className1 + "'>&nbsp;</td>";
            }
            st += "</tr>";
        }
    }
    st += "</table>";
    document.getElementById("tc1_" + i).parentNode.innerHTML = st;
}

function addLink() {
    var obj2;
    obj2 = document.getElementById("login2010");
    var zb = GetZB(obj2);
    document.getElementById("hk1532888_menu3").style.left = zb.left;
    document.getElementById("hk1532888_menu3").style.top = zb.top;
    MM_showHideLayers("hk1532888_menu3", "", "show");
}

function hideadd() {
    MM_showHideLayers("hk1532888_menu3", "", "hide");
}

function hidead() {
    var hiad; //24小时内只可以去掉一次广告。
    hiad = getCookie("hiad");
    if (hiad == "1") {
        alert("抱歉，一天只能隐藏一次广告，请谅解！");
        return;
    }
    else {
        SetCookie("hiad", "1");
    }
    for (var i = 0; i < 8; i++) {
        tid = parseInt(document.getElementById("tchr_" + i).tid);
        document.getElementById("ad01_" + i).innerHTML = "";
        document.getElementById("ad02_" + i).innerHTML = "";
        document.getElementById("ad000_" + i).innerHTML = "";
        document.getElementById("ad01_" + i).className = "addiv1";
        document.getElementById("ad02_" + i).className = "addiv1";
        document.getElementById("ad000_" + i).className = "addiv1";
    }
    for (var i = 0; i < 100; i++) {
        if (document.getElementById("ad03_" + i)) {
            document.getElementById("ad03_" + i).innerHTML = "";
            document.getElementById("ad03_" + i).className = "addiv1";
        }
    }
    for (var i = 0; i < 5; i++) {
        document.getElementById("ad00_" + i).innerHTML = "";
        document.getElementById("ad00_" + i).className= "addiv1";
    }
    document.getElementById("f_left").style.display = "none";
    document.getElementById("f_right").style.display = "none";
}

function chkform(id) {
    if (frm.cls.value == '') {
        alert('请选择网站类型');
        frm.cls.focus();
        return false;
    }
    if (frm.name.value == '') {
        alert('网站名称不能为空');
        frm.name.focus();
        return false;
    }
    if (frm.url.value == '') {
        alert('网站地址不能为空');
        frm.url.focus();
        return false;
    }
    return true;
}


function back_fav2010() {
    var oXmlHttp = zXmlHttp.createRequest();
    var url1 = "back_fav2010.aspx?" + Date.parse(new Date());
    oXmlHttp.open("get", url1, true);
    oXmlHttp.onreadystatechange = function () {
        if (oXmlHttp.readyState == 4 && oXmlHttp.status == 200) {
            var responseText1 = oXmlHttp.responseText;
            if (responseText1 == "") {
                parent.showf0();
            }
            else {
                document.getElementById("tdffff").innerHTML = responseText1;
            }
         }
    };
    oXmlHttp.setRequestHeader("If-Modified-Since", "0");
    oXmlHttp.send(null);
}

var vhide4;

function hide4() {
    MM_showHideLayers("hk1532888_menu4", "", "hide");
} 