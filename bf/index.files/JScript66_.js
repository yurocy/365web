﻿function getCookie(Name) {
    var search = Name + "=";
    if (document.cookie.length > 0) {
        offset = document.cookie.indexOf(search);
        if (offset != -1) {
            offset += search.length;
            end = document.cookie.indexOf(";", offset);
            if (end == -1)
                end = document.cookie.length;
            return unescape(document.cookie.substring(offset, end));
        }
        else {
            return "";
        }
    }
    return "";
}

var flash_div;

flash_div = "100";//  getCookie("fd");
if (flash_div != "2" && flash_div != "0" && flash_div != "1") {
    flash_div = "100";
}

function writeCookie(name, value)
{ var expire = ""; var hours = 365; expire = new Date((new Date()).getTime() + hours * 3600000); expire = ";path=/;expires=" + expire.toGMTString(); document.cookie = name + "=" + value + expire; }

function showflash(x) {
    flash_div = x;
//    writeCookie("fd", x);
    if (x == 100) {
        create_sixdiv();
    }
    else if (x == 1) {
        create_sixdiv(1);
    }
    else if (x == 2) {
        create_sixdiv(1);
    }
}

function show_voi(x) {
    var str = "<table style='font-size:12px;margin:5px;'><tr><td><input type='radio' onclick='javascript:showflash(2)' style='cursor:hand' name='voi_1' id='voi_c2' /><label for='voi_c2' style='cursor:hand'>开启声音播报语音</label></td></tr><tr><td><input type='radio' onclick='javascript:showflash(1)' style='cursor:hand' name='voi_1' id='voi_c1' /><label for='voi_c1' style='cursor:hand'>开启动画播报语音</label></td></tr><tr><td><input type='radio' onclick='javascript:showflash(100)' style='cursor:hand' name='voi_1' id='voi_c3' /><label for='voi_c3' style='cursor:hand'>关闭全部播报语音</label></td></tr></table>";
    str = "<table cellpadding=1 cellspacing=1 bgcolor='#ad7952'><tr><td bgcolor='#ffffff' align='center'><table width='100%' cellpadding=0 cellspacing=0><tr><td style='padding-left:5px;padding-top:5px;background-color:#ffffff'>" + str + "</td></tr></table>";
    show1(str, "cc5_menu1", document.getElementById("iframe1"),x);
}

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

function show1(html1, x, obj,y) {
    document.getElementById(x).innerHTML = html1;
    showcheck();
    var zb = GetZB(obj);
    var zb1 = GetZB(y);
    document.getElementById(x).style.left = zb.left + zb1.left;
    document.getElementById(x).style.top = zb.top + zb1.top + zb1.height-8;
    MM_showHideLayers(x, "", "show");
}

function showcheck() {
    if (flash_div == "1") { document.getElementById("voi_c1").checked = true; }
    else if (flash_div == "2") { document.getElementById("voi_c2").checked = true; }
    else if (flash_div == "100") { document.getElementById("voi_c3").checked = true; }
}

function hide1() {
    MM_showHideLayers("cc5_menu1", "", "hide");
} 