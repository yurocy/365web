﻿function lqdellmove() {
    if (!document.ns) {
        sixdiv.style.left = parseInt((document.documentElement.clientWidth - parseInt(sixdiv.style.width)) / 2);
        sixdiv.style.top = parseInt((document.documentElement.clientHeight - parseInt(sixdiv.style.height)) / 2) + document.documentElement.scrollTop;
    }
    setTimeout("lqdellmove();", 200);
}

function create_sixdiv(x) {
    myxml____ = "";
    OldXmlDoc = "";
    if (flash_div == "1") {
        if (x == 1) {
            if (livestate == 5) { livestate = 104 }
        }
        document.getElementById("sixdiv").style.width = "845px";
        document.getElementById("sixdiv").style.height = "383px";
        document.getElementById("sixdiv").innerHTML = "";
        //        document.write("<div style='width:845px;height:383px;display:none;text-align:center;position: absolute;visibility: visible;z-index:1;line-height:100%;' id='sixdiv'></div>");
        lqdellmove();
    }
    else if (flash_div == "2") {
        if (livestate != 5 || x == 1) {
            document.getElementById("sixdiv").style.width = "1px";
            document.getElementById("sixdiv").style.height = "1px";
            document.getElementById("sixdiv").innerHTML = "";
            //        document.write("<div style='width:1px;height:1px;overflow-x:hidden;overflow-y:hidden;display:none;text-align:center;position: absolute;visibility: visible;z-index:1;line-height:100%;' id='sixdiv'></div>");
        }
    }
    else if (flash_div == "100") {
        document.getElementById("sixdiv").style.width = "1px";
        document.getElementById("sixdiv").style.height = "1px";
        document.getElementById("sixdiv").innerHTML = "";
    }
}

create_sixdiv();

var InternetExplorer = navigator.appName.indexOf("Microsoft") != -1;
var XmlDocVersion = new Array("MSXML4.DOMDocument", "MSXML3.DOMDocument", "MSXML2.DOMDocument", "MSXML.DOMDocument", "Microsoft.XmlDom");
var XmlDoc = null; 		//Dom
var upXmlDoc = null;
for (var i = 0; i < XmlDocVersion.length; i++) {
    try {
        XmlDoc = new ActiveXObject(XmlDocVersion[i]);
        upXmlDoc = new ActiveXObject(XmlDocVersion[i]);
        break;
    }
    catch (e) {
        XmlDoc = null;
        upXmlDoc = null;
    }
}

if (XmlDoc == null) {
    alert('你的浏览器不支持XML组件，请安装XML组件或使用IE浏览器!');
}

function $(ID) {
    return document.getElementById(ID);
}

//加载更新
var OldXmlDoc;
var myxml____ = "";
var my_SataeArr = "";
function reLoad() {
    try {
        if (document.getElementById("sixdiv").style.display == "") {
            var resetit = document.getElementById("tt005").GetVariable("resetit");
            if (resetit == "true") { resetitflash() }
        }
    }
    catch (e) { }
    try {
        if (document.getElementById("sixdiv").style.display == "") {
            var closeit = document.getElementById("tt005").GetVariable("closeit");
            if (closeit == "true") { document.getElementById("sixdiv").style.display = "none"; }
        }
    }
    catch (e) { }
    if (livenow == 3) {
        upXmlDoc.load("liveNow.xml?" + Math.floor(Math.random() * 1000000));
    }
    else {
        upXmlDoc.load("liveNow.aspx?" + Math.floor(Math.random() * 1000000));
    }
    upXmlDoc.onreadystatechange = ActionUpdate;
    window.setTimeout("reLoad()", 1000);
}

var resetitflash1 = 0;

function resetitflash() {
    resetitflash1 = 1;
    createstst();
    myxml____ = "";
    OldXmlDoc = "";
}


function ActionUpdate() {
    if (flash_div != "2" && flash_div != "0" && flash_div != "1") return false;

    if (upXmlDoc.readyState != 4) return false;
    if (upXmlDoc.xml == '') return false;
    if (OldXmlDoc == upXmlDoc.xml) return false;
    var root = upXmlDoc.selectNodes("/root/R");
    var Data = root[0].getAttribute("M");
    if (myxml____ == Data) return false;
    myxml____ = Data;

    if (Data != '') {
        var M = Data.split("^");
        var SataeArr;
        SataeArr = Number(M[11]);
        if (flash_div == "1") {
            if (SataeArr == 5 && livestate != 104) {//不显示
                document.getElementById("sixdiv").style.display = "none";
                OldXmlDoc = upXmlDoc.xml;
                return false;
            }
        }

        if (document.getElementById("tt005")) { }
        else {
            createstst();
        }
        if (my_SataeArr != SataeArr) {
            my_SataeArr = SataeArr;
            if (SataeArr == 0) {//准备开奖
                createstst();
            }
        }
        if (document.getElementById("sixdiv").style.display == "none") {
            document.getElementById("sixdiv").style.display = "";
            myxml____ = "";
            return
        }
        try {
            if (flash_div != "1") {
                document.getElementById("tt005").SetVariable("resetbuttonv", "1");
            }
            if (resetitflash1 == 1) {
                resetitflash1 = 0;
                var Data1 = Data + "^1";
                document.getElementById("tt005").SetVariable("_myxml", Data1);
                document.getElementById("tt005").TCallLabel("/", "loadxml");
            }
            else {
                document.getElementById("tt005").SetVariable("_myxml", Data);
                document.getElementById("tt005").TCallLabel("/", "loadxml");
            }
        }
        catch (e) { }
    }
    OldXmlDoc = upXmlDoc.xml;
}

    ////////////////////////////////////////////////////////////////////////////////////////////////////////
var stst = "<object id='tt005' codebase='http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0' height='383' width='845' align='middle' classid='clsid:d27cdb6e-ae6d-11cf-96b8-444553540000' viewastext>";
    stst += "<param name='_cx' value='4498' />";
    stst += "<param name='_cy' value='6879' />";
    stst += "<param name='FlashVars' value='' />";
    stst += "<param name='Movie' value='6/005.swf?id=0' />";
    stst += "<param name='Src' value='6/005.swf?id=0' />";
    stst += "<param name='Play' value='0' />";
    stst += "<param name='Loop' value='-1' />";
    stst += "<param name='Quality' value='High' />";
    stst += "<param name='SAlign' value='' />";
    stst += "<param name='Menu' value='-1' />";
    stst += "<param name='Base' value='' />";
    stst += "<param name='AllowScriptAccess' value='sameDomain' />";
    stst += "<param name='Scale' value='ShowAll' />";
    stst += "<param name='DeviceFont' value='0' />";
    stst += "<param name='EmbedMovie' value='0' />";
    stst += "<param name='wmode' value='transparent'>";
    stst += "<param name='SWRemote' value='' />";
    stst += "<param name='MovieData' value='' />";
    stst += "<param name='SeamlessTabbing' value='1' />";
    stst += "<param name='Profile' value='0' />";
    stst += "<param name='ProfileAddress' value='' />";
    stst += "<param name='ProfilePort' value='0' />";
    stst += "<embed src='6/005.swf?id=0' quality='high' bgcolor='#ffffff' width='845' height='383' name='005' align='middle' allowScriptAccess='sameDomain' type='application/x-shockwave-flash' pluginspage='http://www.macromedia.com/go/getflashplayer' />";
    stst += "</object>";
    function createstst() {
        if (document.getElementById("sixdiv")) { }
        else {
            create_sixdiv();
        }
        document.getElementById("sixdiv").innerHTML = stst;
    }
    reLoad();
    function kkkkksssss() {
        alert(resetitflash1);
    }
 