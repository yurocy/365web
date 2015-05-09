var _ALL = document.all;
function $(){ return document.getElementById(arguments[0]);}
var Lbase = {};
Lbase.version = function(name){ return window.navigator.userAgent.indexOf(name || 'MSIE 6') != -1; };
Lbase.versionIe67 = function(){ return _ALL && ( Lbase.version() || Lbase.version('MSIE 7') ); };
var Lajax = { _initId : null };
Lajax._arrXmlHttp = ["MSXML2.XmlHttp.6.0","MSXML2.XMLHttp.5.0","MSXML2.XMLHttp.4.0","MSXML2.XmlHttp.3.0","MSXMLL2.XMLHTTP","Microsoft.XMLHTTP"];
Lajax.change = function(o){
    if (this.readyState == 4 && this.status == 200) {this._call(this.responseText);}
};
Lajax.initialize = function(call,par){
    var ajax = null;
    switch(Lajax._initId){
        case -100:
            break;
        case -1:
            ajax = new ActiveXObject("Microsoft.XMLHTTP"); 
            break;
        case -2:
            ajax = new window.XMLHttpRequest(); 
            break;
        case null:
            var id = null;
            if(window.ActiveXObject){ 
                ajax = new ActiveXObject("Microsoft.XMLHTTP"); 
                id = -1;
            }else if(window.XMLHttpRequest){ 
                ajax = new window.XMLHttpRequest(); 
                id = -2;
            }else{
                id = -100;
                var arr = Lajax._arrXmlHttp;
                for(var i = 0 , l = arr.length; i < l; i++){
	                try {ajax = new ActiveXObject(arr[i]); id = i; break;} catch(e) {}
	            }
            }
            Lajax._initId = id;
            break;
        default :
            ajax = new ActiveXObject(Lajax._arrXmlHttp[Lajax._initId]);
            break;
        
    }
    if(ajax && call){
        ajax.onreadystatechange = function(){
            if (ajax.readyState == 4 && ajax.status == 200) { call(ajax.responseText,par);}
        };
    }
    return ajax;
};
Lajax.formatUri = function(uri){ return uri + (uri.indexOf('?') == -1 ? '?' : '&') + 'sendmethod=ajax'; };
Lajax.send = function(call,uri,par){
	var ajax = Lajax.initialize(call,par);
	if(!ajax){ alert('您的浏览器版本太低，请升级！'); return false;}
	ajax.open('GET', uri, true);
	ajax.setRequestHeader("If-Modified-Since","0");
	ajax.send(null);
	return true;
};
var Lxml = {};
Lxml.toXml = function(txt){
    if(_ALL){
        var doc = new ActiveXObject("MSXML2.DOMDocument");
        doc.loadXML(txt);
        return doc;
    }else{
	    return (new DOMParser()).parseFromString(txt, "text/xml");
	}
};
/* Lpositionr */
var Lposition = {};
Lposition.offset = function(o,o2,div,topAlign){
    var oo = o;
    var x = y = x2 = y2 = 0; 
    do{ x += o.offsetLeft || 0;y += o.offsetTop  || 0;o  = o.offsetParent;}while(o);
    if(o2){
        var w = o2.clientWidth;
        var h = o2.clientHeight;
        var dw = document.body.clientWidth;
        x2 = ((x + w > dw) ? dw - w : x);
        if(topAlign){
            y2 = (y - h < 1 ) ? 0 : y - h;
        }else{
            y2 = y + oo.clientHeight;
        }
    }else{
        x2 = x;
        y2 = y;
    }
    if(div){
        div.style.left = x2 + 'px';
        div.style.top = y2 + 'px';
    }
    return {"x" : x, "y" : y , 'x2' : x2, 'y2' : y2};
};

var Ltool = {};
Ltool.create = function(name,type,parentNode,par){
    if(!name){
        name = 'input';
        type = 'button';
    }
    var con = document.createElement(name)
    if(type){ con.type = type;}
    if(par){
        for(var key in par){
            con.setAttribute( key , par[key] );
        }
    }
    if(parentNode){
        parentNode.appendChild(con);
    }
    return con;
};

//注意COOKIE
//_domain : '5532888.com',
var Lcookie={
    _domain : '',
    set:function(name,value,day,hour,minute,path){
        var expires=new Date();
        if(typeof(day)=="number"){expires.setDate(expires.getDate()+day);}
        if(typeof(hour)=="number"){expires.setHours(expires.getHours()+hour);}
        if(typeof(minute)=="number"){expires.setMinutes(expires.getMinutes()+minute);}
        var txt = name+"="+escape(value)+";"+(day||hour||minute? " expires="+expires.toGMTString():"") 
        if(Lcookie._domain){ txt += ';domain=' + Lcookie._domain ; }
        txt += (path==null) ? "; path=/" : "; path="+path;;
        document.cookie = txt;
    },
    get:function(name,dv){var arr=document.cookie.split("; ");for(var i=0; i<arr.length; i++){var kv=arr[i].split("=");if(name==kv[0]){return unescape(kv[1]);}}return dv||'';},
    remove:function(name){
        var txt = name+"=; expires=Fri, 31 Dec 1000 23:59:59 GMT";
        if(Lcookie._domain){ txt += ';domain=' + Lcookie._domain ;} 
        document.cookie= txt ;
    }
};

/* Lanimation */
var Lanimation = {};
Lanimation.tween = function(t,b,c,d){ return -c * ((t = t/d - 1) * t * t * t - 1) + b; };
Lanimation.append = function(o,direction,run,endcall,duration,delay){
    var me = o._Lanimation = {};
    direction = direction || 'left';
    o.style.position = 'absolute';
    me._direction = direction; //方向
    me._duration = duration || 15; //总时间
    me._delay = delay || 20; //延时
    me._topAlign = me._direction == 'top' || me._direction == 'height';
    me._endcall = endcall;
    if(!o._lScrollTop){ o._lScrollTop = me._direction == 'top' ? o.offsetTop : o.clientHeight;}
    if(!o._lScrollLeft){ o._lScrollLeft = me._direction == 'left' ? o.offsetLeft : o.clientWidth;}
    if(run){ Lanimation.run(o); }
};
Lanimation.opacity = function(o,begin,end,run,endcall,duration,delay){
    var me = o._Lanimation = {};
    me._direction = 'opacity'; //方向
    me._duration = duration || 15; //总时间
    me._delay = delay || 20; //延时
    me._endcall = endcall;
    me._end = end;
    //me._Run = o._Run;
    var numto = begin > 0 ? -begin + end : end
    if(run){ Lanimation.run(o,begin,numto,end); }
};
Lanimation.run = function(o,numfrom, numto,target){
    var me = o._Lanimation;
    if(!me){ return;}
    if(numfrom == undefined || numto == undefined){
        if(me._topAlign){
            numfrom = numfrom || (me._direction == 'top' ? o.offsetTop : o.clientHeight);
            numto = numfrom > 0 ? 0 : me._hTop || o._lScrollTop;
            if(numfrom){me._hTop = numfrom;}
        }else{
            numfrom = numfrom || (me._direction == 'left' ? o.offsetLeft : o.clientWidth);
            numto = numfrom > 0 ? 0 : me._hLeft ||  o._lScrollLeft;
            if(numfrom){me._hLeft = numfrom;}
        }
    }
	me._t = 0;
	me._b = numfrom;
	me._c = numto || -numfrom;
	me._target = target || numto;
	me._timer = null;
	if(me._direction == 'opacity'){
	    Lanimation.transparent(o);
	}else{
	    Lanimation.move(o);
	}
};
Lanimation.transparent = function(o){
    if(o._Stop){ return;}
    var me = o._Lanimation;
	var x = me._target;
	clearTimeout(me._timer);
	if (me._t < me._duration) {
		x = Math.round(Lanimation.tween(me._t++, me._b, me._c, me._duration));
		if(x != me._target){
		    me._timer = setTimeout(function(){ Lanimation.transparent(o);}, me._delay);
		}
	}
	if(x < 0){ x = 0;}
	if(x == me._end){
	    clearTimeout(me._timer);
	    if(me._endcall){ me._endcall();}
	}
	if(_ALL){
	    o.style.filter = 'alpha(opacity=' + x + ')';
	}else{
	    o.style.opacity = x * 0.01;
	}
};
Lanimation.move = function(o){
    if(o._Stop){ return;}
    var me = o._Lanimation;
	var left = me._target;
	clearTimeout(me._timer);
	if (me._t < me._duration) {
		left = Math.round(Lanimation.tween(me._t++, me._b, me._c, me._duration));
		if(left != me._target){
		    me._timer = setTimeout(function(){ Lanimation.move(o);}, me._delay);
		}
	}
	o.style[me._direction] = left + "px";
	if(left == me._end){
	    if(me._endcall){ me._endcall();}
	}
};

/* Lad */
var Lad = {_arrHide : []};
Lad.init = function(o){
    o.append = Lad.append;
    o.load = Lad.load;
    o.hide = Lad.hide;
    o.nullId = Lad.nullId;
};
Lad.append = function(id,type,uri,html,style){
    var arr = this[id];
    if(!arr){this[id] = arr = [];}
    arr.push({'uri': uri, 'html': html, 'type': type, 'style': style});
};
Lad.nullId = function(id){
	Lad._arrHide.push(id);
};

Lad.hide = function () {
    var hiad; //24小时内只可以去掉一次广告。
    hiad = getCookie("hiad");
    if (hiad == "1") {
        alert("抱歉，一天只能隐藏一次广告，请谅解！");
        return;
    }
    else {
        SetCookie("hiad", "1");
    }
    var objAd = this;
    for (var id in objAd) {
        if (typeof objAd[id] != 'object') { continue; }
        try {
            document.getElementById(id).style.display = "none";
        }
        catch (e) { }
    }
    for (var j = 0; j < 3; j++) {
        for (var i = 0; i < 8; i++) {
            document.getElementById("ad" + j + "_" + i).style.display = "none";
        }
    }
    for (var j = 0; j < 20; j++) {
        if (document.getElementById("ad100_" + j)) {
            document.getElementById("ad100_" + j).style.display = "none";
        }
    }
    document.getElementById("wAdTopadfadfasdf").innerHTML= "<div style='height:2px'></div>";
    document.getElementById("wAdMyItemdddd").innerHTML = "<div style='height:2px'></div>";
}

Lad.load = function(){
    var objAd = this;
    for(var id in objAd){
        if(typeof objAd[id] != 'object'){ continue;}
        var divId = $(id);
        if(!divId){ continue;}
        divId.innerHTML = '';
        var nodeName = divId.nodeName;
        var arr = objAd[id];
        var index = 0;
        for(var i = 0, l = arr.length; i != l; i++){
            var ad = arr[i];
            var type = '';
            switch(nodeName){
                case 'UL': case 'OL':
                    type = 'li';break;
                case 'TR':
                    type = 'td';break;
                case 'DL':
                    type = 'dd';break;
            }
            var obj = null;
            if(type){
                obj = document.createElement(type);
            }
            var txt = '';
            switch(ad.type){
                case 'txt': case 'img':
                    txt = ad.type == 'txt' ? ad.html : '<img src="' + ad.html + '" />'; 
                    if(obj){
                        var str = '<a href="' + ad.uri + '" target="_blank" ';
                        str += ( ad.style ? 'style="' + ad.style + '"' : '') + '>'+ txt + '</a>';
                        txt = str;
                    }else{
                        obj = document.createElement('a');
                        obj.href = ad.uri;
                        obj.target = '_blank';
                        if(ad.style){ obj.setAttribute('style',ad.style);}
                    }
                    obj.innerHTML = txt;
                    break;
                case 'swf':
                    txt  = '<embed src="'+ ad.uri +'" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" menu="false" wmode="transparent" ' + ( ad.style ? 'style="' + ad.style + '"' : '') + ' ></embed>';
                   if(!obj){ obj = document.createElement('a');}
                   obj.innerHTML = txt;
                   break;
               case 'mytxt':
                   txt = ad.html;
                   if (!obj) { obj = document.createElement('a'); }
                   obj.innerHTML = txt;
                   break;
            }
            if(!obj){ continue;}
            divId.appendChild(obj);
            if(index == 0){ obj.className = 'first';}
            index++;
        }
    }
   //document.title = Lad._arrHide + ',' +Lad._arrHide.length;
    for(var i = 0,l = Lad._arrHide.length; i < l; i++){
	var obj = $(Lad._arrHide[i]);
        //document.title += ',' + $(Lad._arrHide[i]) + '/' + Lad._arrHide[i]
	if(obj){
	obj.style.padding = 0;
	obj.style.lineHeight = '4px';
	obj.style.height = '4px';
	//obj.style.backgroundColor = 'yellow';
	//obj.setAttribute('style','padding:0;height:4px;line-height:4px;background-color:yellow;');
	}
    }
}; 