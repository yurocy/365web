﻿function showcnzz(icon) {

//<script src="http://s14.cnzz.com/stat.php?id=2498650&web_id=2498650" language="JavaScript"></script>
var html1="%3Cscript%20src%3D%22http%3A//s14.cnzz.com/stat.php%3Fid%3D2498650%26web_id%3D2498650%22%20language%3D%22JavaScript%22%3E%3C/script%3E";
//<script src="http://s14.cnzz.com/stat.php?id=2498653&web_id=2498653" language="JavaScript"></script>
var html2="%3Cscript%20src%3D%22http%3A//s14.cnzz.com/stat.php%3Fid%3D2498653%26web_id%3D2498653%22%20language%3D%22JavaScript%22%3E%3C/script%3E";
//<script src="http://s14.cnzz.com/stat.php?id=2498658&web_id=2498658" language="JavaScript"></script>
var html3="%3Cscript%20src%3D%22http%3A//s14.cnzz.com/stat.php%3Fid%3D2498658%26web_id%3D2498658%22%20language%3D%22JavaScript%22%3E%3C/script%3E";
//<script src="http://s14.cnzz.com/stat.php?id=2498661&web_id=2498661" language="JavaScript"></script>
var html4="%3Cscript%20src%3D%22http%3A//s14.cnzz.com/stat.php%3Fid%3D2498661%26web_id%3D2498661%22%20language%3D%22JavaScript%22%3E%3C/script%3E";


    var href__;
    href__ = window.location.href;

    var html0;
    
    if (domain1(href__)){
        html0=html1;
    }
    else if (domain2(href__)){
        html0=html2;
    }
    else if (domain3(href__)){
        html0=html3;
    }
    else{
        html0=html4;
    }
    if (icon == 1) {html0=html0.replace("22%20language","26show%3Dpic%22%20language");}
    document.write(unescape(html0));
}

function domain1(href__)
{
 var isok=0;
 var urls="1532888.hk|2532888.hk|3532888.hk|5532888.hk|6532888.hk|3332888.hk|1532777.com|2532777.com|3532777.com|4532777.com|6532777.com|7532777.com|8532777.com|9532777.com|3332777.com|113.105.169.51|113.105.169.53|113.105.157.98|121.12.122.239|113.105.157.35".split("|");

 for (var i=0;i<urls.length;i++){
     if (href__.indexOf(urls[i]) != -1)
     {
        return true;
     }
 }
 return false;
}

function domain2(href__)
{
 var isok=0;
 var urls="1132777.com|1232777.com|1332777.com|1432777.com|1732777.com|1832777.com|1932777.com".split("|")
 for (var i=0;i<urls.length;i++){
     if (href__.indexOf(urls[i]) != -1)
     {
        return true;
     }
 }
 return false;
}

function domain3(href__)
{
 var isok=0;
 var urls="1332777.com|2332777.com|4332777.com|5332777.com|6332777.com|7332777.com|8332777.com|9332777.com|1632777.com|2632777.com|3632777.com|4632777.com|5632777.com|6632777.com|7632777.com|8632777.com|9632777.com|1532777.net".split("|")
 for (var i=0;i<urls.length;i++){
     if (href__.indexOf(urls[i]) != -1)
     {
        return true;
     }
 }
 return false;
}
 