/*���� */
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
function writeCookie(name, value) 
{ 
	var expire = ""; 
	var hours = 365;
	expire = new Date((new Date()).getTime() + hours * 3600000); 
	expire = ";path=/;expires=" + expire.toGMTString(); 
	document.cookie = name + "=" + value + expire; 
}
function openreg()
{
	var turl='reg.asp'
	window.open(turl,"","");
}

function submitlogin()
{
	document.myform.hid.value="hidd"
	document.myform.submit()
}
function chkform(id)
	{
		if(id.cls.value == ''){
			alert('��ѡ����վ����');
			id.cls.focus();
			return false;
		}
		if(id.name.value == ''){
			alert('��վ���Ʋ���Ϊ��');
			id.name.focus();
			return false;
		}
		if(id.url.value == ''){
			alert('��վ��ַ����Ϊ��');
			id.url.focus();
			return false;
		}
		return true;
	}
var d=document;
function checkdata()
{
	  var typeidstr= d.all['typeidstr'].value;
	   if(typeidstr=='')
	   {
	   		 d.all['typeidstr'].focus();
			alert('��ѡ����վ����');
			
			return false;
	   }
	   var webname=d.all['webname'].value;
	   if(webname=='')
	   {
	   		d.all['webname'].focus();
			alert('��������վ����');
			
			return false;
	   }
	    var urlstr=d.all['urlstr'].value;
	   if(urlstr=='')
	   {
	   		d.all['urlstr'].focus();
			alert('��������ַ');
			
			return false;
	   }
	  
	var dlstr=d.all['dlstr'].value;
	alert(escape(typeidstr)+escape(webname)+escape(urlstr)+escape(dlstr))
	var url='urladd.asp?typeidstr='+escape(typeidstr)+'&webname='+escape(webname)+'&urlstr='+escape(urlstr)+'&dlstr='+escape(dlstr)+'';
	
	
	var str = 'typeidstr='+escape(typeidstr)+'&webname='+escape(webname)+'&urlstr='+escape(urlstr)+'&dlstr='+escape(dlstr)+''; 
	var responsevalue=Fun_PostData(url,str);
	if (responsevalue!='')
	{
		
		
		if(responsevalue=='1')
		{
			alert('�����ɹ�,���ǻᾡ�촦��.');
		}
		else if(responsevalue=='2')
		{
			alert('������ύ����վ.');
			
		}
		else
		{
			alert('�Բ���.�������������.');
			
		}
		
	}
}

	function addFavorites(id,name)
	{
		var urlid=getCookie("urlid_cookies");
		if(urlid==null || urlid.indexOf('undefined')!=-1){
			urlid="";
		}
		
		if(urlid.indexOf(''+id+',')==-1)
		{
			urlid=urlid+""+id+",";
			writeCookie("urlid_cookies", urlid); 
			alert('�ղ�'+name+'�ɹ�'); 
			Fun_PostData('favoritesadd.asp?id='+id+'','') 
			Fun_GetPostData();
		}
		else
		{
			alert('����ӹ�����վ');
		} 
		
		
		
	}
	function removeFavorites(id,name)
	{
		var urlid=getCookie("urlid_cookies");
		if(urlid==null) urlid="";
		
		if(urlid!="")
		{
			urlid=urlid.replace(""+id+",","");
			writeCookie("urlid_cookies", urlid); 
			alert('ȡ��'+name+'�ղ�');  
			Fun_PostData('favoritesdel.asp?id='+id+'','') 
			Fun_GetPostData();
		}
		
		
	}
	
	
function Fun_PostData(url,Body)
{
 try{
	 var xml = new ActiveXObject("Microsoft.XMLHTTP");
	
	 xml.open("POST",url,false);
	 
	//alert(url);
	
	 xml.setRequestHeader("Content-Length",unescape(Body).length);
	  xml.setRequestHeader("Content-Type","application/x-www-form-urlencoded");
	
	 xml.setRequestHeader( "encoding", "gb2312 "); 
	 xml.send(Body);
	//alert(xml.responseText);
	 return xml.responseText; 
 }
 catch(e){
		 return '';
 }
}
//�ղ�
function Fun_GetPostData(){
	var urlid=getCookie("urlid_cookies");
	if(urlid==null || urlid.indexOf('undefined')!=-1){
			urlid="";
		}
	var url='favorites.asp?idstr='+urlid+'';
	//alert(url);
	var str = ''; 
	var responsevalue=Fun_PostData(url,str);
if (responsevalue!='')
	{
		
		favorites_body.innerHTML=responsevalue;
		
	}
else
{
   favorites_body.innerHTML="ע���ղؼ��û�,��������Զ�����վ";

}

}
function selectchar_s(char,id,tableid)
{
	
	var turl="list.asp?id="+id+"&char="+char+"";
	var str="id="+id+"&char="+char+"";
	//alert(turl);
	var responsevalue=Fun_PostData(turl,str);
	
	//if (responsevalue!='')
	//{
		document.all(tableid).innerHTML=responsevalue;
		
	//}
	
}
function select_all(id)
{
	if(id==0)
	{
		document.all("alldata_s").style.display = 'none';
		document.all("alldata").style.display = '';
	}
	else
	{
		var turl="list_all.asp?id="+id+"";
		var str="id="+id+"";
	
		var responsevalue=Fun_PostData(turl,str);
		document.all("alldata").style.display = 'none';
		document.all("alldata_s").innerHTML=responsevalue;
		
	
	}
}