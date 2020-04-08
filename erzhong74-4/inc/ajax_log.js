var useAjaxPost=1;
var xcount=0;

var j$=function(id)
{
	var itm = null;
	if (document.getElementById) {itm = document.getElementById(id);}
	else if (document.all) {itm = document.all[id];}
	else if (document.layers) {itm = document.layers[id];}
	return itm;
}
var getByName=function(_name) {	return document.getElementsByName(_name);}

function AjaxPost(){
	var base=this;
	this.postForm=function(_form){
		if (useAjaxPost != 1) return true;
		var data="";		
		if(_form){		
			//readyPost();
			//document.getElementById('dvlog').style.display="";
			postErr(0,1);
			for(var i=0;i<_form.elements.length;i++)
			{
				if (!_form.elements[i]["name"]) continue;
				if ("username"==_form.elements[i]["name"].toLowerCase()){usernameIndex=i;continue;}
				if(data!=""&&"&"!=data.substr(data.length-1,1)){data += "&"}
					try{
						if (_form.elements[i].type.toLowerCase() == "radio" || _form.elements[i].type.toLowerCase() =="checkbox"){
							var n = getByName(_form.elements[i].name).length;
							for (var j=0; j<n ; j++ ){
								if (_form.elements[i+j].checked){
									data += _form.elements[i+j].name +"="+ base.replace(escape(String(_form.elements[i+j].value)));
									if (_form.elements[i].type.toLowerCase() == "radio"){break;}
								}
							}
							i = i + n-1;
						}
						else{
							data += _form.elements[i].name+"="+base.replace(escape(String(_form.elements[i].value)));
						}
					}
					catch(e) {data += _form.elements[i].name+"="+base.replace(escape(String(_form.elements[i].value)));}	
			
			}
			data+="&ajaxPost=1&username="+escape(escape(String(_form.elements[usernameIndex].value)));			
			var xmlhttp;
			try{
				xmlhttp= new ActiveXObject("Msxml2.XMLHTTP");
			}catch(e){
				try{
					xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
				}catch(e){
					try{
						xmlhttp= new XMLHttpRequest();
					}catch(e){}
				}
			}
			xmlhttp.onreadystatechange=function(){
				if(xmlhttp.readyState==4)
				{
					base.xmlhttp = xmlhttp;
					if(xmlhttp.status==200){
						//postErr(xmlhttp.responseText);
						base.onsuccess();
					}else{
						base.onfailure();
					}
				}				
			}
			xmlhttp.open(_form.method, _form.action, true);
			xmlhttp.setRequestHeader('Content-type','application/x-www-form-urlencoded');
			xmlhttp.send(data.replace("&&","&"));
		}
		return false;
	}
	this.onsuccess=function(){};
	this.onfailure=function(){};
	this.replace=function(str){
		var con=str;
			con=con.replace(/%A0/gi,"%20");
			con=con.replace(/\+/gi,"%2B");
			return con;
	}

}

var ajaxpost=new AjaxPost()
ajaxpost.onsuccess=function()
{
	var rt=this.xmlhttp.responseText;
	var rl=rt.split("@@@@");
	if (rl.length<0){postErr(rt,0);return;}		
	if (parseInt(rl[1])==1)
	{
		delCookie("count");
		if(rl[0].indexOf("showerr")>0){
			window.location.href="index.asp";
			return;
		}else{
			window.location.href=rl[0];
			return;
		}
	}
	else
	{
		/*if(document.getElementById("imgid")){
			document.getElementById("imgid").src='DV_getcode.asp?t='+Math.random();	
		}*/
		get_Code();
		postErr(rl[0],0);	
		xcount+=1;		
		//setTimeout('document.getElementById("dvlog").style.display="none"',300);	
		//setTimeout('document.getElementById("trh").style.display="none"',300);
		if (xcount>2)
		{
			var cookieEnabled=(navigator.cookieEnabled)? true : false;
			if (cookieEnabled){		
				createCookie("count",4,1);
				window.location.href="login.asp";
			}else{
				window.location.href="login.asp?count=4";
			}		
		}
	}	
}
ajaxpost.onfailure=function()
{
	postErr(this.xmlhttp.responseText,0);
}

function postErr(str,tp)
{
	document.getElementById("trh").style.display="";
	var lg=document.getElementById("tdh")
	if(tp==0)
	{
		lg.innerHTML='<img src="images/note_error.gif" width="16" height="16" alt="Err" />&nbsp;&nbsp;<span style="color:#ff0000;">'+str+'</span>';	
	}
	else
	{
		lg.innerHTML='<img src="images/loading_small.gif" alt="��¼��">��¼��...';
	}
}

function checkLogin(theform)
{
	if (theform.username.value=="")
	{		
		postErr("�����������û���",0);
		theform.username.focus();
		return false;
	}
	if (theform.password.value==""/* || theform.password.value.length<6*/)
	{
		postErr("��������������",0);
		theform.password.focus();
		return false;
	}
	return ajaxpost.postForm(theform);
	/*return true;*/
}
/*create cookie*/
function createCookie(name,value,days)
{
	var expires = "";
	if (days)
	{
		var date = new Date();
		date.setTime(date.getTime()+(days*24*60*60*1000));
		var expires = "; expires="+date.toGMTString();
	}
	document.cookie = name+"="+value+expires+"; path=/";
}
/*read cookies*/
function readCookie(name)
{
	var nameEQ = name + "=";
	var ca = document.cookie.split(';');
	for(var i=0;i < ca.length;i++)
	{
		var c =ca[i].replace( /^\s*/,"");
		if (c.indexOf(nameEQ) == 0)
		{
			return c.substring(nameEQ.length,c.length);
		}
	}
	return null;
}
/*delete cookies*/
function delCookie(name)
{
	var exp = new Date();
	exp.setTime(exp.getTime() - 1);
	var cval=readCookie(name);
	if(cval!=null) document.cookie= name + "="+cval+";expires="+exp.toGMTString();
}
/*
var cookieEnabled=(navigator.cookieEnabled)? true : false
//�ж�cookie�Ƿ���
//������������ie4+��ns6+
if (typeof navigator.cookieEnabled=="undefined" && !cookieEnabled)
{ 
	document.cookie="testcookie"
	cookieEnabled=(document.cookie=="testcookie")? true : false
	document.cookie="" //erase dummy value
}
//if (cookieEnabled) //if cookies are enabled on client's browser
//do whatever
*/