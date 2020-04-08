/*//////////////////////////////////////////////////////////////////
* Script For DvSpace Program
* Copyright (C) 2001 - 2008 Artworld.cn
*
* For further information visit:
* 		http://Artworld.cn/
* 		http://www.artistsky.net/
*
* Builder: Sunwin
* Created: 2007-5-10
/*//////////////////////////////////////////////////////////////////

var getNameList = function(strName){return document.getElementsByName(strName);}

function spacetoolber(){
	var Divs = ['space_title','space_intro'];
	var d=document.getElementById("spacetool");
	if (!d){return;}
	if (d.style.display=="none"){
		d.style.display="block";
		Drop.init(document.getElementById("spacemain"));
		Drag.addEvent();
		spacecanedit(Divs);
		IsEdit=1;
		ShowEdit(document.getElementById("spacemain"));
	}
}

function showdisplay(id){
	var obj = document.getElementById(id);
	if (obj)
	{
		obj.style.display = (obj.style.display=="none") ? "block" : "none";
	}
}

function spacecanedit(Divs){
	var obj;
	for(var i=0;i<Divs.length;i++){
		obj = document.getElementById(Divs[i]);
		if(obj){
			obj.innerHTML = '<textarea class="webNote" onblur="webNote.cannotEdit(\''+Divs[i]+'\')" onclick="webNote.canEdit(event,\''+Divs[i]+'\')">'+obj.innerHTML+'</textarea>';
		};
	};

}

function ShowEdit(obj){
	if (!obj){return;}
	var id;

	var editimg = document.getElementsByTagName("div");

	for(var i=0;i<editimg.length;i++){
		if (editimg[i].id=="moduletitle"){
		id=editimg[i].parentNode.id;
		editimg[i].innerHTML = "<img class=\"cleargif\" src=\""+myspaceskin+"close.gif\" alt=\"\" onclick=\"Drag.del(this,'"+id+"')\"/>"+editimg[i].innerHTML;
		}
	}
}

//显示开关
function fold(e,obj){
	//var e;
	e=fixE(e);
	if(e)element=fixElement(e);
	element=element.parentNode.parentNode;
	element.className=element.className.indexOf("hide")>0?"module":"module hide";
	if (obj){
		obj.src = element.className=="module"?""+myspaceskin+"dblock.gif":""+myspaceskin+"dnone.gif"
	}
}

function fixE(e){var e;e=e?e:(window.event?window.event:null);return e}

function fixElement(e){var e;return e.target?(e.target.nodeType==3?e.target.parentNode:e.target):e.srcElement;}

function getskinsdb(objid,objsrc){
	var obj = document.getElementById(objid);
	if (obj) obj.src = "dv_plus/myspace/script/"+objsrc;
}

function getmodules(objid,objsrc){
	var obj = document.getElementById(objid);
	if (obj) obj.src = "http://channels.dvbbs.net/modules.asp?to="+escape(objsrc);
}

function changskin(styleid,stylefile){
	var styleobj= document.getElementById("stylecss");
	if (styleobj&&stylefile!=""){
		styleobj.href = "skins/myspace/"+stylefile+"style.css";
		document.getElementById("styleid").value=styleid;
		document.getElementById("stylepath").value=stylefile;
	}
}

//XML对象
function GetXMLDOM() {
  var xmlDoc=null;
	try{
		xmlDoc = new ActiveXObject("MSXML2.FreeThreadedDOMDocument");
	} catch(e) {
		try
		{
			 xmlDoc = document.implementation.createDocument("","",null);
			 if (xmlDoc.readyState == null) {
				xmlDoc.readyState = 1;
				xmlDoc.addEventListener("load", function () {
					xmlDoc.readyState = 4;
					if (typeof xmlDoc.onreadystatechange == "function")
						xmlDoc.onreadystatechange();
				}, false);
			}
		} catch(ex) {
			 var xmlDoc=null;
		}
	
	}
  return xmlDoc; 
}


//---RSS 加载--------------------------
//Time 加载时间隔
//例子：
//RssList.url('mod_25','http://news.chinabyte.com/index.xml');
//RssList.load;

var RssList = {
	time : 1000,
	clear : new Array(),
	urls : new Array(),
	times : new Array(),
	XslDoc : function(){},
	XslDom : function(){},
	url : function(id,val){
		if (typeof(RssList.urls[id])=='undefined'){
			RssList.urls[id] = val;
			RssList.clear[id] = null;
			RssList.times[id] = 1;
		}
	},
	
	load : function(){
		if (typeof(RssList.XslDom["RssXslt"])=="undefined"){
			RssList.AjaxGet("RssXslt","dv_plus/myspace/script/rss.xslt","get")
		}
		for (var id in RssList.urls){
				RssList.clear[id] = window.setInterval("RssList.readRSS('"+id+"','"+RssList.urls[id]+"')",RssList.time);
				RssList.time +=RssList.time;
		}
	},
	
	AjaxGet : function(id,url,ptype){
		var xmlhttp = new HttpObj();
		xmlhttp.onreadystatechange = function(){
			if(xmlhttp.readyState==4){
				if(xmlhttp.status==200){
					RssList.XslDom[id] = xmlhttp.responseXML;
				}else{

				}
			}
		}
		xmlhttp.open(ptype,url);	
		xmlhttp.send(null);
		return RssList.XslDom[id];
	},

	readRSS : function(objid,url){
		if (url==''){return;}
		var obj = document.getElementById(objid);
		if (!obj){return;}
		var adddiv = obj.getElementsByTagName("div");
		if (adddiv.length>1){
			adddiv = adddiv.item(2);
		}
		if (adddiv.className!='cont'){
			return;
		}
		//var XmlDom = GetXMLDOM();
		//XmlDom.async=false;
		if (RssList.times[objid] >2)
		{
			window.clearInterval(RssList.clear[objid]);
			adddiv.innerHTML = "RssLoad Is Null;"
			return;
		}else{
			RssList.times[objid] += 1;
		}
		url = url.replace(/&amp;/ig,'&')
		url = "dv_rssRead.asp?url="+escape(url);
		RssList.AjaxGet(objid,url,"get");
		if (RssList.XslDom[objid]){
			//var Xsl = GetXMLDOM();
			//Xsl.async = false;
			//Xsl.load("dv_plus/myspace/script/rss.xslt")
			if(document.implementation && document.implementation.createDocument){
				// 定义XSLTProcessor对象 
			   var xsltProcessor = new XSLTProcessor();
			   xsltProcessor.importStylesheet(RssList.XslDom["RssXslt"]);
				// transformToDocument方式
				var result = xsltProcessor.transformToDocument(RssList.XslDom[objid]);
				var xmls = new XMLSerializer(); 
				adddiv.innerHTML = xmls.serializeToString(result);
			}
			else{
				//alert(XmlDom.transformNode(Xsl));
				adddiv.innerHTML = RssList.XslDom[objid].transformNode(RssList.XslDom["RssXslt"]);
			}
			if (RssList.clear[objid]){window.clearInterval(RssList.clear[objid]);}
		}
		
	}
}



//显示开关
function rssnews(e,obj){
	//var e;
	e=fixE(e);
	if(e)element=fixElement(e);
	element=element.parentNode.parentNode;
	//element = element.childNodes(1);
	element = element.getElementsByTagName("div");
	if (element.length>1){
		element = element.item(1);
		element.className=element.className.indexOf("hide")>0?"itemcontent":"itemcontent hide";
		if (obj){
			obj.src = element.className=="itemcontent"?"images/dblock.gif":"images/dnone.gif"
		}
	}
}

//新闻模式分类显示菜单
function NewsSpanBar(){
	this.f=1;
	this.event = "click"
	this.titleid = "";
	this.bodyid="";
	var Tags,TagsCnt,len,flag;
	this.load=function(){
		flag = this.f;
		Tags=document.getElementById(this.titleid).getElementsByTagName('p'); 
		TagsCnt=document.getElementById(this.bodyid).getElementsByTagName('span'); 
		len=Tags.length;
		for(var i=0;i<len;i++){
			Tags[i].value = i;
			if (this.event!='click'){
				Tags[i].onmouseover=function(){changeNav(this.value)};
			}else{
				Tags[i].onclick=function(){changeNav(this.value)};
			}
			TagsCnt[i].className='undis';
		}
		Tags[flag].className='tophit1';
		TagsCnt[flag].className='dis';
	}
	function changeNav(v){
		Tags[flag].className='tophit0';
		TagsCnt[flag].className='undis';
		flag=v;
		Tags[v].className='tophit1';
		TagsCnt[v].className='dis';
	}
}