<!--
var useAjaxPost = 1 ; // �Ƿ����ajax��ʽ�ύ 1Ϊ��  0Ϊ��
var post_time=10;  //����ʱ ����
var n=post_time;
var post=0; //���ڽ�ֹ�ظ��ύ ���
var quoteMode=0; //1Ϊ����Ƕ������ 0Ϊ��

var j$=function(id) {
	var itm = null;
	if (document.getElementById) {itm = document.getElementById(id);}
	else if (document.all) {itm = document.all[id];}
	else if (document.layers) {itm = document.layers[id];}
	return itm;
}
var getByName=function(_name) {	return document.getElementsByName(_name);}

function reply(id,userName,postTime,oDvEdit){
	if (j$('textstyle_'+id) && oDvEdit)
	{
		var contentDiv = j$('textstyle_'+id);
		var content;
		if (quoteMode==0){
			try	{
				var quoteDiv = document.createElement("div");
				quoteDiv.innerHTML = contentDiv.innerHTML;
				var elem = quoteDiv.getElementsByTagName("div");
				for (var i=0;i<elem.length;i++){
					if (elem[i].className.toLowerCase()=="quote"||-1!=elem[i].className.toLowerCase().indexOf("first_body")){
						quoteDiv.removeChild(elem[i]);
					}
				}
				var getImg = quoteDiv.getElementsByTagName("img");
				for (var i=0;i<getImg.length;i++){
					getImg[i].removeAttribute("onload");
				}
				content = quoteDiv.innerHTML;				
			}
			catch (e){}			
		}
		else{
			var getImg = contentDiv.getElementsByTagName("img");
			for (var i=0;i<getImg.length;i++){
				getImg[i].removeAttribute("onload");
			}
			content = contentDiv.innerHTML;
		}
		content = content.replace(/\<script[\s\S]*?\<\/script\>/gi,"");

		oDvEdit.clear();
		if(2!=Dvbbs_Mode)
			oDvEdit.insert('<div class=quote><b>����������<i>'+userName+'</i>��'+postTime+'�ķ��ԣ�</b><br>'+content+'</div><p></p>');
		else
			oDvEdit.insert('[QUOTE][B]����������[I]'+userName+'[/I]��'+postTime+'�ķ��ԣ�[/B][BR]'+content+'[/QUOTE]');
	}	
}

function readyPost(){
	post=1;
	document.Dvform.submit.disabled=true;
	//document.Dvform.Submit2.disabled=true;
	if (j$("errinfo")) {j$("errinfo").innerHTML = "";}
	if (j$("ajaxMsg_1")){
		j$("ajaxMsg_1").style.display = "";
		j$("ajaxMsg_1").innerHTML = '<img src="images/loading_small.gif" width="16" height="16" alt="Loading..." />&nbsp;������Ϣ��<font color="#FF9900">�����ύ���ݣ����Ժ�...</font>'
	}
}

var disp;
function postSucceed(str){
	if (j$("errinfo")) {j$("errinfo").innerHTML = "";}
	if (j$("ajaxMsg_1")){
		j$("ajaxMsg_1").style.display = "";
		j$("ajaxMsg_1").innerHTML = '<img src="images/note_ok.gif" width="19" height="16" alt="Ok" />&nbsp;������Ϣ��<font color="#33CC00">'+str+'</font>';
		window.setTimeout("closeAjaxDiv('ajaxMsg_1')",3000);
	}
	n=post_time;
	disp=window.setInterval(ShowInfo,1000);
	ajaxReset();
}

function ShowInfo(){
	var strInfo =  "�뻺һ���Ժ��ٹ�...";
	n--;
	if(n <= 0){
		document.Dvform.submit.value = "OK! ����ظ�";
		document.Dvform.submit.disabled = false;
		//document.Dvform.Submit2.disabled=false;
		post=0;
		clearInterval(disp);
	}
	else{
		document.Dvform.submit.value = strInfo + n +"��";
	}   
}   

function ajaxReset(){
	//document.Dvform.reset();
	if (dvtextarea){
		dvtextarea.clear();
	}	
}

function postErr(str){
	if (j$("errinfo")) {j$("errinfo").innerHTML = ""; }
	if (j$("ajaxMsg_1")){
		j$("ajaxMsg_1").style.display = "";
		j$("ajaxMsg_1").innerHTML = '<img src="images/note_error.gif" width="16" height="16" alt="Err" />&nbsp;������Ϣ��<font color="#FF0000">'+str+'</font> ������...';
		window.setTimeout("closeAjaxDiv('ajaxMsg_1')",15000);
	}
	document.Dvform.submit.disabled=false;
	//document.Dvform.Submit2.disabled=false;
	post=0;
}

function closeAjaxDiv(id) {j$(id).style.display = "none";}

function AjaxPost(){
	var base=this;
	this.postForm=function(_form){
		if (useAjaxPost != 1) return true;
		var data="",usernameIndex=null,bodyIndex=null,topicIndex=null;		
		if(_form && post==0){			
			readyPost();
			for(var i=0;i<_form.elements.length;i++){
				if (!_form.elements[i]["name"]) continue;
				if ("username"==_form.elements[i]["name"].toLowerCase()){usernameIndex=i;continue;}
				if ("topic"==_form.elements[i]["name"].toLowerCase()){topicIndex=i;continue;}
				if ("body"==_form.elements[i]["name"].toLowerCase()){bodyIndex=i;continue;}
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
			data+="&ajaxPost=1";
			data+="&username="+base.replace(escape(escape(String(_form.elements[usernameIndex].value))));
			data+="&topic="+base.replace(escape(escape(String(_form.elements[topicIndex].value))));
			data+="&body="+base.replace(escape(escape(String(_form.elements[bodyIndex].value))));
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
				if(xmlhttp.readyState==4){
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
/*
˵��
 ���ӣ�
 var ajaxpost=new AjaxPost() 
  
 ����Ҫ���õı����
  onsubmit="return ajaxpost.postForm(this)"
  �ύ���Զ����ӱ���ajaxpost ֵΪ1���Ա�ʶ��Ajax�ύ
*/

var ajaxpost=new AjaxPost()
ajaxpost.onsuccess=function()
{
	//postErr(this.xmlhttp.responseText) return;
	if (addToList(this.xmlhttp.responseText) == 1){
		if (typeof(reload)!="undefined"){
			if(reload==1){location.reload();}
		}
		return;
	}
	try{
		eval("var msg="+this.xmlhttp.responseText);
	}catch(e){postErr(this.xmlhttp.responseText);return;}
	var Str="";	
	for(var i=0;i<msg.message.length;i++){
		Str += ( Str==""?"": msg.message[i]==""?"":"&nbsp;&nbsp;&nbsp;") + msg.message[i];
	}
	if (msg.Suc == 1){
		postSucceed(Str+"&nbsp;&nbsp;��ҳ������ת...");
		// id: Dvbbs.Board_Setting(17)|boardid|RootID|Star|page|returnurl
		if (msg.id != "")
		{
			var id = msg.id.split("|");
			var returnurl;
			switch (parseInt(id[0]))
			{
				//case 1: returnurl = "index.asp"; break;
				//case 2: returnurl = "index.asp?boardid="+id[1]; break;
				//case 3: returnurl = "dispbbs.asp?boardid="+id[1]+"&id="+id[2]+"&star="+id[3]+"&page=1";break;
				//case 4: returnurl = "dispbbs.asp?boardid="+id[1]+"&id="+id[2]+"&star="+id[3]+"&page="+id[4];break;
				default:returnurl = id[5];
			}
			window.setTimeout("window.location.href='"+returnurl+"'",2000);
		}
		
	}
	else
		postErr(Str);
}
ajaxpost.onfailure=function()
{
	var Str="�������"+this.xmlhttp.statu+"<br />"+this.xmlhttp.responseText;
	postErr(Str);
};

function addToList(data){
	if (data.indexOf("ajax.SystemMsg:Post Suc")!=-1){
		var ajaxInsert = j$("ajaxInsert");
		j$("ajaxInsert").innerHTML += data;
		postSucceed("�ظ����ӳɹ���")
		return 1;
	}
	if (data.indexOf("IsAuditInfo")!=-1){
		eval("var msg="+data);
		postSucceed(msg.IsAuditInfo);
		return 1;
	}
	return 0;
}
//->