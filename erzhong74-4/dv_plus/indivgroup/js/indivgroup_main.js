function ShowAppForm(showflag){
	var appformhtml
	var appgroupformobject = document.getElementById('appgroupform')
	if (showflag==0){
		appformhtml = '<form method=\"post\" action="?action=appgroup" name=\"appform\" onsubmit=\"return submitappform();\" target=\"hiddenframe\">\n';
		appformhtml = appformhtml+'<div style=\"float:left;\">Ȧ������: </div><div><input type=\"text\" name=\"groupname\" size=\"40\" />&nbsp;<span id=\"groupnamestr\"></span></div>\n';
		appformhtml += document.getElementById('groupclasslist').innerHTML;
		appformhtml = appformhtml+'<div style=\"float:left;\">Ȧ��˵��: </div><div style=\"padding-top:2px;\"><textarea row=\"6\" cols=\"40\" name=\"groupinfo\"></textarea></div>\n';
		appformhtml = appformhtml+'<div style=\"float:left;\">��Ա����: </div><div style=\"padding-top:2px;\"><input type=\"radio\" class=\"radio\" value=\"0\" name=\"groupsetting\" checked/>���ɼ���	<input type=\"radio\" class=\"radio\" value=\"1\" name=\"groupsetting\"/>��Ҫ���</div>\n';
		appformhtml = appformhtml+'<div style=\"float:left;\">�������: </div><div style=\"padding-top:2px;\"><input type=\"radio\" class=\"radio\" value=\"0\" name=\"viewflag\" checked/>����	<input type=\"radio\" class=\"radio\" value=\"1\" name=\"viewflag\"/>������</div>\n';
		appformhtml = appformhtml+'<div>���������κ��˶������������������ֻ�г�Ա�������</div>\n';
		appformhtml = appformhtml+'<div style=\"padding-left:55px;height:21px;\"><input type=\"submit\" name=\"submit1\" value=\"�ύ\" class=\"input0\"></div>\n';
		
		appformhtml =appformhtml+'</form>\n';
		appgroupformobject.innerHTML = appformhtml;
	}else{
		if (showflag==1){
			appgroupformobject.innerHTML = '<font color=\"#FF0000\">������û�е�¼���������롣</font>��<a href="login.asp">��¼</a>';
		}else{
			appgroupformobject.innerHTML = '<font color=\"#FF0000\">���Բ��������ڵ��û���û������Ȧ��Ȩ�ޡ�</font>';
		}
	}
	if (appgroupformobject.style.display=='none'){
		appgroupformobject.style.display='block';
	}else{
		appgroupformobject.style.display='none';
	}
}

function submitappform(){
	var formobject=document.appform
	if (formobject.groupname.value==''){
		document.getElementById('groupnamestr').innerHTML='<font color=\"#FF0000\">����������дȦ��������</font>'
		formobject.groupname.focus();
	}else{
		if (formobject.groupname.value.length > 50){
			document.getElementById('groupnamestr').innerHTML='<font color=\"#FF0000\">��Ȧ�����Ƴ��Ȳ��ܳ���50���ַ���</font>'
			formobject.groupname.focus();
		}else{
			//document.getElementById("hiddenframe").src="?action=appgroup&groupname="+formobject.groupname.value+"&groupinfo="+formobject.groupinfo.value
			document.getElementById('appgroupform').style.display='none';
			return true;
		}
	}
	return false;
}

function submitappjion(appjionflag,groupid){
	if (appjionflag==1){
		alert("1.����û�е�¼������������롣");
	}else{
		document.getElementById("hiddenframe").src="?action=appjiongroup&groupid="+groupid
	}
}

function information(str){
	alert(str);
}

//��̳��ת�����˵�����,by Lao_Mi
var xsltDoc

function GetIndivGroupBoardXml(XMLFileAdrr,path) {
	xsltDoc = CreateXmlDocument();
	if (!xsltDoc){ throw new Error('Not support!\nplease install a XML parser');return; }
	xsltDoc.async = false;
	
	if (navigator.userAgent.indexOf("MSIE")==-1){
		if (path){
			xsltDoc.load(path+XMLFileAdrr);
		}else{
			xsltDoc.load(XMLFileAdrr);
		}
	}else{
		xsltDoc.load(XMLFileAdrr);
	}
	
	function CreateXmlDocument() {
		if (document.implementation && document.implementation.createDocument) {
			var doc = document.implementation.createDocument("", "", null);
			if (doc.readyState == null) {
				doc.readyState = 1;
				doc.addEventListener("load", function () {
					doc.readyState = 4;
					if (typeof doc.onreadystatechange == "function")
						doc.onreadystatechange();
				}, false);
			}
			return doc;
		}
		else if (window.ActiveXObject) {
			var prefix = ["MSXML3","MSXML2","MSXML","Microsoft"];
			for (var i=0;i<prefix.length;i++) {
				try {
					var doc = new ActiveXObject(prefix[i] + ".DOMDocument");
					//doc.setProperty("SelectionLanguage","XPath");
					if (doc)
					{
						return doc;
					}		
				} catch (e) {}
			}
		}
		return null;
	};
}

function IndivGroupBoardJumpList(groupid){
	if (!xsltDoc){
		GetIndivGroupBoardXml("IndivGroup_index.asp?GroupID="+groupid+"&action=xml");
	}
	if (xsltDoc.parseError){
		if (xsltDoc.parseError.errorCode!=0){
			return "<div class=\"menuitems\">"+xsltDoc.parseError.reason +"</div>";
		}
	}
	var MenuStr="";
	var LoadBoard,depth;
	var nodelist=xsltDoc.documentElement.getElementsByTagName("Board");
	MenuStr+="<div class=\"menuitems\">";
	//alert(nodelist.length)
	if (nodelist!=null){
		for(var i=0;i<nodelist.length;i++){
			LoadBoard = nodelist[i].getAttribute("id");
			boardtype = nodelist[i].getAttribute("boardname");
			var outtext="��";
			MenuStr+="<a href=\"IndivGroup_Index.asp?groupid="+groupid+"&groupboardid="+LoadBoard+"\">"+ outtext + boardtype +"</a><br>";
		}
		MenuStr+="</div>"
		return MenuStr;
	}else{
		return "<div class=\"menuitems\">����</div>";
	}
}

//selected�����б�ѡȡ��()
function IndivGroupBoardJumpListSelect(groupid,boardid,selectname,fristoption,fristvalue,checknopost){
	if (!xsltDoc){
		GetIndivGroupBoardXml("IndivGroup_index.asp?GroupID="+groupid+"&action=xml");
	}
	var sel = 0;
	var sObj = document.getElementById(selectname);
	if (sObj){
		sObj.options[0] =  new Option(fristoption, fristvalue);
		if (xsltDoc.parseError){
			if (xsltDoc.parseError.errorCode!=0) return;
		}
		var nodes = xsltDoc.documentElement.getElementsByTagName("Board")
		if (nodes){
			for (var i = 0,k = 1;i<nodes.length;i++){
				var t = nodes[i].getAttribute("boardname");
				var v = nodes[i].getAttribute("id");
				if (v==boardid) sel=k;
				var outtext="��";
				t = outtext + t
				t = t.replace(/<[^>]*>/g, "")
				t = t.replace(/&[^&]*;/g, "")
				if(checknopost==1 && nodes[i].getAttribute("nopost")=='1') t+="(����ת��)"
				sObj.options[k++] = new Option(t, v);
			}
			sObj.options[sel].selected = true;
		}
	}
}