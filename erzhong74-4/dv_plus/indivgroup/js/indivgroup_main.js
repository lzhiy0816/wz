function ShowAppForm(showflag){
	var appformhtml
	var appgroupformobject = document.getElementById('appgroupform')
	if (showflag==0){
		appformhtml = '<form method=\"post\" action="?action=appgroup" name=\"appform\" onsubmit=\"return submitappform();\" target=\"hiddenframe\">\n';
		appformhtml = appformhtml+'<div style=\"float:left;\">圈子名称: </div><div><input type=\"text\" name=\"groupname\" size=\"40\" />&nbsp;<span id=\"groupnamestr\"></span></div>\n';
		appformhtml += document.getElementById('groupclasslist').innerHTML;
		appformhtml = appformhtml+'<div style=\"float:left;\">圈子说明: </div><div style=\"padding-top:2px;\"><textarea row=\"6\" cols=\"40\" name=\"groupinfo\"></textarea></div>\n';
		appformhtml = appformhtml+'<div style=\"float:left;\">成员加入: </div><div style=\"padding-top:2px;\"><input type=\"radio\" class=\"radio\" value=\"0\" name=\"groupsetting\" checked/>自由加入	<input type=\"radio\" class=\"radio\" value=\"1\" name=\"groupsetting\"/>需要审核</div>\n';
		appformhtml = appformhtml+'<div style=\"float:left;\">浏览设置: </div><div style=\"padding-top:2px;\"><input type=\"radio\" class=\"radio\" value=\"0\" name=\"viewflag\" checked/>公开	<input type=\"radio\" class=\"radio\" value=\"1\" name=\"viewflag\"/>不公开</div>\n';
		appformhtml = appformhtml+'<div>（公开：任何人都可以浏览；不公开：只有成员能浏览）</div>\n';
		appformhtml = appformhtml+'<div style=\"padding-left:55px;height:21px;\"><input type=\"submit\" name=\"submit1\" value=\"提交\" class=\"input0\"></div>\n';
		
		appformhtml =appformhtml+'</form>\n';
		appgroupformobject.innerHTML = appformhtml;
	}else{
		if (showflag==1){
			appgroupformobject.innerHTML = '<font color=\"#FF0000\">←您还没有登录，不能申请。</font>请<a href="login.asp">登录</a>';
		}else{
			appgroupformobject.innerHTML = '<font color=\"#FF0000\">←对不起，您所在的用户组没有申请圈子权限。</font>';
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
		document.getElementById('groupnamestr').innerHTML='<font color=\"#FF0000\">←您忘记填写圈子名称了</font>'
		formobject.groupname.focus();
	}else{
		if (formobject.groupname.value.length > 50){
			document.getElementById('groupnamestr').innerHTML='<font color=\"#FF0000\">←圈子名称长度不能超过50个字符。</font>'
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
		alert("1.您还没有登录，不能申请加入。");
	}else{
		document.getElementById("hiddenframe").src="?action=appjiongroup&groupid="+groupid
	}
}

function information(str){
	alert(str);
}

//论坛跳转下拉菜单部分,by Lao_Mi
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
			var outtext="╋";
			MenuStr+="<a href=\"IndivGroup_Index.asp?groupid="+groupid+"&groupboardid="+LoadBoard+"\">"+ outtext + boardtype +"</a><br>";
		}
		MenuStr+="</div>"
		return MenuStr;
	}else{
		return "<div class=\"menuitems\">错误</div>";
	}
}

//selected下拉列表选取表单()
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
				var outtext="╋";
				t = outtext + t
				t = t.replace(/<[^>]*>/g, "")
				t = t.replace(/&[^&]*;/g, "")
				if(checknopost==1 && nodes[i].getAttribute("nopost")=='1') t+="(不许转移)"
				sObj.options[k++] = new Option(t, v);
			}
			sObj.options[sel].selected = true;
		}
	}
}