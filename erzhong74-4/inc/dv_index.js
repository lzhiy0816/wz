function announcements() //�����С�ֱ� ��ʾ
{
	this_=this;
	var rollText_k=(!document.all?document.getElementsByName("announcementsitem").length:document.getElementById("rollTextMenus").childNodes.length);//�˵�����
	if (rollText_k==0)
	{
		var str=window.location.href; 
		var es=/http:\/\/.+boardid\=(\d+)/; 
		es.exec(str); 
		var boardid=RegExp.$1; 
		if (boardid == "")
		{
			boardid = 0;
		}
		document.getElementById("rollTextMenus").innerHTML = '<div id="rollTextMenu1" name="announcementsitem" style="display: none"><strong>�����棺</strong> <a href="announcements.asp?boardid='+boardid+'" target="_blank"><b>��ǰ��δ�й���</b></a></div>';
		rollText_k=1;
	}
	var rollText_i=1; //�˵�Ĭ��ֵ
	document.getElementById("rollTextMenu1").style.display="block";
	document.getElementById("pageShow").innerHTML = "1/"+rollText_k;
	rollText_tt=setInterval(function(){
		this_.rollText(1)
	},4000);
	this.rollText = function(a){
		clearInterval(rollText_tt);
		rollText_tt=setInterval(function(){
			this_.rollText(1)
		},4000);
		rollText_i+=a;
		if (rollText_i>rollText_k){rollText_i=1;}
		if (rollText_i==0){rollText_i=rollText_k;}
		for (var j=1; j<=rollText_k; j++){
			document.getElementById("rollTextMenu"+j).style.display="none";
		}
		document.getElementById("rollTextMenu"+rollText_i).style.display="block";
		document.getElementById("pageShow").innerHTML = rollText_i+"/"+rollText_k;
	}
}
//����ģʽ������ʾ�˵�
function NewsSpanBar(){
	this.f=1;
	this.event = "click"
	this.titleid = "";
	this.bodyid="";
	this.class_dis = "dis";
	this.class_undis = "undis";
	this.class_hiton = "tab_search_on";
	this.class_hitno = "tab_search";

	var Tags,TagsCnt,len,flag;
	var BClassName;
	this.load=function(){
		if (!document.getElementById(this.titleid)||!document.getElementById(this.bodyid))
		{
			return false;
		}
		flag = this.f;
		BClassName = [this.class_dis,this.class_undis,this.class_hiton,this.class_hitno];
		Tags=document.getElementById(this.titleid).getElementsByTagName('p'); 
		TagsCnt=document.getElementById(this.bodyid).getElementsByTagName('dl'); 
		len=Tags.length;
		for(var i=0;i<len;i++){
			Tags[i].value = i;
			if (this.event!='click'){
				Tags[i].onmouseover=function(){changeNav(this.value)};
			}else{
				Tags[i].onclick=function(){changeNav(this.value)};
			}
			TagsCnt[i].className=BClassName[1];
		}
		Tags[flag].className=BClassName[3];
		TagsCnt[flag].className=BClassName[0];
	}
	function changeNav(v){
		Tags[flag].className=BClassName[2];
		TagsCnt[flag].className=BClassName[1];
		flag=v;
		Tags[v].className=BClassName[3];
		TagsCnt[v].className=BClassName[0];
	}
}
/*user_login_register
function tips(e,str)
{
	var l=e.offsetLeft+120;
	var t=e.offsetTop;
	var tip=document.getElementById("tips");
	tip.innerHTML="��ʾ��"+str;
	if(e=e.offsetParent)
	{
		l+=e.offsetLeft;
		t+=e.offsetTop;		
	}
	tip.style.left=l+"px";
	tip.style.top=t+"px";	
	tip.style.display="";
}
function outtips(){
    document.getElementById("tips").style.display='none';
}*/

function loadRightMenu(){
	//�ұ������л�
	var new1 = new NewsSpanBar();
	new1.f=0;
	new1.titleid = "search_bot";
	new1.bodyid = "searchbody";
	new1.class_hiton = "tabgroup_on";
	new1.class_hitno = "tabgroup";
	new1.load();
	//�ұ������б��л�
	var new3 = new NewsSpanBar();
	new3.f=0;
	new3.titleid = "topic_bot";
	new3.bodyid = "topicbody";
	new3.class_hiton = "tabgroup_on";
	new3.class_hitno = "tabgroup";
	new3.load();
	//�ײ��������Ӻ�����ͳ���л�
	var new_link = new NewsSpanBar();
	new_link.f=0;
	new_link.titleid = "bot_link";
	new_link.bodyid = "body_link";
	new_link.class_hiton = "link_on";
	new_link.class_hitno = "link";
	new_link.load();
}

function setTab(name,cursel,n,t){
	for(var i=1;i<=n;i++){	
		var menu=document.getElementById(name+i);
		if(!menu){
			if (t==0)cursel++;
			continue;
		}
		var con=document.getElementById("con_"+name+"_"+i);
		menu.className=(i==cursel?"hover":"");
		con.style.display=(i==cursel?"block":"none");
	}
}