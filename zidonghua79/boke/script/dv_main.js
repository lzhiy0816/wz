/*//////////////////////////////////////////////////////////////////
 * Script For Dvbbs Program -- �������
 * Version: 8.0
 * Copyright (C) 2000 - 2005 bbs.dvbbs.net
 *
 * For further information visit:
 * 		http://bbs.dvbbs.net/
 * 		http://www.aspsky.net/
 *
 * Builder: Fssunwin
 * Created: 2005-04-23
*///////////////////////////////////////////////////////////////////


var Dvbbs = new MainObjct();
Dvbbs.menuOffX=0;
Dvbbs.menuOffY=18;

function MainObjct(){
	this.VBObjects = new Array();
	this.LoadStats = false;
	/*��ȡ�������Ϣ*/
	this.BrowserInfo = new function(){
		this.sAgent=navigator.userAgent.toLowerCase();
		this.IsIE=this.sAgent.indexOf("msie")!=-1;
		this.IsGecko=!this.IsIE;
		this.IsNC6=document.getElementById&&!this.IsIE;	//����Firefox
		this.IsNC4=document.layers;
		if (this.IsIE){
			this.MajorVer=navigator.appVersion.match(/MSIE (.)/)[1];
			this.MinorVer=navigator.appVersion.match(/MSIE .\.(.)/)[1];
		}else{
			this.MajorVer=0;this.MinorVer=0;
		};
		this.IsIE55rMore=this.IsIE&&(this.MajorVer>5||this.MinorVer>=5);
		this.IsIE4=this.IsIE&&(this.MajorVer<5);
	};
	/*��ǰĿ¼����*/
	this.LastFolder=function(){
		var FolderName = location.pathname;
		FolderName = FolderName.substring(0,FolderName.lastIndexOf("/"));
		if (FolderName.lastIndexOf("/"))
		{
			FolderName = FolderName.substring(FolderName.lastIndexOf("/")+1,FolderName.length);
		};
		return FolderName.toLowerCase();
	};
	/*��ǰ�ĵ�����*/
	this.PageName=function(){
		var Name = location.pathname;
		Name = Name.substring(Name.lastIndexOf("/")+1,Name.length);
		return Name.toLowerCase();
	};

	this.FindObj = function(n, d){
		var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
		d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
		if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
		for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=this.FindObj(n,d.layers[i].document);
		if(!x && d.getElementById) x=d.getElementById(n); return x;
	};
	/*��ȡ����(���ٷ�����ȡ����)*/
	this.Objects = function(idname, forcefetch){
		if (typeof(this.VBObjects[idname]) == "undefined"||forcefetch)
		{
			this.VBObjects[idname] = this.FindObj(idname);
		};
		return this.VBObjects[idname];
	};
	/*����COOKIE �����ƣ�ֵ�����ʱ�䣩*/
	this.createCookie = function (name,value,days){
		var expires = "";
		if (days) {
			var date = new Date();
			date.setTime(date.getTime()+(days*24*60*60*1000));
			expires = "; expires="+date.toGMTString();
		};
		document.cookie = name+"="+value+expires+"; path=/";
	};
	/*��ȡCOOKIE �����ƣ� ����COOKIEֵ*/
	this.readCookie = function (name){
		var nameEQ = name + "=";
		var ca = document.cookie.split(';');
		for(var i=0;i < ca.length;i++) {
			var c = ca[i];
			while (c.charAt(0)==' ') c = c.substring(1,c.length);
			if (c.indexOf(nameEQ) == 0) return c.substring(nameEQ.length,c.length);
		};
		return null;
	};
	/*
	* ��ѡ��ȫѡ�¼� form������,CHECKALL��,Ŀ¼��ѡ��
	* CheckAll(this.form,this,'SelectName') CheckBoxID
	*/
	this.CheckAll = function (form,CheckAll,CheckBoxID){
		var checkBoxs = form.elements[CheckBoxID];
		if (checkBoxs.length>0)
		{
			for (var i=0;i<checkBoxs.length;i++)
			{
				checkBoxs[i].checked = CheckAll.checked;
			};
		}else{checkBoxs.checked = CheckAll.checked;};
	};

	/*��ȡ����Y����*/
	this.getOffsetTop = function(elm) {
		var mOffsetTop = elm.offsetTop;
		var mOffsetParent = elm.offsetParent;
		while(mOffsetParent){
			mOffsetTop += mOffsetParent.offsetTop;
			mOffsetParent = mOffsetParent.offsetParent;
		};
		return mOffsetTop;
	};

	/*��ȡ����X����*/
	this.getOffsetLeft = function(elm) {
		var mOffsetLeft = elm.offsetLeft;
		var mOffsetParent = elm.offsetParent;
		while(mOffsetParent) {
			mOffsetLeft += mOffsetParent.offsetLeft;
			mOffsetParent = mOffsetParent.offsetParent;
		};
		return mOffsetLeft;
	};
};

function ShowList(ImgObj,ObjName,CloseV,OpenV){
	var DivObj = Dvbbs.Objects(ObjName);
	if (DivObj)
	{
		if (DivObj.style.display=='none')
		{
			ImgObj.src=OpenV;
			DivObj.style.display='';
		}else
		{
			ImgObj.src=CloseV;
			DivObj.style.display='none';
		}
	};
};

function bbimg(o){
	var zoom=parseInt(o.style.zoom, 10)||100;zoom+=event.wheelDelta/12;if (zoom>0) o.style.zoom=zoom+'%';
	return false;
}