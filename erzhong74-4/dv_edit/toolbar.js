//������
var dvedit_toolbar_reply_all=['bold','italic','underline','fontsize','fontfamily','fontcolor','fontbgcolor','separator','emot','link','image','media','insertquote','insertcode','inserttable','separator','cleancode','unlink','separator','justifyleft','justifycenter','justifyright','separator','insertorderedlist','insertunorderedlist','outdent','indent','separator'];
var dvedit_toolbar_reply_ubb=['bold','italic','underline','fontsize','justifycenter','separator','emot','link','image','media','insertquote','insertcode','separator'];
var dvedit_toolbar_newpost_all=['bold','italic','underline','fontsize','fontfamily','fontcolor','fontbgcolor','separator','emot','link','image','media','insertquote','insertcode','inserttable','separator','cleancode','unlink','separator','justifyleft','justifycenter','justifyright','separator','insertorderedlist','insertunorderedlist','outdent','indent','separator'];
var dvedit_toolbar_newpost_ubb=['bold','italic','underline','fontsize','justifycenter','separator','emot','link','image','media','insertquote','insertcode','separator'];
var dvedit_toolbar_msg=['bold','italic','underline','fontsize','fontfamily','fontcolor','fontbgcolor','separator','emot','link','image','media','insertquote','insertcode','separator','cleancode','unlink','separator','justifyleft','justifycenter','justifyright','separator','insertorderedlist','insertunorderedlist','outdent','indent'];
var dvedit_toolbar_config=[];
var a=document.location.href.toLowerCase();
if (-1!=a.indexOf("messanger")){
	dvedit_toolbar_config=dvedit_toolbar_msg;
}else if(-1!=a.indexOf("dispbbs")){
	dvedit_toolbar_config=1==Dvbbs_Mode?dvedit_toolbar_reply_all:dvedit_toolbar_reply_ubb;
}else if(-1!=a.indexOf("post")){
	dvedit_toolbar_config=1==Dvbbs_Mode?dvedit_toolbar_newpost_all:dvedit_toolbar_newpost_ubb;
}
function DvEditToolBar(){
var a=arguments;if (0==a.length){return;};
var this_=this;
this.p=a[0];
this.r=null;
this.w=this.p.$(this.p.x.w);
this.m='dvedit_'+this.p.i+'emot_list_page';
this.icon={
 bold:{t:'�Ӵ�',i:'bold'}
,italic:{t:'б��',i:'italic'}
,underline:{t:'�»���',i:'underline'}
,fontsize:{t:'���ִ�С',i:'fontsize'}
,fontfamily:{t:'����',i:'fontfamily'}
,fontcolor:{t:'������ɫ',i:'fontcolor'}
,fontbgcolor:{t:'���ֱ�����ɫ',i:'fontbgcolor'}
,emot:{t:'�������',i:'emot'}
,link:{t:'��������',i:'link'}
,image:{t:'����ͼƬ',i:'image'}
,media:{t:'���붯��/����/��Ӱ..',i:'media'}
,justifyleft:{t:'�����',i:'justifyleft'}
,justifycenter:{t:'����',i:'justifycenter'}
,justifyright:{t:'�Ҷ���',i:'justifyright'}
,insertorderedlist:{t:'���',i:'insertorderedlist'}
,insertunorderedlist:{t:'��Ŀ����',i:'insertunorderedlist'}
,outdent:{t:'��������',i:'outdent'}
,indent:{t:'��������',i:'indent'}
,tenpay:{t:'����һ��������Ϣ',i:'tenpay'}
,separator:{t:'�ָ���',i:'separator'}
,newline:{t:'����������',i:'newline'}
,inserttable:{t:'������',i:'inserttable'}
,unlink:{t:'ȥ������',i:'unlink'}
,insertquote:{t:'��������',i:'insertquote'}
,insertcode:{t:'�������',i:'insertcode'}
,cleancode:{t:'�������',i:'cleancode'}
};
this.menu={
FontFamily:[
 {t:'����'}
,{t:'����_GB2312'}
,{t:'����'}
,{t:'��Բ'}
,{t:'΢���ź�'}
,{t:'Verdana'}
,{t:'Arial'}
,{t:'Impact'}
,{t:'Georgia'}
,{t:'Courier New'}
,{t:'Times New Roman'}
],
FontSize:[
 {t:'1'}
,{t:'2'}
,{t:'3'}
,{t:'4'}
,{t:'5'}
,{t:'6'}
],
FontColor:[
 ['000000'],['993300'],['333300'],['003300'],['003366'],['000080'],['333399'],['333333']
,['800000'],['FF6600'],['808000'],['008000'],['008080'],['0000FF'],['666699'],['808080']
,['FF0000'],['FF9900'],['99CC00'],['339966'],['33CCCC'],['3366FF'],['800080'],['999999']
,['FF00FF'],['FFCC00'],['FFFF00'],['00FF00'],['00FFFF'],['00CCFF'],['993366'],['C0C0C0']
,['FF99CC'],['FFCC99'],['FFFF99'],['CCFFCC'],['CCFFFF'],['99CCFF'],['CC99FF'],['FFFFFF']
]
};
//t:���� b:��ʼ��� e:������� w:��� h:�߶� nw:һ����ʾ���� nh:��ʾ������ p:�����dv_edit/main.js��·��
this.emot=this.p.em||('undefined'!=typeof(global_emot_config)?global_emot_config:null)||[
 {t:"QQϵ��",b:1,e:49,w:21,h:21,nw:10,nh:5,p:'../images/emot/'}
,{t:"������",b:50,e:115,w:40,h:40,nw:5,nh:3,p:'../images/emot/'}
];
this.emot_default='undefined'!=typeof(global_emot_default)?global_emot_default:0;
this.setIcon=function(o,e){var s=o.style;if (0==e){s.border='1px #f9fbfd solid';}else{var c1=(1==e?'#fff':'#ccc'),c2=(1==e?'#ccc':'#fff');s.borderRight='1px '+c1+' solid';s.borderBottom='1px '+c1+' solid';s.borderTop='1px '+c2+' solid';s.borderLeft='1px '+c2+' solid';}};
this.preLoad=function(s){var o=new Image();o.src=s;return o.src};
this.draw=function(){
	var p=this.p.p,st=this.p.s;
	var a=this.preLoad(p+'skins/'+p+'/images/toolbar.gif');
	var b=this.icon;
	var c=this.p.l;
	var s='';
	c=dvedit_toolbar_config.concat(c);
	for (var i in c){
		var d=this.icon[c[i]];
		if (d){
			switch(d['i']){
				case 'separator': s+='<img class="dvediticodiv" src="'+p+'skins/'+st+'/images/separator.gif" border="0" />';break;
				case 'newline': s+='<br style="clear:both;"/>';break;
				default:
					if ('emot'==d['i']&&0==this.emot.length){continue;};
					s+='<a class="dveditico dveditico'+d['i']+'"><img class="dveditbgborder" style="position:relative;" width="100%" height="100%" src="'+p+'skins/'+st+'/images/blank.gif" title="'+d['t']+'" alt="'+d['t']+'"';
					s+=' onclick="'+this.p.d+'.t.exe(\''+d['i']+'\',this);" ';
					s+=' onmousedown="'+this.p.d+'.t.setIcon(this,1)" ';
					s+=' onmouseover="'+this.p.d+'.t.setIcon(this,2)" ';
					s+=' onmouseout="'+this.p.d+'.t.setIcon(this,0)" />';
					s+='</a>';
			}
		}else{s+=c[i];}
	}
	s+='<br style="clear:both;"/>';
	this.p.$(this.p.x.t).innerHTML=s;
	setTimeout(this.p.d+'.$('+this.p.d+'.x.t).style.display=""',0);
};
this.insert=function(s){if('design'!=this.p.m){this.p.$(this.p.c).value+=s;}else{var f=this.p.f;this.p.fcs(this.p.f);if (f.getSelection){var se=f.getSelection(),ra=se.getRangeAt(0),fg=f.document.createDocumentFragment(),div=f.document.createElement('div');div.innerHTML=s;while(div.firstChild){fg.appendChild(div.firstChild);};se.removeAllRanges();ra.deleteContents();var node=ra.startContainer,pos=ra.startOffset;switch(node.nodeType){case 3:if (fg.nodeType == 3){node.insertData(pos, fg.data);ra.setEnd(node, pos + fg.length);ra.setStart(node, pos + fg.length);}else{node=node.splitText(pos);node.parentNode.insertBefore(fg, node);ra.setEnd(node, pos + fg.length);ra.setStart(node, pos + fg.length);};break;case 1:node=node.childNodes[pos];if (!node){return;}node.parentNode.insertBefore(fg, node);ra.setEnd(node, pos + fg.length);ra.setStart(node, pos + fg.length);break;};se.addRange(ra);}else{var ra=this.r;if(ra){ra.pasteHTML(s);}else{f.document.body.innerHTML+=s;}}this.recvR();}};
this.getSel=function(){
	var f=this.p.f;
	if(f.getSelection){
		return f.getSelection();
	}else{
		return f.document.selection.createRange().htmlText;
	}
}
this.saveR=function(){if(this.p.f.document.selection){this.r=this.p.f.document.selection.createRange();}};
this.recvR=function(){this.p.fcs(this.p.f);if(this.r){this.r.select();this.r=null;}};
this.exec=function(c,p){
	this.recvR();
	if(!this.p.um){
		this.p.f.document.execCommand(c,false,p);
	}else{
		var a=this.getSel()||'';
		var s='';
		switch(c){
			case 'bold':s='[B]'+a+'[/B]';break;
			case 'italic':s='[I]'+a+'[/I]';break;
			case 'underline':s='[U]'+a+'[/U]';break;
			case 'fontsize':s='[SIZE='+p+']'+a+'[/SIZE]';break;
			case 'justifycenter':s='[CENTER]'+a+'[/CENTER]';break;
			case 'insertimage':p=p.replace(/http:\/\/.+?dv_edit\//gi,'');s=/emot\/em/gi.test(p)?'[EM'+p.replace(/.*emot\/em(\d+).gif/gi,'$1')+']':'[IMG]'+p+'[/IMG]';break;
			case 'createlink':s='[URL='+a+']'+p+'[/URL]';break;
		}
		if(s){this.saveR();this.insert(s);}
	};
};
this.open=function(o,t){
	var w=this.w;
	w.innerHTML=t;
	w.style.left=o.offsetLeft+'px';
	w.style.top=(o.offsetTop+o.clientHeight-(this_.p.isFF?window.pageYOffset:0))+'px';
	w.style.display="";
};
this.close=function(){var w=this_.w;w.style.display='none';w.innerHTML='';};
this.imgload=function(o){
	var b=0;
	if (o&&false==o.complete){
		o.src=o.src;
		b=1;
	}
	if (!o||b){
		setTimeout(function(){
			this_.imgload(o);
		},500);
	}
};
this.emotsload=function(b,e){
	for (var i=b; i<=e; ++i){this.imgload(this.p.$('dvedit_'+this.p.i+'_emot_'+i));}
}
this.getEmots=function(t,b){
	var a=this.emot[t];
	var n=a.nw*a.nh;
	var s='';
	for (var i=0; i<=this.emot.length-1; ++i){
		s+='<a href="javascript:;"';
		if (i==t){s+=' style="font-weight:bold" ';}else{s+=' onclick="'+this.p.d+'.t.jumpEmots('+i+','+this.emot[i].b+');" ';};
		s+=' >'+this.emot[i].t+'</a>&nbsp;&nbsp;';
	};
	s += '<table unselectable="on" cellpadding="2" cellspacing="2" style="width:auto;padding:0px;margin:0px;"><tr>';
	if (b>a.e) b=a.e;
	for (var i=b; i<=a.e; ++i){
		s+='<td unselectable="on"><img id="dvedit_'+this.p.i+'_emot_'+i+'" src="'+this.preLoad(this.p.p+a.p+'em'+(i>9?i:'0'+i)+'.gif')+'" ';
		s+=' onmousedown="'+this.p.d+'.t.setIcon(this,1)" ';
		s+=' onmouseover="'+this.p.d+'.t.setIcon(this,2)" ';
		s+=' onmouseout="'+this.p.d+'.t.setIcon(this,0)" ';
		s+=' onclick="'+this.p.d+'.t.exec(\'insertimage\',this.src);"';
		s+='/></td>';
		if (0==(i-b+1)%a.nw) {
			s += '</tr><tr>';
			if (i-b+1>=n) i=a.e+1;
		}
	}
	s+='</tr></table>';
	var pages=Math.ceil((a.e-a.b)/n);
	var ipage=Math.ceil((b-a.b)/n)+1;
	var pre=(b-n<a.b?b:(b-n));
	var nex=(b+n>a.e?b:(b+n));
	s+='<div unselectable="on">'+ipage+'/'+pages+'  <a href="javascript:;" onclick="'+this.p.d+'.t.jumpEmots('+t+','+pre+');"><<</a>  <a href="javascript:;" onclick="'+this.p.d+'.t.jumpEmots('+t+','+nex+');">>></a></div>'
	return s;
};
this.jumpEmots=function(i,b){
	var o=this.p.$(this.m);
	if (!o){return;};
	o.innerHTML="loading...";o.innerHTML=this.getEmots(i,b);
	setTimeout(function(){
		this_.emotsload(b>this_.emot[i].e?this_.emot[i].e:b,this_.emot[i].e);
	},1000);
};
this.getColors=function(c){
	var s='';
	s='<table unselectable="on" cellpadding="2" cellspacing="2" style="width:auto;padding:0px;margin:0px;"><tr>';
	for (var i=this.menu.FontColor.length-1; i>=0; --i){
		s += '<td unselectable="on" style="background-color:#'+this.menu.FontColor[i]+';width:10px;height:15px;" onclick="'+this.p.d+'.t.exeRC(\''+c+'\',\'#'+this.menu.FontColor[i]+'\');"></td>';
		if (0==i%8) s += '</tr><tr>';
	}
	s += '</tr></table>';
	return s;
};
this.ubbMedia=function(){
	var u=this.p.$(this.p.x.w+'_media_url').value.replace(' ',''),w=this.p.$(this.p.x.w+'_media_width').value,h=this.p.$(this.p.x.w+'_media_height').value,auto=this.p.$(this.p.x.w+'_media_autostart').checked;
	var ubb=this.p.$(this.p.x.w+'_media_type').value,x='';
	if (''==ubb) {
		try{x=u.substr(u.lastIndexOf('.')+1,u.length);x=x.toLowerCase();}catch(e){;}
		switch(x){
			case 'swf':ubb='flash';au='';break;
			case 'mp3':
			case 'wmv':
			case 'avi':
			case 'asf':
			case 'mov':ubb='mp';break;
			case 'rm':
			case 'rmvb':
			case 'ram':
			case 'ra':ubb='rm';break;
			default:ubb='mp';alert('δ��ʶ��ý���ʽ������Ĭ�ϵĸ�ʽ���롣');
		};
	}
	u=u.replace(/\s/gi,'%20');
	x='['+ubb+'='+w+','+h+('flash'==ubb?'':','+auto)+']'+u+'[/'+ubb+']';
	return x;
};
this.exeRC=function(c,v){this.exec(c,v);this.close();};
this.exe=function(c,o){
	var s="",i=1,e=this.p.d+'.t.exeRC(\'<1>\','+this.p.d+'.$(\''+this.p.x.w+'<2>\').value);',ds='<div unselectable="on"><input unselectable="on" type="button" value="ȷ��" onclick="<1>" /> <input unselectable="on" type="button" value="ȡ��" onclick="'+this.p.d+'.t.close()" /></div>';
	this.saveR();
	switch(c){
		case 'fontcolor':
			this.open(o,this.getColors("forecolor"));
			break;
		case 'fontbgcolor':
			this.open(o,this.getColors(this.p.isFF?"hilitecolor":"backcolor"));
			break;
		case 'fontfamily':
			for (var it in this.menu.FontFamily){
				s+='<div unselectable="on"><a href="javascript:;" onclick="'+this.p.d+'.t.exeRC(\'fontname\',\''+this.menu.FontFamily[it]['t']+'\');" style="font-family:'+this.menu.FontFamily[it]['t']+'">'+this.menu.FontFamily[it]['t']+'</a></div>';
			}
			this.open(o,s);
			break;
		case 'fontsize':
			for (var it in this.menu.FontSize){
				s+='<div unselectable="on"><a href="javascript:;" onclick="'+this.p.d+'.t.exeRC(\'fontsize\',\''+this.menu.FontSize[it]['t']+'\');"><font size='+this.menu.FontSize[it]['t']+'>'+this.menu.FontSize[it]['t']+'</font></a></div>';
			}
			this.open(o,s);
			break;
		case 'emot':
			s+='<div unselectable="on" id="'+this.m+'" style="padding:0px;">loading...</div>';
			this.open(o,s);
			setTimeout(this.p.d+'.t.jumpEmots('+this.emot_default+','+this.p.d+'.t.emot['+this.emot_default+'].b);',100);
			break;
		case 'link':
			e=e.replace('<1>','createlink');e=e.replace('<2>','_link_url');
			s+='<div unselectable="on">��ѡ��һ�����֣�Ȼ��������ַ��<br/>���磺http://bbs.dvbbs.net/��</div>';
			s+='<div unselectable="on"><input id="'+this.p.x.w+'_link_url" type="text" class="dveditborder" value="http://" style="width:300px;" /></div>';
			s+=ds.replace('<1>',e);
			this.open(o,s);
			break;
		case 'image':
			e=e.replace('<1>','insertimage');e=e.replace('<2>','_image_url');
			s+='<div unselectable="on">������ͼƬ��ַ���������������߿�ı��С��<br/>���磺http://bbs.dvbbs.net/images/logo.gif��</div>';
			s+='<div unselectable="on"><input id="'+this.p.x.w+'_image_url" type="text" class="dveditborder" value="http://" style="width:300px;" /></div>';
			s+=ds.replace('<1>',e);
			this.open(o,s);
			break;
		case 'media':
			e=this.p.d+'.t.recvR();'+this.p.d+'.t.insert('+this.p.d+'.t.ubbMedia());'+this.p.d+'.t.close();';
			s+='<div unselectable="on">������ý���ļ���ַ�������Ƕ��������֡���Ӱ...<br/>��ע���Զ������������û����Ȩ���йء���</div>';
			s+='<div unselectable="on"><input id="'+this.p.x.w+'_media_url" type="text" class="dveditborder" value="http://" style="width:300px;" /></div>';
			s+='<div unselectable="on">��� <input id="'+this.p.x.w+'_media_width" type="text" class="dveditborder" value="500" style="width:50px;" />  �߶� <input id="'+this.p.x.w+'_media_height" type="text" class="dveditborder" value="300" style="width:50px;" />  <input id="'+this.p.x.w+'_media_autostart" type="checkbox" class="dveditborder" checked="checked" />�Զ����� <select id="'+this.p.x.w+'_media_type" class="dveditborder" /><option value="">����</option><option value="flash">flash</option><option value="mp">mp3</option><option value="mp">wmv</option><option value="mp">avi</option><option value="mp">asf</option><option value="mp">mov</option><option value="rm">rm</option><option value="rm">rmvb</option><option value="rm">ram</option><option value="rm">ra</option></select></div>  </div>';
			s+=ds.replace('<1>',e);
			this.open(o,s);
			break;
		case 'tenpay':
			s+='<iframe src="'+this.p.p+'tenpay.html?'+this.p.i+'" width="350" height="200" frameborder="0"></iframe>';
			this.open(o,s);
			break;
		case 'inserttable':
			s+='<div unselectable="on">';
			for (var y=1; y<=5; ++y)
			{
				for (var x=1; x<=10; ++x)
				{
					s+='<div id="'+this.p.x.w+'_table_'+x+'_'+y+'" onmouseover="'+this.p.d+'.t.activetable('+x+','+y+')" onclick="'+this.p.d+'.t.createtable()" style="border:1px solid black;width:16px;height:8px;padding:0;margin:1px;float:left;"></div>';
				}
				s+='<br style="clear:both;"/>';
			}
			s+='<div style="padding:0;margin:1px;">���� <input type="text" size="2" value="0" id="'+this.p.x.w+'_table_rows" />  ���� <input type="text" size="2" value="0" id="'+this.p.x.w+'_table_cols" />  �߿�� <input type="text" size="2" value="1" id="'+this.p.x.w+'_table_border" /></div>';
			s+='<br style="clear:both;height:0px;line-height:0px;padding:0;margin:0px;"/>';
			s+='<div style="padding:0;margin:1px;">��� <input type="text" size="2" value="0" id="'+this.p.x.w+'_table_padding" />  �հ� <input type="text" size="2" value="0" id="'+this.p.x.w+'_table_spacing" />  <input type="button" value="����" onclick="'+this.p.d+'.t.createtable()" /></div>';
			s+='</div>';
			this.open(o,s);
			break;
		case 'insertquote':
			s+='<div unselectable="on">';
			s+='<div unselectable="on">�������������ݣ�</div>';
			s+='<br style="clear:both;height:0px;line-height:0px;padding:0;margin:0px;"/>';
			s+='<div style="width:200px;height:100px;"><textarea id="'+this.p.x.w+'_temp_content" style="width:200px;height:100px;"></textarea></div><div>';
			//s+='<input type="button" value="����" onclick="'+this.p.d+'.insert(\''+(this.p.um?'[quote]':'<div class=quote>�������������ݣ�<br/>')+'\'+'+this.p.d+'.$('+this.p.d+'.x.w+\'_temp_content\').value.replace(/\\n/gi,\''+(this.p.um?'[br]':'<br/>')+'\')+\''+(this.p.um?'[/quote]':'</div><br/>')+'\');'+this.p.d+'.t.close()" /></div>';
			s+='<input type="button" value="����" onclick="'+this.p.d+'.insert(\'[quote]\'+'+this.p.d+'.$('+this.p.d+'.x.w+\'_temp_content\').value.replace(/\\n/gi,\''+(this.p.um?'[br]':'<br/>')+'\')+\'[/quote]\');'+this.p.d+'.t.close()" /></div>';
			s+='</div>';
			this.open(o,s);
			break;
		case 'insertcode':
			s+='<div unselectable="on">';
			s+='<div unselectable="on">������������ݣ�</div>';
			s+='<br style="clear:both;height:0px;line-height:0px;padding:0;margin:0px;"/>';
			s+='<div style="width:200px;height:100px;"><textarea id="'+this.p.x.w+'_temp_content" style="width:200px;height:100px;"></textarea></div><div>';
			//s+='<input type="button" value="����" onclick="var codei=0;'+this.p.d+'.insert(\''+(this.p.um?'[code]':'<div class=htmlcode>�����ǽű����ݣ�<br/>')+'\'+(\'\\n\'+'+this.p.d+'.$('+this.p.d+'.x.w+\'_temp_content\').value).replace(/</gi,\'��\').replace(/>/gi,\'��\').replace(/\\t/gi,\'&nbsp;&nbsp;&nbsp;&nbsp;\').replace(/>/gi,\'��\').replace(/\\n/gi,function(){return \''+(this.p.um?'[br]\'+(++codei)+\' ':'<br/><font color=red><b>\'+(++codei)+\'</b></font>&nbsp;&nbsp;')+' \';})+\''+(this.p.um?'[/code]':'</div>')+'\');'+this.p.d+'.t.close()" /></div>';
			s+='<input type="button" value="����" onclick="var codei=0;'+this.p.d+'.insert(\'[code]\'+(\'\\n\'+'+this.p.d+'.$('+this.p.d+'.x.w+\'_temp_content\').value).replace(/</gi,\'��\').replace(/>/gi,\'��\').replace(/\\t/gi,\'&nbsp;&nbsp;&nbsp;&nbsp;\').replace(/>/gi,\'��\').replace(/\\n/gi,function(){return \'[br]\'+(++codei)+\' \';})+\'[/code]\');'+this.p.d+'.t.close()" /></div>';
			s+='</div>';
			this.open(o,s);
			break;
		case 'cleancode':
			this.p.cleancode();
			break;
		default:
			this.exec(c,false,null);
	}
};
this.activetable=function(ix,iy){
	this.p.$(this.p.x.w+'_table_rows').value=iy;
	this.p.$(this.p.x.w+'_table_cols').value=ix;
	for (var y=1; y<=5; ++y)
	{
		for (var x=1; x<=10; ++x)
		{
			var s=this.p.$(this.p.x.w+'_table_'+x+'_'+y).style;
			if (x<=ix && y<=iy)
			{
				s.borderColor='orange';
				s.backgroundColor='yellow';
			}
			else
			{
				s.borderColor='black';
				s.backgroundColor='white';
			}
		}
	}
};
this.createtable=function(){
	var s='<table border="'+(this.p.$(this.p.x.w+'_table_border').value)+'" cellpadding="'+(this.p.$(this.p.x.w+'_table_padding').value)+'" cellspacing="'+(this.p.$(this.p.x.w+'_table_spacing').value)+'">';
	var ix=parseInt(this.p.$(this.p.x.w+'_table_cols').value);
	var iy=parseInt(this.p.$(this.p.x.w+'_table_rows').value);
	for (var y=1; y<=iy; ++y)
	{
		s+='<tr>';
		for (var x=1; x<=ix; ++x)
		{
			s+='<td>&nbsp;&nbsp;&nbsp;</td>';
		}
		s+='</tr>';
	}
	s+='</table>';
	this.recvR();
	this.insert(s);
	this.close();
};
this.clear=function(c){
	c=c.replace(/<\/?SPAN[^>]*>/gi, "" );
	c=c.replace(/<(\w[^>]*) class=([^ |>]*)([^>]*)/gi, "<$1$3");
	c=c.replace(/<(\w[^>]*) style="([^"]*)"([^>]*)/gi, "<$1$3");
	c=c.replace(/<(\w[^>]*) lang=([^ |>]*)([^>]*)/gi, "<$1$3");
	c=c.replace(/<\\?\?xml[^>]*>/gi, "");
	c=c.replace(/<\/?\w+:[^>]*>/gi, "");
	c=c.replace(/<img+.[^>]*>/gi, "");
	c=c.replace(/<p>&nbsp;<\/p>/gi, "<br/>");
	return c;
};
this.enter=function(e){
	var t=this.p.e,k=e.ctrlKey;
	if(13!=e.keyCode){return;}
	if(k&&1==t||(!k)&&2==t){
		var f=null;t=this.p.$(this.p.c);
		while(t=t.parentNode){
			if(t.tagName&&'form'==t.tagName.toLowerCase()){
				var a=t.elements;
				for(var i=a.length-1;i>=0;--i){
					if (a[i].type&&'submit'==a[i].type.toLowerCase()){
						try{if(e.preventDefault){e.preventDefault();}else{e.returnValue=false;}}catch(er){}
						a[i].click();/*a[i].disabled=true;*/return;
					}
				}
			}
		}
	}
	//if(this.p.isIE){
	//	e.returnValue=false;
	//	this.saveR();
	//	this.insert('design'==this.p.m?'<p></p>':'\n');
	//}
	//if(k){
	//	this.saveR();
	//	this.insert('design'==this.p.m?'vvvvvvvvvvvvvvvvvvvv<br/>':'\n');
	//}
};
this.init=function(){
	this.draw();
	this.p.ae(this.p.f.document,"click",this.close);
	this.p.ae(this.p.f.document,"keydown",function(e){this_.enter(e);});
	this.p.ae(this.p.$(this.p.c),"keydown",function(e){this_.enter(e);});
	//this.p.ae(this.p.f.document.body,"mouseup",function(){this_.saveR();});
	//this.p.ae(this.p.f.document.body,"select",function(){this_.saveR();});
	this.p.ae(this.p.f.document.body,"paste",function(){
		if(this_.p.isIE&&clipboardData&&clipboardData.getData&&clipboardData.getData('text')){
			this_.saveR();
			var a=document.createElement("iframe");
			a.style.visibility="hidden";
			a.style.position="absolute";
			document.body.appendChild(a);
			setTimeout(function(){//ע������Ҫ�첽��������Ȼ���ϵò�������
				var b=a.contentWindow.document.body;
				b.innerHTML="";
				b.createTextRange().execCommand("Paste");
				var c=b.innerHTML;
				document.body.removeChild(a);
				var d=/<\w[^>]* class="?MsoNormal"?/gi;
				if (d.test(c)&&confirm("�������Ǵ�word���ƹ������Ѽ�⵽�����ж���Ĵ��룬��ȷ�����������룬��ȡ����������ճ����")){c=this_.clear(c);}
				this_.insert(c);
			},100);
			return false;
		}
	});
};
this.init();
}