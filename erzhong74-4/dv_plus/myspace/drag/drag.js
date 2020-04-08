
var Drag={
	draging : false,
	x : 0,
	y : 0,
	element : null,
	fDiv : null,
	ghost : null,
	ix:2,iy:7,
	ox:6,oy:7,
	fx:6,fy:6,

addEvent : function(){
	var a=document.getElementsByTagName("li");
	for(var i=a.length-1;i>-1;i--){
		if(a[i].className=="module"){
			a[i].onmousedown=Drag.dragStart;

		}	
	}
},

del : function(obj,id){
	if(obj){
		obj.parentNode.parentNode.parentNode.removeChild(obj.parentNode.parentNode);
	}
	if (id!="undefined"){
		//var fobj = document.getElementById(id);
		//if (fobj)fobj.parentNode.removeChild(fobj);
		//document.getElementById("add_"+id).checked=false;
	}
},

//fb添加或去掉：true/false
//例如：Drag.add('usertopic','MainChannal',false)
add : function(id,path,fb){
	if (id=="")return;
	var obj = document.getElementById(id);
	if (document.getElementById("layoutset").value=="2"&&path=="LeftChannal"){path="RightChannal";}
	var addpart = document.getElementById(path);
	var addinput = document.getElementById("add_"+id);
	if (obj){
		obj.parentNode.removeChild(obj);
	}
	if(fb){
		var temppannal = document.getElementById("temppannal").innerHTML;
		temppannal = temppannal.replace(/\[\$id\]/g, id);
		temppannal = temppannal.replace(/\[\$pannaltitle\]/g, addinput.nextSibling.innerHTML);
		temppannal = temppannal.replace(/\[\$cont\]/g,"");
		addpart.innerHTML += temppannal;
		addinput.checked=true;
	}else{
		addinput.checked=false;
	}
	Drag.addEvent();
},

addmodule : function(id,path,title,stype,arr){
	if (document.getElementById("layoutset").value=="2"&&path=="LeftChannal"){path="RightChannal";}
	var addpart = document.getElementById(path);
	var obj = document.getElementById(id);
	if (obj){
		obj.parentNode.removeChild(obj);
	}
	var temppannal = document.getElementById("temppannal").innerHTML;
	temppannal = temppannal.replace(/\[\$id\]/g, id);
	temppannal = temppannal.replace(/\[\$pannaltitle\]/g, title);
	if (stype=="rss"){
		arr = "Rss Channal Data!"
		temppannal = temppannal.replace(/\[\$cont\]/g, arr);
	}else{
		temppannal = temppannal.replace(/\[\$cont\]/g, "Html Channal Data");
	}
	//alert(temppannal)
	addpart.innerHTML += temppannal;
	Drag.addEvent();
},


dragStart : function (e){
	if(Drag.draging)return;
	var e;
	e=fixE(e);
	if(e)element=fixElement(e);
	if(element.className!="title")return;
	element=element.parentNode;
	Drag.element=element;
	//以上获得当前移动的模块

	Drag.x=e.layerX?e.layerX+Drag.fx:(IE?e.x+Drag.ix:e.offsetX+Drag.ox);
	Drag.y=e.layerY?e.layerY+Drag.fy:(IE?e.y+Drag.iy:e.offsetY+Drag.oy);
	//鼠标相对于模块的位置
	Drop.measure();
	if(e.layerX){Drag.floatIt(e);Drag.drag(e);}//fix FF
	B.style.cursor="move";
	document.onmousemove=Drag.drag;
	document.ondragstart=function(){window.event.returnValue = false;}
	document.onselectstart=function(){window.event.returnValue = false;};
	document.onselect=function(){return false};
	document.onmouseup=element.onmouseup=Drag.dragEnd;
	element.onmousedown=null;
},

drag : function (e){
	var e;
	e=fixE(e);
	if(!Drag.fDiv)Drag.floatIt(e);//for IE & Opera
	var x=e.clientX,y=e.clientY;
	Drag.fDiv.style.top=y+H.scrollTop-Drag.y+"px";
	Drag.fDiv.style.left=x+H.scrollLeft-Drag.x+"px";
	Drop.drop(x,y);
	//statu(e);
},

dragEnd : function (e){
	B.style.cursor="";
	document.ondragstart=document.onmousemove=document.onselectstart=document.onselect=document.onmouseup=null;
	Drag.element.onmousedown=Drag.dragStart;
	if(!Drag.draging)return;
	Drag.ghost.parentNode.insertBefore(Drag.element,Drag.ghost);
	Drag.ghost.parentNode.removeChild(Drag.ghost);
	B.removeChild(Drag.fDiv);
	Drag.fDiv=null;
	Drag.draging=false;
	Drop.init(document.getElementById("spacemain"));
},

floatIt : function(e){
	var e,element=Drag.element;
	var ghost=document.createElement("LI");
	Drag.ghost=ghost;
	ghost.className="module ghost";
	ghost.style.height=element.offsetHeight-2+"px";
	element.parentNode.insertBefore(ghost,element);
	//创建模块占位框
	var fDiv=document.createElement("UL");
	Drag.fDiv=fDiv;
	fDiv.className="float";
	B.appendChild(fDiv);
	fDiv.style.width=ghost.parentNode.offsetWidth+"px";
	fDiv.appendChild(element);
	//创建容纳模块的浮动层
	Drag.draging=true;
	}
}

var Drop={
	root : null,
	index : null,
	column : null,

init : function(it){
	if(!it)return;
	Drop.root=it;

	it.firstItem=it.lastItem=null;
	var a=it.getElementsByTagName("ul");
	for(var i=0;i<a.length;i++){
	if(a[i].className!="column")continue;
	if(it.firstItem==null){
		it.firstItem=a[i];
		a[i].previousItem=null;
	}else{
		a[i].previousItem=a[i-1];
		a[i-1].nextItem=a[i];
	}
	a[i].nextItem=null;
	it.lastItem=a[i];
	a[i].index=i;

	a[i].firstItem=a[i].lastItem=null;
	var b=a[i].getElementsByTagName("li");
	for(var j=0;j<b.length;j++){
		if(b[j].className.indexOf("module")==-1)continue;
		if(a[i].firstItem==null){
			a[i].firstItem=b[j];
			b[j].previousItem=null;
		}else{
			b[j].previousItem=b[j-1];
			b[j-1].nextItem=b[j];
		}
		b[j].nextItem=null;
		a[i].lastItem=b[j];
		b[j].index=i+","+j;
		}
	}
},

measure : function(){
	if(!Drop.root)return;
	var currentColumn=Drop.root.firstItem;
	while(currentColumn){
		var currentModule=currentColumn.firstItem;
		while(currentModule){
		currentModule.minY=currentModule.offsetTop;
		//currentModule.minY=currentModule.style.top;
		currentModule.maxY=currentModule.minY+currentModule.offsetHeight;
		currentModule=currentModule.nextItem;
		}
		currentColumn.minX=currentColumn.offsetLeft;
		//currentColumn.minX=currentColumn.style.left;
		currentColumn.maxX=currentColumn.minX+currentColumn.offsetWidth;
		currentColumn=currentColumn.nextItem;
	}
	Drop.index=Drag.element.index;
},

drop : function(x,y){
	if(!Drop.root)return;
	var x,y,currentColumn=Drop.root.firstItem;
	while(x>currentColumn.maxX)if(currentColumn.nextItem)currentColumn=currentColumn.nextItem;else break;
	var currentModule=currentColumn.lastItem;
	if(currentModule)while(y<currentModule.maxY){
	if(y>currentModule.minY-12){
	if(Drop.index==currentModule.index)return;
	Drop.index=currentModule.index;
	if(currentModule.index==Drag.element.index){if(currentModule.nextItem)currentModule=currentModule.nextItem;else break;}
	currentColumn.insertBefore(Drag.ghost,currentModule);
	Drop.column=null;
	return;
	}else if(currentModule.previousItem)currentModule=currentModule.previousItem;else return;}
	if(Drop.column==currentColumn.index)return;
	currentColumn.appendChild(Drag.ghost);
	Drop.index=0;
	Drop.column=currentColumn.index;
	}
}

function statu(e){
	var e,element;
	element=fixElement(e);
	window.status= "e.xy:("+e.x+","+e.y+")<br/>e.offsetXY:("+e.offsetX+","+e.offsetY+")<br/>e.clientXY:("+e.clientX+","+e.clientY+")<br/>element.offsetLeftTop:("+element.offsetLeft+","+element.offsetTop+")<br/>e.layerXY:("+e.layerX+","+e.layerY+")";
	//var aa=document.getElementById("aaa");
	//aa.innerHTML="e.xy:("+e.x+","+e.y+")<br/>e.offsetXY:("+e.offsetX+","+e.offsetY+")<br/>e.clientXY:("+e.clientX+","+e.clientY+")<br/>element.offsetLeftTop:("+element.offsetLeft+","+element.offsetTop+")<br/>e.layerXY:("+e.layerX+","+e.layerY+")";
}

//编辑区域
var webNote={
	obj : null,
	canEdit : function(e,id){
		var e,element;
		e=fixE(e);
		element=fixElement(e);
		if(element.className!='webNote')return;
		if (element.type=="textarea"){
			element.style.borderColor='red';
			webNote.obj=element;
			return;
		}
		if(typeof element.contentEditable!="undefined"){
			element.contentEditable=true;
			element.style.borderColor='red';
			element.focus();
			webNote.obj=element;
		}
	},

	cannotEdit : function(id){
		if(!webNote.obj)return;
		if (webNote.obj.type=="textarea"){
			webNote.obj.style.borderColor='#ffffe0';
			if (document.getElementById("val_"+id)){
				document.getElementById("val_"+id).value = webNote.obj.value;
				//document.getElementById(id).innerHTML = document.getElementById("val_"+id).value;
			};
			return;
		}
		if(typeof webNote.obj.contentEditable!="undefined"){
			if (document.getElementById("val_"+id)){
				document.getElementById("val_"+id).value = webNote.obj.innerText;
			}
			
			webNote.obj.style.borderColor='#ffffe0';
			webNote.obj.contentEditable=false;
			webNote.obj=null;
		}
	}


}

function setlayout(val){
	if (IsEdit==0){return}
	document.getElementById("layoutset").value=val;
	var LeftCannal = document.getElementById("LeftChannal");
	var MainCannal = document.getElementById("MainChannal");
	var RightCannal = document.getElementById("RightChannal");
	LeftCannal.style.display = 'block';MainCannal.style.display = 'block';RightCannal.style.display = 'block';
	switch (val)
	{
	case "1":
		RightCannal.style.display = 'none';
		LeftCannal.innerHTML += RightCannal.innerHTML;
		RightCannal.innerHTML = "";
		LeftCannal.style.width = "30%";
		MainCannal.style.width = "69%";
		RightCannal.style.width = "0%";
		break;
	case "2":
		LeftCannal.style.display = 'none';
		RightCannal.innerHTML += LeftCannal.innerHTML;
		LeftCannal.innerHTML="";
		RightCannal.style.width = "30%";
		MainCannal.style.width = "69%";
		LeftCannal.style.width = "0%";
		break;
	case "3":
		LeftCannal.style.width = "21%";
		RightCannal.style.width = "21%";
		MainCannal.style.width = "54%";
		break;
	}
	if (IsEdit==1){Drag.addEvent();}
}

function saveset(){
	layoutleft.value = get_layoutstr("LeftChannal");
	layoutmain.value = get_layoutstr("MainChannal");
	layoutright.value = get_layoutstr("RightChannal");
}
function get_layoutstr(id){
	var obj = document.getElementById(id);
	var str="";
	if (obj){
		var a=obj.getElementsByTagName("li");
		for(var i=0;i<a.length;i++){
			if(a[i].className=="module hide"||a[i].className=="module"){
				str +=a[i].id;
				if (i<a.length-1){str +=","};
			};
		};
	};
	return str;
}

function putlayout(){
	if (layoutleft.value!=""){
		addlayout(layoutleft.value,"LeftChannal");
	};
	if (layoutmain.value!=""){addlayout(layoutmain.value,"MainChannal");}
	if (layoutright.value!=""){addlayout(layoutright.value,"RightChannal");};

}

function addlayout(str,layval){
	var val = str.split(",");
	var addinput;
	for(var i=0;i<val.length;i++){
		addinput = document.getElementById("add_"+val[i]);
		if (addinput){
			addinput.checked=true;
		}
		//Drag.add(val[i],layval,true);
	};
}
//根据系统数据设定布局
function LoadSet(style){
	layoutleft = document.getElementById("layoutleft");
	layoutmain = document.getElementById("layoutmain");
	layoutright = document.getElementById("layoutright");
	putlayout();
	//setlayout(style);
}

function setmenu(obj){
	var menus = obj.nextSibling;
	if (menus.className=="setmenu"){
		if (menus.style.display=="none"){
			menus.style.display="block";
			menus.style.left = getOffsetLeft(obj)+'px';
			menus.style.top = getOffsetTop(obj)+obj.offsetHeight+'px';
		}else{menus.style.display="none"}
	}
}
