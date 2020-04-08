function agree(boardid,topicid,announceid,t){	
	var postStr="BoardID="+boardid+"&topicid="+topicid+"&announceid="+announceid+"&atype="+t ;
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
			if(xmlhttp.status==200){
				agreeMsg(xmlhttp.responseText);
			}else{
				agreeMsg(xmlhttp.responseText);
			}
		}
	}
	xmlhttp.open("post","Appraise.asp?action=post", true);
	xmlhttp.setRequestHeader('Content-type','application/x-www-form-urlencoded');
	xmlhttp.send(postStr);
	/**/
}

function agreeMsg(str){
	//alert(str);
	var id,s,n;
	id=str.split(",");
	s=document.getElementById("agree_"+id[1]+"_"+id[2]);
	n=document.getElementById("agree_"+id[1]+"_"+id[2]+"_n");
	if (s && n && id.length==3)
	{
		if (!isNaN(id[0]) && !isNaN(id[1]) && !isNaN(id[2]))
		{
			s.innerHTML = "已"+s.innerHTML;
			n.innerHTML = 1 + parseInt(n.innerHTML);
		}
		else{alert(str);}
	}
	else{alert(str);}
}

//更改字体大小
var status0 = '';
function fontSize(type,objname){
	var currentfontsize,currentlineheight;
	currentfontsize=parseInt(document.getElementById(objname).style.fontSize);
	//currentlineheight=parseInt(document.getElementById(objname).style.lineHeight); 
	if (type=="b"){
		if(currentfontsize<64){			
			document.getElementById(objname).style.fontSize=(++currentfontsize)+"pt";
			//document.getElementById(objname).style.lineHeight=(++currentlineheight)+'pt';
		}
	}else{
		if(currentfontsize>8){
			document.getElementById(objname).style.fontSize=(--currentfontsize)+"pt";
			//document.getElementById(objname).style.lineHeight=(--currentlineheight)+'pt';
		}
	}
}