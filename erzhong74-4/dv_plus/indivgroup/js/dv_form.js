/*//////////////////////////////////////////////////////////////////
 * Script For Dvbbs Program -- ���������
 * Version: 8.1.0
 * Copyright (C) 2000 - 2005 bbs.dvbbs.net
 *
 * For further information visit:
 * 		http://bbs.dvbbs.net/
 * 		http://www.aspsky.net/
 *
 * Builder: Fssunwin
 * Created: 2005-04-23
*///////////////////////////////////////////////////////////////////

var Lang = new Array() ;
var Dvbbs_bIsIE5=document.all;
var Dvbbs_Mode = 3;
var Dvbbs_bTextMode=1;
if (Dvbbs_bIsIE5){
}
else{
	var Dvbbs_bIsNC=true;
}
/*
* ������ʾ��Ϣ
*/
Lang["Msg"]				= "��ʾ��";
Lang["NEED_TEL_NUM"]	= "����ȷ��д�绰���룡";
Lang["NEED_NUM"]		= "����ȷ��д��ֵ!";
Lang["NOT_DATE"]		= "����ȷ��д��ѡȡ��ȷ������!";
Lang["NO_Checked"]		= "δѡ��";
Lang["NOT_MIN"]			= "������ֵ����С�ڣ�"
Lang["NOT_MAX"]			= "������ֵ���ܴ��ڣ�"
Lang["NOT_MAXLen"]		= "�ַ����Ʋ�Ҫ������"
Lang["NOT_MINLen"]		= "�ַ����Ʋ�ҪС�ڣ�"
Lang["NEED_English"]	= "ֻ��ʹ��Ӣ����ĸ�����ֻ��»�\"_\"��д��"

/*
* ����ִ����󷽷�
*/
String.prototype.remove=function(start,length){
	var s='';
	if (start>0) s=this.substring(0,start);
	if (start+length<this.length) s+=this.substring(start+length,this.length);
	return s;
};

String.prototype.trim=function(){
	return this.replace(/(^\s*)|(\s*$)/g,'');
};

String.prototype.ltrim=function(){
	return this.replace(/^\s*/g,'');
};

String.prototype.rtrim=function(){
	return this.replace(/\s*$/g,'');
};

String.prototype.replaceNewLineChars=function(replacement){
	return this.replace(/\n/g,replacement);
}

var Dv_Form = new FormObject();
function FormObject(){
	// �޸� textarea �ı�����༭���߶�
	this.set_Size = function(objname,num){
		if (objname=="[object]"){
			var obj = objname;
		}else{
			var obj = Dvbbs.Objects(objname);
		};
		if (obj)
		{
			if (parseInt(obj.rows)+num>=3) {obj.rows = parseInt(obj.rows) + num;	};
		};
	};

	//SELECT��ѡȡ
	this.Set_Select = function(objname,Value){
		if (objname=="[object]"){
			var obj = objname;
		}else{
			var obj = Dvbbs.Objects(objname);
		};
		if (obj)
		{
			for (var i=0;i<obj.length;i++){
				if (obj.options[i].value==Value){obj.options[i].selected=true;break;};
			};
		};
	};

	//Weather SELECT��ѡȡ
	this.Weather_Select = function(objname,Value,WeatherSrc,PicPath){
		if (objname=="[object]"){
			var obj = objname;
		}else{
			var obj = Dvbbs.Objects(objname);
		};
		var SrcFile = "";
		if (obj)
		{
			for (var i=0;i<obj.length;i++){

				if (obj.options[i].value.split("|")[1]==Value){
					SrcFile = obj.options[i].value.split("|")[0];
					obj.options[i].selected=true;break;
				};
			};
		};
		var ObjSrc = Dvbbs.Objects(WeatherSrc);
		if (ObjSrc)
		{
			ObjSrc.src = PicPath + SrcFile;
		}
	};

	//��ѡ��ѡȡ
	this.Set_Radio = function(objname,Value){
		if (objname=="[object]"){
			var obj = objname;
		}else{
			var obj = Dvbbs.Objects(objname);
		};
		if (obj)
		{
			for (var i=0;i<obj.length;i++){
				if (obj[i].value==Value){obj[i].checked=true;break;};
			};
		};
	};

	/**
	* �ı�������
	* <input type="text" onkeyup="javascript:Dv_Form.CheckNumer(this,0,10);" >...
	*/

	//��֤����ֵ object ������ | n_min��Сֵ | n_max ���ֵ
	this.CheckNumer = function (Object,n_min,n_max)
	{
		var Value = Object.value;
		if (isNaN(Value)){
			alert(Lang["NEED_NUM"]);
			Object.value = 0;
		}
		else{
			Value = parseFloat(Value);
			if (!isNaN(n_min)){
				n_min = parseFloat(n_min);
				if (Value<n_min){alert(getMsg(Lang["NOT_MIN"]+n_min));Object.value = n_min;}
			}
			if (!isNaN(n_max)){
				n_max = parseFloat(n_max);
				if (Value>n_max){alert(getMsg(Lang["NOT_MAX"]+n_max));Object.value = n_max;}
			};
		};
	};

	/**
	*�ı��������� object ������ | n_max ���ֵ 	//onkeydown or onkeyup
	* <input type="text" onkeyup="javascript:Dv_Form.CheckLength(this,255);" >...
	*/
	this.CheckLength = function(obj,n_max){
		if (obj){
			Val_Length = obj.value;
			if (!isNaN(n_max)){
				if (Val_Length.length > n_max){
					obj.value = Val_Length.substring(0,n_max);alert(getMsg(Lang["NOT_MAXLen"]+n_max));
				};
			};
		};
	};
	//���ָ���ļ����ֽ���
	this.GetLength = function(_Getobj,Putobj){
		var _Putobj = Dvbbs.Objects(Putobj);
		if (!_Getobj||!_Putobj){alert("test");return false;}
		else{
			var oBjName = _Putobj.nodeName.toLowerCase();
			var TotalLen = _Getobj.value.length;
			if (oBjName=="input"||oBjName=="textarea"){_Putobj.value = TotalLen;}else{_Putobj.innerText = TotalLen;};
		};
	};

	/*
	* ��ɫѡȡ�� ����ͼƬ��ɫ�ɸ���Ĭ��ֵ
	* img_obj ͼ����� ,input_val ���������
	* <input type="text" name="setColor" value="#000000"/><img border="0" src="rect.gif" style="cursor:pointer;background-Color:#000000;" onclick="Dv_Form.Getcolor(this,'setColor');" alt="ѡȡ��ɫ!"/>
	*/
	var ColorImg;
	var ColorValue;
	this.Getcolor = function (img_obj,input_val){
		var obj = document.getElementById("Dv_colourPalette");
		if (!obj)
		{
			obj = document.createElement('iframe');
			obj.id = "Dv_colourPalette";
			obj.width = "260px";
			obj.height = "165px";
			obj.src = "selcolor.htm";
			obj.style.visibility = "hidden"
			obj.style.position = "absolute"
			obj.style.border = "1px gray solid"
			obj.frameborder="0";
			obj.scrolling="no";
			document.getElementsByTagName("body")[0].appendChild(obj);
		};
		ColorImg = img_obj;
		ColorValue = Dvbbs.Objects(input_val);
		if (obj){
			obj.style.left = Dvbbs.getOffsetLeft(ColorImg) + "px";
			obj.style.top = (Dvbbs.getOffsetTop(ColorImg) + ColorImg.offsetHeight) + "px";
			if (obj.style.visibility=="hidden"){obj.style.visibility="visible";}
			else {obj.style.visibility="hidden";};
		};
	};

	this.setColor = function(color){
		if (ColorValue){ColorValue.value = color;}
		if (ColorImg){ColorImg.style.backgroundColor = color;}
		document.getElementById("Dv_colourPalette").style.visibility="hidden";
	};

	/**
	*�ı��ַ��������� object ������ | n_max ���ֵ 	//onkeydown or onkeyup
	* <input type="text" onkeyup="javascript:Dv_Form.CheckEnglish(this);" >...
	*/
	this.CheckEnglish = function(obj){
		if (obj){
			if (CheckIfEnglish(obj.value)==false){
				obj.value = "";
				alert(getMsg(Lang["NEED_English"]));
			};
		};
	};

}

/*
* ��ȡ��ʾ��Ϣ
* @return String
*/
	function getMsg(pKey) {
			return Lang["Msg"] +"\n\n" + pKey ;
	}

/*
*	��ȡ��Ӧ��ʽ������
*	pValΪ���ӵ���������Ϊ+/-�����������뵱�����������ڣ�
*	�� pDiv Ϊ'/'ʱ������ '2004/11/01'�ĸ�ʽ.
*   getDate(0,"/")
*   @param pVal ��������.
*   @param pDiv �ָ���.
*	@return String
*/
	function getDate(pVal, pDiv) {
		var dtnum = Date.parse(new Date());
		dtnum = dtnum + (pVal * (3600000 * 24));
		dt = new Date(dtnum);
		var year = dt.getYear();
		var month = (dt.getMonth() + 1) + '';
		var day = dt.getDate() + '';
		if(month.length == 1)	month = "0" + month;
		if(day.length == 1) 	day = "0" + day;
		rtn = year + pDiv + month + pDiv + day;
		return rtn;
	}


/*
* ����Ƿ�ֻ��Ӣ�ļ����ּ�"_"
* @param pKey String
* @return boolean
*/
	function CheckIfEnglish(pKey)
	{
		var Letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_";
		 var i;
		 var c;
		 for( i = 0; i < pKey.length; i ++ )
		 {
			  c = pKey.charAt( i );
		  if (Letters.indexOf( c ) < 0)
			 return false;
		 };
		 return true;
	};



/*
* ����Ƿ�������
* @param pKey String
* @return boolean
*/
	function CheckIfChina(pKey){
	 var str=pKey;
	 var lst = /[u00-uFF]/; 
	 if (!lst.test(str))
			 return false;
	  return true;
	};

/*
* Email ���
* @param str ����ִ�
* @return boolean
*/
	function CheckEmail(str)
	{
		var filter=/^.+@.+\..{2,3}$/i;
		//var filter=/^([\w-]+(?:\.[\w-]+)*)@((?:[\w-]+\.)*\w[\w-]{0,66})\.([a-z]{2,6}(?:\.[a-z])?)$/i
		if (filter.test(str)) { return true; }
		else { return false; };
	};

/*
*   isNumber('1234')
*   @param pVal  ����ִ�
*	@return boolean
*/
	function isPhone(pVal)
	{
		if((strTrim(pVal)).length == 0 ) {return false;};
		var str = "0123456789()-,*# ";
		for( i = 0; i < pVal.length; i++){
			if( str.indexOf( pVal.charAt( i ) ) < 0 ){
				alert(getMsg('NEED_TEL_NUM'));
				return false;
				break;
			};
		};
		return true;
	};

/*
*	��֤�Ƿ����֤����
*	@param strValue  ����ִ�
*	@return boolean
*/
	function isCertificateNum(strValue)
	{  
	  if(!(strValue.length==15||strValue.length==18)) return(false);	
		var str = "0123456789Xx";
		for( i = 0; i < strValue.length; i++){
			if( str.indexOf( strValue.charAt( i ) ) < 0 ) return false;    
		return true;
		};
	};

/*
* �����滻����
*  strԭ�ִ���aҪ����pattern��b�����ִ���i�Ƿ����ִ�Сд��Msg��ʾ��Ϣ
*/
	function Replace(str,a,b,i,Msg){	
		a = a.replace("?","\\?");
		if (i==null){var r = new RegExp(a,"gi");}
		else if (i) {var r = new RegExp(a,"g");}
		else{var r = new RegExp(a,"gi");
		};
		if (r.test(str))
		{
			if (Msg){alert(getMsg(Msg));};
			return str.replace(r,b); 
		}else{return str;};
	};

/*
OchangeSel_Cat ObjName�б������Stype�������࣬Val=CATID��ֵ��
*/
function OchangeSel_Cat(ObjName,Stype,Val){
	var Obj = Dvbbs.Objects(ObjName);
	var k=0;var sel =0;
	if (Obj&&Stype!="")
	{

		if (Obj.length>0)
		{
			for(var i = 10; i >= 0; i--){
				Obj.options.remove(i);
			};
		}
		Stype = parseInt(Stype)
		if (Stype >=0&&BokeCat_ID[Stype].length>0)
		{
			for (var CatID in BokeCat_ID[Stype]){
				if (Val==BokeCat_ID[Stype][CatID])
				{
					sel = k;
				}
				Obj.options[k++] = new Option(BokeCat_Title[Stype][CatID], BokeCat_ID[Stype][CatID]);
			};
			if (sel>0)
			{
				Obj.options[sel].selected = true;
			};
		};
	};
}

/*
�༭����ع���
*/

function getContent(objname){
	var oEditor = FCKeditorAPI.GetInstance(objname) ;
	alert(oEditor.GetXHTML())

}

function put_ubb(objname,ubbcode,val){
	var oEditor = FCKeditorAPI.GetInstance(objname) ;
	var Ubb_str2= "[/"+ubbcode+"]";
	if (val){
		var Ubb_str1= "["+ubbcode+"="+val+"]";
	}else{
		var Ubb_str1= "["+ubbcode+"]";
	};
	if (oEditor.EditorDocument.selection.type=="Text"){
		var Selval = oEditor.EditorDocument.selection.createRange().text;
		oEditor.InsertHtml(Ubb_str1 + Selval +Ubb_str2);
	}else
	{
		oEditor.InsertHtml(Ubb_str1 + Ubb_str2);
	};
}

function put_html(objname,val){
	var oEditor = FCKeditorAPI.GetInstance(objname) ;
	if (oEditor){
		oEditor.InsertHtml(val);
	};
}

function GetNode(objname,toobj,Maxlen){
	var oEditor = FCKeditorAPI.GetInstance(objname);
	toobj = Dvbbs.Objects(toobj);
	if (oEditor&&toobj){
		if (oEditor.EditorDocument.selection.type=="Text"){
			var Selval = oEditor.EditorDocument.selection.createRange().text;
			if (Selval.length>Maxlen)
			{
				Selval = Selval.substring(0,Maxlen)
			}
			toobj.value = Selval;
		}else
		{
			alert("��ѡȡ��ΪժҪ�����ݣ�");
		};
	};
}

function PlayUbb(ObjName,code){
	var txt1 = "";
	var txt2 = "";
	var val1 = "";
	var val2 = "";
	switch (code)
	{
		case 'FLASH':
			txt1 = "Flash��ȣ��߶�";
			val1 = "500,350";
			txt2 = "Flash �ļ��ĵ�ַ";
			val2 = "http://"
			break;
		case 'RM':
			txt1 = "��Ƶ�Ŀ�ȣ��߶ȣ����Ų���\r(���Ų�����0���ֶ����ţ�1���Զ�����)";
			val1 = "500,350,1";
			txt2 = "��Ƶ�ļ��ĵ�ַ";
			val2 = "http://"
			break;
		case 'MP':
			txt1 = "��Ƶ�Ŀ�ȣ��߶ȣ����Ų���\r(���Ų�����0���ֶ����ţ�1���Զ�����)";
			val1 = "500,350,1";
			txt2 = "��Ƶ�ļ��ĵ�ַ";
			val2 = "http://"
			break;
		case 'QT':
			txt1 = "��Ƶ�Ŀ�ȣ��߶�";
			val1 = "500,350";
			txt2 = "��Ƶ�ļ��ĵ�ַ";
			val2 = "http://"
			break;
		default : return;
	};
	var addtxt1 = prompt(txt1,val1);
	if (addtxt1!=null) {
		var addtxt2=prompt(txt2,val2);
		if (addtxt2!=null) {
			if (addtxt1=="") {             
				var AddTxt="<br/>["+code+"=500,350]"+addtxt2+"[/"+code+"]<br/>";
				put_html(ObjName,AddTxt);
			} else {
				AddTxt="<br/>["+code+"="+addtxt1+"]"+addtxt2+"[/"+code+"]<br/>";
				put_html(ObjName,AddTxt);
			}
	    };
    };
}

function CheckPostForm(Obj){
	document.getElementById('PostContent').value=dvtextarea.save();
	if (Obj.sCatID)
	{
		if (Obj.sCatID.value=="-1"||Obj.sCatID.value=="")
		{
			alert("��ѡȡ����Ļ��⣡");
			return false;
		};
	};
	if (Obj.Catid)
	{
		if (Obj.Catid.value=="-1"||Obj.Catid.value=="")
		{
			alert("��ѡȡ��Ӧ����Ŀ��������Ŀ���ٽ��з���");
			return false;
		};
	};
}

function Dvbbs_foralipay(obj)
{
	var arr = showModalDialog("boke/images/alipay/alipay.htm", window, "dialogWidth:20em; dialogHeight:34em; status:0; help:0");
	if (arr)
	{
		put_html(obj,arr);
		//IframeID.document.body.innerHTML+=arr;
	}
	//IframeID.focus();
}