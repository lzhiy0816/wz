function ajaCreateData(_form){
	var data="";
	if(_form){
		for(var i=0;i<_form.elements.length;i++){
			if(data==""){
				data=_form.elements[i].name+"="+escape(_form.elements[i].value);
			}else{
				data+="&"+_form.elements[i].name+"="+escape(_form.elements[i].value);
			}
		}
	}
	return data;
}

function ajaxSubmit(actionAddress,data){
	var xmlhttp;
	try{
		xmlhttp=new XMLHttpRequest();
	}catch(e){
		try{
			xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
		}catch(e){
			try{
				xmlhttp=new ActiveXObject('Msxml2.XMLHTTP');
			}catch(e){
				alert('��ʹ�õ��������֧��ajax�ύ���ݣ���ʹ��IE6.0��Mozilla Firefox 2.0��������������');
			}
		}
	}
	xmlhttp.onreadystatechange=function(){
		if (4==xmlhttp.readyState){
			if (200==xmlhttp.status){
				eval(xmlhttp.responseText.substring(0));
				if (''==returnInfo.formName){
					alert(returnInfo.formValue);
				}else{
					document.getElementById(returnInfo.formName).style.display = 'block';
					document.getElementById(returnInfo.formName).innerHTML = returnInfo.formValue;
				}
			}else{
				alert(xmlhttp.status+'ͶƱʧ�ܣ������жϻ������æ�����Ժ����ԡ�')
				//������Ը��߿ͻ��ˣ��ύʧ�ܣ�������(״̬��xmlhttp.status)
			}
		}
	}
	xmlhttp.open("POST", actionAddress, true);
	//xmlhttp.setRequestHeader("Method", "POST " + actionAddress + " HTTP/1.1");
	xmlhttp.setRequestHeader('Content-type','application/x-www-form-urlencoded');
	xmlhttp.send(data);
	return false;
}