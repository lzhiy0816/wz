function DvSavePost(){
	var a=arguments;
	this.$=function(o){return document.getElementById(o);}
	this.frm=a[0]||document.forms[0];//���������
	this.evt=a[1]||window.event;//�¼�����
	this.mode=a[2]||'fastre';//����ģʽ
	this.max_title_length=a[3]||100;//������󳤶�
	this.max_content_length=a[4]||16240;//������̳���
	this.isok=true;
	this.chk_topic=function(o){
		if('fastre'!=this.mode){
			if(!o.value){
			this.$(o.name+"_chk").innerHTML='<font color="#FF0000">����������д����</font>';
			o.focus();
			this.isok=false;
			}
		};
		if(o.value.length>this.max_title_length){
			this.$(o.name+"_chk").innerHTML=' <font color="#FF0000">�����ⳤ�Ȳ��ܴ���'+this.max_title_length+'</font>';
			o.focus();
			this.isok=false;
		}
	};
	this.chk_content=function(o){
		//var len=o.value.replace(/<[^>]*>/gi,'').replace(/&\w{4};/gi,'').length;
		var len=o.value.length;
		if(0==len){
			this.$(o.name+"_chk").innerHTML='<font color="#FF0000">��������������û�м�⵽�������ݡ�</font>';
			this.isok=false;
		};
		if(o.value.length>this.max_content_length){
			this.$(o.name+"_chk").innerHTML=' <font color="#FF0000">���������ݳ��Ȳ��ܴ���'+this.max_content_length+',���Ѿ�������'+o.value.length+'����</font>';
			try{o.focus();}catch(e){}
			this.isok=false;
		}
	};
	this.chk_topicmode=function(modevalue,modelimit){
		if (modevalue.value==0&&modelimit.value==2){
			this.$("mode_chk").innerHTML=" <font color=\"#FF0000\">����û��ѡ��ר��</font>"
			this.isok=false;
		}
	};
	this.chk_code=function(){
		if (document.Dvform.codestr)
		{
			document.getElementById("GetCode").innerHTML="";
			if (''==document.Dvform.codestr.value)
			{
				document.getElementById("GetCode").innerHTML=" <font color=\"#FF0000\">����������ȷ����֤��</font>"
				document.Dvform.codestr.focus();
				this.isok=false;
			}
		}
	};
	this.chk_flash=function(){
		if (document.getElementById("phidstatus"))	{
			if (document.getElementById("phidstatus").value!="0")	{
				alert("��ʹ����Flash��ͼ���ܣ��뱣������ύ��");
				this.isok=false;
			}
		}
	}
	this.prevent=function(){
		try{if(this.evt.preventDefault){this.evt.preventDefault();}else{this.evt.returnValue=false;}}catch(er){}
	};
	this.send=function(){		
		if (this.isok){
			try{
				if (typeof(ajaxpost)!="undefined"){
					if(!ajaxpost.postForm(this.frm)){
						this.prevent();
					}
				}
			}
			catch (e){}
		}else{
			this.prevent();
		}
		return this.isok;
	};
}

function ChkPostMoney(n)
{
	var DivInfo = document.getElementById('PostMoneyInfo');
	var DivBuy_Setting = document.getElementById('Buy_setting');
	document.Dvform.ToMoney.value = "";
	document.Dvform.ToMoney.disabled = false;
	if (n!="")
	{
		switch (n)
		{
			case '0':
				document.Dvform.ToMoney.value = 1;
				DivInfo.innerHTML="�Զ������ͽ��������������ʱ�۳����û���Ӧ��ң��û��ɶԲ�ͬ�ظ��û��ڽ�ҷ�Χ�ڷֱ����ͣ�������Ͽɽ�����������ʣ����(δ�ͳ�)�����û�������";
				DivBuy_Setting.style.display="none";
				break;
			case '1':
				document.Dvform.ToMoney.disabled = true;
				DivInfo.innerHTML="�ظ��û����Զ�����������͸�������";
				DivBuy_Setting.style.display="none";
				break;
			case '2':
				document.Dvform.ToMoney.value = 1;
				DivInfo.innerHTML="�����߿��Զ������ӳ��۽���������������Ҫ֧����ҹ���ſ��Բ鿴����ȫ�����ݡ�";
				if (DivBuy_Setting)
				{
					DivBuy_Setting.style.display="";
				}
				break;
		}
	}else{
		DivInfo.innerHTML="";
	}
}