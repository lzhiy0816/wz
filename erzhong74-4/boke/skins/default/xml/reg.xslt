<?xml version="1.0" encoding="Gb2312"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="html" indent="yes" version="4.0"/>
<xsl:preserve-space elements="code"/>
<!--
	Copyright (C) 2004,2005 AspSky.Net. All rights reserved.
	Written by dvbbs.net Fssunwin
	Web: http://www.aspsky.net/,http://www.dvbbs.net/
	Email: eway@aspsky.net
-->
<xsl:variable name="String_0" title="reg">
<![CDATA[
<script language="javascript" src="Boke/Script/Dv_form.js" type="text/javascript"></script>
<script language="JavaScript">
<!--
function CheckForm(theForm){
	if (theForm.NickName.value.length=="") {
		alert("��ʾ��\n\n�û�����������д��");
		theForm.NickName.focus();
		return false;
	};
	if (theForm.BokeName.value.length=="") {
		alert("��ʾ��\n\n�������Ʊ�����д��");
		theForm.BokeName.focus();
		return false;
	};
	if (theForm.BokePassWord.value.length=="") {
		alert("��ʾ��\n\n�������������д��");
		theForm.BokePassWord.focus();
		return false;
	};
	if (theForm.BokeTitle.value.length=="") {
		alert("��ʾ��\n\n���ͱ��������д��");
		theForm.BokeTitle.focus();
		return false;
	};

	if (theForm.BokeCTitle.value.length=="") {
		alert("��ʾ��\n\n����˵��������д��");
		theForm.BokeCTitle.focus();
		return false;
	};
	if (theForm.codestr){
		if (theForm.codestr.value.length=="") {
			alert("��ʾ��\n\n��֤�������д��");
			theForm.codestr.focus();
			return false;
		};
	}

return (true);
}
//-->
</script>
<div id="pagemain">
<center>
<div class="td1" style="width:600px;">
<form method="post" action="?action=savereg" name="apply" onSubmit="return CheckForm(this);">
<div class="td2" style="height:25px;line-height:24px;">
&nbsp;<B>{$RegMsg}</B>
</div>
&nbsp;<B>˵��</B>��<BR />&nbsp;�� �û�������ʾ��������Ĳ��������У�<BR />&nbsp;�� ���ͱ�ʶ�����̺üǣ����ڷ���ϵͳ��������(��ϵͳ֧��)���û�ֱ�ӷ��ʲ�����<BR />&nbsp;�� �����������������͵Ĺ����������̳���벻ͬ��
<BR />
<hr class="post">
&nbsp;<B>��������</B>��
<div class="td0" style="margin-left:20px;">
<li>���ͱ�ʶ��<input type="text" name="BokeName" onkeyup="javascript:Dv_Form.CheckEnglish(this);"/> <input type="submit" value="���̼������Ĳ���"></li>
<li>���ͱ�ʶֻ��ʹ��Ӣ����ĸ�����ּ��»�"_"��д��</li><li><U>����ע���в������뼴Ϊ������̳���룬�û�������Ϊ������̳�û�����</U></li><li>����������ͺ�ʹ�� ��̳���� �����˲��͹������޸���Ӧ�ĸ������á�</li>
</div>
</form>
<form method="post" action="?action=savereg" name="apply" onSubmit="return CheckForm(this);">
<hr class="post">
&nbsp;<B>��������</B>��
<div class="td0" style="margin-left:20px;">
<li>�û�������<input type="text" name="NickName"/></li>
<li>���ͱ�ʶ��<input type="text" name="BokeName" onkeyup="javascript:Dv_Form.CheckEnglish(this);"/> * ֻ��ʹ��Ӣ����ĸ�����ּ��»�"_"��д��</li>
<li>�������룺<input type="password" name="BokePassWord"/></li>
<li>���ͱ��⣺<input type="text" name="BokeTitle" size="50" onkeyup="javascript:Dv_Form.CheckLength(this,150);"/></li>
<li>����˵����<input type="text" name="BokeCTitle" size="50" onkeyup="javascript:Dv_Form.CheckLength(this,150);"/></li>
</div><BR/>
<center><input type="submit" value="�� ��"/>&nbsp;&nbsp;<input type="reset" value="�� ��"/></center>
</form>
</div>
</center>
</div>
]]>
</xsl:variable>
<xsl:variable name="String_1" title="">
<![CDATA[
<li>
��&nbsp;&nbsp;֤&nbsp;�룺{$Dv_GetCode}<!--<input type="text" name="codestr" maxlength="4" size="4">&nbsp;<img src="DV_getcode.asp" height="12">-->
</li>
]]>
</xsl:variable>
<xsl:variable name="String_2" title="">
<![CDATA[

]]>
</xsl:variable>
<xsl:variable name="String_3" title="">
<![CDATA[

]]>
</xsl:variable>
<xsl:variable name="String_4" title="strings"><![CDATA[]]>
</xsl:variable>
<xsl:variable name="String_5" title="strings">
<![CDATA[
�Ĳ���
]]>
</xsl:variable>
<xsl:variable name="String_6" title="strings"><![CDATA[��ӭ <b>{$UserName}</b> ������̳���ͣ�]]>
</xsl:variable>
<xsl:variable name="String_7" title="strings">
<![CDATA[
�Һ�����ʲôҲû����
]]>
</xsl:variable>
</xsl:stylesheet>
