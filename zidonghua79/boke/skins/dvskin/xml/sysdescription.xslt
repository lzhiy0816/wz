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
<xsl:variable name="String_0" title="">
<![CDATA[
{$refresh}
<div id="pagemain">
<center>
<BR/>
<div class="td1" style="width:500px;">
	<div class="td2" style="height:25px;line-height:24px;">
		&nbsp;<b>��ʾ��Ϣ��</b>
	</div>
	<div class="td0" style="margin-left:20px;">
		{$description}
		<br/>
	</div>
	{$refreshinfro}
	<hr class="post" align="center">
	<div align="center">
		<a href="javascript:history.go(-1)">������һҳ</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<a href="#" onclick="window.close()">�رմ���</a>
	</div>
</div>
<BR/>
</center>
</div>
]]>
</xsl:variable>
<xsl:variable name="String_1" title=""><![CDATA[<li>{$msg}</li>]]></xsl:variable>
<xsl:variable name="String_2" title=""><![CDATA[�벻Ҫ�ⲿ�ύ���ݡ�]]></xsl:variable>
<xsl:variable name="String_3" title=""><![CDATA[�벻Ҫ�ظ��ύ���ݡ�]]></xsl:variable>
<xsl:variable name="String_4" title=""><![CDATA[��鿴�����ݲ����ڻ��������]]></xsl:variable>
<xsl:variable name="String_5" title=""><![CDATA[���û������ڻ���д����������!]]></xsl:variable>
<xsl:variable name="String_6" title=""><![CDATA[�����û����Ѿ���ע�ᣬ��������д!]]></xsl:variable>
<xsl:variable name="String_7" title=""><![CDATA[��֤��У��ʧ�ܣ��뷵��ˢ��ҳ�����������֤�롣]]></xsl:variable>
<xsl:variable name="String_8" title=""><![CDATA[�û�������Ϊ�ջ����50���ַ���]]></xsl:variable>
<xsl:variable name="String_9" title=""><![CDATA[�ύ�����ݲ��Ϸ���]]></xsl:variable>
<xsl:variable name="String_10" title=""><![CDATA[��������ֻ����Ӣ����ĸ�����֡��»�����д��]]></xsl:variable>
<xsl:variable name="String_11" title=""><![CDATA[��д���ʺ����벻��Ϊ�գ�]]></xsl:variable>
<xsl:variable name="String_12" title=""><![CDATA[���ͱ����˵������Ϊ�ջ����150���ַ���]]></xsl:variable>
<xsl:variable name="String_13" title=""><![CDATA[���Ĳ����Ѿ����ڻ򲩿ͱ�ʶ�ѱ�ʹ�ã�]]></xsl:variable>
<xsl:variable name="String_14" title=""><![CDATA[����û�� <a href="bokeapply.asp">������˲���</a> �� <a href="login.asp">��δ��¼��̳</a>]]></xsl:variable>
<xsl:variable name="String_15" title=""><![CDATA[������Ĳ�������������������롣]]></xsl:variable>
<xsl:variable name="String_16" title=""><![CDATA[���ɹ��޸��˸������ϡ�]]></xsl:variable>
<xsl:variable name="String_17" title=""><![CDATA[��������²���������ظ�����������벻һ�£����������롣]]></xsl:variable>
<xsl:variable name="String_18" title=""><![CDATA[���ɹ��޸��˲������롣]]></xsl:variable>
<xsl:variable name="String_19" title=""><![CDATA[���ɹ��޸��˸��˲������á�]]></xsl:variable>
<xsl:variable name="String_20" title=""><![CDATA[�ؼ��ֺ��滻�ı�������д��]]></xsl:variable>
<xsl:variable name="String_21" title=""><![CDATA[���͹ؼ������óɹ���]]></xsl:variable>
<xsl:variable name="String_22" title=""><![CDATA[ͬһ�ؼ��ֲ������ö�Ρ�]]></xsl:variable>
<xsl:variable name="String_23" title=""><![CDATA[���͹ؼ���ɾ���ɹ���]]></xsl:variable>
<xsl:variable name="String_24" title=""><![CDATA[������Ŀ���óɹ���]]></xsl:variable>
<xsl:variable name="String_25" title=""><![CDATA[�����벩����Ŀ���ơ�]]></xsl:variable>
<xsl:variable name="String_26" title=""><![CDATA[������Ŀɾ���ɹ���]]></xsl:variable>
<xsl:variable name="String_27" title=""><![CDATA[Ƶ������ҳ�����������³ɹ���]]></xsl:variable>
<xsl:variable name="String_28" title=""><![CDATA[��ʱû�пɹ�ѡȡ��ģ�档]]></xsl:variable>
<xsl:variable name="String_29" title=""><![CDATA[���ģ����ĳɹ���]]></xsl:variable>
<xsl:variable name="String_30" title=""><![CDATA[���ⲻ��Ϊ�ջ����250���ַ���]]></xsl:variable>
<xsl:variable name="String_31" title=""><![CDATA[�����ؼ��ֲ��ܶ���250���ַ���]]></xsl:variable>
<xsl:variable name="String_32" title=""><![CDATA[��ѡȡ�������]]></xsl:variable>
<xsl:variable name="String_33" title=""><![CDATA[��ѡȡ�������⡣]]></xsl:variable>
<xsl:variable name="String_34" title=""><![CDATA[��ѡȡ��������Ŀ�����÷������Ŀδ�������ȴ�����Ŀ�ٽ��з��������]]></xsl:variable>
<xsl:variable name="String_35" title=""><![CDATA[����д�������ݡ�]]></xsl:variable>
<xsl:variable name="String_36" title=""><![CDATA[���ҵ���Ϣ�����ڻ�û�н��иù������Ȩ�ޡ�]]></xsl:variable>
<xsl:variable name="String_37" title=""><![CDATA[����ɹ���]]></xsl:variable>
<xsl:variable name="String_38" title=""><![CDATA[��û��Ȩ�޽��лظ���༭������]]></xsl:variable>
<xsl:variable name="String_39" title=""><![CDATA[�û�������Ϊ�ջ��зǷ��ַ���]]></xsl:variable>
<xsl:variable name="String_40" title=""><![CDATA[����д���û����Ѿ����ڣ����޸��û������½���ٽ��в������ۡ�]]></xsl:variable>
<xsl:variable name="String_41" title=""><![CDATA[��ǰ������ֹͣע��!]]></xsl:variable>
<xsl:variable name="String_42" title=""><![CDATA[
<div class="td3">3����Զ�����{$refreshname}...</div>
]]></xsl:variable>
<xsl:variable name="String_43" title=""><![CDATA[��Ǹ����û��Ȩ�޷���!]]></xsl:variable>
<xsl:variable name="String_44" title=""><![CDATA[�����ڱ�ϵͳע���˸��˲����ʺţ�<a href="boke.asp">��˽������Ĳ���</a>������Ҫ����������������ע����̳�û���]]></xsl:variable>
<xsl:variable name="String_45" title=""><![CDATA[������̳ <a href="reg.asp" alt="��̳ע��">ע��</a> �� <a href="login.asp" alt="��̳��¼">��¼</a> ���ټ��������û����ͣ�]]></xsl:variable>
<xsl:variable name="String_46" title=""><![CDATA[�ò����û������ڻ���д���������������� <a href="bokeapply.asp">���벩��</a> �� <a href="login.asp">��¼��̳</a>]]></xsl:variable>
<xsl:variable name="String_47" title=""><![CDATA[��ϲ�������ĸ��˲����ѳɹ���ͨ��<a href="boke.asp">��˽������ĸ��˲���ҳ��</a>]]></xsl:variable>
<xsl:variable name="String_48" title=""><![CDATA[����Ŀ��δ�����Ϣ]]></xsl:variable>
<xsl:variable name="String_49" title=""><![CDATA[��ѡ�����Ŀ��������]]></xsl:variable>
<xsl:variable name="String_50" title=""><![CDATA[������Ϣ ɾ�����ƶ� �ɹ�]]></xsl:variable>
<xsl:variable name="String_51" title=""><![CDATA[�����ļ�ɾ���ɹ�]]></xsl:variable>
<xsl:variable name="String_52" title=""><![CDATA[��ǰ�����ѱ��رա�]]></xsl:variable>
<xsl:variable name="String_53" title=""><![CDATA[��ǰ�����Ѿ���ͣ�رա�]]></xsl:variable>
<xsl:variable name="String_54" title=""><![CDATA[���� <a href="?s=1&t={$t}&m=2">��ӱ�Ƶ������</a> ���ٽ�����Ϣ����������]]></xsl:variable>
</xsl:stylesheet>
