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
<xsl:variable name="String_0" title="login">
<![CDATA[
<script language="javascript" src="Boke/Script/Dv_form.js" type="text/javascript"></script>
<div id="pagemain">
<center>
<div class="td1" style="width:600px;height:153px;clear:both;">
	<div class="td2" style="line-height:24px;height:25px;">&nbsp;<b>�����¼</b></div>
	<form method="post" action="?Action=Login" name="apply">
		<div style="width:40%;float:left;padding-left:10px;"><BR/>
			{$UserMsg}
			<BR />
			���������Ĳ��������¼��
		</div>
		<div style="width:40%;float:left;border-left:1px #B3B3B3 dashed;padding-left:10px;"><BR/>
		<ul>
			<li>�������룺<input type="password" name="PassWord"/></li>
			<li>��&nbsp;&nbsp;֤&nbsp;�룺<input type="text" name="codestr" maxlength="4" size="4">&nbsp;<img src="DV_getcode.asp">
			</li>
		</ul>
		<BR/>
		<input type="submit" value="��¼���˲��͹���"/>
		<BR/><BR/>
		</div>
		<input type="hidden" value="{$ComeUrl}" name="f">
	</form>
</div>
</center>
</div>
]]>
</xsl:variable>
<xsl:variable name="String_1" title="strings"><![CDATA[��ӭ <b>{$UserName}</b> ��¼�û����͹���]]>
</xsl:variable>
<xsl:variable name="String_2" title="TopNav">
<![CDATA[
<LINK REL="stylesheet" type="text/css" href="Boke/Skins/dvskin/Css/UserManage.css"/>
<div id="navmenu">
<div id="nav_r">
<a href="BokeManage.asp">������ҳ</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="{$bokeurl}{$bokename}.index.html">�ҵĲ���</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="BokeApply.asp?Action=Logout">�˳�����</a>
</div>

<a href="?s=1&t=1">����</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="?s=1&t=2">�ղ�</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="?s=1&t=3">����</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="?s=1&t=4">����</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="?s=1&t=5">���</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="?s=2">����</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="?s=3">�ļ�</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="?s=4">ģ��</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="?s=5">��������</a>
</div>
]]>
</xsl:variable>
<xsl:variable name="String_3" title="LeftMenu">
<![CDATA[
<div id="leftmenu" Style="text-align:left;">
<style>
#leftmenu ul{
	list-style-position : ;
	list-style-type : circle;
	margin:0 0 0 5px;
	padding:0 0 0 0px;
}
#leftmenu ul li{
	margin:0 0 0 25px;
}
</style>
<ul>
<B>��ݲ���</B>
<li><a href="BokeManage.asp?s=1&t=1&m=1">��������</a></li>
<li><a href="BokeManage.asp?s=1&t=2&m=1">����ղ�</a></li>
<li><a href="BokeManage.asp?s=1&t=3&m=1">�������</a></li>
<li><a href="BokeManage.asp?s=1&t=4&m=1">��������</a></li>
<li><a href="BokeManage.asp?s=1&t=5&m=1">������</a></li>
<li><a href="?s=5&t=3">��������</a></li>
<li><a href="?s=5&t=4">�ؼ�������</a></li>
<li><a href="?s=6&t=1">Ƶ����������������</a></li>
<li><a href="?s=7&t=0">��������ͳ��</a></li>
</ul>
</div>
]]>
</xsl:variable>

<xsl:variable name="String_4" title="Main">
<![CDATA[
<div id="main" style="text-align:left;">
<div style="margin-left:20px;">
��ӭ�������Ĳ��͹���ʹ�ø����ǰ����ϸ�Ķ����ҳ���˵������������ɺ�ǵ��˳�����ף��ʹ����졣
<BR />
<BR />
<B>ʹ��С��ʿ</B>
<li>����ʱ��ע���ĸ��˲��ͣ�����������ѷ��������Ƿ�Ϸ�</li>
<li>���¡��ղء����Ӿ��ɽ���Ԥ���������������Ϣʱ����ص���Ϣ�����Ӧ�ķ�����</li>
<li>Ԥ�貿�ֳ��õĹؼ��ּ������ӣ�����ʾ��������������Ӧ�Ĺؼ��ֻ��Զ�Ϊ����ӳ�������</li>
<li>���ֹ��ܿɹرջ���ʾ���粩�ͽ��׹��ܿɹرա���Ŀ�����¿ɶ��岻��ʾ��</li>
<BR />
<BR />
{$sysmessage}
</div>
</div>
]]>
</xsl:variable>
<xsl:variable name="String_5" title="Main_UserInput">
<![CDATA[
<div id="main" style="text-align:left;">
{$UserAction}����&nbsp;&nbsp;��&nbsp;&nbsp;<a href="{$UserAction_1}&m=1">���{$UserAction}</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="{$UserAction_1}&m=2">{$UserAction}��Ŀ����</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="{$UserAction_1}&m=3">{$UserAction}��Ϣ����</a>
<BR />
<BR />
{$UserInputInfo}
</div>
]]>
</xsl:variable>
<xsl:variable name="String_6" title="Main_UserSetting">
<![CDATA[
<div id="main" style="text-align:left;">
��������&nbsp;&nbsp;��&nbsp;&nbsp;<a href="?s=5&t=1">��������</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="?s=5&t=2">��������</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="?s=5&t=3">��������</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="?s=5&t=4">�ؼ�������</a>
<BR />
<BR />
{$UserSettingInfo}
</div>
]]>
</xsl:variable>
<xsl:variable name="String_7" title="Main_UserSetting_Info">
<![CDATA[
<div class="topic">
���������޸�
</div>
<div style="margin-left:20px;">
<BR />
˵����������ϸ�����޸��뵽 <a href="mymodify.asp" target="_blank">��̳�û������޸�</a>
<BR />
<BR />
<FORM METHOD=POST ACTION="?s=5&t=1&Action=Save">
<li>���ͱ�ʶ��
{$BokeName}</li>
<li>�û�������
<input type="text" name="NickName" value="{$NickName}"/></li>
<li>���ͱ��⣺
<input type="text" name="BokeTitle" value="{$BokeTitle}"/></li>
<li>����˵����
<input type="text" name="BokeCTitle" value="{$BokeCTitle}"/>
* ������250�ֽ�</li>
<li>���͹��棺<BR /><textarea name="BokeNote" cols="40" rows="5">{$BokeNote}</textarea></li><BR /><BR />
<input type="submit" name="submit" value="�޸�����">
</FORM>
</div>
]]>
</xsl:variable>

<xsl:variable name="String_8" title="Main_UserSetting_PassWord">
<![CDATA[
<div class="topic">
���������޸�
</div>
<div style="margin-left:20px;">
<BR />
˵�����˴�����Ϊ����½���˲��͹�����������
<BR />
<BR />
<FORM METHOD=POST ACTION="?s=5&t=2&Action=Save">
<li>�������룺
<input type="password" name="PassWord"/></li>
<li>�µ����룺
<input type="password" name="nPass"/></li>
<li>�ظ����룺
<input type="password" name="rnPass"/></li><BR /><BR />
<input type="submit" name="submit" value="�޸�����">
</FORM>
</div>
]]>
</xsl:variable>

<xsl:variable name="String_9" title="Main_UserSetting_Setting">
<![CDATA[
<script language="javascript" src="Boke/Script/Dv_form.js" type="text/javascript"></script>
<div class="topic">����վ������</div>
<br/>
<form method=post action="?s=5&t=3&Action=Save" name="setting">
	<div id="mainlist">
	<div class="set0">
		<div class="Set1">����վ����ࣺ</div>
		<div class="Set2">
			<Select Name="SysCatID" Size="1">{$CatList}</Select>
		</div>
	</div>
	<div class="set0">
		<div class="Set1">�Ƿ��������˷��ʣ�</div>
		<div class="Set2"><input type="radio" name="Setting0" value="0"/> �� <input type="radio" name="Setting0" value="1"/> ��</div>
	</div>
<!-- 	<div class="set0">
		<div class="Set1">Ĭ�ϱ༭��ѡ��</div>
		<div class="Set2"><input type="radio" name="Setting1" value="0"/> �༭��1 <input type="radio" name="Setting1" value="1"/> �༭��2</div>
	</div> -->
	<div class="set0">
		<div class="Set1">�Զ���ȡ��������</div>
		<div class="Set2"><input type="text" size="5" name="Setting2" value="{$Setting2}"/> �֣���ժҪ����Ϊ��ʱ���Զ������ý�ȡ������ΪժҪ��</div>
	</div>
	<div class="set0">
		<div class="Set1">��ʾ������������</div>
		<div class="Set2"><input type="text" size="5" name="Setting3" value="{$Setting3}"/> ��</div>
	</div>
	<div class="set0">
		<div class="Set1">�Ƿ��������ۣ�</div>
		<div class="Set2"><input type="radio" name="Setting4" value="0"/> �� <input type="radio" name="Setting4" value="1"/> ��</div>
	</div>
	<div class="set0">
		<div class="Set1">�û����ۺ󷵻�ҳ�棺</div>
		<div class="Set2"><input type="radio" name="Setting5" value="0"> ����ϵͳ��ҳ <input type="radio" name="Setting5" value="1"/> �û�������ҳ <input type="radio" name="Setting5" value="2"/> ��������</div>
	</div>
	<div class="set0">
		<div class="Set1">��ҳ��ʾ��������</div>
		<div class="Set2"><input type="text" size="5" name="Setting6" value="{$Setting6}"/> ��</div>
	</div>
	<div class="set0">
		<div class="Set1">�б�ÿҳ��ʾ��������</div>
		<div class="Set2"><input type="text" size="5" name="Setting7" value="{$Setting7}"/> ��</div>
	</div>
	<div class="set0">
		<div class="Set1">����ÿҳ��ʾ��������</div>
		<div class="Set2"><input type="text" size="5" name="Setting8" value="{$Setting8}"/> ��</div>
	</div>
	<div class="set0">
		<div class="Set1">����Ĭ�ϴ�С��</div>
		<div class="Set2"><input type="text" size="5" name="Setting9" value="{$Setting9}"/> px</div>
	</div>
	<div class="set0">
		<div class="Set1">����Ĭ���м�ࣺ</div>
		<div class="Set2"><input type="text" size="5" name="Setting10" value="{$Setting10}"/> px</div>
	</div>
	<div class="set0">
		<div class="Set1">������ʾ��ӡ������</div>
		<div class="Set2"><input type="text" size="5" name="Setting11" value="{$Setting11}"/> ��</div>
	</div>
	<div class="set0">
		<div class="Set1">��ӡͼƬ���У���</div>
		<div class="Set2">
			<input type="text" size="20" name="Setting12" value="{$Setting12}" id="Setting12"/>
			<select onchange="Chang(this.value,'DiaryFootMart1','boke/images/sex/','Setting12')">
				<option value="">ѡȡ��ӡͼ��
				<option value="sex1.gif" >sex1.gif</option>
				<option value="sex2.gif" >sex2.gif</option>
				<option value="footmark1.gif" >footmark1.gif</option>
				<option value="footmark3.gif" >footmark3.gif</option>
				<option value="footmark5.gif" >footmark5.gif</option>
			</select>
			&nbsp;<img id=DiaryFootMart1 src="{$Setting12}" border="0" alt="�ÿͽ�ӡͼƬ">
		</div>
	</div>
	<div class="set0">
		<div class="Set1">��ӡͼƬ��Ů����</div>
		<div class="Set2">
			<input type="text" size="20" name="Setting13" value="{$Setting13}" id="Setting13">
			<select onchange="Chang(this.value,'DiaryFootMart2','boke/images/sex/','Setting13')">
				<option value="">ѡȡ��ӡͼ��
				<option value="sex3.gif" >sex3.gif</option>
				<option value="sex4.gif" >sex4.gif</option>
				<option value="footmark2.gif" >footmark2.gif</option>
				<option value="footmark4.gif" >footmark4.gif</option>
				<option value="footmark6.gif" >footmark6.gif</option>
			</select>
			&nbsp;<img id="DiaryFootMart2" src="{$Setting13}" border="0" alt="�ÿͽ�ӡͼƬ">
		</div>
	</div>
	<div class="set0">
		<div class="Set1">���ͽ��׹��ܣ�</div>
		<div class="Set2"><input type="radio" name="Setting14" value="0"/> �ر� <input type="radio" name="Setting14" value="1"/> ����</div>
	</div>
	<div class="set0">
		<div class="Set1">����ģ���������֣�</div>
		<div class="Set2"><input type="text" size="20" name="Setting15" value="{$Setting15}"/> * ��Dopod����ר��</div>
	</div>
	<br />
	<br />
	<div class="set0">
		<div class="Set1"><input type="submit" name="submit" value="�޸�����"/></div>
	</div>
</div>
</form>
<SCRIPT LANGUAGE="JavaScript">
<!--
	Dv_Form.Set_Radio('Setting0','{$Setting0}');
	Dv_Form.Set_Radio('Setting1','{$Setting1}');
	Dv_Form.Set_Radio('Setting4','{$Setting4}');
	Dv_Form.Set_Radio('Setting5','{$Setting5}');
	Dv_Form.Set_Radio('Setting14','{$Setting14}');
	function Chang(picurl,V,S,F)
	{
		var pic=S+picurl
		if (picurl!=''){
		document.getElementById(V).src=pic;
		setting[F].value=pic;
		}
	}
//-->
</SCRIPT>
]]>
</xsl:variable>
<xsl:variable name="String_10" title="Main_UserSetting_KeyWord">
<![CDATA[
<div class="topic">
���͹ؼ�������
</div>
<div style="margin-left:20px;">
<BR />
˵�����������з�������»����ۺ����������õĹؼ�������Զ������滻�ͼ������ӣ������Ĺؼ��ֶ�֮ǰ�����������Ч�����ɱ༭ĳƪ�����Ա���֧�����������õĹؼ���
<BR />
<BR />
<FORM METHOD=POST ACTION="?s=5&t=4&Action=Save">
	<input type="hidden" value="{$KeyID}" name="KeyID"/>
	<li>��&nbsp;&nbsp;��&nbsp;�֣�<input type="text" name="KeyWord" value="{$KeyWord}"/></li>
	<li>�滻�ı���<input type="text" name="nKeyWord" value="{$nKeyWord}"/></li>
	<li>���ӵ�ַ��<input type="text" name="LinkUrl" value="{$LinkUrl}"/><input type="checkbox" name="NewWindows" value="1" {$NewWindows}/>�´��ڴ�</li>
	<li>���ӱ��⣺<input type="text" name="LinkTitle" value="{$LinkTitle}"/><input type="submit" name="submit" value="�ύ�ؼ���"/></li>
</FORM>
<div id="mainlist">
<hr style="clear:both;"/>
<ul>
<li class="Set33"><B>�ؼ���</B></li>
<li class="Set33"><B>�滻�ı�</B></li>
<li class="Set44"><B>���ӵ�ַ</B></li>
<li class="Set33"><B>�������</B></li>
</ul>

<hr style="clear:both;"/>
{$KeyWordList}
</div>
</div>
<SCRIPT LANGUAGE="JavaScript">
<!--
	function alertreadme(str,url){
	{
		if(confirm(str)){
		location.href=url;
		return true;
		}return false;}
	}
//-->
</SCRIPT>
]]>
</xsl:variable>
<xsl:variable name="String_11" title="Main_UserInput_Topic">
<![CDATA[
]]>
</xsl:variable>
<xsl:variable name="String_12" title="Main_UserInput_Cat">
<![CDATA[
<script language="javascript" src="Boke/Script/Dv_form.js" type="text/javascript"></script>
<div class="topic">
{$UserAction}��Ŀ����
</div>
<BR />
<div style="margin-left:20px;">
<FORM METHOD=POST ACTION="{$UserAction_1}&m=2&Action=Save">
<input type="hidden" value="{$uCatID}" name="uCatID"/>
<li>��Ŀ���
<Select Name="sType" Size="1">
<Option value="0">����</Option>
<Option value="1">�ղ�</Option>
<Option value="2">����</Option>
<Option value="3">����</Option>
<Option value="4">���</Option>
</Select>
</li>
<li>��Ŀ���ƣ�
	<input type="text" name="uCatTitle" value="{$uCatTitle}"/>
	<input type="checkbox" name="IsView" value="1" {$IsView}/>
	�Ƿ񿪷�
</li>
<li>��Ŀ˵����
	<input type="text" name="uCatNote" value="{$uCatNote}"/>
	<input type="submit" name="submit" value="�ύ��Ŀ����">
</li>
</FORM>
</div>
<div id="mainlist">
<hr style="clear:both;">
<ul>
<li class="Set3"><B>��Ŀ</B></li>
<li class="Set5"><B>˵��</B></li>
<li class="Set3"><B>�������</B></li>
</ul>
<hr style="clear:both;">
{$InfoList}
</div>
<SCRIPT LANGUAGE="JavaScript">
<!--
	Dv_Form.Set_Select('sType','{$uType}');
	function alertreadme(str,url){
	{
		if(confirm(str)){
		location.href=url;
		return true;
		}return false;}
	}
//-->
</SCRIPT>
]]>
</xsl:variable>
<xsl:variable name="String_13" title="Main_UserInput_mTopic">
<![CDATA[
<script language="JavaScript" src="boke/Script/Pagination.js"></script>
<script type="text/javascript" language="javascript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>
<div class="topic">
	{$UserAction}����
</div>
<div style="margin:10px 0px 10px 0px;">
<FORM METHOD=POST ACTION="?s=1&t={$t}&m=3">
	�ؼ��֣�<input type="text" size="15" name="KeyWord" value="{$KeyWord}"/>
	<input type="submit" value="��ѯ" name="submit"/>
</FORM>
</div>
<div class="list_top">
	<div class="list1">
		<B>ѡ��</B>
	</div>
	<div class="list2">
		<B>����</B>
	</div>
	<div class="list3">
		<B>���</B>
	</div>
	<div class="list4">
		<B>ʱ��</B>
	</div>
	<div class="list3">
		<B>����</B>
	</div>
</div>
<br/>
<FORM METHOD=POST ACTION="?s=1&t={$t}&m=3&Action=Del">
{$InfoList}
	<div class="list_body" style="margin-bottom:10px;">
		<hr class="post" />
		<div class="list1">
			<input type="checkbox" name="chkall" value="on" onclick="CheckAll(this.form);"/>
		</div>
		<div class="list3" style="text-align:left;">
			ȫѡ
		</div>
		<div class="list4" style="text-align:left;">
			<input type="radio" value="0" name="iTopic" checked>
			ɾ��
			<input type="radio" value="1" name="iTopic">
			�ƶ���
		</div>
		<div class="list2" style="text-align:left;">
			<Select Name="uCatID" Size="1">
				<Option value="-1">��ѡ����Ŀ</Option>
				<Option value="0">δ����</Option>
				{$uCatList}
			</Select>
		</div>
		<div class="list3" style="text-align:left;">
			<input type="submit" name="action" value="ִ�в���" onclick="if(confirm('ȷ��ִ��ѡ���Ĳ�����?')==false)return false;" />
		</div>
	</div>
</FORM>
<div style="clear:both;">
<SCRIPT language="JavaScript">
	PageList({$Page},3,{$MaxRows},{$CountNum},"{$PageSearch}",1);
</SCRIPT>
</div>
]]>
</xsl:variable>

<xsl:variable name="String_14" title="Main_Other">
<![CDATA[
]]>
</xsl:variable>

<xsl:variable name="String_15" title="Main_Other">
<![CDATA[
]]>
</xsl:variable>
<xsl:variable name="String_16" title="Main_Other">
<![CDATA[
]]>
</xsl:variable>
<xsl:variable name="String_17" title="Main_Other">
<![CDATA[
]]>
</xsl:variable>
<xsl:variable name="String_18" title="Main_SkinSetting">
<![CDATA[
<script language="javascript" src="Boke/Script/Dv_form.js" type="text/javascript"></script>
<script language="JavaScript" src="boke/Script/Pagination.js"></script>
<div id="main" style="text-align:left;">
<form method="post" action="?s=4&action=Saveskin" name="Skinsform">
<table border="0" cellpadding="3" cellspacing="0" align="center" Style="width:98%">
<tr><td><input type="submit" value="�ύ����"></td></tr>
</table>
<hr/>
<table border="0" cellpadding="3" cellspacing="0" align="center" Style="width:98%">
<tr>
{$skin_list}
</tr>
</table>
</form>
<div style="clear:both;">
<SCRIPT language="JavaScript">
PageList({$Page},3,{$MaxRows},{$CountNum},"{$PageSearch}",1);
Dv_Form.Set_Radio("Skinid","{$S_id}");
</SCRIPT>
</div>
</div>
]]>
</xsl:variable>
<xsl:variable name="String_19" title="Main_SkinSetting_list1">
<![CDATA[
<td align="center" width="25%">
<img src="{$ViewLogo}" alt="" border="0px" width="80px" height="60px"/><br/>
{$S_Name}
<input type="radio" name="Skinid" value="{$S_id}"/>
</td>
]]>
</xsl:variable>
<xsl:variable name="String_20" title="Main_SkinSetting_list2">
<![CDATA[
</tr><tr>
]]>
</xsl:variable>
<xsl:variable name="String_21" title="Main_UserInput_mTopic_list2">
<![CDATA[
<div class="list_body">
<div class="list1">
<input type="checkbox" name="TopicID" value="{$EditID}"/>
</div>
<div class="list2" style="text-align:left;">
<a href="{$bokeurl}{$bokename}.showtopic.{$topicid}.html" target="_blank">{$Topic}</a>
</div>
<div class="list3">
{$cat}
</div>
<div class="list4" style="text-align:left;">
{$DateTime}
</div>
<div class="list3">
<a href="Bokepostings.asp?user={$bokename}&action=edit&rootid={$topicid}&postid={$postid}">�༭</a> | <a href="?s=1&t={$t}&m=3&Action=Del&TopicID={$topicid}&iTopic=0">ɾ��</a>
</div>
</div>
]]>
</xsl:variable>
<xsl:variable name="String_22" title="Main_UserInput_mPost">
<![CDATA[
<script language="JavaScript" src="boke/Script/Pagination.js"></script>
<script type="text/javascript" language="javascript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>
<div id="main" style="text-align:left;">
	<div class="topic">���۹���</div>
	<div style="margin:10px 0px 10px 0px;">
	<FORM METHOD=POST ACTION="?s=2">
		�ؼ��֣�<input type="text" size="15" name="KeyWord" value="{$KeyWord}"/>
		<input type="submit" value="��ѯ" name="submit"/><br/>
	</FORM>
</div>

<div class="list_top">
	<div class="list1">
		<B>ѡ��</B>
	</div>
	<div class="list2">
		<B>����</B>
	</div>
	<div class="list3">
		<B>����</B>
	</div>
	<div class="list4">
		<B>ʱ��</B>
	</div>
	<div class="list3">
		<B>����</B>
	</div>
</div>

<FORM METHOD=POST ACTION="?s=2&Action=Del">
{$InfoList}
<div class="list_body" style="margin-bottom:10px;">
	<hr class="post">
		<div class="list1">
			<input type="checkbox" name="chkall" value="on" onclick="CheckAll(this.form);"/>
		</div>
		<div class="list3" style="text-align:left;">ȫѡ	</div>
		<div class="list2" style="text-align:left;"><input type="radio" value="0" name="iTopic" checked />ɾ��</div>
		<div class="list3" style="text-align:left;">
			<input type="submit" name="action" value="ִ�в���" onclick="if(confirm('ȷ��ִ��ѡ���Ĳ�����?')==false)return false;" />
		</div>
</div>
</FORM>
<div style="clear:both;">
<SCRIPT language="JavaScript">
PageList({$Page},3,{$MaxRows},{$CountNum},"{$PageSearch}",1);
</SCRIPT>
</div>
</div>
]]>
</xsl:variable>

<xsl:variable name="String_23" title="Main_UserInput_mPost_list2">
<![CDATA[
<div class="list_body">
	<div class="list1">
		<input type="checkbox" name="TopicID" value="{$EditID}"/>
	</div>
	<div class="list2" style="text-align:left;">
		<a href="{$bokeurl}{$bokename}.showtopic.{$topicid}.{$postid}.html" target="_blank">{$Topic}</a>
	</div>
	<div class="list3">
		{$cat}
	</div>
	<div class="list4" style="text-align:left;">
		{$DateTime}
	</div>
	<div class="list3">
		<a href="Bokepostings.asp?user={$bokename}&action=edit&rootid={$topicid}&postid={$postid}">�༭</a> | <a href="?s=2&Action=Del&TopicID={$postid}&iTopic=0">ɾ��</a>
	</div>
</div>
]]>
</xsl:variable>
<xsl:variable name="String_24" title="Main_UserFile">
<![CDATA[
<script language="JavaScript" src="boke/Script/Pagination.js"></script>
<div id="main" style="text-align:left;">
�ļ�����&nbsp;&nbsp;��&nbsp;&nbsp;<a href="?s=3&m=1">�����ļ�</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="?s=3&m=2">ͼƬ</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="?s=3&m=3">ѹ��</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="?s=3&m=4">�ĵ�</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="?s=3&m=5">ý��</a>
<BR />
<BR />
<div class="topic">
{$ActionInfo}�ļ�����
</div>
<div style="margin:10px 0px 10px 0px;">
<FORM METHOD=POST ACTION="?s=3">
�ؼ��֣�<input type="text" size="15" name="KeyWord" value="{$KeyWord}"/>
<input type="submit" value="��ѯ" name="submit"/>
</FORM>
</div>
<hr class="post">
<FORM METHOD=POST ACTION="?s=3&Action=Del">
{$topiclist}
</FORM>
<SCRIPT language="JavaScript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</SCRIPT>
</div>
]]>
</xsl:variable>
<xsl:variable name="String_25" title="cat_photo">
<![CDATA[
<table border="0" cellpadding="3" cellspacing="0" align="center" Style="width:98%">
	<tr>
		{$photo_list}
	</tr>
</table>
<div class="list_body">
	<hr class="post">
		<div class="list1">
			<input type="checkbox" name="chkall" value="on" onclick="CheckAll(this.form);"/>
		</div>
		<div class="list3" style="text-align:left;">
			ȫѡ
		</div>
		<div class="list2" style="text-align:left;">
			<input type="radio" value="0" name="iTopic" checked/>
			ɾ��
		</div>
		<div class="list3" style="text-align:left;">
			<input type="submit" name="action" value="ִ�в���" onclick="if(confirm('ȷ��ִ��ѡ���Ĳ�����?')==false)return false;" />
		</div>
</div>
<div style="clear:both;">
<SCRIPT LANGUAGE="JavaScript">
<!--
PageList({$Page},3,{$MaxRows},{$CountNum},"{$PageSearch}",1);
//-->
</SCRIPT>
</div>
]]>
</xsl:variable>

<xsl:variable name="String_26" title="cat_photo">
<![CDATA[
<td align="center" width="25%">
<div class="photo1">
<a href="{$ViewPhoto}" target="_blank"><img src="{$ViewPhoto}" alt="����鿴�������ļ�" Style="border:1px solid #C0C0C0" width="{$width}px" height="{$height}px"/></a>
<div class="photo2">
�� {$topic} ��
<br/><input type="checkbox" name="fileid" value="{$fileid}">
{$PostDate} -- <a href="?s=3&Action=Del&fileid={$fileid}&iTopic=0">ɾ��</a>
</div>
</div>
</td>
]]>
</xsl:variable>
<xsl:variable name="String_27" title="cat_photo">
<![CDATA[
</tr><tr>
]]>
</xsl:variable>
<xsl:variable name="String_28" title="sys_count">
<![CDATA[
<div id="main" style="text-align:left;">
<div class="topic">����վ������ͳ�ƺ͸���</div>
<br/>
	<fieldset title="��ʾ">
	<legend>&nbsp;��ʾ&nbsp;</legend>
	���²�������Ҫ����ִ�У�
	</fieldset>
<br/>
<table border="0" width="98%" align="center" cellpadding="4" cellspacing="2" class="table3">
<tr><td width="10%" class="td1"><input type="button" name="act" value="���͸��˱�������" onclick="location.href='?s=7&t=1'"/></td><td align="left" width="90%" class="td2">�����������Ϣ����Ϊ��ǰʹ�õı�����</td></tr>
<tr><td class="td1"><input type="button" name="act" value="���·�����������" onclick="location.href='?s=7&t=2'"/></td><td class="td2">����ͳ�Ʒ�������������</td></tr>
<tr><td class="td1"><input type="button" name="act" value="���²���������" onclick="location.href='?s=7&t=3'"/></td><td class="td2">���¸��˲�����ͳ������</td></tr>
</table>
</div>
]]>
</xsl:variable>
</xsl:stylesheet>