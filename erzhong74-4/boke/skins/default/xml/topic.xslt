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
<xsl:variable name="String_0" title="topic">
<![CDATA[
<script language="javascript" src="Boke/Script/Dv_form.js" type="text/javascript"></script>
<script type="text/javascript" src="boke/Edit_Plus/FCKeditor/fckeditor.js"></script>
<script language="JavaScript" src="boke/Script/Pagination.js"></script>
<center>
{$nav}
<div class="posthead">
<div class="topic_r">
	<a href="{$bokeurl}{$bokename}.showchannel.{$cat_tid}.{$cat_id}.html">{$catname}</a> | ����({$child}) | �Ķ�({$hits})
</div>
<img src="{$Skins_Path}images/plus.gif" border="0" onclick="ShowList(this,'VisitList','{$Skins_Path}images/plus.gif','{$Skins_Path}images/nofollow.gif');" class="ImgOnclick" alt="�鿴��ϸ�����б�"/>{$Vsisitstr}
<div id="VisitList" style="display:none;">
{$VisitList}
</div>
</div>
<br/>
<div class="topic">
	<div class="topic_r">
	{$WeekName} &nbsp;&nbsp;<img src="boke/images/weather/{$Weather_B}" alt="{$Weather_A}" align="absmiddle">&nbsp;
	</div>
<div><img src="{$Skins_Path}images/topic.gif" border="0" alt="����">&nbsp;{$title}</div>
</div>
{$DispList}
{$replyform}
</center>
]]>
</xsl:variable>
<xsl:variable name="String_1" title="pagesearch">
<![CDATA[
<SCRIPT language="JavaScript">
PageList({$Page},3,{$MaxRows},{$CountNum},"{$PageSearch}",5);
</SCRIPT>
]]>
</xsl:variable>
<xsl:variable name="String_2" title="displist_1">
<![CDATA[
<a name="{$PostID}"></a>
<div class="postbody">
{$ReTitle}
<br/>
<span id="ShowBody" style="font-size:{$fontsize}px;line-height: {$lineheight}px;">
{$Content}
</span>
</div>
<div class="postend">
<a href="dispuser.asp?id={$UserID}" target="_blank" alt="�鿴��������">{$PostUserName}</a> �����ڣ�{$PostDate}
{$Edit}{$Manage}{$Delete}
</div>
<hr class="post"/>
]]>
</xsl:variable>
<xsl:variable name="String_3" title="Edit">
<![CDATA[
[<a href="Bokepostings.asp?user={$bokename}&action=edit&rootid={$RootID}&postid={$PostID}" alt="�༭">�༭</a>]
]]>
</xsl:variable>
<xsl:variable name="String_4" title="Manage">
<![CDATA[
[<a href="Bokepostings.asp?user={$bokename}&action=isbest&rootid={$RootID}&postid={$PostID}" alt="��������">����</a>]
]]>
</xsl:variable>
<xsl:variable name="String_5" title="delete">
<![CDATA[
[<a href="Bokepostings.asp?user={$bokename}&action=delete&rootid={$RootID}&postid={$PostID}" alt="ɾ��">ɾ��</a>]
[<a href="Bokepostings.asp?user={$bokename}&action=reply&rootid={$RootID}&postid={$PostID}" alt="�ظ�">�ظ�</a>]
]]>
</xsl:variable>
<xsl:variable name="String_6" title="visit_str">
<![CDATA[
��û�зÿ����½�ӡ��
]]>
</xsl:variable>
<xsl:variable name="String_7" title="visit_str">
<![CDATA[
�������ķ���ӡ�� <a href="Bokepostings.asp?user={$bokename}&action=visit&rootid={$RootID}" title="��һ�ţ��ڴ˴������ҵķ���ӡ��"><img src="{$VisitPic}" border="0"/></a>
]]>
</xsl:variable>
<xsl:variable name="String_8" title="visit_str">
<![CDATA[
���ѽ�ӡ
]]>
</xsl:variable>
<xsl:variable name="String_9" title="fastreply">
<![CDATA[
<hr class="post"/>
<script language="JavaScript">
<!--
window.onload = function()
{
	var oFCKeditor = new FCKeditor( 'PostContent','100%','200px','Basic' ) ;
	oFCKeditor.BasePath	= "boke/Edit_Plus/FCKeditor/" ;
	oFCKeditor.Config['SkinPath'] = "skins/office2003/"
	oFCKeditor.ReplaceTextarea() ;
}
//-->
</script>
<center>
<div class="reply">
<table border="0" width="100%" align="center" cellpadding="4" cellspacing="2" class="table1">
<FORM METHOD="POST" ACTION="Bokepostings.asp?user={$bokename}&action=save_reply">
	<tr>
		<td height="20" colspan="2" class="reply_t1">����/�ظ�:</td>
	</tr>
	<tr>
		<td width="20%" class="reply_t2">�û�����</td>
		<td width="80%" class="reply_t3"><input type="text" name="PostUserName" value="{$PostUserName}" /></td>
	</tr>
	{$getcode}
	<tr>
		<td class="reply_t2">���⣺</td>
		<td class="reply_t3">
		<input type="text" name="Title" size="50" onkeyup="javascript:Dv_Form.CheckLength(this,250);" value=""/>
		{$isalipay}
		</td>
	</tr>
	<tr>
		<td  class="reply_t2">���ݣ�</td>
		<td  class="reply_t3">
		<textarea rows="4" name="PostContent" id="PostContent"></textarea>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="reply_t1" align="center" height="25px">
		<input type="hidden" name="RootID" value="{$RootID}"/>
		<input type="hidden" name="PostID" value="0"/>
		<input type="submit" value="�ύ��Ϣ"/>
		</td>
	</tr>
</FORM>
</table>
</div>
</center>
]]>
</xsl:variable>

<xsl:variable name="String_10" title="edit">
<![CDATA[
<hr class="post"/>
<script language="javascript" src="Boke/Script/Dv_form.js" type="text/javascript"></script>
<script language="javascript" src="Boke/Script/Sel_calendar.js" type="text/javascript"></script>
<script type="text/javascript" src="boke/Edit_Plus/FCKeditor/fckeditor.js"></script>
<center>
<div class="reply">
<table border="0" width="100%" align="center" cellpadding="4" cellspacing="2" class="table1">
<FORM METHOD="POST" ACTION="{$action}" name="PostForm" onSubmit="return CheckPostForm(this);">
	<tr>
		<td height="20" colspan="2" class="reply_t1">����/����/�ظ�:</td>
	</tr>
	<tr>
		<td width="20%" class="reply_t2">�û�����</td>
		<td width="80%" class="reply_t3"><input type="text" name="PostUserName" value="{$PostUserName}"/>
		{$isalipay}</td>
	</tr>
	{$getcode}
	<tr>
		<td class="reply_t2">���⣺</td>
		<td class="reply_t3">
		<input type="text" name="Title" size="50" onkeyup="javascript:Dv_Form.CheckLength(this,250);" value="{$Title}"/>
		���ڣ�<INPUT TYPE="text" NAME="DDateTime" value="{$PostData}" onfocus="show_cele_date(DDateTime,'','',DDateTime)"/>
		</td>
	</tr>
	{$TopicFunction1}
	{$Show_Payto}
	{$Show_Upload}
	{$Show_PostContent}
	{$TopicFunction2}
	{$TopicFunction3}
	<tr>
		<td colspan="2" class="reply_t1" align="center" height="25px">
		<input type="hidden" name="RootID" value="{$RootID}"/>
		<input type="hidden" name="PostID" value="{$PostID}"/>
		<input type="submit" name="submit" value="�ύ��Ϣ"/>
		</td>
	</tr>
</FORM>
</table>
</div>
</center>
<br/>
<script language="JavaScript">
<!--
function Chang(picurl,V,S)
	{
		picurl=picurl.split("|");
		picurl=picurl[0];
		var pic=S+picurl
		if (picurl!=''){
		document.getElementById(V).src=pic;
		}
	}
//-->
</script>
]]>
</xsl:variable>

<xsl:variable name="String_11" title="edit_AboutTopic1">
<![CDATA[
	<tr>
		<td class="reply_t2">ѡȡ���ࣺ</td>
		<td class="reply_t3">
		<Select Name="sType" Size="1" onchange="OchangeSel_Cat('Catid',this.value);">
		<Option value="-1">ѡ�����</Option>
		<Option value="0">����</Option>
		<Option value="1">�ղ�</Option>
		<Option value="2">����</Option>
		<Option value="3">����</Option>
		<Option value="4">���</Option>
		</Select>
		<Select Name="sCatID" Size="1">
		<Option value="-1">ѡ����</Option>
		{$sCatList}
		</Select>
		<Select Name="Catid" ID="Catid" Size="1">
		<Option value="-1">������Ŀ</Option>
		</Select>
		</td>
	</tr>
]]>
</xsl:variable>
<xsl:variable name="String_12" title="edit_AboutTopic2">
<![CDATA[
{$TitleNote}
<SCRIPT language="JavaScript">
	Dv_Form.Set_Select("sType","{$stype}");
	Dv_Form.Set_Select("Lock","{$Lock}");
	Dv_Form.Set_Select("Best","{$Best}");
	Dv_Form.Weather_Select("Weather","{$Weather}","WeatherSrc","boke/images/weather/");

	OchangeSel_Cat("Catid","{$stype}","{$Cat_Val}");
</SCRIPT>
]]>
</xsl:variable>
<xsl:variable name="String_13" title="edit_Aboutupload">
<![CDATA[
	<tr>
		<td class="reply_t2">�ļ��ϴ���<input type="hidden" name="upfilerename" value=""/></td>
		<td class="reply_t3">
		<iframe name="UPLOADFRAME" frameborder="0" width="100%" height="24" scrolling="no" src="BokeUpload.asp?mode=UploadForm"></iframe></td>
	</tr>
]]>
</xsl:variable>
<xsl:variable name="String_14" title="edit_Aboutupload">
<![CDATA[
<div class="visituser">{$VisitName}<br/><img src="{$VisitPic}" border="0" alt="��ԴIP��{$VisitIP};����ʱ�䣺{$VisitTime}"/></div>
]]>
</xsl:variable>
<xsl:variable name="String_15" title="edit_link">
<![CDATA[
	<tr>
		<td class="reply_t2">���ӵ�ַ��</td>
		<td class="reply_t3"><input type="text" name="PostContent" value="{$Content}" size="60"/></td>
	</tr>
]]>
</xsl:variable>
<xsl:variable name="String_16" title="edit_body">
<![CDATA[
	<tr>
		<td class="reply_t2">���ݣ�</td>
		<td class="reply_t3">
		<textarea rows="4" name="PostContent" id="PostContent">{$Content}</textarea>		
		UBB��ǩ��<a href="javascript:onclick=PlayUbb('PostContent','FLASH');" alt="��� FLASH ����">[FLASH]</a>
		<a href="javascript:onclick=PlayUbb('PostContent','RM');" alt="��� RM ý���ļ�">[RM]</a>
		<a href="javascript:onclick=PlayUbb('PostContent','MP');" alt="��� MEDIA PLAYER ý���ļ�">[MP]</a>
		<a href="javascript:onclick=PlayUbb('PostContent','QT');" alt="">[QT]</a>
		<a href="javascript:onclick=put_ubb('PostContent','CODE');" alt="Դ������">[CODE]</a>
		</td>
	</tr>
<script language="JavaScript">
<!--
	window.onload = function()
{
	var oFCKeditor = new FCKeditor( 'PostContent','100%','350px','{$Editmode}' ) ;
	oFCKeditor.BasePath	= "boke/Edit_Plus/FCKeditor/" ;
	oFCKeditor.Config['SkinPath'] = "skins/office2003/"
	oFCKeditor.ReplaceTextarea() ;
}
//-->
</script>

]]>
</xsl:variable>
<xsl:variable name="String_17" title="edit_titlenote">
<![CDATA[
	<tr>
		<td  class="reply_t2">����ժҪ��</td>
		<td  class="reply_t3">[<a href="javascript:onclick=GetNode('PostContent','PostTitleNote',{$maxlen});" alt="��������ȡժҪ">��ȡժҪ</a>]<br/>
		<textarea style="width:95%;height:50px;" name="PostTitleNote" id="PostTitleNote">{$PostTitleNote}</textarea>
		</td>
	</tr>
]]>
</xsl:variable>
<xsl:variable name="String_18" title="edit_titlenote">
<![CDATA[
	<tr>
		<td  class="reply_t2">����ͼƬ��</td>
		<td  class="reply_t3">
		<input type="text" name="PostTitleNote" value="{$PostTitleNote}" size="60" onkeyup="javascript:Dv_Form.CheckLength(this,250);"/>
		</td>
	</tr>
]]>
</xsl:variable>

<xsl:variable name="String_19" title="displist_2">
<![CDATA[
<div class="retopic"><img src="{$Skins_Path}images/reply.gif" border="0" alt="����">&nbsp;{$RTitle}</div>
]]>
</xsl:variable>
<xsl:variable name="String_20" title="edit_other">
<![CDATA[
]]>
</xsl:variable>
<xsl:variable name="String_21" title="edit_other">
<![CDATA[
]]>
</xsl:variable>
<xsl:variable name="String_22" title="edit_other">
<![CDATA[
]]>
</xsl:variable>
<xsl:variable name="String_23" title="edit_other">
<![CDATA[
]]>
</xsl:variable>
<xsl:variable name="String_24" title="edit_other">
<![CDATA[
]]>
</xsl:variable>
<xsl:variable name="String_25" title="edit_other">
<![CDATA[
]]>
</xsl:variable>
<xsl:variable name="String_26" title="edit_upload">
<![CDATA[
<BODY oncontextmenu="return false;" onselectstart="return false;">
<script>
var upset = {$upset};
if (top.location==self.location){
	top.location="boke.asp"
}
function uploadframe(num)
{
	var obj=parent.document.getElementById("UPLOADFRAME");
	if (obj){
	if (parseInt(obj.height)+num>=24) {
		obj.height = parseInt(obj.height) + num;
		document.getElementById("Setupload").style.display="";
		document.getElementById("allupload").style.display="none";
	}
	}
}
function setid()
{
	str='';
	if(!window.form.upcount.value)
	window.form.upcount.value=1;
	if(window.form.upcount.value>upset){
	alert("�����ֻ��ͬʱ�ϴ�"+upset+"���ļ�!");
	window.form.upcount.value = upset;
	setid();
	}
	else{
	for(i=1;i<=window.form.upcount.value;i++)
	str+='�ļ�'+i+':<input type="file" name="file'+i+'" style="width:200px"/><br/>';
	window.upid.innerHTML=str+'<br/>';
	var num=i*16
	var obj=parent.document.getElementById("UPLOADFRAME");
	if (parseInt(obj.height)+num>=24) {
		obj.height = 24 + num;	
	}
	}
}
</script>
<form name="form" method="post" action="Bokeupload.asp?user={$bokename}&mode=SaveUpload" enctype="multipart/form-data">
<table border="0" cellspacing="0" cellpadding="0" style="width:100%;height:100%">
<tr>
<Input type="hidden" name="act" value="upload"/>
<TD id="upid" valign="top" align="left" class="reply_t3">
<input type="file" name="file1" width="120px" value="" size="25"/></TD>
<td valign="top" width="*" align="left" class="reply_t3">
<input type="submit" name="Submit" value="�ϴ�" onclick="parent.document.PostForm.submit.disabled=true"/>
</td>
<td id="allupload" valign="top" align="left" class="reply_t3">
</td>
<td valign="top" align="left" class="reply_t3">
<div id="Setupload" style="display:none" align="left">
</div>
</td>
</tr>
</table>
</form>
<script language="JavaScript">
<!--
if (upset > 1){
var tempstr1='<input type="button" name="setload" onClick="uploadframe(0);" value="�����ϴ�"/>';
var tempstr2='���ø���<input type="text" value="1" name="upcount" size="1" maxlength="1"/><input type="button" name="Button" onClick="setid();" value="�趨"/>';
document.getElementById("allupload").innerHTML=tempstr1;
document.getElementById("Setupload").innerHTML=tempstr2;
}
//-->
</script>
</body>
</html>
]]>
</xsl:variable>
<xsl:variable name="String_27" title="edit_Saveupload">
<![CDATA[
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<td class="reply_t3" valign="top" height="40" align="left">
{$SucMsg}
</td>
</tr>
</table>
</body>
</html>
]]>
</xsl:variable>

<xsl:variable name="String_28" title="linklist">
<![CDATA[

<div class="topic1">
{$LinkeLogo}
���ƣ�<a href="{$bokeurl}{$bokename}.showtopic.{$TopicID}.html" target="_blank">{$topic}</a>
</div>
<div class="postend" style="text-align:left;">
<div class="topic_r">
<a href="dispuser.asp?id={$UserID}" target="_blank" alt="�鿴��������">{$PostUserName}</a> ����ڣ�{$PostDate}
{$Edit}{$Manage}{$Delete}
</div>
�����{$hits}
</div>
<hr class="post"/><br/>
]]>
</xsl:variable>
<xsl:variable name="String_29" title="linklist">
<![CDATA[
<div class="topic_r">
<a href="{$bokeurl}{$bokename}.showtopic.{$TopicID}.html" target="_blank"><img src="{$LinkImg}" border="0" alt="����ͼ��"/></a>
</div>
]]>
</xsl:variable>
<xsl:variable name="String_30" title="delete">
<![CDATA[
[<a href="Bokepostings.asp?user={$bokename}&action=delete&rootid={$RootID}&postid={$PostID}" alt="ɾ��">ɾ��</a>]
]]>
</xsl:variable>
<xsl:variable name="String_31" title="alipay">
<![CDATA[
&nbsp;<a href="javascript:Dvbbs_foralipay('PostContent');"><img src="boke/images/alipay/icon_alipay.gif" border="0" alt="����һ��֧�������ף�֧�����������ף���ѡ���ȫ�����٣�"></a>
]]>
</xsl:variable>
<xsl:variable name="String_32" title="alipay_a">
<![CDATA[
	<tr>
		<td width="20%" class="reply_t2">֧�����ʺţ�<BR/>����д��ЧEmail��ַ&nbsp;&nbsp;</td>
		<td width="80%" class="reply_t3"><input type="text" name="alipay_seller" size="20" value="{$paytomail}"/>
		δע��֧����������ҽ��׺���ռ����ʼ�</td>
	</tr>
	<tr>
		<td width="20%" class="reply_t2">��Ʒ���ƣ�<BR/><a href="http://server.dvbbs.net/dvbbs/alipay/s.htm" target="_blank"><font color="blue">֧�������װ���</font></a>&nbsp;&nbsp;</td>
		<td width="80%" class="reply_t3"><input type="text" name="alipay_subject" size="30"/>
		&nbsp;&nbsp;�۸�&nbsp;<input type="text" name="alipay_price" size="10" value="0"/> Ԫ
		</td>
	</tr>
	<tr>
		<td width="20%" class="reply_t2">�ʷѳе�����</td>
		<td width="80%" class="reply_t3">
			<input type=radio ID="ialipay_transport" name="ialipay_transport" value="3" onclick="document.getElementById('alipay_mail').disabled=true; document.getElementById('alipay_express').disabled=true;alipay_transport.value=3;"> ������Ʒ�����ʵ�&nbsp;&nbsp;
			<input type=radio ID="ialipay_transport" name="ialipay_transport" checked value="1" onclick="document.getElementById('alipay_mail').disabled=true; document.getElementById('alipay_express').disabled=true;alipay_transport.value=1;"> ���ҳе��˷�&nbsp;&nbsp;
			<input type=radio ID="ialipay_transport" name="ialipay_transport" value="2" onclick="document.getElementById('alipay_mail').disabled=false; document.getElementById('alipay_express').disabled=false;alipay_transport.value=2;"> ��ҳе��˷�<BR>
			<input type="hidden" value="1" ID="alipay_transport">
			ƽ��&nbsp;<input type=text ID="alipay_mail" name="alipay_mail" style="width:45px" size=25 disabled> Ԫ &nbsp;&nbsp;
			���&nbsp;<input type=text ID="alipay_express" name="alipay_express" style="width:45px" size=25 disabled> Ԫ (����Ϊƽ��)</td>
	</tr>
	<tr>
		<td width="20%" class="reply_t2">��ʾ��ַ��</td>
		<td width="80%" class="reply_t3"><input type="text" name="alipay_demo" size="30" value="http://"/>
		��Ʒ˵������д����Ϣ������</td>
	</tr>
	<tr>
		<td width="20%" class="reply_t2">��ϵ��ʽ��</td>
		<td width="80%" class="reply_t3">�Ա�����&nbsp;<input type="text" name="alipay_ww" size="10"/>&nbsp;&nbsp;��ѶQQ&nbsp;<input type="text" name="alipay_qq" size="9"/>
		�����������ϵ��</td>
	</tr>
]]>
</xsl:variable>
<xsl:variable name="String_33" title="topic_11">
<![CDATA[
	<tr>
		<td class="reply_t2">�����ؼ��֣�</td>
		<td class="reply_t3"><input type="text" name="SearchKey" size="50" onkeyup="javascript:Dv_Form.CheckLength(this,250);" value="{$SearchKey}"/>(���ܳ���250���ַ����ö��ŷָ�)
		</td>
	</tr>
	<tr>
		<td class="reply_t2">ѡȡ״̬��</td>
		<td class="reply_t3">
		<Select Name="Lock" Size="1">
		<Option value="0">����</Option>
		<Option value="1">��֤</Option>
		<Option value="2">�ر�</Option>
		<Option value="3">����</Option>
		</Select>
		<Select Name="Best" Size="1">
		<Option value="0">��ͨ</Option>
		<Option value="1">����</Option>
		</Select>
		<Select Name="Weather" Size="1" onchange="Chang(this.value,'WeatherSrc','boke/images/weather/')">{$WeatherList}
		</Select>
		<img id="WeatherSrc" src="boke/images/weather/sun.gif" border="0">
		</td>
	</tr>
]]>
</xsl:variable>
<xsl:variable name="String_34" title="getcode">
<![CDATA[
	<tr>
		<td class="reply_t2">��&nbsp;&nbsp;֤&nbsp;�룺</td>
		<td class="reply_t3">{$Dv_GetCode}<!--<input type="text" name="codestr" maxlength="4" size="4"/>&nbsp;<img name="codepic" src="DV_getcode.asp" onclick="this.src='DV_getcode.asp'" alt="���������֤�룡"/>--></td>
	</tr>
]]>
</xsl:variable>
</xsl:stylesheet>

