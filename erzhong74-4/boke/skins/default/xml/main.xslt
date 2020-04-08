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
<xsl:variable name="String_0" title="htmlhead">
<![CDATA[
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-CN" lang="zh-CN">
<head><title>{$boketitle}-{$bokename}{$stats}</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
<meta name="Copyright" content="{$copyright}"/>
<meta name="description" content="{$bokechildtitle}"/>
<link rel="alternate" type="application/rss+xml" title="����{$bokename}����Ϣ" href="{$rssurl}" />
<link rel="home" href="{$BokeUrl}" />
<link rel="SHORTCUT ICON" href="boke/images/favicon.ico"/>
<link rel="stylesheet" rev="stylesheet" href="{$skinpath}Css/BodyStyle.css" type="text/css" media="screen" />
<script language="javascript" src="Boke/Script/Dv_Main.js" type="text/javascript"></script>
<script language="javascript" src="Boke/Script/Dv_page_inc.js" type="text/javascript"></script>
</head>
<!-- <BODY oncontextmenu="return false;" onselectstart="return false;"> -->
<body>
]]>
</xsl:variable>
<xsl:variable name="String_1" title="Footer">
<![CDATA[
<div id="footer">
<div id="nav_r">
<img src="boke/Images/boke_logo.gif" border="0" alt=""/>
</div>
Power by {$version}
<br/>
Copyright &copy; {$bokeuser} {$sysinfo}
</div>
</body>
</html>
]]>
</xsl:variable>

<xsl:variable name="String_2" title="Top">
<![CDATA[
<div class="menuskin" id="popmenu" onmouseover="DvMenu.clearhidemenu()" onmouseout="DvMenu.dynamichide(event)" style="z-index:100"></div>
<div id="topnav">
<div style="float:right;">
{$SiteUrl}&nbsp;&nbsp;&nbsp;&nbsp;
</div>
<div>
<font style="font-weight : bold ;FONT-SIZE: 22px;line-height: normal;" face="Georgia,Times">
{$boketitle}</font><br/>
&nbsp;&nbsp;&nbsp;&nbsp;{$bokechildtitle}
</div>
</div>
]]>
</xsl:variable>

<xsl:variable name="String_3" title="TopNav">
<![CDATA[
<div id="navmenu">
<div id="nav_r">
{$TopicNum}ƪ���£�{$FavNum}ƪ�ղأ�������Ϣ{$TodayNum}ƪ,������:{$LastUpTime} {$Open}
</div>
<a href="bokeindex.asp">������ҳ</a>&nbsp;&nbsp;|&nbsp;&nbsp;
{$BokeOwnerNav}
<a href="index.asp">��̳</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="{$bokeurl}{$bokename}.showchannel.0.html" onMouseOver="showmenu(event,'','ShowChannel_0');">����Ƶ��</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="{$bokeurl}{$bokename}.showchannel.4.html" onMouseOver="showmenu(event,'','ShowChannel_4');">���Ƶ��</a>
{$BokePayto}
</div>
]]>
</xsl:variable>

<xsl:variable name="String_4" title="LeftMenu">
<![CDATA[
<div id="leftmenu">
{$show_bokenote}
{$show_calendar}
{$show_channel}
{$show_search}
{$show_topicnews}
{$show_photos}
{$show_counts}
{$show_votes}
{$show_links}
<br/>
<a href="http://rss.iboker.com/sub/?{$bokeurl_r}{$bokename}.rss.xml" target="_blank"><img src="boke/Images/xml.gif" border="0" alt="���ı�վ�� RSS 2.0 ������" /></a>
<br/>
<a href="http://validator.w3.org/check?uri=referer" target="_blank"><img src="boke/Images/w3c_xhtml.gif" border="0" alt=""/></a>
<br/>
<a href="http://jigsaw.w3.org/css-validator/check/referer" target="_blank"><img src="boke/Images/w3c_css.gif" border="0" alt=""/></a>
<br/>
<a href="http://www.iboker.com" target="_blank"><img src="boke/Images/boke_icon.gif" border="0" alt=""/></a>
</div>
]]>
</xsl:variable>

<xsl:variable name="String_5" title="Main">
<![CDATA[
<!-- ҳ������ -->
<div id="main">
{$Main}
</div>
]]>
</xsl:variable>

<xsl:variable name="String_6" title="topic">
<![CDATA[
<div class="topic">
	<div class="topic_r">
	{$PostDate} &nbsp;&nbsp;<img src="boke/images/weather/{$Weather_B}" alt="{$Weather_A}" align="absmiddle">&nbsp;
	</div>
<a href="{$bokeurl}{$bokename}.showtopic.{$TopicID}.html" target="_blank"><img src="{$Skins_Path}images/topic.gif" border="0" alt="���´��ڷ��ʸ�����"></a>&nbsp;<a href="{$bokeurl}{$bokename}.showtopic.{$TopicID}.html">{$topic}</a>
</div>
<div class="postbody">
{$Content}
</div>
<div class="postend"><a href="dispuser.asp?id={$UserID}" target="_blank" alt="�鿴��������">{$PostUserName}</a> ������ <a href="{$bokeurl}{$bokename}.showchannel.{$cat_tid}.{$cat_id}.html">{$PChannel}</a> | <a href="{$bokeurl}{$bokename}.showtopic.{$TopicID}.html">����({$Childs})</a> | ����(0) </div>
<hr class="post"/><br/>
]]>
</xsl:variable>

<xsl:variable name="String_7" title="show_bokenote">
<![CDATA[
<!-- ���¹��� -->
<div class="border1">
<div class="lefttitle">���¹���</div>
<div class="msg">
{$bokenote}
</div>
</div>
<br/>
]]>
</xsl:variable>
<xsl:variable name="String_8" title="show_channel">
<![CDATA[
<!-- Ƶ����Ŀ -->
<div class="border1">
<div class="lefttitle">Ƶ����Ŀ</div>
<div class="msg" id="show_channel" style="padding: 3 3 3 20;">
{$bokechannel}
</div>
</div>
<script language="JavaScript">
<!--
//0=����,1=�ղ�,2=����,3=����,4=���
Show_channeldata();

function Show_channeldata(){
	var Channels=["����","�ղ�","����","����","���"];
	var ChannelData="";
	var ChannelPay="{$ChannelPay}";
	for (var i=0;i<5;i++){
		//if (i!=2){
		var Obj = Dvbbs.Objects("ShowChannel_"+i);
		if (Obj){
			if (i==3 && ChannelPay!=''){
			ChannelData = "<b>"+ChannelPay+"</b>"+ Obj.innerHTML;
			}else{
			ChannelData = "<b>"+Channels[i]+"</b>"+ Obj.innerHTML;
			}
			Dvbbs.Objects("show_channel").innerHTML += ChannelData;
		};
		//}
	};
}
//-->
</script>
<br/>
]]>
</xsl:variable>
<xsl:variable name="String_9" title="show_search">
<![CDATA[
<!-- վ������ -->
<div class="border1">
<div class="lefttitle">��ѯ</div>
<form method="post" action="bokesearch.asp">
<div class="msg">
<input type="hidden" name="User" value="{$bokename}">
<input type="text" name="keyword" value="" size="25"/>
<br/>
<input type="radio" name="sel" value="0" checked="true">����
<input type="radio" name="sel" value="1">����
<select name="stype">
	<option value="-1">����</option>
	<option value="0" selected="true">���� </option>
	<option value="1">�ղ�</option>
	<option value="2">��ǩ</option>
	<option value="3">����</option>
	<option value="4">���</option>
</select>
<input type="submit" value="��ѯ"/>
</div>
</form>
</div>
<br/>
]]>
</xsl:variable>
<xsl:variable name="String_10" title="show_topicnews">
<![CDATA[
<!-- �������� -->
<div class="border1">
<div class="lefttitle">��������</div>
<div class="msg">
{$boketopicnews}
</div>
</div>
<br/>
]]>
</xsl:variable>

<xsl:variable name="String_11" title="show_photos">
<![CDATA[
<!-- ����ͼ�� -->
<div class="border1">
<div class="lefttitle">����ͼ��</div>
<div class="msg">
{$bokephotos}
</div>
</div>
<br/>
]]>
</xsl:variable>
<xsl:variable name="String_12" title="show_links">
<![CDATA[
<!-- �������� -->
<div class="border1">
<div class="lefttitle">��������</div>
<div class="msg">
{$bokelinks}
</div>
</div>
<br/>
]]>
</xsl:variable>

<xsl:variable name="String_13" title="show_counts">
<![CDATA[
<!-- ����ͳ�� -->
<div class="border1">
<div class="lefttitle">����ͳ��</div>
<div class="msg">
{$bokecounts}
</div>
</div>
<br/>
]]>
</xsl:variable>

<xsl:variable name="String_14" title="show_votes">
<![CDATA[
<!-- ����ͶƱ -->
<!--
<div class="border1">
<div class="lefttitle">����ͶƱ</div>
<div class="msg">
{$bokevotes}
</div>
</div>
<br/>
-->
]]>
</xsl:variable>

<xsl:variable name="String_15" title="show_calendar">
<![CDATA[
<!-- ���� -->
<!--
<div class="border1">
<div class="msg">
{$bokecalendar}
</div>
</div>
<br/>
-->
]]>
</xsl:variable>

<xsl:variable name="String_16" title="show_channel">
<![CDATA[
<!--����Ƶ���б�-->
<div class="Menu_popup" id="ShowChannel_{$cat_tid}">
<div class="menuitems">
{$channellist}
</div>
</div>
]]>
</xsl:variable>
<xsl:variable name="String_17" title="show_channellist">
<![CDATA[
<li><a href="{$bokeurl}{$bokename}.showchannel.{$cat_tid}.{$cat_id}.html" alt="">{$channelname}</a></li>
]]>
</xsl:variable>

<xsl:variable name="String_18" title="show_Topiclist">
<![CDATA[
<script language="JavaScript" src="boke/Script/Pagination.js"></script>
{$nav}
{$topiclist}
<SCRIPT language="JavaScript">
PageList({$Page},3,{$MaxRows},{$CountNum},"{$PageSearch}",5);
</SCRIPT>
]]>
</xsl:variable>
<xsl:variable name="String_19" title="channel_nav">
<![CDATA[
<div class="channel_nav">
<a href="{$bokeurl}{$bokename}.index.html" alt="����{$bokeuser}��ҳ">{$bokeuser}��ҳ</a> >> <a href="{$bokeurl}{$bokename}.showchannel.{$cat_tid}.html">{$stype}</a> >> <a href="{$bokeurl}{$bokename}.showchannel.{$cat_tid}.{$cat_id}.html">{$catname}</a> >> {$stats}
</div>
<br/>
<div class="channel_intro">{$Channel_Intro}</div>
]]>
</xsl:variable>
<xsl:variable name="String_20" title="cat_photo">
<![CDATA[
<table border="0" cellpadding="3" cellspacing="0" align="center" Style="width:98%">
<tr>
{$photo_list}
</tr>
</table>
]]>
</xsl:variable>

<xsl:variable name="String_21" title="cat_photo">
<![CDATA[
<td align="center" width="25%">
<div class="photo1">
<a href="{$bokeurl}{$bokename}.showtopic.{$TopicID}.html"><img src="{$ViewPhoto}" alt="" Style="border:1px solid #C0C0C0" width="{$width}px" height="{$height}px"/></a>
<div class="photo2">
�� {$topic} ��
<br/>{$PostUserName} -- {$PostDate}
</div>
</div>
</td>
]]>
</xsl:variable>
<xsl:variable name="String_22" title="cat_photo">
<![CDATA[
</tr><tr>
]]>
</xsl:variable>
<xsl:variable name="String_23" title="boke_link">
<![CDATA[
<li><a href="{$linkurl}" target="_blank">{$linkname}</a></li>
]]>
</xsl:variable>
<xsl:variable name="String_24" title="boke_photo">
<![CDATA[
<div class="photo1">
<a href="{$bokeurl}{$bokename}.showtopic.{$TopicID}.html"><img src="{$ViewPhoto}" alt="" Style="border:1px solid #C0C0C0" width="{$width}px" height="{$height}px"/></a>
<div class="photo2">
�� {$topic} ��
</div>
</div>
]]>
</xsl:variable>
<xsl:variable name="String_25" title="boke_postnews">
<![CDATA[
<li><a href="{$bokeurl}{$bokename}.showtopic.{$TopicID}.{$PostID}.html#{$PostID}">{$title}</a><br/>{$postusername}</li>
]]>
</xsl:variable>
<xsl:variable name="String_26" title="boke_bokecounts">
<![CDATA[
<li>������:{$TodayNum}</li>
<li>������:{$TopicNum}</li>
<li>�ղ���:{$FavNum}</li>
<li>ͼƬ��:{$PhotoNum}</li>
<li>������:{$PostNum}</li>
<li>����ʱ��:{$JoinBokeTime}</li>
<li>����ʱ��:{$LastUpTime}</li>
]]>
</xsl:variable>
<xsl:variable name="String_27" title="IS_BokeOwnerNav">
<![CDATA[
<a href="BokeManage.asp" target="_blank" title="�����ҵĸ��˲���">������</a>&nbsp;&nbsp;|&nbsp;&nbsp;
]]>
</xsl:variable>
<xsl:variable name="String_28" title="Not_BokeOwnerNav">
<![CDATA[
<a href="BokeApply.asp">���벩��</a>&nbsp;&nbsp;|&nbsp;&nbsp;
]]>
</xsl:variable>
<xsl:variable name="String_29" title="other_BokeOwnerNav">
<![CDATA[
]]>
</xsl:variable>
<xsl:variable name="String_30" title="BokePayto">
<![CDATA[
&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="{$bokeurl}{$bokename}.showchannel.3.html" onMouseOver="showmenu(event,'','ShowChannel_3');">{$PaytoStr}</a>
]]>
</xsl:variable>
<xsl:variable name="String_31" title="BokePayto">
<![CDATA[
����Ƶ��
]]>
</xsl:variable>
<xsl:variable name="String_32" title="BokeNone">
<![CDATA[
<ul>
<li><B>������Ϣ</B>���ò����û������ڻ���д����������
</ul>
]]>
</xsl:variable>
<xsl:variable name="String_33" title="TopNav_Sys">
<![CDATA[
<div id="navmenu">
<div id="nav_r">
{$BokeOwnerNav}
<a href="index.asp">������̳</a>
</div>
<a href="bokeindex.asp">������ҳ</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="{$ibokeurl}show_user.html">��������</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="{$ibokeurl}show_topic.html">��������</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="{$ibokeurl}show_topic.1.html">����</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="{$ibokeurl}show_topic.2.html">�ղ�</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="{$ibokeurl}show_topic.5.html">���</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="{$ibokeurl}show_topic.3.html">��ǩ</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="{$ibokeurl}show_topic.4.html">����</a>
</div>
]]>
</xsl:variable>
<xsl:variable name="String_34" title="iBokeOwnerNav">
<![CDATA[
<a href="{$bokeurl}{$bokename}.index.html">�ҵĲ���</a>&nbsp;&nbsp;|&nbsp;&nbsp;
]]>
</xsl:variable>
<xsl:variable name="String_35" title="SiteUrl">
<![CDATA[
{$SiteUrl}
]]>
</xsl:variable>
<xsl:variable name="String_36" title="BokeStats">
<![CDATA[
<a href="Bokepostings.asp?user={$bokename}&action=bokestats"><img src="{$skinpath}images/stat_open.gif" border="0" alt="����ǰ����״̬�Ĳ�����Ϊ�ر�״̬��"/></a>
]]>
</xsl:variable>
<xsl:variable name="String_37" title="BokeStats">
<![CDATA[
<a href="Bokepostings.asp?user={$bokename}&action=bokestats"><img src="{$skinpath}images/stat_close.gif" border="0" alt="����ǰ�ر�״̬�Ĳ�����Ϊ����״̬��"/></a>
]]>
</xsl:variable>
</xsl:stylesheet>
