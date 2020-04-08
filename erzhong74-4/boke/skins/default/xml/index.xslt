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
<xsl:variable name="String_0" title="index_main">
<![CDATA[
<style>
@import url({$skinpath}Css/IndexStyle.css);
</style>
<center>
<div class="pageall">
<!-- �������ݿ�ʼ -->
{$Page_Main}
<!-- �������ݽ��� -->
</div>
</center>
<div style="clear:both"><br/></div>
]]>
</xsl:variable>
<xsl:variable name="String_1" title="index">
<![CDATA[
<div class="PageRow" style="background-color:#fff">
<!--��һ�У����� -->
{$Part_TopNotice}

<!--��һ�У��������� -->
{$Part_TopNews}
</div>
<hr class="dashed"/>
<!--�ڶ��У�������-->
<div class="PageRow">
{$SearchTool}
</div>
<hr class="dashed"/>
<!--�����У���������-->
<div class="pagerow">
	<li class="part" style="padding-right:6px;">
		{$Page_NewJoinBoker}
	</li>
	<li class="part" style="padding-right:6px;">
		{$Page_HotBoker}
	</li>
	<li class="part">
		{$Page_UpBoker}
	</li>
</div>
<hr class="dashed"/>
<!--�����У�����Ŀ2-->
<div class="pagerow">
	<div class="partleft">
		{$Page_NewTopicList}
	</div>
	<div class="partright">
	
		{$Page_WeekPostList}

	</div>
</div>
<div class="pagerow">
	<hr class="dashed"/>
	<div class="partleft">
		{$Page_NewPostList}
	</div>
	<div class="partright">
		{$Page_NewLinkList}
	</div>
</div>
<hr class="dashed"/>
<!--�����У�����ͼƬ-->
<div class="pagerow">
<div id="photolist">
	<div class="title"><img src="{$skinpath}images/photo_title.gif" border="0" align="absMiddle"/>&nbsp;�������</div>
	{$Page_Photos}
</div>
</div>
<!--�����У����ͷ���Ŀ¼-->
<hr class="dashed"/>
<div class="pagerow">
<div id="TopNotice" style="width:29%;">
	<div class="noticeicon"><img src="{$skinpath}images/notice.gif"></div>
	<div class="noticetitle">ϵͳ��Ϣ</div>
	<div class="msg">
		{$SystemInfo}
	</div>
</div>
	<div style="float:right;width:535px;">
	<div id="bokeclass">
		<div style="border-left:10px solid #56B1F5;margin:2px;">
		<div class="title1">��������</div>
		</div>
		<div class="msg">
		{$SysCat}
		</div>
		<div style="border-left:10px solid #91BE06;margin:2px;">
		<div class="title1">��������</div>
		</div>
		<div class="msg">
		{$SysChatCat}
		</div>
	</div>
	</div>
</div>
]]>
</xsl:variable>
<xsl:variable name="String_2" title="index">
<![CDATA[
	<div id="bokeclass">
		<div style="border-left:10px solid #56B1F5;margin:2px;">
			<div class="title1">��������</div>
		</div>
		<div class="msg">
			{$SysCat}
		</div>
	</div>
	<div class="pagerow">
	<hr class="dashed"/>
	<div class="classinfo"><a href="{$ibokeurl}show_user.html">�����û�����</a>&nbsp;��&nbsp;{$Descriptions}</div>
	<hr class="dashed"/>
	</div>
	<div class="pagerow">
	{$BokeUserList}
	</div>
]]>
</xsl:variable>
<xsl:variable name="String_3" title="index">
<![CDATA[
	<script language="JavaScript" src="boke/Script/Pagination.js"></script>
	{$BokeChatCat}
	<div class="pagerow">
	<hr class="dashed"/>
	<div class="classinfo"><a href="{$ibokeurl}show_topic.html">��������</a>&nbsp;��&nbsp;{$Descriptions}{$Descriptions_a}</div>
	<hr class="dashed"/>
	</div>
	<div class="pagerow">
	{$BokeTopicList}
	<hr class="dashed"/>
	<SCRIPT language="JavaScript">
		PageList({$Page},3,{$MaxRows},{$CountNum},"{$PageSearch}",5);
	</SCRIPT>
	</div>
]]>
</xsl:variable>

<xsl:variable name="String_4" title="Page_PhotoList">
<![CDATA[
<div class="showphotos">
{$PhotoLine}
</div>
]]>
</xsl:variable>

<xsl:variable name="String_5" title="Part_TopNotice">
<![CDATA[
<div id="TopNotice">
	<div class="noticeicon" ><img src="{$skinpath}images/notice.gif"></div>
	<div class="noticetitle">�û���Ϣ</div>
	<div class="msg">
		{$UserInfo}
	</div>
</div>	
]]>
</xsl:variable>

<xsl:variable name="String_6" title="Part_TopNews">
<![CDATA[
<div id="TopNews">
	<div class="newsicon"><img src="{$skinpath}images/news.gif"></div>
	<div class="newstitle">�� ��</div>
	<div class="msg">
		{$TopNewsMsg}
	</div>
</div>
]]>
</xsl:variable>

<xsl:variable name="String_7" title="SearchTool">
<![CDATA[
<form method="post" ACTION="?show_topic" onSubmit="this.action = '?show_topic.'+this.stype.options[this.stype.selectedIndex].value+'.html'">
	<div class="search">
		�ؼ��֣�
		<input type="text" name="searchword"/>
		<input type="radio" name="sel" value="0" checked="true">����
		<input type="radio" name="sel" value="1">����
		<select name="stype">
		<option value="">����</option>
		<option value="1" selected="true">���� </option>
		<option value="2">�ղ�</option>
		<option value="3">��ǩ</option>
		<option value="4">����</option>
		<option value="5">���</option>
		</select>
	</div>
	<div class="searchright" style="">		<input style="border:0px" src="{$skinpath}images/searchgo.gif" type="image" alt="�� ��" align="absMiddle"/>
</div>
</form>
]]>
</xsl:variable>

<xsl:variable name="String_8" title="Page_NewJoinBoker">
<![CDATA[
<!-- �²����û� -->
<div class="border0">
	<div class="toptitle1"><img src="{$skinpath}images/icon_green.gif" border="0" align="absMiddle"/>&nbsp;�²����û�</div>
	<div class="msg">
		{$Page_NewJoinBoker}
	</div>
</div>
]]>
</xsl:variable>

<xsl:variable name="String_9" title="Page_HotBoker">
<![CDATA[
<!-- ���Ų��� -->
<div class="border0">
	<div class="toptitle2"><img src="{$skinpath}images/icon_orange.gif" border="0" align="absMiddle"/>&nbsp;���Ų���</div>
	<div class="msg">
		{$Page_HotBoker}
	</div>
</div>
]]>
</xsl:variable>

<xsl:variable name="String_10" title="Page_UpBoker">
<![CDATA[
<!-- �Ƽ����� -->
	<div class="border0">
		<div class="toptitle3"><img src="{$skinpath}images/icon_blue.gif" border="0" align="absMiddle"/>&nbsp;���͸���</div>
		<div class="msg">
			{$Page_UpBoker}
		</div>
	</div>
]]>
</xsl:variable>

<xsl:variable name="String_11" title="Page_NewTopicList">
<![CDATA[
<!-- �������� -->
<table cellSpacing="0" cellPadding="2" border="0" width="98%" align="center">
<tr>
	<td class="t_left" width="8px" height="25px"><img src="{$skinpath}images/subject_left.gif"/></td>
	<td class="t_title" width="94px"><img src="{$skinpath}images/new_article.gif" align="absMiddle"/> ������Ϣ</td>
	<td class="t_title" width="*"><span style="float:right;">|</span>��&nbsp;&nbsp;&nbsp;&nbsp;��</td>
	<td class="t_title" width="85px"><span style="float:right;">|</span>����</td>
	<td class="t_title" width="100px">����ʱ��</td>
	<td class="t_right" width="8px"><img src="{$skinpath}images/subject_right.gif"/></td>
</tr>
{$Page_NewTopicList}
</table>
]]>
</xsl:variable>

<xsl:variable name="String_12" title="Page_NewPostList">
<![CDATA[
<!-- �������� -->
<table cellSpacing="0" cellPadding="2" border="0" width="98%" align="center">
<tr>
	<td class="t_left" width="8px" height="25px"><img src="{$skinpath}images/reply_left.gif"/></td>
	<td class="t_title1" width="94px"><img src="{$skinpath}images/new_reply.gif" align="absMiddle"/> ��������</td>
	<td class="t_title1" width="*"><span style="float:right;">|</span>��&nbsp;&nbsp;&nbsp;&nbsp;��</td>
	<td class="t_title1" width="85px"><span style="float:right;">|</span>����</td>
	<td class="t_title1" width="100px">����ʱ��</td>
	<td class="t_right" width="8px"><img src="{$skinpath}images/reply_right.gif"/></td>
</tr>
{$Page_NewPostList}
</table>
]]>
</xsl:variable>
<xsl:variable name="String_13" title="SysCat">
<![CDATA[
<!-- ���ർ�� -->
{$SysCat}
]]>
</xsl:variable>
<xsl:variable name="String_14" title="SysChatCat">
<![CDATA[
<!-- ���⵼�� -->
{$SysChatCat}
]]>
</xsl:variable>
<xsl:variable name="String_15" title="Main_System_Post">
<![CDATA[

]]>
</xsl:variable>
<xsl:variable name="String_16" title="Main_System_Photos">
<![CDATA[

]]>
</xsl:variable>
<xsl:variable name="String_17" title="Main_System_msg1">
<![CDATA[
</div>
]]>
</xsl:variable>
<xsl:variable name="String_18" title="Main_System_msg2">
<![CDATA[
]]>
</xsl:variable>
<!-- �����б� -->
<xsl:variable name="String_19" title="Msg_NewJoinBoker">
<![CDATA[
<span class="list_left">
<img src="{$skinpath}images/top_num{$num}.gif" border="0" align="absMiddle"/>&nbsp;<a href="{$bokeurl}{$Boke_Name}.index.html" target="_blank">{$Boke_Title}</a>
</span>
<span class="list_right">
	{$Boke_User}
</span>
]]>
</xsl:variable>
<xsl:variable name="String_20" title="Msg_HotBoker">
<![CDATA[
<span class="list_left">
	<a href="{$bokeurl}{$Boke_Name}.index.html" target="_blank">{$Boke_Title}</a>
</span>
<span class="list_right">
	{$Boke_User}
</span>
]]>
</xsl:variable>
<xsl:variable name="String_21" title="Msg_TopBoker">
<![CDATA[
<span class="list_left">
	<a href="{$bokeurl}{$Boke_Name}.index.html" target="_blank">{$Boke_Title}</a>
</span>
<span class="list_right">
	{$Boke_User}
</span>
]]>
</xsl:variable>

<xsl:variable name="String_22" title="Page_NewTopicList">
<![CDATA[
<tr>
<td></td>
<td class="t_body1">{$CatName}</td>
<td class="t_body1" align="left"><a href="{$bokeurl}Userid_{$UserID}.showtopic.{$TopicID}.html" alt="���������" target="_blank">{$Title}</a></td>
<td class="t_body1"><a href="boke.asp?userid={$UserID}" title="�������߲���" target="_blank"><b>{$PostUser}</b></a></td>
<td class="t_body1">{$LastPostTime}</td>
<td></td>
</tr>
]]>
</xsl:variable>
<xsl:variable name="String_23" title="Page_NewTopicList">
<![CDATA[
<tr>
<td></td>
<td class="t_body1">{$CatName}</td>
<td class="t_body1" align="left"><a href="{$bokeurl}Userid_{$UserID}.showtopic.{$TopicID}.html" alt="���������" target="_blank">{$Title}</a></td>
<td class="t_body1"><b>{$PostUser}</b></td>
<td class="t_body1">{$LastPostTime}</td>
<td></td>
</tr>
]]>
</xsl:variable>
<xsl:variable name="String_24" title="Page_SysCats">
<![CDATA[
<a href="{$ibokeurl}show_user.{$catid}.html">{$CatName}</a>&nbsp;&nbsp;
]]>
</xsl:variable>
<xsl:variable name="String_25" title="Page_SysChatCats">
<![CDATA[
<a href="{$ibokeurl}show_topic.{$t}.{$catid}.html">{$CatName}</a>&nbsp;&nbsp;
]]>
</xsl:variable>
<xsl:variable name="String_26" title="Page_SignPhoto">
<![CDATA[
<li>
<div class="photo">
<a href="{$bokeurl}Userid_{$UserID}.showtopic.{$TopicID}.html" target="_blank"><img src="{$ViewPhoto}" border="0" width="115" height="100"/></a>
</div>
</li>
]]>
</xsl:variable>
<xsl:variable name="String_27" title="Page_SysPhotolist1">
<![CDATA[
<table border="0" cellpadding="3" cellspacing="0" align="center" Style="width:98%">
<tr>
{$photo_list}
</tr>
</table>
]]>
</xsl:variable>
<xsl:variable name="String_28" title="Page_SysPhotolist2">
<![CDATA[
<td align="center" width="25%">

<table border="0" width="100%" height="100%" cellspacing="1" class="table2">
	<tr>
		<td height="20px" colspan="4" class="t_title2">�� {$Title} ��</td>
	</tr>
	<tr>
		<td colspan="4" class="td2" style="text-align:center;">
		<a href="{$bokeurl}Userid_{$UserID}.showtopic.{$TopicID}.html" target="_blank"><img src="{$ViewPhoto}" border="0" width="115" height="100"/></a>
		</td>
	</tr>
	<tr>
		<td width="20%" class="td1">����</td>
		<td width="30%" class="td2"><a href="boke.asp?UserID={$UserID}" title="������߲���" target="_blank">{$UserName}</a></td>
		<td width="20%" class="td1">ʱ��</td>
		<td width="30%" class="td2">{$DateTime}</td>
	</tr>
	<tr>
		<td class="td1">��С</td>
		<td class="td2" colspan="3">{$FileSize} KB</td>
	</tr>
</table>
</td>
]]>
</xsl:variable>
<xsl:variable name="String_29" title="other">
<![CDATA[
]]>
</xsl:variable>
<xsl:variable name="String_30" title="Page_Syscat">
<![CDATA[
<script language="JavaScript" src="boke/Script/Pagination.js"></script>
<br/>
<table cellSpacing="0" cellPadding="2" border="0" width="98%" align="center">
<tr>
	<td class="t_left" width="8px" height="25px"><img src="{$skinpath}images/subject_left.gif"/></td>
	<td class="t_title" width="85px"><span style="float:right;">|</span>����</td>
	<td class="t_title" width="*"><span style="float:right;">|</span>��������</td>
	<td class="t_title" width="100px"><span style="float:right;">|</span>�û�</td>
	<td class="t_title" width="45px"><span style="float:right;">|</span>����</td>
	<td class="t_title" width="45px"><span style="float:right;">|</span>����</td>
	<td class="t_title" width="45px"><span style="float:right;">|</span>�ղ�</td>
	<td class="t_title" width="45px"><span style="float:right;">|</span>ͼƬ</td>
	<td class="t_title" width="100px">����ʱ��</td>
	<td class="t_right" width="8px"><img src="{$skinpath}images/subject_right.gif"/></td>
</tr>
{$Page_BokeUserList}
</table>
<SCRIPT language="JavaScript">
PageList({$Page},3,{$MaxRows},{$CountNum},"{$PageSearch}",5);
</SCRIPT>
]]>
</xsl:variable>
<xsl:variable name="String_31" title="Page_Syscat">
<![CDATA[
<tr>
<td></td>
<td class="t_body1">{$CatName}</td>
<td class="t_body1">
<a href="{$bokeurl}{$BokeSn}.index.html" title="������û�����" target="_blank">{$BokeTitle}</a>
</td>
<td class="t_body1"><a href="dispuser.asp?userid={$UserID}" title="������û�����" target="_blank"><b>{$BokeUser}</b></a></td>
<td class="t_body1">{$TopicNum}</td>
<td class="t_body1">{$PostNum}</td>
<td class="t_body1">{$FavNum}</td>
<td class="t_body1">{$PhotoNum}</td>
<td class="t_body1">{$JoinTime}</td>
<td></td>
</tr>
]]>
</xsl:variable>
<xsl:variable name="String_32" title="Page_SysTopiclist">
<![CDATA[
<div class="cat_title">
<div class="cat_note">����({$Child}) | �Ķ�({$Hits})</div>
{$Num}. <a href="{$bokeurl}Userid_{$UserID}.showtopic.{$TopicID}.html" target="_blank">{$Title}</a>
</div>
<div class="cat_end">
���ࣺ[{$CatName}] -- ���ߣ�<a href="boke.asp?UserID={$UserID}" title="������߲���" target="_blank">{$UserName}</a> -- ����ʱ�䣺<font color="gray">{$PostTime}</font>
</div>
<hr class="dashed"/>
]]>
</xsl:variable>
<xsl:variable name="String_33" title="Page_SysLinklist">
<![CDATA[
<div class="cat_title">
<div class="cat_note">{$Logo}</div>
{$Num}. <a href="{$bokeurl}Userid_{$UserID}.showtopic.{$TopicID}.html" target="_blank">{$Title}</a>
</div>
<div class="cat_end">
���ࣺ[{$CatName}] -- ����ߣ�<a href="boke.asp?UserID={$UserID}" title="������߲���" target="_blank">{$UserName}</a> -- ���ʱ�䣺<font color="gray">{$PostTime}</font> -- ���(<font color="red">{$Hits}</font>)
</div>
<hr class="dashed"/>
]]>
</xsl:variable>
<xsl:variable name="String_34" title="Page_SysChatCat">
<![CDATA[
	<div id="bokeclass">
		<div style="border-left:10px solid #56B1F5;margin:2px;">
			<div class="title1">��������</div>
		</div>
		<div class="msg">
			{$SysCat}
		</div>
	</div>
]]>
</xsl:variable>

<xsl:variable name="String_35" title="Page_WeekPostList">
<![CDATA[
<!-- �������� -->
<div class="border0">
	<div class="toptitle1"><img src="{$skinpath}images/icon_green.gif" border="0" align="absMiddle"/>&nbsp;��������</div>
	<div class="msg">
		{$Page_WeekPostList}
	</div>
</div>
]]>
</xsl:variable>

<xsl:variable name="String_36" title="Page_NewLinkList">
<![CDATA[
<!-- ������ǩ -->
<div class="border0">
	<div class="toptitle2"><img src="{$skinpath}images/icon_orange.gif" border="0" align="absMiddle"/>&nbsp;������ǩ</div>
	<div class="msg">
		{$Page_NewLinkList}
	</div>
</div>
]]>
</xsl:variable>
<xsl:variable name="String_37" title="Msg_WeekPostList">
<![CDATA[
<span class="list_left_a">
	{$num}.&nbsp;<a href="{$bokeurl}Userid_{$UserID}.showtopic.{$TopicID}.html" target="_blank" title="{$PostUser}������{$LastPostTime},����{$Child}������">{$Title}</a>
</span>
]]>
</xsl:variable>
<xsl:variable name="String_38" title="Msg_NewLinkList">
<![CDATA[
<span class="list_left_a">
	{$num}.&nbsp;<a href="{$bokeurl}Userid_{$UserID}.showtopic.{$TopicID}.html" target="_blank" title="{$PostUser}�����{$LastPostTime}">{$Title}</a>
</span>
]]>
</xsl:variable>
<xsl:variable name="String_39" title="SystemInfo">
<![CDATA[
	<div style="margin-left:20px;clear:both;">
		<li>������Ϣ��{$TodayNum}</li>
		<li>�����û���{$UserNum}</li>
		<li>�������£�{$TopicNum}</li>
		<li>�����ղأ�{$FavNum}</li>
		<li>����ͼƬ��{$PhotoNum}</li>
		<li>��Ϣ���ۣ�{$PostNum}</li>
	</div>
]]>
</xsl:variable>
<xsl:variable name="String_40" title="UserInfo_Login">
<![CDATA[<BR/>
<div style="margin-left:8px;">
<FORM METHOD=POST ACTION="login.asp?action=chk">
��̳�û���<input tyep="text" name="username" size="12" /><BR/>
��̳���룺<input type="password" name="password" size="12" /><BR/>
{$GetCode}
��¼���棺<select name="CookieDate"><option value="0" selected="selected">������</option><option value="1">����һ��</option><option value="2">����һ��</option><option value="3">����һ��</option></select><BR/><BR/>
<input type="hidden" value="bokeindex.asp" name="comeurl">
<input type="submit" name="submit" value="��¼ϵͳ" />
</FORM>
</div>
]]>
</xsl:variable>
<xsl:variable name="String_41" title="UserInfo_isLogin">
<![CDATA[
	<div style="margin-left:20px;clear:both;">
		<li>��ӭ <B>{$UserName}</B> ���벩��</li>
		<li>������Ϣ��{$TodayNum}</li>
		<li>�������£�{$TopicNum}</li>
		<li>����ͼƬ��{$PhotoNum}</li>
		<li>��Ϣ���ۣ�{$PostNum}</li>
		<li>{$UserMsg}</li>
	</div>
]]>
</xsl:variable>
<xsl:variable name="String_42" title="UserInfo_GetCode">
<![CDATA[
�� ֤ �룺{$Dv_GetCode}<!--<input type="text" name="codestr" size="4" /> <img src="DV_getcode.asp" alt="��֤��" height="18px;" align="middle"/>--><BR/>
]]>
</xsl:variable>
<xsl:variable name="String_43" title="UserInfo_a">
<![CDATA[
<a href="{$bokeurl}{$bokename}.index.html">���˲���</a>&nbsp;&nbsp;<a href="bokemanage.asp">���͹���</a>
]]>
</xsl:variable>
<xsl:variable name="String_44" title="UserInfo_b">
<![CDATA[
<a href="bokeapply.asp" title="����������ĸ��˲���"><font color="blue">�ڱ�վ�������ĸ��˲���</font></a>
]]>
</xsl:variable>
<xsl:variable name="String_45" title="msg3">
<![CDATA[
������Ϣ�б�
]]>
</xsl:variable>
<xsl:variable name="String_46" title="msg3">
<![CDATA[
��&nbsp;{$showcat}
]]>
</xsl:variable>
</xsl:stylesheet>