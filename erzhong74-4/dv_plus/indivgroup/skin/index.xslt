<?xml version="1.0" encoding="gb2312"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" >
<xsl:output method="xml" omit-xml-declaration = "yes" indent="yes" version="4.0"/>
	<!--
	����Ȧ����Ŀ�б�ҳ��
	Copyright (C) 2004,2005 AspSky.Net. All rights reserved.
	Written by dvbbs.net Lao Mi
	Web: http://www.aspsky.net/,http://www.dvbbs.net/
	Email: eway@aspsky.net
	-->
<xsl:template match="/">
<xsl:for-each select="IndivGroup">
	<div class="mainbar4" id="boardmaster">
		<xsl:if test="@powerflag='1' or @powerflag='2' or @powerflag='3'">
		<div id="boardmanage">
			����ѡ��:��<a href="IndivGroup_Manage.asp?managetype=board&amp;action=boardmanage&amp;groupid={@groupid}">�����Ŀ</a>| <a href="IndivGroup_Manage.asp?managetype=user&amp;action=user&amp;groupid={@groupid}">��Ա����</a>| <a href="IndivGroup_Manage.asp?groupid={@groupid}">Ȧ�ӹ���</a>
		</div>
		</xsl:if>
		<div id="masterpic"></div>
		<div style="float:left;text-indent:5px;">����Ա:
		<xsl:for-each select="MasterList/Master">
			<a href="dispuser.asp?id={@userid}"><xsl:value-of select="@username" /></a>
		</xsl:for-each>
		</div>
	</div>
	<div class="th">
		<img src="{@picurl}/nofollow.gif" border="0"/>
		<xsl:value-of select="@groupname" disable-output-escaping="yes" /> - ����Ŀ�б�
	</div>
	<xsl:choose>
		<xsl:when test="BoardList/Board">
			<xsl:call-template name="showGroupBoard" />
		</xsl:when>
		<xsl:otherwise>
			<div class="mainbar2" style="height:25px;overflow :hidden;">��Ȧ��Ŀǰû����Ŀ</div>
		</xsl:otherwise>
	</xsl:choose>
</xsl:for-each>
<br />
<div class="itableborder" style="text-align:center">
<img src="{IndivGroup/forum_setting/@pic_0}" align="absmiddle" alt="û���µ�����" /> û���µ����ۡ���<img src="{IndivGroup/forum_setting/@pic_1}" align="absmiddle" alt="���µ�����" /> ���µ����ۡ���<img src="{IndivGroup/forum_setting/@pic_2}" align="absmiddle" alt="����������̳" /> ����������Ŀ
</div>
</xsl:template>

<xsl:template name="showGroupBoard">
<xsl:for-each select="BoardList/Board">
	<xsl:variable name="Boardid" select="@id"/>
	<div class="mainbar" onmouseover="boardbarover(this);" onmouseout="boardbarover(this);" style="height:60px;">
		<div class="index_right" style="height:44px;line-height:normal;margin-top:6px;">
			<div>���⣺<a href="GroupDispbbs.asp?GroupBoardid={@lastpost_7}&amp;ID={@lastpost_6}&amp;replyID={@lastpost_1}&amp;skin=1"><xsl:value-of select="@lastpost_3" disable-output-escaping="yes" /></a></div>
			<div>������<a href="dispuser.asp?id={@lastpost_5}"><xsl:value-of select="@lastpost_0"/></a></div>
			<div>���ڣ�<xsl:value-of select="@lastpost_2"/></div>
		</div>
		<div class="index_left_states"><img>
		<xsl:choose>
			<xsl:when test="@boardstats=0">
				<xsl:choose>
					<xsl:when test="@newpost=1"><xsl:attribute name="title">������Ŀ,��������</xsl:attribute></xsl:when>
					<xsl:otherwise><xsl:attribute name="title">������Ŀ,��������</xsl:attribute></xsl:otherwise>
				</xsl:choose>
				<xsl:attribute name="src"><xsl:value-of select="/IndivGroup/forum_setting/@pic_2"/></xsl:attribute>
			</xsl:when>
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test="@newpost=1">
						<xsl:attribute name="alt">���ŵ���Ŀ,��������</xsl:attribute>
						<xsl:attribute name="src"><xsl:value-of select="/IndivGroup/forum_setting/@pic_1"/></xsl:attribute>
					</xsl:when>
					<xsl:otherwise>
						<xsl:attribute name="alt">���ŵ���Ŀ,��������</xsl:attribute>
						<xsl:attribute name="src"><xsl:value-of select="/IndivGroup/forum_setting/@pic_0"/></xsl:attribute>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
		</img></div>
		<div style="text-align :left;text-indent:5px;">
			<xsl:if test="@indeximg!=''"><a href="IndivGroup_Index.asp?GroupID={@rootid}&amp;GroupBoardid={$Boardid}"><img src="{@indeximg}" alt="" class="boardlogo" /></a></xsl:if>
			<div class="board_style">
				<a href="IndivGroup_Index.asp?GroupID={@rootid}&amp;GroupBoardid={$Boardid}"><xsl:value-of select="@boardname" disable-output-escaping="yes"/></a>
				<font face="Arial" color="#005596"><small> [<font color="#ff6600"><xsl:value-of select="concat(' ',@todaynum,' ')"/></font>|<xsl:value-of select="concat(' ',@topicnum,' ')"/>|<xsl:value-of select="concat(' ',@postnum,' ')"/>]</small></font>
			</div>
			<div style="overflow :hidden;height:38px;"><font face="Arial"><img alt="" src="Skins/Default/Forum_readme.gif" align="middle"/><xsl:value-of select="@boardinfo" disable-output-escaping="yes"/></font></div>
		</div>
	</div>
	<div class="mainbar2" style="height:1px;overflow :hidden;"> </div>
</xsl:for-each>
</xsl:template>
</xsl:stylesheet>