<?xml version="1.0" encoding="gb2312"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" >
<xsl:output method="xml" omit-xml-declaration = "yes" indent="yes" version="4.0"/>
	<!--
	Copyright (C) 2004,2005 AspSky.Net. All rights reserved.
	Written by dvbbs.net Lao Mi
	Web: http://www.aspsky.net/,http://www.dvbbs.net/
	Email: eway@aspsky.net
	论坛公告模板
	-->
<xsl:variable name="marquee" select="0"/><!--设置为1则移动公告-->
<xsl:variable name="maxposition" select="5"/><!--移动公告最多显示多少条-->
<xsl:template  match="/">
<xsl:variable name="boardid" select="xml/@boardid"/>
<div class="itableborder">
<xsl:choose>
		<xsl:when test="xml/news[@boardid=$boardid]">
		<xsl:choose>
		<xsl:when test="$marquee=1">
		<div style="width:80%">
		<marquee scrolldelay="150" scrollamount="4" onmouseout="if (document.all!=null)this.start()" onmouseover="if (document.all!=null)this.stop()">
		公告：<xsl:for-each select="xml/news[@boardid=$boardid][position() &lt; ($maxposition+1)]"> <a href="javascript:openScript('announcements.asp?boardid={$boardid}',500,400)"><b><xsl:value-of select="@title" disable-output-escaping="yes"/></b></a>(<xsl:value-of select="translate(@addtime,'T',' ')" />) </xsl:for-each>
		</marquee>
		</div>
		</xsl:when>
		<xsl:otherwise>
		<xsl:if test="xml/news[@boardid=$boardid]/@bgs and xml/news[@boardid=$boardid]/@bgs !=''"><bgsound  src="{xml/news[@boardid=$boardid]/@bgs}"/><img src="Skins/Default/filetype/mid.gif" border="0" alt="" /> </xsl:if><a href="javascript:openScript('announcements.asp?action=showone&amp;boardid={$boardid}',500,400)"><b><xsl:value-of select="xml/news[@boardid=$boardid]/@title" disable-output-escaping="yes"/></b></a>(<xsl:value-of select="translate(xml/news[@boardid=$boardid]/@addtime,'T',' ')" />)</xsl:otherwise>
			</xsl:choose>
		</xsl:when>
		<xsl:otherwise>
		<a href="javascript:openScript('announcements.asp?action=showone&amp;boardid={$boardid}',500,400)"><b>当前还未有公告</b></a>
		</xsl:otherwise>
</xsl:choose>
</div>
</xsl:template>
</xsl:stylesheet>