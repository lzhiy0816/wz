<?xml version="1.0" encoding="gb2312"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" >
<xsl:output method="xml" omit-xml-declaration = "yes" indent="yes" version="4.0"/>
		<!--
	Copyright (C) 2004,2005 AspSky.Net. All rights reserved.
	Written by dvbbs.net Lao Mi
	Web: http://www.aspsky.net/,http://www.dvbbs.net/
	Email: eway@aspsky.net
	-->
<xsl:template  match="/">
<a href="cookies.asp?action=stylemod&amp;skinid=&amp;boardid=$boardid" >恢复默认设置</a>
<xsl:for-each select="xml/style"><br /><a href="cookies.asp?action=stylemod&amp;skinid={@id}_&amp;boardid=$boardid" title="使用模板[{@stylename}]"><b><xsl:value-of select="@stylename"/></b></a>
<xsl:call-template name="cssmenu">
<xsl:with-param name="id" select="@id"/>
<xsl:with-param name="stylename" select="@stylename"/>
</xsl:call-template>
</xsl:for-each>
</xsl:template>
<xsl:template name="cssmenu">
<xsl:param name="id"/>
<xsl:param name="stylename"/>
<xsl:for-each select="/xml/xml/css[tid=$id]"><br /><a href="cookies.asp?action=stylemod&amp;skinid={$id}_{@id}&amp;boardid=$boardid" title="使用模板[{$stylename}],使用Css样式[{@type}]" ><xsl:value-of select="@type"/></a></xsl:for-each>
</xsl:template>
</xsl:stylesheet>