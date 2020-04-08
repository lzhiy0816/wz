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
<xsl:for-each select="xml/plus[@plus_type = 0]">
<xsl:variable name="id" select="@id"/>
<xsl:variable name="haschild" ><xsl:if test="/xml/plus[@plus_type=$id]">1</xsl:if></xsl:variable>
<xsl:if test="$haschild = 1"><xsl:call-template name="childmenu">
<xsl:with-param name="plus_type" select="$id"/>
</xsl:call-template></xsl:if>
<div class="menudiv1"><a><xsl:choose>
		<xsl:when test="@plus_setting = 0 or @plus_setting = 1"><xsl:attribute name="href"><xsl:value-of select="@mainpage"/></xsl:attribute><xsl:if test="@plus_setting = 1"><xsl:attribute name="target">_blank</xsl:attribute></xsl:if></xsl:when>
		<xsl:otherwise>
		<xsl:attribute name="href">JavaScript:openScript('<xsl:value-of select="@mainpage"/>',<xsl:choose>
		<xsl:when test="@plus_setting = 2"><xsl:value-of select="concat(@width,',',@height)"/></xsl:when>
		<xsl:otherwise>screen.width,screen.height</xsl:otherwise>
</xsl:choose>)</xsl:attribute>
		</xsl:otherwise>
</xsl:choose><xsl:attribute name="title"><xsl:value-of select="@plus_copyright"/></xsl:attribute><xsl:if test="$haschild = 1"><xsl:attribute name="onmouseover">showmenu(event,'','Menu_plus_<xsl:value-of select="$id" />');</xsl:attribute></xsl:if><xsl:value-of select="@plus_name" disable-output-escaping="yes"/></a></div>
</xsl:for-each>
</xsl:template>
<xsl:template name="childmenu">
<xsl:param name="plus_type"/>
<div class="menu_popup" id="Menu_plus_{$plus_type}">
<div class="menuitems">
<xsl:for-each select="/xml/plus[@plus_type = $plus_type]">
<a><xsl:choose>
		<xsl:when test="@plus_setting = 0 or @plus_setting = 1"><xsl:attribute name="href"><xsl:value-of select="@mainpage"/></xsl:attribute><xsl:if test="@plus_setting = 1"><xsl:attribute name="target">_blank</xsl:attribute></xsl:if></xsl:when>
		<xsl:otherwise>
		<xsl:attribute name="href">JavaScript:openScript('<xsl:value-of select="@mainpage"/>',<xsl:choose>
		<xsl:when test="@plus_setting = 2"><xsl:value-of select="concat(@width,',',@height)"/></xsl:when>
		<xsl:otherwise>screen.width,screen.height</xsl:otherwise>
</xsl:choose>)</xsl:attribute>
		</xsl:otherwise>
</xsl:choose><xsl:attribute name="title"><xsl:value-of select="@plus_copyright"/></xsl:attribute><xsl:value-of select="@plus_name" disable-output-escaping="yes"/></a>
<br />
</xsl:for-each>
</div>
</div>
</xsl:template>
</xsl:stylesheet>