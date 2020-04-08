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
<xsl:variable name="tdcount">
<xsl:choose>
		<xsl:when test="count(xml/text) &gt; 6">6</xsl:when>
		<xsl:otherwise><xsl:value-of select="count(xml/text)" /></xsl:otherwise>
</xsl:choose>
</xsl:variable>
<xsl:if test="$tdcount !=0">
<xsl:variable name="tdwidth"><xsl:value-of select="floor((100 div $tdcount))"/>%</xsl:variable>
<table cellpadding="3" cellspacing="1" align="center" class="tableborder1">
<xsl:call-template name="showrows">
<xsl:with-param name="tdwidth" select="$tdwidth"/>
<xsl:with-param name="positions" select="0"/>
<xsl:with-param name="tdcount" select="$tdcount"/>
</xsl:call-template>
</table>
</xsl:if>
</xsl:template>
<xsl:template name="showrows">
<xsl:param name="tdwidth"/>
<xsl:param name="positions"/>
<xsl:param name="tdcount"/>
<tr>
<xsl:for-each select="xml/text[position() &gt; $positions and position() &lt; ($tdcount+$positions+1)]">
		 	<td width="{$tdwidth}" class="tablebody1" height="20" style="text-align : center; "><xsl:value-of select="." disable-output-escaping="yes"/></td>
</xsl:for-each>
</tr>
<xsl:if test="xml/text[position() &gt; ($tdcount+$positions)]">
<xsl:call-template name="showrows">
<xsl:with-param name="tdwidth" select="$tdwidth"/>
<xsl:with-param name="positions" select="$positions+$tdcount"/>
<xsl:with-param name="tdcount" select="$tdcount"/>
</xsl:call-template>
</xsl:if>
</xsl:template>
</xsl:stylesheet>