<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" omit-xml-declaration = "yes" indent="yes" version="4.0"/>
	<!--
	Copyright (C) 2004,2007 AspSky.Net. All rights reserved.
	Written by dvbbs.net Sunwin
	-->
	
	<xsl:template  match="/">
		<xsl:apply-templates select="//channel" />
	</xsl:template>
	<xsl:template match="channel">
		<table class="spacetable" cellspacing="1" cellpadding="0" align="center">
			<tr>
				<td>
					<div id="rsschannel">
						<!--注意，此处不能使用UL-->
						<ol>
							<xsl:apply-templates select="item[position() &lt;= 5]">
								<xsl:sort select="pubDate" order="descending"/>
							</xsl:apply-templates>
						</ol>
					</div>
				</td>
			</tr>
		</table>

	</xsl:template>
	<xsl:template match="item">
		<li>
			<div class="itemtitle">
				<img src="images/dnone.gif" alt="" onclick="rssnews(event,this);" class="block" align="absmiddle"/>
				<a href="{link}" target="_blank">
					<xsl:value-of select="title"/>
				</a>
				<xsl:if test="category!=''">
					<span class="category">
						[<xsl:value-of select="category"/>]
					</span>
				</xsl:if>
				
			</div>
			<div name="decodeable" class="itemcontent hide">
				<xsl:call-template name="outputContent" />
			</div>
			<div class="itemposttime hide">
				<xsl:value-of select="pubDate"/>
			</div>
		</li>
	</xsl:template>
	<xsl:template name="outputContent">
		<xsl:choose>
			<xsl:when test="xhtml:body" xmlns:xhtml="http://www.w3.org/1999/xhtml">
				<xsl:copy-of select="xhtml:body/*"/>
			</xsl:when>
			<xsl:when test="xhtml:div" xmlns:xhtml="http://www.w3.org/1999/xhtml">
				<xsl:copy-of select="xhtml:div"/>
			</xsl:when>
			<xsl:when test="description">
				<xsl:value-of select="description" disable-output-escaping="yes"/>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
</xsl:stylesheet>