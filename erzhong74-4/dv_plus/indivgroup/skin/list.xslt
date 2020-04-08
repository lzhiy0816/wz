<?xml version="1.0" encoding="gb2312"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" >
	<xsl:output method="xml" omit-xml-declaration = "yes" indent="yes"/>
	<!--
	Copyright (C) 2004,2005 AspSky.Net. All rights reserved.
	Written by dvbbs.net Lao Mi
	Web: http://www.aspsky.net/,http://www.dvbbs.net/
	Email: eway@aspsky.net
	-->
<xsl:template  match="/">
<!--个性圈子排行整页面-->
<script language="JavaScript" src="Dv_plus/IndivGroup/js/IndivGroup_Main.js"></script>
<script language="JavaScript" src="inc/Pagination.js"></script>
<table cellpadding="6" cellspacing="1" align="center" class="tableborder1">
	<tr><th width="50%">最新创建圈子</th><th width="50%">热门圈子</th></tr>
	<tr><td class="tablebody1"><script src="Dv_News.asp?GetName=newig"></script></td><td class="tablebody1"><script src="Dv_News.asp?GetName=hotig"></script></td></tr>
</table>
<br/>
<table cellpadding="6" cellspacing="1" align="center" class="tableborder1">
	<tr>
		<td colspan="4" class="tabletitle2" width="50%">
			<div style="float:left;padding-left:12px;">个性圈子总数： <xsl:value-of select="xml/Forum/@IndivGroupTotal" /> 个</div>
			<div style="float:left;padding-left:12px;height:21px;"><input type="button" name="appgroup" value="申请圈子" onclick="ShowAppForm({xml/Forum/@AppPowerFlag});" /></div>
			<div id="appgroupform" style="display:none;float:left;padding-left:12px;"></div>
			<iframe style="border:0px;width:0px;height:0px;" src="" name="hiddenframe" id="hiddenframe"></iframe>
		</td>
		<td colspan="5" align="right" valign="top" class="tabletitle2">
			<form method="post" action="?action=grouplist" name="IndivQueryFrom">
			个性圈子搜索：<input name="keyword" type="text" value="{xml/Forum/@keyword}" size="15" /> <input name="Submit" type="submit" value="搜索" />
			<select name="orders" onchange="javascript:submit()">
				<option value="1">按发贴总数排序</option>
				<option value="2">按成员总数排序</option>
				<option value="3">按最新建立排序</option>
			</select>
			<script language="JavaScript">ChkSelected(document.IndivQueryFrom.orders,'<xsl:value-of select="xml/Forum/@Orders" />');</script>
			</form>
		</td>
	</tr>
	<tr>
		<th>圈子名称</th>
		<th>管理员</th>
		<th>状态</th>
		<th>圈子成员数</th>
		<th>圈子发贴数</th>
		<th>圈子主题数</th>
		<th>圈子今日贴数</th>
		<th>圈子建立日期</th>
		<th></th>
	</tr>
	<xsl:choose>
		<xsl:when test="xml/IndivGroup/row">
			<xsl:for-each select="xml/IndivGroup/row">
				<xsl:call-template name="IndivGroupList" />
			</xsl:for-each>
		</xsl:when>
		<xsl:otherwise><tr><td colspan="9" class="tablebody1">　　搜索不到个性圈子数据。</td></tr></xsl:otherwise>
	</xsl:choose>
</table>
<table cellpadding="2" cellspacing="1" align="center" width="98%">
	<tr>
		<td>
<script language="javascript">
	PageList(<xsl:value-of select="xml/Forum/@Page"/>,10,<xsl:value-of select="xml/Forum/@pagesize"/>,<xsl:value-of select="xml/Forum/@IndivGroupTotal"/>,'?action=grouplist',1);
</script>
		</td>
	</tr>
</table>
</xsl:template>

<xsl:template  name="IndivGroupList">
<!--个性圈子排行循环部分-->
<tr>
	<td class="tablebody1"><a href="IndivGroup_index.asp?groupid={@id}"><xsl:value-of select="@groupname"/></a></td>
	<td align="center" class="tablebody2"><a href="dispuser.asp?id={@appuserid}" target="_blank"><xsl:value-of select="@appusername"/></a></td>
	<td align="center" class="tablebody1">
		<xsl:call-template name="IndivGroupStatsStr">
			<xsl:with-param name="Sid"><xsl:value-of select="@stats"/></xsl:with-param>
		</xsl:call-template>
	</td>
	<td align="center" class="tablebody2"><xsl:value-of select="@usernum"/></td>
	<td align="center" class="tablebody1"><xsl:value-of select="@postnum"/></td>
	<td align="center" class="tablebody2"><xsl:value-of select="@topicnum"/></td>
	<td align="center" class="tablebody1"><xsl:value-of select="@todaynum"/></td>
	<td align="center" class="tablebody2"><xsl:value-of select="@passdate"/></td>
	<td align="center" class="tablebody1">
	<xsl:choose>
		<xsl:when test="@islock='1'"><font color="red">待审核中...</font></xsl:when>
		<xsl:when test="@islock='2'">已加入</xsl:when>
		<xsl:otherwise>
			<xsl:choose>
				<xsl:when test="@usernum=@limituser or @limituser='' or @stats=2 or @stats=3">
					<input type="button" id="appJionGroup_{@id}" name="appJionGroup_{@id}" value="申请加入" onclick="submitappjion({/xml/Forum/@AppJionFlag},{@id});" disabled="true" />
				</xsl:when>
				<xsl:otherwise>
					<input type="button" id="appJionGroup_{@id}" name="appJionGroup_{@id}" value="申请加入" onclick="submitappjion({/xml/Forum/@AppJionFlag},{@id});" />
				</xsl:otherwise>
			</xsl:choose>
		</xsl:otherwise>
	</xsl:choose>
	</td>
</tr>
</xsl:template>

<xsl:template name="IndivGroupStatsStr">
<xsl:param name="Sid"/>
<xsl:choose>
	<xsl:when test="$Sid=1">正常</xsl:when>
	<xsl:when test="$Sid=2">锁定</xsl:when>
	<xsl:when test="$Sid=3">关闭</xsl:when>
	<xsl:when test="$Sid=0">审核</xsl:when>
	<xsl:otherwise>未知</xsl:otherwise>
</xsl:choose>
</xsl:template>
</xsl:stylesheet>