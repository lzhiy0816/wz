<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" omit-xml-declaration = "yes" indent="yes" version="4.0"/>
	<!--
	Copyright (C) 2004,2005 AspSky.Net. All rights reserved.
	Written by dvbbs.net Sunwin
	-->
	<xsl:include href="./fuc_system.xslt"/>
	<xsl:include href="./fuc_spacetool.xslt"/>
	<xsl:include href="./fuc_setting.xslt"/>
	<xsl:variable name="canedit_space">
		<xsl:choose>
			<xsl:when test="dv_space/space_info/@isadmin=-1 and dv_space/forum_info/@act=''">
				<xsl:text>1</xsl:text>
			</xsl:when>
			<xsl:otherwise>
				<xsl:text>0</xsl:text>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:variable>

	<xsl:variable name="leftchannal_style">
		<xsl:choose>
			<xsl:when test="dv_space/space_info/@s_style=1">
				<xsl:text>width:240px;display:block;</xsl:text>
			</xsl:when>
			<xsl:when test="dv_space/space_info/@s_style=2">
				<xsl:text>width:0;display:none;</xsl:text>
			</xsl:when>
			<xsl:otherwise>
				<xsl:text>width:200px;display:block;</xsl:text>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	<xsl:variable name="mainchannal_style">
		<xsl:choose>
			<xsl:when test="dv_space/space_info/@s_style=3">
				<xsl:text>width:536px;display:block;</xsl:text>
			</xsl:when>
			<xsl:otherwise>
				<xsl:text>width:705px;display:block;</xsl:text>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	<xsl:variable name="rightchannal_style">
		<xsl:choose>
			<xsl:when test="dv_space/space_info/@s_style=1">
				<xsl:text>width:0;display:none;</xsl:text>
			</xsl:when>
			<xsl:when test="dv_space/space_info/@s_style=2">
				<xsl:text>width:240px;display:block;</xsl:text>
			</xsl:when>
			<xsl:otherwise>
				<xsl:text>width:22%;display:block;</xsl:text>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	<xsl:template  match="/">
		<link rel="stylesheet" type="text/css" id="stylecss" href="{dv_space/space_info/@skinpath}style.css"/>
			<script type="text/javascript">
				var IE=false,FF=false,H,B;IsEdit=0;
				var layoutleft,layoutmain,layoutright;
				var myspaceskin = '<xsl:value-of select="dv_space/space_info/@skinpath"/>';
				window.onload=function(){
				B=document.getElementsByTagName("body")[0];
				H=document.getElementsByTagName("html")[0];
				}
			</script>
		<xsl:if test="$canedit_space=1">
			<script language = "javaScript" src = "dv_plus/myspace/drag/drag.js" type="text/javascript"></script>
		</xsl:if>
		<script language = "javaScript" src = "dv_plus/myspace/drag/space.js" type="text/javascript"></script>
		<div id="top">
			<div class="left"></div>
			<div class="center">
				<xsl:call-template name="space_top"/>
			</div>
			<div class="right"></div>
			<div class="space"></div>
		</div>
		<div id="head">
			<div class="left"></div>
			<div class="center">
				<xsl:call-template name="space_head"/>
			</div>
			<div class="right"></div>
			<div class="space"></div>
		</div>
		<div id="menus">
			<div class="left">
			</div>
			<div class="center">
				<xsl:call-template name="space_menus"/>
			</div>
			<div class="right"></div>
			<div class="space"></div>
		</div>
		<div id="spacemain">
		<div id="top_style">
		<h2></h2>
		<h1></h1>
		</div>
			<xsl:if test="$canedit_space=1">
				<div class="center">
					<xsl:call-template name="space_toolbar"/>
				</div>
			</xsl:if>
			<div class="center">
				<xsl:call-template name="space_main"/>
				<div class="space"></div>
			</div>
		</div>
		<div id="bot_style">
		<h2></h2>
		<h1></h1>
		</div>
		<div id="footer">
			<div class="left"></div>
			<div class="center">
				<xsl:call-template name="space_footer"/>
			</div>
			<div class="right"></div>
		</div>

		<xsl:if test="$canedit_space=1">
		<script type="text/javascript">
			LoadSet(document.getElementById("layoutset").value);
		</script>
		</xsl:if>
	</xsl:template>
	<xsl:template name="space_top">
		<!-- 顶部信息-->
			<span>
				欢迎您：<xsl:value-of select="dv_space/forum_user/userinfo/@username"/>
				<xsl:choose>
					<xsl:when test="dv_space/forum_user/userinfo/@userid != 0">
						| <a href="logout.asp" class="menu0">退出</a>
						<xsl:if test="$canedit_space=1">
							| <a href="#" onclick="spacetoolber();" class="menu0">自定义</a>
						</xsl:if>
					</xsl:when>
					<xsl:otherwise>
						| <a href="login.asp" class="menu0">登陆</a>
					</xsl:otherwise>
				</xsl:choose>
				| <a href="userspace.asp?sid={dv_space/forum_user/userinfo/@userid}" title="访问自已的个性首页" class="menu0">个性首页</a>
				| <a href="index.asp" class="menu0">返回论坛</a>
			</span>
	</xsl:template>
	<xsl:template name="space_head">
		<!--顶部-->
			<div class="title" id="space_title">
				<xsl:value-of select="dv_space/space_info/@title"/>
			</div>
			<div class="intro" id="space_intro">
				<xsl:value-of select="dv_space/space_info/@intro"/>
			</div>
	</xsl:template>
	<xsl:template name="space_menus">
		<!--菜单部分-->
		<ul>
			<li>
				<a href="?sid={dv_space/space_info/@userid}" class="menu0">
					<span>空间首页</span>
				</a>
			</li>
			<li>
				<a href="?sid={dv_space/space_info/@userid}&amp;act=userinfo" class="menu0">
					<span>个人信息</span>
				</a>
			</li>
			<li>
				<a href="?sid={dv_space/space_info/@userid}&amp;act=topic" class="menu0">
					<span>个人主题</span>
				</a>
			</li>
			<li>
				<a href="?sid={dv_space/space_info/@userid}&amp;act=reply" class="menu0">
					<span>参与的评论</span>
				</a>
			</li>

			<li>
				<a href="?sid={dv_space/space_info/@userid}&amp;act=board" class="menu0">
					<span>收藏的论坛</span>
				</a>
			</li>
			<li>
				<a href="?sid={dv_space/space_info/@userid}&amp;act=modifyset" class="menu0">
					<span>个人空间管理</span>
				</a>
			</li>
		</ul>
	</xsl:template>
	<xsl:template name="space_footer">
		<!--底部-->
		<div class="copyright">
			<xsl:value-of select="dv_space/forum_info/@type"  disable-output-escaping="yes"/>,<xsl:value-of select="dv_space/forum_info/@copyright"  disable-output-escaping="yes"/>
			<br/>
			页面执行时间 <xsl:value-of select="dv_space/forum_info/@runtime"/> 秒, <xsl:value-of select="dv_space/forum_info/@querynum"/> 次数据查询。<xsl:value-of select="dv_space/forum_info/@powered"  disable-output-escaping="yes"/>
		</div>
		<div class="intro"></div>
		<script>
			RssList.load();
		</script>
	</xsl:template>
	<xsl:template name="space_main">
		<!--中间主体部分-->
			<ul class="column" id="LeftChannal" style="{$leftchannal_style}">
				<xsl:apply-templates select="dv_space/space_info/leftchannal/channals"></xsl:apply-templates>
			</ul>
			<ul class="column" id="MainChannal" style="{$mainchannal_style}">
				<xsl:choose>
					<xsl:when test="dv_space/forum_info/@act='userinfo'">
						<xsl:call-template name="forum_userinfo"/>
					</xsl:when>
					<xsl:when test="dv_space/forum_info/@act='topic'">
						<xsl:call-template name="topiclist"/>
					</xsl:when>
					<xsl:when test="dv_space/forum_info/@act='reply'">
						<xsl:call-template name="replylist"/>
					</xsl:when>
					<xsl:when test="dv_space/forum_info/@act='modifyset'">
						<xsl:call-template name="modify_setting"/>
					</xsl:when>
					<xsl:when test="dv_space/forum_info/@act='modifyskin'">
						<xsl:call-template name="modify_skins"/>
					</xsl:when>
					<xsl:when test="dv_space/forum_info/@act='board'">
						<xsl:call-template name="template_userboard"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:apply-templates select="dv_space/space_info/mainchannal/channals"></xsl:apply-templates>
					</xsl:otherwise>
				</xsl:choose>
			</ul>
			<ul class="column" id="RightChannal" style="{$rightchannal_style}">
				<xsl:apply-templates select="dv_space/space_info/rightchannal/channals"></xsl:apply-templates>
			</ul>
	</xsl:template>

	<xsl:template match="channals">
		<xsl:choose>
			<xsl:when test="@id='userinfo'">
				<xsl:call-template name="template_userinfo"/>
			</xsl:when>
			<xsl:when test="@id='userfav'">
				<xsl:call-template name="template_userfav"/>
			</xsl:when>
			<xsl:when test="@id='userfriend'">
				<xsl:call-template name="template_userfriend"/>
			</xsl:when>
			<xsl:when test="@id='usermsg'">
				<xsl:call-template name="template_usermsg"/>
			</xsl:when>
			<xsl:when test="@id='usertopic'">
				<xsl:call-template name="template_usertopic"/>
			</xsl:when>
			<xsl:when test="@id='userupload'">
				<xsl:call-template name="template_userupload"/>
			</xsl:when>
			<xsl:when test="@id='userreply'">
				<xsl:call-template name="template_userreply"/>
			</xsl:when>
			<xsl:when test="@id='userbest'">
				<xsl:call-template name="template_userbest"/>
			</xsl:when>
			<xsl:when test="@id='userboard'">
				<xsl:call-template name="template_userboard"/>
			</xsl:when>
			<xsl:when test="@id='foruminfo'">
				<xsl:call-template name="template_foruminfo"/>
			</xsl:when>
			<xsl:when test="substring-before(@id,'_')='mod'">
				<xsl:call-template name="template_module"/>
			</xsl:when>
			
		</xsl:choose>
	</xsl:template>
</xsl:stylesheet>