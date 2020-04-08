<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" omit-xml-declaration = "yes" indent="yes" version="4.0"/>
	<!--
	Copyright (C) 2004,2005 AspSky.Net. All rights reserved.
	Written by dvbbs.net Sunwin
	-->
	<xsl:template  name="showpic">
		<xsl:choose>
			<xsl:when test="@isbest=1">
				<img border="0" src="{../@bestpic}" alt="含精华的主题"/>
			</xsl:when>
			<xsl:when test="@locktopic=1">
				<img border="0" src="{../@islockpic}" alt="已锁定的主题"/>
			</xsl:when>
			<xsl:when test="@isvote=1">
				<img border="0" src="{../@votepic}" alt="带投票的主题"/>
			</xsl:when>
			<xsl:when test="@child &gt; 20">
				<img border="0" src="{../@hotpic}" alt="热门的主题"/>
			</xsl:when>
			<xsl:otherwise>
				<img border="0" src="{../@openpic}" alt="开放的主题"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="template_module">
		<!--扩展模块-->
		<xsl:variable name="modid" select="@id"/>
		<li class="module" id="{@id}">
			<div class="mobule_top">
				<h2></h2>
				<h1></h1>
			</div>
			<div class="title" id="moduletitle">
				<img src="{/dv_space/space_info/@skinpath}dblock.gif" alt="" onclick="fold(event,this);" class="block"/>
					<xsl:value-of select="//Module/ModulePrefs[@id=$modid]/@title"/>
			</div>
			<div class="cont">
				<xsl:choose>
					<xsl:when test="//Module/ModulePrefs[@id=$modid]/../Content/@type='rss'">
						<img src="images/loading.gif" border="0"/>
						<br/>
						数据通信加载中...
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="//Module/ModulePrefs[@id=$modid]/../Content" disable-output-escaping="yes"/>
					</xsl:otherwise>
				</xsl:choose>
			</div>
			<div class="mobule_bot">
				<h2></h2>
				<h1></h1>
			</div>
		</li>
		<xsl:if test="//Module/ModulePrefs[@id=$modid]/../Content/@type='rss'">
			<script>
				RssList.url('<xsl:value-of select="@id"/>','<xsl:value-of select="//Module/ModulePrefs[@id=$modid]/../Content"/>');
			</script>
		</xsl:if>
	</xsl:template>

	<xsl:template name="template_userboard">
		<!--收藏的版块-->
		<li class="module" id="{@id}">
		<div class="mobule_top"><h2></h2><h1></h1></div>
			<div class="title" id="moduletitle">
				<img src="{/dv_space/space_info/@skinpath}dblock.gif" alt="" onclick="fold(event,this);" class="block"/>我的版块
			</div>
			<div class="cont">
				<xsl:for-each select="/dv_space/space_user/favbid">
					<xsl:variable name="boardid" select="@bid"/>
					<div class="topic">
						<a href="index.asp?boardid={@bid}" title="浏览该版块">
							<xsl:value-of select="/dv_space/BoardList/board[@boardid=$boardid]/@boardtype"  disable-output-escaping="yes"/>
						</a>
					</div>
					<div class="bottom">
						<div class="bottom1">
							<xsl:value-of select="/dv_space/BoardList/board[@boardid=$boardid]/@readme"  disable-output-escaping="yes"/>
							
						</div>
						<div class="bottom2">
							<xsl:if test="/dv_space/BoardList/board[@boardid=$boardid]/@indeximg!=''">
								<img src="{/dv_space/BoardList/board[@boardid=$boardid]/@indeximg}" border="0"/>
							</xsl:if>
						</div>
					</div>
				</xsl:for-each>
			</div>
		<div class="mobule_bot"><h2></h2><h1></h1></div>
		</li>
	</xsl:template>
	<xsl:template name="template_userinfo">
		<!-- 用户信息-->
		<li class="module" id="{@id}">
		<div class="mobule_top"><h2></h2><h1></h1></div>
			<div class="title" id="moduletitle">
				<img src="{/dv_space/space_info/@skinpath}dblock.gif" alt="" onclick="fold(event,this);" class="block"/>个人信息</div>
			<div class="cont">
				<div class="userface">
					<xsl:choose>
						<xsl:when test="/dv_space/space_user/@userid != 0">
							<img src="{/dv_space/space_user/@userface}" border="0"/>
						</xsl:when>
						<xsl:otherwise>
							<img src="{/dv_space/forum_info/@skinpath}Default/anyon.gif" border="0"/>
						</xsl:otherwise>
					</xsl:choose>
				</div>
				<table class="spacetable" cellspacing="1" cellpadding="0" align="center">
					<tr>
						<td class="border1" width="30%" align="right">名字：</td>
						<td class="border1" width="70%">
							<xsl:value-of select="/dv_space/space_user/@username"/>
						</td>
					</tr>
					<tr>
						<td class="border1" align="right" valign="top">简介：</td>
						<td class="border1" >
							<xsl:value-of select="/dv_space/space_user/personal" disable-output-escaping="yes"/>
						</td>
					</tr>
					<tr>
						<td class="border0" colspan="2" align="center">
							<a href="messanger.asp?action=new&amp;touser={/dv_space/space_user/@username}" target="_blank"><img src="{/dv_space/space_info/@skinpath}message.gif" border="0" alt="给{/dv_space/space_user/@username}发送一个短消息" align="middle" /></a><a href="friendlist.asp?action=addF&amp;myFriend={/dv_space/space_user/@username}" target="_blank"><img src="{/dv_space/space_info/@skinpath}friend.gif" border="0" alt="把{/dv_space/space_user/@username}加入好友" align="middle" /></a><a href="mailto:{/dv_space/space_user/@useremail}"><img alt="点击这里发送电子邮件给{/dv_space/space_user/@username}" border="0" src="{/dv_space/space_info/@skinpath}email.gif" align="middle" /></a>
						</td>
					</tr>
				</table>

				<xsl:value-of select="dv_space/forum_info/@runtime"/>
			</div>
		<div class="mobule_bot"><h2></h2><h1></h1></div>
		</li>
	</xsl:template>
	<xsl:template name="template_userfav">
		<!--个人收藏-->
		<li class="module" id="{@id}">
		<div class="mobule_top"><h2></h2><h1></h1></div>
			<div class="title" id="moduletitle">
				<img src="{/dv_space/space_info/@skinpath}dblock.gif" alt="" onclick="fold(event,this);" class="block"/>个人收藏
			</div>
			<div class="cont">
				<table class="spacetable" cellspacing="1" cellpadding="0" align="center">
					<tr>
						<td class="border2" width="65%" align="center">信息</td>
						<td class="border2" width="35%" align="center">添加时间</td>
					</tr>
					<xsl:for-each select="userfav/row">
						<tr>
							<td class="border1">
								<a href="{@url}" target="_blank">
									<xsl:value-of select="@topic" disable-output-escaping="yes"/>
								</a>
							</td>
							<td class="border1" align="right">
								<xsl:value-of select="@addtime"/>
							</td>
						</tr>
					</xsl:for-each>
				</table>
			</div>
		<div class="mobule_bot"><h2></h2><h1></h1></div>
		</li>
	</xsl:template>
	<xsl:template name="template_userfriend">
		<!--我的好友-->
		<li class="module" id="{@id}">
		<div class="mobule_top"><h2></h2><h1></h1></div>
			<div class="title" id="moduletitle">
				<img src="{/dv_space/space_info/@skinpath}dblock.gif" alt="" onclick="fold(event,this);" class="block"/>我的好友
			</div>
			<div class="cont">
				<table class="spacetable" cellspacing="1" cellpadding="0" align="center">
					<tr>
						<td class="border2" width="35%" align="center">分类</td>
						<td class="border2" width="65%" align="center">姓名</td>
					</tr>
					<xsl:for-each select="userfriend/row">
						<xsl:variable name="f_mod" select="@f_mod"/>
						<tr>
							<td class="border1" align="center">
								<xsl:value-of select="/dv_space/space_user/userfav[@m=$f_mod]/@name" />
							</td>
							<td class="border1">
								<a href="dispuser.asp?name={@f_friend}" target="_blank">
								<xsl:value-of select="@f_friend"/>
								</a>
							</td>
						</tr>
					</xsl:for-each>
				</table>
			</div>
		<div class="mobule_bot"><h2></h2><h1></h1></div>
		</li>
	</xsl:template>
	<xsl:template name="template_usermsg">
		<!--我的短信-->
		<li class="module" id="{@id}">
		<div class="mobule_top"><h2></h2><h1></h1></div>
			<div class="title" id="moduletitle">
				<img src="{/dv_space/space_info/@skinpath}dblock.gif" alt="" onclick="fold(event,this);" class="block"/>我的短信
			</div>
			<div class="cont">
				<table class="spacetable" cellspacing="1" cellpadding="0" align="center">
					<tr>
						<td class="border1" width="35%" align="center" height="30px" colspan="2">
							当前共有（<xsl:value-of select="count(usermsg/row[@flag=0])" />）条新短信未阅！
						</td>
					</tr>
					<!--只允许该用户浏览-->
					<xsl:if test="/dv_space/space_info/@userid = /dv_space/forum_user/userinfo/@userid">
						<tr>
							<td class="border2" width="60%" align="center">标题</td>
							<td class="border2" width="40%" align="center">发送者</td>
						</tr>
						<xsl:for-each select="usermsg/row">
							<tr>
								<td class="border1" align="left">
									<a href="javascript:openScript('messanger.asp?action=read&amp;id={@id}&amp;sender=admin',500,400)" title="浏览短信">
										<xsl:value-of select="@title" />
									</a>
								</td>
								<td class="border1">
									<a href="dispuser.asp?name={@sender}" target="_blank" title="浏览该用户资料"><xsl:value-of select="@sender"/></a>
									<a href="javascript:openScript('messanger.asp?action=new&amp;touser=admin&amp;id={@id}',500,400)" title="回复该短信">
										<img src="{/dv_space/space_info/@skinpath}re_mail.gif"/>
									</a>
								</td>
							</tr>
						</xsl:for-each>
					</xsl:if>
				</table>
			</div>
		<div class="mobule_bot"><h2></h2><h1></h1></div>
		</li>
	</xsl:template>
	<xsl:template name="topiclist">
		<!--主题列表-->
		<script language="JavaScript" src="inc/Pagination.js"></script>
		<div width="100%"><img src="{/dv_space/space_info/@skinpath}warning.gif"/>点击状态图标，可以打开新窗口浏览该信息！ </div>
		<table class="spacetopic" cellspacing="1" cellpadding="0" align="center" width="100%">
			<tr>
				<th width="8%">状态</th>
				<th width="17%">版块</th>
				<th width="48%">标题</th>
				<th width="17%">回复/点击</th>
				<th width="20%">时间</th>
			</tr>
			<xsl:for-each select="dv_space/space_info/mainchannal/topic/row">
				<xsl:variable name="boardid" select="@boardid"/>
				<tr>
					<td class="border2">
						<a href="dispbbs.asp?boardID={@boardid}&amp;ID={@topicid}" target="_blank">
							<xsl:call-template name="showpic"/>
						</a>
					</td>
					<td class="border2">
						<a href="index.asp?boardid={@boardid}">
							<xsl:value-of select="/dv_space/BoardList/board[@boardid=$boardid]/@boardtype" disable-output-escaping="yes"/>
						</a>
					</td>
					<td class="border2" align="left">
						<xsl:choose>
							<xsl:when test="contains(@expression,'|')">
								<img src="Skins/Default/topicface/{substring-after(@expression,'|')}" class="listexpression" alt="" />
							</xsl:when>
							<xsl:otherwise>
								<img src="Skins/Default/topicface/{@expression}" class="listexpression" alt="" />
							</xsl:otherwise>
						</xsl:choose>
						<xsl:if test="@lastpost_4!=''">
							<img src="{/dv_space/forum_info/@picurl}filetype/{@lastpost_4}.gif" width="16" height="16" class="filetype" />
						</xsl:if>
						<a href="dispbbs.asp?boardID={@boardid}&amp;ID={@topicid}" title="《{@title}》&#xA;作者：{@postusername}&#xA;发表于：{@dateandtime}&#xA;最后发贴：{@lastposttime}">
							<xsl:value-of select="@title" disable-output-escaping="yes"/>
						</a>
					</td>
					<td class="border2">
						<xsl:value-of select="@child"/>/<xsl:value-of select="@hits"/>
						<xsl:if test="@isvote= 1 ">
							/<font color="red">
								<b>
									<xsl:value-of select="@votetotal" />
								</b>
							</font>
						</xsl:if>
					</td>
					<td class="border2">
						<xsl:value-of select="@dateandtime"/>
					</td>
				</tr>
			</xsl:for-each>
		</table>
		<xsl:if test="dv_space/space_info/mainchannal/topic/@page!=''">
			<script>
				PageList(<xsl:value-of select="dv_space/space_info/mainchannal/topic/@page"/>,10,<xsl:value-of select="dv_space/space_info/mainchannal/topic/@maxrows"/>,<xsl:value-of select="dv_space/space_info/mainchannal/topic/@countnum"/>,"sid=<xsl:value-of select="dv_space/space_info/@userid"/>&amp;act=topic",2);
			</script>
		</xsl:if>
	</xsl:template>
	<xsl:template name="template_usertopic">
		<!--我的主题-->
		<li class="module" id="{@id}">
			<div class="mobule_top">
				<h2></h2>
				<h1></h1>
			</div>
			<div class="title" id="moduletitle">
				<img src="{/dv_space/space_info/@skinpath}dblock.gif" alt="" onclick="fold(event,this);" class="block"/>我的主题(<xsl:value-of select="topic/@countnum"/>)
			</div>
			<div class="cont">
				<xsl:for-each select="topic/row">
					<xsl:variable name="boardid" select="@boardid"/>
					<div class="topic">
						<a href="dispbbs.asp?boardID={@boardid}&amp;ID={@topicid}" target="_blank" title="点击开新窗口浏览">
							<xsl:call-template name="showpic"/>
						</a>
						<a href="dispbbs.asp?boardID={@boardid}&amp;ID={@topicid}" title="《{@title}》&#xA;作者：{@postusername}&#xA;发表于：{@dateandtime}&#xA;最后发贴：{@lastposttime}">
							<xsl:value-of select="@title" disable-output-escaping="yes"/></a>
					</div>
					<div class="bottom">
						<div class="bottom1">
							<a href="index.asp?boardid={@boardid}" title="浏览该版块">
								<xsl:value-of select="/dv_space/BoardList/board[@boardid=$boardid]/@boardtype"  disable-output-escaping="yes"/>
							</a>
							|
							<a href="dispbbs.asp?boardid={@boardid}&amp;id={@topicid}&amp;replyid={@lastpost_1}&amp;skin=1" title="查看最后回复" target="_blank">
							<xsl:value-of select="@lastposttime"/>
							</a>
						</div>
						<div class="bottom2">
							查看(<xsl:value-of select="@hits"/>) | 回复(<xsl:value-of select="@child"/>)
							<xsl:if test="@isvote= 1 ">| 投票(<font color="red"><b><xsl:value-of select="@votetotal" /></b></font>)</xsl:if>
							| <a href="favlist.asp?action=add&amp;BoardID={@boardid}&amp;id={@topicid}" title="将本贴加入论坛收藏夹">收藏</a>
							| <a href="TopicOther.asp?t=7&amp;BoardID={@boardid}&amp;id={@topicid}" title="推荐本贴给好友">推荐</a>
						</div>
					</div>
				</xsl:for-each>
			</div>
		<div class="mobule_bot"><h2></h2><h1></h1></div>
		</li>
	</xsl:template>
	<xsl:template name="template_userupload">
		<!--附件显示-->
		<li class="module" id="{@id}">
		<div class="mobule_top"><h2></h2><h1></h1></div>
			<div class="title" id="moduletitle">
				<img src="{/dv_space/space_info/@skinpath}dblock.gif" alt="" onclick="fold(event,this);" class="block"/>我的图片
			</div>
			<div class="cont">
					<div id="newstitle">
						<div id="newtitle_bot">
							<p class="tophit0">图片</p>
							<p class="tophit0">Flash</p>
							<p class="tophit0">多媒体</p>
							<p class="tophit0">附件</p>
							<div style="clear:both;"></div>
						</div>
						<div id="newsbody">
							<span>
								<!--图片-->
								<xsl:for-each select="userfile/row[@f_type=1]">
									<xsl:variable name="filepath">
										<xsl:choose>
											<xsl:when test="../@ishide=1">
												showimg.asp?Boardid=<xsl:value-of select="@f_boardid"/>&amp;filename=<xsl:value-of select="@f_filename"/>
											</xsl:when>
											<xsl:otherwise>
												<xsl:choose>
													<xsl:when test="@f_viewname!=''">
														<xsl:value-of select="@f_viewname"/>
													</xsl:when>
													<xsl:otherwise>
														<xsl:value-of select="concat(../@filepath,@f_filename)"/>
													</xsl:otherwise>
												</xsl:choose>
											</xsl:otherwise>
										</xsl:choose>
									</xsl:variable>
									<xsl:variable name="title" select="@f_readme"/>
									<div id="up_img">
										<a href="dispbbs.asp?Boardid={@f_boardid}&amp;ID={@rootid}&amp;replyID={@announceid}&amp;skin=1" target="_blank" title="查看相关评论">
											<xsl:choose>
												<xsl:when test="string-length($title) &gt; 10 ">
													<xsl:value-of select="concat(substring(@f_readme,1,8),'...')"/>
												</xsl:when>
												<xsl:otherwise>
													<xsl:value-of select="@f_readme"/>
												</xsl:otherwise>
											</xsl:choose>
										</a>
										<br/>
										<a href="fileshow.asp?boardid={@f_boardid}&amp;id={@f_id}" title="" target="_blank">
											<img src="{$filepath}" class="viewshow" alt="浏览图片"/>
										</a>
										<br/>
									</div>
								</xsl:for-each>
								<div style="clear:both;"></div>
							</span>
							<span>
								<!--flash-->
								<table class="spacetable" cellspacing="1" cellpadding="0" align="center">
									<tr>
										<td class="border2" width="10%" align="center">类型</td>
										<td class="border2" width="40%">附件相关评论</td>
										<td class="border2" width="15%" align="center">下载</td>
										<td class="border2" width="15%" align="center">点击</td>
										<td class="border2" width="20%">时间</td>
									</tr>
									<xsl:for-each select="userfile/row[@f_type=2]">
										<tr>
											<td class="border1" align="center">
												<img src="{/dv_space/forum_info/@picurl}filetype/{@f_filetype}.gif"/>
											</td>
											<td class="border1">
												<a href="dispbbs.asp?Boardid={@f_boardid}&amp;ID={@rootid}&amp;replyID={@announceid}&amp;skin=1" target="_blank" title="查看相关评论">
													<xsl:value-of select="@f_readme"/>
												</a>
											</td>
											<td class="border1" align="center">
												<a href="viewfile.asp?BoardID={@f_boardid}&amp;ID={@f_id}" target="_blank" title="点击下载">下载</a>
											</td>
											<td class="border1" align="center">
												<xsl:value-of select="@f_viewnum"/>
											</td>
											<td class="border1">
												<xsl:value-of select="@f_addtime"/>
											</td>
										</tr>
									</xsl:for-each>
								</table>
							</span>
							<span>
								<!--多媒体-->
								<table class="spacetable" cellspacing="1" cellpadding="0" align="center">
									<tr>
										<td class="border2" width="10%" align="center">类型</td>
										<td class="border2" width="40%">附件相关评论</td>
										<td class="border2" width="15%" align="center">下载</td>
										<td class="border2" width="15%" align="center">点击</td>
										<td class="border2" width="20%">时间</td>
									</tr>
									<xsl:for-each select="userfile/row[@f_type=3 or @f_type=4] ">
										<tr>
											<td class="border1" align="center">
												<img src="{/dv_space/forum_info/@picurl}filetype/{@f_filetype}.gif"/>
											</td>
											<td class="border1">
												<a href="dispbbs.asp?Boardid={@f_boardid}&amp;ID={@rootid}&amp;replyID={@announceid}&amp;skin=1" target="_blank" title="查看相关评论">
													<xsl:value-of select="@f_readme"/>
												</a>
											</td>
											<td class="border1" align="center">
												<a href="viewfile.asp?BoardID={@f_boardid}&amp;ID={@f_id}" target="_blank" title="点击下载">下载</a>
											</td>
											<td class="border1" align="center">
												<xsl:value-of select="@f_viewnum"/>
											</td>
											<td class="border1">
												<xsl:value-of select="@f_addtime"/>
											</td>
										</tr>
									</xsl:for-each>
								</table>
							</span>
							<span>
								<!--附件-->
								<table class="spacetable" cellspacing="1" cellpadding="0" align="center">
									<tr>
										<td class="border2" width="10%" align="center">类型</td>
										<td class="border2" width="40%">附件相关评论</td>
										<td class="border2" width="15%" align="center">下载</td>
										<td class="border2" width="15%" align="center">点击</td>
										<td class="border2" width="20%">时间</td>
									</tr>
									<xsl:for-each select="userfile/row[@f_type=0]">
										<tr>
											<td class="border1" align="center">
												<img src="{/dv_space/forum_info/@picurl}filetype/{@f_filetype}.gif"/>
											</td>
											<td class="border1">
												<a href="dispbbs.asp?Boardid={@f_boardid}&amp;ID={@rootid}&amp;replyID={@announceid}&amp;skin=1" target="_blank" title="查看相关评论">
													<xsl:value-of select="@f_readme"/>
												</a>
											</td>
											<td class="border1" align="center">
												<a href="viewfile.asp?BoardID={@f_boardid}&amp;ID={@f_id}" target="_blank" title="点击下载">下载</a> </td>
											<td class="border1" align="center">
												<xsl:value-of select="@f_viewnum"/>
											</td>
											<td class="border1">
												<xsl:value-of select="@f_addtime"/>
											</td>
										</tr>
									</xsl:for-each>
								</table>
							</span>
						</div>
						<div style="clear:both;"></div>
					</div>
			</div>
		<div class="mobule_bot"><h2></h2><h1></h1></div>
		</li>
		<script>
			var new1 = new NewsSpanBar();
			new1.f=<xsl:value-of select="/dv_space/space_info/@set_4"/>;
			new1.titleid = "newtitle_bot";
			new1.bodyid = "newsbody";
			new1.load();
		</script>
		<div style="clear:both;"></div>
	</xsl:template>
	
	<xsl:template name="replylist">
		<!--回复评论列表-->
		<script language="JavaScript" src="inc/Pagination.js"></script>
		<div width="100%">
			<img src="{/dv_space/space_info/@skinpath}warning.gif"/> 点击状态图标，可以打开新窗口浏览该信息！
		</div>
		<table class="spacetopic" cellspacing="1" cellpadding="0" align="center" width="100%">
			<tr>
				<th width="8%">状态</th>
				<th width="17%">版块</th>
				<th width="50%">标题</th>
				<th width="35%">时间</th>
			</tr>
			<xsl:for-each select="dv_space/space_info/mainchannal/reply/row">
				<xsl:variable name="boardid" select="@boardid"/>
				<tr>
					<td class="border2">
						<a href="dispbbs.asp?boardID={@boardid}&amp;ID={@rootid}&amp;replyid={@announceid}&amp;skin=1" target="_blank">
							<xsl:call-template name="showpic"/>
						</a>
					</td>
					<td class="border2">
						<a href="index.asp?boardid={@boardid}">
							<xsl:value-of select="/dv_space/BoardList/board[@boardid=$boardid]/@boardtype" disable-output-escaping="yes"/>
						</a>
					</td>
					<td class="border2" align="left">
						<span style="float:right;color:gray;" title="共有{@length}字符!">
							(<xsl:value-of select="@length"/>)
						</span>
						<img src="Skins/Default/topicface/{@expression}" class="listexpression" alt="" />
						<a href="dispbbs.asp?boardID={@boardid}&amp;ID={@rootid}&amp;replyid={@announceid}&amp;skin=1">
							<xsl:value-of select="@topic" disable-output-escaping="yes"/>
						</a>
					</td>
					<td class="border2">
						<xsl:value-of select="@dateandtime"/>
					</td>
				</tr>
			</xsl:for-each>
		</table>
		<xsl:if test="dv_space/space_info/mainchannal/reply/@page!=''">
			<script>
				PageList(<xsl:value-of select="dv_space/space_info/mainchannal/reply/@page"/>,10,<xsl:value-of select="dv_space/space_info/mainchannal/reply/@maxrows"/>,<xsl:value-of select="dv_space/space_info/mainchannal/reply/@countnum"/>,"sid=<xsl:value-of select="dv_space/space_info/@userid"/>&amp;act=reply",2);
			</script>
		</xsl:if>
	</xsl:template>
	<xsl:template name="template_userreply">
		<li class="module" id="{@id}">
			<div class="mobule_top">
				<h2></h2>
				<h1></h1>
			</div>
			<div class="title" id="moduletitle">
				<img src="{/dv_space/space_info/@skinpath}dblock.gif" alt="" onclick="fold(event,this);" class="block"/>我的评论(<xsl:value-of select="reply/@countnum"/>)
			</div>
			<div class="cont">
				<xsl:for-each select="reply/row">
					<xsl:variable name="boardid" select="@boardid"/>
					<div class="topic">
						<a href="dispbbs.asp?boardID={@boardid}&amp;ID={@rootid}" target="_blank" title="点击开新窗口浏览该主题">
							<xsl:call-template name="showpic"/>
						</a>
						<a href="dispbbs.asp?boardID={@boardid}&amp;ID={@rootid}&amp;replyid={@announceid}&amp;skin=1" title="《{@topic}》&#xA;发表于：{@dateandtime}">
							<xsl:value-of select="@topic" disable-output-escaping="yes"/>
								</a>
							</div>
							<div class="bottom">
								<div class="bottom1">
									<a href="index.asp?boardid={@boardid}" title="浏览该版块">
										<xsl:value-of select="/dv_space/BoardList/board[@boardid=$boardid]/@boardtype"  disable-output-escaping="yes"/>
									</a>
								</div>
								<div class="bottom2">
									字符(<xsl:value-of select="@length"/>)
									| <a href="post.asp?action=re&amp;BoardID={@boardid}&amp;replyID={@announceid}&amp;ID={@rootid}&amp;star=1&amp;reply=true" title="引用回复这个贴子">引用</a>
									| <a href="post.asp?action=re&amp;BoardID={@boardid}&amp;replyID={@announceid}&amp;ID={@rootid}&amp;star=1" title="回复这个贴子">回复</a>
								</div>
							</div>
						</xsl:for-each>
			</div>
		<div class="mobule_bot"><h2></h2><h1></h1></div>
		</li>
	</xsl:template>
	<xsl:template name="template_userbest">
		<li class="module" id="{@id}">
		<div class="mobule_top"><h2></h2><h1></h1></div>
			<div class="title" id="moduletitle">
				<img src="{/dv_space/space_info/@skinpath}dblock.gif" alt="" onclick="fold(event,this);" class="block"/>我的精华
			</div>
			<div class="cont">
				<xsl:for-each select="userbest/row">
					<xsl:variable name="boardid" select="@boardid"/>
					<div class="topic">
						<a href="dispbbs.asp?boardID={@boardid}&amp;ID={@rootid}" target="_blank" title="点击开新窗口浏览该主题">
							<xsl:call-template name="showpic"/>
						</a>
						<a href="dispbbs.asp?boardID={@boardid}&amp;ID={@rootid}&amp;replyid={@announceid}&amp;skin=1" title="《{@title}》&#xA;发表于：{@dateandtime}">
							<xsl:value-of select="@title" disable-output-escaping="yes"/>
						</a>
					</div>
					<div class="bottom">
						<div class="bottom1">
							<a href="index.asp?boardid={@boardid}" title="浏览该版块">
								<xsl:value-of select="/dv_space/BoardList/board[@boardid=$boardid]/@boardtype" />
							</a>
						</div>
						<div class="bottom2">
							<a href="post.asp?action=re&amp;BoardID={@boardid}&amp;replyID={@announceid}&amp;ID={@rootid}&amp;star=1&amp;reply=true" title="引用回复这个贴子">引用</a>
							| <a href="post.asp?action=re&amp;BoardID={@boardid}&amp;replyID={@announceid}&amp;ID={@rootid}&amp;star=1" title="回复这个贴子">回复</a>
						</div>
					</div>
				</xsl:for-each>
			</div>
		<div class="mobule_bot"><h2></h2><h1></h1></div>
		</li>
	</xsl:template>

	<xsl:template name="template_foruminfo">
		<li class="module" id="{@id}">
			<div class="mobule_top">
				<h2></h2>
				<h1></h1>
			</div>
			<div class="title" id="moduletitle">
				<img src="{/dv_space/space_info/@skinpath}dblock.gif" alt="" onclick="fold(event,this);" class="block"/>论坛信息
			</div>
			<div class="cont">
				<div class="body1">
				<table class="spacetable" cellspacing="1" cellpadding="0" align="center">
					<tr>
						<td class="border2" width="45%" align="right">论坛主题：</td>
						<td class="border2" width="55%">
							<xsl:value-of select="/dv_space/forum_info/@topicnum" />篇
						</td>
					</tr>
					<tr>
						<td class="border2" align="right">论坛帖子：</td>
						<td class="border2">
							<xsl:value-of select="/dv_space/forum_info/@postnum" />篇
						</td>
					</tr>
					<tr>
						<td class="border2" align="right">今日帖数：</td>
						<td class="border2">
							<xsl:value-of select="/dv_space/forum_info/@todaynum" />篇
						</td>
					</tr>
					<tr>
						<td class="border2" align="right">昨日帖数：</td>
						<td class="border2">
							<xsl:value-of select="/dv_space/forum_info/@yesterdaynum" />篇
						</td>
					</tr>
					<tr>
						<td class="border2" align="right">最高日发帖：</td>
						<td class="border2">
							<xsl:value-of select="/dv_space/forum_info/@maxpostnum" />篇
						</td>
					</tr>
					<tr>
						<td class="border2" align="right">建站时间：</td>
						<td class="border2">
							<xsl:value-of select="/dv_space/forum_info/@createtime" />
						</td>
					</tr>
				</table>
				</div>
				<div class="body1">
					<table class="spacetable" cellspacing="1" cellpadding="0" align="center">
						<tr>
							<td class="border2" width="45%" align="right">最新会员：</td>
							<td class="border2" width="55%">
								<xsl:value-of select="/dv_space/forum_info/@lastuser" />
							</td>
						</tr>
						<tr>
							<td class="border2" align="right">会员总数：</td>
							<td class="border2">
								<xsl:value-of select="/dv_space/forum_info/@usernum" />人
							</td>
						</tr>
						<tr>
							<td class="border2" align="right">当前在线：</td>
							<td class="border2">
								<xsl:value-of select="/dv_space/forum_info/@online" />人
							</td>
						</tr>
						<tr>
							<td class="border2" align="right">在线会员：</td>
							<td class="border2">
								<xsl:value-of select="/dv_space/forum_info/@useronline" />人
							</td>
						</tr>
						<tr>
							<td class="border2" align="right">在线访客：</td>
							<td class="border2">
								<xsl:value-of select="/dv_space/forum_info/@guestonline" />人
							</td>
						</tr>
						<tr>
							<td class="border2" align="right">最高在线：</td>
							<td class="border2">
								<xsl:value-of select="/dv_space/forum_info/@maxonline" />人
							</td>
						</tr>
					</table>
				</div>
				<div class="space" style="height:0px;"></div>
			</div>
			<div class="mobule_bot">
				<h2></h2>
				<h1></h1>
			</div>
		</li>
	</xsl:template>

</xsl:stylesheet>