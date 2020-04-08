<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" omit-xml-declaration = "yes" indent="yes" version="4.0"/>
	<!--
	Copyright (C) 2004,2005 AspSky.Net. All rights reserved.
	Written by dvbbs.net Sunwin
	-->
	<xsl:template  name="modify_tools">
		<div id="spacetool" style="height:30px;text-align:left;">
			<input type="button" name="tools" class="button" value="基本设置" onclick="window.location='?sid={dv_space/space_info/@userid}&amp;act=modifyset'"/>
			<input type="button" name="tools" class="button" value="自定义风格" onclick="window.location='?sid={dv_space/space_info/@userid}&amp;act=modifyskin'"/>
		</div>
	</xsl:template>
	<xsl:template  name="modify_skins">
		<xsl:call-template name="modify_tools"/>
		<form method="post" name="myspace" action="?act=saveskin">
			<table class="spacetable" cellspacing="1" cellpadding="0" align="center" width="98%">
				<tr>
					<th colspan="2">风格DIY</th>
				</tr>
				<tr>
					<td class="border0" width="70%">
						<ol>
							<li>修改CSS样式，定制自已的空间风格；</li>
							<li>点击右边的文件管理，可以添加gif,jpg,png,swf,jpeg,bmp到个人空间，单个文件不能超过100K，文件数限制为50个；</li>
							<li>风格修改完成后，可以点击右边的推荐按钮，推荐自已的风格；</li>
							<li>推荐的风格经管理员审核后将出现在风格模板列表；</li>
						</ol>
					</td>
					<td class="border0" width="30%">
						<input type="button" class="button" value="文件管理" style="height:60px;width:80px;" onclick="window.open('Dv_plus/myspace/script/filemange.asp','DvMyspace_filemange','width=600px,height=550px,resizable=0,scrollbars=yes,menubar=no,status=yes');"/>
						<input type="button" class="button" value="推荐" style="height:60px;width:50px;" onclick="showdisplay('set1');"/>
					</td>
				</tr>
			</table>
			<table id="set1" class="spacetable" cellspacing="3" cellpadding="4" align="center" width="98%" style="display:none;">
				<tr>
					<th colspan="2">推荐风格资料</th>
				</tr>
				<tr>
					<td class="border1" width="30%" align="right">风格名称：</td>
					<td class="border1" width="70%">
						<input type="text"  name="skinname" size="40" value=""/>
					</td>
				</tr>
				<tr>
					<td class="border1" width="30%" align="right">是否推荐：</td>
					<td class="border1" width="70%">
						<input type="radio" name="addtoskins" value="0" checked="true"/>
						<input type="radio" name="addtoskins" value="1"/>
					</td>
				</tr>
			</table>
			<table class="spacetable" cellspacing="1" cellpadding="0" align="center" width="98%">
				<tr>
					<th colspan="2">样式风格CSS修改：</th>
				</tr>
				<tr>
					<td class="border1" colspan="2" align="center">
						<textarea name="s_css" id="s_css"  style="width:98%;height:80px;">
							<xsl:value-of select="dv_space/space_info/@s_css"/>
						</textarea>
						<img src="dv_plus/myspace/drag/minus.gif" unselectable="on" class="cleargif" onclick="textarea_size(-200,'s_css');"/>
						<img src="dv_plus/myspace/drag/plus.gif" unselectable="on" class="cleargif" onclick="textarea_size(200,'s_css');"/>
					</td>
				</tr>
				<tr>
					<td class="border2" colspan="2" align="center">
						<input type="submit" name="submit" value="保 存 设 置"/>
					</td>
				</tr>
			</table>
		</form>
	</xsl:template>
	<xsl:template  name="modify_setting">
		<!--个性空间设置-->
		<xsl:call-template name="modify_tools"/>
		<table class="spacetable" cellspacing="1" cellpadding="0" align="center" width="98%">
			<form method="post" name="myspace" action="?act=saveedit">
				<input type="hidden" name="modify" value="1"/>
			<tr>
				<th colspan="2">基本资料</th>
			</tr>
			<tr>
				<td class="border1" width="30%" align="right">空间标题：</td>
				<td class="border1" width="70%">
					<input type="text"  name="spacetitle" size="40" value="{dv_space/space_info/@title}"/>
				</td>
			</tr>
			<tr>
				<td class="border1" align="right">简介说明：</td>
				<td class="border1">
					<textarea name="spaceintro" style="width:90%;height:50px;">
						<xsl:value-of select="dv_space/space_info/@intro"/>
					</textarea>
				</td>
			</tr>
			<tr>
				<th colspan="2">界面布局</th>
			</tr>
			<tr>
				<td class="border1" colspan="2" align="center">
					<img src="dv_plus/myspace/drag/layout1.gif" class="imgonclick" alt=""/>
					<input type="radio" name="layoutset" value="1"/>左侧边栏
					<img src="dv_plus/myspace/drag/layout3.gif" class="imgonclick" alt=""/>
					<input type="radio" name="layoutset" value="3"/>双侧边栏
					<img src="dv_plus/myspace/drag/layout2.gif" class="imgonclick" alt=""/>
					<input type="radio" name="layoutset" value="2"/>右侧边栏
				</td>
			</tr>
			<tr>
				<th colspan="2">个性设置</th>
			</tr>
			<tr>
				<td class="border1" align="right">作为默认首页：<br/>（作为访问论坛的默认首页）</td>
				<td class="border1">
					<input type="radio" name="ismyindex" value="0"/>否
					<input type="radio" name="ismyindex" value="1"/>是
				</td>
			</tr>
			<tr>
				<td class="border1" align="right">
					首页缓存间隔时间：
				</td>
				<td class="border1">
					<select name="set_5">
						<option value="15">15分钟</option>
						<option value="30">30分钟</option>
						<option value="60">1小时</option>
						<option value="180">3小时</option>
						<option value="360">6小时</option>
						<option value="720">12小时</option>
						<option value="1440">24小时</option>
					</select>
				</td>
			</tr>
			<tr>
				<td class="border1" align="right">
					首页我的主题显示记录数：
				</td>
				<td class="border1">
					<select name="set_0">
						<option value="5">5</option>
						<option value="10">10</option>
						<option value="15">15</option>
						<option value="20">20</option>
						<option value="30">30</option>
					</select>
				</td>
			</tr>
			<tr>
				<td class="border1" align="right">
					首页我的精华显示记录数：
				</td>
				<td class="border1">
					<select name="set_1">
						<option value="5">5</option>
						<option value="10">10</option>
						<option value="15">15</option>
						<option value="20">20</option>
						<option value="30">30</option>
					</select>
				</td>
			</tr>
			<tr>
				<td class="border1" align="right">
					首页我的评论显示记录数：
				</td>
				<td class="border1">
					<select name="set_2">
						<option value="5">5</option>
						<option value="10">10</option>
						<option value="15">15</option>
						<option value="20">20</option>
						<option value="30">30</option>
					</select>
				</td>
			</tr>
				<tr>
				<td class="border1" align="right">
					首页我的短信显示记录数：
				</td>
				<td class="border1">
					<select name="set_8">
						<option value="5">5</option>
						<option value="10">10</option>
						<option value="15">15</option>
						<option value="20">20</option>
						<option value="30">30</option>
					</select>
				</td>
			</tr>
				<tr>
				<td class="border1" align="right">
					首页我的收藏显示记录数：
				</td>
				<td class="border1">
					<select name="set_9">
						<option value="5">5</option>
						<option value="10">10</option>
						<option value="15">15</option>
						<option value="20">20</option>
						<option value="30">30</option>
					</select>
				</td>
			</tr>
				<tr>
				<td class="border1" align="right">
					首页我的好友显示记录数：
				</td>
				<td class="border1">
					<select name="set_10">
						<option value="5">5</option>
						<option value="10">10</option>
						<option value="15">15</option>
						<option value="20">20</option>
						<option value="30">30</option>
					</select>
				</td>
			</tr>
			<tr>
				<td class="border1" align="right">
					首页我的文件显示记录数：
				</td>
				<td class="border1">
					<select name="set_3">
						<option value="5">5</option>
						<option value="10">10</option>
						<option value="15">15</option>
						<option value="20">20</option>
						<option value="30">30</option>
					</select>
				</td>
			</tr>
			<tr>
				<td class="border1" align="right">
					首页我的文件显示类型：
				</td>
				<td class="border1">
					<input type="radio" name="set_4" value="0"/>图片
					<input type="radio" name="set_4" value="1"/>FLASH
					<input type="radio" name="set_4" value="2"/>多媒体
					<input type="radio" name="set_4" value="3"/>其它
				</td>
			</tr>
			<tr>
				<td class="border1" align="right">
					个人主题显示记录数：
				</td>
				<td class="border1">
					<select name="set_6">
						<option value="5">5</option>
						<option value="10">10</option>
						<option value="15">15</option>
						<option value="20">20</option>
						<option value="30">30</option>
					</select>
				</td>
			</tr>
			<tr>
				<td class="border1" align="right">
					参与的评论显示记录数：
				</td>
				<td class="border1">
					<select name="set_7">
						<option value="5">5</option>
						<option value="10">10</option>
						<option value="15">15</option>
						<option value="20">20</option>
						<option value="30">30</option>
					</select>
				</td>
			</tr>
			<tr>
				<td class="border2" colspan="2" align="center">
					<input type="submit" name="submit" value="保 存 设 置"/>
				</td>
			</tr>
			</form>
		</table>
		<SCRIPT LANGUAGE="JavaScript">
			chkradio(document.myspace.ismyindex,"<xsl:value-of select="dv_space/space_user/@set4"/>");
			chkradio(document.myspace.layoutset,"<xsl:value-of select="dv_space/space_info/@s_stylevalue"/>");
			chkradio(document.myspace.set_4,"<xsl:value-of select="dv_space/space_info/@set_4"/>");
			ChkSelected(document.myspace.set_0,"<xsl:value-of select="dv_space/space_info/@set_0"/>");
			ChkSelected(document.myspace.set_1,"<xsl:value-of select="dv_space/space_info/@set_1"/>");
			ChkSelected(document.myspace.set_2,"<xsl:value-of select="dv_space/space_info/@set_2"/>");
			ChkSelected(document.myspace.set_3,"<xsl:value-of select="dv_space/space_info/@set_3"/>");
			ChkSelected(document.myspace.set_5,"<xsl:value-of select="dv_space/space_info/@set_5"/>");
			ChkSelected(document.myspace.set_6,"<xsl:value-of select="dv_space/space_info/@set_6"/>");
			ChkSelected(document.myspace.set_7,"<xsl:value-of select="dv_space/space_info/@set_7"/>");
			ChkSelected(document.myspace.set_8,"<xsl:value-of select="dv_space/space_info/@set_8"/>");
			ChkSelected(document.myspace.set_9,"<xsl:value-of select="dv_space/space_info/@set_9"/>");
			ChkSelected(document.myspace.set_10,"<xsl:value-of select="dv_space/space_info/@set_10"/>");
		</SCRIPT>
	</xsl:template>
	
	<xsl:template name="forum_userinfo">
		<!--用户信息显示-->
		<script language = "javaScript" src = "inc/star.js" type="text/javascript"></script>
		<xsl:if test="dv_space/space_info/@isadmin=-1 or /dv_space/space_user/@set1='1'">
			<table class="spacetable" cellspacing="1" cellpadding="0" align="center" width="98%">
				<tr>
					<th colspan="4">基本资料</th>
				</tr>
				<tr>
					<td class="border1" width="15%" align="right">性 别：</td>
					<td class="border1" width="35%">
						<xsl:choose>
							<xsl:when test="/dv_space/space_user/@usersex=1">
								先生
							</xsl:when>
							<xsl:otherwise>
								女士
							</xsl:otherwise>
						</xsl:choose>
					</td>
					<td class="border1" width="15%" align="right">出 生：</td>
					<td class="border1" width="35%">
						<xsl:value-of select="/dv_space/space_user/@userbirthday"/>
					</td>
				</tr>
				<tr>
					<td class="border1" align="right" width="20%">星 座：</td>
					<td class="border1">
						<script>
							document.write(astro('<xsl:value-of select="/dv_space/space_user/@userbirthday"/>'));
						</script>
					</td>
					<td class="border1" align="right">Email：</td>
					<td class="border1">
						<xsl:value-of select="/dv_space/space_user/@useremail"/>
					</td>
				</tr>
				<tr>
					<td class="border1" align="right">腾讯ＱＱ：</td>
					<td class="border1">
						<xsl:value-of select="/dv_space/space_user/qq"/>
					</td>
					<td class="border1" align="right">ＩＣＱ：</td>
					<td class="border1">
						<xsl:value-of select="/dv_space/space_user/icq"/>
					</td>
				</tr>
				<tr>
					<td class="border1" align="right">ＭＳＮ：</td>
					<td class="border1">
						<xsl:value-of select="/dv_space/space_user/msn"/>
					</td>
					<td class="border1" align="right">ＹＡＨＯＯ：</td>
					<td class="border1">
						<xsl:value-of select="/dv_space/space_user/yahoo"/>
					</td>
				</tr>
				<tr>
					<td class="border1" align="right">ＡＩＭ：</td>
					<td class="border1">
						<xsl:value-of select="/dv_space/space_user/aim"/>
					</td>
					<td class="border1" align="right">ＵＣ：</td>
					<td class="border1">
						<xsl:value-of select="/dv_space/space_user/uc"/>
					</td>
				</tr>
				<tr>
					<td class="border1" align="right">主页：</td>
					<td class="border1" colspan="3" >
						<xsl:value-of select="/dv_space/space_user/homepage"/>
					</td>
				</tr>
			</table>
		</xsl:if>
		<xsl:if test="dv_space/space_info/@isadmin=-1 or /dv_space/space_user/@set2='1'">
			<table class="spacetable" cellspacing="1" cellpadding="0" align="center">
				<tr>
					<th colspan="4">用户详细资料</th>
				</tr>
				<tr>
					<td class="border1" width="20%" align="right">真实姓名：</td>
					<td class="border1" width="30%">
						<xsl:value-of select="/dv_space/space_user/realname"/>
					</td>
					<td class="border1" width="20%" align="right">国　　家：</td>
					<td class="border1" width="30%">
						<xsl:value-of select="/dv_space/space_user/contry"/>
					</td>
				</tr>
				<tr>
					<td class="border1" align="right">联系电话：</td>
					<td class="border1">
						<xsl:value-of select="/dv_space/space_user/phone"/>
					</td>
					<td class="border1" align="right">省　　份：</td>
					<td class="border1" >
						<xsl:value-of select="/dv_space/space_user/province"/>
					</td>
				</tr>
				<tr>
					<td class="border1" align="right">通信地址：</td>
					<td class="border1">
						<xsl:value-of select="/dv_space/space_user/address"/>
					</td>
					<td class="border1" align="right">城　　市：</td>
					<td class="border1">
						<xsl:value-of select="/dv_space/space_user/city"/>
					</td>
				</tr>
				<tr>
					<td class="border1" align="right" width="20%">生　　肖：</td>
					<td class="border1">
						<xsl:value-of select="/dv_space/space_user/shengxiao"/>
					</td>
					<td class="border1" align="right">血　　型：</td>
					<td class="border1">
						<xsl:value-of select="/dv_space/space_user/blood"/>
					</td>
				</tr>
				<tr>
					<td class="border1" align="right">信　　仰：</td>
					<td class="border1">
						<xsl:value-of select="/dv_space/space_user/belief"/>

					</td>
					<td class="border1" align="right">职　　业：</td>
					<td class="border1">
						<xsl:value-of select="/dv_space/space_user/occupation"/>
					</td>
				</tr>
				<tr>
					<td class="border1" align="right">婚姻状况：</td>
					<td class="border1">
						<xsl:value-of select="/dv_space/space_user/marital"/>
					</td>
					<td class="border1" align="right">最高学历：</td>
					<td class="border1">
						<xsl:value-of select="/dv_space/space_user/education"/>
					</td>
				</tr>
				<tr>
					<td class="border1" align="right" width="20%">毕业院校：</td>
					<td class="border1">
						<xsl:value-of select="/dv_space/space_user/college"/>
					</td>
					<td class="border1" align="right">
					</td>
					<td class="border1">
					</td>
				</tr>
			</table>
		</xsl:if>
		<table class="spacetable" cellspacing="1" cellpadding="0" align="center">
			<tr>
				<th colspan="4">用户论坛信息</th>
			</tr>
			<tr>
				<td class="border1" width="20%" align="right">经验值：</td>
				<td class="border1" width="30%">
					<xsl:value-of select="/dv_space/space_user/@userep"/>
				</td>
				<td class="border1" width="20%" align="right">精华帖子：</td>
				<td class="border1" width="30%">
					<xsl:value-of select="/dv_space/space_user/@userisbest"/>
				</td>
			</tr>
			<tr>
				<td class="border1" align="right">魅力值：</td>
				<td class="border1">
					<xsl:value-of select="/dv_space/space_user/@usercp"/>
				</td>
				<td class="border1" align="right">帖子总数：</td>
				<td class="border1" >
					<xsl:value-of select="/dv_space/space_user/@userpost"/>
				</td>
			</tr>
			<tr>
				<td class="border1" align="right">威望值：</td>
				<td class="border1">
					<xsl:value-of select="/dv_space/space_user/@userpower"/>
				</td>
				<td class="border1" align="right">被删主题：</td>
				<td class="border1">
					<xsl:value-of select="/dv_space/space_user/@userdel"/>
				</td>
			</tr>
			<tr>
				<td class="border1" align="right" width="20%">论坛等级：</td>
				<td class="border1">
					<xsl:value-of select="/dv_space/space_user/@userclass"/>
				</td>
				<td class="border1" align="right">被删除率：</td>
				<td class="border1">
					<xsl:value-of select="/dv_space/space_user/@userdelpercent"/>
				</td>
			</tr>
			<tr>
				<td class="border1" align="right">注册日期：</td>
				<td class="border1">
					<xsl:value-of select="/dv_space/space_user/@joindate"/>
				</td>
				<td class="border1" align="right">上次登录：</td>
				<td class="border1">
					<xsl:value-of select="/dv_space/space_user/@lastlogin"/>
				</td>
			</tr>
		</table>

	</xsl:template>
</xsl:stylesheet>