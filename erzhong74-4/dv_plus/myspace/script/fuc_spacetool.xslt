<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" omit-xml-declaration = "yes" indent="yes" version="4.0"/>
	<!--
	Copyright (C) 2004,2005 AspSky.Net. All rights reserved.
	Written by dvbbs.net Sunwin
	-->
		<xsl:template name="space_toolbar">
		<!--管理工具-->
			<xsl:if test="/dv_space/space_info/@id=0">
				<style type="text/css">
					/*屏蔽层样式定义*/
					#Shade {
					position:absolute;
					left: 0;
					top: 0;
					width: 100%;
					filter:alpha(opacity=50);
					-moz-opacity:.5;
					opacity:.5;
					display: none;
					background-color: #000;
					z-index: 2007;
					}
					#AutoPostPrompt{
					font-size:12px;
					line-height:18px;
					padding:5px 5px 5px 32px;
					background:#FFFFE6;
					border:2px solid #222;
					text-align:left;
					z-index: 2008;
					display: none;
					position: absolute;
					}
				</style>
				<img src="{/dv_space/space_info/@skinpath}warning.gif"/> 您的个性首页暂未正式开通，请点击导航上的“自定义”，进行修改后开通个性首页频道！	
			</xsl:if>
		<div id="spacetool" style="display:none;">
			<ul id="temppannal" style="display:none;">
				<li class="module" id="[$id]">
					<div class="title">
						<img class="cleargif" src="{/dv_space/space_info/@skinpath}close.gif" alt="" onclick="Drag.del(this,'[$id]')"/> [$pannaltitle]
					</div>
					<div class="cont">[$cont]</div>
				</li>
			</ul>
			<table border="0" cellpadding="3" cellspacing="3" width="100%" height="100%">
				<form method="post" name="myspace" action="?act=saveedit" onsubmit="return saveset();">
					<tr>
						<td width="75%" align="left">
							<!-- 个性布局菜单 -->
							<input type="button" name="layout" class="button" value="选取个性布局" onclick="setmenu(this)"/><div id="div_layout" style="display:none;" class="setmenu">
								<div class="title">
									<img class="cleargif" src="{/dv_space/space_info/@skinpath}close.gif" alt="" onclick="setmenu(parentNode.parentNode.previousSibling)"/>个性布局设置
								</div>
								<div class="content">
									<img src="{/dv_space/space_info/@skinpath}layout1.gif" class="imgonclick" alt="" onclick="setlayout('1')"/>
									<img src="{/dv_space/space_info/@skinpath}layout3.gif" class="imgonclick" alt="" onclick="setlayout('3')"/>
									<img src="{/dv_space/space_info/@skinpath}layout2.gif" class="imgonclick" alt="" onclick="setlayout('2')"/>
								</div>
							</div>
							<!-- 个性布局菜单 -->
							<input type="button" name="layout" class="button" value="选取系统模块"  onclick="setmenu(this)"/><div id="div_layout" style="display:none;" class="setmenu">
								<div class="title">
									<img class="cleargif" src="{/dv_space/space_info/@skinpath}close.gif" alt="" onclick="setmenu(parentNode.parentNode.previousSibling)"/>系统模块
								</div>
								<div class="content" >
									<ul style="float:left">
										<li>
											<input type="checkbox" name="addleft" id="add_foruminfo" value="foruminfo" onclick="Drag.add(this.value,'LeftChannal',this.checked)"/><label>论坛信息</label>
										</li>
										<li>
											<input type="checkbox" name="addleft" id="add_userinfo" value="userinfo" onclick="Drag.add(this.value,'LeftChannal',this.checked)"/><label>个人信息</label>
										</li>
										<li>
											<input type="checkbox" name="addleft" id="add_userfav" value="userfav" onclick="Drag.add(this.value,'LeftChannal',this.checked)"/><label>个人收藏</label>
										</li>
										<li>
											<input type="checkbox" name="addleft" id="add_userfriend" value="userfriend" onclick="Drag.add(this.value,'LeftChannal',this.checked)"/><label>我的好友</label>
										</li>
										<li>
											<input type="checkbox" name="addleft" id="add_usermsg" value="usermsg" onclick="Drag.add(this.value,'LeftChannal',this.checked)"/><label>我的短信</label>
										</li>
									</ul>
									<ul style="float:left">
										<li>
											<input type="checkbox" name="addmain" id="add_userboard" value="userboard" onclick="Drag.add(this.value,'MainChannal',this.checked)"/><label>收藏版块</label>
										</li>
										<li>
											<input type="checkbox" name="addmain" id="add_usertopic" value="usertopic" onclick="Drag.add(this.value,'MainChannal',this.checked)"/><label>我的主题</label>
										</li>
										<li>
											<input type="checkbox" name="addmain" id="add_userbest" value="userbest" onclick="Drag.add(this.value,'MainChannal',this.checked)"/><label>我的精华</label>
										</li>
										<li>
											<input type="checkbox" name="addmain" id="add_userupload" value="userupload" onclick="Drag.add(this.value,'MainChannal',this.checked)"/><label>我的图片</label>
										</li>
										<li>
											<input type="checkbox" name="addmain" id="add_userreply" value="userreply" onclick="Drag.add(this.value,'MainChannal',this.checked)"/><label>我的评论</label>
										</li>
									</ul>
									<div style="clear:both;"></div>
								</div>
							</div>
							<input type="button" name="layout" class="button" value="选取扩展模块" onclick="setmenu(this);getmodules('selmodules',window.location.href);"/><div id="div_layout" class="setmenu" style="display:none;width:460px;" >
								<div class="title">
									<img class="cleargif" src="{/dv_space/space_info/@skinpath}close.gif" alt="" onclick="setmenu(parentNode.parentNode.previousSibling)"/>选取频道模板
								</div>
								<div class="content">
									<iframe id="selmodules" src="" width="100%" height="450px"></iframe>
								</div>
							</div>
							<input type="button" name="layout" class="button" value="风格模板定制" onclick="setmenu(this);getskinsdb('selskins','spaceskin.asp');"/><div id="div_layout" class="setmenu" style="display:none;width:430px;" >
								<div class="title">
									<img class="cleargif" src="{/dv_space/space_info/@skinpath}close.gif" alt="" onclick="setmenu(parentNode.parentNode.previousSibling)"/>选取风格模板
								</div>
								<div class="content">
									<iframe id="selskins" src="" width="100%" height="450px"></iframe>
								</div>
							</div>
						</td>
						<td whdth="25%" rowspan="2">
							<input type="submit" name="dragsubmit" class="button"  value="保存设置"/>
							<button onclick="window.location='userspace.asp?sid={/dv_space/space_info/@userid}';" class="button">退出设置</button>
						</td>
					</tr>
					<tr>
						<td align="left">
							<input type="hidden" name="styleid" id="styleid" value=""/>
							<input type="hidden" name="stylepath" id="stylepath" value="{dv_space/space_info/@s_path}"/>
							<input type="hidden" name="layoutset" id="layoutset" value="{dv_space/space_info/@s_style}"/>
							<input type="hidden" name="layoutleft" id="layoutleft" value="{dv_space/space_info/@s_left}"/>
							<input type="hidden" name="layoutmain" id="layoutmain" value="{dv_space/space_info/@s_center}"/>
							<input type="hidden" name="layoutright" id="layoutright" value="{dv_space/space_info/@s_right}"/>
							<input type="hidden" name="spacetitle" id="val_space_title" value="{dv_space/space_info/@title}"/>
							<input type="hidden" name="spaceintro" id="val_space_intro" value="{dv_space/space_info/@intro}"/>
							<img src="{/dv_space/space_info/@skinpath}warning.gif"/>
							点击个性空间标题和简介内容可以直接修改，若要重新排列模块，只要将其拖放到页面上的其他位置即可，完成更改后，请单击“保存设置”。
						</td>
					</tr>
				</form>
			</table>
		</div>
	</xsl:template>

</xsl:stylesheet>