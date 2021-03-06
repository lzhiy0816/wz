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
<!-- 圈子管理页面 -->
<table cellpadding="4" cellspacing="1" class="tableborder1" align="center">
	<tr >
		<th height="24" align="center" colspan="2">欢迎进入圈子管理页面</th>
	</tr>
	<tr >
		<td width="200" class="tablebody1,menu" valign="top">
		<div class="info">
		<ul class="info">
<!--
		<div><img src="skins/Default/nofollow.gif"/><a href="?managetype=info&amp;action=info&amp;groupid={IndivGroup/@groupid}">圈子基本信息</a>[<a href="?managetype=info&amp;action=infoedit&amp;groupid={IndivGroup/@groupid}">编辑</a>]</div>
		<div style="margin-top:4px;"><img src="skins/Default/plus.gif"/><a href="?managetype=board&amp;action=board&amp;groupid={IndivGroup/@groupid}">栏目管理</a></div>
		<div style="margin-left:10px;margin-top:4px;"><img src="skins/Default/nofollow.gif"/><a href="?managetype=board&amp;action=boardmanage&amp;groupid={IndivGroup/@groupid}">添加栏目</a></div>
		<div style="margin-top:4px;"><img src="skins/Default/plus.gif"/><a href="?managetype=user&amp;action=user&amp;groupid={IndivGroup/@groupid}">成员管理</a></div>
		<div style="margin-top:4px;"><img src="skins/Default/nofollow.gif"/><a href="?managetype=updatedata&amp;action=updatedata&amp;groupid={IndivGroup/@groupid}">更新圈子数据</a></div>
-->
		<li><a href="?managetype=info&amp;action=info&amp;groupid={IndivGroup/@groupid}">圈子基本信息</a>[<a href="?managetype=info&amp;action=infoedit&amp;groupid={IndivGroup/@groupid}">编辑</a>]</li>
		<li><a href="?managetype=board&amp;action=board&amp;groupid={IndivGroup/@groupid}">栏目管理</a></li>
		<ul>
			<li><a href="?managetype=board&amp;action=boardmanage&amp;groupid={IndivGroup/@groupid}">添加栏目</a></li>
		</ul>
		<li><a href="?managetype=user&amp;action=user&amp;groupid={IndivGroup/@groupid}">成员管理</a></li>
		<li><a href="?managetype=updatedata&amp;action=updatedata&amp;groupid={IndivGroup/@groupid}">更新圈子数据</a></li>
		</ul>
		</div>
		</td>
		<td class="tablebody1">
		<xsl:choose>
			<xsl:when test="IndivGroup/@action='info'">
				<xsl:call-template name="Info"/>
			</xsl:when>
			<xsl:when test="IndivGroup/@action='infoedit'">
				<xsl:call-template name="InfoManage"/>
			</xsl:when>
			<xsl:when test="IndivGroup/@action='board'">
				<xsl:call-template name="Board"/>
			</xsl:when>
			<xsl:when test="IndivGroup/@action='boardmanage'">
				<xsl:call-template name="BoardManage"/>
			</xsl:when>
			<xsl:when test="IndivGroup/@action='user'">
				<xsl:call-template name="User"/>
			</xsl:when>
			<xsl:when test="IndivGroup/@action='usermanage'">
				<xsl:call-template name="UserManage"/>
			</xsl:when>
			<xsl:when test="IndivGroup/@action='updatedata'">
				<xsl:call-template name="UserDateData"/>
			</xsl:when>
			<xsl:otherwise><xsl:call-template name="Info"/></xsl:otherwise>
		</xsl:choose>
		</td>
	</tr>
</table>
<BR/>
<iframe style="border:0px;width:0px;height:0px;" src="" name="hiddenframe"></iframe>
</xsl:template>

<xsl:template name="Info">
<!-- 圈子基本信息显示 -->
<xsl:for-each select="IndivGroup/Info">
<table cellpadding="3" cellspacing="1" style="width:100%;word-break:break-all;">
	<tr>
		<td width="100%" height="24" colspan="6" class="tablebody2"><b> 圈子基本信息 </b></td>
	</tr>
	<tr>
		<td width="10%" height="24" class="tablebody2">圈子名称：</td>
		<td width="30%" class="tablebody2"><xsl:value-of select="@groupname"/></td>
		<td width="10%" class="tablebody2">今日新帖：</td>
		<td width="20%" class="tablebody2"><font color="red"><xsl:value-of select="@todaynum"/></font></td>
		<td width="10%" class="tablebody2">成员上限：</td>
		<td width="20%" class="tablebody2"><xsl:value-of select="@limituser"/></td>
	</tr>
	<tr>
		<td height="24" class="tablebody2">创 建 人：</td>
		<td class="tablebody2"><xsl:value-of select="@appusername"/></td>
		<td class="tablebody2">主题帖子：</td>
		<td class="tablebody2"><xsl:value-of select="@topicnum"/></td>
		<td class="tablebody2">成员加入：</td>
		<td class="tablebody2">
		<xsl:choose>
			<xsl:when test="@topicnum=1">需要审核</xsl:when>
			<xsl:otherwise>自由加入</xsl:otherwise>
		</xsl:choose>
		</td>
	</tr>
	<tr>
		<td height="24" class="tablebody2">管理成员：</td>
		<td class="tablebody2">
		<xsl:for-each select="/IndivGroup/MasterList/Master">
			<a href="dispuser.asp?id={@userid}" target="_black"><xsl:value-of select="@username"/></a>
			<xsl:if test="position()!=last()">, </xsl:if>
		</xsl:for-each>
		</td>
		<td class="tablebody2">帖子总数：</td>
		<td class="tablebody2"><xsl:value-of select="@postnum"/></td>
		<td class="tablebody2">圈子状态：</td>
		<td class="tablebody2">
		<xsl:call-template name="IndivGroupStatsStr">
			<xsl:with-param name="Sid"><xsl:value-of select="@stats"/></xsl:with-param>
		</xsl:call-template>
		</td>
	</tr>
	<tr>
		<td height="24" class="tablebody2" colspan="2">圈子简介：</td>
		<td height="24" class="tablebody2">创建时间：</td>
		<td class="tablebody2"><xsl:value-of select="@passdate"/></td>
		<td height="24" class="tablebody2"><xsl:value-of select="' '"/></td>
		<td class="tablebody2"><xsl:value-of select="' '"/></td>
	</tr>
	<tr>
		<td colspan="6" class="tablebody2"><xsl:value-of select="@groupinfo" disable-output-escaping="yes"/></td>
	</tr>
</table>
</xsl:for-each>
</xsl:template>

<xsl:template  name="InfoManage">
<!-- 圈子基本信息编辑 -->
<xsl:for-each select="IndivGroup/Info">
<table cellpadding="4" cellspacing="1" style="width:100%;word-break:break-all;">
	<tr>
		<td width="100%" height="24" class="tablebody2"><b>基本信息编辑</b></td>
	</tr> 
	<tr>
		<td class="tablebody2" id="infoForm">
		<form method="post" action="?managetype=info&amp;action=infosave" name="infoForm" onsubmit="return submitform();" target="hiddenframe">
		<div style="padding-left:12px;padding-top:6px;">圈子名称：<input type="hidden" name="groupid" value="{@id}"/><input type="text" name="groupname" value="{@groupname}" size="40"/><span id="groupnamestr"></span></div>
		<div style="padding-left:12px;padding-top:6px">
			<div style="float:left">圈子简介：</div>
			<textarea name="groupinfo" rows="4" cols="40"><xsl:value-of select="@groupinfo"/></textarea>
		</div>
		<xsl:if test="/IndivGroup/@powerflag &lt; 3">
			<div style="padding-left:12px;padding-top:6px;">
			圈子转让：<select name="appuserid">
				<xsl:for-each select="/IndivGroup/MasterList/Master">
					<option value="{@userid}"><xsl:value-of select="@username"/></option>
				</xsl:for-each>
			</select>
			<script language="javascript">ChkSelected(document.infoForm.appuserid,'<xsl:value-of select="@appuserid"/>');</script>
			</div>
		</xsl:if>
		<div style="padding-left:12px;padding-top:6px;">圈子状态：<input type="radio" value="1" name="groupstats"/>正常
			<input type="radio" value="2" name="groupstats"/>锁定
			<input type="radio" value="3" name="groupstats"/>关闭
		</div>
		<div style="padding-left:12px;padding-top:6px;">成员加入：<input type="radio" value="0" name="groupsetting"/>自由加入
			<input type="radio" value="1" name="groupsetting"/>需要审核
		</div>
		<div style="padding-left:12px;padding-top:6px;">浏览设置：<input type="radio" value="0" name="viewflag"/>公开
			<input type="radio" value="1" name="viewflag"/>不公开
		</div>
		<div style="padding-left:12px;padding-top:6px;">（公开：任何人都可以浏览；不公开：只有成员能浏览）（公开：任何人都可以浏览；不公开：只有成员能浏览）</div>
		<div style="padding-left:72px;padding-top:6px;padding-bottom:8px"><input type="submit" name="Submit" value="保  存"/></div>
		</form>
		</td>
	</tr>
</table>
<script language="JavaScript">
chkradio(document.infoForm.groupstats,'<xsl:value-of select="@stats"/>');
chkradio(document.infoForm.groupsetting,'<xsl:value-of select="@locked"/>');
chkradio(document.infoForm.viewflag,'<xsl:value-of select="@viewflag"/>');
function submitform(){
	var formobject=document.infoForm;
	var i,groupInfo,groupStats;
	if (formobject.groupname.value==''){
		document.getElementById('groupnamestr').innerHTML='<font color="#ff0000">←您忘记填写圈子名称了</font>';
		formobject.groupname.focus();
		return false;
	}else{
		if (formobject.groupname.value.length <xsl:value-of select="'&gt;'" disable-output-escaping="yes"/> 50){
			document.getElementById('groupnamestr').innerHTML='<font color="#ff0000">←圈子名称长度不能超过50个字符。</font>';
			formobject.groupname.focus();
			return false;
		}
	}
}
</script>
</xsl:for-each>
</xsl:template>

<xsl:template  name="Board">
<table cellpadding="4" cellspacing="1" style="width:100%;height:100px;word-break:break-all;">
	<tr>
		<td width="100%" height="24" class="tablebody2"><b>栏目列表</b></td>
	</tr> 
	<tr>
		<td class="tablebody2">
		<xsl:for-each select="IndivGroup/BoardList/Board">
			<div id="board_{@id}">
			<a href="IndivGroup_Index.asp?GroupID={@rootid}&amp;GroupBoardid={@id}" target="_black"><xsl:value-of select="@boardname"/></a>
			(今日评论数：<xsl:value-of select="@todaynum"/>，评论主题：<xsl:value-of select="@topicnum"/>，评论总数：<xsl:value-of select="@postnum"/>)
			[<a href="?managetype=board&amp;action=boardmanage&amp;groupid={@rootid}&amp;groupboardid={@id}">编辑</a>]
			[<a href="" onclick="return submitdelete({@id},'{@boardname}');">删除</a>]
			</div>
		</xsl:for-each>
		</td>
	</tr>
</table>
<script language="JavaScript">
function submitdelete(boardid,boardname){
	if(confirm("您确定删除栏目“"+boardname+"”吗?")){
		document.getElementById("hiddenframe").src = <xsl:value-of select="'&quot;?managetype=board&amp;action=boarddelete&amp;groupid='" disable-output-escaping="yes"/><xsl:value-of select="IndivGroup/@groupid"/><xsl:value-of select="'&amp;groupboardid=&quot;+boardid'" disable-output-escaping="yes"/>;
	}
	return false;
}
</script>
</xsl:template>

<xsl:template  name="BoardManage">
<table cellpadding="4" cellspacing="1" style="width:100%;word-break:break-all;">
	<tr>
		<td width="100%" height="24" class="tablebody2"><b>栏目添加/编辑</b></td>
	</tr>
	<tr>
		<td class="tablebody2" id="boardForm">
		<form method="post" name="boardForm" onsubmit="return submitform();">
		<input type="hidden" name="groupboardid" value="{IndivGroup/Board/@id}"/>
		<div style="padding-left:12px;padding-top:6px;">
			栏目名称：<input type="text" name="boardname" value="{IndivGroup/Board/@boardname}" size="40"/><span id="boardnamestr"></span>
		</div>
		<div style="padding-left:12px;padding-top:6px">
			<div style="float:left">栏目说明：</div>
			<textarea name="boardinfo" rows="4" cols="40"><xsl:value-of select="IndivGroup/Board/@boardinfo"/></textarea>
		</div>
		<div style="padding-left:12px;padding-top:6px;">
			主题图片：<input type="text" name="indeximg" value="{IndivGroup/Board/@indeximg}" size="40"/>
		</div>
		<div style="padding-left:12px;padding-top:6px">
			<div style="float:left">栏目规则：</div>
			<textarea name="boardrules" rows="4" cols="40"><xsl:value-of select="IndivGroup/Board/@rules"/></textarea>
		</div>
		<div style="padding-left:12px;padding-top:6px;">
			栏目状态：<input type="radio" value="1" name="boardstats"/>开放<input type="radio" value="0" name="boardstats"/>锁定
		</div>
		<div style="padding-left:70px;padding-top:6px;padding-bottom:8px"><input type="submit" name="Submit" value="保 存"/></div>
		</form>
		</td>
	</tr>
</table>
<script language="JavaScript">
chkradio(document.boardForm.boardstats,'<xsl:value-of select="IndivGroup/Board/@boardstats"/>');
function submitform(){
	var formobject=document.boardForm;
	var i,boardInfo,boardRules,boardStats;
	if (formobject.boardname.value==''){
		document.getElementById('boardnamestr').innerHTML='<font color="#ff0000">←您忘记填写栏目名称了</font>';
		formobject.boardname.focus();
	}else{
		boardInfo = formobject.boardinfo.value.replace(/\r\n/g,'<BR />');
		boardRules = formobject.boardrules.value.replace(/\r\n/g,'<BR />');
		for (i=0;i <xsl:value-of select="'&lt;'" disable-output-escaping="yes"/> formobject.boardstats.length;i++){
			if (formobject.boardstats[i].checked==true){
				boardStats=formobject.boardstats[i].value;
				break;
			}
		}
		document.getElementById("hiddenframe").src = <xsl:value-of select="'&quot;?managetype=board&amp;action=boardsave&amp;groupid='" disable-output-escaping="yes"/><xsl:value-of select="IndivGroup/@groupid"/><xsl:value-of select="'&amp;groupboardid=&quot;+formobject.groupboardid.value+&quot;&amp;boardname=&quot;+formobject.boardname.value+&quot;&amp;boardinfo=&quot;+boardInfo+&quot;&amp;indeximg=&quot;+formobject.indeximg.value+&quot;&amp;boardrules=&quot;+boardRules+&quot;&amp;boardstats=&quot;+boardStats'" disable-output-escaping="yes"/>;
		//document.write(document.getElementById("hiddenframe").src);
	}
	return false;
}
</script>
</xsl:template>

<xsl:template  name="User">
<script language="JavaScript" src="inc/Pagination.js"></script>
<div style="padding-left:12px;padding-bottom:6px;width:100%">
	<input type="button" class="button" name="action" value="全部成员" onclick="checkqueryform(0);"/>
	<input type="button" class="button" name="action" value="待审核成员" onclick="checkqueryform(1);"/>
	<input type="button" class="button" name="action" value="普通成员" onclick="checkqueryform(2);"/>
	<input type="button" class="button" name="action" value="管理员" onclick="checkqueryform(3);"/>
</div>
<table cellpadding="4" cellspacing="1" align="center" style="width:98%;word-break:break-all;">
	<tr>
		<th height="27px" align="left">
		<div style="padding-left:12px;width:100%;">
			<div style="float:left;width:200px;">成员名称</div>
			<div style="float:left;width:200px;">成员状态</div>
			<div>操　　作</div>
		</div>
		</th>
	</tr>
	<tr>
		<td class="tablebody2" id="userForm">
		<xsl:for-each select="IndivGroup/UserList/User">
		<div style="padding-left:12px;width:100%">
			<div style="float:left;width:200px;"><a href="Dispuser.asp?id={@userid}" target="_blank"><xsl:value-of select="@username"/></a></div>
			<div style="float:left;width:200px;">
				<xsl:call-template name="UserStatsStr">
					<xsl:with-param name="Sid"><xsl:value-of select="@islock"/></xsl:with-param>
				</xsl:call-template>
			</div>
			<div>
				<xsl:if test="@islock=0">
				<a href="?{/IndivGroup/@PageSearch}&amp;page={/IndivGroup/@Page}&amp;groupuserid={@id}&amp;action=passuser" target="hidden">通过</a> |
				</xsl:if>
				<a href="?{/IndivGroup/@PageSearch}&amp;page={/IndivGroup/@Page}&amp;groupuserid={@id}&amp;action=usermanage">修改</a> | 
				<a href="?{/IndivGroup/@PageSearch}&amp;page={/IndivGroup/@Page}&amp;groupuserid={@id}&amp;action=deleteuser" target="hidden">删除</a> |
				<a href="?{/IndivGroup/@PageSearch}&amp;page={/IndivGroup/@Page}&amp;groupuserid={@id}&amp;action=setadmin" target="hidden">设置成管理员</a>
			</div>
			<xsl:if test="@intro != ''">
			<div style="margin-top:10px;">
				<div style="float:left;">成员备注：</div>
				<div><xsl:value-of select="@intro" disable-output-escaping="yes"/></div>
			</div>
			</xsl:if>
		</div>
		<div style="padding-left:12px;width:98%;height:1px;"><hr/></div>
		</xsl:for-each>
		<xsl:if test="not(IndivGroup/UserList/User)">没有搜索到任何用户数据！</xsl:if>
		</td>
	</tr>
</table>
<div style="padding-left:12px;width:98%;height:1px;">
<SCRIPT language="javascript">
PageList('<xsl:value-of select="IndivGroup/@Page"/>',3,'<xsl:value-of select="IndivGroup/@MaxRows"/>','<xsl:value-of select="IndivGroup/@CountNum"/>','<xsl:value-of select="IndivGroup/@PageSearch"/><xsl:value-of select="'&amp;action=user'" disable-output-escaping="yes"/>',1);
function checkqueryform(v){
	window.location.href='<xsl:value-of select="'?managetype=user&amp;groupid='" disable-output-escaping="yes"/><xsl:value-of select="IndivGroup/@groupid"/><xsl:value-of select="'&amp;action=user&amp;query='" disable-output-escaping="yes"/>'+v;
}
</SCRIPT>
</div>
</xsl:template>

<xsl:template  name="UserManage">
<table cellpadding="4" cellspacing="1" style="width:100%;word-break:break-all;">
	<tr>
		<td width="100%" height="24" class="tablebody2"><b>成员信息编辑</b></td>
	</tr>
	<tr>
		<td class="tablebody2" id="userForm">
		<xsl:for-each select="IndivGroup/User">
		<form method="post" name="userForm" onsubmit="return submitform();">
		<input type="hidden" name="groupuserid" value="{@id}"/>
		<div>成员名称：<input type="text" name="username" value="{@username}" size="20" readonly="true"/></div>
		<div>用户状态：<input type="radio" value="1" name="userstats"/>普通成员
			<input type="radio" value="2" name="userstats"/>管理员
			<input type="radio" value="0" name="userstats"/>审核
		</div>
		<div style="padding-top:6px;">
			<div style="float:left">成员备注：</div>
			<textarea name="userintro" rows="4" cols="40"><xsl:value-of select="@intro"/></textarea>
		</div>
		<div style="padding-left:65px;"><input type="submit" name="Submit" value="保 存"/></div>
		</form>
		<script language="JavaScript">
		chkradio(document.userForm.userstats,'<xsl:value-of select="@islock"/>');
		function submitform(){
			var formobject=document.userForm;
			var i,userIntro,userStats;
			userIntro = formobject.userintro.value.replace(/\r\n/g,'<BR />');
			for (i=0;i <xsl:value-of select="'&lt;'" disable-output-escaping="yes"/> formobject.userstats.length;i++){
				if (formobject.userstats[i].checked==true){
					userStats=formobject.userstats[i].value;
					break;
				}
			}
			document.getElementById("hiddenframe").src = <xsl:value-of select="'&quot;?managetype=user&amp;action=usersave&amp;groupid='" disable-output-escaping="yes"/><xsl:value-of select="/IndivGroup/@groupid"/><xsl:value-of select="'&amp;groupuserid=&quot;+formobject.groupuserid.value+&quot;&amp;userintro=&quot;+userIntro+&quot;&amp;userstats=&quot;+userStats+&quot;&amp;page='" disable-output-escaping="yes"/><xsl:value-of select="/IndivGroup/@Page"/><xsl:value-of select="'&quot;'" disable-output-escaping="yes"/>;
			return false;
		}
		</script>
		</xsl:for-each>
		<xsl:if test="not(IndivGroup/User)"><font color="red">错误：你要编辑的成员不存在或被删除.</font></xsl:if>
		</td>
	</tr>
</table>
</xsl:template>

<xsl:template  name="UserDateData">
<table cellpadding="4" cellspacing="1" style="width:100%;word-break:break-all;">
	<tr>
		<td width="100%" height="24" class="tablebody2"><b>圈子数据更新</b></td>
	</tr>
	<tr>
		<td class="tablebody2" height="60"><xsl:value-of select="IndivGroup/@InfoStr" disable-output-escaping="yes"/></td>
	</tr>
</table>
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

<xsl:template name="UserStatsStr">
<xsl:param name="Sid"/>
<xsl:choose>
	<xsl:when test="$Sid=1">普通成员</xsl:when>
	<xsl:when test="$Sid=2">管理员</xsl:when>
	<xsl:when test="$Sid=0">审核</xsl:when>
	<xsl:otherwise>未知</xsl:otherwise>
</xsl:choose>
</xsl:template>

<xsl:template  name="dddd">
</xsl:template>
</xsl:stylesheet>