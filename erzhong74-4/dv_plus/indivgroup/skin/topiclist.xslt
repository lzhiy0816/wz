<?xml version="1.0" encoding="gb2312"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" >
<xsl:output method="xml" omit-xml-declaration = "yes" indent="yes" version="4.0"/>
	<!--
	Copyright (C) 2004,2005 AspSky.Net. All rights reserved.
	Written by dvbbs.net Lao Mi
	Web: http://www.aspsky.net/,http://www.dvbbs.net/
	Email: eway@aspsky.net
	-->
<!--模板设置-->
<xsl:variable name="showexpression" select="0" /><!--设置为1显示发贴心情图标-->
<xsl:variable name="showtitleinfo" select="1" /><!--设置为1显示标题中悬停的提示内容,就是title属性标签里的内容.-->
<xsl:variable name="showstatslink" select="1" /><!--设置为1贴子状态图表中加入开新窗口的链接.-->
<xsl:variable name="showspage" select="1" /><!--是否显示单贴分页设置为1则显示-->
<xsl:variable name="splittoptiopc" select="0" /><!--是否分离固顶区-->
<xsl:variable name="rulesstyle" select="1" /><!--版规样式 ,0为表格样式,1为fieldset样式-->
<xsl:variable name="showfiletype" select="1" /><!--0不显示附件类型图片-->
<xsl:variable name="rulespage" select="2" /> <!--第几页开始不显示版规内容--> 
<!--模板设置结束-->

<xsl:variable name="myscript">
<![CDATA[
function CheckAll(form){
	for (var i=0;i<form.elements.length;i++){
		var e = form.elements[i];
		if (e.name == 'Announceid') e.checked = form.chkall.checked;
	}
}
function clicksubmit(form){
	if(confirm('您确定执行的操作吗?')){
		this.document.batch.submit();
		return true;
	}else{
		return false;
	}
}
BoardJumpListSelect(0,"newboard","移动帖子请选择","",1)
]]>
</xsl:variable>
	
<xsl:template match="/">
<script language="JavaScript" src="Dv_plus/IndivGroup/js/IndivGroup_Main.js"></script>
<script language="JavaScript" src="inc/Pagination.js"></script>
<xsl:if test="IndivGroup/Board/@rules != '' and IndivGroup/@page &lt; $rulespage">
	<xsl:choose>
		<xsl:when test="$rulesstyle=0">
			<div class="th"><xsl:text disable-output-escaping="yes" >&amp;nbsp;&amp;nbsp;</xsl:text>>>本版规则</div>
			<div class="mainbar" id="rules">
				<div id="rulesbody"><xsl:value-of select="IndivGroup/Board/@rules" disable-output-escaping="yes"/></div>
			</div>
			<br />
		</xsl:when>
		<xsl:otherwise>
			<div class="mainbar0" style="margin-bottom:10px;">
				<div style="padding:0px 10px;">
					<div style="padding:10px;background-color : #f5f5f5;border : 1px dotted #ccc;text-align:left;line-height:18px;">
						<strong><font color="#ff0000">版 规 和 导 读</font></strong><br />
						<xsl:value-of select="IndivGroup/Board/@rules" disable-output-escaping="yes"/>
					</div>
				</div>
			</div>
		</xsl:otherwise>
	</xsl:choose>
</xsl:if>
<div class="main" style="margin-top:4px;margin-bottom:4px;height:28px;line-height:28px;">
	<xsl:if test="IndivGroup/forum_setting/@rss=0 or IndivGroup/forum_setting/@wap=1">
		<div class="tableborder5" style="line-height:18px;height:20px;float:right;margin-right:1px;margin-top:3px;font-size:9px;font-family:tahoma,arial;">
			<div class="tabletitle1" style="float:left;width:25px;margin : 1px; 1px; 1px; 1px;">XML</div>
			<xsl:if test="IndivGroup/forum_setting/@rss=0">
						<div style="float:left;width:45px;margin: 1px 1px 1px 1px ;background-color : #fff;border:1px inset;line-height:16px;">
					<a href="http://rss.iboker.com/sub/?{IndivGroup/forum_setting/@ForumUrl}rssfeed.asp?rssid=4&amp;boardid={IndivGroup/Board/@id}" target="_blank" title="订阅本版面最新帖子" style="color:#000;">RSS 2.0</a>
				</div>
			</xsl:if>
			<xsl:if test="IndivGroup/forum_setting/@wap=1">
				<div class="tabletitle1" style="float:left;width:25px;margin : 1px;">
					<a href="wap.asp?Action=readme" target="_blank" title="通过手机访问论坛，把论坛动起来" style="color:#fff;">WAP</a>
				</div>
			</xsl:if>
		</div>
	</xsl:if>
	<div id="postbutton" class="posttopic"  onclick="return window.location='IndivGroup_Post.asp?action=new&amp;groupid={IndivGroup/Board/@rootid}&amp;groupboardid={IndivGroup/Board/@id}';" title="发表一个新帖子"></div>
</div>
<div class="mainbar4" id="boardmaster">
	<div style="float:left;margin-left:4px;">精华评论</div>
	<xsl:if test="IndivGroup/@powerflag=1 or IndivGroup/@powerflag=2 or IndivGroup/@powerflag=3">
		<div id="boardmanage">管理选项:　<a href="?action=batch&amp;GroupID={IndivGroup/Board/@rootid}&amp;GroupBoardid={IndivGroup/Board/@id}">帖子批量管理</a></div>
	</xsl:if>
</div>
<div class="th">
	<div class="list_r">
		<div class="list1" style="width:80px;">作 者</div>
		<div class="list1" style="width:50px;">回复</div>
		<div class="list1" style="width:50px;">人气</div>
	</div>
	<div class="list1">
	<xsl:choose>
		<xsl:when test="IndivGroup/@powerflag &lt; 4">
			<a href="?groupid={IndivGroup/Board/@rootid}&amp;groupboardid={IndivGroup/Board/@id}&amp;page={IndivGroup/@page}&amp;action=batch">状态</a>
		</xsl:when>
		<xsl:otherwise>状态</xsl:otherwise>
	</xsl:choose>
	</div>
	<div >主 题</div>
</div>
<xsl:choose>
	<xsl:when test="not (IndivGroup/toptopic/row or IndivGroup/topic/row)"><div class="list" style="text-align:center;">本栏目没有评论</div></xsl:when>
	<xsl:otherwise>
		<form action="IndivGroup_PostsManage.asp?groupid={IndivGroup/Board/@rootid}&amp;groupboardid={IndivGroup/Board/@id}" method="post" name="batch">
		<xsl:if test="IndivGroup/toptopic/row">
			<xsl:if test="$splittoptiopc=1"><div class="mainbar3"><hr /></div></xsl:if>
			<xsl:for-each select="IndivGroup/toptopic/row">
				<xsl:call-template name="topic"/>
			</xsl:for-each>
			<xsl:if test="$splittoptiopc=1"><div class="mainbar3"><hr /></div></xsl:if>
		</xsl:if>
		<xsl:for-each select="IndivGroup/topic/row">
			<xsl:call-template name="topic"/>
		</xsl:for-each>
		<xsl:if test="IndivGroup/@action='batch' and IndivGroup/@powerflag &lt; 4">
			<div class="list">
				<input type="hidden" value="{IndivGroup/Board/@id}" name="groupboardid"/>
				<input type="checkbox" name="chkall" value="on" onclick="CheckAll(this.form);" class="chkbox"/>全选/取消 
				<input type="radio" name="action" value="dele" class="radio"/>批量删除
				<input class="radio" type="radio" name="action" value="isbest"/>批量精华
				<input type="radio" name="action" value="lock" class="radio"/>批量锁定
				<input type="submit" name="Submit" value="执行" onclick="return clicksubmit(this.form);"/>
				<script type="text/javascript" language="javascript">
					<xsl:value-of select="$myscript" disable-output-escaping="yes" />
				</script>
			</div >
		</xsl:if>
		</form>
		<xsl:variable name="CountNum"><xsl:value-of select="IndivGroup/Board/@topicnum - count(IndivGroup/toptopic/row)"/></xsl:variable>
		<div class="mainbar0" style="margin-top :2px;height : 26px;">
			<div style="float:right;">
				<!-- 分页信息-->
				<script language="javascript">PageList(<xsl:value-of select="IndivGroup/@page"/>,10,<xsl:value-of select="IndivGroup/forum_setting/@pagesize"/>,<xsl:value-of select="$CountNum"/>,"groupid=<xsl:value-of select="IndivGroup/Board/@rootid"/>&amp;groupboardid=<xsl:value-of select="IndivGroup/Board/@id"/>&amp;action=<xsl:value-of select="/IndivGroup/@action"/>",4);</script>
			</div>
		</div>
	</xsl:otherwise>
</xsl:choose>
<div class="mainbar0" style="margin-top :2px;height : 22px;">
	<div style="float:right;">
		<select name="BoardJumpList" id="BoardJumpList" onchange="if(this.options[this.selectedIndex].value!='')location='?groupid={IndivGroup/Board/@rootid}&amp;groupboardid='+this.options[this.selectedIndex].value;"></select>
	</div>
	<script type="text/javascript" language="javascript">
		IndivGroupBoardJumpListSelect(<xsl:value-of select="IndivGroup/Board/@rootid"/>,'<xsl:value-of select="IndivGroup/Board/@id"/>',"BoardJumpList","跳转到版面....","",0);
	</script>
	<div style="float:left;">
		<form method="post" action="Query.asp">
			<input name="boardid" type="hidden" value="{IndivGroup/boarddata/@boardid}"/>
			<input name="sType" type="hidden" value="2"/>
			<input name="SearchDate" type="hidden" value="30"/>
			<input name="pSearch" type="hidden" value="1"/>
			<input name="nSearch" type="hidden" value="1"/>
			<input name="isWeb" type="hidden" value="1"/>
			<input type="text" name="keyword"/>
			<xsl:text disable-output-escaping="yes" >&amp;nbsp;</xsl:text>
			<input type="submit" name="submit" value="网页搜索"/>
		</form>
	</div>
</div>
<br/>
<div class="th">
	<xsl:text disable-output-escaping="yes" >&amp;nbsp;&amp;nbsp;</xsl:text>-=> <xsl:value-of select="IndivGroup/boarddata/@boardtype" disable-output-escaping="yes"/>
</div>
<div class="mainbar3" style="height:38px;">
	<div style="padding:10px;line-height:20px;">
		<div style="width:15%;float:left;"><img align="absmiddle" src="{IndivGroup/@opentopic}" alt="开放的话题"/> 开放的话题</div>
		<div style="width:15%;float:left;"><img align="absmiddle" src="{IndivGroup/@hottopic}" alt="回复超过{IndivGroup/forum_setting/@ishot}贴"/> 热门的话题</div>
		<div style="width:15%;float:left;"><img align="absmiddle" src="{IndivGroup/@ilocktopic}" alt="锁定的话题"/> 锁定的话题</div>
		<div style="width:15%;float:left;"><img align="absmiddle" src="{IndivGroup/@besttopic}" alt="含精华贴的话题"/> 含精华贴的话题</div>
		<div style="width:15%;float:left;"><img align="absmiddle" src="{IndivGroup/@istopic}" alt="固顶"/>  固顶话题</div>
	</div>
</div>
</xsl:template>

<xsl:template  name="topic">
<!-- 主题列表循环部分-->
<xsl:variable name="endpage">
	<xsl:choose>
		<xsl:when test="(@child+1) mod (/IndivGroup/forum_setting/@dispsize)=0"><xsl:value-of select="(@child+1) div (/IndivGroup/forum_setting/@dispsize)"/></xsl:when>
		<xsl:otherwise><xsl:value-of select="floor((@child) div (/IndivGroup/forum_setting/@dispsize))+1"/></xsl:otherwise>
	</xsl:choose>
</xsl:variable>
<div class="list">
	<div class="list_r1">
		<div class="list_a"><a href="dispuser.asp?id={@postuserid}" target="_blank"><xsl:value-of select="@postusername" /></a></div>
		<div class="list_c"><xsl:value-of select="@child"/></div>
		<div class="list_c"><xsl:value-of select="@hits"/></div>
		<div class="list_t"><a href="IndivGroup_Dispbbs.asp?GroupID={@groupid}&amp;groupboardid={@boardid}&amp;id={@topicid}&amp;star={$endpage}&amp;page={/IndivGroup/@page}#{@lastpost_1}"><xsl:value-of select="@lastpost_2"/></a></div> <font color="#FF0000">|</font> <a href="dispuser.asp?id={@lastpost_5}" target="_blank"><xsl:value-of select="@lastpost_0"/></a>
	</div>
	<div class="list_s"><xsl:call-template name="topicstats"/></div>
	<div class="list_img" id="followImg{@topicid}">
		<xsl:choose>
			<xsl:when test="@child!=0"><img src="{/IndivGroup/@pic_follow}" alt="" class="listimg1"/></xsl:when>
			<xsl:otherwise><img src="{/IndivGroup/@pic_nofollow}" alt="" class="listimg1"/></xsl:otherwise>
		</xsl:choose>
	</div>
	<div class="listtitle" style="margin-top:2px;">
		<xsl:if test="$showexpression=1">
			<xsl:choose>
				<xsl:when test="contains(@expression,'|')"><img src="Skins/Default/topicface/{substring-after(@expression,'|')}" class="listexpression" alt="" /></xsl:when>
				<xsl:otherwise><img src="Skins/Default/topicface/{@expression}" class="listexpression" alt="" /></xsl:otherwise>
			</xsl:choose>
		</xsl:if>
		<xsl:if test="$showfiletype=1">
			<xsl:if test="@lastpost_4!=''"><img src="{/IndivGroup/@picurl}filetype/{@lastpost_4}.gif" width="16" height="16" class="filetype" /></xsl:if>
		</xsl:if>
		<xsl:call-template name="showtitle"/>
	</div>
	<xsl:if test="$showspage = 1 "> 
		<xsl:if test="$endpage &gt; 1 ">
			<div style="float:left;">[ </div>
			<div style="float:left;"><img src="{/IndivGroup/@picurl}pagelist.gif"  alt="" style="margin-top:7px;"/></div>
			<div style="float:left;">
			<xsl:call-template name="pagelist">
				<xsl:with-param name="i" select="2"/>
				<xsl:with-param name="endpage" select="$endpage" />
			</xsl:call-template>
			</div> 
			<div style="float:left;"> ]</div>
		</xsl:if>
	</xsl:if>
	<xsl:if test="/IndivGroup/forum_setting/@newflagpic !='0'  and /IndivGroup/forum_setting/@newflagpic !=''">
		<xsl:if test="@datedifftime"> 
			<div><img src="{/IndivGroup/forum_setting/@newflagpic}" border="0" alt="{@datedifftime}分钟前更新!" style="margin-top: 9px;"/></div>
		</xsl:if>
	</xsl:if>
</div>
<div class="mainbar3" id="follow{@topicid}" style="display:none;text-align : left; " ><div id="followTd{@topicid}" style="padding:2px;"></div></div>
</xsl:template>

<xsl:template  name="topicstats">
<xsl:choose>
	<xsl:when test="/IndivGroup/@action='batch' and /IndivGroup/@powerflag &lt; 4">
		<input type="checkbox" name="topicid" value="{@topicid}" class="chkbox"/>
	</xsl:when>
	<xsl:otherwise>
		<xsl:choose>
			<xsl:when test="$showstatslink = 1"><a href="IndivGroup_Dispbbs.asp?GroupID={@groupid}&amp;groupboardid={@boardid}&amp;ID={@topicid}&amp;page={/IndivGroup/@page}" target="_blank"><xsl:call-template name="showpic"/></a></xsl:when>
			<xsl:otherwise><xsl:call-template name="showpic"/></xsl:otherwise>
		</xsl:choose>
	</xsl:otherwise>
</xsl:choose>
</xsl:template>

<xsl:template  name="showpic">
<xsl:choose>
	<xsl:when test="@istop=1 or @isbest=1">
		<xsl:choose>
			<xsl:when test="@istop=1"><img border="0" src="{/IndivGroup/@istopic}" alt="固顶的主题,点击开新窗口浏览" /></xsl:when>
			<xsl:otherwise><img border="0" src="{/IndivGroup/@besttopic}" alt="含精华的主题,点击开新窗口浏览" /></xsl:otherwise>
		</xsl:choose>
	</xsl:when>
	<xsl:otherwise>
		<xsl:choose>
			<xsl:when test="@locktopic=1">
				<img border="0" src="{/IndivGroup/@ilocktopic}" alt="已锁定的主题,点击开新窗口浏览"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test="@child &gt; /IndivGroup/forum_setting/@ishot">
						<img border="0" src="{/IndivGroup/@hottopic}" alt="热门的主题,点击开新窗口浏览" />
					</xsl:when>
					<xsl:otherwise>
						<img border="0" src="{/IndivGroup/@opentopic}" alt="开放的主题,点击开新窗口浏览" />
					</xsl:otherwise>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:otherwise>
</xsl:choose>
</xsl:template>

<!--显示各种模式的标题-->
<xsl:template name="showtitle">
<a href="IndivGroup_Dispbbs.asp?GroupID={@groupid}&amp;groupboardid={@boardid}&amp;ID={@topicid}&amp;page={/IndivGroup/@page}">
<xsl:if test="$showtitleinfo=1">
	<xsl:attribute name="title">《<xsl:value-of select="@title" disable-output-escaping="yes"/>》
作者：<xsl:value-of select="@postusername"/>
发表于：<xsl:value-of select="@dateandtime"/>
最后发贴：<xsl:value-of select="@lastpost_3"/>
	</xsl:attribute>
</xsl:if>
<xsl:choose>
	<xsl:when test="@topicmode='1'"><xsl:value-of select="@title" disable-output-escaping="yes"/></xsl:when>
	<xsl:otherwise>
		<xsl:choose>
			<xsl:when test="string-length(@title) &gt; 30 ">
				<xsl:value-of select="concat(substring(@title,0,20),'....',substring(@title,(string-length(@title)- 10 ),string-length(@title)))" disable-output-escaping="yes"/>[长]
			</xsl:when>
			<xsl:otherwise><xsl:value-of select="@title" disable-output-escaping="yes"/></xsl:otherwise>
		</xsl:choose>
	</xsl:otherwise>
</xsl:choose>
</a>
<xsl:value-of select="' '" />
</xsl:template>

<xsl:template name="pagelist">
	<xsl:param name="i"/>
	<xsl:param name="endpage"/>
	<xsl:if test="$i &lt; 10 and $i &lt;=$endpage">
		<xsl:value-of select="' '" /><b><a href="IndivGroup_Dispbbs.asp?GroupID={@groupid}&amp;groupboardid={@boardid}&amp;ID={@topicid}&amp;star={$i}&amp;page={/aml/@page}"><font color="red"><xsl:value-of select="$i"/></font></a></b>
		<xsl:call-template name="pagelist">
			<xsl:with-param name="i" select="$i+1"/>
			<xsl:with-param name="endpage" select="$endpage"/>
		</xsl:call-template>
	</xsl:if>
	<xsl:if test="$i = 9 and $endpage &gt; 9">....<b><a href="IndivGroup_Dispbbs.asp?GroupID={@groupid}&amp;groupboardid={@boardid}&amp;ID={@topicid}&amp;star={$endpage}&amp;page={/IndivGroup/@page}"><font color="red"><xsl:value-of select="$endpage"/></font></a></b></xsl:if>
</xsl:template>
</xsl:stylesheet>