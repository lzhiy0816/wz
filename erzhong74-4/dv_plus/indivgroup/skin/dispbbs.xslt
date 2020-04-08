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
	<script language="JavaScript" src="Dv_plus/IndivGroup/js/IndivGroup_Main.js"></script>
	<script language="JavaScript" src="inc/Pagination.js"></script>
	<script language="JavaScript" src="inc/Font.js"></script>
	<!--�ж��Ƿ���IE��������Ӧ��style��class-->
	<xsl:variable name="postclass">
		<xsl:choose>
			<xsl:when test="post/agent/@browser='Microsoft Internet Explorer' and not(post/agent/@version > 6 ) ">postie</xsl:when>
			<xsl:otherwise>post</xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	<xsl:variable name="bodystyle">
		<xsl:choose>
			<xsl:when test="post/agent/@browser='Microsoft Internet Explorer' and not(post/agent/@version > 6 )">height:200px;width:97%;padding-right:0px; overflow-x: hidden;</xsl:when>
			<xsl:otherwise>min-height:200px;</xsl:otherwise>
		</xsl:choose>font-size:<xsl:value-of select="post/setting/@fontsize" />pt;line-height:<xsl:value-of select="post/setting/@lineheight" />;text-indent:<xsl:value-of select="post/setting/@indent" />px;
	</xsl:variable>
	<!--end -->
	<script type="text/javascript" language="javascript" src="inc/star.js"></script>
	<div class="main" style="margin-top:4px;height:28px;line-height:28px;margin-bottom:4px;">
		<div class="posttopic" title="����һ��������" onclick="location.href='IndivGroup_Post.asp?action=new&amp;groupid={/post/postinfo/@groupid}&amp;groupboardid={/post/postinfo/@boardid}'"></div>
		<div class="repost" title="�ظ�����" onclick="location.href='IndivGroup_Post.asp?action=revert&amp;groupid={/post/postinfo/@groupid}&amp;groupboardid={/post/postinfo/@boardid}&amp;id={/post/postinfo/@topicid}&amp;page={/post/postinfo/@page}'"></div>
		<div style="float:right;">���Ǳ����ĵ� <b><xsl:value-of select="post/postinfo/@hits"/></b> ���Ķ���</div>
	</div>
	<div class="th">
		<div style="height:24px;text-indent:10px;text-align:left;">
			���⣺
			<xsl:choose>
				<xsl:when test="post/postinfo/@topicmode=0 or post/postinfo/@topicmode=1">
					<xsl:value-of select="post/postinfo/@title" disable-output-escaping="yes"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="post/postinfo/@title"/>
				</xsl:otherwise>
			</xsl:choose>
		</div>
	</div>
	<xsl:for-each select="post/row">
		<a name="{@announceid}" ><xsl:value-of select="''" /></a>
		<xsl:variable name="userid" select="@postuserid"/>
		<div class="postlary{2-(position() mod 2)}">
			<div class="postuserinfo">
				<div style="padding: 10px 0px 0px 5px;line-height:30px;height:30px;">
					<div style="float:left;width :120px;filter:glow(color='{/post/userlist/user[@userid=$userid]/@nameglow}',strength='2');"><xsl:value-of select="substring-before(/post/userlist/user[@userid=$userid]/@namestyle,'��')" disable-output-escaping="yes"/><xsl:value-of select="@username" /><xsl:value-of select="substring-after(/post/userlist/user[@userid=$userid]/@namestyle,'��')" disable-output-escaping="yes"/></div>
					<div style="float:left;width :23px;text-indent:0px;margin:3px;">
					<xsl:choose>
						<xsl:when test="/post/userlist/user[@userid=$userid]/@userhidden='1'">
							<xsl:choose>
								<xsl:when test="/post/userlist/user[@userid=$userid]/@usersex!='1'"><img src="{/post/setting/@picurl}ofFeMale.gif" alt="��Ůѽ�����ߣ����Ը��Ұɣ�"/></xsl:when>
								<xsl:otherwise><img src="{/post/setting/@picurl}ofMale.gif" alt="˧��Ӵ�����ߣ�����������"/></xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:otherwise>
							<xsl:choose>
								<xsl:when test="/post/userlist/user[@userid=$userid]/@usersex!='1'"><img src="{/post/setting/@picurl}FeMale.gif" alt="��Ůѽ�����ߣ��������Ұɣ�"/></xsl:when>
								<xsl:otherwise><img src="{/post/setting/@picurl}Male.gif" alt="˧��Ӵ�����ߣ�����������"/></xsl:otherwise>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
					</div>
					<div style="float:left;width :15px;text-indent:0px;margin:5px;"><script type="text/javascript" language="javascript">document.write (astro('<xsl:value-of select="/post/userlist/user[@userid=$userid]/@userbirthday"/>'));</script></div>
				</div>
				<xsl:variable name="userface" select="/post/userlist/user[@userid=$userid]/@userface"/>
				<xsl:choose>
					<xsl:when test="contains($userface,'|')">
						<div><img src="{substring-after($userface,'|')}" width="{/post/userlist/user[@userid=$userid]/@userwidth}" height="{/post/userlist/user[@userid=$userid]/@userheight}"/></div>
						<xsl:if test="substring-before($userface,'|') != '0'">
							<div><a href="javascript:DispMagicEmot({substring-before($userface,'|')},350,500)">�鿴ħ��ͷ��</a></div>
						</xsl:if>
					</xsl:when>
					<xsl:otherwise>
						<div><img  src="{$userface}" width="{/post/userlist/user[@userid=$userid]/@userwidth}" height="{/post/userlist/user[@userid=$userid]/@userheight}"/></div>
					</xsl:otherwise>
				</xsl:choose>
				<div><img style="margin:5px 0px 5px 0px;" src="{/post/setting/@picurl}star/{/post/userlist/user[@userid=$userid]/@titlepic}"/></div>
				<xsl:if test="/post/setting/@usertitle=1 and /post/userlist/user[@userid=$userid]/@usertitle != ''">
					<div>ͷ�Σ�<xsl:value-of select="/post/userlist/user[@userid=$userid]/@usertitle"/></div>
				</xsl:if>
				<div>�ȼ���<xsl:value-of select="/post/userlist/user[@userid=$userid]/@userclass"/></div>
				<xsl:if test="/post/userlist/user[@userid=$userid]/@userpower != 0">
					<div><font color="red"><b>������<xsl:value-of select="/post/userlist/user[@userid=$userid]/@userpower"/></b></font></div>
				</xsl:if>
				<div>���£�<xsl:value-of select="/post/userlist/user[@userid=$userid]/@userpost"/></div>
				<div>���֣�<xsl:value-of select="/post/userlist/user[@userid=$userid]/@userep"/></div>
				<div>ע�᣺<xsl:value-of select="/post/userlist/user[@userid=$userid]/@joindate"/></div>
			</div>
			<div class="{$postclass}">
				<div>
					<div class="user_menu_info">
					<div style="float:right;margin-top:3px;color:#333;">
					<xsl:choose>
						<xsl:when test="/post/postinfo/@star =1">
							<xsl:choose>
								<xsl:when test="position()=1">¥��</xsl:when>
								<xsl:otherwise>�� <font color="red"><xsl:value-of select="position()"/></font> ¥</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:otherwise>	�� <font color="red"><xsl:value-of select="position() + (/post/postinfo/@star -1) * /post/setting/@pagesize"/></font> ¥</xsl:otherwise>
					</xsl:choose>
					</div>
					<div class="text_style" style="float:right">
						<a href="###" onClick="fontSize('m','textstyle_{@announceid}')">С</a>
						<a href="###" onClick="fontSize('b','textstyle_{@announceid}')">��</a>
					</div>
					<xsl:text disable-output-escaping="yes" >&amp;nbsp;</xsl:text>
					<xsl:if test="(not (/post/userinfo/@cananony=1) or @signflag !=2 or /post/userinfo/@truemaster=1 or /post/userinfo/@userid = @postuserid) and @postuserid!=0">
						<a href="userspace.asp?sid={@postuserid}">������ҳ</a>
						<xsl:if test="/post/setting/@isboke=1">
							 | <a href="boke.asp?UserID={@postuserid}" title="����{@username}�ĸ��˲���" target="_blank">����</a>
						</xsl:if>
						<xsl:if test="/post/userlist/user[@userid=$userid]/@oicq !=''">
							| <a href="http://wpa.qq.com/msgrd?V=1&amp;Uin={/post/userlist/user[@userid=$userid]/@oicq}&amp;Site=By Dvbbs&amp;Menu=yes" title="�������QQ��Ϣ��{@username}" target="_blank">QQ</a>
						</xsl:if>
						<xsl:if test="/post/userinfo/@userid  !=0">
							| <a href="messanger.asp?action=new&amp;touser={@username}" target="_blank">����</a>
							| <a href="friendlist.asp?action=addF&amp;myFriend={@username}" target="_blank">����</a>
							| <a href="dispuser.asp?id={@postuserid}" target="_blank">��Ϣ</a>
							| <a href="query.asp?stype=1&amp;nSearch=3&amp;keyword={@username}&amp;groupid={/post/postinfo/@groupid}&amp;GroupBoardID={/post/postinfo/@boardid}&amp;SearchDate=ALL" target="_blank">����</a>
						</xsl:if>
						<xsl:if test="/post/userlist/user[@userid=$userid]/@useremail !=''">
							| <a href="mailto:{/post/userlist/user[@userid=$userid]/@useremail}">����</a>
						</xsl:if>
						<xsl:if test="/post/userlist/user[@userid=$userid]/@homepage !=''">
							| <a href="{/post/userlist/user[@userid=$userid]/@homepage}" target="_blank">��ҳ</a>
						</xsl:if>
						<xsl:if test="/post/userlist/user[@userid=$userid]/@userim !=''">
							| <a href="http://service.51uc.com/user_info/show_user_info_base.shtml?UID={/post/userlist/user[@userid=$userid]/@userim}" title="{@username}[{/post/userlist/user[@userid=$userid]/@userim}]��UC���" target="_blank">UC</a>
						</xsl:if>
					</xsl:if>
				</div>
				</div>
				<xsl:if test="@locktopic=2">
					<xsl:choose>
						<xsl:when test="/post/userinfo/@truemaster=1"><div class="lockuser">���ݱ�����</div></xsl:when>
						<xsl:otherwise><div style="{$bodystyle}"><div class="lockuser" >���ݱ�����</div></div></xsl:otherwise>
					</xsl:choose>
				</xsl:if>
				<xsl:if test="@locktopic != 2 or /post/userinfo/@truemaster =1">
					<div style="height:22px;line-height:22px;">
						<b><xsl:value-of select="@topic" disable-output-escaping="yes"/></b>
					</div>
					<!-- �������ݲ��� -->
					<!-- �������ݶ�����沿�� -->
					<xsl:if test="position()=1 and /post/postinfo/@star='1' and /post/postinfo/forum_ads[@id=18]/@set='1'">
						<div class="body_adv_top"><xsl:value-of select="/post/postinfo/forum_ads[@id='18']/@code" disable-output-escaping="yes"/></div><br/>
					</xsl:if>
					<div id="textstyle_{@announceid}" style="{$bodystyle}margin-top:10px;word-wrap : break-word ;word-break : break-all ;" onload="this.style.overflowX='auto';">
						<!-- �������ҹ�沿�� -->
						<xsl:if test="position()=1 and /post/postinfo/@star=1">
							<xsl:choose>
								<xsl:when test="/post/postinfo/forum_ads[@id=22]/@set='1'">
									<div class="body_adv_left"><xsl:value-of select="/post/postinfo/forum_ads[@id='22']/@code" disable-output-escaping="yes"/></div>
								</xsl:when>
								<xsl:when test="/post/postinfo/forum_ads[@id=22]/@set='2'">
									<div class="body_adv_right"><xsl:value-of select="/post/postinfo/forum_ads[@id='22']/@code" disable-output-escaping="yes"/></div>
								</xsl:when>
							</xsl:choose>
						</xsl:if>
						<!-- ���������ı� -->
						<xsl:value-of select="@body" disable-output-escaping="yes"/>
						<br />
						<!-- �������ݵײ���沿�� -->
						<xsl:if test="position()=1 and /post/postinfo/@star=1 and /post/postinfo/forum_ads[@id='20']/@set='1'">
							<br/>
							<div class="body_adv_bottom"><xsl:value-of select="/post/postinfo/forum_ads[@id='20']/@code" disable-output-escaping="yes"/></div>
						</xsl:if>
					</div>
					<xsl:if test="/post/userlist/user[@userid=$userid]/@usersign !=''">
						<div style="width:85%;overflow-x: hidden;" id="sigline_{position()}">
							<img src="{/post/setting/@picurl}sigline.gif" />
							<br />
							<xsl:value-of select="/post/userlist/user[@userid=$userid]/@usersign" disable-output-escaping="yes"/>
						</div>
						<script type="text/javascript" language="javascript">
							fixheight('sigline_<xsl:value-of select="position()" />');
						</script>
					</xsl:if>
				</xsl:if>
				<xsl:if test="@isbest='1'">
					<div class="info"><img src="{/post/setting/@picurl}jing.gif" border="0" title="��������Ϊ����" align="absmiddle"/>��������Ϊ����</div>
				</xsl:if>
			</div>
		</div>
		<div class="postbottom{2 - (position() mod 2)}" >
			<xsl:if test="position()=last()"><xsl:attribute name="id">postend</xsl:attribute></xsl:if>
			<div class="postuserinfo" style="height:28px;">
				<div style="float:left;">
					<xsl:choose>
						<xsl:when test="/post/userinfo/@truemaster=1">
							<a href="TopicOther.asp?t=1&amp;groupid={/post/postinfo/@groupid}&amp;GroupBoardID={/post/postinfo/@boardid}&amp;userid={$userid}&amp;ip={@ip}&amp;action=lookip" target="_blank"><img border="0" src="Skins/Default/ip.gif" style="margin-top:4px;" alt="����鿴�û���Դ������������IP��{@ip}" /></a>
						</xsl:when>
						<xsl:otherwise>
							<img src="{/post/setting/@picurl}ip.gif" style="margin: 4px;" alt="ip��ַ�����ñ���" />
						</xsl:otherwise>
					</xsl:choose>
				</div>
				<div style="float:left; margin-left:-20px;"><xsl:value-of select="@dateandtime" /></div>
			</div>
			<div class="postie" style="height:28px;padding:0px;text-indent:10px;">
				<div style="float:right;margin-right:5px;">
					<xsl:if test="/post/userinfo/@userid  !=0">
						<div class="menu_popup" id="Menu_{position()}">
							<div class="menuitems"><xsl:value-of select="''" />
								<xsl:if test="position() != 1">
									<xsl:if test="/post/userinfo/@userid= $userid or /post/userinfo/@truemaster=1">
										<a href="IndivGroup_PostManage.asp?action=deletepost&amp;groupid={/post/postinfo/@groupid}&amp;GroupBoardID={/post/postinfo/@boardid}&amp;replyID={@announceid}&amp;ID={/post/postinfo/@topicid}&amp;star={/post/postinfo/@star}&amp;userid={@postuserid}">ɾ������</a>
										<br />
									</xsl:if>
								</xsl:if>
								<xsl:if test="/post/userinfo/@truemaster =1 ">
									<xsl:choose>
										<xsl:when test="@isbest= 0 ">
											<a href="IndivGroup_PostManage.asp?action=setbest&amp;groupid={/post/postinfo/@groupid}&amp;GroupBoardID={/post/postinfo/@boardid}&amp;replyID={@announceid}&amp;ID={/post/postinfo/@topicid}&amp;star={/post/postinfo/@star}&amp;userid={@postuserid}">��Ϊ����</a>
										</xsl:when>
										<xsl:otherwise>
											<a href="IndivGroup_PostManage.asp?action=unchainbest&amp;groupid={/post/postinfo/@groupid}&amp;GroupBoardID={/post/postinfo/@boardid}&amp;replyID={@announceid}&amp;ID={/post/postinfo/@topicid}&amp;star={/post/postinfo/@star}&amp;userid={@postuserid}">�������</a>
										</xsl:otherwise>
									</xsl:choose>
									<xsl:choose>
										<xsl:when test="@locktopic !=2">
											<br/>
											<a href="IndivGroup_PostManage.asp?action=LockPage&amp;groupid={/post/postinfo/@groupid}&amp;GroupBoardID={/post/postinfo/@boardid}&amp;replyID={@announceid}&amp;ID={/post/postinfo/@topicid}&amp;star={/post/postinfo/@star}&amp;userid={@postuserid}">��������</a>
										</xsl:when>
										<xsl:otherwise>
											<br/>
											<a href="IndivGroup_PostManage.asp?action=UnchainLockPage&amp;groupid={/post/postinfo/@groupid}&amp;GroupBoardID={/post/postinfo/@boardid}&amp;replyID={@announceid}&amp;ID={/post/postinfo/@topicid}&amp;star={/post/postinfo/@star}&amp;userid={@postuserid}">�������</a>
										</xsl:otherwise>
									</xsl:choose>
								</xsl:if>
							</div>
						</div>
						<xsl:if test="(/post/postinfo/@locktopic != 1 and /post/userinfo/@canreply =1)or (/post/userinfo/@truemaster =1 ) ">
							<a href="IndivGroup_Post.asp?action=revert&amp;groupid={/post/postinfo/@groupid}&amp;GroupBoardID={/post/postinfo/@boardid}&amp;PostID={@announceid}&amp;ID={/post/postinfo/@topicid}&amp;star={/post/postinfo/@star}&amp;reply=true">����</a> |
							<a href="IndivGroup_Post.asp?action=revert&amp;groupid={/post/postinfo/@groupid}&amp;GroupBoardID={/post/postinfo/@boardid}&amp;PostID={@announceid}&amp;ID={/post/postinfo/@topicid}&amp;star={/post/postinfo/@star}">�ظ�</a> |
						</xsl:if>
						<xsl:if test="/post/userinfo/@truemaster=1 or (/post/userinfo/@userid= $userid and /post/userinfo/@caneditmypost=1)">
							<a href="IndivGroup_Post.asp?action=edit&amp;groupid={/post/postinfo/@groupid}&amp;GroupBoardID={/post/postinfo/@boardid}&amp;PostID={@announceid}&amp;page={/post/postinfo/@page}">�༭</a> |
							<a href="###" onmouseover="showmenu(event,'','Menu_{position()}');">���Ӳ���</a>
						</xsl:if>
					</xsl:if>
					<a href="#top"><img src="{/post/setting/@picurl}p_up.gif" alt="�ص�����" style="border:0px;"/></a>
				</div>
				<div><xsl:value-of select="@ubblist" disable-output-escaping="yes"/></div>
			</div>
		</div>
	</xsl:for-each>
	<br />
	<div class="mainbar0" style="height:26px;text-align : left;">
		<div  style="height:26px;float:right;">
			<xsl:choose>
				<xsl:when test="post/postinfo/@skin='0'">
				<!--��ҳӦ��-->
				<script language="javascript">PageList(<xsl:value-of select="post/postinfo/@star"/>,10,<xsl:value-of select="post/setting/@pagesize"/>,<xsl:value-of select="post/postinfo/@child+1"/>,"?groupid=<xsl:value-of select="post/postinfo/@groupid"/>&amp;GroupBoardID=<xsl:value-of select="post/postinfo/@boardid"/>&amp;replyid=<xsl:value-of select="post/postinfo/@replyid"/>&amp;id=<xsl:value-of select="post/postinfo/@topicid"/>&amp;skin=<xsl:value-of select="post/postinfo/@skin"/>",3);</script>
				</xsl:when>
				<xsl:otherwise>
					<span id="showpagelist"></span>
				</xsl:otherwise>
			</xsl:choose>
			<span id="showclose"> </span>
		</div>
		<select name="BoardJumpList" id="BoardJumpList" onchange="if(this.options[this.selectedIndex].value!='')location='IndivGroup_Index.asp?groupid={/post/postinfo/@groupid}&amp;GroupBoardID='+this.options[this.selectedIndex].value;" ></select>
	</div>
	<script type="text/javascript" language="javascript">
		IndivGroupBoardJumpListSelect('<xsl:value-of select="post/postinfo/@groupid"/>','<xsl:value-of select="post/postinfo/@boardid"/>',"BoardJumpList","��̳��ת��....","",0);
	</script>
	<xsl:variable name="postuserid" select="post/postinfo/@postuserid"/>
	<xsl:choose>
		<xsl:when test="post/userinfo/@truemaster =1">
			<xsl:call-template name="managemenu" />
		</xsl:when>
		<xsl:otherwise>
			<xsl:if test="post/userinfo/@userid = $postuserid ">
				<xsl:call-template name="managemenume" />
			</xsl:if>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>
<xsl:template name="checkuser">
	<xsl:param name="userid"/>
	<xsl:choose>
		<xsl:when test="@locktopic=2">���ݱ�����</xsl:when>
		<xsl:otherwise>
			<xsl:choose>
				<xsl:when test="/post/userlist/user[@userid=$userid]/@lockuser=1">�û��ѱ�����</xsl:when>
				<xsl:when test="/post/userlist/user[@userid=$userid]/@lockuser=2">�û��Ѿ�������</xsl:when>
				<xsl:otherwise>
					<xsl:choose>
						<xsl:when test="@locktopic=3">���ݴ����</xsl:when>
						<xsl:otherwise></xsl:otherwise>
					</xsl:choose>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<xsl:template name="managemenu">
<div class="mainbar0" style="height:36px;line-height:36px;">
	<div style="height:36px;float:right;">
		<b>����ѡ��:</b>
		<a href="IndivGroup_PostManage.asp?action=�޸�&amp;groupid={post/postinfo/@groupid}&amp;GroupBoardID={post/postinfo/@boardid}&amp;ID={post/postinfo/@topicid}" title="�޸����ӵĻظ���������ʱ�䡣" style="margin-right: 10px;font-weight : normal; ">�޸�</a>
		<xsl:choose>
			<xsl:when test="post/postinfo/@locktopic = 1">| <a href="IndivGroup_PostManage.asp?action=UnchainLock&amp;groupid={/post/postinfo/@groupid}&amp;GroupBoardID={/post/postinfo/@boardid}&amp;ID={/post/postinfo/@topicid}" title="��������⿪����">����</a></xsl:when>
			<xsl:otherwise>| <a href="IndivGroup_PostManage.asp?action=Lock&amp;groupid={/post/postinfo/@groupid}&amp;GroupBoardID={/post/postinfo/@boardid}&amp;ID={/post/postinfo/@topicid}" title="����������">����</a></xsl:otherwise>
		</xsl:choose>
		| <a href="IndivGroup_PostManage.asp?action=DeleteTopic&amp;groupid={/post/postinfo/@groupid}&amp;GroupBoardID={/post/postinfo/@boardid}&amp;ID={/post/postinfo/@topicid}" title="ע�⣺��������ɾ���������������ӣ����ָܻ�">ɾ������</a>
		<xsl:choose>
			<xsl:when test="post/postinfo/@istop = 0">
			| <a href="IndivGroup_PostManage.asp?action=SetTop&amp;groupid={/post/postinfo/@groupid}&amp;GroupBoardID={/post/postinfo/@boardid}&amp;ID={/post/postinfo/@topicid}" title="�����������ù̶�">���ù̶�</a>
			</xsl:when>
			<xsl:otherwise>
			| <a href="IndivGroup_PostManage.asp?action=UnchainTop&amp;groupid={/post/postinfo/@groupid}&amp;GroupBoardID={/post/postinfo/@boardid}&amp;ID={/post/postinfo/@topicid}" title="�����������ù̶�">����̶�</a>
			</xsl:otherwise>
		</xsl:choose>
	</div>
</div>
</xsl:template>

<xsl:template name="managemenume">
	<div class="mainbar0" style="height:36px;line-height:36px;">
		<div style="height:36px;float:right;">
			<b>�������ӹ���:</b>
			<a href="IndivGroup_PostManage.asp?action=DeleteTopic&amp;groupid={/post/postinfo/@groupid}&amp;GroupBoardID={/post/postinfo/@boardid}&amp;ID={/post/postinfo/@topicid}" title="ע�⣺��������ɾ���������������ӣ����ָܻ�">ɾ������</a>
		</div>
	</div>
</xsl:template>
</xsl:stylesheet>