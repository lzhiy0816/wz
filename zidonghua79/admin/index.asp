<!--#include file="../Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
Dvbbs.Stats="论坛管理控制面板"
Select Case Request("action")
Case "admin_left"
	Call admin_left()
Case "admin_main"
	Call admin_main()
Case "admin_top"
	Call admin_top()
Case "xml"
	Call Showxml()
Case Else
	Call Main()
End Select
Sub Showxml()
	Dim XMLDOM
	Set XMLDOM=Application(Dvbbs.CacheName&"_sBoradlist").cloneNode(True)
	Dim node
	Response.Clear
	Response.CharSet="gb2312"  
	Response.ContentType="text/xml"
	Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"&vbNewLine
	Response.Write XMLDom.documentElement.XML
	Set XMLDOM=Nothing
End Sub

Sub Main()
If Not Dvbbs.Master Or Session("flag")="" Then
	Response.Redirect "../admin_login.asp"
Else
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN"
     "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312" />
<title><%=Dvbbs.Forum_info(0)%>--控制面板</title>
<meta name="author" content="Dynamic Vanguard" />
<meta name="author's email" content="eway@aspsky.net" />
<meta name="author's homepage" content="www.aspsky.net" />
</head>
<frameset rows="33,*">
<frame longdesc="" src="index.asp?action=admin_top" noresize="noresize" frameborder="0" name="ads" scrolling="no" marginwidth="0" marginheight="0" />
<frameset rows="675" cols="180,*">
<frame longdesc="" src="index.asp?action=admin_left" name="list" noresize="noresize" marginwidth="0" marginheight="0" frameborder="0" scrolling="yes" />
<frame longdesc="" src="index.asp?action=admin_main" name="main" marginwidth="0" marginheight="0" frameborder="0" scrolling="yes" />
</frameset>
<noframes><body></body></noframes>
</frameset>
</html>
<%
End If
End Sub
Sub admin_top()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="refresh" content="60">
<title>top</title>
</head>
<%=Replace(template.html(1),"{$path}","images/")%>
<style>
#AdminTableTitleLink A {
	COLOR: #ffffff; TEXT-DECORATION: none
}
#AdminTableTitleLink A:hover {
	COLOR: #ffffff; TEXT-DECORATION: underline
}
</style>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" height="32" border="0" cellpadding="0" cellspacing="0" background="images/top_bg.gif">
  <tr> 
    <td width="147" bgcolor="#5298E4" style="BORDER-BOTTOM: #ffffff 0px solid"><a href="http://www.dvbbs.net/" target="_blank"><img src="images/top_logo.gif" width="147" height="32" border="0"></a></td>
    <td style="BORDER-BOTTOM: #ffffff 0px solid"> 
      <table width="100%" border="0" align="left" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="right" width=15 style="BORDER-BOTTOM: #ffffff 0px solid"></td>
          <td align="left" style="BORDER-BOTTOM: #ffffff 0px solid" id=AdminTableTitleLink><b><font color=#FFffff>欢迎 <%=Dvbbs.MemberName%> 进入论坛设置面板&nbsp;&nbsp;&nbsp;&nbsp;<a href="../index.asp" target=_top>论坛首页</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="http://bbs.dvbbs.net/index.asp?boardid=143" target=_blank title="欢迎广大站长和论坛爱好者在这里交流与探讨网站的定位、管理、空间问题、网站建设技术、经验以及推广">网站建设交流</a></font></b>&nbsp;&nbsp;&nbsp;&nbsp;
	<a href="http://bbs.dvbbs.net/index.asp?boardid=111" target=_blank title="论坛相关技术问题咨询、插件和模板交流、经验交流等"><font color=#FFFFAA><B>论坛技术交流</B></font></a>
		  </td>
          <td align="Left" valign=top style="BORDER-BOTTOM: #ffffff 0px solid" id=AdminTableTitleLink width="70"><b><font color=#FFffff><a href="http:/www.dvbbs.net" target=_blank title="点击浏览更多官方信息">官方公告</a>：</b></td>
		  <td align="Left" valign=top style="BORDER-BOTTOM: #ffffff 0px solid" id=AdminTableTitleLink width="200"><b>
<%
	If Session("packjs") =empty Then 
		Dim Yrs,dupack,packjs
		set Yrs=Dvbbs.Execute("select * from Dv_setup")
		If IsNull(Yrs("forum_pack")) Or Yrs("forum_pack")="" Then
			dupack = "1|||admin|||admin888"
			dupack = Split(dupack,"|||")
		Else
			dupack = Yrs("forum_pack")
			dupack = Split(dupack,"|||")
			If Ubound(dupack)<>2 Then
				dupack(0) = 1
				dupack(1) = "admin"
				dupack(2) = "admin888"
			End If
		End If
		packjs ="http://bbs.dvbbs.net/packjs.asp?a="&dupack(1)&"&b="&dupack(2)&"&c="&Dvbbs.Forum_info(0)&"&d="&Dvbbs.Get_ScriptNameUrl&"&e="&Yrs("Forum_usernum")&"&f="&Yrs("Forum_PostNum")&"&g="&Dvbbs.Forum_info(5)&"&h="&IsSqlDatabase&"&i=Dvbbs Version "&Dvbbs.Forum_Version&"&j="&Yrs("Forum_TopicNum")&"&k="&MyBoardOnline.Forum_Online&"&l="&Dvbbs.Forum_ChanSetting(0)
		Yrs.close
		set Yrs=nothing
		Session("packjs")=packjs
	End If
	Response.Write "<script src="""&Session("packjs")&"""></script>"
%>
		  </font></b></td>
          <td align="right" width=15 style="BORDER-BOTTOM: #ffffff 0px solid"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="1" bgcolor="#135093" style="BORDER-BOTTOM: #135093 0px solid"> </td>
  </tr>
</table>
</body>
</html>
<%
end sub

sub admin_left()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<title><%=dvbbs.Forum_info(0)%>--管理页面</title>
<style type=text/css>
body  { background:#799AE1; margin:0px; font:normal 12px 宋体; 
SCROLLBAR-FACE-COLOR: #799AE1; SCROLLBAR-HIGHLIGHT-COLOR: #799AE1; 
SCROLLBAR-SHADOW-COLOR: #799AE1; SCROLLBAR-DARKSHADOW-COLOR: #799AE1; 
SCROLLBAR-3DLIGHT-COLOR: #799AE1; SCROLLBAR-ARROW-COLOR: #FFFFFF;
SCROLLBAR-TRACK-COLOR: #AABFEC;
}
table  { border:0px; }
td  { font:normal 12px 宋体; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px 宋体; color:#000000; text-decoration:none; }
a:hover  { color:#428EFF;text-decoration:underline; }
.sec_menu  { border-left:1px solid white; border-right:1px solid white; border-bottom:1px solid white; overflow:hidden; background:#D6DFF7; }
.menu_title  { }
.menu_title span  { position:relative; top:2px; left:8px; color:#215DC6; font-weight:bold; }
.menu_title2  { }
.menu_title2 span  { position:relative; top:2px; left:8px; color:#428EFF; font-weight:bold; }
</style>
<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
</SCRIPT>
<%
REM 管理栏目设置
dim menu(8,10),trs,k
i=0
k=0
set rs=Dvbbs.Execute("select * from dv_help where h_type=1 and h_parentid=0 and not h_stype=1 order by h_id")
do while not rs.eof
	menu(i,k)=rs("h_title")
	'response.write "menu("&i&","&k&")="""&rs("h_title")&""""
	'response.write chr(10)
	k=k+1
	set trs=Dvbbs.Execute("select * from dv_help where h_type=1 and h_parentid="&rs(0)&" and not h_stype=1 order by h_id")
	do while not trs.eof
		menu(i,k)="<a href=help.asp?action=view&id="&trs(0)&" target=main><img src=images/bullet.gif border=0 alt=点击查看该项目的帮助></a>" & trs("h_title")
		'response.write "menu("&i&","&k&")="""&trs("h_title")&""""
		'response.write chr(10)
		k=k+1
	trs.movenext
	loop
	trs.close
	set trs=nothing
	i=i+1
	k=0
rs.movenext
loop
rs.close
set rs=nothing
'response.end
%>
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table width=100% cellpadding=0 cellspacing=0 border=0 align=left>
    <tr><td valign=top>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=42 valign=bottom>
	  <img src="images/title.gif" width=158 height=38>
    </td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=images/title_bg_quit.gif  >
	  <span><a href="index.asp" target=_top><b>管理首页</b></a> | <a href="logout.asp" target=_top><b>退出</b></a></span>
    </td>
  </tr>
</table>
&nbsp;
<%
	dim j,i
	dim tmpmenu
	dim menuname
	dim menurl
	Dim TempStr,Menu_1,Menu_2
	TempStr = template.html(0)
	Menu_1 = Split(TempStr,"||")
for i=0 to ubound(Menu_1)
%>
<table cellpadding=0 cellspacing=0 width=158 align=center>
<%
	Menu_2 = Split(Menu_1(i),"@@")
	For j = 0 To Ubound(Menu_2)
	If j=0 Then
%>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin_left_<%=i+1%>.gif" id=menuTitle1 onclick="showsubmenu(<%=i%>)">
	  <span><%=Menu_2(0)%></span>
	</td>
  </tr>
  <tr>
    <td style="display" id='submenu<%=i%>'><div class=sec_menu style="width:158"><table cellpadding=0 cellspacing=0 align=center width=150><TBODY>
<%
	Else
	if j=1 then response.write "<tr><td height=5></td></tr>"
%>
<tr><td height=20><img alt src="images/bullet.gif" border="0" width="15" height="20"><%=Menu_2(j)%></td></tr>
<%
	End If
	If i = 0 And j = Ubound(Menu_2) And Dvbbs.Forum_Setting(99)="1" Then
%>
<tr><td height=20><img alt src="images/bullet.gif" border="0" width="15" height="20"><a href="../bokeadmin.asp" target=main>论坛博客系统管理</a></td></tr>
<%
	End If
	next
%><TBODY></table></div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=135>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>
<%next%>
&nbsp;
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin_left_9.gif" id=menuTitle1>
	  <span>动网论坛信息</span>
	</td>
  </tr>
  <tr>
    <td>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=135>
<tr><td height=20>&nbsp;<br><a href="http://www.aspsky.net/" target=_blank>版权所有：<BR>动网先锋(<font face=Verdana, Arial, Helvetica, sans-serif>AspSky<font color=#CC0000>.Net</font></font>)</a><BR>
<a href="http://www.dvbbs.net/" target=_blank>支持论坛：<BR>动网论坛(<font face=Verdana, Arial, Helvetica, sans-serif>Dvbbs<font color=#CC0000>.Net</font></font>)</a><BR><BR>
</td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>
&nbsp;
<%
end sub

sub admin_main()
If Not Dvbbs.Master Or Session("flag")="" Then
	Response.Redirect "../showerr.asp?action=OtherErr&ErrCodes=<li>本页面为管理员专用，请<a href=admin_login.asp target=_top>登录</a>后进入。"
Else
	Dim theInstalledObjects(20)
    theInstalledObjects(0) = "MSWC.AdRotator"
    theInstalledObjects(1) = "MSWC.BrowserType"
    theInstalledObjects(2) = "MSWC.NextLink"
    theInstalledObjects(3) = "MSWC.Tools"
    theInstalledObjects(4) = "MSWC.Status"
    theInstalledObjects(5) = "MSWC.Counters"
    theInstalledObjects(6) = "IISSample.ContentRotator"
    theInstalledObjects(7) = "IISSample.PageCounter"
    theInstalledObjects(8) = "MSWC.PermissionChecker"
    theInstalledObjects(9) = "Scripting.FileSystemObject"
    theInstalledObjects(10) = "adodb.connection"
    
    theInstalledObjects(11) = "SoftArtisans.FileUp"
    theInstalledObjects(12) = "SoftArtisans.FileManager"
    theInstalledObjects(13) = "JMail.SMTPMail"	'Jamil 4.2
    theInstalledObjects(14) = "CDONTS.NewMail"
    theInstalledObjects(15) = "Persits.MailSender"
    theInstalledObjects(16) = "LyfUpload.UploadFile"
    theInstalledObjects(17) = "Persits.Upload.1"
	theInstalledObjects(18) = "JMail.Message"	'Jamil 4.3
	theInstalledObjects(19) = "Persits.Upload"
	theInstalledObjects(20) = "SoftArtisans.FileUp"
	Head()
%>
<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>论坛信息统计</th><tr>
<tr><td class="bodytitle" height=23 colspan=2>
<%
dim isaudituser
set rs=Dvbbs.Execute("select count(*) from [dv_user] where usergroupid=5")
isaudituser=rs(0)
if isnull(isaudituser) then isaudituser=0
Dim BoardListNum
set rs=dvbbs.execute("select count(*) from dv_board")
BoardListNum=rs(0)
If isnull(BoardListNum) then BoardListNum=0
set rs=Dvbbs.Execute("select * from dv_setup")
if not rs.eof then
%>
系统信息：论坛帖子数 <B><%=rs("Forum_PostNum")%></B> 主题数 <B><%=rs("Forum_topicnum")%></B> 用户数 <B><%=rs("Forum_usernum")%></B> 待审核用户数 <B><%=isaudituser%></B> 版面总数 <B><%=BoardListNum%></B>
<%
end if
rs.close
set rs=nothing
%>
</td></tr>
<tr><td  class="forumRowHighlight" height=23 colspan=2>
本论坛由动网先锋（aspsky.net）授权给 <%=dvbbs.Forum_info(0)%> 使用，当前使用版本为 动网论坛
<%
If IsSqlDatabase=1 Then
	Response.Write "SQL数据库"
Else
	Response.Write "Access数据库"
End If
Response.Write " Dvbbs " & Dvbbs.Forum_Version
%>
</td></tr>
<tr>
<td width="50%"  class="forumRow" height=23>服务器类型：<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
<td width="50%" class="forumRow">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
</tr>
<tr>
<td width="40%" class="forumRow" height=23>FSO文本读写：<%If Not IsObjInstalled(theInstalledObjects(9)) Then%><font color="<%=dvbbs.mainsetting(1)%>"><b>×</b></font><%else%><b>√</b><%end if%></td>
<td width="60%" class="forumRow">
<%If IsObjInstalled(theInstalledObjects(18)) Then%>Jmail4.3邮箱组件支持：<%else%>Jmail4.2组件支持：<%end if%>
<%If IsObjInstalled(theInstalledObjects(18)) or IsObjInstalled(theInstalledObjects(13)) Then%>
<b>√</b>
<%else%>
<font color="<%=dvbbs.mainsetting(1)%>"><b>×</b></font>
<%end if%>
&nbsp;&nbsp;<a href="data.asp?action=SpaceSize">>>查看更详细服务器信息检测</a>
</td>
</tr>
<tr><td class="forumRow" height=23 colspan=2>
<a href="challenge.asp"><font color=red>网络支付和手机短信相关设置</font></a>：主要用于论坛点券充值等相关导致论坛收益的服务，支付宝帐号为<%=Dvbbs.Forum_ChanSetting(4)%>
</td></tr>
<tr><td class="forumRow" height=23 colspan=2>
<a href="data.asp"><font color=red>数据定期备份</font></a>：请注意做好定期数据备份，数据的定期备份可最大限度的保障您论坛数据的安全
</td></tr>
</table>
<p></p>

<%
Dim Forum_Pack
Set Rs=Dvbbs.Execute("Select Forum_Pack From Dv_Setup")
Forum_Pack = Rs(0)
Forum_Pack=Split(Forum_Pack,"|||")
If UBound(Forum_Pack)<2 Then ReDim Forum_Pack(3)
If Forum_Pack(0) = "1" Then
%>
<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align=center style="line-height:14pt">
<tr><th class="tableHeaderText" colspan=2 height=25 ID=TableTitleLink><a href="http://bbs.dvbbs.net/union_post.asp" target=SiteAdmin>论坛站长之家</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=143" target=SiteAdmin>论坛管理交流</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=134" target=SiteAdmin>论坛新手区</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=8" target=SiteAdmin>论坛技术区</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=13" target=SiteAdmin>论坛插件区</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=102" target=SiteAdmin>论坛风格区</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=-1&UserName=<%=Forum_Pack(1)%>&PassWord=<%=Forum_Pack(2)%>" target=SiteAdmin>我的问题</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=-2&UserName=<%=Forum_Pack(1)%>&PassWord=<%=Forum_Pack(2)%>" target=SiteAdmin>发表问题</a></th><tr>
<tr><td class="forumRow" height=23 valign=top>注意：完全使用本功能必须正确设置 <a href="Setting.asp">动网官方自动通讯设置</a></tr>
<tr><td class="forumRow" height=23 valign=top>
<iframe src="http://bbs.dvbbs.net/union_post.asp" name="SiteAdmin" height=180 width="100%" MARGINWIDTH=0 MARGINHEIGHT=0 HSPACE=0 VSPACE=0 FRAMEBORDER=0 SCROLLING=no></iframe>
</tr>
</table>
<p></p>
<%
End If
Rs.Close
Set Rs=Nothing
%>

<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align=center style="line-height:14pt">
<tr><th class="tableHeaderText" colspan=2 height=25>论坛管理小贴士</th><tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>用户组权限</B>
</td><td class="forumRow" height=23 width="*">动网论坛将注册用户分成不同的用户组，每个用户组可以拥有不同的论坛操作权限，并且在动网论坛7.0版本之后，用户等级结合到了用户组中，假如用户等级没有自定义权限，那么这个等级的权限就使用他所属的用户组权限，反之则拥有这个等级自己的权限。<font color=red>每个等级或者用户组所设定的权限都是是针对整个论坛的</font></td></tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>分版面权限</B>
</td><td class="forumRow" height=23 width="*">
每个用户组或有自定义权限设置的等级，都可以设置其在论坛中各个版面拥有不同的权限，比如说您可以设置注册用户或者新手上路在版面A不能发贴可以浏览等等权限设置，极大的扩充了论坛权限的设置，<font color=blue>从理论上来说可以分出很多个不同功能类型的论坛</font>
</td></tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>用户权限设定</B>
</td><td class="forumRow" height=23 width="*">
每个用户都可以设置其在论坛中各个版面拥有不同的权限或者特殊的权限，比如说您可以设置用户A在版面A中拥有所有管理权限。<U>对于上述三种权限需要注意的是其优先顺序为：用户权限设置(<font color=gray>自定义</font>)<font color=blue> <B>></B> </font>分版面权限设定(<font color=gray>自定义</font>)<font color=blue> <B>></B> </font>用户组权限设定(<font color=gray>默认</font>)</U>
</td></tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>对风格模板的管理</B>
</td><td class="forumRow" height=23 width="*">
其中包含对论坛所有模板的管理，模板中论坛的基本CSS设置，论坛主风格的更改，论坛分页面风格的更改，图片的设置，语言包的设置，新建模板页面和模板，模板中新建不同的语言、图片、风格等模板元素等等功能，并且拥有模板的导入导出功能，从真正意义上实现了论坛风格的在线编辑和切换
</td></tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>一句话贴士</B>
</td><td class="forumRow" height=23 width="*">
① 对于不同功能模块的页面，要仔细看页面中的说明，以免误操作
<BR>
② 用户组及其扩展的权限设置，对论坛的各种设置有极大扩充性，要充分明白其优先和有效顺序
<BR>
③ 添加论坛大分类的时候，别忘了回头看看该版面高级设置是否正确
<BR>
④ 有问题请到动网论坛官方站点提问，有很多热心的朋友会帮忙，<a href="help.asp">查看更多贴士请点击</a>
</td></tr>
</table>
<p></p>

<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>论坛管理快捷方式</th><tr>
<FORM METHOD=POST ACTION="user.asp?action=userSearch&userSearch=9&usernamechk=yes"><tr>
<td width="20%"  class="forumRow" height=23>快速查找用户</td>
<td width="80%" class="forumRow">
<input type="text" name="username" size="30"> <input type="submit" value="立刻查找">
<input type="hidden" name="userclass" value="0">
<input type="hidden" name="searchMax" value=100>
</td></FORM>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>快捷功能链接</td>
<td width="80%" class="forumRow"><a href=board.asp?action=add>添加论坛类别</a> | <a href=board.asp>管理论坛版面</a> | <a href="ReloadForumCache.asp">更新服务器缓存</a></td>
</tr>
<tr><form action="update.asp?action=updat" method=post>
<td width="20%" class="forumRow" height=23>快速更新数据</td>
<td width="80%" class="forumRow">
<Input Type="hidden" Name="index" value="1">
<input type="submit" name="Submit" value="更新分版面数据">&nbsp;
<input type="submit" name="Submit" value="更新论坛总数据">&nbsp;
<input type="submit" name="Submit" value="清空在线用户">
</td></form>
</tr>
</table>
<script language='javascript'> function jumpto(url) { if (url != '') { window.open(url); } } </script>
<p></p>

<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>动网先锋论坛系统[动网论坛]</th></tr>
<tr>
<td width="20%" class="forumRow" height=23>产品开发</td>
<td width="80%" class="forumRow">
<a href="http://www.dvbbs.net" target=_blank>海口动网先锋网络科技有限公司&nbsp;&nbsp;<font color=blue>中国国家版权局著作权登记号2004SR00001</font>
</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>产品负责</td>
<td width="80%" class="forumRow">
网站事业部 动网论坛项目组&nbsp;&nbsp;<a href="http://www.dvbbs.net/busslist.asp" target=_blank><font color=blue>企业典型案例</font></a>
</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>联系方法</td>
<td width="80%" class="forumRow">
网站事业部：0898-68557467<BR>
<a href="http://www.aspsky.cn/" target=_blank>主机事业部</a>：0898-68592224 68592294<BR>
传　　　真：0898-68556467<BR>
<a href="http://www.dvbbs.net/contact.asp" target=_blank>点击查看详细联系方法</a>
</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>插件开发</td>
<td width="80%" class="forumRow">
动网论坛插件组织（Dvbbs Plus Organization）
</td>
</tr>
</table>

<%
footer
end if
end sub

Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
%>