<!--#include file="../Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
Dvbbs.Stats="��̳����������"
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
<title><%=Dvbbs.Forum_info(0)%>--�������</title>
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
          <td align="left" style="BORDER-BOTTOM: #ffffff 0px solid" id=AdminTableTitleLink><b><font color=#FFffff>��ӭ <%=Dvbbs.MemberName%> ������̳�������&nbsp;&nbsp;&nbsp;&nbsp;<a href="../index.asp" target=_top>��̳��ҳ</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="http://bbs.dvbbs.net/index.asp?boardid=143" target=_blank title="��ӭ���վ������̳�����������ｻ����̽����վ�Ķ�λ�������ռ����⡢��վ���輼���������Լ��ƹ�">��վ���轻��</a></font></b>&nbsp;&nbsp;&nbsp;&nbsp;
	<a href="http://bbs.dvbbs.net/index.asp?boardid=111" target=_blank title="��̳��ؼ���������ѯ�������ģ�彻�������齻����"><font color=#FFFFAA><B>��̳��������</B></font></a>
		  </td>
          <td align="Left" valign=top style="BORDER-BOTTOM: #ffffff 0px solid" id=AdminTableTitleLink width="70"><b><font color=#FFffff><a href="http:/www.dvbbs.net" target=_blank title="����������ٷ���Ϣ">�ٷ�����</a>��</b></td>
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
<title><%=dvbbs.Forum_info(0)%>--����ҳ��</title>
<style type=text/css>
body  { background:#799AE1; margin:0px; font:normal 12px ����; 
SCROLLBAR-FACE-COLOR: #799AE1; SCROLLBAR-HIGHLIGHT-COLOR: #799AE1; 
SCROLLBAR-SHADOW-COLOR: #799AE1; SCROLLBAR-DARKSHADOW-COLOR: #799AE1; 
SCROLLBAR-3DLIGHT-COLOR: #799AE1; SCROLLBAR-ARROW-COLOR: #FFFFFF;
SCROLLBAR-TRACK-COLOR: #AABFEC;
}
table  { border:0px; }
td  { font:normal 12px ����; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px ����; color:#000000; text-decoration:none; }
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
REM ������Ŀ����
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
		menu(i,k)="<a href=help.asp?action=view&id="&trs(0)&" target=main><img src=images/bullet.gif border=0 alt=����鿴����Ŀ�İ���></a>" & trs("h_title")
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
	  <span><a href="index.asp" target=_top><b>������ҳ</b></a> | <a href="logout.asp" target=_top><b>�˳�</b></a></span>
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
<tr><td height=20><img alt src="images/bullet.gif" border="0" width="15" height="20"><a href="../bokeadmin.asp" target=main>��̳����ϵͳ����</a></td></tr>
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
	  <span>������̳��Ϣ</span>
	</td>
  </tr>
  <tr>
    <td>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=135>
<tr><td height=20>&nbsp;<br><a href="http://www.aspsky.net/" target=_blank>��Ȩ���У�<BR>�������ȷ�(<font face=Verdana, Arial, Helvetica, sans-serif>AspSky<font color=#CC0000>.Net</font></font>)</a><BR>
<a href="http://www.dvbbs.net/" target=_blank>֧����̳��<BR>��������̳(<font face=Verdana, Arial, Helvetica, sans-serif>Dvbbs<font color=#CC0000>.Net</font></font>)</a><BR><BR>
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
	Response.Redirect "../showerr.asp?action=OtherErr&ErrCodes=<li>��ҳ��Ϊ����Աר�ã���<a href=admin_login.asp target=_top>��¼</a>����롣"
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
<tr><th class="tableHeaderText" colspan=2 height=25>��̳��Ϣͳ��</th><tr>
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
ϵͳ��Ϣ����̳������ <B><%=rs("Forum_PostNum")%></B> ������ <B><%=rs("Forum_topicnum")%></B> �û��� <B><%=rs("Forum_usernum")%></B> ������û��� <B><%=isaudituser%></B> �������� <B><%=BoardListNum%></B>
<%
end if
rs.close
set rs=nothing
%>
</td></tr>
<tr><td  class="forumRowHighlight" height=23 colspan=2>
����̳�ɶ����ȷ棨aspsky.net����Ȩ�� <%=dvbbs.Forum_info(0)%> ʹ�ã���ǰʹ�ð汾Ϊ ������̳
<%
If IsSqlDatabase=1 Then
	Response.Write "SQL���ݿ�"
Else
	Response.Write "Access���ݿ�"
End If
Response.Write " Dvbbs " & Dvbbs.Forum_Version
%>
</td></tr>
<tr>
<td width="50%"  class="forumRow" height=23>���������ͣ�<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
<td width="50%" class="forumRow">�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
</tr>
<tr>
<td width="40%" class="forumRow" height=23>FSO�ı���д��<%If Not IsObjInstalled(theInstalledObjects(9)) Then%><font color="<%=dvbbs.mainsetting(1)%>"><b>��</b></font><%else%><b>��</b><%end if%></td>
<td width="60%" class="forumRow">
<%If IsObjInstalled(theInstalledObjects(18)) Then%>Jmail4.3�������֧�֣�<%else%>Jmail4.2���֧�֣�<%end if%>
<%If IsObjInstalled(theInstalledObjects(18)) or IsObjInstalled(theInstalledObjects(13)) Then%>
<b>��</b>
<%else%>
<font color="<%=dvbbs.mainsetting(1)%>"><b>��</b></font>
<%end if%>
&nbsp;&nbsp;<a href="data.asp?action=SpaceSize">>>�鿴����ϸ��������Ϣ���</a>
</td>
</tr>
<tr><td class="forumRow" height=23 colspan=2>
<a href="challenge.asp"><font color=red>����֧�����ֻ������������</font></a>����Ҫ������̳��ȯ��ֵ����ص�����̳����ķ���֧�����ʺ�Ϊ<%=Dvbbs.Forum_ChanSetting(4)%>
</td></tr>
<tr><td class="forumRow" height=23 colspan=2>
<a href="data.asp"><font color=red>���ݶ��ڱ���</font></a>����ע�����ö������ݱ��ݣ����ݵĶ��ڱ��ݿ�����޶ȵı�������̳���ݵİ�ȫ
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
<tr><th class="tableHeaderText" colspan=2 height=25 ID=TableTitleLink><a href="http://bbs.dvbbs.net/union_post.asp" target=SiteAdmin>��̳վ��֮��</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=143" target=SiteAdmin>��̳������</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=134" target=SiteAdmin>��̳������</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=8" target=SiteAdmin>��̳������</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=13" target=SiteAdmin>��̳�����</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=102" target=SiteAdmin>��̳�����</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=-1&UserName=<%=Forum_Pack(1)%>&PassWord=<%=Forum_Pack(2)%>" target=SiteAdmin>�ҵ�����</a> | <a href="http://bbs.dvbbs.net/union_post.asp?iBoardID=-2&UserName=<%=Forum_Pack(1)%>&PassWord=<%=Forum_Pack(2)%>" target=SiteAdmin>��������</a></th><tr>
<tr><td class="forumRow" height=23 valign=top>ע�⣺��ȫʹ�ñ����ܱ�����ȷ���� <a href="Setting.asp">�����ٷ��Զ�ͨѶ����</a></tr>
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
<tr><th class="tableHeaderText" colspan=2 height=25>��̳����С��ʿ</th><tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>�û���Ȩ��</B>
</td><td class="forumRow" height=23 width="*">������̳��ע���û��ֳɲ�ͬ���û��飬ÿ���û������ӵ�в�ͬ����̳����Ȩ�ޣ������ڶ�����̳7.0�汾֮���û��ȼ���ϵ����û����У������û��ȼ�û���Զ���Ȩ�ޣ���ô����ȼ���Ȩ�޾�ʹ�����������û���Ȩ�ޣ���֮��ӵ������ȼ��Լ���Ȩ�ޡ�<font color=red>ÿ���ȼ������û������趨��Ȩ�޶��������������̳��</font></td></tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>�ְ���Ȩ��</B>
</td><td class="forumRow" height=23 width="*">
ÿ���û�������Զ���Ȩ�����õĵȼ�������������������̳�и�������ӵ�в�ͬ��Ȩ�ޣ�����˵����������ע���û�����������·�ڰ���A���ܷ�����������ȵ�Ȩ�����ã��������������̳Ȩ�޵����ã�<font color=blue>����������˵���Էֳ��ܶ����ͬ�������͵���̳</font>
</td></tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>�û�Ȩ���趨</B>
</td><td class="forumRow" height=23 width="*">
ÿ���û�����������������̳�и�������ӵ�в�ͬ��Ȩ�޻��������Ȩ�ޣ�����˵�����������û�A�ڰ���A��ӵ�����й���Ȩ�ޡ�<U>������������Ȩ����Ҫע�����������˳��Ϊ���û�Ȩ������(<font color=gray>�Զ���</font>)<font color=blue> <B>></B> </font>�ְ���Ȩ���趨(<font color=gray>�Զ���</font>)<font color=blue> <B>></B> </font>�û���Ȩ���趨(<font color=gray>Ĭ��</font>)</U>
</td></tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>�Է��ģ��Ĺ���</B>
</td><td class="forumRow" height=23 width="*">
���а�������̳����ģ��Ĺ���ģ������̳�Ļ���CSS���ã���̳�����ĸ��ģ���̳��ҳ����ĸ��ģ�ͼƬ�����ã����԰������ã��½�ģ��ҳ���ģ�壬ģ�����½���ͬ�����ԡ�ͼƬ������ģ��Ԫ�صȵȹ��ܣ�����ӵ��ģ��ĵ��뵼�����ܣ�������������ʵ������̳�������߱༭���л�
</td></tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>һ�仰��ʿ</B>
</td><td class="forumRow" height=23 width="*">
�� ���ڲ�ͬ����ģ���ҳ�棬Ҫ��ϸ��ҳ���е�˵�������������
<BR>
�� �û��鼰����չ��Ȩ�����ã�����̳�ĸ��������м��������ԣ�Ҫ������������Ⱥ���Ч˳��
<BR>
�� �����̳������ʱ�򣬱����˻�ͷ�����ð���߼������Ƿ���ȷ
<BR>
�� �������뵽������̳�ٷ�վ�����ʣ��кܶ����ĵ����ѻ��æ��<a href="help.asp">�鿴������ʿ����</a>
</td></tr>
</table>
<p></p>

<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>��̳�����ݷ�ʽ</th><tr>
<FORM METHOD=POST ACTION="user.asp?action=userSearch&userSearch=9&usernamechk=yes"><tr>
<td width="20%"  class="forumRow" height=23>���ٲ����û�</td>
<td width="80%" class="forumRow">
<input type="text" name="username" size="30"> <input type="submit" value="���̲���">
<input type="hidden" name="userclass" value="0">
<input type="hidden" name="searchMax" value=100>
</td></FORM>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>��ݹ�������</td>
<td width="80%" class="forumRow"><a href=board.asp?action=add>�����̳���</a> | <a href=board.asp>������̳����</a> | <a href="ReloadForumCache.asp">���·���������</a></td>
</tr>
<tr><form action="update.asp?action=updat" method=post>
<td width="20%" class="forumRow" height=23>���ٸ�������</td>
<td width="80%" class="forumRow">
<Input Type="hidden" Name="index" value="1">
<input type="submit" name="Submit" value="���·ְ�������">&nbsp;
<input type="submit" name="Submit" value="������̳������">&nbsp;
<input type="submit" name="Submit" value="��������û�">
</td></form>
</tr>
</table>
<script language='javascript'> function jumpto(url) { if (url != '') { window.open(url); } } </script>
<p></p>

<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>�����ȷ���̳ϵͳ[������̳]</th></tr>
<tr>
<td width="20%" class="forumRow" height=23>��Ʒ����</td>
<td width="80%" class="forumRow">
<a href="http://www.dvbbs.net" target=_blank>���ڶ����ȷ�����Ƽ����޹�˾&nbsp;&nbsp;<font color=blue>�й����Ұ�Ȩ������Ȩ�ǼǺ�2004SR00001</font>
</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>��Ʒ����</td>
<td width="80%" class="forumRow">
��վ��ҵ�� ������̳��Ŀ��&nbsp;&nbsp;<a href="http://www.dvbbs.net/busslist.asp" target=_blank><font color=blue>��ҵ���Ͱ���</font></a>
</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>��ϵ����</td>
<td width="80%" class="forumRow">
��վ��ҵ����0898-68557467<BR>
<a href="http://www.aspsky.cn/" target=_blank>������ҵ��</a>��0898-68592224 68592294<BR>
���������棺0898-68556467<BR>
<a href="http://www.dvbbs.net/contact.asp" target=_blank>����鿴��ϸ��ϵ����</a>
</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>�������</td>
<td width="80%" class="forumRow">
������̳�����֯��Dvbbs Plus Organization��
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