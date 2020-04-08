<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/chan_const.asp"-->
<!--#include file="inc/md5.asp"-->
<%
Dvbbs.Stats="用户升级"
Dvbbs.Loadtemplates("")
Dvbbs.Nav()
If Dvbbs.UserID=0 then Response.redirect "showerr.asp?ErrCodes=<li>您还没有登录，请登录后进行操作。&action=OtherErr"

If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(9)=1) Then Response.redirect "showerr.asp?ErrCodes=<li>本论坛没有开启阳光会员注册、修改资料和密码的功能。&action=OtherErr"

Select case request("action")
	case "submobile"
		dvbbs.stats="提交资料"
		Dvbbs.Head_var 0,0,"普通用户升级","challenge_up.asp"
		call reg_2()
	Case "redir"
		dvbbs.stats="提交资料"
		Dvbbs.Head_var 0,0,"普通用户升级","challenge_up.asp"
		call redir()
	Case else
		dvbbs.stats="输入资料"
		Dvbbs.Head_var 0,0,"普通用户升级","challenge_up.asp"
		call reg_1()
End Select

Dvbbs.Footer()

sub reg_1()
dim rs
set rs=dvbbs.execute("select IsChallenge from [dv_user] where userid="&Dvbbs.userid)
if rs(0)=1 Then Response.redirect "showerr.asp?iErrCodes=<li>您已经是高级用户，如果需要<a href=challenge_mod.asp>修改资料请点击下面连接</a>。&action=OtherErr"
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<form action="challenge_up.asp?action=submobile" method=post>
<tr><th align=center colspan=2>普通用户升级为高级用户</td></tr>
<tr><td class=tablebody1 align=right width="40%"><B>请输入您的论坛密码</B>：</td>
	<td class=tablebody1 width="60%">
<input type=password size=30 name="password">
	</td></tr>
<tr><td class=tablebody1 align=right width="40%"><B>请输入您的手机号码</B>：</td>
	<td class=tablebody1 width="60%">
<input type=text size=30 name="mobile">
	</td></tr>
<tr><td align=center class=tablebody2 colspan=2><input type=submit value="提 交"></td></tr>
</form>
</table>
<%
end sub

sub reg_2()
dim rs
if request("mobile")="" then Response.redirect "showerr.asp?ErrCodes=<li>请输入您的手机号。&action=OtherErr"

if request("password")="" then Response.redirect "showerr.asp?ErrCodes=<li>请输入您的论坛密码。&action=OtherErr"

Get_ChallengeWord

set rs=dvbbs.execute("select top 1 * from Dv_ChallengeInfo")
Dim MyForumID_Up
MyForumID_Up=rs("D_ForumID")

Set  Rs=dvbbs.execute("select * from [dv_user] where Usermobile='"&Dvbbs.CheckStr(request("mobile"))&"' and IsChallenge=1")
If Not (rs.bof and rs.bof) Then Response.redirect "showerr.asp?ErrCodes=<li>您使用的手机号别人已经在本论坛使用，请确认您在本论坛是否有其它帐号使用该手机号。<li>您可以使用该 <a href=login.asp>手机号快速登录论坛</a>。&action=iOtherErr"

Dim UserIM
set rs=dvbbs.execute("select * from [dv_user] where userid="&Dvbbs.userid)
if md5(trim(request("password")),16) <> rs("userpassword") then Response.redirect "showerr.asp?ErrCodes=<li>您输入的论坛密码不正确，请重新输入。&action=OtherErr"
dim sex
if cint(rs("usersex"))=1 then
	sex="F"
else
	sex="M"
end if
UserIM = Split(Rs("UserIM"),"|||")
%>
正在提交数据，请稍后……
<form name="redir_up" action="<%=Dvbbs_Server_Url%>ray.asp?Action=UserReg" method="post">
<INPUT type=hidden name="password" value="<%=Request.form("password")%>">
<INPUT type=hidden name="mobile" value="<%=request("mobile")%>">
<INPUT type=hidden name="forumid" value="<%=MyForumID_Up%>">
<input type=hidden value="<%=Session("challengeWord")%>" name="seqno">
<input type=hidden value="<%=Dvbbs.Get_ScriptNameUrl%>challenge_up.asp?action=redir" name="backurl">
</form>
<script LANGUAGE=javascript>
<!--
redir_up.submit();
//-->
</script>
<%
set rs=nothing
end sub

sub redir()
dim rs
dim ErrorCode,ErrorMsg
dim remobile,rechallengeWord,retokerWord
dim challengeWord_key,rechallengeWord_key

ErrorCode=trim(request("ErrorCode"))
ErrorMsg=trim(request("ErrorMsg"))
remobile=trim(Dvbbs.CheckStr(request("mobile")))
rechallengeWord=trim(Dvbbs.CheckStr(request("seqno")))
retokerWord=trim(request("token"))

If ErrorCode = "1" Then
	challengeWord_key=session("challengeWord_key")
	If challengeWord_key=retokerWord Then
		Dvbbs.Execute("update [dv_user] set UserMobile='"&remobile&"',IsChallenge=1 where userid="&Dvbbs.UserID)
	Else
		Response.redirect "showerr.asp?ErrCodes=<li>非法的提交过程。&action=OtherErr"
		'Response.Write challengeWord_key
		'Response.Write "<BR>"
		'Response.Write retokerWord
		'Response.End
	End If
Else
	Response.redirect "showerr.asp?ErrCodes=<li>"&ErrorMsg&"&action=OtherErr"
	Exit Sub
End If
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>注册成功：<%=Dvbbs.Forum_Info(0)%>欢迎您的到来</th>
</tr>
<tr><td class=tablebody1><br>
<ul><li>您在本站成功的注册成为高级用户<br><li><a href="index.asp">进入讨论区</a></li></ul>
</td></tr>
</table>
<%
'Session(Dvbbs.CacheName & "UserID")=Empty
end sub
%>