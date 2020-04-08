<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/chan_const.asp"-->
<!--#include file="inc/md5.asp"-->
<%
Dvbbs.Stats="高级用户修改密码"	
if Dvbbs.UserID=0 then Response.redirect "showerr.asp?ErrCodes=<li>您还没有登录，请登录后进行操作。&action=OtherErr"

If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(9)=1) Then Response.redirect "showerr.asp?ErrCodes=<li>本论坛没有开启阳光会员注册、修改资料和密码的功能。&action=OtherErr"

Dvbbs.Loadtemplates("")
Dvbbs.Nav()

Select case request("action")
	case "submod"
		Dvbbs.stats="提交资料"
		Dvbbs.Head_var 0,0,"阳光会员修改信息","challenge_mod_pw.asp"
		call reg_2()
	Case "redir"
		Redir()
	Case else
		Dvbbs.stats="修改资料"
		Dvbbs.Head_var 0,0,"阳光会员修改信息","challenge_mod_pw.asp"
		call reg_1()
End select

Dvbbs.Footer()

sub reg_1()
dim rs
set rs=dvbbs.execute("select * from [dv_user] where userid="&dvbbs.userid)
if rs("IsChallenge")=0 or isnull(rs("IsChallenge")) then Response.redirect "showerr.asp?ErrCodes=<li>您还不是阳光会员，请先升级为阳光会员。&action=OtherErr"

%>
<FORM name=theForm action="challenge_mod.asp?action=submod" method=post>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TBODY>
<TR align=middle>
<Th colSpan=2 height=24>高级用户修改手机号码资料</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>用户名</B>：</TD>
<TD width=60%  class=tablebody1>
<%=Rs("Username")%></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>原手机号码</B>：</TD>
<TD class=tablebody1>
<%=rs("usermobile")%>
</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>新手机号码</B>：</TD>
<TD class=tablebody1>
<input type=text name="newmobile" size=15> <font color=red><B>*</B></font>
</TD>
</TR>
<tr><td align=center class=tablebody2 colspan=2><input type=submit value="提交修改"></td></form></tr>
</tbody>
</table>
</form>
<%

set rs=nothing
end sub

sub reg_2()
if request("newmobile")="" then Response.redirect "showerr.asp?ErrCodes=<li>请输入您的新手机号码。&action=OtherErr"
If Len(Trim(request("newmobile")))<>11 Then Response.redirect "showerr.asp?ErrCodes=<li>请输入正确的新手机号码格式。&action=OtherErr"
dim rs
set rs=Dvbbs.Execute("select * from [dv_user] where userid="&dvbbs.userid)
if rs("IsChallenge")=0 or isnull(rs("IsChallenge")) then Response.redirect "showerr.asp?ErrCodes=<li>您还不是手机高级会员，请先升级为手机高级会员。&action=OtherErr"
dim mobile
mobile=rs("UserMobile")
if Trim(request("newmobile"))=mobile then Response.redirect "showerr.asp?ErrCodes=<li>您输入的新手机号码和原号码相同。&action=OtherErr"

Set Rs=Dvbbs.Execute("Select UserMobile From Dv_User Where UserMobile = '"&Dvbbs.CheckStr(Request("newmobile"))&"'")
If Not (Rs.Eof And Rs.Bof) Then Response.redirect "showerr.asp?ErrCodes=<li>您要修改的手机号别人已经在本论坛使用，请确认您在本论坛是否有其它帐号使用该手机号。<li>您可以使用该 <a href=login.asp>手机号快速登录论坛</a>。。&action=iOtherErr"

Get_ChallengeWord

set rs=dvbbs.execute("select top 1 * from Dv_ChallengeInfo")
Dim MyForumID_Up
MyForumID_Up=rs("D_ForumID")

set rs=nothing
%>
正在提交数据，请稍后……
<form name="redir" action="<%=Dvbbs_Server_Url%>ray.asp?Action=ChanMob" method="post">
<INPUT type=hidden name="mobile" value="<%=mobile%>">
<INPUT type=hidden name="newmobile" value="<%=Trim(request("newmobile"))%>">
<INPUT type=hidden name="forumid" value="<%=MyForumID_Up%>">
<INPUT type=hidden name="backurl" value="<%=Dvbbs.Get_ScriptNameUrl%>challenge_mod.asp?action=redir">
<INPUT type=hidden name="seqno" value="<%=Session("challengeWord")%>">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
end sub

sub redir()
dim rs
dim ErrorCode,ErrorMsg
dim remobile,newremobile,rechallengeWord,retokerWord
dim challengeWord_key,rechallengeWord_key

ErrorCode=trim(request("ErrorCode"))
ErrorMsg=trim(request("ErrorMsg"))
remobile=trim(Dvbbs.CheckStr(request("mobile")))
newremobile=trim(Dvbbs.CheckStr(request("newmobile")))
rechallengeWord=trim(Dvbbs.CheckStr(request("seqno")))
retokerWord=trim(request("token"))

If remobile="" Or newremobile="" Or rechallengeWord="" Or retokerWord="" Then Response.redirect "showerr.asp?ErrCodes=<li>非法的参数。&action=OtherErr"

If ErrorCode = "1" Then
	challengeWord_key=session("challengeWord_key")
	If challengeWord_key=retokerWord Then
		Dvbbs.Execute("update [dv_user] set UserMobile='"&newremobile&"',IsChallenge=1 where UserMobile='"&remobile&"'")
	Else
		Response.redirect "showerr.asp?ErrCodes=<li>非法的提交过程,"&retokerWord&","&challengeWord_key&","&newremobile&","&rechallengeWord&"。&action=OtherErr"
	End If
Else
	Response.redirect "showerr.asp?ErrCodes=<li>"&ErrorMsg&"&action=OtherErr"
	Exit sub
End If
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>注册成功：<%=Dvbbs.Forum_Info(0)%>欢迎您的到来</th>
</tr>
<tr><td class=tablebody1><br>
<ul><li>您在本站成功的修改了高级用户手机号码<br><li><a href="index.asp">进入讨论区</a></li></ul>
</td></tr>
</table>
<%
'Session(Dvbbs.CacheName & "UserID")=Empty
end sub

function checkreal(v)
dim w
if not isnull(v) then
	w=replace(v,"|||","§§§")
	checkreal=w
end if
end function
%>