<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/chan_const.asp"-->
<!--#include file="inc/md5.asp"-->
<%
Dvbbs.Stats="�߼��û��޸�����"	
if Dvbbs.UserID=0 then Response.redirect "showerr.asp?ErrCodes=<li>����û�е�¼�����¼����в�����&action=OtherErr"

If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(9)=1) Then Response.redirect "showerr.asp?ErrCodes=<li>����̳û�п��������Աע�ᡢ�޸����Ϻ�����Ĺ��ܡ�&action=OtherErr"

Dvbbs.Loadtemplates("")
Dvbbs.Nav()

Select case request("action")
	case "submod"
		Dvbbs.stats="�ύ����"
		Dvbbs.Head_var 0,0,"�����Ա�޸���Ϣ","challenge_mod_pw.asp"
		call reg_2()
	Case "redir"
		Redir()
	Case else
		Dvbbs.stats="�޸�����"
		Dvbbs.Head_var 0,0,"�����Ա�޸���Ϣ","challenge_mod_pw.asp"
		call reg_1()
End select

Dvbbs.Footer()

sub reg_1()
dim rs
set rs=dvbbs.execute("select * from [dv_user] where userid="&dvbbs.userid)
if rs("IsChallenge")=0 or isnull(rs("IsChallenge")) then Response.redirect "showerr.asp?ErrCodes=<li>�������������Ա����������Ϊ�����Ա��&action=OtherErr"

%>
<FORM name=theForm action="challenge_mod.asp?action=submod" method=post>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TBODY>
<TR align=middle>
<Th colSpan=2 height=24>�߼��û��޸��ֻ���������</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>�û���</B>��</TD>
<TD width=60%  class=tablebody1>
<%=Rs("Username")%></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>ԭ�ֻ�����</B>��</TD>
<TD class=tablebody1>
<%=rs("usermobile")%>
</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>���ֻ�����</B>��</TD>
<TD class=tablebody1>
<input type=text name="newmobile" size=15> <font color=red><B>*</B></font>
</TD>
</TR>
<tr><td align=center class=tablebody2 colspan=2><input type=submit value="�ύ�޸�"></td></form></tr>
</tbody>
</table>
</form>
<%

set rs=nothing
end sub

sub reg_2()
if request("newmobile")="" then Response.redirect "showerr.asp?ErrCodes=<li>�������������ֻ����롣&action=OtherErr"
If Len(Trim(request("newmobile")))<>11 Then Response.redirect "showerr.asp?ErrCodes=<li>��������ȷ�����ֻ������ʽ��&action=OtherErr"
dim rs
set rs=Dvbbs.Execute("select * from [dv_user] where userid="&dvbbs.userid)
if rs("IsChallenge")=0 or isnull(rs("IsChallenge")) then Response.redirect "showerr.asp?ErrCodes=<li>���������ֻ��߼���Ա����������Ϊ�ֻ��߼���Ա��&action=OtherErr"
dim mobile
mobile=rs("UserMobile")
if Trim(request("newmobile"))=mobile then Response.redirect "showerr.asp?ErrCodes=<li>����������ֻ������ԭ������ͬ��&action=OtherErr"

Set Rs=Dvbbs.Execute("Select UserMobile From Dv_User Where UserMobile = '"&Dvbbs.CheckStr(Request("newmobile"))&"'")
If Not (Rs.Eof And Rs.Bof) Then Response.redirect "showerr.asp?ErrCodes=<li>��Ҫ�޸ĵ��ֻ��ű����Ѿ��ڱ���̳ʹ�ã���ȷ�����ڱ���̳�Ƿ��������ʺ�ʹ�ø��ֻ��š�<li>������ʹ�ø� <a href=login.asp>�ֻ��ſ��ٵ�¼��̳</a>����&action=iOtherErr"

Get_ChallengeWord

set rs=dvbbs.execute("select top 1 * from Dv_ChallengeInfo")
Dim MyForumID_Up
MyForumID_Up=rs("D_ForumID")

set rs=nothing
%>
�����ύ���ݣ����Ժ󡭡�
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

If remobile="" Or newremobile="" Or rechallengeWord="" Or retokerWord="" Then Response.redirect "showerr.asp?ErrCodes=<li>�Ƿ��Ĳ�����&action=OtherErr"

If ErrorCode = "1" Then
	challengeWord_key=session("challengeWord_key")
	If challengeWord_key=retokerWord Then
		Dvbbs.Execute("update [dv_user] set UserMobile='"&newremobile&"',IsChallenge=1 where UserMobile='"&remobile&"'")
	Else
		Response.redirect "showerr.asp?ErrCodes=<li>�Ƿ����ύ����,"&retokerWord&","&challengeWord_key&","&newremobile&","&rechallengeWord&"��&action=OtherErr"
	End If
Else
	Response.redirect "showerr.asp?ErrCodes=<li>"&ErrorMsg&"&action=OtherErr"
	Exit sub
End If
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>ע��ɹ���<%=Dvbbs.Forum_Info(0)%>��ӭ���ĵ���</th>
</tr>
<tr><td class=tablebody1><br>
<ul><li>���ڱ�վ�ɹ����޸��˸߼��û��ֻ�����<br><li><a href="index.asp">����������</a></li></ul>
</td></tr>
</table>
<%
'Session(Dvbbs.CacheName & "UserID")=Empty
end sub

function checkreal(v)
dim w
if not isnull(v) then
	w=replace(v,"|||","����")
	checkreal=w
end if
end function
%>