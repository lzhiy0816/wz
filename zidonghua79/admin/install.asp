<!--#include file="../Conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="../inc/chan_const.asp"-->
<!--#include file="../inc/md5.asp"-->
<%
Head()
If Not Dvbbs.Master Or Session("flag")="" Then
	Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=../admin_login.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	Dvbbs_Error1()
Else
	Select Case request("action")
		Case "apply"
			If request("isnew")="0" Then
				reg_2()
			ElseIf request("isnew")="1" Then
				reg_2a()
			Else
				reg_2b()
			End If
		Case "redir1"
			Call redir1()
		Case "redir2"
			call redir2()
		Case "redir3"
			call redir3()
		Case "redir4"
			call redir4()
		Case "redir5"
			call redir5()
		Case "redir6"
			call redir6()
		Case "redir7"
			call redir7()
		Case "redir8"
			call redir8()
		Case "redir9"
			call redir9()
		Case "redir10"
			call redir10()
		Case Else
			reg_1
	End Select
	If Errmsg<>"" Then Dvbbs_Error1()
End If
Footer()
'ҳ�������ʾ��Ϣ
Sub dvbbs_error1()
	Response.Write"<br>"
	Response.Write"<table cellpadding=3 cellspacing=1 align=center class=""tableBorder"" style=""width:75%"">"
	Response.Write"<tr align=center>"
	Response.Write"<th width=""100%"" height=25 colspan=2>������Ϣ"
	Response.Write"</td>"
	Response.Write"</tr>"
	Response.Write"<tr>"
	Response.Write"<td width=""100%"" class=""Forumrow"" colspan=2>"
	Response.Write ErrMsg
	Response.Write"</td></tr>"
	Response.Write"<tr>"
	Response.Write"<td class=""Forumrow"" valign=middle colspan=2 align=center><a href=""javascript:history.go(-2)""><<������һҳ</a></td></tr>"
	Response.Write"</table>"
	footer()
	Response.End 
End Sub 

Function reg_1()
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function submitclick()
{
	if(dvform.isnew[0].checked==false&&dvform.isnew[1].checked==false&&dvform.isnew[2].checked==false)
	{
		alert('ע��ע�����\n1������ϸ�Ķ�ע��Э�飬������ѡ����վ��ע�ᡱ����ע�ᣬ����ȷ�����ϡ���\n2�����������վ����û��ע���������ע��������ѡ����վ��ע�ᡱ\n3��������Ѿ�ע���������̳վ����ֻ�Ǹı��˵�ַ�������°�װ��������ע��������ѡ����ע��վ��������վ����ע��վ���������ϡ� ')
	}
	else
	{
		dvform.submit()
	}
}
//-->
</SCRIPT>
<table cellpadding=3 cellspacing=1 align=center class=tableBorder><form name="dvform"  action="install.asp?action=apply" method=post>
<tr><th align=center colSpan=2>�������������</td></tr>
<tr><td class=Forumrow align=left colSpan=2>
<iframe src="http://bbs.ray5198.com/bbsapp/protocol.jsp" height=400 width="100%" MARGINWIDTH=0 MARGINHEIGHT=0 HSPACE=0 VSPACE=0 FRAMEBORDER=0></iframe>
</td></tr>
<TR align=middle>
<Th colSpan=2 height=24>��ѡ��ע������</TD>
</TR>
<tr>
<td  width=100% align=center  class=Forumrow  height=24> <input type="radio" name="isnew" value="0"><b>��վ��ע��</b>&nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="isnew" value="1" ><b>��ע��վ��������վ</b>&nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="isnew" value="2" ><b>��ע��վ����������</b></td>
</TR>
<TR align=middle>
<td align=center class=forumRowHighlight colSpan=2><input type="button" value="��ͬ��" Onclick="submitclick();"></td></tr>
</form></table>
<%
End Function

Function reg_3()
	Response.Write "<script>dvbbs_install_reg_3();</script>"
End Function

Function reg_2()
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0

  window.open(theURL,winName,features);

}
//-->
</SCRIPT>
<FORM name=theForm action=install.asp?action=redir1 method=post>
<table cellpadding=3 cellspacing=1 align=center class=tableBorder>
<TR align=middle>
<Th colSpan=2 height=24>��վ��������д</TD>
</TR>
<TR>
<TR>
<Td colSpan=2 class=Forumrow>˵�����ڴ˰�װ�����У���Ĭ����ͬ�������̳���Ż��������������Ͻ��ύ��������������ȷ�ϣ�����дǰ��ȷ������д����������Ƿ���ȷ���⽫Ӱ�쵽������ڷ����о�����������֡�</TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>�û���</B>��<br>������20λ����ĸ�����ֵ����</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="username">&nbsp;<B>*</B> <a style="CURSOR:hand" onclick="MM_openBrWindow('install.asp?action=redir10','getpass','width=300,height=100')">����û�</a></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>�û�����</B>��<br>8-20λ����ĸ�����ֵ����</TD>
<TD width=60%  class=Forumrow>
<INPUT type=password size=30 name="password">&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>�ظ�����</B>��</TD>
<TD width=60%  class=Forumrow>
<INPUT type=password size=30 name="repassword">&nbsp;<B>*</B></TD>
</TR>
<%GetUserInput()%>
</tbody>
</table>
</td></tr></tbody></table>
<table cellpadding=0 cellspacing=0 border=0 width='98%' align=center>
<tr>
<td width=50% height=24> </td>
<td width=50% ><input type=submit value="ע ��" name=Submit>����<input type=reset value="�� ��" name=Submit2></td>
</tr></table>
</form>
<%
End Function

Function reg_2a()
Set Rs=Dvbbs.Execute("Select Forum_ChanName From Dv_Setup")
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0

  window.open(theURL,winName,features);

}
//-->
</SCRIPT>
<FORM name=theForm action="install.asp?action=redir2" method=post>
<table cellpadding=3 cellspacing=1 align=center class=tableBorder>
<TR align=middle>
<Th colSpan=2 height=24>��ע��վ��������վ</TD>
</TR>
<TR>
<TR>
<Td colSpan=2 class=Forumrow>˵����������Ѿ�ע������Ż���������ֱ���ڴ���д������֤��Ϣ�ύ����������������֤��</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>�û���</B>��<br>������20λ����ĸ�����ֵ����</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="username" value="<%=Rs(0)%>"> ��ע���ʱ��ʹ�õ��û��� <a style="CURSOR:hand" onclick="MM_openBrWindow('install.asp?action=redir10','getpass','width=300,height=100')">����û�</a></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>�û�����</B>��<br>8-20λ����ĸ�����ֵ����</TD>
<TD width=60%  class=Forumrow>
<INPUT type=password size=34 name="password"> <a style="CURSOR:hand" onclick="MM_openBrWindow('install.asp?action=redir9','getpass','width=300,height=100')">��������</a></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>��̳����</B>��</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="forumname" value="<%=Server.HtmlEncode(Dvbbs.Forum_Info(0))%>"></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>��̳��ַ</B>��</TD>
<TD width=60%  class=Forumrow>
<INPUT type=hidden value="<%=Replace(Dvbbs.Get_ScriptNameUrl,Lcase(Dvbbs.CacheData(33,0)),"")%>" name="forumUrl">
<%=Replace(Dvbbs.Get_ScriptNameUrl,Lcase(Dvbbs.CacheData(33,0)),"")%></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>��̳�ṩ��</B>��</TD>
<TD width=60%  class=Forumrow>
�����ȷ�</TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>��̳�汾</B>��</TD>
<TD width=60%  class=Forumrow>
Dvbbs 7.1.0
</TD>
</TR>
</tbody>
</table>
</td></tr></tbody></table>
<table cellpadding=0 cellspacing=0 border=0 width='+Dvbbs.Forum_Body[12]+' align=center>
<tr>
<td width=50% height=24> </td>
<td width=50% ><input type=submit value="�����������" name=Submit>����<input type=reset value="�� ��" name=Submit2></td>
</tr></table>
</form>
<%
Rs.Close
Set Rs=Nothing
End Function

Function reg_2b()
Set Rs=Dvbbs.Execute("Select Forum_ChanName From Dv_Setup")
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0

  window.open(theURL,winName,features);

}
//-->
</SCRIPT>
<FORM name=theForm action="install.asp?action=redir5" method=post>
<table cellpadding=3 cellspacing=1 align=center class=tableBorder>
<TR align=middle>
<Th colSpan=2 height=24>��ע��վ����������</TD>
</TR>
<TR>
<TR>
<Td colSpan=2 class=Forumrow>˵����������Ѿ�ע������Ż���������ֱ���ڴ���д������֤��Ϣ�ύ����������������֤��</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>�û���</B>��<br>������20λ����ĸ�����ֵ����</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="username" value="<%=Rs(0)%>"> <a style="CURSOR:hand" onclick="MM_openBrWindow('install.asp?action=redir10','getpass','width=300,height=100')">����û�</a></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>����</B>��</TD>
<TD width=60%  class=Forumrow>
<INPUT type=password size=34 name="password"> <a style="CURSOR:hand" onclick="MM_openBrWindow('install.asp?action=redir9','getpass','width=300,height=100')">��������</a></TD>
</TR>
</tbody>
</table>
</td></tr></tbody></table>
<table cellpadding=0 cellspacing=0 border=0 width='+Dvbbs.Forum_Body[12]+' align=center>
<tr>
<td width=50% height=24> </td>
<td width=50% ><input type=submit value="�����������" name=Submit>����<input type=reset value="�� ��" name=Submit2></td>
</tr></table>
</form>
<%
Rs.Close
Set Rs=Nothing
End Function

Function redir5()
	If request("username")="" Then
		Errmsg=ErrMsg + "<BR><li>�����������û�����"
		Exit Function
	End If
	If request("password")="" Then
		Errmsg=ErrMsg + "<BR><li>�����������û����롣"
		Exit Function
	End If
	Get_ChallengeWord

	session("challengeWord_key")=md5(Session("challengeWord") & ":raynetwork",32)
%>
�����ύ���ݣ����Ժ󡭡�
<form name="redir" action="http://bbs.ray5198.com/bbsapp/stationModify.jsp" method="post">
<INPUT type=hidden name="username" value="<%=checkreal(request("username"))%>">
<INPUT type=hidden name="password" value="<%=checkreal(request("password"))%>">
<input type=hidden value="<%=Session("challengeWord")%>" name="challengeWord">
<input type=hidden value="<%=Replace(Dvbbs.Get_ScriptNameUrl,Lcase(Dvbbs.CacheData(33,0)),"")%>" name="forumUrl">
<input type=hidden value="<%=Dvbbs.CacheData(33,0)%>install.asp?action=redir6" name="dirPage">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
End Function

'��ע��վ������ע����վ���ؽ��ҳ��
Function redir6()
	Dim ErrorCode,ErrorMsg
	Dim reForumID,reNewPkey,rechallengeWord,retokerWord
	Dim challengeWord_key,rechallengeWord_key
	Dim Forum_Master_Reg_Temp_2,OldForumID,i
	ErrorCode=trim(request("ErrorCode"))
	ErrorMsg=trim(request("ErrorMsg"))
	rechallengeWord=trim(Dvbbs.CheckStr(request("challengeWord")))
	retokerWord=trim(request("tokenWord"))
	'rechallengeWord_key=md5_32(rechallengeWord & ":" & Dvbbs.CacheData(21,0))
	Dim vipboardsetting,vipboardslist,vipisupdate
	vipisupdate=false

	Select Case ErrorCode
	Case 100
		challengeWord_key=session("challengeWord_key")
		'If challengeWord_key=retokerWord Then
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0

  window.open(theURL,winName,features);

}
//-->
</SCRIPT>
<FORM name=theForm action="install.asp?action=redir7" method=post>
<table cellpadding=3 cellspacing=1 align=center class=tableBorder>
<TR align=middle>
<Th colSpan=2 height=24>����վ������</TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>�û���</B>��<br>������20λ����ĸ�����ֵ����</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="username" value="<%=request("username")%>">&nbsp;<B>*</B> <a style="CURSOR:hand" onclick="MM_openBrWindow('install.asp?action=redir10','getpass','width=300,height=100')">����û�</a></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>�û�����</B>��<br>8-20λ����ĸ�����ֵ����</TD>
<TD width=60%  class=Forumrow>
<INPUT type=password size=30 name="password" value="<%=request("password")%>">&nbsp;<B>*</B> <a style="CURSOR:hand" onclick="MM_openBrWindow('install.asp?action=redir9','getpass','width=300,height=100')">��������</a></TD>
</TR>
<%GetUserInput()%>
</tbody>
</table>
</td></tr></tbody></table>
<table cellpadding=0 cellspacing=0 border=0 width='98%' align=center>
<tr>
<td width=50% height=24> </td>
<td width=50% ><input type=submit value="����վ������" name=Submit>����<input type=reset value="�� ��" name=Submit2></td>
</tr></table>
</form>
<%
		'Else
		'	Errmsg=ErrMsg + "<BR><li>��֤ʧ�ܣ��Ƿ��Ĳ�����"
		'	Exit Function
		'End If
	Case 101
		Errmsg=ErrMsg + "<BR><li>���ڶ��Ż���������������ע��ʧ�ܣ�"&ErrorMsg&"��"
		Exit Function
	Case 102
		Errmsg=ErrMsg + "<BR><li>��ע����û����Ͷ��Ż����������������ϵ��û����ظ���"&ErrorMsg&"��"
		Exit Function
	Case 201
		Errmsg=ErrMsg + "<BR><li>����д����Ϣ�ڶ��Ż����������������ϵ�¼��֤ʧ�ܣ�"&ErrorMsg&"��"
		Exit Function
	Case Else
		Errmsg=ErrMsg + "<BR><li>�Ƿ����ύ���̣�"&ErrorMsg&"��"
		Exit Function
	End Select
	Emp_ChallengeWord
	session("Forum_Master_Reg_Temp_2")=""
End Function

Function redir7()
	If request("username")="" Then
		Errmsg=ErrMsg + "<BR><li>�����������û�����"
		Exit Function
	End If
	If request("password")="" Then
		Errmsg=ErrMsg + "<BR><li>�����������û����롣"
		Exit Function
	End If
	If request("realname")="" Then
		Errmsg=ErrMsg + "<BR><li>������������ʵ������"
		Exit Function
	End If
	If request("identityNo")="" Then
		Errmsg=ErrMsg + "<BR><li>�������������֤�š�"
		Exit Function
	End If
	If request("sex")="" Then
		Errmsg=ErrMsg + "<BR><li>��ѡ�������Ա�"
		Exit Function
	End If
	If request("postcode")="" Then
		Errmsg=ErrMsg + "<BR><li>�����������������롣"
		Exit Function
	End If
	If request("address")="" Then
		Errmsg=ErrMsg + "<BR><li>���������ĵ�ַ��"
		Exit Function
	End If
	If request("email")="" Then 
		Errmsg=ErrMsg + "<BR><li>�����������ʼ���ַ��"
		Exit Function
	End If
	If request("forumname")="" Then
		Errmsg=ErrMsg + "<BR><li>������������̳���ơ�"
		Exit Function
	End If
	If request("bankname")="" Then
		Errmsg=ErrMsg + "<BR><li>��ѡ�����Ŀ������С�"
		Exit Function
	End If
	If request("bankid")="" Then
		Errmsg=ErrMsg + "<BR><li>���������������ʺš�"
		Exit Function
	End If
	If request("bankaddress")="" Then
		Errmsg=ErrMsg + "<BR><li>�������������е�ַ��"
		Exit Function
	End If
	If request("telephone")="" Then
		Errmsg=ErrMsg + "<BR><li>������������ϵ�绰��"
		Exit Function
	End If
	If request("receiver")="" Then
		Errmsg=ErrMsg + "<BR><li>�����������տ��ˡ�"
		Exit Function
	End If
	session("Forum_Master_Reg_Temp_1")=checkreal(request("username")) & "|||" & checkreal(request("email")) & "|||" & checkreal(request("forumname")) & "|||" & checkreal(request("forumurl")) & "|||�����ȷ�|||7.1.0"

	Get_ChallengeWord
	session("challengeWord_key")=md5(Session("challengeWord") & ":raynetwork",32)

%>
�����ύ���ݣ����Ժ󡭡�
<form name="redir" action="http://bbs.ray5198.com/bbsapp/stationMdfOK.jsp" method="post">
<INPUT type=hidden name="username" value="<%=checkreal(request("username"))%>">
<INPUT type=hidden name="password" value="<%=checkreal(request("password"))%>">
<INPUT type=hidden name="realname" value="<%=checkreal(request("realname"))%>">
<INPUT type=hidden name="identityNo" value="<%=checkreal(request("identityNo"))%>">
<INPUT type=hidden name="sex" value="<%=checkreal(request("sex"))%>">
<INPUT type=hidden name="postcode" value="<%=checkreal(request("postcode"))%>">
<INPUT type=hidden name="address" value="<%=checkreal(request("address"))%>">
<INPUT type=hidden name="receiver" value="<%=checkreal(request("realname"))%>">
<INPUT type=hidden name="email" value="<%=checkreal(request("email"))%>">
<INPUT type=hidden name="backupemail" value="<%=checkreal(request("backupemail"))%>">
<INPUT type=hidden name="forumName" value="<%=checkreal(request("forumname"))%>">
<INPUT type=hidden name="forumUrl" value="<%=Replace(Dvbbs.Get_ScriptNameUrl,Lcase(Dvbbs.CacheData(33,0)),"")%>">
<INPUT type=hidden name="telephone" value="<%=checkreal(request("telephone"))%>">
<INPUT type=hidden name="mobile" value="<%=checkreal(request("mobile"))%>">
<INPUT type=hidden name="bankname" value="<%=checkreal(request("bankname"))%>">
<INPUT type=hidden name="bankid" value="<%=checkreal(request("bankid"))%>">
<INPUT type=hidden name="bankaddress" value="<%=checkreal(request("bankaddress"))%>">
<INPUT type=hidden name="forumProvider" value="�����ȷ�">
<INPUT type=hidden name="version" value="Dvbbs 7.1.0">
<INPUT type=hidden name="web_intro" value="<%=checkreal(request("web_intro"))%>">
<input type=hidden value="<%=Session("challengeWord")%>" name="challengeWord">
<input type=hidden value="<%=Dvbbs.CacheData(33,0)%>install.asp?action=redir8" name="dirPage">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
End Function

Function redir8()

	Dim ErrorCode,ErrorMsg
	Dim reForumID,reNewPkey,rechallengeWord,retokerWord
	Dim challengeWord_key,rechallengeWord_key
	Dim Forum_Master_Reg_Temp_2,OldForumID,i
	ErrorCode=trim(request("ErrorCode"))
	ErrorMsg=trim(request("ErrorMsg"))
	rechallengeWord=trim(Dvbbs.CheckStr(request("challengeWord")))
	retokerWord=trim(request("tokenWord"))
	'rechallengeWord_key=md5_32(rechallengeWord & ":" & Dvbbs.CacheData(21,0))
	Dim vipboardsetting,vipboardslist,vipisupdate
	vipisupdate=false

	Select Case ErrorCode
	Case 100
		challengeWord_key=session("challengeWord_key")
		If challengeWord_key=retokerWord Then
			Forum_Master_Reg_Temp_2=Split(Dvbbs.CheckStr(session("Forum_Master_Reg_Temp_2")),"|||")
		Else
			Errmsg=ErrMsg + "<BR><li>��֤ʧ�ܣ��Ƿ��Ĳ�����"
			Exit Function
		End If
	Case 101
		Errmsg=ErrMsg + "<BR><li>���ڶ��Ż���������������ע��ʧ�ܣ�"&ErrorMsg&"��"
		Exit Function
	Case 102
		Errmsg=ErrMsg + "<BR><li>��ע����û����Ͷ��Ż����������������ϵ��û����ظ���"&ErrorMsg&"��"
		Exit Function
	Case 201
		Errmsg=ErrMsg + "<BR><li>����д����Ϣ�ڶ��Ż����������������ϵ�¼��֤ʧ�ܣ�"&ErrorMsg&"��"
		Exit Function
	Case Else
		Errmsg=ErrMsg + "<BR><li>�Ƿ����ύ���̣�"&ErrorMsg&"��"
		Exit Function
	End Select
	Emp_ChallengeWord
	session("Forum_Master_Reg_Temp_2")=""
	Dim iForum_ChanSetting,mForum_ChanSetting
	Set Rs=Dvbbs.Execute("Select Forum_ChanSetting,Forum_challengePassWord From Dv_Setup")
	iForum_ChanSetting = Rs(0)
	iForum_ChanSetting = Split(iForum_ChanSetting,",")
	For i = 0 To Ubound(iForum_ChanSetting)
		If i = 0 Then
			mForum_ChanSetting = 1
		Else
			Select Case i
			Case 1
				mForum_ChanSetting = mForum_ChanSetting & ",1"
			Case 13
				mForum_ChanSetting = mForum_ChanSetting & ",1"
			Case Else
				mForum_ChanSetting = mForum_ChanSetting & "," & iForum_ChanSetting(i)
			End Select
		End If
	Next
	Dvbbs.Execute("Update Dv_Setup Set Forum_ChanSetting='"&Dvbbs.CheckStr(mForum_ChanSetting)&"'")
	Dvbbs.loadSetup()
%>
<table cellpadding=3 cellspacing=1 align=center class=tableBorder>
<tr>
<th height=24>�ɹ���Ϣ��������̳���Ż���������Ϣ���³ɹ�</th>
</tr>
<tr><td class=Forumrow><br>
<ul><li><a href="challenge.asp">����������̳</a></li></ul>
</td></tr>
</table>
<%
End Function

Function redir1()
	If request("username")="" Then
		Errmsg=ErrMsg + "<BR><li>�����������û�����"
		Exit Function
	End If
	If request("password")="" Then
		Errmsg=ErrMsg + "<BR><li>�����������û����롣"
		Exit Function
	End If
	If request("repassword")="" Then
		Errmsg=ErrMsg + "<BR><li>����������ȷ�����롣"
		Exit Function
	End If
	If Lcase(request("password"))<>Lcase(request("repassword")) Then
		Errmsg=ErrMsg + "<BR><li>������������ȷ�����벻һ�¡�"
		Exit Function
	End If
	If request("realname")="" Then
		Errmsg=ErrMsg + "<BR><li>������������ʵ������"
		Exit Function
	End If
	If request("identityNo")="" Then
		Errmsg=ErrMsg + "<BR><li>�������������֤�š�"
		Exit Function
	End If
	If request("sex")="" Then
		Errmsg=ErrMsg + "<BR><li>��ѡ�������Ա�"
		Exit Function
	End If
	If request("postcode")="" Then
		Errmsg=ErrMsg + "<BR><li>�����������������롣"
		Exit Function
	End If
	If request("address")="" Then
		Errmsg=ErrMsg + "<BR><li>���������ĵ�ַ��"
		Exit Function
	End If
	If request("email")="" Then 
		Errmsg=ErrMsg + "<BR><li>�����������ʼ���ַ��"
		Exit Function
	End If
	If request("forumname")="" Then
		Errmsg=ErrMsg + "<BR><li>������������̳���ơ�"
		Exit Function
	End If
	If request("bankname")="" Then
		Errmsg=ErrMsg + "<BR><li>��ѡ�����Ŀ������С�"
		Exit Function
	End If
	If request("bankid")="" Then
		Errmsg=ErrMsg + "<BR><li>���������������ʺš�"
		Exit Function
	End If
	If request("bankaddress")="" Then
		Errmsg=ErrMsg + "<BR><li>�������������е�ַ��"
		Exit Function
	End If
	If request("telephone")="" Then
		Errmsg=ErrMsg + "<BR><li>������������ϵ�绰��"
		Exit Function
	End If
	If request("receiver")="" Then
		Errmsg=ErrMsg + "<BR><li>�����������տ��ˡ�"
		Exit Function
	End If
	session("Forum_Master_Reg_Temp_1")=checkreal(request("username")) & "|||" & checkreal(request("email")) & "|||" & checkreal(request("forumname")) & "|||" & checkreal(request("forumurl")) & "|||�����ȷ�|||7.1.0"

	Get_ChallengeWord
	Session("challengeWord_key")=md5(Session("challengeWord") & ":raynetwork",32)

%>
�����ύ���ݣ����Ժ󡭡�
<form name="redir" action="http://bbs.ray5198.com/bbsapp/stationReg.jsp" method="post">
<INPUT type=hidden name="username" value="<%=checkreal(request("username"))%>">
<INPUT type=hidden name="password" value="<%=checkreal(request("password"))%>">
<INPUT type=hidden name="realname" value="<%=checkreal(request("realname"))%>">
<INPUT type=hidden name="identityNo" value="<%=checkreal(request("identityNo"))%>">
<INPUT type=hidden name="sex" value="<%=checkreal(request("sex"))%>">
<INPUT type=hidden name="postcode" value="<%=checkreal(request("postcode"))%>">
<INPUT type=hidden name="address" value="<%=checkreal(request("address"))%>">
<INPUT type=hidden name="receiver" value="<%=checkreal(request("realname"))%>">
<INPUT type=hidden name="email" value="<%=checkreal(request("email"))%>">
<INPUT type=hidden name="backupemail" value="<%=checkreal(request("backupemail"))%>">
<INPUT type=hidden name="forumName" value="<%=checkreal(request("forumname"))%>">
<INPUT type=hidden name="forumUrl" value="<%=Replace(Dvbbs.Get_ScriptNameUrl,Lcase(Dvbbs.CacheData(33,0)),"")%>">
<INPUT type=hidden name="telephone" value="<%=checkreal(request("telephone"))%>">
<INPUT type=hidden name="mobile" value="<%=checkreal(request("mobile"))%>">
<INPUT type=hidden name="bankname" value="<%=checkreal(request("bankname"))%>">
<INPUT type=hidden name="bankid" value="<%=checkreal(request("bankid"))%>">
<INPUT type=hidden name="bankaddress" value="<%=checkreal(request("bankaddress"))%>">
<INPUT type=hidden name="forumProvider" value="�����ȷ�">
<INPUT type=hidden name="version" value="Dvbbs 7.1.0">
<INPUT type=hidden name="web_intro" value="<%=checkreal(request("web_intro"))%>">
<input type=hidden value="<%=Session("challengeWord")%>" name="challengeWord">
<input type=hidden value="<%=Dvbbs.CacheData(33,0)%>install.asp?action=redir3" name="dirPage">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
End Function

Function redir2()

	If request("username")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�����������û�����"
		Exit Function
	End If
	If request("password")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�������������롣"
		Exit Function
	End If
	If request("forumname")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>������������̳���ơ�"
		Exit Function
	End If
	session("Forum_Master_Reg_Temp_2")=checkreal(request("username")) & "|||" & checkreal(request("password")) & "|||" & checkreal(request("forumname")) & "|||" & checkreal(request("forumurl")) & "|||�����ȷ�|||7.1.0"

	Get_ChallengeWord

	session("challengeWord_key")=md5(Session("challengeWord") & ":raynetwork",32)
%>
�����ύ���ݣ����Ժ󡭡�
<form name="redir" action="http://bbs.ray5198.com/bbsapp/oldStnReg.jsp" method="post">
<INPUT type=hidden name="username" value="<%=checkreal(request("username"))%>">
<INPUT type=hidden name="password" value="<%=checkreal(request("password"))%>">
<INPUT type=hidden name="forumName" value="<%=checkreal(request("forumname"))%>">
<INPUT type=hidden name="forumUrl" value="<%=Replace(Dvbbs.Get_ScriptNameUrl,Lcase(Dvbbs.CacheData(33,0)),"")%>">
<INPUT type=hidden name="forumProvider" value="�����ȷ�">
<INPUT type=hidden name="version" value="Dvbbs 7.1.0">
<input type=hidden value="<%=Session("challengeWord")%>" name="challengeWord">
<input type=hidden value="<%=Dvbbs.CacheData(33,0)%>install.asp?action=redir4" name="dirPage">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
End Function

'��һ��ע�Ჿ�ֽ��շ�����Ϣ
Function redir3()
	Dim ErrorCode,ErrorMsg
	Dim reForumID,reNewPkey,rechallengeWord,retokerWord
	Dim challengeWord_key,rechallengeWord_key
	Dim Forum_Master_Reg_Temp_1,OldForumID
	Dim vipboardsetting,vipboardslist,vipisupdate,i
	vipisupdate=false
	ErrorCode=trim(request("ErrorCode"))
	ErrorMsg=trim(request("ErrorMsg"))
	reForumID=trim(Dvbbs.CheckStr(request("ForumID")))
	reNewPkey=trim(Dvbbs.CheckStr(request("NewPkey")))
	rechallengeWord=trim(Dvbbs.CheckStr(request("challengeWord")))
	retokerWord=trim(request("tokenWord"))
	'rechallengeWord_key=md5_32(rechallengeWord & ":" & Dvbbs.CacheData(21,0))

	Select Case ErrorCode
	Case 100
		challengeWord_key=Session("challengeWord_key")							
		If challengeWord_key=retokerWord Then
			Forum_Master_Reg_Temp_1=split(Dvbbs.CheckStr(session("Forum_Master_Reg_Temp_1")),"|||")
			Dim iForum_ChanSetting,mForum_ChanSetting
			Set Rs=Dvbbs.Execute("Select Forum_ChanSetting,Forum_challengePassWord From Dv_Setup")
			iForum_ChanSetting = Rs(0)
			iForum_ChanSetting = Split(iForum_ChanSetting,",")
			For i = 0 To Ubound(iForum_ChanSetting)
				If i = 0 Then
					mForum_ChanSetting = 1
				Else
					Select Case i
					Case 1
						mForum_ChanSetting = mForum_ChanSetting & ",1"
					Case 13
						mForum_ChanSetting = mForum_ChanSetting & ",1"
					Case Else
						mForum_ChanSetting = mForum_ChanSetting & "," & iForum_ChanSetting(i)
					End Select
				End If
			Next

			Dvbbs.Execute("update dv_Setup set Forum_challengePassWord='"&reNewPkey&"',Forum_ChanSetting='"&Dvbbs.CheckStr(mForum_ChanSetting)&"',Forum_ChanName='"&Forum_Master_Reg_Temp_1(0)&"',Forum_IsInstall=1,Forum_Version='"&Forum_Master_Reg_Temp_1(5)&"'")
			Set Rs=Dvbbs.Execute("select top 1 * from Dv_ChallengeInfo")
			OldForumID=rs("D_ForumID")
			Rs.close
			Set Rs=Nothing
			'session("Forum_Master_Reg_Temp_1")=checkreal(request("username")) & "|||" & checkreal(request("email")) & "|||" & checkreal(request("forumname")) & "|||" & checkreal(request("forumurl")) & "|||�����ȷ�|||7.1.0"
			Dvbbs.Execute("update Dv_ChallengeInfo set D_ForumID='"&reForumID&"',D_UserName='"&Forum_Master_Reg_Temp_1(0)&"',D_Password='"&reNewPkey&"',D_email='"&Forum_Master_Reg_Temp_1(1)&"',D_forumname='"&Forum_Master_Reg_Temp_1(2)&"',D_forumurl='"&Forum_Master_Reg_Temp_1(3)&"',D_forumProvider='"&Forum_Master_Reg_Temp_1(4)&"',D_version='"&Forum_Master_Reg_Temp_1(5)&"',D_challengePassWord='"&OldForumID&"'")
			Dvbbs.loadSetup()
		Else
			Response.Write retokerWord
			Response.Write "<BR>"
			Response.Write challengeWord_key
			Response.end
			'Errmsg=ErrMsg + "<BR><li>��֤ʧ�ܣ��Ƿ��Ĳ�����"
			'Exit Function
		End If
	Case 101
		Errmsg=ErrMsg + "<BR><li>ע��ʧ�ܣ�"&ErrorMsg&"��"
		Exit Function
	Case 102
		Errmsg=ErrMsg + "<BR><li>��ע����û����Ͷ��Ż����������������ϵ��û����ظ���"&ErrorMsg&"��"
		Exit Function
	Case 201
		Errmsg=ErrMsg + "<BR><li>����д����Ϣ�ڶ��Ż����������������ϵ�¼��֤ʧ�ܣ�"&ErrorMsg&"��"
		Exit Function
	Case Else
		Errmsg=ErrMsg + "<BR><li>�Ƿ����ύ���̣�"&ErrorMsg&"��"
		Exit Function
	End Select
	Emp_ChallengeWord
	Session("Forum_Master_Reg_Temp_1")=""
%>
<table cellpadding=3 cellspacing=1 align=center class=tableBorder>
<tr>
<th height=24>ע��ɹ���������̳���Ż�������ע��ɹ�</th>
</tr>
<tr><td class=Forumrow><br>
<ul><li>���ɹ��Ŀ�������̳�ֻ����ŷ����ܡ�</li></ul>
</td></tr>
</table>
<%
End Function

'��ע��վ������ע����վ���ؽ��ҳ��
Function redir4()

	Dim ErrorCode,ErrorMsg
	Dim reForumID,reNewPkey,rechallengeWord,retokerWord
	Dim challengeWord_key,rechallengeWord_key
	Dim Forum_Master_Reg_Temp_2,OldForumID,i
	ErrorCode=trim(request("ErrorCode"))
	ErrorMsg=trim(request("ErrorMsg"))
	reForumID=trim(Dvbbs.CheckStr(request("ForumID")))
	reNewPkey=trim(Dvbbs.CheckStr(request("NewPkey")))
	rechallengeWord=trim(Dvbbs.CheckStr(request("challengeWord")))
	retokerWord=trim(request("tokenWord"))
	'rechallengeWord_key=md5_32(rechallengeWord & ":" & Dvbbs.CacheData(21,0))
	Dim vipboardsetting,vipboardslist,vipisupdate
	vipisupdate=false

	Select Case ErrorCode
	Case 100
		challengeWord_key=session("challengeWord_key")
		If challengeWord_key=retokerWord Then
			Forum_Master_Reg_Temp_2=Split(Dvbbs.CheckStr(session("Forum_Master_Reg_Temp_2")),"|||")
			Dim iForum_ChanSetting,mForum_ChanSetting
			Set Rs=Dvbbs.Execute("Select Forum_ChanSetting,Forum_challengePassWord From Dv_Setup")
			iForum_ChanSetting = Rs(0)
			iForum_ChanSetting = Split(iForum_ChanSetting,",")
			For i = 0 To Ubound(iForum_ChanSetting)
				If i = 0 Then
					mForum_ChanSetting = 1
				Else
					Select Case i
					Case 1
						mForum_ChanSetting = mForum_ChanSetting & ",1"
					Case 13
						mForum_ChanSetting = mForum_ChanSetting & ",1"
					Case Else
						mForum_ChanSetting = mForum_ChanSetting & "," & iForum_ChanSetting(i)
					End Select
				End If
			Next
			Dvbbs.Execute("update dv_Setup set Forum_challengePassWord='"&reNewPkey&"',Forum_ChanSetting='"&Dvbbs.CheckStr(mForum_ChanSetting)&"',Forum_ChanName='"&Forum_Master_Reg_Temp_2(0)&"',Forum_IsInstall=1,Forum_Version='"&Forum_Master_Reg_Temp_2(5)&"'")
			Set Rs=Dvbbs.Execute("select top 1 * from Dv_ChallengeInfo")
			OldForumID=rs("D_ForumID")
			Rs.Close
			Set Rs=Nothing
			Dvbbs.Execute("update Dv_ChallengeInfo set D_ForumID='"&reForumID&"',D_UserName='"&Forum_Master_Reg_Temp_2(0)&"',D_Password='"&reNewPkey&"',D_forumname='"&Forum_Master_Reg_Temp_2(2)&"',D_forumurl='"&Forum_Master_Reg_Temp_2(3)&"',D_forumProvider='"&Forum_Master_Reg_Temp_2(4)&"',D_version='"&Forum_Master_Reg_Temp_2(5)&"',D_challengePassWord='"&OldForumID&"'")

			Dvbbs.loadSetup()
		Else
			Errmsg=ErrMsg + "<BR><li>��֤ʧ�ܣ��Ƿ��Ĳ�����"
			Exit Function
		End If
	Case 101
		Errmsg=ErrMsg + "<BR><li>���ڶ��Ż���������������ע��ʧ�ܣ�"&ErrorMsg&"��"
		Exit Function
	Case 102
		Errmsg=ErrMsg + "<BR><li>��ע����û����Ͷ��Ż����������������ϵ��û����ظ���"&ErrorMsg&"��"
		Exit Function
	Case 201
		Errmsg=ErrMsg + "<BR><li>����д����Ϣ�ڶ��Ż����������������ϵ�¼��֤ʧ�ܣ�"&ErrorMsg&"��"
		Exit Function
	Case Else
		Errmsg=ErrMsg + "<BR><li>�Ƿ����ύ���̣�"&ErrorMsg&"��"
		Exit Function
	End Select
	Emp_ChallengeWord
	session("Forum_Master_Reg_Temp_2")=""
%>
<table cellpadding=3 cellspacing=1 align=center class=tableBorder>
<tr>
<th height=24>�ɹ���Ϣ��������̳���Ż���������Ϣ���³ɹ�</th>
</tr>
<tr><td class=Forumrow><br>
<ul><li><a href="challenge.asp">����������̳</a></li></ul>
</td></tr>
</table>
<%
End Function

function checkreal(v)
Dim w
if not isnull(v) then
	w=replace(v,"|||","����")
	checkreal=w
end if
end function

Sub GetUserInput()
%>
<TR>
<TD width=40% class=Forumrow><B>��ʵ����</B>��</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="realname" value="<%=Request("realname")%>">&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>���֤��</B>��<BR>֧��15λ��18λ����</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="identityNo" value="<%=Request("identityNo")%>">&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40%  class=Forumrow><B>�Ա�</B>��<BR>��ѡ�������Ա�</font></TD>
<TD width=60%  class=Forumrow> <INPUT type=radio CHECKED value="0" name=sex <%if request("sex")="0" then Response.Write "checked"%>>
�� &nbsp;&nbsp;
<INPUT type=radio value="1" name=sex <%if request("sex")="1" then Response.Write "checked"%>>
Ů&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>�ʱ�</B>��</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="postcode" value="<%=Request("postcode")%>">&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>��ַ</B>��</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="address" value="<%=Request("address")%>">&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>�ʼ���ַ</B>��<br>Ϊ���ܳɹ���װ��̳���������д��ʵ��Ч��Email��ַ</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="email" value="<%=Request("email")%>">&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>�����ʼ�</B>��</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="backupemail" value="<%=Request("backupemail")%>"></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>��̳����</B>��</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="forumname" value="<%=Request("forumname")%>">&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>��̳��ַ</B>��</TD>
<TD width=60%  class=Forumrow>
<INPUT type=hidden value="<%=Replace(Dvbbs.Get_ScriptNameUrl,Lcase(Dvbbs.CacheData(33,0)),"")%>" name="forumUrl">
<%=Replace(Dvbbs.Get_ScriptNameUrl,Lcase(Dvbbs.CacheData(33,0)),"")%></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>�绰</B>��<BR>Ϊ�˷������Ǽ�ʱ������ϵ������ʵ��д</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="telephone" value="<%=Request("telephone")%>">&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>�ֻ�</B>��</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="mobile" value="<%=Request("mobile")%>"></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>��������</B>��</TD>
<TD width=60%  class=Forumrow>
<select name="bankname">
<option value="��������" <%if request("bankname")="��������" then Response.Write "selected"%>>��������</option>
<option value="��������" <%if request("bankname")="��������" then Response.Write "selected"%>>��������</option>
<option value="�й�����" <%if request("bankname")="�й�����" then Response.Write "selected"%>>�й�����</option>
<option value="ũҵ����" <%if request("bankname")="ũҵ����" then Response.Write "selected"%>>ũҵ����</option>
<option value="��������" <%if request("bankname")="��������" then Response.Write "selected"%>>��������</option>
<option value="��ͨ����" <%if request("bankname")="��ͨ����" then Response.Write "selected"%>>��ͨ����</option>
</select>
&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>���е�ַ</B>��<BR>XXʡXX��/�ء���������ʡ�人��</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="bankaddress" value="<%=Request("bankaddress")%>">&nbsp;<B>*</B></TD>
</TR>
<TD width=40% class=Forumrow><B>�����ʺ�</B>��<BR>Ϊȷ�������յ��������ϸ��д���м䲻Ҫ�ӿո�</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="bankid" value="<%=Request("bankid")%>">&nbsp;<B>*</B></TD>
</TR>
<TD width=40% class=Forumrow><B>�տ�������</B>(������)��<BR>Ϊȷ�������յ��������ϸ��д</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="receiver" value="<%=Request("receiver")%>">&nbsp;<B>*</B></TD>
</TR>
<TR align=middle>
<Th colSpan=2 height=24>��̳������д</TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>��̳�ṩ��</B>��</TD>
<TD width=60%  class=Forumrow>
�����ȷ�</TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>��̳�汾</B>��</TD>
<TD width=60%  class=Forumrow>
Dvbbs 7.1.0</TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>��վ���ݽ���</B>��</TD>
<TD width=60%  class=Forumrow>
<textarea class=smallarea cols=50 name=web_intro rows=7 wrap=VIRTUAL><%=Request("web_intro")%></textarea>
</TD>
</TR>
<%
End Sub

'ȡ������
Sub Redir9()
	If Request("t")="t" Then
		If Request("ErrorCode")="-1" Then
			Response.Write "����������Ϣ�ѳɹ�������ע������䣬���������"
		Else
			Response.Write "����ʧ��"
		End If
	Else
%>
<table cellpadding=3 cellspacing=1 align=center class=tableBorder>
<FORM METHOD=POST ACTION="http://bbs.ray5198.com/bbsapp/sendpassword.jsp">
<tr>
<th height=24>��������</th>
</tr>
<tr><td class=Forumrow><br>
�����������û�����<input type=text size=30 name="username"> <input type=submit name=submit value="�ύ">
<input type=hidden value="<%=Replace(Dvbbs.Get_ScriptNameUrl,Lcase(Dvbbs.CacheData(33,0)),"")%><%=Dvbbs.CacheData(33,0)%>install.asp?action=redir9&t=t" name="dirPage">
</td></tr>
</FORM>
</table>
<%
	End If
End Sub
'����û�
Sub Redir10()
	If Request("t")="t" Then
		If Request("ErrorCode")="-1" Then
			Response.Write "���ʺ���δע��"
		Else
			Response.Write "���ʺ��ѱ�ע��"
		End If
	Else
%>
<table cellpadding=3 cellspacing=1 align=center class=tableBorder>
<FORM METHOD=POST ACTION="http://bbs.ray5198.com/bbsapp/userexist.jsp">
<tr>
<th height=24>����û��Ƿ����</th>
</tr>
<tr><td class=Forumrow><br>
��������Ҫע����û�����<input type=text size=30 name="username"> <input type=submit name=submit value="�ύ">
<input type=hidden value="<%=Replace(Dvbbs.Get_ScriptNameUrl,Lcase(Dvbbs.CacheData(33,0)),"")%><%=Dvbbs.CacheData(33,0)%>install.asp?action=redir10&t=t" name="dirPage">
</td></tr>
</FORM>
</table>
<%
	End If
End Sub
%>