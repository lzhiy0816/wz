<!--#include file="../Conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="../inc/chan_const.asp"-->
<!--#include file="../inc/md5.asp"-->
<%
Head()
If Not Dvbbs.Master Or Session("flag")="" Then
	Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=../admin_login.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
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
'页面错误提示信息
Sub dvbbs_error1()
	Response.Write"<br>"
	Response.Write"<table cellpadding=3 cellspacing=1 align=center class=""tableBorder"" style=""width:75%"">"
	Response.Write"<tr align=center>"
	Response.Write"<th width=""100%"" height=25 colspan=2>错误信息"
	Response.Write"</td>"
	Response.Write"</tr>"
	Response.Write"<tr>"
	Response.Write"<td width=""100%"" class=""Forumrow"" colspan=2>"
	Response.Write ErrMsg
	Response.Write"</td></tr>"
	Response.Write"<tr>"
	Response.Write"<td class=""Forumrow"" valign=middle colspan=2 align=center><a href=""javascript:history.go(-2)""><<返回上一页</a></td></tr>"
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
		alert('注册注意事项：\n1、请仔细阅读注册协议，您必须选择“新站长注册”或“已注册，重新确认资料”。\n2、如果您的网站从来没有注册过，请在注册类型中选择“新站长注册”\n3、如果您已经注册过阳光论坛站长，只是改变了地址或者重新安装，可以在注册类型中选择“已注册站长绑定新网站或已注册站长更新资料” ')
	}
	else
	{
		dvform.submit()
	}
}
//-->
</SCRIPT>
<table cellpadding=3 cellspacing=1 align=center class=tableBorder><form name="dvform"  action="install.asp?action=apply" method=post>
<tr><th align=center colSpan=2>服务条款和声明</td></tr>
<tr><td class=Forumrow align=left colSpan=2>
<iframe src="http://bbs.ray5198.com/bbsapp/protocol.jsp" height=400 width="100%" MARGINWIDTH=0 MARGINHEIGHT=0 HSPACE=0 VSPACE=0 FRAMEBORDER=0></iframe>
</td></tr>
<TR align=middle>
<Th colSpan=2 height=24>请选择注册类型</TD>
</TR>
<tr>
<td  width=100% align=center  class=Forumrow  height=24> <input type="radio" name="isnew" value="0"><b>新站长注册</b>&nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="isnew" value="1" ><b>已注册站长绑定新网站</b>&nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="isnew" value="2" ><b>已注册站长更新资料</b></td>
</TR>
<TR align=middle>
<td align=center class=forumRowHighlight colSpan=2><input type="button" value="我同意" Onclick="submitclick();"></td></tr>
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
<Th colSpan=2 height=24>新站长资料填写</TD>
</TR>
<TR>
<TR>
<Td colSpan=2 class=Forumrow>说明：在此安装过程中，将默认您同意加入论坛短信互动服务，所有资料将提交至主服务器进行确认，在填写前请确认您填写的相关资料是否正确，这将影响到您今后在服务中具体利益的体现。</TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>用户名</B>：<br>不超过20位的字母或数字的组合</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="username">&nbsp;<B>*</B> <a style="CURSOR:hand" onclick="MM_openBrWindow('install.asp?action=redir10','getpass','width=300,height=100')">检测用户</a></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>用户密码</B>：<br>8-20位的字母或数字的组合</TD>
<TD width=60%  class=Forumrow>
<INPUT type=password size=30 name="password">&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>重复密码</B>：</TD>
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
<td width=50% ><input type=submit value="注 册" name=Submit><input type=reset value="清 除" name=Submit2></td>
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
<Th colSpan=2 height=24>已注册站长绑定新网站</TD>
</TR>
<TR>
<TR>
<Td colSpan=2 class=Forumrow>说明：如果您已经注册过短信互动服务，请直接在此填写您的认证信息提交至主服务器进行验证。</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>用户名</B>：<br>不超过20位的字母或数字的组合</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="username" value="<%=Rs(0)%>"> 您注册的时候使用的用户名 <a style="CURSOR:hand" onclick="MM_openBrWindow('install.asp?action=redir10','getpass','width=300,height=100')">检测用户</a></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>用户密码</B>：<br>8-20位的字母或数字的组合</TD>
<TD width=60%  class=Forumrow>
<INPUT type=password size=34 name="password"> <a style="CURSOR:hand" onclick="MM_openBrWindow('install.asp?action=redir9','getpass','width=300,height=100')">忘记密码</a></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>论坛名称</B>：</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="forumname" value="<%=Server.HtmlEncode(Dvbbs.Forum_Info(0))%>"></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>论坛地址</B>：</TD>
<TD width=60%  class=Forumrow>
<INPUT type=hidden value="<%=Replace(Dvbbs.Get_ScriptNameUrl,Lcase(Dvbbs.CacheData(33,0)),"")%>" name="forumUrl">
<%=Replace(Dvbbs.Get_ScriptNameUrl,Lcase(Dvbbs.CacheData(33,0)),"")%></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>论坛提供者</B>：</TD>
<TD width=60%  class=Forumrow>
动网先锋</TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>论坛版本</B>：</TD>
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
<td width=50% ><input type=submit value="申请更新资料" name=Submit><input type=reset value="清 除" name=Submit2></td>
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
<Th colSpan=2 height=24>已注册站长更新资料</TD>
</TR>
<TR>
<TR>
<Td colSpan=2 class=Forumrow>说明：如果您已经注册过短信互动服务，请直接在此填写您的认证信息提交至主服务器进行验证。</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>用户名</B>：<br>不超过20位的字母或数字的组合</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="username" value="<%=Rs(0)%>"> <a style="CURSOR:hand" onclick="MM_openBrWindow('install.asp?action=redir10','getpass','width=300,height=100')">检测用户</a></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>密码</B>：</TD>
<TD width=60%  class=Forumrow>
<INPUT type=password size=34 name="password"> <a style="CURSOR:hand" onclick="MM_openBrWindow('install.asp?action=redir9','getpass','width=300,height=100')">忘记密码</a></TD>
</TR>
</tbody>
</table>
</td></tr></tbody></table>
<table cellpadding=0 cellspacing=0 border=0 width='+Dvbbs.Forum_Body[12]+' align=center>
<tr>
<td width=50% height=24> </td>
<td width=50% ><input type=submit value="申请更新资料" name=Submit><input type=reset value="清 除" name=Submit2></td>
</tr></table>
</form>
<%
Rs.Close
Set Rs=Nothing
End Function

Function redir5()
	If request("username")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的用户名。"
		Exit Function
	End If
	If request("password")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的用户密码。"
		Exit Function
	End If
	Get_ChallengeWord

	session("challengeWord_key")=md5(Session("challengeWord") & ":raynetwork",32)
%>
正在提交数据，请稍后……
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

'已注册站长重新注册网站返回结果页面
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
<Th colSpan=2 height=24>更新站长资料</TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>用户名</B>：<br>不超过20位的字母或数字的组合</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="username" value="<%=request("username")%>">&nbsp;<B>*</B> <a style="CURSOR:hand" onclick="MM_openBrWindow('install.asp?action=redir10','getpass','width=300,height=100')">检测用户</a></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>用户密码</B>：<br>8-20位的字母或数字的组合</TD>
<TD width=60%  class=Forumrow>
<INPUT type=password size=30 name="password" value="<%=request("password")%>">&nbsp;<B>*</B> <a style="CURSOR:hand" onclick="MM_openBrWindow('install.asp?action=redir9','getpass','width=300,height=100')">忘记密码</a></TD>
</TR>
<%GetUserInput()%>
</tbody>
</table>
</td></tr></tbody></table>
<table cellpadding=0 cellspacing=0 border=0 width='98%' align=center>
<tr>
<td width=50% height=24> </td>
<td width=50% ><input type=submit value="更新站长资料" name=Submit><input type=reset value="清 除" name=Submit2></td>
</tr></table>
</form>
<%
		'Else
		'	Errmsg=ErrMsg + "<BR><li>验证失败，非法的参数。"
		'	Exit Function
		'End If
	Case 101
		Errmsg=ErrMsg + "<BR><li>您在短信互动服务主服务器注册失败，"&ErrorMsg&"。"
		Exit Function
	Case 102
		Errmsg=ErrMsg + "<BR><li>您注册的用户名和短信互动服务主服务器上的用户名重复，"&ErrorMsg&"。"
		Exit Function
	Case 201
		Errmsg=ErrMsg + "<BR><li>您填写的信息在短信互动服务主服务器上登录验证失败，"&ErrorMsg&"。"
		Exit Function
	Case Else
		Errmsg=ErrMsg + "<BR><li>非法的提交过程，"&ErrorMsg&"。"
		Exit Function
	End Select
	Emp_ChallengeWord
	session("Forum_Master_Reg_Temp_2")=""
End Function

Function redir7()
	If request("username")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的用户名。"
		Exit Function
	End If
	If request("password")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的用户密码。"
		Exit Function
	End If
	If request("realname")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的真实姓名。"
		Exit Function
	End If
	If request("identityNo")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的身份证号。"
		Exit Function
	End If
	If request("sex")="" Then
		Errmsg=ErrMsg + "<BR><li>请选择您的性别。"
		Exit Function
	End If
	If request("postcode")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的邮政编码。"
		Exit Function
	End If
	If request("address")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的地址。"
		Exit Function
	End If
	If request("email")="" Then 
		Errmsg=ErrMsg + "<BR><li>请输入您的邮件地址。"
		Exit Function
	End If
	If request("forumname")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的论坛名称。"
		Exit Function
	End If
	If request("bankname")="" Then
		Errmsg=ErrMsg + "<BR><li>请选择您的开户银行。"
		Exit Function
	End If
	If request("bankid")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的银行帐号。"
		Exit Function
	End If
	If request("bankaddress")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的银行地址。"
		Exit Function
	End If
	If request("telephone")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的联系电话。"
		Exit Function
	End If
	If request("receiver")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的收款人。"
		Exit Function
	End If
	session("Forum_Master_Reg_Temp_1")=checkreal(request("username")) & "|||" & checkreal(request("email")) & "|||" & checkreal(request("forumname")) & "|||" & checkreal(request("forumurl")) & "|||动网先锋|||7.1.0"

	Get_ChallengeWord
	session("challengeWord_key")=md5(Session("challengeWord") & ":raynetwork",32)

%>
正在提交数据，请稍后……
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
<INPUT type=hidden name="forumProvider" value="动网先锋">
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
			Errmsg=ErrMsg + "<BR><li>验证失败，非法的参数。"
			Exit Function
		End If
	Case 101
		Errmsg=ErrMsg + "<BR><li>您在短信互动服务主服务器注册失败，"&ErrorMsg&"。"
		Exit Function
	Case 102
		Errmsg=ErrMsg + "<BR><li>您注册的用户名和短信互动服务主服务器上的用户名重复，"&ErrorMsg&"。"
		Exit Function
	Case 201
		Errmsg=ErrMsg + "<BR><li>您填写的信息在短信互动服务主服务器上登录验证失败，"&ErrorMsg&"。"
		Exit Function
	Case Else
		Errmsg=ErrMsg + "<BR><li>非法的提交过程，"&ErrorMsg&"。"
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
<th height=24>成功信息：您在论坛短信互动服务信息更新成功</th>
</tr>
<tr><td class=Forumrow><br>
<ul><li><a href="challenge.asp">进入您的论坛</a></li></ul>
</td></tr>
</table>
<%
End Function

Function redir1()
	If request("username")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的用户名。"
		Exit Function
	End If
	If request("password")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的用户密码。"
		Exit Function
	End If
	If request("repassword")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的确认密码。"
		Exit Function
	End If
	If Lcase(request("password"))<>Lcase(request("repassword")) Then
		Errmsg=ErrMsg + "<BR><li>您输入的密码和确认密码不一致。"
		Exit Function
	End If
	If request("realname")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的真实姓名。"
		Exit Function
	End If
	If request("identityNo")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的身份证号。"
		Exit Function
	End If
	If request("sex")="" Then
		Errmsg=ErrMsg + "<BR><li>请选择您的性别。"
		Exit Function
	End If
	If request("postcode")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的邮政编码。"
		Exit Function
	End If
	If request("address")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的地址。"
		Exit Function
	End If
	If request("email")="" Then 
		Errmsg=ErrMsg + "<BR><li>请输入您的邮件地址。"
		Exit Function
	End If
	If request("forumname")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的论坛名称。"
		Exit Function
	End If
	If request("bankname")="" Then
		Errmsg=ErrMsg + "<BR><li>请选择您的开户银行。"
		Exit Function
	End If
	If request("bankid")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的银行帐号。"
		Exit Function
	End If
	If request("bankaddress")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的银行地址。"
		Exit Function
	End If
	If request("telephone")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的联系电话。"
		Exit Function
	End If
	If request("receiver")="" Then
		Errmsg=ErrMsg + "<BR><li>请输入您的收款人。"
		Exit Function
	End If
	session("Forum_Master_Reg_Temp_1")=checkreal(request("username")) & "|||" & checkreal(request("email")) & "|||" & checkreal(request("forumname")) & "|||" & checkreal(request("forumurl")) & "|||动网先锋|||7.1.0"

	Get_ChallengeWord
	Session("challengeWord_key")=md5(Session("challengeWord") & ":raynetwork",32)

%>
正在提交数据，请稍后……
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
<INPUT type=hidden name="forumProvider" value="动网先锋">
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
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请输入您的用户名。"
		Exit Function
	End If
	If request("password")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请输入您的密码。"
		Exit Function
	End If
	If request("forumname")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请输入您的论坛名称。"
		Exit Function
	End If
	session("Forum_Master_Reg_Temp_2")=checkreal(request("username")) & "|||" & checkreal(request("password")) & "|||" & checkreal(request("forumname")) & "|||" & checkreal(request("forumurl")) & "|||动网先锋|||7.1.0"

	Get_ChallengeWord

	session("challengeWord_key")=md5(Session("challengeWord") & ":raynetwork",32)
%>
正在提交数据，请稍后……
<form name="redir" action="http://bbs.ray5198.com/bbsapp/oldStnReg.jsp" method="post">
<INPUT type=hidden name="username" value="<%=checkreal(request("username"))%>">
<INPUT type=hidden name="password" value="<%=checkreal(request("password"))%>">
<INPUT type=hidden name="forumName" value="<%=checkreal(request("forumname"))%>">
<INPUT type=hidden name="forumUrl" value="<%=Replace(Dvbbs.Get_ScriptNameUrl,Lcase(Dvbbs.CacheData(33,0)),"")%>">
<INPUT type=hidden name="forumProvider" value="动网先锋">
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

'第一次注册部分接收返回信息
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
			'session("Forum_Master_Reg_Temp_1")=checkreal(request("username")) & "|||" & checkreal(request("email")) & "|||" & checkreal(request("forumname")) & "|||" & checkreal(request("forumurl")) & "|||动网先锋|||7.1.0"
			Dvbbs.Execute("update Dv_ChallengeInfo set D_ForumID='"&reForumID&"',D_UserName='"&Forum_Master_Reg_Temp_1(0)&"',D_Password='"&reNewPkey&"',D_email='"&Forum_Master_Reg_Temp_1(1)&"',D_forumname='"&Forum_Master_Reg_Temp_1(2)&"',D_forumurl='"&Forum_Master_Reg_Temp_1(3)&"',D_forumProvider='"&Forum_Master_Reg_Temp_1(4)&"',D_version='"&Forum_Master_Reg_Temp_1(5)&"',D_challengePassWord='"&OldForumID&"'")
			Dvbbs.loadSetup()
		Else
			Response.Write retokerWord
			Response.Write "<BR>"
			Response.Write challengeWord_key
			Response.end
			'Errmsg=ErrMsg + "<BR><li>验证失败，非法的参数。"
			'Exit Function
		End If
	Case 101
		Errmsg=ErrMsg + "<BR><li>注册失败；"&ErrorMsg&"。"
		Exit Function
	Case 102
		Errmsg=ErrMsg + "<BR><li>您注册的用户名和短信互动服务主服务器上的用户名重复；"&ErrorMsg&"。"
		Exit Function
	Case 201
		Errmsg=ErrMsg + "<BR><li>您填写的信息在短信互动服务主服务器上登录验证失败；"&ErrorMsg&"。"
		Exit Function
	Case Else
		Errmsg=ErrMsg + "<BR><li>非法的提交过程；"&ErrorMsg&"。"
		Exit Function
	End Select
	Emp_ChallengeWord
	Session("Forum_Master_Reg_Temp_1")=""
%>
<table cellpadding=3 cellspacing=1 align=center class=tableBorder>
<tr>
<th height=24>注册成功：您在论坛短信互动服务注册成功</th>
</tr>
<tr><td class=Forumrow><br>
<ul><li>您成功的开启了论坛手机短信服务功能。</li></ul>
</td></tr>
</table>
<%
End Function

'已注册站长重新注册网站返回结果页面
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
			Errmsg=ErrMsg + "<BR><li>验证失败，非法的参数。"
			Exit Function
		End If
	Case 101
		Errmsg=ErrMsg + "<BR><li>您在短信互动服务主服务器注册失败，"&ErrorMsg&"。"
		Exit Function
	Case 102
		Errmsg=ErrMsg + "<BR><li>您注册的用户名和短信互动服务主服务器上的用户名重复，"&ErrorMsg&"。"
		Exit Function
	Case 201
		Errmsg=ErrMsg + "<BR><li>您填写的信息在短信互动服务主服务器上登录验证失败，"&ErrorMsg&"。"
		Exit Function
	Case Else
		Errmsg=ErrMsg + "<BR><li>非法的提交过程，"&ErrorMsg&"。"
		Exit Function
	End Select
	Emp_ChallengeWord
	session("Forum_Master_Reg_Temp_2")=""
%>
<table cellpadding=3 cellspacing=1 align=center class=tableBorder>
<tr>
<th height=24>成功信息：您在论坛短信互动服务信息更新成功</th>
</tr>
<tr><td class=Forumrow><br>
<ul><li><a href="challenge.asp">进入您的论坛</a></li></ul>
</td></tr>
</table>
<%
End Function

function checkreal(v)
Dim w
if not isnull(v) then
	w=replace(v,"|||","§§§")
	checkreal=w
end if
end function

Sub GetUserInput()
%>
<TR>
<TD width=40% class=Forumrow><B>真实姓名</B>：</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="realname" value="<%=Request("realname")%>">&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>身份证号</B>：<BR>支持15位和18位号码</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="identityNo" value="<%=Request("identityNo")%>">&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40%  class=Forumrow><B>性别</B>：<BR>请选择您的性别</font></TD>
<TD width=60%  class=Forumrow> <INPUT type=radio CHECKED value="0" name=sex <%if request("sex")="0" then Response.Write "checked"%>>
男 &nbsp;&nbsp;
<INPUT type=radio value="1" name=sex <%if request("sex")="1" then Response.Write "checked"%>>
女&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>邮编</B>：</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="postcode" value="<%=Request("postcode")%>">&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>地址</B>：</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="address" value="<%=Request("address")%>">&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>邮件地址</B>：<br>为了能成功安装论坛，请务必填写真实有效的Email地址</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="email" value="<%=Request("email")%>">&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>备用邮件</B>：</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="backupemail" value="<%=Request("backupemail")%>"></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>论坛名称</B>：</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="forumname" value="<%=Request("forumname")%>">&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>论坛地址</B>：</TD>
<TD width=60%  class=Forumrow>
<INPUT type=hidden value="<%=Replace(Dvbbs.Get_ScriptNameUrl,Lcase(Dvbbs.CacheData(33,0)),"")%>" name="forumUrl">
<%=Replace(Dvbbs.Get_ScriptNameUrl,Lcase(Dvbbs.CacheData(33,0)),"")%></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>电话</B>：<BR>为了方便我们及时和您联系，请如实填写</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="telephone" value="<%=Request("telephone")%>">&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>手机</B>：</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="mobile" value="<%=Request("mobile")%>"></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>银行名称</B>：</TD>
<TD width=60%  class=Forumrow>
<select name="bankname">
<option value="招商银行" <%if request("bankname")="招商银行" then Response.Write "selected"%>>招商银行</option>
<option value="工商银行" <%if request("bankname")="工商银行" then Response.Write "selected"%>>工商银行</option>
<option value="中国银行" <%if request("bankname")="中国银行" then Response.Write "selected"%>>中国银行</option>
<option value="农业银行" <%if request("bankname")="农业银行" then Response.Write "selected"%>>农业银行</option>
<option value="建设银行" <%if request("bankname")="建设银行" then Response.Write "selected"%>>建设银行</option>
<option value="交通银行" <%if request("bankname")="交通银行" then Response.Write "selected"%>>交通银行</option>
</select>
&nbsp;<B>*</B></TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>银行地址</B>：<BR>XX省XX市/县。例：湖北省武汉市</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="bankaddress" value="<%=Request("bankaddress")%>">&nbsp;<B>*</B></TD>
</TR>
<TD width=40% class=Forumrow><B>银行帐号</B>：<BR>为确保您能收到款项，请仔细填写，中间不要加空格</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="bankid" value="<%=Request("bankid")%>">&nbsp;<B>*</B></TD>
</TR>
<TD width=40% class=Forumrow><B>收款人姓名</B>(开户人)：<BR>为确保您能收到款项，请仔细填写</TD>
<TD width=60%  class=Forumrow>
<INPUT type=text size=30 name="receiver" value="<%=Request("receiver")%>">&nbsp;<B>*</B></TD>
</TR>
<TR align=middle>
<Th colSpan=2 height=24>论坛资料填写</TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>论坛提供者</B>：</TD>
<TD width=60%  class=Forumrow>
动网先锋</TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>论坛版本</B>：</TD>
<TD width=60%  class=Forumrow>
Dvbbs 7.1.0</TD>
</TR>
<TR>
<TD width=40% class=Forumrow><B>网站内容介绍</B>：</TD>
<TD width=60%  class=Forumrow>
<textarea class=smallarea cols=50 name=web_intro rows=7 wrap=VIRTUAL><%=Request("web_intro")%></textarea>
</TD>
</TR>
<%
End Sub

'取回密码
Sub Redir9()
	If Request("t")="t" Then
		If Request("ErrorCode")="-1" Then
			Response.Write "您的密码信息已成功发往您注册的邮箱，请留意查收"
		Else
			Response.Write "发送失败"
		End If
	Else
%>
<table cellpadding=3 cellspacing=1 align=center class=tableBorder>
<FORM METHOD=POST ACTION="http://bbs.ray5198.com/bbsapp/sendpassword.jsp">
<tr>
<th height=24>忘记密码</th>
</tr>
<tr><td class=Forumrow><br>
请输入您的用户名：<input type=text size=30 name="username"> <input type=submit name=submit value="提交">
<input type=hidden value="<%=Replace(Dvbbs.Get_ScriptNameUrl,Lcase(Dvbbs.CacheData(33,0)),"")%><%=Dvbbs.CacheData(33,0)%>install.asp?action=redir9&t=t" name="dirPage">
</td></tr>
</FORM>
</table>
<%
	End If
End Sub
'检测用户
Sub Redir10()
	If Request("t")="t" Then
		If Request("ErrorCode")="-1" Then
			Response.Write "该帐号尚未注册"
		Else
			Response.Write "该帐号已被注册"
		End If
	Else
%>
<table cellpadding=3 cellspacing=1 align=center class=tableBorder>
<FORM METHOD=POST ACTION="http://bbs.ray5198.com/bbsapp/userexist.jsp">
<tr>
<th height=24>检测用户是否存在</th>
</tr>
<tr><td class=Forumrow><br>
请输入您要注册的用户名：<input type=text size=30 name="username"> <input type=submit name=submit value="提交">
<input type=hidden value="<%=Replace(Dvbbs.Get_ScriptNameUrl,Lcase(Dvbbs.CacheData(33,0)),"")%><%=Dvbbs.CacheData(33,0)%>install.asp?action=redir10&t=t" name="dirPage">
</td></tr>
</FORM>
</table>
<%
	End If
End Sub
%>