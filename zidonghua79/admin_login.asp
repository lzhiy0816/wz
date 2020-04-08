<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/md5.asp" -->
<!-- #include file="inc/myadmin.asp" -->
<%
if not session("checked")="z79yes" then
response.Redirect "z79login.asp"

else

Rem ===============================================================
Rem 开启使用允许IP登陆功能 Chk_IPLogin : 0- 关闭，1=开启
Const Chk_IPLogin = 1
Rem ===============================================================
Rem ===============================================================
Rem CHECK_CODE 设置为1 开通登录验证码，设置为0关闭登录验证码，以方便视障人士如盲人等朋友使用。
Const CHECK_CODE=1
Rem ===============================================================
Dim Rs,sql,i
Dvbbs.LoadTemplates("")
Set Rs=Dvbbs.Execute("Select H_Content From Dv_Help Where H_ID=1")
template.value = Rs(0)

Dvbbs.Stats="论坛管理登录"

Admin_Login()

Sub Admin_Login()
	'Response.Write Dvbbs.CacheData(33,0)
	Dvbbs.Head()

	If (Dvbbs.GroupSetting(70)="1" and Dvbbs.UserGroupID>1 and Dvbbs.UserID>0) or Dvbbs.Master or Dvbbs.UserID=0  Then
		Dvbbs.Master = True
	Else
		Dvbbs.Master = False
	End If
	If Not Dvbbs.Master Then Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>您不是系统管理员！"

	If Dvbbs.Master And Session("flag")<>"" Then Response.Redirect Dvbbs.CacheData(33,0) & "index.asp"
	If Request.form("reaction")="chklogin" Then
		ChkLogin()
	Else
		Admin_Login_Main()
	End If
End Sub

Sub Admin_Login_Main()
	Dim version
	If IsSqlDataBase = 1 Then version="SQL 版" Else version="ACCESS 版"
	Response.Write Replace(template.html(1),"{$path}",Dvbbs.CacheData(33,0) & "images/")
%>
<p>&nbsp;</p>
<p>&nbsp;</p>
<form action="admin_login.asp" method="post">
<input name="reaction" type="hidden" value="chklogin" />
<table cellpadding="1" cellspacing="0" border="0" align=center style="border: outset 3px;width:0;">
<tr>
	<td>
	<table style="width:500px;" border="0" cellspacing="0" cellpadding="3" align="center" class="tablefoot">
		<tr><th valign="middle" colspan="2"><%=dvbbs.Forum_info(0)%>管理登录</th></tr>
	</table>
	<table style="width:500px;" border="0" cellspacing="0" cellpadding="3" align="center" class="tablefoot">
	<tr>
	<td valign="middle" colspan="2" align="center" class="forumRowHighlight" style="background-image: url(<%=Dvbbs.CacheData(33,0)%>images/loginbg.jpg);height:75px;">
		<table border="0" width="100%" height="100%">
		<tr>
		<td width="61%" height="100%" rowspan="3"></td>
		<td width="39%" height="0"></td>
		</tr>
		<tr>
		<td valign="top" class="tdfoot" align="left"><br /><a href="index.asp"><b><%=Dvbbs.Forum_info(0)%></b></a><br />版本：Dvbbs v7.1.0 <%If Dvbbs.UserID>0 Then Response.Write Version%></td>
		</tr>
		<tr><td></td></tr>
		</table>
	</td>
	</tr>
	</table>
	<table style="width:500px;" border="0" cellspacing="0" cellpadding="3" align="center">
	<tr>
	<td valign="middle" colspan="2" align="center" class="forumRowHighligh"t height="4"></td>
	</tr>
	<%If Dvbbs.UserID=0 Or (Dvbbs.UserGroupID>1 And Dvbbs.GroupSetting(70)="0") Then%>
	<tr>
	<td valign="middle" class="forumRow" width="30%" align="right"><b>前台用户名：</b></td>
	<td valign="middle" class="forumRow"><input name="adduser" type="text" /></td></tr>
	<tr>
	<td valign="middle" class="forumRow" align="right"><b>前台密码：</b></font></td>
	<td valign="middle" class="forumRow"><input name="password2" type="password" /></td></tr>
	<%End If%>
	<tr>
	<td valign="middle" class="forumRow" width="30%" align="right"><b>用户名：</b></td>
	<td valign="middle" class="forumRow" align="left"><input name="username" type="text" /></td></tr>
	<tr>
	<td valign="middle" class="forumRow" align="right"><b>密　码：</b></font></td>
	<td valign="middle" class="forumRow" align="left"><input name="password" type="password" /></td></tr>
	<tr>
	<%If CHECK_CODE=1 Then%>
	<td valign="middle" class="forumRow" align="right"><b>附加码：</b></td>
	<td valign="middle" class="forumRow" align="left" ><input name="verifycode" type="text">&nbsp;请在附加码框输入 <img src="Dv_getcode.asp"  alt="验证码,看不清楚?请点击刷新验证码" height="16" style="cursor : pointer;" onclick="this.src='DV_getcode.asp'" /></td></tr>
	<%End If%>
	<tr>
	<td valign="middle" colspan="2" align="center" class="forumRowHighlight"><input class="button" type="submit" name="submit" value="登 录" /></td>
	</tr>
	</table>
	</td>
</tr>
</table>
</form>

</body>
</html>
<%

End Sub

Sub ChkLogin()
	Dim ip
	Dim UserName
	Dim PassWord
	UserName=Replace(Request("username"),"'","")
	PassWord=md5(request("password"),16)
	If CHECK_CODE=1 Then
		If Request("verifycode")="" Then
			Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请返回输入确认码。<b>返回后请刷新登录页面后重新输入正确的信息。</b>"
			Exit Sub
		Elseif Session("getcode")="9999" then
			Session("getcode")=""
		Elseif Session("getcode")="" then
			Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请不要重复提交，如需重新登录请返回登录页面。<b>返回后请刷新登录页面后重新输入正确的信息。</b>"
			Exit Sub
		ElseIf Cstr(Session("getcode"))<>Lcase(Cstr(Trim(Request("verifycode")))) Then
			Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>您输入的确认码和系统产生的不一致，请重新输入。<b>返回后请刷新登录页面后重新输入正确的信息。</b>"
			Exit Sub
		End If
		Session("getcode")=""
	End If
	if UserName="" Or PassWord="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请输入您的用户名或密码。<b>返回后请刷新登录页面后重新输入正确的信息。</b>"
		Exit Sub
	End If
	ip=Dvbbs.UserTrueIP
	Dim MemberName
	If Dvbbs.MemberName=""  Or Request("adduser") <>"" Then 
		MemberName=Dvbbs.Checkstr(Request("adduser"))
	Else
		MemberName=Dvbbs.MemberName
	End If
	Set Rs=Dvbbs.Execute("Select a.*,u.userpassword,u.usergroupid From "&admintable&"  a Inner Join Dv_user u On a.adduser=u.userName Where a.UserName='"&username&"' And AddUser='"&MemberName&"'")
	If Rs.Eof And Rs.Bof Then
		Rs.Close
		Set Rs=Nothing
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>您输入的用户名和密码不正确或者您不是系统管理员。请<a href=admin_login.asp>重新输入</a>您的密码。<b>返回后请刷新登录页面后重新输入正确的信息。</b>"
		Exit Sub
	Else
		If Rs("AcceptIP")<>"" And Chk_IPLogin=1 Then
			If ChkLoginIP(Rs("AcceptIP"),ip)=False Then
				Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>你不是合法的后台管理员。请<a href=admin_login.asp>重新输入</a>您的密码。"
				Exit Sub
			End If
		End If
		If Trim(Rs("password"))<>PassWord then
			Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>您输入的用户名和密码不正确或者您不是系统管理员。请<a href=admin_login.asp>重新输入</a>您的密码。<b>返回后请刷新登录页面后重新输入正确的信息。</b>"
			Exit Sub
		Else
			If Dvbbs.MemberName=""  Or Request("adduser") <>"" Then 
				If Trim(Rs("userpassword"))<>md5(Request("password2"),16) Then
					Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>您输入的用户名和密码不正确或者您不是系统管理员。请<a href=admin_login.asp>重新输入</a>您的密码。<b>返回后请刷新登录页面后重新输入正确的信息。</b>"		
					Exit Sub
				End If
			End If
			Dim Rs1	'在此验证GroupSetting(70)，轻飘飘
			Set Rs1=Dvbbs.Execute("Select GroupSetting From Dv_UserGroups Where UserGroupID="&Rs("usergroupid"))
			If Rs1.Eof Or Rs1.Bof Then
				Rs.Close
				Set Rs=Nothing
				Set Rs1=Nothing
				Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>您输入的用户名和密码不正确或者您不是系统管理员。请<a href=admin_login.asp>重新输入</a>您的密码。<b>返回后请刷新登录页面后重新输入正确的信息。</b>"
			Else
				If Split(Rs1(0),",")(70)="1" Then
					Dvbbs.Execute("Update "&admintable&" Set LastLogin="&SqlNowString&",LastLoginIP='"&ip&"' Where UserName='"&UserName&"'")
					Session("flag")=Rs("flag")
					Session.Timeout=45
					Session("MemberName")=MemberName
					Response.Redirect Dvbbs.CacheData(33,0) & "index.asp"
				Else
					Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>您没有登陆后台管理的权限！"
				End If
			End If		
			Rs.Close
			Set Rs=Nothing
			Set Rs1=Nothing
		End If
	End If
End Sub

Function ChkLoginIP(AcceptIP,ChkIp)
	Dim i,LoginIP,TempIP
	ChkLoginIP = False
	If Instr("|"&AcceptIP&"|","|"&ChkIp&"|") Then ChkLoginIP = True : Exit Function
	LoginIP = Split(ChkIp,".")
	TempIP = LoginIP(0)&"."&LoginIP(1)&"."&LoginIP(2)&".*"
	If Instr("|"&AcceptIP&"|","|"&TempIP&"|") Then ChkLoginIP = True : Exit Function
	TempIP = LoginIP(0)&"."&LoginIP(1)&".*.*"
	If Instr("|"&AcceptIP&"|","|"&TempIP&"|") Then ChkLoginIP = True : Exit Function
	TempIP = LoginIP(0)&".*.*.*"
	If Instr("|"&AcceptIP&"|","|"&TempIP&"|") Then ChkLoginIP = True : Exit Function
End Function
end if
%>