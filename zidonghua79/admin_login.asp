<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/md5.asp" -->
<!-- #include file="inc/myadmin.asp" -->
<%
if not session("checked")="z79yes" then
response.Redirect "z79login.asp"

else

Rem ===============================================================
Rem ����ʹ������IP��½���� Chk_IPLogin : 0- �رգ�1=����
Const Chk_IPLogin = 1
Rem ===============================================================
Rem ===============================================================
Rem CHECK_CODE ����Ϊ1 ��ͨ��¼��֤�룬����Ϊ0�رյ�¼��֤�룬�Է���������ʿ��ä�˵�����ʹ�á�
Const CHECK_CODE=1
Rem ===============================================================
Dim Rs,sql,i
Dvbbs.LoadTemplates("")
Set Rs=Dvbbs.Execute("Select H_Content From Dv_Help Where H_ID=1")
template.value = Rs(0)

Dvbbs.Stats="��̳�����¼"

Admin_Login()

Sub Admin_Login()
	'Response.Write Dvbbs.CacheData(33,0)
	Dvbbs.Head()

	If (Dvbbs.GroupSetting(70)="1" and Dvbbs.UserGroupID>1 and Dvbbs.UserID>0) or Dvbbs.Master or Dvbbs.UserID=0  Then
		Dvbbs.Master = True
	Else
		Dvbbs.Master = False
	End If
	If Not Dvbbs.Master Then Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>������ϵͳ����Ա��"

	If Dvbbs.Master And Session("flag")<>"" Then Response.Redirect Dvbbs.CacheData(33,0) & "index.asp"
	If Request.form("reaction")="chklogin" Then
		ChkLogin()
	Else
		Admin_Login_Main()
	End If
End Sub

Sub Admin_Login_Main()
	Dim version
	If IsSqlDataBase = 1 Then version="SQL ��" Else version="ACCESS ��"
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
		<tr><th valign="middle" colspan="2"><%=dvbbs.Forum_info(0)%>�����¼</th></tr>
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
		<td valign="top" class="tdfoot" align="left"><br /><a href="index.asp"><b><%=Dvbbs.Forum_info(0)%></b></a><br />�汾��Dvbbs v7.1.0 <%If Dvbbs.UserID>0 Then Response.Write Version%></td>
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
	<td valign="middle" class="forumRow" width="30%" align="right"><b>ǰ̨�û�����</b></td>
	<td valign="middle" class="forumRow"><input name="adduser" type="text" /></td></tr>
	<tr>
	<td valign="middle" class="forumRow" align="right"><b>ǰ̨���룺</b></font></td>
	<td valign="middle" class="forumRow"><input name="password2" type="password" /></td></tr>
	<%End If%>
	<tr>
	<td valign="middle" class="forumRow" width="30%" align="right"><b>�û�����</b></td>
	<td valign="middle" class="forumRow" align="left"><input name="username" type="text" /></td></tr>
	<tr>
	<td valign="middle" class="forumRow" align="right"><b>�ܡ��룺</b></font></td>
	<td valign="middle" class="forumRow" align="left"><input name="password" type="password" /></td></tr>
	<tr>
	<%If CHECK_CODE=1 Then%>
	<td valign="middle" class="forumRow" align="right"><b>�����룺</b></td>
	<td valign="middle" class="forumRow" align="left" ><input name="verifycode" type="text">&nbsp;���ڸ���������� <img src="Dv_getcode.asp"  alt="��֤��,�������?����ˢ����֤��" height="16" style="cursor : pointer;" onclick="this.src='DV_getcode.asp'" /></td></tr>
	<%End If%>
	<tr>
	<td valign="middle" colspan="2" align="center" class="forumRowHighlight"><input class="button" type="submit" name="submit" value="�� ¼" /></td>
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
			Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�뷵������ȷ���롣<b>���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ��</b>"
			Exit Sub
		Elseif Session("getcode")="9999" then
			Session("getcode")=""
		Elseif Session("getcode")="" then
			Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�벻Ҫ�ظ��ύ���������µ�¼�뷵�ص�¼ҳ�档<b>���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ��</b>"
			Exit Sub
		ElseIf Cstr(Session("getcode"))<>Lcase(Cstr(Trim(Request("verifycode")))) Then
			Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�������ȷ�����ϵͳ�����Ĳ�һ�£����������롣<b>���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ��</b>"
			Exit Sub
		End If
		Session("getcode")=""
	End If
	if UserName="" Or PassWord="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�����������û��������롣<b>���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ��</b>"
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
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>��������û��������벻��ȷ����������ϵͳ����Ա����<a href=admin_login.asp>��������</a>�������롣<b>���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ��</b>"
		Exit Sub
	Else
		If Rs("AcceptIP")<>"" And Chk_IPLogin=1 Then
			If ChkLoginIP(Rs("AcceptIP"),ip)=False Then
				Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�㲻�ǺϷ��ĺ�̨����Ա����<a href=admin_login.asp>��������</a>�������롣"
				Exit Sub
			End If
		End If
		If Trim(Rs("password"))<>PassWord then
			Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>��������û��������벻��ȷ����������ϵͳ����Ա����<a href=admin_login.asp>��������</a>�������롣<b>���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ��</b>"
			Exit Sub
		Else
			If Dvbbs.MemberName=""  Or Request("adduser") <>"" Then 
				If Trim(Rs("userpassword"))<>md5(Request("password2"),16) Then
					Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>��������û��������벻��ȷ����������ϵͳ����Ա����<a href=admin_login.asp>��������</a>�������롣<b>���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ��</b>"		
					Exit Sub
				End If
			End If
			Dim Rs1	'�ڴ���֤GroupSetting(70)����ƮƮ
			Set Rs1=Dvbbs.Execute("Select GroupSetting From Dv_UserGroups Where UserGroupID="&Rs("usergroupid"))
			If Rs1.Eof Or Rs1.Bof Then
				Rs.Close
				Set Rs=Nothing
				Set Rs1=Nothing
				Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>��������û��������벻��ȷ����������ϵͳ����Ա����<a href=admin_login.asp>��������</a>�������롣<b>���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ��</b>"
			Else
				If Split(Rs1(0),",")(70)="1" Then
					Dvbbs.Execute("Update "&admintable&" Set LastLogin="&SqlNowString&",LastLoginIP='"&ip&"' Where UserName='"&UserName&"'")
					Session("flag")=Rs("flag")
					Session.Timeout=45
					Session("MemberName")=MemberName
					Response.Redirect Dvbbs.CacheData(33,0) & "index.asp"
				Else
					Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>��û�е�½��̨�����Ȩ�ޣ�"
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