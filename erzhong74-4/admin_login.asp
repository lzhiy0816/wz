<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/md5.asp" -->
<!-- #include file="inc/myadmin.asp" -->

<%if not session("checked")="ez74yes" then
response.Redirect "ez74login.asp"

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
Dvbbs.LoadTemplates("Admin")
template.ChildFolder="Admin"
'Set Rs=Dvbbs.Execute("Select H_Content From Dv_Help Where H_ID=1")
'Response.Write Rs(0)
'template.value = Rs(0)
'Response.End
Dvbbs.Stats="��̳������¼"

Admin_Login()
Dvbbs.PageEnd()
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
	If IsSqlDataBase = 1 Then version="SQL ��"&fversion Else version="ACCESS ��"&fversion
	'Response.Write Replace(template.html(1),"{$path}",Dvbbs.CacheData(33,0) & "images/")
	Response.Write Replace(template.html(1),"{$path}","")
%>
<style type="text/css">
body { background:#fff; background-image : url("skins/images/body_bg.gif");background-repeat: repeat-x ;  }
td { font-size:12px;}
input { border:1px solid #999; }
.button { color: #135294; border:1px solid #666; height:21px; line-height:18px; background:url("images/button_bg.gif")}
div#nifty{margin: 0 10%;background: #ABD4EF;width: 420px;word-break:break-all; margin-top:60px;}
b.rtop, b.rbottom{display:block;background: #FFF}
b.rtop b, b.rbottom b{display:block;height: 1px;overflow: hidden; background: #ABD4EF}
b.r1{margin: 0 5px}
b.r2{margin: 0 3px}
b.r3{margin: 0 2px}
b.rtop b.r4, b.rbottom b.r4{margin: 0 1px;height: 2px}
</style>
<center>
<div id="nifty">
<b class="rtop"><b class="r1"></b><b class="r2"></b><b class="r3"></b><b class="r4"></b></b>
<div style="width:403px; height:26px; line-height:26px; background:none; font-size:12px; text-align:left;"><%=dvbbs.Forum_info(0)%> -- ������¼</div>
<div style="width:403px; height:46px; background:#166CA3;"><img src="images/login.gif" alt="" /></div>
<div style="width:401px !important; width:403px; height:auto; background:#fff; border-left:1px solid #649EB2; border-right:1px solid #649EB2; ">
<table width="100%" border="0" cellspacing="3" cellpadding="0">
<form action="admin_login.asp" method="post">
<input name="reaction" type="hidden" value="chklogin" />
	<%If Dvbbs.UserID=0 Or (Dvbbs.UserGroupID>1 And Dvbbs.GroupSetting(70)="0") Then%>
	<tr>
		<td align="right" width="35%"><b>ǰ̨�û�����</b></td>
		<td align="left"><input name="adduser" type="text" tabindex="2"/></td>
	</tr>
	<tr>
		<td align="right" width="35%"><b>ǰ̨���룺</b></td>
		<td align="left"><input name="password2" type="password" tabindex="3"/></td>
	</tr>
	<%End If%>
	<tr>
		<td align="right"><b>�û�����</b></td>
		<td align="left"><input name="username" type="text" tabindex="4"/></td>
	</tr>
	<tr>
		<td align="right"><b>�ܡ��룺</b></td>
		<td align="left"><input name="password" type="password" tabindex="5"/></td>
	</tr>
	<%If CHECK_CODE=1 Then%>
	<tr>	
		<td align="right"><b>�����룺</b></td>
		<td align="left"><%=Dvbbs.GetCode%></td>
	</tr>
	<%End If%>
	<tr>
	<td align="right"></td>
	<td align="left"><input  class="button" type="submit" name="submit" value="�� ¼"/></td>
	</tr>	
  </form>
</table>
</div>
<div style="width:401px !important; width:403px; height:20px; background:#F7F7E7; border:1px solid #649EB2; border-top:1px solid #ddd; margin-bottom:5px; font-size:12px; line-height:20px; "> <%=Dvbbs.Forum_info(0)%> <%If Dvbbs.UserID>0 Then Response.Write Version%></div>
<b class="rbottom"><b class="r4"></b><b class="r3"></b><b class="r2"></b><b class="r1"></b>
</div>
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
		If Request("codestr")="" Then
			Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�뷵������ȷ���롣<b>���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ��</b>"
			Exit Sub
		Elseif Session("getcode")="9999" then
			Session("getcode")=""
		Elseif Session("getcode")="" then
			Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�벻Ҫ�ظ��ύ���������µ�¼�뷵�ص�¼ҳ�档<b>���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ��</b>"
			Exit Sub
		ElseIf Cstr(Session("getcode"))<>Lcase(Cstr(Trim(Request("codestr")))) Then
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
					Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>��û�е�½��̨������Ȩ�ޣ�"
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