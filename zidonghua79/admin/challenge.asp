<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<%	
	Head()
	Dim admin_flag,rs_c
	admin_flag=",1,"
	If Not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 Then 
		Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=../admin_login.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		Call dvbbs_error()
	Else
		If Request("action")="Save1" Then
			Save1()
		ElseIf Request("action")="Save2" Then
			Save2()
		ElseIf Request("action")="Save3" Then
			Save3()
		ElseIf Request("action")="Save4" Then
			Save4()
		Else
			consted()
		End If
		If Errmsg <> "" Then Dvbbs_Error()
		Footer()
	End If


Sub consted()
Dim  sel
%>
<form method="POST" action="Challenge.asp?action=Save1">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder" style="LINE-HEIGHT: 110%;"> 
<th height=25 colspan=2 align=center id=tabletitlelink><a name="setting20"></a><b>RSS/WAP/手机短信/在线支付</b>
</tr>

<!--网络支付部分开始-->
<tr>
<td width="50%" class=Forumrow> <U>是否开启网络银行充值点券</U><br>开通后可通过网络支付手段来充值论坛点券</td>
<td width="50%" class=Forumrow>
<input type=radio name="Forum_ChanSetting(3)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(3))=1 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_ChanSetting(3)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(3))=0 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr>
<td width="50%" class=Forumrow> <U>支付宝账号</U><br>请到<a href="https://www.alipay.com/" target=_blank><font color=red>支付宝网站</font></a>申请一个支付宝账号，然后填写到表单中</td>
<td width="50%" class=Forumrow>  
<input type=text size=35 value="<%=Dvbbs.Forum_ChanSetting(4)%>" name="Forum_ChanSetting(4)">
</td>
</tr>
<tr>
<td width="50%" class=Forumrow> <U>安全校验码</U><br>请参阅下述说明流程进行设置</td>
<td width="50%" class=Forumrow>  
<input type=text size=35 value="<%=Dvbbs.Forum_ChanSetting(6)%>" name="Forum_ChanSetting(6)">&nbsp;&nbsp;&nbsp;<input type="submit" name="Submit" value="提交修改">
</td>
</tr>
<tr>
<td width="100%" class=forumRowHighlight colspan=2 height=25><B>说明</B>：<BR>
1、网络银行支付功能由“阿里巴巴支付宝”提供，开通此功能需到<a href="https://www.alipay.com/" target=_blank><font color=red>支付宝网站</font></a>申请一个支付宝账号并填入上述表单，不正确的账号将会影响到您的网站收益<BR>
2、<font color=blue>请设置正确的安全校验码，流程<BR>登录<a href="https://www.alipay.com/" target=_blank><font color=blue>支付宝网站</font></a> -> 进入商家工具 -> 点击<B>开发集成</B>一栏表格 -> 点击右边的<B>设置安全校验码</B> -> <font color=red>获取校验码并填入上述表单</font> -> 点击<B>设置通知接受URL</B> -> 接受URL设置为<B><%=Dvbbs_PayTo_Url%>top.asp</B></font><BR>
3、<font color=red><B>通过本功能交易手续费用全免，支付宝交易免费、安全、便捷！用户支付的款项将直接转到您指定的支付宝帐号！</B></font><BR>
</td>
</tr>
</form>
<form method="POST" action="Challenge.asp?action=Save2">
<!--手机部分开始-->
<tr> 
<td width="50%" class=Forumrow> <U>是否开启论坛手机相关功能</U><BR>手机相关功能总开关</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(0)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(0))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_ChanSetting(0)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(0))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<td width="50%" class=Forumrow> <U>是否开启论坛WAP</U><br>开启后通过手机可浏览论坛和进行发贴等操作</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(1)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(1))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_ChanSetting(1)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(1))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr>
<td width="50%" class=Forumrow> <U>是否开启手机短信充值点券</U><br>开启后可通过手机短信进行论坛点券充值　</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(13)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(13))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_ChanSetting(13)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(13))=1 then%>checked<%end if%>>是&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="Submit" value="提 交">
</td>
</tr>
<tr>
<td width="100%" class=forumRowHighlight colspan=2 height=25><B>说明</B>：<BR>相关的手机无线产品服务由北京阳光加信科技有限公司提供，开通此功能则默认接受相关的服务条款<BR><BR>
<%
dim trs
set trs=Dvbbs.Execute("select * from Dv_ChallengeInfo")
set rs=Dvbbs.Execute("select * from dv_setup")
%>
<%if Dvbbs.Forum_ChanSetting(0)="1" and rs("Forum_isinstall")=1 then%>
您已经安装了论坛的手机无线产品服务，具体事项请看相关说明<BR>
您当前注册的手机无线产品资料是，用户名：<%=trs("d_username")%>，网站名：<%=trs("d_forumname")%>，论坛地址：<%if trs("d_forumurl")="" then%><%=Dvbbs.Get_ScriptNameUrl()%><%else%><%=trs("d_forumurl")%><%end if%>，如果这些资料和您当前所使用的论坛不符（如论坛地址或用户名），您将不能得到相关的短信收益。<BR>
<a href="install.asp?isnew=1"><font color=blue>您可以点击此处进行资料更新或者重新注册站长资料</font></a>
<%else%>
您还没有安装论坛的手机无线产品服务，通过手机无线产品服务，您可以享受到各种不同的网站收益，具体请看关于<a href=""><font color=red>手机无线产品服务的说明</font></a>，<a href="install.asp"><font color=blue>点击此处开启论坛手机无线产品服务</font></a>
<%end if%>
<%
rs.close
set rs=nothing
trs.close
set trs=nothing
%>
</td>
</tr>
<!--手机部分结束-->
</form>
<form method="POST" action="Challenge.asp?action=Save3">
<tr>
<td width="50%" class=Forumrow> <U>充值的点券兑换率</U><BR>包括手机短信和网络支付方式</td>
<td width="50%" class=Forumrow>  
<input type=text size=5 value="<%=Dvbbs.Forum_ChanSetting(14)%>" name="Forum_ChanSetting(14)"> 张点券=1元人民币或短信
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否开启RSS订阅功能</U><BR>开启后可通过一些RSS阅读软件订阅</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(2)" value=1 <%if (Dvbbs.Forum_ChanSetting(2))="1" then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_ChanSetting(2)" value=0 <%if (Dvbbs.Forum_ChanSetting(2))="0" then%>checked<%end if%>>是&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="Submit" value="提 交">
</td>
</tr>
</table>
</form>
<%
end sub

'保存网络支付相关
Sub Save1()
	'3是否开启/4支付宝账号/5动网关联ID/6Forum_AliPayKey
	Dim Forum_ChanSetting,iForum_ChanSetting,mForum_ChanSetting
	Set Rs=Dvbbs.Execute("Select Forum_ChanSetting From Dv_Setup")
	Forum_ChanSetting = Rs(0)
	Forum_ChanSetting = Split(Forum_ChanSetting,",")
	Rs.Close
	Set Rs=Nothing

	For i = 0 To Ubound(Forum_ChanSetting)
		If i = 0 Then
			iForum_ChanSetting = Forum_ChanSetting(i)
		Else
			Select Case i
			Case 3
				iForum_ChanSetting = iForum_ChanSetting & "," & Replace(Replace(Request.Form("Forum_ChanSetting(3)"),"'",""),",","")
			Case 4
				iForum_ChanSetting = iForum_ChanSetting & "," & Replace(Replace(Request.Form("Forum_ChanSetting(4)"),"'",""),",","")
			Case 5
				iForum_ChanSetting = iForum_ChanSetting & ",0"
			Case 6
				iForum_ChanSetting = iForum_ChanSetting & "," & Replace(Replace(Request.Form("Forum_ChanSetting(6)"),"'",""),",","")
			Case Else
				iForum_ChanSetting = iForum_ChanSetting & "," & Forum_ChanSetting(i)
			End Select
		End If
	Next
	Sql="Update Dv_Setup Set Forum_ChanSetting='"&iForum_ChanSetting&"'"
	Dvbbs.Execute(Sql)
	Dvbbs.loadSetup()
	Dv_suc("网络支付设置成功！")
End Sub

'保存从主服务器返回信息
Sub Save4()
	'3是否开启/4支付宝账号/5动网关联ID/6Forum_AliPayKey
	If Request("UserID")="" Or Request("Email")="" Or Request("ForumKey")="" Then
		Errmsg=ErrMsg + "<BR><li>非法的返回参数。"
		Exit Sub
	End If
	Dim Forum_ChanSetting,iForum_ChanSetting,mForum_ChanSetting
	Set Rs=Dvbbs.Execute("Select Forum_ChanSetting From Dv_Setup")
	Forum_ChanSetting = Rs(0)
	Forum_ChanSetting = Split(Forum_ChanSetting,",")
	mForum_ChanSetting = False
	For i = 0 To Ubound(Forum_ChanSetting)
		If i = 0 Then
			iForum_ChanSetting = Forum_ChanSetting(i)
		Else
			Select Case i
			Case 3
				iForum_ChanSetting = iForum_ChanSetting & ",0"
			Case 4
				iForum_ChanSetting = iForum_ChanSetting & "," & Replace(Replace(Request("Email"),"'",""),",","")
			Case 5
				iForum_ChanSetting = iForum_ChanSetting & "," & Replace(Replace(Request("UserID"),"'",""),",","")
			Case 6
				iForum_ChanSetting = iForum_ChanSetting & "," & Replace(Replace(Request("ForumKey"),"'",""),",","")
			Case Else
				iForum_ChanSetting = iForum_ChanSetting & "," & Forum_ChanSetting(i)
			End Select
		End If
	Next
	Sql="Update Dv_Setup Set Forum_ChanSetting='"&iForum_ChanSetting&"'"
	Dvbbs.Execute(Sql)
	Dvbbs.loadSetup()
	Rs.Close
	Set Rs=Nothing
	Dv_suc("主服务器网络支付功能设置成功！" & Request("msg"))
End Sub

'保存手机短信相关
Sub Save2()
	Dim Forum_ChanSetting,iForum_ChanSetting,mForum_ChanSetting
	Set Rs=Dvbbs.Execute("Select Forum_ChanSetting,Forum_challengePassWord From Dv_Setup")
	Forum_ChanSetting = Rs(0)
	Forum_ChanSetting = Split(Forum_ChanSetting,",")
	mForum_ChanSetting = False
	For i = 0 To Ubound(Forum_ChanSetting)
		If i = 0 Then
			iForum_ChanSetting = Request.Form("Forum_ChanSetting("&i&")")
			If (Cint(Replace(Request.Form("Forum_ChanSetting("&i&")"),",",""))=1 And Rs(1)="raynetwork") Or (Cint(Replace(Request.Form("Forum_ChanSetting("&i&")"),",",""))=1 And Cint(Forum_ChanSetting(i))=0) Then
				mForum_ChanSetting = True
			End If
		Else
			Select Case i
			Case 1
				iForum_ChanSetting = iForum_ChanSetting & "," & Replace(Request.Form("Forum_ChanSetting("&i&")"),",","")
			Case 13
				iForum_ChanSetting = iForum_ChanSetting & "," & Replace(Request.Form("Forum_ChanSetting("&i&")"),",","")
			Case Else
				iForum_ChanSetting = iForum_ChanSetting & "," & Forum_ChanSetting(i)
			End Select
		End If
	Next
	Sql="Update Dv_Setup Set Forum_ChanSetting='"&iForum_ChanSetting&"'"
	Dvbbs.Execute(Sql)
	Dvbbs.loadSetup()
	Rs.Close
	Set Rs=Nothing
	If mForum_ChanSetting Then Response.Redirect "install.asp?isnew=1"
	Dv_suc("手机短信设置成功！")
End Sub

'保存杂项
Sub Save3()
	Dim Forum_ChanSetting,iForum_ChanSetting,mForum_ChanSetting
	Set Rs=Dvbbs.Execute("Select Forum_ChanSetting,Forum_challengePassWord From Dv_Setup")
	Forum_ChanSetting = Rs(0)
	Forum_ChanSetting = Split(Forum_ChanSetting,",")
	mForum_ChanSetting = False
	For i = 0 To Ubound(Forum_ChanSetting)
		If i = 0 Then
			iForum_ChanSetting = Forum_ChanSetting(i)
		Else
			Select Case i
			Case 2
				iForum_ChanSetting = iForum_ChanSetting & "," & Replace(Request.Form("Forum_ChanSetting("&i&")"),",","")
			Case 14
				iForum_ChanSetting = iForum_ChanSetting & "," & Replace(Request.Form("Forum_ChanSetting("&i&")"),",","")
			Case Else
				iForum_ChanSetting = iForum_ChanSetting & "," & Forum_ChanSetting(i)
			End Select
		End If
	Next
	Sql="Update Dv_Setup Set Forum_ChanSetting='"&iForum_ChanSetting&"'"
	Dvbbs.Execute(Sql)
	Rs.Close
	Set Rs=Nothing
	Dv_suc("RSS设置成功！")
	Dvbbs.loadSetup()
End Sub
%>