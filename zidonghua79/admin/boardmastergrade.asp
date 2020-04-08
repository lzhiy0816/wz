<!--#include file="../conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
Head()
Dim admin_flag
FoundErr=False 
admin_flag=",14,"
If Not Dvbbs.Master Or InStr(","&Session("flag")&",",admin_flag)=0 Then
	Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=../admin_login.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	Dvbbs_Error()
Else
	Main()
	If FoundErr Then Dvbbs_Error()
	Footer()
End if

'显示内容
'用户名、职务、所管理版面，帖子总数，精华、被删数，登陆次数，最后登陆时间和IP，管理日志（所有日志操作次数和30天内操作次数，点击数字可查看其所管理的日志列表），30天内发贴情况
'评定指标
'列出所管理版面列表，其中包含对应版面该版主
Sub Main()
End Sub
%>