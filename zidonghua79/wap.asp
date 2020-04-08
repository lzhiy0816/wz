<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/chan_const.asp"-->
<!--#include file="inc/md5.asp"-->
<%
Dvbbs.LoadTemplates("")

If Request("Action")="readme" Then
	Main()
ElseIf Request("Action")="wappush" Then
	WapPush()
Else
	MyRedirect()
End If

Sub Main()
	Dvbbs.stats="WAP移动论坛说明"
	Dvbbs.Nav()
	Dvbbs.Head_var 0,0,"WAP移动论坛说明","wap.asp?Action=readme"
	Dvbbs.Showerr()
	Get_ChallengeWord
	Dim Rs
	Set Rs=Dvbbs.Execute("Select Top 1 * From Dv_ChallengeInfo")
	If Rs.Eof And Rs.Bof Then
		Response.redirect "showerr.asp?ErrCodes=<li>错误的数据，请联系动网论坛官方解决。&action=OtherErr"
		Exit Sub
	End If
%>
<table border="0" cellpadding=3 cellspacing=1 align=center class=Tableborder1>
	<tr>
	<th height=23>WAP移动论坛，把论坛动起来，随时随地的沟通</th></tr>
	<FORM METHOD=POST ACTION="http://bbs.ray5198.com/bbsapp/midwappush.jsp">
	<input type=hidden name="mouseid" value="<%=Rs("D_UserName")%>">
	<input type=hidden name="forumname" value="<%=Dvbbs.Forum_Info(0)%>">
	<input type=hidden name="forumid" value="<%=Rs("D_ForumID")%>">
	<input type=hidden name="forumurl" value="<%=Dvbbs.Get_ScriptNameUrl()%>wap.asp">
	<input type=hidden name="backurl" value="<%=Dvbbs.Get_ScriptNameUrl()%>wap.asp?Action=wappush">
	<input type=hidden name="type" value="<%
	If Request("t")<>"" Then
		Response.Write "1"
	Else
		If Request("s")<>"" Then
			Response.Write Request("s")
		Else
			Response.Write "0"
		End If
	End If
	%>">
	<input type=hidden name="seqno" value="<%=Session("challengeWord")%>">
	<input type=hidden name="imgurl" value="<%
	If InStr(Lcase(Request("t")),"showimg") Then
		Response.Write Request("t") & "&filename="&Request("filename")
	Else
		Response.Write Request("t")
	End If
	%>">
	<input type=hidden name="provider" value="1">
	<tr>
	<td height=23 class=Tablebody1 style="line-height: 18px">
	<BLOCKQUOTE><BR>
	<center><%=Dvbbs.Forum_Info(0)%> 目前已开通手机移动论坛（移动博客）功能，让您和您的朋友们情景互动、随时随地的沟通</center>
	<hr size=1 color=#000000>
	<B>手机访问地址</B>：<font color=red><%=Dvbbs.Get_ScriptNameUrl()%>wap.asp</font>或<font color=red><%=Replace(Lcase(Dvbbs.Get_ScriptNameUrl()),"http://","wsp://")%>wap.asp</font><BR>
	说明：只要您开通了移动GPRS或联通CDMA的数据服务，在支持WAP的手机上即可通过此网址访问 <%=Dvbbs.Forum_Info(0)%><BR><BR>
	<%If Request("t")="" Then%><B>WAP PUSH下发<font color=blue>免费</font>论坛地址</B><%Else%>WAP PUSH下发免费论坛图片地址（可通过彩信转发至手机）<%End If%>，请输入您的手机号码：<input type=text size=20 name="mobile" value="13">&nbsp;<input type=submit name=submit value="<%
	If Request("t")<>"" Then
		Response.Write "发送图片到手机"
	Else
		If Request("s")<>"" Then
			Response.Write "免费发送免费图铃和游戏"
		Else
			Response.Write "免费发送论坛地址"
		End If
	End If%>">
	<BR><BR>
	<%
	If Request("t")<>"" Then
		If InStr(Lcase(Request("t")),"showimg") Then
			Response.Write "<div align=center><img src="&Request("t")&"&filename="&Request("filename")&" onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></div><BR><BR>"
		Else
			Response.Write "<div align=center><img src="&Request("t")&" onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></div><BR><BR>"
		End If
	End If
	%>
	<B>主要手机论坛应用范围</B>：<BR>
	&nbsp;&nbsp;&nbsp;&nbsp;手机也“博客”－通过手机您可以随时随地的浏览您在网络论坛上的活动信息，了解您发表的帖子（日记）的阅读、回复情况。您可以随时随地的将您的问题、生活趣事、趣味照片（手机拍照）等等信息通过手机发布，无线沟通，共享快乐。<BR>
	&nbsp;&nbsp;&nbsp;&nbsp;信息的“获取”－重要资料不能随身携带，或是忘记携带怎么办？现在只要您在 <%=Dvbbs.Forum_Info(0)%> 上保存了相关资料，通过手机您就可随时随地的重新翻阅或查找相关的资料信息。<BR>
	&nbsp;&nbsp;&nbsp;&nbsp;问题的“解决”－碰到紧急的问题自己不能解决？没有传统网络设备可供上网查询？现在您只要通过手机在热门的相关论坛版面提出您的问题，相信很快就能得到热心网友的回答。<BR><BR>
	<B>免费图片铃声：</B><BR>
	这里为大家提供大量的免费图片铃声，只要不是明确提示了收费的项目，全部都是免费的。但是要提醒大家，用手机上网中国移动是要收取GPRS流量费的，参考价是每K3分钱，当然各地移动会有不同的套餐或是优惠政策，只能麻烦大家咨询当地的1860移动客服台去了解细节了，<a href="?Action=readme&s=2"><font color=blue>免费图片铃声</font></a><BR><BR>
	<B>免费游戏下载：</B><BR>
	这里为大家提供大量的免费手机游戏，只要不是明确提示了收费的项目，全部都是免费的。但是要提醒大家，用手机上网中国移动是要收取GPRS流量费的，参考价是每K3分钱，当然各地移动会有不同的套餐或是优惠政策，只能麻烦大家咨询当地的1860移动客服台去了解细节了，<a href="?Action=readme&s=2"><font color=blue>免费游戏下载</font></a><BR><BR>	<B>手机上网资费说明</B>：移动GPRS或联通CDMA的数据服务包月或流量费用（参考价是每K3分钱，当然各地移动会有不同的套餐或是优惠政策，只能麻烦大家咨询当地的1860移动客服台去了解细节了），无需任何额外的费用支出<BR><BR>
	<B>下载免费资源的注意事项？</B><BR>
	不管是下载图片、铃声还是游戏，都要在事先确认你的机型是不是和提供的文件匹配，因为手机的机型不同技术标准也不同，所以各种程序的通用性很差，一定要对号使用，不然不仅下载的文件不能使用，还要耗费你下载的流量费，真的很心疼啊。
	</BLOCKQUOTE>
	</td>
	</tr>
	</FORM>
</table>
<%
	Rs.Close
	Set Rs=Nothing
	Dvbbs.ActiveOnline()
	Dvbbs.Footer()
End Sub

Sub WapPush()
	Dvbbs.stats="WAP PUSH下发"
	Dvbbs.Nav()
	Dvbbs.Head_var 0,0,"WAP PUSH下发","wap.asp?Action=readme"
	Dvbbs.Showerr()
%>
<table border="0" cellpadding=3 cellspacing=1 align=center class=Tableborder1 style="width:600;">
	<tr>
	<th height=23>WAP移动论坛，把论坛动起来，随时随地的沟通</th></tr>
	<tr>
	<td height=23 class=Tablebody1 style="line-height: 18px">
<%
	If Request("errorcode")="1" Then
		Response.Write "<li>您成功的将相关的论坛地址用WAP PUSH下发到指定的手机。"
	Else
		Response.Write "<li>错误，WAP PUSH下发失败，"&request("errormsg")&"。"
	End If
%>
	</td>
	</tr>
</table>
<%
	Dvbbs.ActiveOnline()
	Dvbbs.Footer()
End Sub

Sub MyRedirect()
	Dim ServerUrl,ScriptName
	Dim Tmpstr
	Tmpstr = Request.ServerVariables("PATH_INFO")
	Tmpstr = Split(Tmpstr,"/")
	ScriptName = Lcase(Tmpstr(UBound(Tmpstr)))
	
	If request.servervariables("SERVER_PORT")="80" Then
		ServerUrl="http://" & request.servervariables("server_name")&replace(lcase(request.servervariables("script_name")),ScriptName,"")
	Else
		ServerUrl="http://" & request.servervariables("server_name")&":"&request.servervariables("SERVER_PORT")&replace(lcase(request.servervariables("script_name")),ScriptName,"")
	End If
	'Response.Write "http://bbs.ray5198.com/wap/first.jsp?url="&ServerUrl
	'Response.End
	Response.Redirect "http://bbs.ray5198.com/wap/first.jsp?url="&ServerUrl
End Sub
%>