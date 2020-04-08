<!--#include FILE="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="inc/ubblist.asp"-->
<!--#include file="Dv_plus/IndivGroup/Dv_IndivGroup_Config.asp"-->
<!--#include file="Dv_plus/IndivGroup/Dv_IndivGroup_MainCls.asp"-->
<%
Dim Rs,SQL
Dim Action,PageHtml,ActionName
'Dvbbs.ShowSQL = 1
Action = LCase(Request("action"))
Dvbbs.Loadtemplates("IndivGroup")

If Dv_IndivGroup_MainClass.ID=0 Or Dv_IndivGroup_MainClass.Name="" Then Response.redirect "showerr.asp?ErrCodes=对不起，你访问的圈子不存在或已经被删除&action=OtherErr"
If Dv_IndivGroup_MainClass.PowerFlag>0 Then
	If Dv_IndivGroup_MainClass.PowerFlag>3 And Dv_IndivGroup_MainClass.GroupStats=3 Then Response.redirect "showerr.asp?ErrCodes=<li>圈子“"&Dv_IndivGroup_MainClass.Name&"”已关闭，只有圈子管理员才能进入。&action=OtherErr"
Else
	Response.redirect "showerr.asp?ErrCodes=<li>抱歉，圈子“"&Dv_IndivGroup_MainClass.Name&"”不公开，只有圈子成员才能进入。&action=OtherErr"
End If
If Dv_IndivGroup_MainClass.PowerFlag > 7 Then Response.redirect "showerr.asp?ErrCodes=<li>抱歉，你不是圈子“"&Dv_IndivGroup_MainClass.Name&"”的成员，不能发贴。&action=OtherErr"
If Dv_IndivGroup_MainClass.BoardStats=0 Then Response.redirect "showerr.asp?ErrCodes=<li>栏目“"&Dv_IndivGroup_MainClass.BoardName&"”已经锁定，不能发贴。&action=OtherErr"

Dv_IndivGroup_MainClass.Stats = "发表帖子"
Dvbbs.Nav()
Dv_IndivGroup_MainClass.Head_var 1,"",""

If Action="new" Or Action="revert" Or Action="edit" Then
	Call LoadForm
ElseIf Action="savenew" Or Action="saverevert" Or Action="saveedit" Then
	Select Case Action
		Case "savenew"
			Action=1
		Case "saverevert"
			Action=2
		Case "saveedit"
			Action=3
	End Select
	Call SavePost
Else
	Call LoadForm
End If

Dvbbs.ActiveOnline
Dvbbs.Footer
Dvbbs.PageEnd()
Sub LoadForm()
	Dim GroupID,GroupBoardID,TopicID,PostID
	Dim UserName,Topic,Content,IsBest,IsTop
	Dim Page,QueryStr,XMLDom,Node,XSLTemplate,XMLStyle,proc
	GroupID = Dv_IndivGroup_MainClass.ID
	GroupBoardID = Dv_IndivGroup_MainClass.BoardID
	TopicID = Dvbbs.CheckNumeric(Request("ID"))
	PostID = Dvbbs.CheckNumeric(Request("PostID"))
	Page = Dvbbs.CheckNumeric(Request("page"))
	QueryStr = "?action=save"&Action&"&groupid="&GroupID&"&groupboardid="&GroupBoardID

	Set XMLDom=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	XMLDom.appendChild(XMLDom.createElement("xml"))
	Set Node=XMLDom.createNode(1,"postdata","")
	node.attributes.setNamedItem(XMLDom.createNode(2,"editmode","")).text="Default"
	node.attributes.setNamedItem(XMLDom.createNode(2,"groupid","")).text=GroupID
	node.attributes.setNamedItem(XMLDom.createNode(2,"boardid","")).text=GroupBoardID
	node.attributes.setNamedItem(XMLDom.createNode(2,"rootid","")).text=TopicID
	node.attributes.setNamedItem(XMLDom.createNode(2,"postid","")).text=PostID

	Select Case Action
		Case "edit"
			ActionName = "编辑帖子"
			'AnnounceID=0,GroupID=1,BoardID=2,UserName=3,PostUserID=4,Topic=5,Body=6,RootID=7,IsBest=8
			SQL = "Select AnnounceID,GroupID,BoardID,UserName,PostUserID,Topic,Body,RootID,IsBest From Dv_Group_BBS Where GroupID="&GroupID&" And BoardID="&GroupBoardID&" And AnnounceID="&PostID
			Set Rs=Dv_IndivGroup_MainClass.Execute(SQL)
			If Not Rs.Eof Then
				If Dv_IndivGroup_MainClass.PowerFlag=0 Or Dv_IndivGroup_MainClass.PowerFlag>3 Then 
					If Rs(4)<>Dvbbs.UserID Then
						Response.Redirect "showerr.asp?ErrCodes=<li>你没有编辑该帖子的权限。&action=OtherErr"
					End If
				End If
				node.attributes.setNamedItem(XMLDom.createNode(2,"postusername","")).text=Rs(3)&""
				node.attributes.setNamedItem(XMLDom.createNode(2,"topic","")).text=Rs(5)&""
				node.attributes.setNamedItem(XMLDom.createNode(2,"content","")).text=Ubb2Html(Rs(6))&""
				node.attributes.setNamedItem(XMLDom.createNode(2,"isbest","")).text=Rs(8)&""
				node.attributes.setNamedItem(XMLDom.createNode(2,"actionname","")).text="编辑帖子"
			Else
				Response.Redirect "showerr.asp?ErrCodes=<li>编辑帖子的参数错误,或该帖子已经被删除了。&action=OtherErr"
			End If
			Rs.Close:Set Rs=Nothing
		Case "revert"
			ActionName = "回复帖子"
			If Dv_IndivGroup_MainClass.PowerFlag=0 Or Dv_IndivGroup_MainClass.PowerFlag>3 Then 
				SQL = "Select LockTopic From Dv_Group_Topic Where GroupID="&GroupID&" And BoardID="&GroupBoardID&" And TopicID="&TopicID
				Set Rs=Dv_IndivGroup_MainClass.Execute(SQL)
				If Not Rs.Eof Then
					If Rs(0)=1 Then Response.Redirect "showerr.asp?ErrCodes=<li>帖子已经被锁定，不能回复。&action=OtherErr"
				Else
					Response.Redirect "showerr.asp?ErrCodes=<li>帖子的参数错误,或该帖子已经被删除了。&action=OtherErr"
				End IF
				Rs.Close
			End If
			Content=""
			If Request("Reply")="true" Then
				SQL = "Select UserName,Body,DateAndTime From Dv_Group_BBS Where GroupID="&GroupID&" And BoardID="&GroupBoardID&" And AnnounceID="&PostID
				Set Rs=Dv_IndivGroup_MainClass.Execute(SQL)
				If Not Rs.Eof Then
					If Rs(0)<>"" Then
						Content = reubbcode(Rs(1))
						Content = Ubb2Html(Content)
						Content = "<DIV class=quote><B>以下是引用<i>"&Rs(0)&"</i>在"&Rs(2)&"的发言：</B><br>"& Content & "</DIV><p>"
						'Content = Server.HtmlEncode(Content)
					End If
				End If
			End If
			node.attributes.setNamedItem(XMLDom.createNode(2,"postusername","")).text=Server.HTMLEncode(Dvbbs.MemberName)
			node.attributes.setNamedItem(XMLDom.createNode(2,"topic","")).text=""
			node.attributes.setNamedItem(XMLDom.createNode(2,"content","")).text=Content
			node.attributes.setNamedItem(XMLDom.createNode(2,"isbest","")).text=0
			node.attributes.setNamedItem(XMLDom.createNode(2,"actionname","")).text="回复帖子"
		Case Else
			node.attributes.setNamedItem(XMLDom.createNode(2,"postusername","")).text=Server.HTMLEncode(Dvbbs.MemberName)
			node.attributes.setNamedItem(XMLDom.createNode(2,"topic","")).text=""
			node.attributes.setNamedItem(XMLDom.createNode(2,"content","")).text=""
			node.attributes.setNamedItem(XMLDom.createNode(2,"isbest","")).text=0
			node.attributes.setNamedItem(XMLDom.createNode(2,"actionname","")).text="发表新帖子"
	End Select
	QueryStr = QueryStr&"&PostID="&PostID&"&Page="&Page
	node.attributes.setNamedItem(XMLDom.createNode(2,"action","")).text=QueryStr
	XMLDom.documentElement.appendChild(Node)

	Set XSLTemplate=Dvbbs.iCreateObject("Msxml2.XSLTemplate" & MsxmlVersion)
	Set XMLStyle=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	XMLStyle.loadxml template.html(4)
	'XMLStyle.load Server.MapPath("IndivGroup/Skin/post.xslt")
	XSLTemplate.stylesheet=XMLStyle
	Set proc = XSLTemplate.createProcessor()
	proc.input = XMLDom
	proc.transform()
	Response.Write  proc.output
	Set XMLDom=Nothing 
	Set proc=Nothing
End Sub

Sub SavePost()
	Dim TopicID,PostID,ReplayID,LastPost,sucmsg,QueryString
	Dim Title,GroupBoardID,GroupID,UserName,DateTimeStr
	Dim AnnounceID,RootID,PostContent,IsBest,UbblistBody
	Dim IsAudit,LockTopic	'add hantg.20071204

	TopicID = Dvbbs.CheckNumeric(Request.Form("rootid"))
	PostID = Dvbbs.CheckNumeric(Request.Form("PostID"))
	Title = Replace(Dvbbs.CheckStr(Trim(Request.Form("Title"))),"","")
	GroupID = Dv_IndivGroup_MainClass.ID
	GroupBoardID = Dv_IndivGroup_MainClass.BoardID
	UserName = Dvbbs.MemberName
	DateTimeStr = Replace(Replace(CSTR(NOW()+Dvbbs.Forum_Setting(0)/24),"上午",""),"下午","")

	PostContent = Dvbbs.CheckStr(Request.Form("PostContent"))
	IsBest = Dvbbs.CheckNumeric(Request.Form("Best"))
	UbblistBody = PostContent
	UbblistBody = Ubblist(PostContent)
	QueryString = ""

	'--start add hantg.20071204-------------------------------------
	'发帖审核处理
	LockTopic=0
	If Dvbbs.CheckNumeric(Dvbbs.Forum_Setting(105))=1 Then
		If Dvbbs.Forum_Setting(106)="" Or Dvbbs.Forum_Setting(106)="0" Or InStr("||"&(Dvbbs.Forum_Setting(106)&"||"),"||"&(Dv_IndivGroup_MainClass.Name&"||"))>0 Then
			IsAudit=0
			If AccessPost =1 Then
				IsAudit=1
			End If
			locktopic=0
			If IsAudit=1 Then
				LockTopic=GroupBoardID
				GroupBoardID=-2
			End If
		End If
	End If
	'--end add hantg.20071204-------------------------------------
	
	Select Case Action
		Case "1"
			'插入主题
			If Title="" Then Response.redirect "showerr.asp?ShowErrType=1&ErrCodes=<li>您忘记填写标题<BR>2秒后自动返回上一页面。&action=OtherErr&"
			SQL="insert into Dv_Group_Topic (Title,Boardid,GroupID,PostUsername,PostUserid,DateAndTime,LastPost,LastPostTime,locktopic,istop) values ('"&Title&"',"&GroupBoardID&","&GroupID&",'"&UserName&"',"&Dvbbs.Userid&",'"&DateTimeStr&"','$$"&DateTimeStr&"$$$$','"&DateTimeStr&"',"&LockTopic&",0)"
			Dv_IndivGroup_MainClass.Execute(sql)
			Set Rs=Dv_IndivGroup_MainClass.Execute("select Max(topicid) From Dv_Group_Topic Where PostUserid="&Dvbbs.UserID)
			RootID=Rs(0)
			'插入帖子
			SQL="insert into Dv_Group_bbs(GroupID,Boardid,ParentID,username,topic,body,DateAndTime,length,RootID,ip,locktopic,isbest,PostUserID,Ubblist) values("&GroupID&","&GroupBoardID&",0,'"&username&"','"&Title&"','"&PostContent&"','"&DateTimeStr&"','"&Dvbbs.strlength(PostContent)&"',"&RootID&",'"&Dvbbs.UserTrueIP&"',"&LockTopic&","&IsBest&","&Dvbbs.userid&",'"&UbblistBody&"')"
			Dv_IndivGroup_MainClass.Execute(sql)
			Set Rs=Dv_IndivGroup_MainClass.Execute("select Max(AnnounceID) From Dv_Group_bbs Where PostUserID="&Dvbbs.UserID)
			AnnounceID=Rs(0)

			QueryString = ",TopicNum=TopicNum+1"
			If IsAudit=1 Then
				sucmsg = "<li>由于您发表的评论含敏感内容或尚未达到豁免审核的条件，您的贴子需要管理员审核过才可以见到。</li>"
			Else
				sucmsg = "<li>成功发布评论</li>"
			End If
		Case "2"
			'SQL = "Select TopicID From Dv_Group_Topic Where GroupID="&GroupID&" And TopicID="&TopicID
			SQL = "Select AnnounceID From Dv_Group_BBS Where GroupID="&GroupID&" And ParentID=0 And RootID="&TopicID
			Set Rs=Dv_IndivGroup_MainClass.Execute(SQL)
			If Rs.Eof Then Response.Redirect "showerr.asp?ShowErrType=1&ErrCodes=<li>您回复帖子参数错误,可能该帖子已经被删除。&action=OtherErr"
			RootID=TopicID : PostID=Rs(0)

			SQL="insert into Dv_Group_bbs(GroupID,Boardid,ParentID,username,topic,body,DateAndTime,length,RootID,ip,locktopic,isbest,PostUserID,Ubblist) values("&GroupID&","&GroupBoardID&","&PostID&",'"&username&"','"&Title&"','"&PostContent&"','"&DateTimeStr&"','"&Dvbbs.strlength(PostContent)&"',"&RootID&",'"&Dvbbs.UserTrueIP&"',"&LockTopic&","&IsBest&","&Dvbbs.userid&",'"&UbblistBody&"')"
			Dv_IndivGroup_MainClass.Execute(SQL)
			Set Rs=Dv_IndivGroup_MainClass.Execute("select Max(AnnounceID) From Dv_Group_bbs Where PostUserID="&Dvbbs.UserID)
			AnnounceID=Rs(0)
	
			If Title = "" Then Title = PostContent
			Title = Dv_IndivGroup_MainClass.cutStr(Title,25)
			LastPost=Replace(username,"$","") & "$" & AnnounceID & "$" & DateTimeStr & "$" & Replace(Title,"$","&#36;") & "$$" & Dvbbs.UserID & "$" & RootID & "$" & GroupBoardID

			Set Rs = Dv_IndivGroup_MainClass.Execute("Select Title,LastPost From Dv_Group_Topic Where TopicID="&RootID&" And GroupID="&Dv_IndivGroup_MainClass.ID)
			If Not Rs.Eof Then
				Title = Rs(0)
				Dv_IndivGroup_MainClass.Execute("Update Dv_Group_Topic Set Child=Child+1,LastPost='"&LastPost&"',LastPostTime="&SqlNowString&" Where TopicID="&RootID&" And GroupID="&Dv_IndivGroup_MainClass.ID)
			End If
			If IsAudit=1 Then
				sucmsg = "<li>由于您回复的评论含敏感内容或尚未达到豁免审核的条件，您的贴子需要管理员审核过才可以见到。</li>"
			Else
				sucmsg = "<li>成功回复评论</li>"
			End If
		Case "3"
			Dv_IndivGroup_MainClass.ShowSQL = 1
			SQL = "Select * From Dv_Group_BBS Where AnnounceID="&PostID
			Set Rs=Dvbbs.iCreateObject("ADODB.RecordSet")
			If Not IsObject(Conn) Then ConnectionDatabase
			Rs.Open SQL,Conn,1,3
			If Not Rs.Eof Then
				If Rs("postuserid") <> Dvbbs.Userid And Not (Dvbbs.Master Or Dvbbs.Superboardmaster) Then 
					Response.Redirect "showerr.asp?ShowErrType=1&ErrCodes=<li>您没有修改别人帖子的权限。&action=OtherErr"
				End If 
				If Rs("ParentID")=0 Then 
					If GroupBoardID=-2 Then
						SQL = ",BoardID="&GroupBoardID&",LockTopic="&LockTopic
					Else
						SQL = ""
					End If
					SQL = "Update Dv_Group_Topic Set Title='"&Title&"'"& SQL &" Where GroupID="&GroupID&" And TopicID="&Rs("RootID")
					Dv_IndivGroup_MainClass.Execute(SQL)
				End If
				RootID = Rs("RootID")
				Rs("topic") = Title
				Rs("body") = PostContent
				If GroupBoardID=-2 Then
					Rs("BoardID")=Dvbbs.BoardID
					Rs("LockTopic")=LockTopic
				End If
				Rs("length") = Dvbbs.strlength(PostContent)
				Rs("Ubblist") = UbblistBody
				Rs.Update
			Else
				Response.Redirect "showerr.asp?ShowErrType=1&ErrCodes=<li>您编辑的帖子参数错误,可能该帖子已经被删除。&action=OtherErr"
			End If
			If IsAudit=1 Then
				sucmsg = "<li>由于您编辑的评论含敏感内容或尚未达到豁免审核的条件，您的贴子需要管理员审核过才可以见到。</li>"
			Else
				sucmsg = "<li>成功编辑评论</li>"
			End If
	End Select
	Rs.Close:Set Rs=Nothing
	If Action<>"saveedit" And IsAudit=0 Then
		LastPost=Replace(username,"$","") &"$"& AnnounceID &"$"& DateTimeStr &"$"& Dv_IndivGroup_MainClass.cutStr(Replace(Title,"$","&#36;"),25) &"$$"& Dvbbs.UserID &"$"& RootID &"$"& GroupBoardID
		Dv_IndivGroup_MainClass.Execute("Update Dv_Group_Board Set LastPost='"&LastPost&"',PostNum=PostNum+1,todaynum=todaynum+1"&QueryString&" Where ID="&GroupBoardID&" And RootID="&GroupID)
		Dv_IndivGroup_MainClass.Execute("Update Dv_GroupName Set TodayNum=TodayNum+1,PostNum=PostNum+1"&QueryString&" Where ID="&Dv_IndivGroup_MainClass.ID)
	End If

	sucmsg = "<div style=""width:100%;text-align:left;"">"& sucmsg &"<li><a href=IndivGroup_index.asp?GroupID="& Dv_IndivGroup_MainClass.ID &"&groupboardid="& Dv_IndivGroup_MainClass.BoardID &">返回评论列表</a></li>"
	If IsAudit=0 Then
		sucmsg = sucmsg & "<li><a href=IndivGroup_Dispbbs.asp?GroupID="& Dv_IndivGroup_MainClass.ID &"&groupboardid="& Dv_IndivGroup_MainClass.BoardID &"&id="& RootID &" >返回评论：《"& server.htmlencode(Title) &"》</a></li>"
	End If
	sucmsg = sucmsg &"</div>"

	Dvbbs.Dvbbs_Suc(sucmsg)
End Sub

'--start add hantg.20071204-------------------------------------
'判断是否需要审核
Function AccessPost()
	Dim Dom,node,ischeckcontent
	If Dvbbs.UserID>0 Then
		If Dvbbs.GroupSetting(64)="1" Then
			AccessPost=0
			Exit Function
		End If
	End If
'	If Cint(Dvbbs.Board_Setting(3))=0 Then
'		AccessPost=0
'		Exit Function
'	End If
	ischeckcontent=0
	Set Dom=Application(Dvbbs.CacheName & "_accesstopic")
	'检查审核是否有打开设置
	Set Node=Dom.documentElement.selectSingleNode("setting[@use=1]")
	If Node is Nothing Then
		AccessPost=0
		Exit Function
	End If
	'检查是否属于需要审核的用户组
	Set Node=Dom.documentElement.selectSingleNode("setting[@use=1]/checkuser[@use=1 and @usergroupid="&dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usergroupid").text&"]")
	Dim CheckUserGroup
	If Node is Nothing Then
		Set Node=Dom.documentElement.selectSingleNode("setting[@use=1]/checkuser[@use=1 and @usergroupid=4]")
		If Not (Node Is Nothing) Then
			Response.Charset = "GB2312"
			Dim ParentGID
			If Not TypeName(Application(Dvbbs.CacheName &"_groupsetting"))="DOMDocument" Then LoadGroupSetting()
			ParentGID=Application(Dvbbs.CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usergroupid").text &"']/@parentgid").text
			If Not ParentGID=3 Then AccessPost=0:Exit Function
			CheckUserGroup=4
		Else
			AccessPost=0
			Exit Function
		End If
	Else
		CheckUserGroup=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usergroupid").text
	End If

	'用户被豁免审核判断
	Set Node=Dom.documentElement.selectSingleNode("setting[@use=1]/nocheck/username[.='"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@username").text&"']")
	If Not Node is Nothing Then
		AccessPost=0
		Exit Function
	End If
	Rem action=5 主题 actiom=6 回复 action=7 投票 action=8 编辑
	Dim tmpstr,laction,regdate,delPercent
	 Select Case action
		Case "1"
			tmpstr="setting[@use=1]/check[@new=1]"
			laction="new"
		Case "2"
			tmpstr="setting[@use=1]/check[@re=1]"
			laction="re"
		Case "3"
			tmpstr="setting[@use=1]/check[@edit=1]"
			laction="edit"
	 End Select
	If Dom.documentElement.selectSingleNode(tmpstr) Is Nothing Then
		AccessPost=0
		Exit Function
	End If 
	regdate=Datediff("d",Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@joindate").text,Now())
	If CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text) > 0 Or CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userdel").text) < 0 Then
		delPercent=CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text)-CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userdel").text)
		delPercent=(0-CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userdel").text))/delPercent
	Else
		delPercent=0
	End If
	delPercent=FormatPercent(delPercent)
	For Each node in Dom.documentElement.selectNodes("setting[@use=1 and check/@"&laction&"=1]/checkuser[@use=1 and @usergroupid="&CheckUserGroup&"]")
		'主题数判断
		If Node.selectSingleNode("usertopic/@use").text="1" Then
			If CCur(Node.selectSingleNode("usertopic/@value").text) > CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usertopic").text) Then
				AccessPost=1
				Exit For
			End If
		End If
		'回贴数
		If Node.selectSingleNode("userpost/@use").text="1" Then
			If CCur(Node.selectSingleNode("userpost/@value").text) > CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text) Then
				AccessPost=1
				Exit For
			End If
		End If
		'注册天数的比较
		If Node.selectSingleNode("regdate/@use").text="1" Then
			If CCur(Node.selectSingleNode("regdate/@value").text) > regdate Then
				AccessPost=1
				Exit For
			End If
		End If
		'删贴百分比比较
		If Node.selectSingleNode("userdel/@use").text="1" Then
			If FormatNumber(Node.selectSingleNode("userdel/@value").text) <  FormatNumber(delPercent) Then
				AccessPost=1
				Exit For
			End If
		End If
		'对被屏蔽用户是否审核的判断
		If Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@lockuser").text="2" And Node.selectSingleNode("lockuser/@use").text="1" Then
			AccessPost=1
			Exit For
		End If
		If Node.selectSingleNode("checkcontent/@use").text="1" Then
			Ischeckcontent=1
		End If
	Next
	If AccessPost=1 Then
		Exit Function
	End If
	If ischeckcontent=0 Then
		AccessPost=0
	Else
		Set Node=Dom.documentElement.selectNodes("setting[@use=1]/checkcontent")
		AccessPost=checkcontent(node)
	End If
End Function

Function checkcontent(node1)
	Dim re,s,node,tmpnode
	checkcontent=0
	s=Request("Title")&Request("PostContent")
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	For Each Node in node1
		'含图片检查
		If Node.selectSingleNode("checkpic/@use").text="1" Then
			re.Pattern="(\[IMG\]|<img)"
			If re.Test(s) Then
				checkcontent=1
				Exit For
			End If
		End If
		'含链接检查
		If Node.selectSingleNode("checklink/@use").text="1" Then
			re.Pattern="(\[url\]|<a |http://|https://)"
			If re.Test(s) Then
				checkcontent=1
				Exit For
			End If
		End If
		'含flash检查
		If Node.selectSingleNode("checkflash/@use").text="1" Then
			re.Pattern="(\[flash|<object )"
			If re.Test(s) Then
				checkcontent=1
				Exit For
			End If
		End If
		'含媒体播放检查
		If Node.selectSingleNode("checkmp/@use").text="1" Then
			re.Pattern="(\[mp)"
			If re.Test(s) Then
				checkcontent=1
				Exit For
			End If
		End If
		'含媒rm检查
		If Node.selectSingleNode("checkrm/@use").text="1" Then
			re.Pattern="(\[rm)"
			If re.Test(s) Then
				checkcontent=1
				Exit For
			End If
		End If
		For Each tmpnode in Node.selectNodes("checkword")
			re.Pattern = Trim(tmpnode.getAttribute("content"))
			If re.Test(s) Then
				checkcontent=1
				Exit For
			End If
		Next
	Next
	set re=Nothing
End Function
'--start add hantg.20071204-------------------------------------

Function reUBBCode(strContent)
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="<\/div>"
	strContent=re.Replace(strContent,Chr(2))
	re.Pattern="<div class=quote><b>以下是引用"
	strContent=re.Replace(strContent,Chr(1))
	re.Pattern="<div class=quote>([^\x01\x02]*)\x02"
	Do While re.Test(strContent)
		strContent=re.Replace(strContent,"[quote]$1[/quote]")
	Loop
	re.Pattern="\x01[^\x02]*\x02"
	strContent=re.Replace(strContent,"")
	re.Pattern="\[quote\]((?:.|\n)*?)\[\/quote\]"
	Do While re.Test(strContent)
		strContent=re.Replace(strContent,"<div class=quote>$1</div>")
	Loop
	re.Pattern="\x02"
	strContent=re.Replace(strContent,"</div>")
	re.Pattern="<div align=right><font color=#000066>(?:.|\n)*?<\/font><\/div>"
	strContent=re.Replace(strContent,"")
	re.Pattern="\[align=right\]\[color=#000066\](?:.|\n)*?\[\/color\]\[\/align\]"
	strContent=re.Replace(strContent,"")
'	re.Pattern="(\[QUOTE\])(.|\n)*?(\[\/QUOTE\])"
'	strContent=re.Replace(strContent,"$2")
	re.Pattern="\[point=*([0-9]*)\](?:.|\n)*?\[\/point\]"
	strContent=re.Replace(strContent,"")
	re.Pattern="\[post=*([0-9]*)\](?:.|\n)*?\[\/post\]"
	strContent=re.Replace(strContent,"")
	re.Pattern="\[power=*([0-9]*)\](?:.|\n)*?\[\/power\]"
	strContent=re.Replace(strContent,"")
	re.Pattern="\[usercp=*([0-9]*)\](?:.|\n)*?\[\/usercp\]"
	strContent=re.Replace(strContent,"")
	re.Pattern="\[money=*([0-9]*)\](?:.|\n)*?\[\/money\]"
	strContent=re.Replace(strContent,"")
	re.Pattern="\[replyview\](?:.|\n)*?\[\/replyview\]"
	strContent=re.Replace(strContent,"")
	re.Pattern="\[usemoney=*([0-9]*)\](?:.|\n)*?\[\/usemoney\]"
	strContent=re.Replace(strContent,"")
	re.Pattern="\[UserName=(.[^\[]*)\](?:.|\n)*?\[\/UserName\]"
	strContent=re.Replace(strContent,"")
	re.Pattern="  "
	strContent=re.Replace(strContent,"&nbsp;&nbsp;")
	re.Pattern="<I><\/I>"
	strContent=re.Replace(strContent,"")
	set re=Nothing
	reUBBCode=strContent
End Function

'编辑时用（对旧数据兼容）
Function Ubb2Html(str)
	If Str<>"" And Not IsNull(Str) Then
		Dim re
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="(>)("&vbNewLine&")(<)"
		Str=re.Replace(Str,"$1$3")
		re.Pattern="(>)("&vbNewLine&vbNewLine&")(<)"
		Str=re.Replace(Str,"$1$3")
		re.Pattern=vbNewLine
		Str=re.Replace(Str,"<br>")
		re.Pattern="  "
		Str=re.Replace(Str,"&nbsp;&nbsp;")
		re.Pattern="	"
		Str=re.Replace(Str,"&nbsp;")
		re.Pattern="<I><\/I>"
		Str=re.Replace(Str,"")
		re.Pattern="<(\w+)(?:&nbsp;)+([^>]*)>"
		Str = re.Replace(Str,"<$1 $2>")
		If Request("reply")="true" Then
			re.Pattern="<DIV class=quote><b>以下是引用(?:.|\n)*<\/div>"
			Str=re.Replace(Str,"")
			re.Pattern="<div class=""quote""><b>以下是引用(?:.|\n)*<\/div>"
			Str=re.Replace(Str,"")
			re.Pattern="\[quote\]<b>以下是引用(?:.|\n)*\[\/quote\]"
			Str=re.Replace(Str,"")
			re.Pattern="\[quote\]\[b\]以下是引用(?:.|\n)*\[\/quote\]"
			Str=re.Replace(Str,"")
		End If
		Set Re=Nothing 
		Ubb2Html = Str
	Else
		Ubb2Html = ""
	End If
End Function
%>