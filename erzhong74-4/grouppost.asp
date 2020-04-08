<!--#include FILE="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="boke/checkinput.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/ubblist.asp"-->
<!--#include file="inc/Group_MainCls.asp"-->
<%
Dim Rs,SQL
Dim Action,PageHtml,ActionName
'Dvbbs.ShowSQL = 1
Action = LCase(Request("action"))
Dvbbs.Loadtemplates("Group")
If Group_MainClass.GroupName="" Then Response.redirect "showerr.asp?ErrCodes="&template.Strings(5)&"&action=OtherErr"
If Group_MainClass.PowerFlag=0 Then Response.redirect "showerr.asp?ErrCodes=<li>抱歉，你不是圈子“"&Group_MainClass.GroupName&"”成员不能进入圈子。&action=OtherErr"
Group_MainClass.Stats = "发表帖子"
Dvbbs.Nav()
Group_MainClass.Head_var 1,"",""

If Action="new" Or Action="revert" Or Action="edit" Then
	Call LoadForm
ElseIf Action="savenew" Or Action="saverevert" Or Action="saveedit" Then
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
	GroupID = Group_MainClass.GroupID
	GroupBoardID = Group_MainClass.GroupBoardID
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
	node.attributes.setNamedItem(XMLDom.createNode(2,"postid","")).text=PostID

	Select Case Action
		Case "edit"
			ActionName = "编辑帖子"
			'AnnounceID=0,GroupID=1,BoardID=2,UserName=3,PostUserID=4,Topic=5,Body=6,RootID=7,IsBest=8
			SQL = "Select AnnounceID,GroupID,BoardID,UserName,PostUserID,Topic,Body,RootID,IsBest From Dv_Group_BBS Where GroupID="&GroupID&" And BoardID="&GroupBoardID&" And AnnounceID="&PostID
			Set Rs=Dvbbs.Execute(SQL)
			If Not Rs.Eof Then
				If Group_MainClass.PowerFlag=0 Or Group_MainClass.PowerFlag>3 Then 
					If Rs(4)<>Dvbbs.UserID Then
						Response.Redirect "showerr.asp?ShowErrType=1&ErrCodes=<li>你没有编辑该帖子的权限。&action=OtherErr"
					End If
				End If
				node.attributes.setNamedItem(XMLDom.createNode(2,"postusername","")).text=Rs(3)&""
				node.attributes.setNamedItem(XMLDom.createNode(2,"topic","")).text=Rs(5)&""
				node.attributes.setNamedItem(XMLDom.createNode(2,"content","")).text=Rs(6)&""
				node.attributes.setNamedItem(XMLDom.createNode(2,"isbest","")).text=Rs(8)&""
				node.attributes.setNamedItem(XMLDom.createNode(2,"actionname","")).text="编辑帖子"
			Else
				Response.Redirect "showerr.asp?ShowErrType=1&ErrCodes=<li>编辑帖子的参数错误,或该帖子已经被删除了。&action=OtherErr"
			End If
			Rs.Close:Set Rs=Nothing
		Case "revert"
			ActionName = "回复帖子"
			If Group_MainClass.PowerFlag=0 Or Group_MainClass.PowerFlag>3 Then 
				SQL = "Select LockTopic From Dv_Group_Topic Where GroupID="&GroupID&" And BoardID="&GroupBoardID&" And TopicID="&TopicID
				Set Rs=Dvbbs.Execute(SQL)
				If Not Rs.Eof Then
					If Rs(0)=1 Then Response.Redirect "showerr.asp?ShowErrType=1&ErrCodes=<li>帖子已经被锁定，不能回复。&action=OtherErr"
				Else
					Response.Redirect "showerr.asp?ShowErrType=1&ErrCodes=<li>编辑帖子的参数错误,或该帖子已经被删除了。&action=OtherErr"
				End IF
			End If
			node.attributes.setNamedItem(XMLDom.createNode(2,"postusername","")).text=Server.HTMLEncode(Dvbbs.MemberName)
			node.attributes.setNamedItem(XMLDom.createNode(2,"topic","")).text=""
			node.attributes.setNamedItem(XMLDom.createNode(2,"content","")).text=""
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
	'XMLStyle.loadxml template.html(10)
	XMLStyle.load Server.MapPath("grouppost.xslt")
	XSLTemplate.stylesheet=XMLStyle
	Set proc = XSLTemplate.createProcessor()
	proc.input = XMLDom
	proc.transform()
	Response.Write  proc.output
	Set XMLDom=Nothing 
	Set proc=Nothing
End Sub

Sub SavePost()
	Dim PostID,ReplayID,LastPost,sucmsg
	Dim Title,GroupBoardID,GroupID,UserName,DateTimeStr
	Dim AnnounceID,RootID,PostContent,IsBest,UbblistBody

	PostID = Dvbbs.CheckNumeric(Request.Form("PostID"))
	Title = Dvbbs.CheckStr(Request.Form("Title"))
	GroupID = Group_MainClass.GroupID
	GroupBoardID = Group_MainClass.GroupBoardID
	UserName = Dvbbs.MemberName
	DateTimeStr = Replace(Replace(CSTR(NOW()+Dvbbs.Forum_Setting(0)/24),"上午",""),"下午","")

	PostContent = Dvbbs.CheckStr(Request.Form("PostContent"))
	IsBest = Dvbbs.CheckNumeric(Request.Form("Best"))
	UbblistBody = PostContent
	UbblistBody = Ubblist(PostContent)
	Select Case Action
		Case "savenew"
			'插入主题
			SQL="insert into Dv_Group_Topic (Title,Boardid,GroupID,PostUsername,PostUserid,DateAndTime,LastPost,LastPostTime,locktopic,istop,HideName) values ('"&Title&"',"&GroupBoardID&","&GroupID&",'"&UserName&"',"&Dvbbs.Userid&",'"&DateTimeStr&"','$$"&DateTimeStr&"$$$$','"&DateTimeStr&"',0,0,0)"
			Dvbbs.Execute(sql)
			Set Rs=Dvbbs.Execute("select Max(topicid) From Dv_Group_Topic Where PostUserid="&Dvbbs.UserID)
			RootID=Rs(0)
			Dvbbs.Execute("Update Dv_Group_Board Set Topicnum=Topicnum+1 Where ID="&GroupBoardID&" And RootID="&GroupID)
			'插入帖子
			SQL="insert into Dv_Group_bbs(GroupID,Boardid,ParentID,username,topic,body,DateAndTime,length,RootID,ip,locktopic,isbest,PostUserID,Ubblist) values("&GroupID&","&GroupBoardID&","&RootID&",'"&username&"','"&Title&"','"&PostContent&"','"&DateTimeStr&"','"&Dvbbs.strlength(PostContent)&"',"&RootID&",'"&Dvbbs.UserTrueIP&"',0,"&IsBest&","&Dvbbs.userid&",'"&UbblistBody&"')"
			Dvbbs.Execute(sql)
			Set Rs=Dvbbs.Execute("select Max(AnnounceID) From Dv_Group_bbs Where PostUserID="&Dvbbs.UserID)
			AnnounceID=Rs(0)
			sucmsg = "<li>成功发布评论"
		Case "saverevert"
			SQL = "Select RootID From Dv_Group_BBS Where GroupID="&GroupID&" And Boardid="&GroupBoardID&" And AnnounceID="&PostID
			Set Rs=Dvbbs.Execute(SQL)
			If Not Rs.Eof Then
				RootID=Rs(0)
				SQL="insert into Dv_Group_bbs(GroupID,Boardid,ParentID,username,topic,body,DateAndTime,length,RootID,ip,locktopic,isbest,PostUserID,Ubblist) values("&GroupID&","&GroupBoardID&","&PostID&",'"&username&"','"&Title&"','"&PostContent&"','"&DateTimeStr&"','"&Dvbbs.strlength(PostContent)&"',"&RootID&",'"&Dvbbs.UserTrueIP&"',0,"&IsBest&","&Dvbbs.userid&",'"&UbblistBody&"')"
				Dvbbs.Execute(SQL)
				Set Rs=Dvbbs.Execute("select Max(AnnounceID) From Dv_Group_bbs Where PostUserID="&Dvbbs.UserID)
				AnnounceID=Rs(0)
			Else
				Response.Redirect "showerr.asp?ShowErrType=1&ErrCodes=<li>您回复帖子参数错误,可能该帖子已经被删除。&action=OtherErr"
			End If
			sucmsg = "<li>成功回复评论"
		Case "saveedit"
			SQL = "Select * From Dv_Group_BBS Where GroupID="&GroupID&" And Boardid="&GroupBoardID&" And AnnounceID="&PostID
			Set Rs=Dvbbs.iCreateObject("ADODB.RecordSet")
			If Not IsObject(Conn) Then ConnectionDatabase
			Rs.Open SQL,Conn,1,3
			If Not Rs.Eof Then
				RootID = Rs("RootID")
				Rs("topic") = Title
				Rs("body") = PostContent
				Rs("length") = Dvbbs.strlength(PostContent)
				Rs("IsBest") = IsBest
				Rs("Ubblist") = UbblistBody
				Rs.Update
			Else
				Response.Redirect "showerr.asp?ShowErrType=1&ErrCodes=<li>您编辑的帖子参数错误,可能该帖子已经被删除。&action=OtherErr"
			End If
			sucmsg = "<li>成功编辑评论"
	End Select
	Rs.Close:Set Rs=Nothing
	LastPost=Replace(username,"$","")&"$"&AnnounceID&"$"&DateTimeStr&"$"&Replace(Title,"$","&#36;")&"$$"&Dvbbs.UserID&"$"&RootID&"$"&GroupBoardID
	Dvbbs.Execute("Update Dv_Group_Board Set LastPost='"&LastPost&"',PostNum=PostNum+1,todaynum=todaynum+1 Where ID="&GroupBoardID&" And RootID="&GroupID)

	sucmsg="<div style=""width:100%;text-align:left;"">"+sucmsg+"<li>"+"<a href=GroupTopicList.asp?GroupID="&Group_MainClass.GroupID&"&groupboardid="&Group_MainClass.GroupBoardID&">返回评论列表</a>"
	sucmsg=sucmsg+"<li>"+"<a href=Groupdispbbs.asp?GroupID="&Group_MainClass.GroupID&"&groupboardid="&Group_MainClass.GroupBoardID&"&id="&RootID&" >返回评论：《"&server.htmlencode(Title)&"》</a></div>"
	Dvbbs.Dvbbs_Suc(sucmsg)
End Sub
%>