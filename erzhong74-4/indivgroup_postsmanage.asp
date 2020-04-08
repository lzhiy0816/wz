<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="Dv_plus/IndivGroup/Dv_IndivGroup_Config.asp"-->
<!--#include file="Dv_plus/IndivGroup/Dv_IndivGroup_MainCls.asp"-->
<%
Dim Rs,SQL,i,TempStr
Dim ID,sucmsg,ErrCodes
Dvbbs.LoadTemplates("indivgroup")

If Dv_IndivGroup_MainClass.ID=0 Or Dv_IndivGroup_MainClass.Name="" Then Response.redirect "showerr.asp?ErrCodes=对不起，你访问的圈子不存在或已经被删除1&action=OtherErr"
If Dv_IndivGroup_MainClass.PowerFlag=0 Or Dv_IndivGroup_MainClass.PowerFlag>3 Then Response.redirect "showerr.asp?ErrCodes=<li>只有该圈子管理员或论坛管理员才能进入圈子管理页面。&action=OtherErr"
If Request.form("topicid")="" Then Response.redirect "showerr.asp?ErrCodes=<li>参数错误，主题ID为空。&action=OtherErr"

Dv_IndivGroup_MainClass.stats="帖子批量管理"
Dvbbs.Nav()
Dv_IndivGroup_MainClass.Head_var 1,"",""
Select Case LCase(Request("action"))
	Case "lock"
		Topic_Batch_Lock
	Case "dele"
		Topic_Batch_Delete
	Case "isbest"
		Topic_Batch_Isbest
	Case Else
		Dvbbs.AddErrCode(35)
End Select
Dvbbs.Footer()
Dvbbs.PageEnd()
'批量锁定
Function Topic_Batch_Lock
	For i=1 To Request.Form("topicid").Count
		ID=Dvbbs.CheckNumeric(Request.Form("topicid")(i))
		Dvbbs.Execute("Update Dv_Group_Topic Set LockTopic=1 Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&Dv_IndivGroup_MainClass.BoardID&" And TopicID=" & ID)
	Next
	sucmsg="<li>帖子批量锁定成功。"
	Dvbbs.Dvbbs_suc(sucmsg)
End Function

'批量删除
Function Topic_Batch_Delete
	Dim toprs,iBoardTopStr
	Dim ChildNum,istop,Forum_AllTopNum,j,AllTopNum,Board_TopNum,TopNum,BoardTopStr
	ChildNum = 0
	For i=1 To Request.Form("topicid").Count
		ID=Dvbbs.CheckNumeric(Request.Form("topicid")(i))
		Set Rs=Dvbbs.Execute("Select Child From Dv_Group_Topic Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&Dv_IndivGroup_MainClass.BoardID&" And TopicID="&ID)
		If Not Rs.Eof Then
			ChildNum=ChildNum+Rs(0)
			'更新帖子数据
			Dvbbs.Execute("Update Dv_Group_BBS Set Locktopic=BoardID,BoardID=-1,IsBest=0 Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&Dv_IndivGroup_MainClass.BoardID&" And RootID="&ID)
			Dvbbs.Execute("Update Dv_Group_Topic Set Locktopic=BoardID,BoardID=-1,IsBest=0 Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&Dv_IndivGroup_MainClass.BoardID&" And TopicID="&ID)
		End If
		Rs.Close:Set Rs=Nothing
	Next
	Del_Math_Forum_Count i-1,ChildNum
	Math_Forum_Today(Dv_IndivGroup_MainClass.BoardID)
	LastCount(Dv_IndivGroup_MainClass.BoardID)
	sucmsg="<li>帖子批量删除成功。"
	Dvbbs.Dvbbs_suc(sucmsg)
End Function

Function Topic_Batch_Isbest
	For i=1 To Request.Form("topicid").Count
		ID=Dvbbs.CheckNumeric(Request.Form("topicid")(i))
		Set Rs=Dvbbs.Execute("Select Top 1 * From Dv_Group_BBS Where GroupID="&Dv_IndivGroup_MainClass.ID&" And RootID="&ID&" And ParentID=0")
		If Not Rs.Eof Then
			Dvbbs.Execute("Update Dv_Group_BBS Set Isbest=1 Where GroupID="&Dv_IndivGroup_MainClass.ID&" And AnnounceID="&Rs("AnnounceID"))
			Dvbbs.Execute("update Dv_Group_Topic Set Isbest=1 Where GroupID="&Dv_IndivGroup_MainClass.ID&" And TopicID="&ID)
		End If
	Next
	Rs.Close:Set Rs=Nothing
	sucmsg="<li>帖子批量精华成功。"
	Dvbbs.Dvbbs_suc(sucmsg)
End Function

'更新指定论坛信息
Function LastCount(BoardID)
	Dim LastTopic,body,LastRootid,LastPostTime,LastPostUser
	Dim LastPost,uploadpic_n,Lastpostuserid,Lastid
	Set Rs=Dvbbs.Execute("Select Top 1 Announceid,Dateandtime,Username,Postuserid,Rootid From Dv_Group_BBS Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&BoardID&" Order By Announceid Desc")
	If Not Rs.Eof Then
		Lastid=Rs(0)
		LastPostTime=Rs(1)
		LastPostUser=Rs(2)
		LastPostUserid=Rs(3)
		LastRootid=Rs(4)
		Rs.Close:Set Rs=Nothing
		Set Rs=Dvbbs.Execute("Select Title From Dv_Group_Topic Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&BoardID&" And TopicID="&LastRootid)
		Lasttopic=Replace(Left(Rs(0),15),"$","")
	Else
		LastTopic="无"
		LastRootid=0
		LastPostTime=now()
		LastPostUser="无"
		LastPostUserid=0
		Lastid=0
	End If
	Rs.Close:Set Rs=Nothing
	LastPost = LastPostUser&"$"&Lastid&"$"&LastPostTime&"$"&LastTopic&"$$"&LastPostUserID&"$"&LastRootID&"$"&BoardID
	LastPost = Dvbbs.CheckStr(LastPost)
	Dvbbs.Execute("update dv_Board Set LastPost='"&LastPost&"' Where BoardID="&Dv_IndivGroup_MainClass.BoardID)
End Function

'更新圈子数据和版面数据
Function Del_Math_Forum_Count(TopicNum,BBSNUM)
	Dvbbs.Execute("Update Dv_GroupName Set postNum=postNum-"&BBSNUM+TopicNum&",TopicNum=TopicNum-"&TopicNum&" Where ID="&Dv_IndivGroup_MainClass.ID)
	Dvbbs.Execute("update Dv_Group_Board Set postnum=postnum-"&BBSNUM+TopicNum&",topicnum=topicnum-"&TopicNum&" Where RootID="&Dv_IndivGroup_MainClass.ID&" And ID="&Dv_IndivGroup_MainClass.BoardID)
End Function

'今日帖子
Function Math_Forum_Today(BoardID)
	Dim MathForumToday
	If IsSqlDataBase=1 then
		Set Rs=Dvbbs.Execute("Select Count(*) From Dv_Group_BBS Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&BoardID&" And Datediff(d,Dateandtime,"&SqlNowString&")=0")
	Else
		Set Rs=Dvbbs.Execute("Select Count(*) From Dv_Group_BBS Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&BoardID&" And Datediff('d',Dateandtime,"&SqlNowString&")=0")
	End if
	MathForumToday=Rs(0)
	Rs.Close:Set Rs=Nothing
	If Isnull(MathForumToday) Then MathForumToday=0
	Dvbbs.Execute("Update Dv_Group_Board Set Todaynum="&MathForumToday&" Where ID="&BoardID)
	If IsSqlDataBase=1 then
		Set Rs=Dvbbs.Execute("Select Count(*) From Dv_Group_BBS Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID>0 And Datediff(d,Dateandtime,"&SqlNowString&")=0")
	Else
		Set Rs=Dvbbs.Execute("Select Count(*) From Dv_Group_BBS Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID>0 And Datediff('d',Dateandtime,"&SqlNowString&")=0")
	End if
	MathForumToday=Rs(0)
	Rs.Close:Set Rs=Nothing
	If Isnull(MathForumToday) Then MathForumToday=0
	Dvbbs.Execute("Update Dv_GroupName Set Todaynum="&MathForumToday&" Where ID="&Dv_IndivGroup_MainClass.ID)
End Function
%>