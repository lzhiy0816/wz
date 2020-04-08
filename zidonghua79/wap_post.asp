<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/Class_Mobile.asp"-->
<!--#include file="inc/ubblist.asp"-->
<%

Dim Topic,Body,ID,RootID,AnnounceID,ParentID,FoundErr
Dim BoardInfo
Dim DateTimeStr,Expression,MyLastPostTime,TotalUseTable,layer,orders,LastPost,LastPost_1
Dim UbblistBody
Dim OpType
Dim Toptopic,uploadpic_n,IsAudit
Dim IsTop,Child,smsuserlist
Dim Tools_LastPostTime,Tools_UseTools
Dim Myistop
Myistop=0

OpType = DvbbsWap.ChkNumeric(Trim(Request("optype")))
FoundErr = False

'http://dvbbs1:81/wap_Post.asp?path=/119/1/4938/&Stype=4&optype=4&Mobile=1234567890123&name=&content=12356789102
'http://dvbbs1:81/wap_Post.asp?path=/119/1/&Stype=3&optype=3&Mobile=1234567890123&name=&content=12356789102

DvbbsWap.ShowXMLStar
If DvbbsWap.Stype<3 or Dvbbs.UserID=0 or DvbbsWap.PathCount<2 or OpType<3 Then 
	DvbbsWap.ShowErr 0,"参数错误，请确认从有效的地址访问！"
Else
	PostData
	CheckBoard
	ChkPostData
	Select case OpType
		Case 3 : SaveData_New
		Case 4 : SaveData_Reply
		Case 5 : SaveData_Edit
	End Select
End If
DvbbsWap.ShowXMLEnd

'获取提交数据
Sub PostData()
	Topic = DvbbsWap.Checkstr(Request("Name"))
	Body = DvbbsWap.Checkstr(Request("content"))
	If Body = "" Then
		DvbbsWap.ShowErr 0,"内容不能为空！"
		FoundErr = True
		Exit Sub
	End If
	If OpType=3 and DvbbsWap.Stype=3 Then	'发表
		Dvbbs.BoardID = DvbbsWap.ChkNumeric(DvbbsWap.Path(DvbbsWap.PathCount-1))
	ElseIf OpType=4 and DvbbsWap.Stype=4 Then	'回复主题
		Dvbbs.BoardID = DvbbsWap.ChkNumeric(DvbbsWap.Path(DvbbsWap.PathCount-2))
		ID = DvbbsWap.ChkNumeric(DvbbsWap.Path(DvbbsWap.PathCount-1))	'主题ID
		RootID = ID
	ElseIf OpType=5 and DvbbsWap.Stype=4 and DvbbsWap.PathCount>2 Then	'编辑
		Dvbbs.BoardID = DvbbsWap.ChkNumeric(DvbbsWap.Path(DvbbsWap.PathCount-3))
		RootID = DvbbsWap.ChkNumeric(DvbbsWap.Path(DvbbsWap.PathCount-2))
		ID = DvbbsWap.ChkNumeric(DvbbsWap.Path(DvbbsWap.PathCount-1))
	Else
		DvbbsWap.ShowErr 0,"参数错误，操作终止！"
		FoundErr = True
		Exit Sub
	End If
	If Dvbbs.BoardID=0 Then
		DvbbsWap.AddErrCode(29)
		FoundErr = True
		Exit Sub
	End If

	If OpType>3 and (ID=0 Or RootID=0) Then
		DvbbsWap.ShowErr 0,"帖子参数错误！"
		FoundErr = True
		Exit Sub
	End If
End Sub

'获取版块信息及权限
Sub CheckBoard()
	If FoundErr Then Exit Sub
	Dim Rs,Sql
	Sql = "Select BoardID,BoardType,Child,Readme,BoardMaster,BoardUser,Board_User,Board_Setting,IsGroupSetting From Dv_board where BoardID="&Dvbbs.BoardID
	Set Rs = Dvbbs.Execute(Sql)
	If not Rs.Eof Then
		BoardInfo = Rs.GetRows(1)
	Else
		DvbbsWap.ShowErr 0,"错误的版面参数！"
		FoundErr = True
		Exit Sub
	End If
	Rs.Close
  	Set Rs = Nothing

	DvbbsWap.LoadBoardPass
	''''版块权限，用户组是否有自定义权限，是否版主，是否认证等权限（待加）
End Sub

'发表验证
Sub ChkPostData()
	If FoundErr Then Exit Sub

	If Dvbbs.Board_Setting(43)="1" Then
		DvbbsWap.AddErrCode(72)
		FoundErr = True
		Exit Sub
	End If
	If Dvbbs.Board_Setting(1)="1" and Dvbbs.GroupSetting(37)="0" Then 
		Dvbbs.AddErrCode(26)
		FoundErr = True
		Exit Sub
	End If

	If Dvbbs.UserID>0 Then
		If Clng(Dvbbs.GroupSetting(52))>0 And DateDiff("s",Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@joindate").text(14),Now)<Clng(Dvbbs.GroupSetting(52))*60 Then
			DvbbsWap.ShowErr 0,"新注册用户在"&Dvbbs.GroupSetting(52)&"分钟内不能发言。"
			FoundErr = True
			Exit Sub
		End If
		If Dvbbs.GroupSetting(62)<>"0" And Not OpType=5 Then
			If Clng(Dvbbs.GroupSetting(62))<=Clng(Dvbbs.UserToday(0)) Then 
				DvbbsWap.ShowErr 0,"你的今日发帖数已超出了最大限制"&Dvbbs.GroupSetting(62)&"篇"
				FoundErr = True
				Exit Sub
			End If
		End If
	End If
	If Dvbbs.GroupSetting(3)="0" And OpType=3 Then
		DvbbsWap.AddErrCode(28)
		FoundErr = True
		Exit Sub
	End If
	
	If Dvbbs.GroupSetting(5)="0" And OpType=5 Then
		DvbbsWap.AddErrCode(29)
		FoundErr = True
		Exit Sub
	End If

	If Len(Topic)>250 Then Topic = Left(DvbbsWap.ForMatHtml(Topic),255)
	'Body = DvbbsWap.ForMatHtml(Body)
	DateTimeStr=Replace(Replace(CSTR(NOW()+Dvbbs.Forum_Setting(0)/24),"上午",""),"下午","")
	Expression = "face1.gif"
	MyLastPostTime=Replace(Replace(CSTR(NOW()+Dvbbs.Forum_Setting(0)/24),"上午",""),"下午","")
	TotalUseTable = Dvbbs.NowUseBBS
	If Dvbbs.strLength(topic)> CLng(Dvbbs.Board_Setting(45)) Then
		DvbbsWap.ShowErr 0,"主题不能超过该版限制"&Dvbbs.Board_Setting(45)&"个字符！"
		FoundErr = True
		Exit Sub
	End If
	If Dvbbs.strLength(Body) > CLng(Dvbbs.Board_Setting(16)) Then
		DvbbsWap.ShowErr 0,"内容不能超过该版限制"&Dvbbs.Board_Setting(16)&"个字符！"
		FoundErr = True
		Exit Sub
	End If
	If Dvbbs.strLength(Body) < CLng(Dvbbs.Board_Setting(52)) Then
		DvbbsWap.ShowErr 0,"内容不能少于该版限制"&Dvbbs.Board_Setting(52)&"个字符！"
		FoundErr = True
		Exit Sub
	End If

	If Dvbbs.GroupSetting(64)="1" Then
 		IsAudit=0
 	Else
 		If Dvbbs.Board_Setting(57)="1" Then
 			IsAudit=NeedIsAudit(Body)
 		End If
 	End If
End Sub

'保存数据=================================================
'发表
Sub SaveData_New()
	If FoundErr Then Exit Sub
	If Topic = "" Then
		DvbbsWap.ShowErr 0,"主题不能为空！"
		FoundErr = True
		Exit Sub
	End If
	layer = 1
	orders = 0
	ParentID = 0
	Insert_To_Topic()
	Insert_To_Announce()
	UpdateData
	'更新版块发帖数，论坛统计，用户积分
	DvbbsWap.ShowErr 1,"发表成功！"
End Sub

'回复
Sub SaveData_Reply()
	If FoundErr Then Exit Sub
	Get_SaveRe_TopicInfo
	If FoundErr Then Exit Sub
	Get_ForumTreeCode
	If FoundErr Then Exit Sub
	Insert_To_Announce
	UpdateData
	DvbbsWap.ShowErr 1,"回复成功！"
End Sub

Sub UpdateData()
	Dim Sql
	topic=Replace(Replace(cutStr(topic,14),chr(10),""),"'","")
	If topic="" Then
		topic=Body
		topic=Replace(cutStr(topic,14),chr(10),"")
	Else
		topic=Replace(cutStr(topic,14),chr(10),"")
	End If

	LastPost=Replace(Dvbbs.MemberName,"$","") & "$" & AnnounceID & "$" & DateTimeStr & "$" & Replace(cutStr(topic,20),"$","&#36;") & "$" & uploadpic_n & "$" & Dvbbs.UserID & "$" & RootID & "$" & Dvbbs.BoardID
		
	LastPost=reubbcode(Replace(LastPost,"'",""))
	LastPost=Dvbbs.ChkBadWords(LastPost)
	If IsAudit<>1 Then
			Dim IsUpDateLastPostTime
			IsUpDateLastPostTime = False
			If istop>0 Or InStr("," & Tools_UseTools & ",",",13,")>0 Or InStr("," & Tools_UseTools & ",",",14,")>0 Then IsUpDateLastPostTime = True
			If OpType = 4 Then 
				If istop=3 And IsUpDateLastPostTime Then
					If IsSqlDataBase=1 Then
						SQL="update Dv_topic set child=child+1,LastPostTime=dateadd(day,300,"&SqlNowString&"),LastPost='"&LastPost&"' where TopicID="&RootID
					Else
						SQL="update Dv_topic set child=child+1,LastPostTime=dateadd('d',300,"&SqlNowString&"),LastPost='"&LastPost&"' where TopicID="&RootID
					End If
				ElseIf istop=2 And IsUpDateLastPostTime Then
					If IsSqlDataBase=1 Then
						SQL="update Dv_topic set child=child+1,LastPostTime=dateadd(day,200,"&SqlNowString&"),LastPost='"&LastPost&"' where TopicID="&RootID
					Else
						SQL="update Dv_topic set child=child+1,LastPostTime=dateadd('d',200,"&SqlNowString&"),LastPost='"&LastPost&"' where TopicID="&RootID
					End If
				ElseIf istop=1 And IsUpDateLastPostTime Then
					If IsSqlDataBase=1 Then
						SQL="update Dv_topic set child=child+1,LastPostTime=dateadd(day,100,"&SqlNowString&"),LastPost='"&LastPost&"' where TopicID="&RootID
					Else
						SQL="update Dv_topic set child=child+1,LastPostTime=dateadd('d',100,"&SqlNowString&"),LastPost='"&LastPost&"' where TopicID="&RootID
					End If
				ElseIf IsUpDateLastPostTime Then
					If DateDiff("h",Now(),Tools_LastPostTime)<=0 Then
						SQL="update Dv_topic set child=child+1,LastPostTime="&SqlNowString&",LastPost='"&LastPost&"' where TopicID="&RootID
					Else
						SQL="update Dv_topic set child=child+1,LastPost='"&LastPost&"' where TopicID="&RootID
					End If
				Else
					SQL="update Dv_topic set child=child+1,LastPostTime="&SqlNowString&",LastPost='"&LastPost&"' where TopicID="&RootID
				End If
				Dvbbs.Execute(SQL)
				'Child=Child+2
				'If Child mod Dvbbs.Board_Setting(27)=0 Then 
				'	star=Child\Dvbbs.Board_Setting(27)
				'Else
				'	star=(Child\Dvbbs.Board_Setting(27))+1
				'End If
				'Get_Chan_TopicOrder()
			Else
				Dvbbs.Execute("update Dv_topic set LastPost='"&LastPost&"' where topicid="&RootID)
			End If
	End If
		If OpType = 3 Then
			toptopic=Replace(topic,"$","&#36;")
		Else
			toptopic=Replace(cutStr(toptopic,20),"$","&#36;")
		End If
		If Dvbbs.Board_Setting(2)="1" Then
			LastPost_1="保密$" & AnnounceID & "$" & DateTimeStr & "$请认证用户进入查看$" & uploadpic_n & "$" & Dvbbs.UserID & "$" & RootID & "$" & Dvbbs.BoardID
		Else
			LastPost_1=Replace(Dvbbs.MemberName,"$","") & "$" & AnnounceID & "$" & DateTimeStr & "$" & toptopic & "$" & uploadpic_n & "$" & Dvbbs.UserID & "$" & RootID & "$" & Dvbbs.BoardID
		End If
		LastPost_1=reubbcode(Replace(LastPost_1,"'",""))
		LastPost_1=Dvbbs.ChkBadWords(LastPost_1)
		If IsAudit=0 Then
			UpDate_BoardInfoAndCache()
			UpDate_ForumInfoAndCache()
		End If

End Sub

'更新版面数据和缓存
Sub UpDate_BoardInfoAndCache()
		Dim Sql
		Dim UpdateBoardID
		If Application(Dvbbs.CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"&Dvbbs.BoardID&"']/@parentstr").text <> "" Then 
			UpdateBoardID=Application(Dvbbs.CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"&Dvbbs.BoardID&"']/@parentstr").text & "," & Dvbbs.BoardID	
		Else
			UpdateBoardID=Dvbbs.BoardID
		End If

		If OpType = 4 Then
			SQL="update Dv_board set PostNum=PostNum+1,todaynum=todaynum+1,LastPost='"&SimEncodeJS(LastPost_1)&"' where boardid in ("&UpdateBoardID&")"
		ElseIf OpType = 3 Then
			SQL="update Dv_board set PostNum=PostNum+1,TopicNum=TopicNum+1,todaynum=todaynum+1,LastPost='"&SimEncodeJS(LastPost_1)&"' where boardid in ("&UpdateBoardID&")"
		End If
		Dvbbs.Execute(sql)
		Dvbbs.ReloadBoardInfo(UpdateBoardID)
End Sub
Sub UpDate_ForumInfoAndCache()
		Dim Sql
		Dim updateinfo,LastPostTime,LastPostTimes
		Dim Forum_LastPost,Forum_TodayNum,Forum_MaxPostNum
		Forum_LastPost=Dvbbs.CacheData(15,0)
		Forum_TodayNum=Dvbbs.CacheData(9,0)
		Forum_MaxPostNum=Dvbbs.CacheData(12,0)
		LastPostTimes=split(Forum_LastPost,"$")
		LastPostTime=LastPostTimes(2)
		If Not IsDate(LastPostTime) Then LastPostTime=Now()
		If datediff("d",LastPostTime,Now())=0 Then
			If CLng(Forum_TodayNum)+1 > CLng(Forum_MaxPostNum) Then 
				updateinfo=",Forum_MaxPostNum=Forum_TodayNum+1,Forum_MaxPostDate="&SqlNowString&""
				Dvbbs.ReloadSetupCache Now(),13
				Dvbbs.ReloadSetupCache CLng(Forum_TodayNum)+1,12
			End If
			Dvbbs.ReloadSetupCache CLng(Forum_TodayNum)+1,9
			If OpType = 4 Then
				SQL="update Dv_setup set Forum_PostNum=Forum_PostNum+1,Forum_TodayNum=Forum_TodayNum+1,Forum_LastPost='"&LastPost&"' "&updateinfo&" "
			Else
				SQL="update Dv_setup set Forum_TopicNum=Forum_TopicNum+1,Forum_PostNum=Forum_PostNum+1,Forum_TodayNum=Forum_TodayNum+1,Forum_LastPost='"&LastPost&"' "&updateinfo&" "
			End If
		Else
			If OpType = 4 Then
				SQL="update Dv_setup set Forum_PostNum=Forum_PostNum+1,forum_YesTerdayNum="&CLng(Forum_TodayNum)&",Forum_TodayNum=1,Forum_LastPost='"&LastPost&"' "
			Else
				SQL="update Dv_setup set Forum_TopicNum=Forum_TopicNum+1,Forum_PostNum=Forum_PostNum+1,forum_YesTerdayNum="&CLng(Forum_TodayNum)&",Forum_TodayNum=1,Forum_LastPost='"&LastPost&"' "
			End If
			Dvbbs.ReloadSetupCache 1,9
		End If
		'更新总固顶部分数据和缓存
		If Not OpType = 4 Then
			If Myistop=2 Then
				Dim tmpstr
				If Dvbbs.CacheData(28,0)="" Then
					tmpstr=", Forum_alltopnum='"&RootID&"'"
					Dvbbs.ReloadSetupCache RootID,28
				Else
					tmpstr=", Forum_alltopnum='"&Dvbbs.CacheData(28,0)&","&RootID&"'"
					Dvbbs.ReloadSetupCache Dvbbs.CacheData(28,0)&","&RootID,28
				End If 
				SQL=SQl&tmpstr
			End If
			Dvbbs.ReloadSetupCache CLng(Dvbbs.CacheData(7,0))+1,7'主题数
		End If
		Dvbbs.ReloadSetupCache CLng(Dvbbs.CacheData(8,0))+1,8'文章数
		Dvbbs.ReloadSetupCache LastPost,15
		Dvbbs.Execute(SQL)
End Sub

Sub Get_SaveRe_TopicInfo()
	Dim Rs,Sql

		SQL="select locktopic,LastPost,title,smsuserlist,IsSmsTopic,istop,Child,UseTools,LastPostTime,GetMoneyType,PostUserID, DateAndTime from Dv_topic where BoardID="&Dvbbs.BoardID&" And topicid="&cstr(RootID)
		Set Rs=dvbbs.Execute(sql)
		If Not (Rs.EOF And Rs.BOF) Then
			toptopic = Rs(2)
			istop = Rs(5)
			Child = Rs(6)
			Tools_UseTools = "," & Rs(7) & ","
			Tools_LastPostTime = Rs(8)
			'If Rs("IsSmsTopic")=1 Then smsuserlist=Rs("smsuserlist")
			If Rs("LockTopic")=1 And Not (Dvbbs.Master Or Dvbbs.BoardMaster Or Dvbbs.SuperBoardMaster) Then
				DvbbsWap.AddErrCode(78)
				FoundErr = True
				Exit Sub
			End If
			'锁定多少天前的帖子判断 2004-9-16 Dv.Yz
			If Ubound(Dvbbs.Board_Setting) > 70 And Not (Dvbbs.Master Or Dvbbs.BoardMaster Or Dvbbs.SuperBoardMaster) Then
				If Not Clng(Dvbbs.Board_Setting(71)) = 0 And Datediff("d", Rs(11), Now()) > Clng(Dvbbs.Board_Setting(71)) Then 
					DvbbsWap.AddErrCode(78)
					FoundErr = True
					Exit Sub
				End If
			End If
			If Not IsNull(Rs(1)) Then
				LastPost=split(Rs(1),"$")
				If ubound(LastPost)=7 Then
					UpLoadPic_n=LastPost(4)
				Else
					UpLoadPic_n=""
				End If
			End If
			If Cint(Rs(9)) = 5 Then
				DvbbsWap.ShowErr 0,"主题已结帖，不允许回复。"
				FoundErr = True
				Exit Sub
			End If
		Else
			DvbbsWap.ShowErr 0,"该帖子不存在！"
			FoundErr = True
			Exit Sub
		End If
		Set Rs = Nothing
End Sub

Public Sub Get_ForumTreeCode()
	Dim Rs,Sql
	Sql = "Select AnnounceID,PostUserID From "&TotalUseTable&" where ParentID=0 and RootID="&RootID
	Set Rs = Dvbbs.Execute(Sql)
	If Rs.Eof Then
		DvbbsWap.ShowErr 0,"该帖子不存在！"
		FoundErr = True
		Exit Sub
	Else
		ParentID = Rs(0)
		If Rs(1)=Dvbbs.UserID Then
			If Cint(Dvbbs.GroupSetting(4))=0 Then
				DvbbsWap.AddErrCode(73)
				FoundErr = True
				Exit Sub
			End If
		End If
	End If
	Rs.Close
	Sql = "Select Max(layer),Max(orders) From "&TotalUseTable&" where RootID="&RootID
	Set Rs=Dvbbs.Execute(sql)
	If Not(rs.EOF And rs.BOF) Then
			If IsNull(Rs(0)) Then
				Layer=1
			Else
				Layer=Rs(0)+1
			End If
			If IsNull(Rs(1)) Then
				Orders=0
			Else
				Orders=Rs(1)+1
			End If
	Else
			Layer=1
			Orders=0
	End If
	Rs.Close
	Set Rs=Nothing
End Sub


'插入主题
Sub Insert_To_Topic()
	Dim Sql
	SQL="insert into Dv_topic (Title,Boardid,PostUsername,PostUserid,DateAndTime,Expression,LastPost,LastPostTime,PostTable,locktopic,istop,TopicMode,isvote,PollID,Mode,GetMoney,GetMoneyType,IsSmsTopic) values ('"&Topic&"',"&Dvbbs.Boardid&",'"&Dvbbs.MemberName&"',"&Dvbbs.Userid&",'"&DateTimeStr&"','"&Expression&"','$$"&DateTimeStr&"$$$$','"&MyLastPostTime&"','"&TotalUseTable&"',0,0,0,0,0,0,0,0,1)"
	Dvbbs.Execute(sql)
	RootID=Dvbbs.Execute("select Max(topicid) From Dv_topic Where PostUserid="&Dvbbs.UserID)(0)
End Sub

'插入回复
Sub Insert_To_Announce()
	Dim Sql
	Body = Html2Ubb(Body)
	UbblistBody = Ubblist(Body)
	SQL="insert into "&TotalUseTable&"(Boardid,ParentID,username,topic,body,DateAndTime,length,RootID,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID,isupload,IsAudit,Ubblist,GetMoney,GetMoneyType) values ("&Dvbbs.boardid&","&ParentID&",'"&Dvbbs.MemberName&"','"&Topic&"','"&Body&"','"&DateTimeStr&"','"&Dvbbs.strlength(Body)&"',"&RootID&","&layer&","&orders&",'"&Dvbbs.UserTrueIP&"','"&Expression&"',0,0,0,0,"&Dvbbs.userid&",2,0,'"&UbblistBody&"',0,0)"
	Dvbbs.Execute(sql)
	AnnounceID=Dvbbs.Execute("select Max(AnnounceID) From "&TotalUseTable&" Where PostUserID="&Dvbbs.UserID)(0)
End Sub


'编辑
Sub SaveData_Edit()
	If FoundErr Then Exit Sub
	Dim Rs,Sql
	Dim PostUserID,CanEditPost,UserGroupID,IsTopic,LockTopic,istop,dateandtime
	CanEditPost = False
	IsTopic = False
	Sql = "Select Title,LockTopic,PostTable,PostUserID,istop From [Dv_Topic] where BoardID="&Dvbbs.BoardID&" and TopicID="&RootID
	Set Rs = Dvbbs.Execute(sql)
	If Rs.Eof Then
		DvbbsWap.ShowErr 0,"该帖子不存在！"
		FoundErr = True
		Exit Sub
	Else
		TotalUseTable = Rs(2)
		istop = Rs(4)
	End If
	Rs.Close

	Sql = "Select B.AnnounceID,B.Topic,B.Body,B.PostUserID,B.UbbList,B.ParentID,B.locktopic,B.DateAndTime,U.UserGroupID From "&TotalUseTable&" B, [Dv_user] U where B.PostUserID=U.UserID and BoardID="&Dvbbs.BoardID&" and AnnounceID="&ID
	Set Rs = Dvbbs.Execute(sql)

	If Rs.Eof Then
		DvbbsWap.ShowErr 0,"该帖子不存在！"
		FoundErr = True
		Exit Sub
	Else
		AnnounceID = Rs(0)
		PostUserID = Rs(3)
		If Rs(5)=0 Then
			IsTopic = True
		Else
			Topic = Rs(1)
		End If
		LockTopic = Rs(6)
		DateAndTime = Rs(8)
		UserGroupID = Rs(8)
	End If
	Rs.Close

	If IsTopic and Topic="" Then
		DvbbsWap.ShowErr 0,"主题不能为空！"
		FoundErr = True
		Exit Sub
	End If
	If PostUserID=Dvbbs.UserID Then
		If Dvbbs.GroupSetting(10)="0" then
			DvbbsWap.AddErrCode(74)
			CanEditPost=False
			FoundErr = True
			Exit Sub
		Else 
			CanEditPost=True
		End If
	Else
		If (Dvbbs.Master or Dvbbs.Superboardmaster or Dvbbs.Boardmaster) and Dvbbs.GroupSetting(23)="1" then
			CanEditPost=True
		Else 
			CanEditPost=False
		End If 
		If Cint(Dvbbs.UserGroupID) > 3 And Dvbbs.GroupSetting(23)="1" Then CanEditPost=True
		If Dvbbs.GroupSetting(23)="1" and Dvbbs.founduserPer Then 
				CanEditPost=True
		ElseIf Dvbbs.GroupSetting(23)="0" And Dvbbs.founduserPer Then 
				CanEditPost=False
		End If
		If Not CanEditPost Then
			DvbbsWap.AddErrCode(74)
			FoundErr = True
			Exit Sub
		End If
		If Cint(Dvbbs.UserGroupID) < 4 And Cint(Dvbbs.UserGroupID) = UserGroupID Then 
				DvbbsWap.AddErrCode(75)
				FoundErr = True
		ElseIf Cint(Dvbbs.UserGroupID) < 4 and Cint(Dvbbs.UserGroupID) > UserGroupID Then
				DvbbsWap.AddErrCode(76)
				FoundErr = True
		End If
	End If
	If FoundErr Then
		Exit Sub
	End If
	If Not Dvbbs.master And LockTopic=1 then
		DvbbsWap.AddErrCode(78)
		FoundErr = True
		Exit Sub
	End If

	Dim char_changed
	Dim re,LastBoard,LastTopic,LastPost
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="\[align=right\]\[color=#000066\](.|\n)*\[\/color\]\[\/align\]"
	Body = re.Replace(Body,"")
	re.Pattern="<div align=right><font color=#000066>(.|\n)*<\/font><\/div>"
	Body = re.Replace(Body,"")
	Set re=Nothing

	If PostUserID<>Dvbbs.UserID Then 
		If Dvbbs.forum_setting(49)="1" Then char_changed = "[align=right][color=#000066][此贴子已经被"&Dvbbs.membername&"于"&Now()&"编辑过][/color][/align]"
	Else
		If Dvbbs.forum_setting(48)="1" Then char_changed = "[align=right][color=#000066][此贴子已经被作者于"&Now()&"编辑过][/color][/align]"
	End If

	If Clng(Dvbbs.forum_setting(50))>0 then
		If Datediff("s",DateAndTime,Now())>Clng(Dvbbs.forum_setting(50))*60 then
			Body = Body+chr(13)+chr(10)+char_changed+chr(13)
		End If
	Else
		Body = Body+chr(13)+chr(10)+char_changed+chr(13)
	End If
	If Clng(Dvbbs.forum_setting(51))>0 and not (Dvbbs.master or Dvbbs.boardmaster or Dvbbs.superboardmaster) Then 
		If DateDiff("s",DateAndTime,Now())>Clng(Dvbbs.forum_setting(51))*60 Then
			DvbbsWap.ShowErr 0,"论坛限制在："&Dvbbs.forum_setting(51)&"秒内不能编辑！"
			FoundErr = True
			Exit Sub
		End If
	End If 
		'取出当前版面最后回复id,如果本帖为最后回复则更新相应数据
		Set Rs = Dvbbs.Execute("select LastPost from dv_board where boardid="&Dvbbs.BoardID)
		If not (Rs.EOF And Rs.BOF) Then
			If Not IsNull(rs(0)) And rs(0)<>"" then
				LastBoard=split(rs(0),"$")
				If ubound(LastBoard)=7 Then
					If cCur(LastBoard(6))=cCur(AnnounceID) Then
						LastPost=LastBoard(0) & "$" & LastBoard(1) & "$" & Now() & "$" & Replace(cutStr(reubbcode(topic),20),"$","&#36;") & "$" & LastBoard(4) & "$" & LastBoard(5) & "$" & LastBoard(6) & "$" & Dvbbs.BoardID
						dvbbs.execute("update dv_board set LastPost='"&SimEncodeJS(Replace(LastPost,"'",""))&"' where boardid="&Dvbbs.BoardID)
					End If
				End If
			End If
		End If

		'取得当前主题最后回复id,如果本帖为最后回复则更新相应数据
		Set Rs=Dvbbs.Execute("select LastPost,istop from dv_topic where topicid="&rootid)
		If Not (Rs.Eof And Rs.Bof) Then
			istop=rs(1)
			If Not Isnull(Rs(0)) And Rs(0)<>"" Then
				LastTopic=split(rs(0),"$")
				If Ubound(LastTopic)=7 Then
					If cCur(LastTopic(1))=cCur(Announceid) Then
						LastPost=LastTopic(0) & "$" & LastTopic(1) & "$" & Now() & "$" & Replace(cutStr(reubbcode(body),20),"$","&#36;") & "$" & LastTopic(4) & "$" & LastTopic(5) & "$" & LastTopic(6) & "$" & Dvbbs.BoardID
						dvbbs.execute("update dv_topic set LastPost='"&Replace(LastPost,"'","")&"' where topicid="&rootid)
					End If
				End If
			End If
		End If

		Set Rs = Server.CreateObject("ADODB.Recordset")
		SQL="SELECT * FROM "&TotalUseTable&" where AnnounceID="&Announceid
		rs.Open SQL,conn,1,3
		If not (Rs.EOF And Rs.BOF) Then
			If Rs("parentid")=0 then
				If istop=1 Then
					If IsSqlDataBase=1 Then
						dvbbs.execute("update dv_topic set title='"&topic&"',LastPostTime=dateadd(day,100,"&SqlNowString&") where topicid="&rootid)
					Else
						dvbbs.execute("update dv_topic set title='"&topic&"',LastPostTime=dateadd('d',100,"&SqlNowString&") where topicid="&rootid)
					End If
				ElseIf istop=3 Then
					If IsSqlDataBase=1 Then
						dvbbs.execute("update dv_topic set title='"&topic&"',LastPostTime=dateadd(day,300,"&SqlNowString&") where topicid="&rootid)
					Else
						dvbbs.execute("update dv_topic set title='"&topic&"',LastPostTime=dateadd('d',300,"&SqlNowString&") where topicid="&rootid)
					End If
				Else
					dvbbs.execute("update dv_topic set title='"&topic&"' where topicid="&rootid)
				End If
			End If
			Body = Html2Ubb(Body)
			Rs("Topic") = Topic
			Rs("Body") = Body
			Rs("length")= Dvbbs.strlength(Body)
			Rs("ip")= Dvbbs.UserTrueIP
			'If Rs("isupload")=0 And ihaveupfile=1 Then Rs("isupload")=1
			Rs("isupload")=2	'WAP标识
			UbblistBody = Ubblist(Body)
			Rs("Ubblist")=UbblistBody
			Rs.Update
			'If ihaveupfile=1 Then dvbbs.execute("update dv_upfile set F_AnnounceID='"&rootid&"|"&AnnounceID&"',F_Readme='"&Replace(Rs("Topic"),"'","''")&"',F_flag=0 where F_ID in ("&upfileinfo&")")
			DvbbsWap.ShowErr 1,"编辑成功！"
		End If	
		Rs.Close
		Set Rs=Nothing
End Sub

'截取指定字符
Function cutStr(str,strlen)
	'去掉所有HTML标记
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="<(.[^>]*)>"
	str=re.Replace(str,"")	
	set re=Nothing
	Dim l,t,c,i
	l=Len(str)
	t=0
	For i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		If c>255 Then
			t=t+2
		Else
			t=t+1
		End If
		If t>=strlen Then
			cutStr=left(str,i)&"..."
			Exit For
		Else
			cutStr=str
		End If
	Next
	cutStr=Replace(cutStr,chr(10),"")
	cutStr=Replace(cutStr,chr(13),"")
End Function
'过滤不必要UBB
Function reUBBCode(strContent)
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	strContent=Replace(strContent,"&nbsp;"," ")
	re.Pattern="(\[QUOTE\])(.|\n)*(\[\/QUOTE\])"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\[point=*([0-9]*)\])(.|\n)*(\[\/point\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[post=*([0-9]*)\])(.|\n)*(\[\/post\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[power=*([0-9]*)\])(.|\n)*(\[\/power\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[usercp=*([0-9]*)\])(.|\n)*(\[\/usercp\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[money=*([0-9]*)\])(.|\n)*(\[\/money\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[replyview\])(.|\n)*(\[\/replyview\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[usemoney=*([0-9]*)\])(.|\n)*(\[\/usemoney\])"
	strContent=re.Replace(strContent,"&nbsp;")
	strContent=Replace(strContent,"<I></I>","")
	set re=Nothing
	reUBBCode=strContent
End Function

Function SimEncodeJS(str)
	If Not IsNull(str) Then
		str = Replace(str, "\", "\\")
		str = Replace(str, chr(34), "\""")
		str = Replace(str, chr(39), "\'")
		str = Replace(str, chr(10), "\n")
		str = Replace(str, chr(13), "\r")
		SimEncodeJS=str
	End If
End Function

'发贴时用，为了减少入库量
Function Html2Ubb(str)
	If Str<>"" And Not IsNull(Str) Then
		Dim re,tmpstr
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern = "(<br>)"
		Str = re.Replace(Str,"[br]")
		If Dvbbs.Board_Setting(5)="0" Then
			'先去掉标记中的换行
			re.Pattern="(<(i|b|p)>)"
			Str=re.Replace(Str,"[$2]")
			re.Pattern="(<(\/i|\/b|\/p)>)"
			Str=re.Replace(Str,"[$2]")
			re.Pattern="(>)("&vbNewLine&")(<)"
			Str=re.Replace(Str,"$1$3") 
			re.Pattern="(>)("&vbNewLine&vbNewLine&")(<)"
			Str=re.Replace(Str,"$1$3")
			re.Pattern="(<DIV class=quote>)((.|\n)*)(<\/div>)"
			Str=re.Replace(Str,"[quote]$2[/quote]")
			re.Pattern="<(.[^>]*)>"
			Str=re.Replace(Str,"")
			re.Pattern="(\[(i|b|p)\])"
			Str=re.Replace(Str,"<$2>")
			re.Pattern="(\[(\/i|\/b|\/p)\])"
			Str=re.Replace(Str,"<$2>")
		End If
		Str = Replace(Str, "[br]", CHR(13) & CHR(10))
		re.Pattern = "(&nbsp;)"
		Str = re.Replace(Str,Chr(9))
		re.Pattern = "(<STRONG>)"
		Str = re.Replace(Str,"<b>")
		re.Pattern = "(<\/STRONG>)"
		Str = re.Replace(Str,"</b>")
		re.Pattern ="(<TBODY>)"
		Str = re.Replace(Str,"")
		re.Pattern ="(<\/TBODY>)"
		Str = re.Replace(Str,"")
		Set Re=Nothing
		Html2Ubb = Str
	Else
		Html2Ubb = ""
	End If
End Function
'检查贴中是否含过滤字
Function NeedIsAudit(Content)
		NeedIsAudit=0
		Dim i,ChecKData
		If Dvbbs.Board_Setting(58)<>"0" Then
			ChecKData=split(Dvbbs.Board_Setting(58),"|")
			For i=0 to UBound(ChecKData)
				If Trim(ChecKData(i))<>"" Then
					If InStr(Content,ChecKData(i))>0 Or InStr(Topic,ChecKData(i))>0 Then
						NeedIsAudit=1
						Exit Function
					End If
				End If
			Next
		End If		
End Function
%>