<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/Class_Mobile.asp"-->
<%
'版面列表
'http://bbs.dvbbs.net/wap_board.asp?path=/&stype=3&startid=19&child=4&echild=2&mobile=123456789123&number=1
Dim EmotPath,Forum_Url
Dim FoundTopTopic
FoundTopTopic = False
Dvbbs.LoadTemplates("dispbbs")
Forum_Url = Dvbbs.Get_ScriptNameUrl
EmotPath= Forum_Url & Split(Dvbbs.Forum_emot,"|||")(0)		'em心情路径

DvbbsWap.ShowXMLStar
If DvbbsWap.PathCount = -1 or DvbbsWap.Mobile=0 or DvbbsWap.Child=0 Then
	DvbbsWap.ShowErr 0,"参数错误，请确认从有效的地址访问！"
Else
	Select Case DvbbsWap.Child
		Case 1 :	'只显示版块
			ShowBoard()
		Case 2 :	'只显示主题
			ShowTopic()
		Case 3 :	'显示主题及版块
			ShowBoard()
			ShowTopic()
		Case 4 :	'浏览帖子
			Dim Announceid,TotalUseTable,FoundErr,Star
			Dim TopicCount,Topic
			FoundErr = False
			If DvbbsWap.StartID="" Then DvbbsWap.StartID=1
			Star = Clng(DvbbsWap.StartID)	'分页
			Chk_Topic_Err
			If FoundTopTopic Then
				ShowTopic()
			Else
				ShowDispbbs()
			End If
		Case Else
			DvbbsWap.ShowErr 0,"参数错误，请确认从有效的地址访问！"
	End Select
End If

DvbbsWap.ShowXMLEnd

'显示版块
Sub ShowBoard()
	Dim Rs,Sql,SearchSQL
	Dim i,Board_Datas,LastPost,Setings,Loadboard
	Dim Show_Content
	Dim TotalNum,n,page
	Sql = "select boardid,BoardType,ParentID,ParentStr,Depth,RootID,Child,Readme,BoardMaster,PostNum,TopicNum,indexIMG,todayNum,boarduser,LastPost,Sid,Board_Setting,Board_user,IsGroupSetting,BoardTopStr From Dv_board "
	If DvbbsWap.PathCount = 1 or DvbbsWap.PathCount = -1 Then
		SearchSQL = "where ParentID=0 "
		Sql = Sql & SearchSQL
		DvbbsWap.Stype = 1
	Else
		If Not IsNumeric(DvbbsWap.Path(DvbbsWap.PathCount-1)) or DvbbsWap.Path(DvbbsWap.PathCount-1)="" Then Exit Sub
		SearchSQL = "where ParentID="& DvbbsWap.Path(DvbbsWap.PathCount-1)
		Sql = Sql & SearchSQL
		DvbbsWap.Stype = 2
	End If
	Sql = Sql & " order by orders,boardid"

	If DvbbsWap.StartID="" or DvbbsWap.StartID=0 Then DvbbsWap.StartID=1
	page = Clng(DvbbsWap.StartID)	'分页
	If DvbbsWap.Number = 0 THen DvbbsWap.Number = 10	'每页记录数
	TotalNum = Dvbbs.Execute("Select Count(*) From Dv_Board "&SearchSQL)(0)
	Set Rs = Dvbbs.Execute(Sql)
	If Rs.Eof And Rs.Bof Then
		DvbbsWap.ShowErr 0,"找不到数据，请返回！"
		Rs.Close
		Exit Sub
	End If
	If TotalNum Mod DvbbsWap.Number=0 Then
				n = TotalNum \ DvbbsWap.Number
	Else
	     		n = TotalNum \ DvbbsWap.Number+1
  	End If
	Rs.MoveFirst
	Response.Write "<pagecounter>"&N&"</pagecounter>"
	If page > n Then
		page = n
		DvbbsWap.ShowErr 0,"参数错误，请返回！"
		Exit Sub
	End If
	If page < 1 Then page = 1
	If page > 1 Then 				
		Rs.Move (page-1) * DvbbsWap.Number
	End if
	If Rs.Eof Then
		Rs.Close
		DvbbsWap.AddErrCode(29)
		Exit Sub
	Else
		Board_Datas = Rs.GetRows(DvbbsWap.Number)
		Rs.Close
	End If

	'显示我发表的主题链接
	'ShowCodes(S_Self ,S_Child ,S_Sid ,S_Stype ,S_Name ,S_Content ,S_OtherContent,S_Author ,S_Createtime ,S_Modifytime)
	If (DvbbsWap.PathCount = 1 or DvbbsWap.PathCount = -1) And Dvbbs.UserID > 0 Then
		DvbbsWap.ShowCodes DvbbsWap.Self,2,-1,DvbbsWap.Stype ,"我发表的主题" ,DvbbsWap.Format_Content( 0 , "我在此论坛发表过的主题列表" ) ,"","自己" ,"" ,Now()
	End If
	'显示我发表的主题链接
	For i=0 To Ubound(Board_Datas,2)
			Loadboard = True
			Setings = Split(Board_Datas(16,i),",")
			If CInt(Setings(1))=1 And CInt(Dvbbs.GroupSetting(37))<>1 Then Loadboard = False
			If Loadboard Then
				LastPost = Split(Board_Datas(14,i),"$")(2)
				If Not IsDate(LastPost) Then LastPost = Now()
				Show_Content = DvbbsWap.Format_Content( 0 , DvbbsWap.ForMatHtml(Board_Datas(7,i)) )
				If Clng(Board_Datas(6,i))>0 Then	'当有下属版块，设置显示主题则为3，不显示为1，没有下属版块则为2
					If Cint(Setings(43)) = 1 Then
						DvbbsWap.Child = 1
					Else
						DvbbsWap.Child = 3
					End If
				Else
					DvbbsWap.Child = 2
				End If
				DvbbsWap.ShowCodes DvbbsWap.Self,DvbbsWap.Child,Board_Datas(0,i),DvbbsWap.Stype ,Board_Datas(1,i) ,Show_Content ,"",Board_Datas(8,i) ,"" ,LastPost
			End If
	Next
	Set Rs=Nothing
End Sub

'显示主题
Sub ShowTopic()
	Dim Rs,Sql
	Dvbbs.BoardID = DvbbsWap.Path(DvbbsWap.PathCount-1)
	If Dvbbs.BoardID<>"" And IsNumeric(Dvbbs.BoardID) Then
		Dvbbs.BoardID = Clng(Dvbbs.BoardID)
	Else
		DvbbsWap.AddErrCode(29)
		Exit Sub
	End If

	Select Case Dvbbs.BoardID
	Case -1
		'查看自己的帖子
		'Exit Function
	Case -2
		'查看固顶类帖子
		Dvbbs.BoardID = DvbbsWap.Path(DvbbsWap.PathCount-2)
		If Dvbbs.BoardID<>"" And IsNumeric(Dvbbs.BoardID) Then
			Dvbbs.BoardID = Clng(Dvbbs.BoardID)
		Else
			DvbbsWap.AddErrCode(29)
			Exit Sub
		End If
		DvbbsWap.LoadBoardPass
	Case Else
		DvbbsWap.LoadBoardPass
	End Select

	'查看固顶和总固顶主题
	If Not Dvbbs.BoardID = -1 And Not FoundTopTopic Then DvbbsWap.ShowCodes DvbbsWap.Self,4,-2 ,1 ,"查看本论坛重要通知和固顶主题" ,"查看本论坛总固顶、区域固顶、版面固顶主题" ,"","管理员" ,Now() ,Now()

	Dim TopicMode,page,ti,n,Cmd,limitime,LastPost,LastPostTime
	Dim S_Content
	Dim TopicNum
	TopicMode = 0
	If DvbbsWap.StartID="" or DvbbsWap.StartID=0 Then DvbbsWap.StartID=1
	page = Clng(DvbbsWap.StartID)	'分页
	If DvbbsWap.Number = 0 THen DvbbsWap.Number = 10	'每页记录数

	If Dvbbs.BoardID > 0 Then TopicNum = Application(Dvbbs.CacheName &"_information_" & node.text).documentElement.selectSingleNode("information/@topicnum").text	'该版块主题数

	If Not IsObject(Conn) Then ConnectionDatabase
	If FoundTopTopic Then
		'查询总固顶、区域固顶、版面固顶帖子
		Dim Forum_AllTopNum
		Forum_AllTopNum=Dvbbs.CacheData(28,0)
		If Trim(Application(Dvbbs.CacheName &"_information_" & Dvbbs.boardid).documentElement.selectSingleNode("information/@boardtopstr").text)<>"" Then
			If Trim(Forum_AllTopNum)<>"" Then
				Forum_AllTopNum = Forum_AllTopNum & "," & Application(Dvbbs.CacheName &"_information_" & Dvbbs.boardid).documentElement.selectSingleNode("information/@boardtopstr").text
			Else
				Forum_AllTopNum = Application(Dvbbs.CacheName &"_information_" & Dvbbs.boardid).documentElement.selectSingleNode("information/@boardtopstr").text
			End If
		End If
		If Trim(Forum_AllTopNum)<>"" Then
			Sql="Select TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,TopicMode,Mode,GetMoney,GetMoneyType,UseTools From Dv_Topic Where istop>0 and TopicID in ("&Forum_AllTopNum&") Order By istop desc, Lastposttime Desc"
			Set Rs = server.CreateObject ("adodb.recordset")
			Rs.Open Sql,Conn,1,1
			TopicNum = Ubound(Split(Forum_AllTopNum,","))
		Else
			Response.Write "<pagecounter>0</pagecounter>"
		End If
	ElseIf Dvbbs.UserID > 0 And Dvbbs.BoardID = -1 Then
		'查询我发表的主题
		Set Rs=Dvbbs.Execute("Select Count(*) From Dv_Topic Where PostUserID = "&Dvbbs.UserID)
		TopicNum = Rs(0)
		If IsNull(TopicNum) Then TopicNum = 0
		Sql="Select TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,TopicMode,Mode,GetMoney,GetMoneyType,UseTools From Dv_Topic Where PostUserID = "&Dvbbs.UserID&" Order By LastPostTime Desc"
		Set Rs = server.CreateObject ("adodb.recordset")
		Rs.Open Sql,Conn,1,1
	ElseIf IsSqlDataBase=1 And IsBuss=1 Then
		Set Cmd = Server.CreateObject("ADODB.Command")
		Set Cmd.ActiveConnection=conn
		Cmd.CommandText="dv_list"
		Cmd.CommandType=4
		Cmd.Parameters.Append cmd.CreateParameter("@boardid",3)
		Cmd.Parameters.Append cmd.CreateParameter("@pagenow",3)
		Cmd.Parameters.Append cmd.CreateParameter("@pagesize",3)
		Cmd.Parameters.Append cmd.CreateParameter("@tl",3)
		Cmd.Parameters.Append cmd.CreateParameter("@topicmode",3)
		Cmd.Parameters.Append cmd.CreateParameter("@totalrec",3,2)
		Cmd("@boardid")=Dvbbs.BoardID
		Cmd("@pagenow")=page
		Cmd("@pagesize")=DvbbsWap.Number
		Cmd("@topicmode")=TopicMode
		If limitime="" Then
			Cmd("@tl")=0
		Else
			Cmd("@tl")=limitime
		End If
		set Rs=Cmd.Execute
	Else
		Sql="Select TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,TopicMode,Mode,GetMoney,GetMoneyType,UseTools From Dv_Topic Where BoardID="&Dvbbs.BoardID&" And IsTop=0 Order By LastPostTime Desc"
		Set Rs = server.CreateObject ("adodb.recordset")
		Rs.Open Sql,Conn,1,1
	End If
	If Not (Rs.Eof And Rs.Bof) Then
		If TopicNum Mod DvbbsWap.Number=0 Then
			n = TopicNum \ DvbbsWap.Number
		Else
	     	n = TopicNum \ DvbbsWap.Number+1
  		End If
		Response.Write "<pagecounter>"&N&"</pagecounter>"
		If IsSqlDatabase = 1 And IsBuss=1 And Dvbbs.BoardID > 0 Then
			SQL=Rs.GetRows(-1)
		Else
			Rs.MoveFirst
			'If page > n Then page = n
			If page > n Then
				page = n
				DvbbsWap.ShowErr 0,"参数错误，请返回！"
				Exit Sub
			End If
			If page < 1 Then page = 1
			If page > 1 Then 				
				Rs.Move (page-1) * DvbbsWap.Number
			End if
			If Rs.Eof Then Exit Sub
			SQL=Rs.GetRows(DvbbsWap.Number)
		End If
		
		For ti=0 To Ubound(SQL,2)
			LastPost = Split(SQL(9,ti),"$")
			If Ubound(LastPost)>=3 Then
				S_Content = "最后回复人："& LastPost(0)
				LastPostTime = LastPost(2)
			Else
				S_Content = ""
				LastPostTime = SQL(7,ti)
			End If
			S_Content = DvbbsWap.Format_Content( 0 , DvbbsWap.ForMatHtml(S_Content) )
			DvbbsWap.ShowCodes DvbbsWap.Self,4,SQL(0,ti) ,3 ,SQL(2,ti) ,S_Content ,"",SQL(3,ti) ,SQL(5,ti) ,LastPostTime
		Next
	Else
		Response.Write "<pagecounter>0</pagecounter>"
	End If
	Rs.Close
	Set Rs=Nothing
	Set Cmd=Nothing
End Sub

'显示帖子内容
Sub ShowDispbbs()
	If FoundErr Then Exit Sub
	Dim Rs,Sql
	SQL="B.AnnounceID,B.BoardID,B.UserName,B.Topic,B.dateandtime,B.body,B.Expression,B.ip,B.RootID,B.signflag,B.isbest,B.PostUserid,B.layer,b.isagree,U.useremail,U.UserIM,U.UserMobile,U.Usersign,U.userclass,U.Usertitle,U.Userwidth,U.Userheight,U.UserPost,U.Userface,U.JoinDate,U.userWealth,U.userEP,U.userCP,U.Userbirthday,U.Usersex,U.UserGroup,U.LockUser,U.userPower,U.titlepic,U.UserGroupID,U.LastLogin,B.PostBuyUser,U.UserHidden,U.IsChallenge,B.Ubblist,B.LockTopic,B.GetMoney,B.UseTools,U.UserMoney,U.UserTicket,B.GetMoneyType"

	Dim AnnounceIDlists
	AnnounceIDlists=AnnounceIDlist()
	If FoundErr Then Exit Sub
	SQL="Select "&SQL&" From "&TotalUseTable&" B Inner Join [dv_user] U On U.UserID=B.PostUserID Where B.RootID="&Announceid&" And B.BoardID="&Dvbbs.BoardID&" And  B.AnnounceID in ("&AnnounceIDlists&") Order BY B.AnnounceID, B.DateAndTime"

	Set Rs = Dvbbs.Execute(SQL)
	If Rs.EOF And Rs.BOF Then
		DvbbsWap.AddErrCode(33)
		Exit Sub
	End If
	Dim Pcount,i
	'Pcount = 0
	If Not(Rs.EOF And Rs.BOF) Then
		If TopicCount mod Cint(DvbbsWap.Number) = 0 Then
			Pcount= TopicCount \ Cint(DvbbsWap.Number)
		Else
			Pcount= TopicCount \ Cint(DvbbsWap.Number)+1
		End If
		'Rs.MoveFirst
		'If star > Pcount Then star = Pcount
		Response.Write "<pagecounter>"&Pcount&"</pagecounter>"
		If star > Pcount Then
				star = Pcount
				DvbbsWap.ShowErr 0,"参数错误，请返回！"
				Exit Sub
		End If
		If star < 1 Then star = 1
		SQL=Rs.GetRows(DvbbsWap.Number)
		Set Rs=Nothing
		
		'AnnounceID=0,BoardID=1,UserName=2,Topic=3,dateandtime=4,body=5,
		'Expression=6,ip=7,RootID=8,signflag=9,isbest=10,PostUserid=11,
		'layer=12,isagree=13,useremail=14,UserIM=15,UserMobile=16,sign=17,
		'userclass=18,title=19,width=20,height=21,article=22,face=23,JoinDate=24,
		'userWealth=25,userEP=26,userCP=27,birthday=28,sex=29,UserGroup=30,LockUser=31,
		'userPower=32,titlepic=33,UserGroupID=34,LastLogin=35,PostBuyUser=36,UserHidden=37,IsChallenge=38,Ubblists=39,LockTopic=40,
		'GetMoney=41,UseTools=42,UserMoney=43,UserTicket=44,GetMoneyType=45
		For i=0 To Ubound(SQL,2)
			'SQL(5,i) = SQL(5,i)&"[code][align=right][b]test[/b][fly]fly[/fly][move]move[/move][color=red]文字文字文字[/color][center]center[/center][/align][rm=500,60,true]http://218.64.81.237/b/i/0326/60534.rm[/rm]aspscript[/code][em10][MP=500,60,true]http://218.64.81.237/b/i/0326/60534.wma[/MP][img]http://218.64.81.237/b/i/0326/60534.gif[/img]"
			SQL(5,i) = DvbbsWap.ForMatHtml(SQL(5,i))
			DvbbsWap.ShowCodes DvbbsWap.Self,"",SQL(0,i) ,4 ,SQL(3,i) ,DvbbsWap.Format_Content(0,SQL(5,i)) ,DvbbsWap.OtherContent,SQL(2,i) ,SQL(4,i) ,SQL(4,i)
		Next

	End If

End Sub


Function Chk_Topic_Err
	Announceid = DvbbsWap.Path(DvbbsWap.PathCount-1)
	If AnnounceID="" Or Not IsNumeric(AnnounceID) Then
		DvbbsWap.AddErrCode(30)
		FoundErr = True
		Exit Function
	End If
	Announceid = Clng(Announceid)
	If AnnounceID = -2 Then
		FoundTopTopic = True
		Exit Function
	End If
	'ReplyID=Request("ReplyID")
	'If ReplyID="" Or Not IsNumeric(ReplyID) Then ReplyID=AnnounceID
	Dim Rs,Sql
	Dvbbs.BoardID = DvbbsWap.Path(DvbbsWap.PathCount-2)
	'Response.Write dvbbs.boardid
	'response.end
	If Dvbbs.BoardID<>"" And IsNumeric(Dvbbs.BoardID) Then
		Dvbbs.BoardID = Clng(Dvbbs.BoardID)
	Else
		DvbbsWap.AddErrCode(29)
		FoundErr = True
		Exit Function
	End If
	Select Case Dvbbs.BoardID
	Case -1
	Case -2
		'Exit Function
	Case Else
		DvbbsWap.LoadBoardPass
	End Select
	Dim MyCanReply,CanRead,CanReply
	'浏览购买帖权限
	CanRead=False
	If Dvbbs.Master or Dvbbs.SuperBoardMaster or Dvbbs.BoardMaster Then CanRead=True
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	SQL="Select title,istop,isbest,PostUserName,PostUserid,hits,isvote,child,pollid,LockTopic,PostTable,BoardID,TopicMode,GetMoney,UseTools,GetMoneyType, DateAndTime From DV_topic where TopicID="&Announceid

	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open SQL,Conn,1,3
	Dvbbs.SqlQueryNum=Dvbbs.SqlQueryNum+1
	'Set Rs=Dvbbs.Execute(SQL)

	Dim istop,isVote,pollid,Locktopic
	Dim TopicMode,ViewNum,T_GetMoney,T_UseTools,T_GetMoneyType

	If Not(Rs.BOF and Rs.EOF) then

		If Dvbbs.BoardID = -1 Or Dvbbs.BoardID = -2 Then
			Dvbbs.BoardID = Rs(11)
			DvbbsWap.LoadBoardPass
		ElseIf Rs(11)<>Dvbbs.BoardID Then
			DvbbsWap.AddErrCode(29)
			FoundErr = True
			Exit Function
		End If

		Rs(5)=Rs(5)+1
		Rs.Update
		Topic=Rs(0)
		istop=rs(1)
		isVote=rs(6)
		TopicCount=rs(7)+1
		pollid=rs(8)
		'锁定多少天前的帖子判断 2004-9-16 Dv.Yz
		If Not Ubound(Dvbbs.Board_Setting) > 70 Then
			Locktopic = Rs(9)
		Else
			If Not Clng(Dvbbs.Board_Setting(71)) = 0 And Datediff("d", Rs(16), Now()) > Clng(Dvbbs.Board_Setting(71)) Then
				Locktopic = 1
			Else
				Locktopic = Rs(9)
			End If
		End If
		TotalUseTable=rs(10)
		TopicMode=rs(12)
		ViewNum=Rs(5)
		T_GetMoney = cCur(Rs(13))
		T_UseTools = Rs(14)
		T_GetMoneyType = Cint(Rs(15))
		
		If Rs(4)=Dvbbs.UserID Then
			MyCanReply=Dvbbs.GroupSetting(4)
			CanRead=True
		Else
			MyCanReply=Dvbbs.GroupSetting(5)
			If Cint(Dvbbs.GroupSetting(2))=0 Then
				DvbbsWap.AddErrcode(31)
				FoundErr = True
				Exit Function
			End If
		End If
		If Len(Topic) > Cint(Dvbbs.Board_Setting(25)) And Not TopicMode>0 Then
			Topic=Left(Topic,Dvbbs.Board_Setting(25))&"..."
		End If
		Topic=Dvbbs.ChkBadWords(Topic)
	Else
		DvbbsWap.AddErrcode(32)
		FoundErr = True
		Exit Function
	End If
	Rs.Close
	Set Rs=Nothing
End Function

Function AnnounceIDlist()
	Dim Rs,SQL,i,starcount
	If Star<1 Then Star=1
	If DvbbsWap.Number = 0 Then DvbbsWap.Number = 10	'每页记录数
	starcount=(Star-1)*DvbbsWap.Number

	SQL="Select Announceid From "&TotalUseTable&" Where BoardID="&Dvbbs.BoardID&" And RootID="&Announceid&" Order By AnnounceID"
	Set Rs=Dvbbs.Execute(SQL)
	If Not Rs.Eof Then
		Rs.Move Starcount
		REM 修正最后页面出错信息 2004-5-22 Dv.Yz
		If Rs.Eof Then
			DvbbsWap.AddErrcode(33)
			FoundErr = True
			Exit Function
		End If
		AnnounceIDlist = Rs(0)
		Rs.Movenext
		For i = 1 To DvbbsWap.Number-1
			If Rs.Eof Then Exit For
			AnnounceIDlist = AnnounceIDlist & "," & Rs(0)
			Rs.Movenext
		Next
	Else
		DvbbsWap.AddErrcode(32)
		FoundErr = True
	End If 
	Set Rs=Nothing
End Function
%>