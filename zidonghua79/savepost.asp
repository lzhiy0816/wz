<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/ubblist.asp"-->
<!--#include file="inc/Email_Cls.asp"-->
<%
Dim parameter
If Dvbbs.BoardID < 1 Then
	Response.Write "参数错误"
	Set Dvbbs=Nothing
	Response.End
End If
Dim MyPost
Dim postbuyuser,bgcolor,abgcolor,FormID
Dvbbs.Loadtemplates("post")
Set MyPost = New Dvbbs_Post
Dvbbs.Stats = MyPost.ActionName
Dvbbs.Nav()
Dvbbs.Head_var 1,Application(Dvbbs.CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"&Dvbbs.BoardID&"']/@depth").text,"",""
MyPost.Save_CheckData
Set MyPost = Nothing
Dvbbs.ActiveOnline
Dvbbs.Footer

Class Dvbbs_Post
	Public Action,ActionName,Star,Page,IsAudit,TotalUseTable,ToAction,TopicMode,Reuser
	Private AnnounceID,ReplyID,ParentID,RootID,Topic,Content,char_changed,signflag,mailflag,iLayer,iOrders
	Private TopTopic,IsTop,LastPost,LastPost_1,UpLoadPic_n,ihaveupfile,smsuserlist,upfileinfo
	Private UserName,UserPassWord,UserPost,GroupID,UserClass,DateAndTime,DateTimeStr,Expression,MyLastPostTime,LastPostTimes
	Private LockTopic,MyLockTopic,MyIsTop,MyIsTopAll,MyTopicMode,Child
	Private CanLockTopic,CanTopTopic,CanTopTopic_a,CanEditPost,Rs,SQL,i,IsAuditcheck
	Private vote,votetype,votenum,votetimeout,voteid,isvote,ErrCodes
	Private GetPostType,ToMoney,UseTools,ToolsBuyUser,GetMoneyType,Tools_UseTools,Tools_LastPostTime,ToolsInfo,ToolsSetting
	Private tMagicFace,iMagicFace,tMagicMoney,tMagicTicket,FoundUseMagic,isAlipayTopic
	Private Sub Class_Initialize()
		ErrCodes = ""
		'管理员及该版版主允许在锁定论坛发帖
		If Dvbbs.Board_Setting(0)="1" And Not (Dvbbs.Master or Dvbbs.Boardmaster) Then
			parameter="showerr.asp?action=lock&boardid="&dvbbs.boardID&"" 
			Set Dvbbs=Nothing
			Response.redirect parameter
		End If
		If Dvbbs.IsReadonly()  And Not Dvbbs.Master Then 
			parameter="showerr.asp?action=readonly&boardid="&dvbbs.boardID&"" 
			Set Dvbbs=Nothing
			Response.redirect parameter
		End If
		Action = Request("Action")
		TotalUseTable = Dvbbs.NowUseBBS
		Select Case Action
		Case "snew"
			Action = 5
			ActionName = template.Strings(1)
			If Dvbbs.GroupSetting(3)="0" Then Dvbbs.AddErrCode(70)
		Case "sre"
			Action = 6
			ActionName = template.Strings(3)
			If Dvbbs.GroupSetting(5)="0" then Dvbbs.AddErrCode(71)
		Case "svote"
			Action = 7
			ActionName = template.Strings(5)
			If Dvbbs.GroupSetting(8)="0" then Dvbbs.AddErrCode(56)
		Case "sedit"
			Action = 8
			ActionName = template.Strings(7)
		Case Else
			Action = 1
			ActionName = template.Strings(0)
		End Select
		Star = Request("star")
		If Star = "" Or Not IsNumeric(Star) Then Star = 1
		Star = Clng(Star)
		Page = Request("page")
		If Page = "" Or Not IsNumeric(Page) Then Page = 1
		Page = Clng(Page)
		'IsAudit = Cint(Dvbbs.Board_Setting(3))
		IsAudit=0
		Reuser = False'此变量标识是否更名发贴
		FoundUseMagic = False
	End Sub
	Public Function inpostlist()
		Dim Rs
		Set Rs=Dvbbs.Execute("Select * From Dv_lastpost Where PostuserID=" & Dvbbs.userID & " order by DateAndtime Desc")
		If Not Rs.EOF Then
			If Datediff("s", Rs("DateAndtime"), Now()) < 10 Then
				inpostlist="<li>您不能在10秒内多次发贴"
			Else
				Do while Not Rs.EOF
				If Rs("PostTitle")=Left(Trim(Request.Form("topic")&"")& Request.Form("body"),150) Then
					inpostlist="<li>您不能重复发贴"
					Exit Do
				End If
			 	Rs.MoveNext
				Loop		
			End If
		Else
			inpostlist=""
		End If
		Set Rs=Nothing 
	End Function
	Public Function CheckFormID(id)
		CheckFormID=false
		Dim i,Str
		For i=1 to Len(id)
			Str=Str & Asc(Mid(id,i,1))-97
		Next
		If Session.SessionID=Str Then
			CheckFormID=True
		End If
	End Function
	'通用判断
	Public Function Chk_Post()
		FormID=Request("Dvbbs")
		If FormID="" Then FormID=Request.Cookies("Dvbbs"):Response.Cookies("Dvbbs")=""
		'If Not CheckFormID(FormID) Then Set Dvbbs=Nothing:Response.redirect "showerr.asp?ErrCodes=您提交的参数错误&action=OtherErr"
		If Dvbbs.Board_Setting(43)="1" Then Dvbbs.AddErrCode(72)
		If Dvbbs.Board_Setting(1)="1" and Dvbbs.GroupSetting(37)="0" Then Dvbbs.AddErrCode(26)
		If Dvbbs.UserID>0 Then
			If Clng(Dvbbs.GroupSetting(52))>0 And DateDiff("s",Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@joindate").text,Now)<Clng(Dvbbs.GroupSetting(52))*60 Then 
				parameter="showerr.asp?ErrCodes=<li>"&Replace(template.Strings(21),"{$timelimited}",Dvbbs.GroupSetting(52))&"&action=OtherErr"
				Set Dvbbs=Nothing
				Response.redirect parameter
			End If
			If Dvbbs.GroupSetting(62)<>"0" And Not Action = 8 Then
				If Clng(Dvbbs.GroupSetting(62))<=Clng(Dvbbs.UserToday(0)) Then 
					parameter="showerr.asp?ErrCodes=<li>"&Replace(template.Strings(27),"{$topiclimited}",Dvbbs.GroupSetting(62))&"&action=OtherErr"
					Set Dvbbs=Nothing
					Response.redirect parameter
				End If
			End If
		End If
		If Dvbbs.GroupSetting(3)="0" And (Action = 5 Or Action = 7) Then Set Dvbbs=Nothing:Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(28)&"&action=OtherErr"
		If Dvbbs.GroupSetting(5)="0" And (Action = 6) Then Set Dvbbs=Nothing:Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(29)&"&action=OtherErr"
	End Function

	'返回判断和参数
	Public Function Get_M_Request()
		AnnounceID = Request("ID")
		If AnnounceID = "" Or Not IsNumeric(AnnounceID) Then Dvbbs.AddErrCode(30)
		Dvbbs.ShowErr()
		AnnounceID = cCur(AnnounceID)
	End Function
	

	'判断发表类型及权限 GetPostType 0=赠送金币贴(求回复答案),1=获赠金币贴,2=金币购买贴
	Private Sub Chk_PostType()
		Dim ToolsID
		ToolsID = Trim(Request.Form("ToolsID"))
		GetPostType = Trim(Request.Form("GetPostType"))
		ToMoney = Trim(Request.Form("ToMoney"))
		If ToMoney="" or Not Isnumeric(ToMoney) Then ToMoney = 0
		If ToolsID="" or Not Isnumeric(ToolsID) Then
			ToolsID = ""
		Else
			ToolsID = Cint(ToolsID)
		End If
		ToMoney = cCur(ToMoney)
		UseTools = ""
		ToolsBuyUser = ""
		GetMoneyType = 0
		If Dvbbs.GroupSetting(59)<>1 Then Exit Sub
		If GetPostType<>"" and (Action = 5 or Action = 7) Then
			Select Case GetPostType
			Case "0"
				If ToMoney = 0 or ToMoney > CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text)  Or ToMoney < 0 Then Set Dvbbs=Nothing:Response.redirect "showerr.asp?ErrCodes=<li>您设置的金币值为空或者多于您拥有的金币数量。&action=OtherErr"
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text  = CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text) -ToMoney
				'UseTools = "-1111"
				ToolsBuyUser = "0|||$SendMoney"
				GetMoneyType = 1
			Case "1"
				ToolsBuyUser = "0|||$GetMoney"
				GetMoneyType = 2
				'UseTools = ToolsInfo(4)
			Case "2"
				If ToMoney = 0 Or ToMoney < 0 Then Set Dvbbs=Nothing:Response.redirect "showerr.asp?ErrCodes=<li>请正确填写购买帖的金币数量。&action=OtherErr"
				Dim Buy_Orders,Buy_VIPType,Buy_UserList
				Buy_Orders = Request.FORM("Buy_Orders")
				Buy_VIPType = Request.FORM("Buy_VIPType")
				Buy_UserList = Request.FORM("Buy_UserList")
				If Buy_Orders<>"" and IsNumeric(Buy_Orders) Then
					Buy_Orders = cCur(Buy_Orders)
				Else
					Buy_Orders = -1
				End If
				If Not IsNumeric(Buy_VIPType) Then Buy_VIPType = 0
				If Buy_UserList<>"" Then Buy_UserList = Replace(Replace(Replace(Buy_UserList,"|||",""),"@@@",""),"$PayMoney","")
				ToolsBuyUser = "0@@@"&Buy_Orders&"@@@"&Buy_VIPType&"@@@"&Buy_UserList&"|||$PayMoney|||"
				GetMoneyType = 3
				'UseTools = ToolsInfo(4)
			End Select
		End If
		'回复获赠金币帖判断
		If Action = 6 and GetPostType = "1" Then
			If ToMoney = 0 or ToMoney > CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text)  Or ToMoney < 0 Then Set Dvbbs=Nothing:Response.redirect "showerr.asp?ErrCodes=<li>您设置的金币值为空或者多于您拥有的金币数量。&action=OtherErr"
		End If
	End Sub

	'更新用户及系统道具数量(用户ID，道具ID，减少数量)
	Private Sub UpdateUserTools(U_UserID,U_ToolsID,n)
		Dim Sql,Rs
		If Clng(n)<0 Then
			n = "+" & n
		Else
			n = "-" & n
		End If
		Set Rs = Dvbbs.Plus_Execute("Select ID From [Dv_Plus_Tools_Buss] Where UserID="& U_UserID &" and ToolsID="& U_ToolsID)
		If Rs.Eof And Rs.Bof Then
			Dim Trs
			Set Trs = Dvbbs.Plus_Execute("Select ToolsName From Dv_Plus_Tools_Info Where ID=" & U_ToolsID)
			If Not (Trs.Eof And Trs.Bof) Then
				Sql = "Insert Into [Dv_Plus_Tools_Buss] (UserID,UserName,ToolsID,ToolsName,ToolsCount) Values ("&U_UserID&",'"&Dv_Tools.ToUserInfo(1)&"',"&U_ToolsID&",'"&Trs(0)&"',"&Clng(Replace(Replace(n,"+",""),"-",""))&")"
				Dvbbs.Plus_Execute(Sql)
			End If
			Trs.Close
			Set Trs=Nothing
		Else
			Sql = "Update [Dv_Plus_Tools_Buss] Set ToolsCount = ToolsCount"&n&" Where UserID="& U_UserID &" and ToolsID="& U_ToolsID
			Dvbbs.Plus_Execute(Sql)
		End If
		Rs.Close
		Set Rs=Nothing
		Sql = "Update [Dv_Plus_Tools_Info] Set UserStock =  UserStock"&n&" Where ID="& U_ToolsID
		Dvbbs.Plus_Execute(Sql)
	End Sub
	'回复获赠金币帖时更新
	Public Sub GetMoney_SaveRe()
		Dim ToUserID,ToUserName,UpPostBuyUser,UpGetMoney,Tempstr
		SQL="Select PostUserID,GetMoney,PostBuyUser,Username From "&TotalUseTable&" Where RootID="&RootID&" and ParentID=0 and GetMoneyType=2"
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.Open SQL,conn,1,3
			If Not Rs.Eof Then
				ToUserID = Clng(Rs(0))
				UpGetMoney = Rs(1) + ToMoney
				UpPostBuyUser = Rs(2)
				ToUserName = Rs(3)
				Tempstr = Split(UpPostBuyUser,"|||",2)
				Tempstr(0) = UpGetMoney
				Tempstr(1) = Tempstr(1) & "|||" & UserName &","& ToMoney
				UpPostBuyUser = Tempstr(0) & "|||" & Tempstr(1)
				UpPostBuyUser = Trim(UpPostBuyUser)
				Rs(1) = UpGetMoney
				Rs(2) = UpPostBuyUser
			Rs.update
			End If
		Rs.Close : Set Rs=Nothing

		'更新主题MONEY值
		Sql = "Update [Dv_Topic] Set GetMoney="&UpGetMoney&" where TopicID="&RootID
		Dvbbs.Execute(sql)
		'主题作者获取金币
		Sql = "Update [Dv_User] Set UserMoney=UserMoney + "&ToMoney&" where UserID="&ToUserID
		Dvbbs.Execute(sql)
		'插入道具日志
		'Dim LogMsg
		'LogMsg = "<a href=""dispbbs.asp?boardid="&Dvbbs.boardid&"&id="&RootID&"&star="&Star&"#"&Announceid&""" target=_blank>回复主题：《<B>"&topic&"</B>》</a> 赠道给作者:<b>"&ToUserName&"</b><font color="&Dvbbs.mainsetting(1)&">"&ToMoney&"</font>个金币成功，并插入道具日志记录。"
		'Call Dvbbs.ToolsLog(0,0,ToMoney,0,0,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text &"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
	End Sub

	Rem ------------------------
	Rem 保存部分函数开始
	Rem ------------------------
	'检查数据,提取数据，获得贴子数据表名等。
	Public Sub Save_CheckData()
		Chk_Post()
		CheckfromScript()
		'把提交的数据保存到session
		Content = CheckAlipay()
		isAlipayTopic = 2
		If Content = "" Then
			Content = Dvbbs.Checkstr(Request.Form("body"))
			isAlipayTopic = 0
		End If
		If InStr(Content,"[/payto]") > 0 And InStr(Content,"[payto]") > 0 And InStr(Content,"(/seller)") > 0 And InStr(Content,"(seller)") > 0 Then isAlipayTopic = 2
		Dvbbs.UserSession.documentElement.selectSingleNode("userinfo").attributes.setNamedItem(Dvbbs.UserSession.createNode(2,"postdata","")).text= Request.Form("body")
		If Dvbbs.Board_Setting(4) = "1" Then
			If Not Dvbbs.CodeIsTrue() Then
				Set Dvbbs=Nothing
				 Response.redirect "showerr.asp?ErrCodes=<li>验证码校验失败，2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
			End If
		End If
		Chk_PostType()

		'魔法表情检查部分
		tMagicFace = Request("tMagicFace")
		If tMagicFace = "" Or Not IsNumeric(tMagicFace) Then tMagicFace = 0
		tMagicFace = Cint(tMagicFace)
		iMagicFace = Request("iMagicFace")
		If iMagicFace = "" Or Not IsNumeric(iMagicFace) Then iMagicFace = 0
		iMagicFace = Clng(iMagicFace)
		Expression = Dvbbs.Checkstr(Request.Form("Expression"))
		If Expression = "" Then
			Expression = "face1.gif"
		Else
			Expression = Replace(Expression,"|","")
		End If
		If tMagicFace = 1 And iMagicFace > 0 And Dvbbs.Forum_Setting(98)="1" Then
			Set Rs = Dvbbs.Plus_Execute("Select tMoney,tTicket,MagicSetting From Dv_Plus_Tools_MagicFace Where MagicFace_s = " & iMagicFace)
			If Rs.Eof And Rs.Bof Then
				Expression = "0|" & Expression
				tMagicMoney = 0
				tMagicTicket = 0
			Else
				If cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text) < Rs(1) And cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text ) < Rs(0) Then Set Dvbbs=Nothing:Response.redirect "showerr.asp?ErrCodes=<li>您没有足够的金币或点券使用魔法表情，2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
				Dim iMagicSetting
				iMagicSetting = Split(Rs(2),"|")
				If cCur(iMagicSetting(0)) > cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text ) Then Set Dvbbs=Nothing:Response.redirect "showerr.asp?ErrCodes=<li>您的帖子数没有达到使用魔法表情的标准，2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
				If cCur(iMagicSetting(1)) > cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text) Then Set Dvbbs=Nothing:Response.redirect "showerr.asp?ErrCodes=<li>您的金钱数没有达到使用魔法表情的标准，2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
				If cCur(iMagicSetting(2)) > cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text) Then Set Dvbbs=Nothing:Response.redirect "showerr.asp?ErrCodes=<li>您的经验数没有达到使用魔法表情的标准，2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
				If cCur(iMagicSetting(3)) > cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text ) Then Set Dvbbs=Nothing:Response.redirect "showerr.asp?ErrCodes=<li>您的魅力数没有达到使用魔法表情的标准，2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
				If cCur(iMagicSetting(4)) > cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpower").text) Then Set Dvbbs=Nothing:Response.redirect "showerr.asp?ErrCodes=<li>您的威望数没有达到使用魔法表情的标准，2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
				Expression = iMagicFace & "|" & Expression
				tMagicMoney = Rs(0)
				tMagicTicket = Rs(1)
				FoundUseMagic = True
			End If
			Rs.Close
			Set Rs=Nothing
		Else
			Expression = "0|" & Expression
		End If
		Expression = Split(Expression,"|")
		Topic = Dvbbs.Checkstr(Trim(Request.Form("topic")))
		signflag = Dvbbs.Checkstr(Trim(Request.Form("signflag")))
		mailflag = Dvbbs.Checkstr(Trim(Request.Form("emailflag")))
		MyTopicMode = Dvbbs.Checkstr(Trim(Request.Form("topicximoo")))
		MyLockTopic = Dvbbs.Checkstr(Trim(Request.Form("locktopic")))
		Myistop = Dvbbs.Checkstr(Trim(Request.Form("istop")))
		Myistopall = Dvbbs.Checkstr(Trim(Request.Form("istopall")))
		TopicMode = Request.Form("topicmode")
		If Dvbbs.strLength(topic)> CLng(Dvbbs.Board_Setting(45)) Then 
			parameter="showerr.asp?ErrCodes=<li>"&Replace(template.Strings(23),"{$topiclimited}",Dvbbs.Board_Setting(45))&"<BR>2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
			Set Dvbbs=Nothing
			Response.redirect parameter
		End If
		Rem 限制提交数据不能大于64K
		If Len(Content) > 64*1024*1024 Then
			parameter="showerr.asp?ErrCodes=<li>您提交的数据过大，提交数据不能大于64K&action=OtherErr&autoreload=1"
			Set Dvbbs=Nothing
			Response.redirect parameter
		End If
		Dim TMPData
		'TMPData=inpostlist
		If TMPData<>"" Then
			If Action <> 8 Then
				parameter="showerr.asp?ErrCodes="&TMPData&"&action=OtherErr&autoreload=1"
				Set Dvbbs=Nothing
				Response.redirect parameter
			End If
		End If
		Rem 老迷增加xhtml格式限制
		Dim XMLPOST,XHTML
		XHTML=True
		If XHTML Then
			Set XMLPOST=Server.CreateObject("msxml2.DOMDocument"& MsxmlVersion)
			If XMLPOST.loadxml("<xhtml>" & xmlencode(Content) &"</xhtml>") Then
				Content=Rexmlencode(Mid(XMLPOST.documentElement.xml,8,Len(XMLPOST.documentElement.xml)-15))
			Else
				parameter="showerr.asp?ErrCodes=<li>您提交的数据不合法(必须提交XHTML格式)&action=OtherErr&autoreload=1"
				Set Dvbbs=Nothing
				Response.redirect parameter
			End If
			Set XMLPOST=Nothing
		End If
		If Dvbbs.strLength(Content) > CLng(Dvbbs.Board_Setting(16)) Then Response.redirect "showerr.asp?ErrCodes=<li>"&Replace(template.Strings(24),"{$bodylimited}",Dvbbs.Board_Setting(16))&"<BR>2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
			REM 2004-4-23添加限制帖子内容最小字节数,下次在模板中添加。Dvbbs.YangZheng
			If Dvbbs.strLength(Content) < CLng(Dvbbs.Board_Setting(52)) And Not CLng(Dvbbs.Board_Setting(52)) = 0 Then
				parameter="showerr.asp?ErrCodes=<li>"&Replace(template.Strings(24),"大于{$bodylimited}","小于"&Dvbbs.Board_Setting(52))&"<BR>2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
				Set Dvbbs=Nothing
			Response.redirect parameter
		End If
		Dim testContent
		testContent=Content
		testContent=Replace(testContent,vbNewLine,"")
		testContent=Replace(testContent," ","")
		testContent=Replace(testContent,"&nbsp;","")		
		testContent=Trim(Dvbbs.Replacehtml(testContent))
		If testContent="" and InStr(Content,"<img")=0 and InStr(Content,"<input")=0 and InStr(Content,"<object")=0  and InStr(Content,"<embed")=0 Then Set Dvbbs=Nothing:Response.redirect "showerr.asp?ErrCodes=<li>您没有填写内容.2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
		If Dvbbs.UserID=0 Then
			mailflag=0:signflag=0
		Else
			If Not IsNumeric(mailflag) Or mailflag="" Then mailflag=0
			mailflag=CInt(mailflag)
			If Not IsNumeric(signflag) Or signflag="" Then signflag=1
			signflag=CInt(signflag)
		End If
		If TopicMode<>"" and IsNumeric(TopicMode) Then TopicMode=Cint(TopicMode) Else TopicMode=0
		If Request.form("upfilerename")<>"" Then
			ihaveupfile=1
			upfileinfo=Replace(Request.form("upfilerename"),"'","")
			upfileinfo=Replace(upfileinfo,";","")
			upfileinfo=Replace(upfileinfo,"--","")
			upfileinfo=Replace(upfileinfo,")","")
			Dim fixid,upfilelen
			fixid=Replace(upfileinfo," ","")
			fixid=Replace(fixid,",","")
			If Not IsNumeric(fixid) Then ihaveupfile=0
			upfilelen=len(upfileinfo)
			upfileinfo=left(upfileinfo,upfilelen-1)
		Else
			ihaveupfile=0
		End If
		voteid=0
		isvote=0
		If Action = 7 Then
			votetype=Dvbbs.Checkstr(request.Form("votetype"))
			If IsNumeric(votetype)=0 or votetype="" Then votetype=0
			vote=Dvbbs.Checkstr(trim(Replace(request.Form("vote"),"|","")))
			Dim j,k,vote_1,votelen,votenumlen
			If vote="" Then
				Dvbbs.AddErrCode(81)
			Else
				vote=split(vote,chr(13)&chr(10))
				j=0
				For i = 0 To ubound(vote)
					If Not (vote(i)="" Or vote(i)=" ") Then
						vote_1=""&vote_1&""&vote(i)&"|"
						j=j+1
					End If
					If i>cint(Dvbbs.Board_Setting(32))-2 Then Exit For
				Next
				For k = 1 to j
					votenum=""&votenum&"0|"
				Next
				votelen=len(vote_1)
				votenumlen=len(votenum)
				votenum=left(votenum,votenumlen-1)
				vote=left(vote_1,votelen-1)
			End If
			If Not IsNumeric(request("votetimeout")) Then
				Dvbbs.AddErrCode(82)
			Else
				If request("votetimeout")="0" Then
					votetimeout=dateadd("d",9999,Now())
				Else
					votetimeout=dateadd("d",CCur(request("votetimeout")),Now())
				End If
				votetimeout=Replace(Replace(CSTR(votetimeout+Dvbbs.Forum_Setting(0)/24),"上午",""),"下午","")
			End If
		End If
		If Action = 5 Or Action = 7 Then
			CanLockTopic=False
			CanTopTopic=False
			CanTopTopic_a=False
			If Topic="" OR Replace(Topic&"","","")="" Then Set Dvbbs=Nothing:Response.redirect "showerr.asp?ErrCodes=<li>您忘记填写标题<BR>2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
			'减少判断，如果不为固顶，锁定等等的操作的话，不检查权限。
			If MyLockTopic="yes" Or Myistopall="yes" Or Myistop="yes" Then
				'判断用户是否有固顶/解除固顶帖子权限
				If (dvbbs.master or dvbbs.superboardmaster or dvbbs.boardmaster) Then 
					If Dvbbs.GroupSetting(21)="1" Then CanTopTopic=True
					If  Dvbbs.GroupSetting(20)="1" Then CanLockTopic=True
				End If
				If Dvbbs.GroupSetting(21)="1" and Cint(Dvbbs.UserGroupID)>3 Then CanTopTopic=True
				If Dvbbs.FoundUserPer and Dvbbs.GroupSetting(21)="1" Then
					CanTopTopic=True
				ElseIf Dvbbs.FoundUserPer and Dvbbs.GroupSetting(21)="0" Then
					CanTopTopic=False
				End If
				'判断用户是否有总固顶帖子权限
				If (dvbbs.master or dvbbs.superboardmaster or dvbbs.boardmaster) and Dvbbs.GroupSetting(38)="1" Then CanTopTopic_a=True
				If Dvbbs.GroupSetting(38)="1" and Cint(Dvbbs.UserGroupID)>3 Then CanTopTopic_a=True
				If Dvbbs.FoundUserPer and Dvbbs.GroupSetting(38)="1" Then
					CanTopTopic_a=True
				ElseIf Dvbbs.FoundUserPer and Dvbbs.GroupSetting(38)="0" Then
					CanTopTopic_a=False
				End If	
			End If
			If MyLockTopic="yes" Then
				MyLockTopic=1
			Else
				MyLockTopic=0
			End If
			If Myistopall="yes" Then
				Myistopall=1
			Else
				Myistopall=0
			End If
			If Not CanTopTopic_a Then Myistopall=0
			If Myistop="yes" and Myistopall=0 Then
				Myistop=1
				If Not CanTopTopic Then Myistop=0
			ElseIf Myistopall=1 Then
				Myistop=3
			Else
				Myistop=0
			End If
			If Not IsNumeric(MyTopicMode) or Dvbbs.GroupSetting(51)="0" Then MyTopicMode=0
			If Not CanLockTopic Then MyLockTopic=0
			TotalUseTable=Dvbbs.NowUseBbs
		ElseIf Action = 6 Then
			AnnounceID = Request("followup")
			RootID = Request("RootID")
			If AnnounceID = "" Or Not IsNumeric(AnnounceID) Then Dvbbs.AddErrCode(30)
			If RootID = "" Or Not IsNumeric(RootID) Then Dvbbs.AddErrCode(30)
			TotalUseTable=Request.Form("TotalUseTable")
			TotalUseTable=checktable(TotalUseTable)
			Dvbbs.ShowErr()
			AnnounceID = cCur(AnnounceID)
			ParentID = AnnounceID
			RootID = cCur(RootID)
			MyLockTopic=0
		ElseIf Action = 8 Then
			If Not IsNumeric(MyTopicMode) or Dvbbs.GroupSetting(51)="0" Then MyTopicMode=0
			AnnounceID = Request("replyID")
			RootID = Request("ID")
			If AnnounceID = "" Or Not IsNumeric(AnnounceID) Then Dvbbs.AddErrCode(30)
			If RootID = "" Or Not IsNumeric(RootID) Then Dvbbs.AddErrCode(30)
			TotalUseTable = Checktable(Request.Form("TotalUseTable"))
			Dvbbs.ShowErr()
			AnnounceID = cCur(AnnounceID)
			RootID = cCur(RootID)
			MyLockTopic=0
			Set Rs=Dvbbs.Execute("select AnnounceID from "&TotalUseTable&" where ParentID=0 and rootid="&RootID)
			If Not Rs.eof Then 
				If AnnounceID=Rs(0) Then
					If Topic="" OR Replace(Topic&"","","")="" Then Dvbbs.AddErrCode(79)
				End If
			Else
				Dvbbs.AddErrCode(30)			
			End If
		End If
		Dvbbs.Showerr()
		SaveData()
	End Sub
	'保存数据
	Private Sub SaveData()
		If Not (Dvbbs.Master or Action = 8) Then CheckpostTime()
		Dim Forumupload
		Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@postdata").text =""
 		If Dvbbs.GroupSetting(64)="1" Then
 			IsAudit=0
 		Else
 			If Dvbbs.Board_Setting(57)="1" Then
 				IsAudit=NeedIsAudit()
 				IsAuditcheck=IsAudit
 			End If
 		End If
		If Dvbbs.UserID=0 Then IsAudit=1:IsAuditcheck=2
 		If IsAudit=0 Then
 			If AccessPost =1 Then
 					IsAudit=1
	 				IsAuditcheck=2
	 		End If
 		End If
 		locktopic=0
		If MyLockTopic=1 Then locktopic=1
		If IsAudit=1 And Action <> 8 Then
			LockTopic=Dvbbs.BoardID
			Dvbbs.BoardID=777
			Response.Cookies("Dvbbs")=LockTopic
		Else
			Response.Cookies("Dvbbs")=Dvbbs.Boardid
		End If
		Forumupload=split(Dvbbs.Board_Setting(19),"|")
		For i=0 to ubound(Forumupload)
			If Instr(Content,"[upload="&Forumupload(i)&"]") or Instr(Content,"."&Forumupload(i)&"") or Instr(Content,"["&Forumupload(i)&"]") then
				uploadpic_n=Forumupload(i)
				Exit For
			End If
		Next
		If InStr(Content,"viewfile.asp?ID=") Then uploadpic_n="down"
		If Not Action = 8 Then 
			savepost()
			If Dvbbs.UserID>0 Then updatepostuser()
		Else
			Update_Edit_Announce()
		End If
		If Dvbbs.GroupSetting(59)=1 Then
			If (Action = 5 Or Action = 7) and (GetPostType="1" or GetPostType="2") Then
				'Dim LogMsg
				'LogMsg = "<a href=""dispbbs.asp?boardid="&Dvbbs.boardid&"&id="&RootID&""" target=_blank>主题：《<B>"&topic&"</B>》</a> 使用<B>"&ToolsInfo(5)&"</B>道具成功，并插入道具日志记录。"
				'Call Dvbbs.ToolsLog(ToolsInfo(4),1,0,0,1,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text &"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
				'Call UpdateUserTools(Dvbbs.UserID,ToolsInfo(4),1)
			ElseIf Action = 6 and GetPostType="1" and GetMoneyType=4 Then
				Call GetMoney_SaveRe
			End If
		End If
		If Action <>8 Then
			'If IsSqlDataBase = 1 Then
			'	Dvbbs.Execute("Delete From [Dv_lastpost] Where Datediff(Mi, DateAndtime, " & SqlNowString & ") > 20")
			'Else
			'	Dvbbs.Execute("Delete From [Dv_lastpost] Where Datediff('s', DateAndtime, " & SqlNowString & ") > 20*60")
			'End If
			'Dvbbs.Execute("Insert Into Dv_lastpost (PostTitle,PostUserID,DateAndtime) values ('"&Dvbbs.Checkstr(Left(Trim(Request.Form("topic")&"")& Request.Form("body"),150))&"',"&Dvbbs.UserID&",'"&Now()&"')")
		End If
		succeed()
	End Sub
	'保存发贴，投票和回贴
	Public Sub Savepost()
		If Action = 5 Or Action = 7 Then
			ilayer=1:iOrders=0:ParentID=0
			MyLastPostTime=Replace(Replace(CSTR(NOW()+Dvbbs.Forum_Setting(0)/24),"上午",""),"下午","")
			DateTimeStr=MyLastPostTime'Replace(Replace(CSTR(NOW()+Dvbbs.Forum_Setting(0)/24),"上午",""),"下午","")
			If Action = 7 Then Insert_To_Vote()
			Insert_To_Topic()
			'更新总固顶和固顶的数据以及缓存数据
			If MyIsTop=3 Then
				'将总固顶ID插入总设置表
				Dim iForum_AllTopNum
				Set Rs=Dvbbs.Execute("Select Forum_AllTopNum From Dv_Setup")
				If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
					iForum_AllTopNum = RootID
				Else
					iForum_AllTopNum = Rs(0) & "," & RootID
				End If
				Dvbbs.Execute("Update Dv_Setup Set Forum_AllTopNum='"&iForum_AllTopNum&"'")
				Dvbbs.ReloadSetupCache iForum_AllTopNum,28
				Set Rs=Nothing
			ElseIf MyIsTop=1 Then
				Dim BoardTopStr
				Set Rs=Dvbbs.Execute("Select BoardID,BoardTopStr From Dv_Board Where BoardID="&Clng(Dvbbs.BoardID))
				If Not (Rs.Eof And Rs.Bof) Then
					If Rs(1)="" Or IsNull(Rs(1)) Then
						BoardTopStr = RootID
					Else
						If InStr(","&Rs(1)&",","," & RootID & ",")>0 Then
							BoardTopStr = Rs(1)
						Else
							BoardTopStr = Rs(1) & "," & RootID
						End If
					End If
					Dvbbs.Execute("Update Dv_Board Set BoardTopStr='"&BoardTopStr&"' Where BoardID="&Rs(0))
					Dvbbs.LoadBoardinformation Rs(0)
				End If
				Set Rs=Nothing
			End If
		ElseIf Action = 6 Then
			Get_SaveRe_TopicInfo()
			Get_ForumTreeCode()
			DateTimeStr=Replace(Replace(CSTR(NOW()+Dvbbs.Forum_Setting(0)/24),"上午",""),"下午","")
		End If
		Insert_To_Announce()
		If Action = 6 Then
			topic=Replace(Replace(cutStr(reubbcode(topic),14),chr(10),""),"'","")
			If topic="" Then
				topic=Content
				topic=Replace(cutStr(reubbcode(topic),14),chr(10),"")
			Else
				topic=Replace(cutStr(reubbcode(topic),14),chr(10),"")
			End If
			If ihaveupfile=1 Then Dvbbs.Execute("update dv_upfile set F_AnnounceID='"&RootID&"|"&AnnounceID&"',F_Readme='"&Replace(toptopic,"'","")&"',F_flag=0 where F_ID in ("&upfileinfo&")")
		Else
			If ihaveupfile=1 Then Dvbbs.Execute("update dv_upfile set F_AnnounceID='"&RootID&"|"&AnnounceID&"',F_Readme='"&Topic&"',F_flag=0  where F_ID in ("&upfileinfo&")")
		End If
		If signflag=2 Then
			LastPost=Replace("匿名用户","$","") & "$" & AnnounceID & "$" & DateTimeStr & "$" & Replace(topic,"$","&#36;") & "$" & uploadpic_n & "$" & " " & "$" & RootID & "$" & Dvbbs.BoardID
		Else
			LastPost=Replace(username,"$","") & "$" & AnnounceID & "$" & DateTimeStr & "$" & Replace(topic,"$","&#36;") & "$" & uploadpic_n & "$" & Dvbbs.UserID & "$" & RootID & "$" & Dvbbs.BoardID
		End If
		LastPost=Replace(LastPost,"'","")
		LastPost=Dvbbs.ChkBadWords(LastPost)
		If IsAudit<>1 Then
			Dim IsUpDateLastPostTime
			IsUpDateLastPostTime = False
			If  InStr("," & Tools_UseTools & ",",",13,")>0 Or InStr("," & Tools_UseTools & ",",",14,")>0 Then IsUpDateLastPostTime = True
			If Action = 6 Then 
				If IsUpDateLastPostTime Then
					If DateDiff("h",Now(),Tools_LastPostTime)<=0 Then
						SQL="update Dv_topic set child=child+1,LastPostTime="&SqlNowString&",LastPost='"&LastPost&"' where TopicID="&RootID
					Else
						SQL="update Dv_topic set child=child+1,LastPost='"&LastPost&"' where TopicID="&RootID
					End If
				Else
					SQL="update Dv_topic set child=child+1,LastPostTime="&SqlNowString&",LastPost='"&LastPost&"' where TopicID="&RootID
				End If
				Dvbbs.Execute(SQL)
				Child=Child+2
				If Child mod Dvbbs.Board_Setting(27)=0 Then 
					star=Child\Dvbbs.Board_Setting(27)
				Else
					star=(Child\Dvbbs.Board_Setting(27))+1
				End If
			Else
				 Dvbbs.Execute("update Dv_topic set LastPost='"&LastPost&"' where topicid="&RootID)
			End If
		Else
			Rem 待审核取消回复更新最后发贴。
			If Action = 5 Or Action = 7 Then  Dvbbs.Execute("update Dv_topic set LastPost='"&LastPost&"' where topicid="&RootID)
		End If
		If Action = 5 Or Action = 7 Then
			toptopic=Replace(cutStr(topic,20),"$","&#36;")
		Else
			toptopic=Replace(cutStr(toptopic,20),"$","&#36;")
		End If
		'toptopic =主标题,Content=内容
		'判断是否是加密的论坛，如果是，则不显示最后发贴内容。
		If Dvbbs.Board_Setting(2)="1" Then
			LastPost_1="保密$" & AnnounceID & "$" & DateTimeStr & "$请认证用户进入查看$" & uploadpic_n & "$" & Dvbbs.UserID & "$" & RootID & "$" & Dvbbs.BoardID
		ElseIf signflag=2 Then
			LastPost_1=Replace("匿名用户","$","") & "$" & AnnounceID & "$" & DateTimeStr & "$" & toptopic & "$" & uploadpic_n & "$" & " " & "$" & RootID & "$" & Dvbbs.BoardID
		Else
			LastPost_1=Replace(username,"$","") & "$" & AnnounceID & "$" & DateTimeStr & "$" & toptopic & "$" & uploadpic_n & "$" & Dvbbs.UserID & "$" & RootID & "$" & Dvbbs.BoardID
		End If
		LastPost_1=reubbcode(Replace(LastPost_1,"'",""))
		LastPost_1=Dvbbs.ChkBadWords(LastPost_1)
		If IsAudit=0 Then
			UpDate_BoardInfoAndCache()
			UpDate_ForumInfoAndCache()
		End If
		Response.Cookies("Dvbbs")=Dvbbs.BoardID
	End Sub
	'判断用户是否有编辑权限且提取相关信息
	Public Function Get_Edit_PermissionInfo()
		Dim old_user
		If Action = 4 Then
			Set Rs=Dvbbs.Execute("select b.username,b.topic,b.body,b.dateandtime,u.UserGroupID,b.signflag,b.emailflag from "&TotalUseTable&" b left outer join [dv_user] u on b.postuserid=u.userid where b.RootID="&AnnounceID&" and b.AnnounceID="&replyID)
		Else
			Set Rs=Dvbbs.Execute("select b.username,b.topic,b.body,b.dateandtime,u.UserGroupID,b.signflag,b.emailflag from "&TotalUseTable&" b left outer join [dv_user] u on b.postuserid=u.userid where b.RootID="&RootID&" and b.AnnounceID="&AnnounceID)
		End If
		If Rs.Eof And Rs.Bof Then
			Dvbbs.AddErrCode(48)
		Else
			If Action = 4 Then
				signflag=Rs("signflag")
				mailflag=Rs("emailflag")
				Topic=rs("topic")
				If Topic<>"" Then Topic = Server.HtmlEncode(Topic)
				Content=rs("body")
				old_user=rs("username")
			Else
				If Clng(Dvbbs.forum_setting(50))>0 then
					If Datediff("s",rs("dateandtime"),Now())>Clng(Dvbbs.forum_setting(50))*60 then
						Content = Content+chr(13)+chr(10)+char_changed+chr(13)
					End If
				Else
					Content = Content+chr(13)+chr(10)+char_changed+chr(13)
				End If
			End If
			If Clng(Dvbbs.forum_setting(51))>0 and not (Dvbbs.master or Dvbbs.boardmaster or Dvbbs.superboardmaster) Then 
				If DateDiff("s",rs("dateandtime"),Now())>Clng(Dvbbs.forum_setting(51))*60 then 
					parameter="showerr.asp?ErrCodes=<li>"&Replace(Replace(template.Strings(22),"{$posttime}",Datediff("s",rs("dateandtime"),Now())/60),"{$etlimited}",Dvbbs.forum_setting(51))&"&action=OtherErr"
					Set Dvbbs=Nothing
					Response.redirect parameter
				End If
			End If 
			If Rs("username")=Dvbbs.membername Then 
				If Dvbbs.GroupSetting(10)="0" then
					Dvbbs.AddErrCode(74)
					CanEditPost=False
				Else 
					CanEditPost=True
				End If 
			Else
				signflag=Rs("signflag")
				mailflag=Rs("emailflag")
				If (Dvbbs.master or Dvbbs.superboardmaster or Dvbbs.boardmaster) and Dvbbs.GroupSetting(23)="1" then
					CanEditPost=True
				Else 
					CanEditPost=False
				End  If 
				If Cint(Dvbbs.UserGroupID) > 3 And Dvbbs.GroupSetting(23)="1" Then CanEditPost=true
				If Dvbbs.GroupSetting(23)="1" and Dvbbs.founduserPer Then 
					CanEditPost=True
				ElseIf Dvbbs.GroupSetting(23)="0" And Dvbbs.founduserPer Then 
					CanEditPost=False
				End If
				If Cint(Dvbbs.UserGroupID) < 4 And Cint(Dvbbs.UserGroupID) = rs("UserGroupID") Then 
					Dvbbs.AddErrCode(75)
				ElseIf Cint(Dvbbs.UserGroupID) < 4 and Cint(Dvbbs.UserGroupID) > rs("UserGroupID") Then
					Dvbbs.AddErrCode(76)
				End If 
				If Not CanEditPost Then Dvbbs.AddErrCode(77)
			End If
		End If
		Set Rs=Nothing
		Dvbbs.ShowErr()
		If Action = 4 Then Dvbbs.MemberName=old_user
	End Function
	Public Sub Get_SaveRe_TopicInfo()
		SQL="select locktopic,LastPost,title,smsuserlist,IsSmsTopic,istop,Child,UseTools,LastPostTime,GetMoneyType,PostUserID, DateAndTime from Dv_topic where BoardID="&Dvbbs.BoardID&" And topicid="&cstr(RootID)
		Set Rs=dvbbs.Execute(sql)
		If Not (Rs.EOF And Rs.BOF) Then
			toptopic = Rs(2)
			istop = Rs(5)
			Child = Rs(6)
			Tools_UseTools = "," & Rs(7) & ","
			Tools_LastPostTime = Rs(8)
			If Rs("IsSmsTopic")=1 Then smsuserlist=Rs("smsuserlist")
			If Rs("LockTopic")=1 And Not (Dvbbs.Master Or Dvbbs.BoardMaster Or Dvbbs.SuperBoardMaster) Then Dvbbs.AddErrCode(78)
			'锁定多少天前的帖子判断 2004-9-16 Dv.Yz
			If Ubound(Dvbbs.Board_Setting) > 70 And Not (Dvbbs.Master Or Dvbbs.BoardMaster Or Dvbbs.SuperBoardMaster) Then
				If Not Clng(Dvbbs.Board_Setting(71)) = 0 And Datediff("d", Rs(11), Now()) > Clng(Dvbbs.Board_Setting(71)) Then Dvbbs.AddErrCode(78)
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
				Set Dvbbs=Nothing
				Response.redirect"showerr.asp?ErrCodes=<br>"+"<li>"&template.Strings(32)&"&action=OtherErr"
				Exit Sub
			End If
			If Clng(Rs(10)) <> Dvbbs.UserID and Cint(Rs(9)) = 2 and GetPostType = "1" Then
				GetMoneyType = 4
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text  = Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text -ToMoney
			Else
				GetMoneyType = 0
				ToMoney=0
			End If
		End If
		Set Rs = Nothing
		Dvbbs.showErr()
	End Sub
	Public Sub Insert_To_Vote()
		'插入投票记录 GroupSetting(68)投票项目中是否可以使用HTML
		Dim UArticle,UWealth,UEP,UCP,UPower
		UArticle = Request("UArticle")
		UWealth = Request("UWealth")
		UEP = Request("UEP")
		UCP = Request("UCP")
		UPower = Request("UPower")
		If Not IsNumeric(UArticle) Then UArticle = 0
		If Not IsNumeric(UWealth) Then UWealth = 0
		If Not IsNumeric(UEP) Then UEP = 0
		If Not IsNumeric(UCP) Then UArticle = 0
		If Not IsNumeric(UPower) Then UPower = 0
		If Dvbbs.GroupSetting(68)<>"1" Then vote = server.htmlencode(vote)
		Dvbbs.execute("insert into dv_vote (vote,votenum,votetype,timeout,UArticle,UWealth,UEP,UCP,UPower) values ('"&vote&"','"&votenum&"',"&votetype&",'"&votetimeout&"',"&UArticle&","&UWealth&","&UEP&","&UCP&","&UPower&")")
		set rs=dvbbs.execute("select Max(voteid) from dv_vote")
		voteid=Rs(0)
		isvote=1
		Set Rs=Nothing
	End Sub
	Public Sub Insert_To_Topic()
		'插入主题表
		Dim hidename
		If signflag=2 Then
			hidename=1
		Else
			hidename=0
		End If
		SQL="insert into Dv_topic (Title,Boardid,PostUsername,PostUserid,DateAndTime,Expression,LastPost,LastPostTime,PostTable,locktopic,istop,TopicMode,isvote,PollID,Mode,GetMoney,UseTools,GetMoneyType,isSmsTopic,HideName) values ('"&topic&"',"&Dvbbs.boardid&",'"&username&"',"&Dvbbs.userid&",'"&DateTimeStr&"','"&Expression(0)&"|"&Expression(1)&"','$$"&DateTimeStr&"$$$$','"&MyLastPostTime&"','"&TotalUseTable&"',"&locktopic&","&Myistop&","&MyTopicMode&","&isvote&","&voteid&","&TopicMode&","&ToMoney&",'"&UseTools&"',"&GetMoneyType&","&isAlipayTopic&","&hidename&")"
		Dvbbs.Execute(sql)
		Set Rs=Dvbbs.Execute("select Max(topicid) From Dv_topic Where PostUserid="&Dvbbs.UserID)
		RootID=Rs(0)
	End Sub
	Public Sub Insert_To_Announce()
		'插入回复表
		DIM UbblistBody
		UbblistBody = Content
		UbblistBody = Ubblist(Content)
		SQL="insert into "&TotalUseTable&"(Boardid,ParentID,username,topic,body,DateAndTime,length,RootID,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID,isupload,IsAudit,Ubblist,GetMoney,UseTools,PostBuyUser,GetMoneyType) values ("&Dvbbs.boardid&","&ParentID&",'"&username&"','"&topic&"','"&Content&"','"&DateTimeStr&"','"&Dvbbs.strlength(Content)&"',"&RootID&","&ilayer&","&iorders&",'"&Dvbbs.UserTrueIP&"','"&Expression(1)&"',"&locktopic&","&signflag&","&mailflag&",0,"&Dvbbs.userid&","&ihaveupfile&","&IsAudit&",'"&UbblistBody&"',"&ToMoney&",'"&UseTools&"','"&dvbbs.checkstr(ToolsBuyUser)&"',"&GetMoneyType&")"
		Dvbbs.Execute(sql)
		Set Rs=Dvbbs.Execute("select Max(AnnounceID) From "&TotalUseTable&" Where PostUserID="&Dvbbs.UserID)
		AnnounceID=Rs(0)
	End Sub
	'检查贴中是否含过滤字
	Public Function NeedIsAudit()
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
	Public Sub Get_ForumTreeCode()
		Dim mailbody
		Set Rs=Dvbbs.Execute("select b.layer,b.orders,b.EmailFlag,b.username,u.userEmail from "&TotalUseTable&" b inner join [Dv_user] u on b.postuserid=u.userid where b.AnnounceID="&ParentID)
		If Not(rs.EOF And rs.BOF) Then
			If IsNull(Rs(0)) Then
				iLayer=1
			Else
				iLayer=Rs(0)
			End If
			If IsNull(Rs(1)) Then
				iOrders=0
			Else
				iOrders=Rs(1)
			End If
			If Rs(3)=Dvbbs.membername Then
				If Cint(Dvbbs.GroupSetting(4))=0 Then Dvbbs.AddErrCode(73)
			Else
				If Rs(2)>0 Then
					Dim sUrl,Email,TempArray,etopic
					TempArray = Split(template.html(10),"||")
					sUrl=Dvbbs.Get_ScriptNameUrl
					If Rs(2)=1 Or Rs(2)=3 Then
						etopic=Replace(template.Strings(25),"{$forumname}",Dvbbs.Forum_info(0))
						email=Rs(4)
						mailbody = TempArray(0)
						mailbody = Replace(mailbody,"{$boardid}",Dvbbs.BoardID)
						mailbody = Replace(mailbody,"{$forumname}",Dvbbs.Forum_info(0))
						mailbody = Replace(mailbody,"{$topicid}",RootID)
						mailbody = Replace(mailbody,"{$star}",Star)
						mailbody = Replace(mailbody,"{$surl}",sUrl)
						mailbody = Replace(mailbody,"{$parentid}",ParentID)
						Dim DvEmail
						Set DvEmail = New Dv_SendMail
						DvEmail.SendObject = Cint(Dvbbs.Forum_Setting(2))	'设置选取组件 1=Jmail,2=Cdonts,3=Aspemail
						DvEmail.ServerLoginName = Dvbbs.Forum_info(12)	'您的邮件服务器登录名
						DvEmail.ServerLoginPass = Dvbbs.Forum_info(13)	'登录密码
						DvEmail.SendSMTP = Dvbbs.Forum_info(4)			'SMTP地址
						DvEmail.SendFromEmail = Dvbbs.Forum_info(5)		'发送来源地址
						DvEmail.SendFromName = Dvbbs.Forum_info(0)		'发送人信息
						If DvEmail.ErrCode = 0 Then
							DvEmail.SendMail email,etopic,mailbody	'执行发送邮件
						End If
						Set DvEmail = Nothing
					End If
					If Rs(2)=2 Or Rs(2)=3 Then
						mailbody = TempArray(1)
						mailbody = Replace(mailbody,"{$boardid}",Dvbbs.BoardID)
						mailbody = Replace(mailbody,"{$topicid}",RootID)
						mailbody = Replace(mailbody,"{$star}",Star)
						mailbody = Replace(mailbody,"{$surl}",sUrl)
						mailbody = Replace(mailbody,"{$parentid}",ParentID)
						Dvbbs.Execute("insert into dv_message(incept,sender,title,content,sendtime,flag,issend) values('"&Rs(3)&"','"&Dvbbs.Forum_info(0)&"','"&template.Strings(26)&"','"&mailbody&"',"&SqlNowString&",0,1)")
						update_user_msg(Rs(3))
					End If
				End If
			End If
		Else
			iLayer=1
			iOrders=0
		End If
		Set Rs=Nothing
		If RootID<>0 Then 
			iLayer=ilayer+1
			Dvbbs.Execute "update "&TotalUseTable&" set orders=orders+1 where RootID="&cstr(RootID)&" and orders>"&cstr(iOrders)
			iOrders=iOrders+1
		End If
	End Sub
	Public Sub Update_Edit_Announce()
		Dim re,LastBoard,LastTopic
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="<br>"
		Content = re.Replace(Content,"[br]")
		re.Pattern="\[align=right\]\[color=#000066\](.*)\[\/color\]\[\/align\]"
		Content = re.Replace(Content,"")
		re.Pattern="<div align=right><font color=#000066>(.*)<\/font><\/div>"
		Content = re.Replace(Content,"")
		re.Pattern="\[br\]"
		Content = re.Replace(Content,"<br>")		
		Set re=Nothing
		If Dvbbs.membername<>UserName Then 
			If Dvbbs.forum_setting(49)="1" Then char_changed = "[align=right][color=#000066][此贴子已经被"&Dvbbs.membername&"于"&Now()&"编辑过][/color][/align]"
		Else
			If Dvbbs.forum_setting(48)="1" Then char_changed = "[align=right][color=#000066][此贴子已经被作者于"&Now()&"编辑过][/color][/align]"
		End If
		Get_Edit_PermissionInfo
		Dim Contentdata,LastPost_l
		Contentdata=Content
		Dvbbs.ShowErr
		'取出当前版面最后回复id,如果本帖为最后回复则更新相应数据
		Set Rs = Dvbbs.Execute("select LastPost from dv_board where boardid="&Dvbbs.BoardID)
		If not (Rs.EOF And Rs.BOF) Then
			If Not IsNull(rs(0)) And rs(0)<>"" then
				LastBoard=split(rs(0),"$")
				If ubound(LastBoard)=7 Then
					If cCur(LastBoard(1))=cCur(AnnounceID) Then
						If signflag=2 Then LastBoard(0)="匿名用户":LastBoard(5)=""
						LastPost=LastBoard(0) & "$" & LastBoard(1) & "$" & Now() & "$" & LastBoard(3) & "$" & LastBoard(4) & "$" & LastBoard(5) & "$" & LastBoard(6) & "$" & Dvbbs.BoardID
						dvbbs.execute("update dv_board set LastPost='"&Replace(Replace(LastPost,"'",""),"<","&lt;")&"' where boardid="&Dvbbs.BoardID)
						UpDate_BoardInfoAndCache()
					End If
				End If
			End If
		End If

		'取得当前主题最后回复id,如果本帖为最后回复则更新相应数据
		Dim iFoundUseMagic,MyMagicFace
		Set Rs=Dvbbs.Execute("select LastPost,istop,Expression from dv_topic where topicid="&rootid)
		If Not (Rs.Eof And Rs.Bof) Then
			istop=rs(1)
			MyMagicFace = Split(Rs(2),"|")
			If Ubound(MyMagicFace) = 1 Then iFoundUseMagic = MyMagicFace(0)
			If Not Isnull(Rs(0)) And Rs(0)<>"" Then
				LastTopic=split(rs(0),"$")
				If Ubound(LastTopic)=7 Then
					If cCur(LastTopic(1))=cCur(Announceid) Then
						If signflag=2 Then LastTopic(0)="匿名用户":LastTopic(5)=""
						LastPost=LastTopic(0) & "$" & LastTopic(1) & "$" & Now() & "$" & Replace(cutStr(reubbcode(Contentdata),20),"$","&#36;") & "$" & LastTopic(4) & "$" & LastTopic(5) & "$" & LastTopic(6) & "$" & Dvbbs.BoardID
						dvbbs.execute("update dv_topic set LastPost='"&Replace(LastPost,"'","")&"' where topicid="&rootid)
					End If
				End If
			End If
		End If

		Set Rs = Server.CreateObject("ADODB.Recordset")
		SQL="SELECT * FROM "&TotalUseTable&" where AnnounceID="&Announceid&" And username='"&username&"'"
		rs.Open SQL,conn,1,3
		If Rs.EOF And Rs.BOF Then
			If Not CanEditPost Then Dvbbs.AddErrCode(77)
		ElseIf Not Dvbbs.master And rs("locktopic")=1 then
			Dvbbs.AddErrCode(78)
		Else
			If Rs("parentid")=0 then
				Dim iExpression
				iExpression = ",Expression='"&Expression(0)&"|"&Expression(1)&"'"
				If signflag=2 Then
					Dvbbs.Execute("update dv_topic set title='"&topic&"',TopicMode="&MyTopicMode&",isSmsTopic="&isAlipayTopic&" "&iExpression&",HideName=1 where topicid="&rootid)
				Else
					Dvbbs.Execute("update dv_topic set title='"&topic&"',TopicMode="&MyTopicMode&",isSmsTopic="&isAlipayTopic&" "&iExpression&" where topicid="&rootid)
				End If
			End If
			If IsAudit = 1 Then
				Rs("locktopic")=3
			ElseIf Rs("locktopic")=3 And IsAudit = 0 Then
				Rs("locktopic")=0
			End If
			Rs("Topic") =Replace(Topic,"''","'")
			rs("Body") =Replace(Content,"''","'")
			rs("length")=Dvbbs.strlength(Content)
			rs("ip")=Dvbbs.UserTrueIP
			rs("Expression")=Expression(1)
			rs("signflag")=signflag
			rs("emailflag")=mailflag
			If Rs("isupload")=0 And ihaveupfile=1 Then Rs("isupload")=1
			Dim UbblistBody
			UbblistBody = Ubblist(Content)
			Rs("Ubblist")=UbblistBody
			rs.Update
			If ihaveupfile=1 Then dvbbs.execute("update dv_upfile set F_AnnounceID='"&rootid&"|"&AnnounceID&"',F_Readme='"&Replace(Rs("Topic"),"'","''")&"',F_flag=0 where F_ID in ("&upfileinfo&")")
			If FoundUseMagic Then
				If iMagicFace <> Clng(iFoundUseMagic) Then
					If cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text ) > tMagicMoney Then
						Dvbbs.Execute("Update Dv_User Set UserMoney=UserMoney-"&tMagicMoney&" Where UserID="&Dvbbs.UserID)
						Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text =Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text -tMagicMoney
						Dvbbs.ToolsLog -88,1,tMagicMoney,0,1,"使用金币购买魔法表情",Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text  & "|" & Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text
					Else
						Dvbbs.Execute("Update Dv_User Set UserTicket=UserTicket-"&tMagicTicket&" Where UserID="&Dvbbs.UserID)
						Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text-tMagicTicket
						Dvbbs.ToolsLog -88,1,0,tMagicTicket,1,"使用点券购买魔法表情",Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text  & "|" & Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text
					End If
				End If
			End If
		End If	
		Rs.Close
		Set Rs=Nothing
		Dvbbs.ShowErr()
	End Sub
	'更新版面数据和缓存
	Public Sub UpDate_BoardInfoAndCache()
		Dim UpdateBoardID,parentstr
		parentstr =Application(Dvbbs.CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"&Dvbbs.BoardID&"']/@parentstr").text
		If parentstr <> "0" Then 
			UpdateBoardID= parentstr & "," & Dvbbs.BoardID
		Else
			UpdateBoardID=Dvbbs.BoardID
		End If
		Dim updateboard,i
		LastPost_1=replace(LastPost_1,"<","&lt;")
		If Action = 6 Then
			SQL="update Dv_board set PostNum=PostNum+1,todaynum=todaynum+1,LastPost='"&LastPost_1&"' where boardid in ("&UpdateBoardID&")"
		ElseIf Action = 5 Or Action = 7 Then
			SQL="update Dv_board set PostNum=PostNum+1,TopicNum=TopicNum+1,todaynum=todaynum+1,LastPost='"&LastPost_1&"' where boardid in ("&UpdateBoardID&")"
		ElseIf Action = 8 Then
			LastPost_1=replace(LastPost,"<","&lt;")
		End If
		Dvbbs.Execute(sql)
		UpdateBoardID=Split(UpdateBoardID,",")
		LastPost_1=Split(LastPost_1,"$")
		For Each updateboard in UpdateBoardID
			If IsObject(Application(Dvbbs.CacheName &"_information_" & updateboard)) Then
				If Not Action = 8 Then
					Application(Dvbbs.CacheName &"_information_" & updateboard).documentElement.selectSingleNode("information/@postnum").text=CLng(Application(Dvbbs.CacheName &"_information_" & updateboard).documentElement.selectSingleNode("information/@postnum").text)+1
					Application(Dvbbs.CacheName &"_information_" & updateboard).documentElement.selectSingleNode("information/@todaynum").text=CLng(Application(Dvbbs.CacheName &"_information_" & updateboard).documentElement.selectSingleNode("information/@todaynum").text)+1
				End If
				If Action = 5 Or Action = 7 Then
					Application(Dvbbs.CacheName &"_information_" & updateboard).documentElement.selectSingleNode("information/@topicnum").text=CLng(Application(Dvbbs.CacheName &"_information_" & updateboard).documentElement.selectSingleNode("information/@topicnum").text)+1
				End If
				For i=0 to UBound(LastPost_1)
					Application(Dvbbs.CacheName &"_information_" & updateboard).documentElement.selectSingleNode("information/@lastpost_"& i &"").text=LastPost_1(i)
				Next
			End If
		Next
	End Sub
	Public Sub UpDate_ForumInfoAndCache()
		Dim updateinfo,LastPostTime
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
			If Action = 6 Then
				SQL="update Dv_setup set Forum_PostNum=Forum_PostNum+1,Forum_TodayNum=Forum_TodayNum+1,Forum_LastPost='"&LastPost&"' "&updateinfo&" "
			Else
				SQL="update Dv_setup set Forum_TopicNum=Forum_TopicNum+1,Forum_PostNum=Forum_PostNum+1,Forum_TodayNum=Forum_TodayNum+1,Forum_LastPost='"&LastPost&"' "&updateinfo&" "
			End If
		Else
			If Action = 6 Then
				SQL="update Dv_setup set Forum_PostNum=Forum_PostNum+1,forum_YesTerdayNum="&CLng(Forum_TodayNum)&",Forum_TodayNum=1,Forum_LastPost='"&LastPost&"' "
			Else
				SQL="update Dv_setup set Forum_TopicNum=Forum_TopicNum+1,Forum_PostNum=Forum_PostNum+1,forum_YesTerdayNum="&CLng(Forum_TodayNum)&",Forum_TodayNum=1,Forum_LastPost='"&LastPost&"' "
			End If
			Dvbbs.ReloadSetupCache 1,9
		End If
		'更新总固顶部分数据和缓存
		If Not Action = 6 Then
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
	Public Sub succeed()
		Dim TempStr,PostRetrunName,tourl,returnurl
		If IsAudit=1 And Action <> 8 Then Dvbbs.BoardID=LockTopic
		Dvbbs.Stats = Dvbbs.Stats & template.Strings(20)
		TempStr = template.html(8)
		Select case Dvbbs.Board_Setting(17)
		case "1"
			tourl = "index.asp"
			PostRetrunName=template.Strings(13)
		case "2"
			tourl="index.asp?boardid="&Dvbbs.boardid
			PostRetrunName=template.Strings(14)
		case "3"
			If IsAudit=1  Then
				tourl="index.asp?boardid="&Dvbbs.boardid
				If IsAuditcheck=1 Then
					PostRetrunName="由于您发表的贴子含敏感内容，您的贴子需要管理员审核过才可以见到。"
				ElseIf IsAuditcheck=2 Then
					PostRetrunName="由于您尚未达到豁免审核的条件，您的贴子需要管理员审核过才可以见到。" 
				Else
					PostRetrunName=template.Strings(19)
				End If 
			Else
				Select Case Action
				Case 5
				tourl="dispbbs.asp?boardid="&Dvbbs.boardid&"&id="&RootID
				PostRetrunName=template.Strings(15)
				Case 6
				tourl="dispbbs.asp?boardid="&Dvbbs.boardid&"&id="&RootID&"&star="&Star&"#"&Announceid
				PostRetrunName=template.Strings(16)
				Case 7
				tourl="dispbbs.asp?boardid="&Dvbbs.boardid&"&id="&RootID
				PostRetrunName=template.Strings(17)
				Case 8
				tourl="dispbbs.asp?boardid="&Dvbbs.boardid&"&id="&RootID&"&star="&Star&"#"&RootID
				PostRetrunName=template.Strings(18)
				End Select
			End If
		End Select
		returnurl="dispbbs.asp?boardid="&Dvbbs.boardid&"&id="&RootID
		TempStr = Replace(TempStr,"{$tourl}",tourl)
		TempStr = Replace(TempStr,"{$returnurl}",returnurl)
		TempStr = Replace(TempStr,"{$stats}",Dvbbs.Stats)
		TempStr = Replace(TempStr,"{$boardname}",Dvbbs.BoardType)
		TempStr = Replace(TempStr,"{$boardid}",Dvbbs.BoardID)
		TempStr = Replace(TempStr,"{$page}",page)
		TempStr = Replace(TempStr,"{$PostRetrunName}",PostRetrunName)
		Response.Write TempStr
	End Sub
	Private Function checktable(Table)
		Table=Right(Trim(Table),2)
		If Not IsNumeric(table) Then Table=Right(Trim(Table),1)
		If Not IsNumeric(table) Then Dvbbs.AddErrCode(30)
		checktable="Dv_bbs"&table
	End Function
	'检查提交来源
	Public Sub CheckfromScript()
		If Not Dvbbs.ChkPost() Then Dvbbs.AddErrCode(16):Dvbbs.Showerr()
 		If CStr(Request.Cookies("Dvbbs"))=CStr(Dvbbs.Boardid) Then Dvbbs.AddErrCode(30):Dvbbs.Showerr()
 		If (Not ChkUserLogin) And (Action = 5 Or Action = 6 Or Action = 7) And Dvbbs.UserID>0 Then Dvbbs.AddErrCode(12):Dvbbs.Showerr()	
	End Sub
	'判断发贴时间间隔
	Private Sub CheckpostTime()
		If Dvbbs.Board_Setting(30)="1"  Then 
			If IsDate(Session(Dvbbs.CacheName & "posttime"))  Then
				If DateDiff("s",Session(Dvbbs.CacheName & "posttime"),Now())<CLng(Dvbbs.Board_Setting(31)) Then
					template.Strings(33) = Replace(template.Strings(33),"{$PostTimes}",Dvbbs.Board_Setting(31))
					 parameter="showerr.asp?ErrCodes=<Br>"+"<li>"&template.Strings(33)&"&action=OtherErr"
					Set Dvbbs=Nothing
					Response.redirect parameter
				End If
			End If
			Session(Dvbbs.CacheName & "posttime")=Now()
		End If
	End Sub 
	'检查用户身份
	Public Function ChkUserLogin()
 		ChkUserLogin=False
 		'取得发贴用户名和密码
		If Dvbbs.UserID=0 Then
			UserName="客人"
		Else
			UserName=Dvbbs.Checkstr(Request.Form("username"))
		End If
		'校验用户名和密码是否合法
		If UserName="" Or Dvbbs.strLength(userName)>Cint(Dvbbs.Forum_setting(41)) Or Dvbbs.strLength(userName) < Cint(Dvbbs.Forum_setting(40)) Then Dvbbs.AddErrCode(17)
		If Not IstrueName(UserName) Then Dvbbs.AddErrCode(18)
		Dvbbs.ShowErr()
		If Action = 8 Then
			'编辑贴子，检查用户身份
			UserPassWord=Dvbbs.checkStr(Trim(Request.Cookies(Dvbbs.Forum_sn)("password")))
			SQL = "Select JoinDate,UserID,UserPost,UserGroupID,userclass,lockuser,TruePassWord From [Dv_User] Where UserID="&Dvbbs.UserID
		Else
			'检查用户是否当前用户
			If UserName<>Dvbbs.MemberName Then
				Reuser=True
				UserPassWord=Dvbbs.Checkstr(Trim(Request.Form("passwd")))
				UserPassWord=md5(UserPassWord,16)
				SQL = "Select JoinDate,UserID,UserPost,UserGroupID,userclass,lockuser,userpassword From [Dv_User] Where UserName='"&UserName&"' "
			Else
				UserPassWord=Dvbbs.checkStr(Trim(Request.Cookies(Dvbbs.Forum_sn)("password")))
				SQL = "Select JoinDate,UserID,UserPost,UserGroupID,userclass,lockuser,TruePassWord From [Dv_User] Where UserID="&Dvbbs.UserID		
			End If
		End If
		If Len(UserPassWord)<>16 AND Len(UserPassWord)<>32 Then Dvbbs.AddErrCode(18)
 		Set Rs=Dvbbs.Execute(SQL)
 		If Not Rs.EOF Then
			If Not (UserPassWord<>rs(6) Or rs(5)=1 or rs(3)=5) Then
				'不允许使用马甲
				If Dvbbs.UserID<>Rs(1) Then
					ChkUserLogin = False
				Else
					Dvbbs.UserID=Rs(1)	
 					UserPost=Rs(2)
 					GroupID=Rs(3)
 					userclass=Rs(4)
 					ChkUserLogin=True
				End If
				Response.cookies("upNum")=0
 			Else
  				Dvbbs.LetGuestSession()			
			End If
 		End If
 		Set Rs = Nothing
 	End Function
 	'更新用户积分，所需外部变量,UserPost,userid,（外加发贴回贴的积分设置数据）
	Public Sub updatepostuser()
		'投票，发贴，更新积分
		'更新最后发贴时间
		Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@activetime").text = Now()
		If IsAudit> 0 Then Exit Sub
		Dim MagicSql
		If Action = 5 Or Action = 7 Then 
			If FoundUseMagic Then
				If cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text ) > tMagicMoney Then
					Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text =Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text -tMagicMoney
					Dvbbs.ToolsLog -88,1,tMagicMoney,0,1,"使用金币购买魔法表情",Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text  & "|" & Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text
				Else
					MagicSql = ",UserTicket=UserTicket-"&tMagicTicket&""
					Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text-tMagicTicket
					Dvbbs.ToolsLog -88,1,0,tMagicTicket,1,"使用点券购买魔法表情",Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text  & "|" & Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text
				End If
			End If
			Dvbbs.Execute("update [Dv_user] set UserLastIP='"&Dvbbs.usertrueip&"',UserPost=UserPost+1,UserTopic=UserTopic+1,userWealth=userWealth+"&Clng(Dvbbs.Forum_user(1))&",userEP=userEP+"&Clng(Dvbbs.Forum_user(6))&",userCP=userCP+"&Clng(Dvbbs.Forum_user(11))&",UserToday='"&Clng(Dvbbs.UserToday(0))+1&"|"&Clng(Dvbbs.UserToday(1))&"|"&Clng(Dvbbs.UserToday(2))&"|"&Clng(Dvbbs.UserToday(3))&"|"&Clng(Dvbbs.UserToday(4))&"',UserMoney="&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text &" "&MagicSql&" Where UserID="&Dvbbs.userID)
			If Not Reuser Then
				UserPost=UserPost+1
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usertopic").text =CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usertopic").text)+1
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text+Clng(Dvbbs.Forum_user(1))
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text+Clng(Dvbbs.Forum_user(6))
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text =Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text +Clng(Dvbbs.Forum_user(11))
			End If
		ElseIf Action = 6 Then '回贴更新积分。
			If Not Reuser Then
				Dvbbs.Execute("update [Dv_user] set UserLastIP='"&Dvbbs.usertrueip&"',UserPost=UserPost+1,userWealth=userWealth+"&Clng(Dvbbs.Forum_user(2))&",userEP=userEP+"&Clng(Dvbbs.Forum_user(7))&",userCP=userCP+"&Clng(Dvbbs.Forum_user(12))&",UserToday='"&Clng(Dvbbs.UserToday(0))+1&"|"&Clng(Dvbbs.UserToday(1))&"|"&Clng(Dvbbs.UserToday(2))&"|"&Clng(Dvbbs.UserToday(3))&"|"&Clng(Dvbbs.UserToday(4))&"',UserMoney="&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text &" Where UserID="&Dvbbs.userID)
				UserPost=UserPost+1
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text+Clng(Dvbbs.Forum_user(2))
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text+Clng(Dvbbs.Forum_user(7))
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text =Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text +Clng(Dvbbs.Forum_user(12))
			Else
				Dvbbs.Execute("update [Dv_user] set UserLastIP='"&Dvbbs.usertrueip&"',UserPost=UserPost+1,userWealth=userWealth+"&Clng(Dvbbs.Forum_user(2))&",userEP=userEP+"&Clng(Dvbbs.Forum_user(7))&",userCP=userCP+"&Clng(Dvbbs.Forum_user(12))&",UserMoney="&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text &" Where UserID="&Dvbbs.userID)
			End If
		End If
		If Not Reuser Then
			Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text =CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text)+1
			Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usertoday").text=Clng(Dvbbs.UserToday(0))+1 & "|" & Clng(Dvbbs.UserToday(1)) & "|" & Clng(Dvbbs.UserToday(2))&"|"&Clng(Dvbbs.UserToday(3))&"|"&Clng(Dvbbs.UserToday(4))
		End If
		'发贴数字能整除十则更新用户等级。(Updategrade())
		If UserPost mod 10 < 1  Then Updategrade()
	End Sub
	'更新用户等级，所需外部变量,UserPost,GroupID,userid
	Public Sub Updategrade()
		Dim FoundGrade,TitlePic
		REM 判断用户组(等级)资料，当用户级别为跟随文章数增长则自动更新用户组(等级)
		REM 自动更新用户数据
		REM 如果属于系统或特殊或多属性组
		Set Rs=Dvbbs.Execute("Select MinArticle,IsSetting,ParentGID,UserTitle,GroupPic From Dv_UserGroups Where UserGroupID="&GroupID)
		If Not (Rs.Eof And Rs.Bof) Then
			If Rs(2)<>3 Then
				'用户等级不按照文章升级，用户为系统或特殊或多属性组
				UserClass=Rs(3)
				TitlePic=Rs(4)
				FoundGrade=True
			End If
		End If
		If Not FoundGrade Then
			'如果不属于系统或特殊或多属性组,则将该用户属于注册用户组且按照其文章数自动更新其用户组(等级)
			Set Rs=Dvbbs.Execute("Select Top 1 usertitle,GroupPic,UserGroupID From Dv_UserGroups Where ParentGID=3 And Minarticle<="&UserPost&" Order By MinArticle Desc,UserGroupID")
			If Not (Rs.Eof And Rs.Bof) Then
				UserClass=Rs(0)
				TitlePic=Rs(1)
				GroupID=Rs(2)
				FoundGrade=True
			End If
		End If
		Set Rs=Nothing
		If Not FoundGrade Then Set Dvbbs=Nothing:Response.Redirect "showerr.asp?ErrCodes=<li>系统没有找到您的注册用户组资料，请联系管理员进行修正。&action=OtherErr"
		Dvbbs.Execute("Update [Dv_User] Set UserClass='"&UserClass&"',TitlePic='"&TitlePic&"',UserGroupID="&GroupID&" Where UserID="&Dvbbs.UserID)
		'If Not Reuser Then
			Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userclass").text = UserClass
			Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usergroupid").text = GroupID
		'End If
	End Sub
	Function AccessPost()
		Dim Dom,node,ischeckcontent
		If Cint(Dvbbs.Board_Setting(3))=0 Then
			AccessPost=0
			Exit Function
		End If
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
		If Node is Nothing Then
		 	AccessPost=0
			Exit Function
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
		 	Case "5"
		 		tmpstr="setting[@use=1]/check[@new=1]"
		 		laction="new"
		 	Case "6"
		 			tmpstr="setting[@use=1]/check[@re=1]"
		 			laction="re"
		 	Case "7"
		 		tmpstr="setting[@use=1]/check[@new=1]"
		 		laction="new"
		 	Case "8"
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
		For Each node in Dom.documentElement.selectNodes("setting[@use=1 and check/@"&laction&"=1]/checkuser[@use=1 and @usergroupid="&dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usergroupid").text&"]")
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
		s=Request("topic")&Request("Body")
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		For Each Node in node1
			
			'含图片检查
			If Node.selectSingleNode("checkpic/@use").text="1" Then
				re.Pattern="(\[IMG\]|<img)"
				If re.Test(s) Then
					checkcontent=1
					set re=Nothing
					Exit For
				End If
			End If
			'含链接检查
			If Node.selectSingleNode("checklink/@use").text="1" Then
				re.Pattern="(\[url\]|<a |http://|https://)"
				If re.Test(s) Then
					checkcontent=1
					set re=Nothing
					Exit For
				End If
			End If
			'含flash检查
			If Node.selectSingleNode("checkflash/@use").text="1" Then
				re.Pattern="(\[flash|<object )"
				If re.Test(s) Then
					checkcontent=1
					set re=Nothing
					Exit For
				End If
			End If
			'含媒体播放检查
			If Node.selectSingleNode("checkmp/@use").text="1" Then
				re.Pattern="(\[mp)"
				If re.Test(s) Then
					checkcontent=1
					set re=Nothing
					Exit For
				End If
			End If
			'含媒rm检查
			If Node.selectSingleNode("checkrm/@use").text="1" Then
				re.Pattern="(\[rm)"
				If re.Test(s) Then
					checkcontent=1
					set re=Nothing
					Exit For
				End If
			End If
			For Each tmpnode in Node.selectNodes("checkword")
				re.Pattern=tmpnode.selectSingleNode("@content").text
				If re.Test(s) Then
					checkcontent=1
					set re=Nothing
					Exit For
				End If
			Next
		Next
		set re=Nothing
	End Function
End Class

'截取指定字符
Function cutStr(str,strlen)
	'去掉所有HTML标记
	'Dim re
	'Set re=new RegExp
	're.IgnoreCase =True
	're.Global=True
	're.Pattern="<([^<>]*?)>"
	'Do while re.Test(str)
	'	str=re.Replace(str,"")
	'Loop
	'set re=Nothing
	Str=Dvbbs.Replacehtml(Str)
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
'过滤不必要UBB,和HTML
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
	Set re=Nothing
	strContent=Replace(strContent,Chr(10),"")
	strContent=Replace(strContent,Chr(13),"")
	strContent=Replace(strContent, CHR(32), " ")
	strContent=Replace(strContent, CHR(9), " ")
	strContent=Dvbbs.Replacehtml(strContent)
	reUBBCode=replace(strContent,"<","&lt;")
End Function

'通用函数
Function IstrueName(uName)
	IstrueName=False
	If InStr(uName,"=")>0 Then Exit Function
	If InStr(uName,"%")>0 Then Exit Function 
	If InStr(uName,Chr(32))>0 Then Exit Function 
	If InStr(uName,"?")>0 Then Exit Function 
	If InStr(uName,"&")>0 Then Exit Function 
	If InStr(uName,";")>0 Then Exit Function 
	If InStr(uName,",")>0 Then Exit Function 
	If InStr(uName,"'")>0 Then Exit Function 
	If InStr(uName,Chr(34))>0 Then Exit Function 
	If InStr(uName,chr(9))>0 Then Exit Function 
	If InStr(uName,"")>0 Then Exit Function 
	If InStr(uName,"$")>0 Then Exit Function
	IstrueName=True 	
End Function
'发贴时用，为了减少入库量
'更新用户短信通知信息（新短信条数||新短讯ID||发信人名）
Sub UPDATE_User_Msg(username)
	Dim msginfo,i,UP_UserInfo,newmsg
	newmsg=newincept(username)
	If newmsg>0 Then
		msginfo=newincept(username) & "||" & inceptid(1,username) & "||" & inceptid(2,username)
	Else
		msginfo="0||0||null"
	End If
	Dvbbs.execute("UPDATE [Dv_User] Set UserMsg='"&Dvbbs.CheckStr(msginfo)&"' WHERE username='"&Dvbbs.CheckStr(username)&"'")
	If username=Dvbbs.MemberName Then
		Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermsg").text=msginfo
	Else
		Call Dvbbs.NeedUpdateList(username,1)
	End If
End Sub

'统计留言
Function newincept(iusername)
Dim Rs
Rs=Dvbbs.execute("SELECT Count(id) FROM Dv_Message WHERE flag=0 And issend=1 And DelR=0 And incept='"& Dvbbs.CheckStr(iusername) &"'")
    newincept=Rs(0)
	Set Rs=nothing
	If isnull(newincept) Then newincept=0
End Function

Function inceptid(stype,iusername)
	Dim Rs
	Set Rs=Dvbbs.execute("SELECT top 1 id,sender FROM Dv_Message WHERE flag=0 And issend=1 And DelR=0 And incept='"& Dvbbs.CheckStr(iusername) &"'")
	If not rs.eof Then
		If stype=1 Then
			inceptid=Rs(0)
		Else
			inceptid=Rs(1)
		End If
	Else
		If stype=1 Then
			inceptid=0
		Else
			inceptid="null"
		End If
	End If
	Set Rs=nothing
End Function

%>