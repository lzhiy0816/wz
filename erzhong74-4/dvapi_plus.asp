<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/ubblist.asp"-->
<%

Const PassServer = ""
Const PassIP = ""
Const PassIP2 = ""

Dim XmlDoc,PassUser
Dim F,P1
Dim RootID
F = Request("F")

If PassServer<>"" and Checkserver(PassServer) = False Then
    Response.Write "访问已取消！"
	Response.Write "<br />"

	Response.End
End If

If PassIP<>"" and PassIP2<>"" and ChkIpLimited(PassIP,PassIP2) = False then
	Response.Write "REMOTE_ADDR="&Request.ServerVariables("REMOTE_ADDR")
	Response.Write "<br />"
	Response.Write "访问已取消！"
	Response.End
End If

Set XmlDoc = Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
XmlDoc.LoadXml "<?xml version=""1.0"" encoding=""gb2312""?><root/>"
Select Case F
	Case "1" : GetBBSList()
	Case "2" : AddOneAd()
	Case "3" : CheckOneAd()
	Case "4" : DelOneAd()
	Case "5" : EditOneAd()
	Case "6" : GetOneAdNo()
End Select
ShowXml()
Set XmlDoc = Nothing
Dvbbs.PageEnd()
Set Dvbbs = Nothing

Sub ShowXml()
	Response.Clear
	Response.CharSet="gb2312"  
	Response.ContentType="text/xml"
	Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"&vbNewLine
	Response.Write XmlDoc.documentElement.XML
End Sub

Sub GetBBSList()
	Dim P1,P2
	Dim ChildNode,Node,BbsNode,i
	Dim XMLDOM
	
	P1 = Request("P1")
	P2 = Request("P2")
	Set XMLDOM = Application(Dvbbs.CacheName&"_boardlist").cloneNode(True)
	Set ChildNode = XmlDoc.createNode(1,"SiteID","")
	ChildNode.text = P1
	XmlDoc.documentElement.appendChild(ChildNode)
	Set ChildNode = XmlDoc.createNode(1,"BBSCount","")
	ChildNode.text = XMLDOM.documentElement.getElementsByTagName("board").length
	XmlDoc.documentElement.appendChild(ChildNode)

	i=0
	For Each BbsNode in XMLDOM.documentElement.getElementsByTagName("board")
		i = i + 1
		If BbsNode.attributes.getNamedItem("hidden").text<>"1" Then
			Set Node = XmlDoc.createNode(1,"B"&i,"")
			Set ChildNode = XmlDoc.createNode(1,"ID","")
			ChildNode.text = BbsNode.getAttribute("boardid")
			Node.appendChild(ChildNode)
			Set ChildNode = XmlDoc.createNode(1,"Name","")
			ChildNode.text = BbsNode.getAttribute("boardtype")
			Node.appendChild(ChildNode)
			Set ChildNode = XmlDoc.createNode(1,"URL","")
			ChildNode.text = Dvbbs.Get_ScriptNameUrl()&"index.asp?boardid="&BbsNode.getAttribute("boardid")
			Node.appendChild(ChildNode)
			Set ChildNode = XmlDoc.createNode(1,"Time","")
			ChildNode.text = "论坛创建时间"
			Node.appendChild(ChildNode)
			XmlDoc.documentElement.appendChild(Node)
		End If
	Next
End Sub

Sub CheckUserName(Newname,Userpass)
	Dim Sql,Rs
	Sql = "SELECT top 1 * FROM Dv_User WHERE UserName = '" & Newname & "'"
	Set Rs = Dvbbs.iCreateObject("Adodb.Recordset")
	If Not Isobject(Conn) Then ConnectionDatabase
	Rs.Open Sql,Conn,1,3
	If Not (Rs.Eof And Rs.Bof) Then
		Dvbbs.Userid = Rs(0)
	Else
		'加入用户表
		Dim Titlepic,TRs,UserClass
		Dim TruePassWord
		TruePassWord = Dvbbs.Createpass
		Set TRs = Dvbbs.Execute("Select UserTitle,GroupPic,UserGroupID,IsSetting,ParentGID From Dv_UserGroups Where ParentGID=3 Order By MinArticle")
			UserClass = Trs(0)
			TitlePic = Trs(1)
			Dvbbs.UserGroupID = TRs(2)
		TRs.close

		Rs.Addnew
		Rs("UserName") = Newname
		Rs("Userpassword") = Md5(Userpass,16)
		Rs("Userclass") = UserClass
		Rs("UserGroupID") = Dvbbs.UserGroupID
		Rs("Titlepic") = Titlepic
		Rs("UserPost") = 0
		Rs("UserWealth") = 100
		Rs("Userep") = 30
		Rs("Usercp") = 30
		Rs("Userisbest") = 0
		Rs("Userdel") = 0
		Rs("Userpower") = 0
		Rs("Lockuser") = 0
		Rs("UserSex") = 1
		Rs("UserEmail") = Newname & "@daqi.com"
		Rs("UserFace") = "Images/userface/image1.gif"
		Rs("UserWidth") = 32
		Rs("UserHeight") = 32 
		Rs("UserIM") = "||||||||||||||||||"
		Rs("UserFav") = "陌生人,我的好友,黑名单"
		Rs("LastLogin") = Now()
		Rs("JoinDate") = Now()
		Rs.Update
		Dvbbs.Execute("UpDate Dv_Setup Set Forum_UserNum=Forum_UserNum+1,Forum_lastUser='"&Dvbbs.HtmlEncode(Newname)&"'")

		Set Trs=Dvbbs.execute("select top 1 userid from [Dv_user] order by userid desc")
			Dvbbs.userid=Trs(0)
		Trs.close
		Set Trs=nothing
	End If
	Set Rs = Nothing
End Sub

Sub AddOneAd()
	Dim P1,P2
	Dim ChildNode,Node
	Dim Topic,Body,Boardid
	P1 = Request("P1")
	Boardid = Request("P2")
	Topic = Dvbbs.Checkstr(Request("P3"))
	Body = Dvbbs.Checkstr(Request("P4"))
	'daqi.com
	If Not IsNumeric(Boardid) Or Boardid="" Then
		Exit Sub
	Else
		Boardid = Int(Boardid)
	End If
	If Topic="" or Body="" Then
		Exit Sub
	End If
	
	Dim PostDate,TotalUseTable
	Dim Sql,Rs
	TotalUseTable = Dvbbs.NowUseBbs
	PostDate = DateTimeStr
	
	Set Rs = Dvbbs.Execute("select Boardid From Dv_Board Where Boardid="&Boardid)
	If Rs.Eof Or Rs.Bof Then
		Exit Sub
	End If
	Rs.Close
	Dim UserName,Userid,Userpassword
	Dim RanName
	RanName = "ArabianSun,vincent119,daqi.com,vfootball,linkflag,u00000l" '随机用户名单，以逗号分隔
	RanName = Split(RanName,",")
	Randomize
	UserName = RanName(Int((Ubound(RanName) + 1) * Rnd))
	Userid = 0

	Userpassword = Dvbbs.Createpass
	If Dvbbs.Userid>0 Then
		UserName = Dvbbs.Membername
		Userid = Dvbbs.Userid
	Else
		CheckUserName UserName,Userpassword
		Userid = Dvbbs.Userid
	End If

	SQL="insert into Dv_topic (Title,Boardid,PostUsername,PostUserid,DateAndTime,Expression,LastPost,LastPostTime,PostTable,locktopic,istop,TopicMode,isvote,PollID,Mode,GetMoney,UseTools,GetMoneyType,isSmsTopic,HideName) values ('"&Topic&"',"&Boardid&",'"&UserName&"',"&Userid&",'"&PostDate&"','0|face1.gif','$$"&PostDate&"$$$$','"&PostDate&"','"&TotalUseTable&"',0,1,0,0,0,0,0,'',0,0,0)"
	Dvbbs.Execute(sql)
	Set Rs=Dvbbs.Execute("select Max(topicid) From Dv_topic Where PostUserid="&Userid)
	RootID=Rs(0)
	
	DIM UbblistBody
	UbblistBody = Ubblist(Body)
	SQL="insert into "&TotalUseTable&"(Boardid,ParentID,username,topic,body,DateAndTime,length,RootID,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID,isupload,IsAudit,Ubblist,GetMoney,UseTools,PostBuyUser,GetMoneyType) values ("&Boardid&",0,'"&UserName&"','"&topic&"','"&Body&"','"&PostDate&"','"&Dvbbs.strlength(Body)&"',"&RootID&",1,0,'"&Dvbbs.UserTrueIP&"','face1.gif',0,0,0,0,"&Userid&",0,0,'"&UbblistBody&"',0,'','',0)"
	
	Dvbbs.Execute(sql)
	Dim AnnounceID,LastPost_1
	Set Rs=Dvbbs.Execute("select Max(AnnounceID) From "&TotalUseTable&" Where PostUserID="&UserID)
	AnnounceID=Rs(0)

	LastPost_1=Replace(UserName,"$","") & "$" & AnnounceID & "$" & DateTimeStr & "$" & Replace(cutStr(topic,20),"$","&#36;") & "$$" & Userid & "$" & RootID & "$" & BoardID
	Dim BoardTopStr
	Set Rs=Dvbbs.Execute("Select BoardID,BoardTopStr From Dv_Board Where BoardID="&BoardID)
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
		Dvbbs.Execute("Update Dv_Board Set BoardTopStr='"&BoardTopStr&"',PostNum=PostNum+1,todaynum=todaynum+1,LastPost='"&LastPost_1&"' Where BoardID="&Rs(0))
		Dvbbs.LoadBoardinformation Rs(0)
	End If
	Set Rs=Nothing
	Dvbbs.Execute("Update Dv_Topic Set LastPost='"&LastPost_1&"' where Topicid="&RootID)
	
	Set ChildNode = XmlDoc.createNode(1,"R1","")
	ChildNode.text = P1
	XmlDoc.documentElement.appendChild(ChildNode)
	Set ChildNode = XmlDoc.createNode(1,"R2","")
	ChildNode.text = RootID
	XmlDoc.documentElement.appendChild(ChildNode)
	Set ChildNode = XmlDoc.createNode(1,"R3","")
	ChildNode.text = Dvbbs.Get_ScriptNameUrl()&"dispbbs.asp?boardid="&Boardid&"&id="&RootID
	XmlDoc.documentElement.appendChild(ChildNode)

End Sub

Sub CheckOneAd()
	Dim P1,P2,TopidID,Boardid,IsFound,Rs,Topic
	P1 = Request("P1")
	Boardid = Request("P2")
	TopidID = Request("P3")
	Topic = Dvbbs.Checkstr(Request("P4"))
	IsFound = 1
	If Not IsNumeric(Boardid) Or Boardid="" Then
		Exit Sub
	Else
		Boardid = Int(Boardid)
	End If

	If Not IsNumeric(TopidID) Or TopidID="" Then
		Exit Sub
	Else
		TopidID = Int(TopidID)
	End If
	Set Rs = Dvbbs.Execute("Select Topicid From Dv_Topic where Topicid="&TopidID&" and Title='"&Topic&"'")
	If Not (Rs.Eof And Rs.Bof) Then
		IsFound = 0
	End If
	Rs.Close
	Set Rs = Nothing
	Dim ChildNode
	Set ChildNode = XmlDoc.createNode(1,"R1","")
	ChildNode.text = P1
	XmlDoc.documentElement.appendChild(ChildNode)
	Set ChildNode = XmlDoc.createNode(1,"R2","")
	ChildNode.text = IsFound
	XmlDoc.documentElement.appendChild(ChildNode)

End Sub

Sub DelOneAd()
	Dim P1,P2,TopidID,Boardid,IsFound,Rs,Sql
	P1 = Request("P1")
	Boardid = Request("P2")
	TopidID = Request("P3")
	IsFound = 1
	If Not IsNumeric(Boardid) Or Boardid="" Then
		Exit Sub
	Else
		Boardid = Int(Boardid)
	End If

	If Not IsNumeric(TopidID) Or TopidID="" Then
		Exit Sub
	Else
		TopidID = Int(TopidID)
	End If
	Dim PostTable
	Set Rs = Dvbbs.Execute("Select Topicid,PostTable,Boardid,PostUserName From Dv_Topic where Topicid="&TopidID)
	If Not Rs.Eof Then
		SQL = Rs.GetRows(1)
		Rs.Close:Set Rs = Nothing
		If CanDelTopic(SQL(3,0)) Then
			IsFound = 0
			PostTable = Rs(1)
			Boardid = Rs(2)
			Dvbbs.Execute("Delete From dv_topic where Topicid="&TopidID)
			Dvbbs.Execute("Delete From "&PostTable&" where Rootid="&TopidID)

			Dim BoardTopStr
			Set Rs=Dvbbs.Execute("Select BoardID,BoardTopStr From Dv_Board Where BoardID="&BoardID)
			If Not Rs.Eof Then
				If Not (Rs(1)="" Or IsNull(Rs(1))) Then
					If InStr(","&Rs(1)&",","," & TopidID & ",")>0 Then
						BoardTopStr  = Replace(","&Rs(1)&",","," & TopidID & ",",",")
						If BoardTopStr<>"" Then
							If Left(BoardTopStr,1)="," Then
								BoardTopStr = Mid(BoardTopStr,2)
							End If
							If Right(BoardTopStr,1)="," Then
								BoardTopStr = Left(BoardTopStr,Len(BoardTopStr)-1)
							End If
						End If
						Dvbbs.Execute("Update Dv_Board Set BoardTopStr='"&BoardTopStr&"',PostNum=PostNum-1 Where BoardID="&Rs(0))
						Dvbbs.LoadBoardinformation Rs(0)
					End If
				End If
			End If
			Rs.Close:Set Rs = Nothing
		End If
	Else
		Rs.Close:Set Rs = Nothing
	End If

	Dim ChildNode
	Set ChildNode = XmlDoc.createNode(1,"R1","")
	ChildNode.text = P1
	XmlDoc.documentElement.appendChild(ChildNode)
	Set ChildNode = XmlDoc.createNode(1,"R2","")
	ChildNode.text = IsFound
	XmlDoc.documentElement.appendChild(ChildNode)
End Sub

Function CanDelTopic(TopicUsername)
	CanDelTopic=False
	If Dvbbs.FoundUserPer Then
		If TopicUsername=Dvbbs.MemberName Then
			If Cint(Dvbbs.GroupSetting(11))=1 Then CanDelTopic=True:Exit Function
		Else
			If Cint(Dvbbs.GroupSetting(18))=1 Then CanDelTopic=True:Exit Function
		End If
	End If
	If Cint(Dvbbs.GroupSetting(18))=1 Then
		If Dvbbs.master or Dvbbs.superboardmaster or Dvbbs.boardmaster Then CanDelTopic=True:Exit Function
		If Dvbbs.UserGroupID>3 Then CanDelTopic=True:Exit Function
	End If
	If Cint(Dvbbs.GroupSetting(11))=1 and TopicUsername=Dvbbs.MemberName Then CanDelTopic=True:Exit Function
End Function

Sub EditOneAd()
	Dim P1,P2,TopidID,Boardid,Rs,Sql,Old_User
	Dim Topic,Body,ErrCode
	P1 = Request("P1")
	Boardid = Dvbbs.CheckNumeric(Request("P2"))
	TopidID = Dvbbs.CheckNumeric(Request("P3"))
	Topic = Dvbbs.Checkstr(Request("P4"))
	Body = Dvbbs.Checkstr(Request("P5"))
	If Application(Dvbbs.CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"& Boardid&"']/@nopost").text="1" Then CanEditAd=1:Exit Sub
	If Application(Dvbbs.CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"& Boardid&"']/@hidden").text="1" and Dvbbs.GroupSetting(37)="0" Then CanEditAd=1:Exit Sub

	Set Rs = Dvbbs.Execute("Select T.Boardid,T.PostUserName,T.Title,T.DateAndTime,T.PostTable,u.UserGroupID From Dv_Topic T left outer join [dv_user] u on T.postuserid=u.userid where T.Topicid="&TopidID)
	If Not Rs.Eof Then
		Dim R_BoardID,R_DateAndTime,R_PostUserName,R_UserGroupID,R_PostTable
		R_BoardID = Rs("Boardid")
		R_DateAndTime = Rs("dateandtime")
		R_PostUserName = Rs("PostUserName")
		R_UserGroupID = Rs("UserGroupID")
		R_PostTable = Rs("PostTable")
		Rs.Close:Set Rs = Nothing
		If R_BoardID=Boardid And CanEditAd(R_DateAndTime,R_PostUserName,R_UserGroupID) Then
			ErrCode = 0
			Sql = "Update dv_topic set title='"&Topic&"' where topicid="&TopidID
			Dvbbs.Execute(Sql)
			DIM UbblistBody
			UbblistBody = Ubblist(Body)
			Sql = "Update "&R_PostTable&" set topic='"&Topic&"',body='"&body&"',Ubblist='"&UbblistBody&"' where parentid=0 and Rootid="&TopidID
			Dvbbs.Execute(Sql)
		Else
			ErrCode = 1
		End If
	Else
		Rs.Close:Set Rs = Nothing
		ErrCode = 2
	End If

	Dim ChildNode
	Set ChildNode = XmlDoc.createNode(1,"R1","")
	ChildNode.text = P1
	XmlDoc.documentElement.appendChild(ChildNode)
	Set ChildNode = XmlDoc.createNode(1,"R2","")
	ChildNode.text = ErrCode
	XmlDoc.documentElement.appendChild(ChildNode)
End Sub

Function CanEditAd(DateAndTime,PostUserName,UserGroupID)
	CanEditAd=0
	If Clng(Dvbbs.forum_setting(51))>0 and Not (Dvbbs.Master or Dvbbs.BoardMaster or Dvbbs.SuperBoardMaster) Then
		If DateDiff("s",DateAndTime,Now())>Clng(Dvbbs.Forum_Setting(51))*60 Then Exit Function
	End If
	If PostUserName=Dvbbs.membername Then 
		If Clng(Dvbbs.GroupSetting(10))=1 Then CanEditAd=1
	Else
		If Cint(Dvbbs.UserGroupID) < 4 And Cint(Dvbbs.UserGroupID) >= UserGroupID Then Exit Function
		If Clng(Dvbbs.GroupSetting(23))=1 Then
			If Dvbbs.founduserPer Then CanEditAd=1:Exit Function
			If Cint(Dvbbs.UserGroupID) > 3 Then CanEditAd=1:Exit Function
			If Dvbbs.master or Dvbbs.superboardmaster or Dvbbs.boardmaster Then CanEditAd=1
		End If
	End If
End Function

Sub GetOneAdNo()
	Dim P1,P2,TopidID,Boardid,Rs
	Dim Errmsg,Hits,Childs
	P1 = Request("P1")
	Boardid = Request("P2")
	TopidID = Request("P3")
	If Not IsNumeric(Boardid) Or Boardid="" Then
		Exit Sub
	Else
		Boardid = Int(Boardid)
	End If

	If Not IsNumeric(TopidID) Or TopidID="" Then
		Exit Sub
	Else
		TopidID = Int(TopidID)
	End If
	Hits = 0
	Childs = 0
	Errmsg = ""
	Set Rs = Dvbbs.Execute("Select Topicid,Child,Hits From Dv_Topic where Topicid="&TopidID)
	If Not (Rs.Eof And Rs.Bof) Then
		Hits = Rs(2)
		Childs = Rs(1)
	Else
		Errmsg = "相关主题未找到!"
	End If
	Rs.Close
	Set Rs = Nothing
	Dim ChildNode
	Set ChildNode = XmlDoc.createNode(1,"R1","")
	ChildNode.text = P1
	XmlDoc.documentElement.appendChild(ChildNode)
	Set ChildNode = XmlDoc.createNode(1,"R2","")
	ChildNode.text = Hits
	XmlDoc.documentElement.appendChild(ChildNode)
	Set ChildNode = XmlDoc.createNode(1,"R3","")
	ChildNode.text = Childs
	XmlDoc.documentElement.appendChild(ChildNode)
	Set ChildNode = XmlDoc.createNode(1,"R4","")
	ChildNode.text = Errmsg
	XmlDoc.documentElement.appendChild(ChildNode)
End Sub

'截取指定字符
Function cutStr(str,strlen)
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

Function DateTimeStr()
	DateTimeStr = Replace(Replace(CSTR(NOW()+Dvbbs.Forum_Setting(0)/24),"上午",""),"下午","")
End Function

Function ChkIpLimited(ViewIpLimited,KeyIP2)
	Dim ReServerIp
	ChkIpLimited = False
	If ViewIpLimited = "" Then Exit Function
  If KeyIP2 = "" Then Exit Function		
  	
	ReServerIp = Trim(Request.ServerVariables("REMOTE_ADDR"))
	If Instr(ViewIpLimited,ReServerIp) or Instr(KeyIP2,ReServerIp) Then
		ChkIpLimited = True
	End If
End Function

Function CheckServer(str)
	Dim i,Servername
	CheckServer = False
	If Str = "" Then  Exit Function
	Str = split(Cstr(str),",")
	Servername = Request.ServerVariables("HTTP_REFERER")
	for i=0 to Ubound(str)
	if right(str(i),1)="/" then str(i)=left(trim(str(i)),len(str(i))-1)
		if Lcase(left(servername,len(str(i))))=Lcase(str(i)) then
			CheckServer = true
			exit for
		else
			CheckServer = false
		end if
	next
End Function
%>