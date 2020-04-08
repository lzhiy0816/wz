<%
'=========================================================
' File: Dv_ClsMain.asp
' Version:7.1.0 sp1
' Date: 2005-8-1
' Script Written by dvbbs.net
'=========================================================
' Copyright (C) 2003,2004 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: eway@aspsky.net
'=========================================================
'是否商业版，非官方SQL版本请在此设置为0以及在Conn中设置论坛为SQL数据库，否则显示不正常
Const IsBuss=1
Const Dvbbs_Server_Url = "http://server.dvbbs.net/"
Const Dvbbs_PayTo_Url = "http://pay.dvbbs.net/"
Dim IP_MAX
Dim tourl
Class Cls_Forum
	Rem Const
	Public BoardID,SqlQueryNum,Forum_Info,Forum_Setting,Forum_user,Forum_Copyright,Forum_ads,Forum_ChanSetting,Forum_UploadSetting
	Public Forum_sn,Forum_Version,Stats,StyleName,ErrCodes,NowUseBBS,Cookiepath,ScriptFolder,BoardInfoData,UserSession
	Public lanstr,mainhtml,mainsetting,sysmenu,mainpic,UserToday,BoardJumpList,BoardList,CacheData,Maxonline
	Public VipGroupUser,Vipuser,Boardmaster,Superboardmaster,Master,FoundIsChallenge,FoundUser
	Public ScriptName,MemberName,MemberWord,MemberClass,UserHidden,UserID,UserTrueIP,UserPermission
	Public sendmsgnum,sendmsgid,sendmsguser,Page_Admin
	Public BadWords,rBadWord,Forum_emot,Forum_PostFace,Forum_UserFace,SkinID,Forum_PicUrl
	Private Forum_CSS,Main_Sid,Nowstats,CssID
	Public Reloadtime,CacheName,UserGroupID,Lastlogin,GroupSetting,FoundUserPer,UserGroupParent,UserGroupParentID
	Private LocalCacheName,IsTopTable,ShowErrType
	Public Board_Setting,LastPost,Board_user,BoardType,Board_Data,Sid,Boardreadme,BoardRootID,BoardParentID
	Private Is_Isapi_Rewrite,iArchiverUrl
	Public ModHtmlLinked,ArchiverUrl,ArchiverType
	Public Browser,version ,platform,IsSearch,Cls_IsSearch
	Public IsUserPermissionOnly,IsUserPermissionAll,ShowSQL,actforip
	Rem Sub 
	Private Sub Class_Initialize()
		Forum_sn="DvForum"'如果一个虚拟目录或站点开多个论坛，则每个要错开，不能定义同一个名称
		CacheName="DvCache"'如果一个虚拟目录或站点开多个论坛，则每个要错开，不能定义同一个名称
		If Not Response.IsClientConnected Then
			Session(CacheName & "UserID")=empty
			Set Dvbbs=Nothing
			Response.End
		End If
		IsUserPermissionOnly = 0
		IsUserPermissionAll = 0
		ShowErrType = 0 '错误信息显示模式
		SqlQueryNum = 0
		Reloadtime=28800
		IsTopTable = 0
		VipGroupUser = False:IsSearch=False:Cls_IsSearch=False
		Vipuser = False:Boardmaster = False
		Superboardmaster = False:Master = False:FoundIsChallenge = False:FoundUser = False
		BoardID = Request("BoardID")
		If IsNumeric(BoardID) = 0 or BoardID = "" Then BoardID = 0
		BoardID = Clng(BoardID)
		MemberName = checkStr(Trim(Request.Cookies(Forum_sn)("username")))
		MemberWord = checkStr(Trim(Request.Cookies(Forum_sn)("password")))
		UserHidden = Trim(Request.Cookies(Forum_sn)("userhidden"))
		UserID = Trim(Request.Cookies(Forum_sn)("UserID"))
		If IsNumeric(UserHidden) = 0 or Userhidden = "" Then UserHidden = 2
		If IsNumeric(UserID) = 0 Or UserID="" Then UserID=0
		UserID = Clng(UserID)
		UserTrueIP = getIP()
		IP_MAX=0
		Dim Tmpstr
		Tmpstr = Request.ServerVariables("PATH_INFO")
		Tmpstr = Split(Tmpstr,"/")
		ScriptName = Lcase(Tmpstr(UBound(Tmpstr)))
		ScriptFolder = Lcase(Tmpstr(UBound(Tmpstr)-1)) & "/"
		MemberClass = checkStr(Request.Cookies(Forum_sn)("userclass"))
		Page_Admin=False
		If InStr(ScriptName,"showerr")>0 Or InStr(ScriptName,"login")>0 Or InStr(ScriptName,"admin_")>0  Then Page_Admin=True
		sendmsgnum=0:sendmsgid=0:sendmsguser=""
		'模拟HTML部分开始
		Is_Isapi_Rewrite = 0
		If Is_Isapi_Rewrite = 0 Then ModHtmlLinked = "?"
		ArchiverType = 0
		If InStr(ScriptName,"indexhtml.asp") > 0 Then
			iArchiverUrl = Lcase(Request.ServerVariables("QUERY_STRING"))
			If iArchiverUrl <> "" Then
				ArchiverUrl = iArchiverUrl
				iArchiverUrl = Split(iArchiverUrl,"_")
				If iArchiverUrl(0) = "list" And Ubound(iArchiverUrl) = 5 Then
					If IsNumeric(iArchiverUrl(1)) Then
						ArchiverType = 1
						BoardID = Clng(iArchiverUrl(1))
					End If
				End If
			End If
		End If
		'模拟HTML部分结束
		'Response.Write Server.MapPath("index.asp")
		'response.end
	End Sub
	Private Function getIP() 
		Dim strIPAddr 
		If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then 
			strIPAddr = Request.ServerVariables("REMOTE_ADDR") 
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then 
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1) 
			actforip=Request.ServerVariables("REMOTE_ADDR")
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then 
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)
			actforip=Request.ServerVariables("REMOTE_ADDR")
		Else 
			strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
			actforip=Request.ServerVariables("REMOTE_ADDR")
		End If 
		getIP = CheckStr(Trim(Mid(strIPAddr, 1, 30)))
	End Function 

	Private Sub class_terminate()
		If EnabledSession Then
			If Not UserSession Is Nothing  Then Session(CacheName & "UserID")= UserSession.xml
		End If
		Set UserSession=Nothing 
		If IsObject(Conn) Then Conn.Close : Set Conn = Nothing
		If IsObject(Plus_Conn) Then Plus_Conn.Close : Set Plus_Conn = Nothing
	End Sub
	Public Sub Sendmessanger(touserid,senduser,messangertext)
		Dim Node
		If Not IsObject( Application(Dvbbs.CacheName&"_messanger")) Then
			Set  Application(Dvbbs.CacheName&"_messanger")=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			 Application(Dvbbs.CacheName&"_messanger").appendChild( Application(Dvbbs.CacheName&"_messanger").createElement("xml"))
		End If
		For Each Node in Application(Dvbbs.CacheName&"_messanger").documentElement.SelectNodes("messanger")
			If datediff("s",Node.selectSingleNode("@sendtime").text,Now()) > 72000 Then
				Application(Dvbbs.CacheName&"_messanger").documentElement.removeChild(Node)
			End If
		Next
		Set Node=Application(Dvbbs.CacheName&"_messanger").documentElement.appendChild(Application(Dvbbs.CacheName&"_messanger").createNode(1,"messanger",""))
		Node.attributes.setNamedItem(Application(Dvbbs.CacheName&"_messanger").createNode(2,"sendtime","")).text=Now()
		Node.attributes.setNamedItem(Application(Dvbbs.CacheName&"_messanger").createNode(2,"touserid","")).text=touserid
		Node.attributes.setNamedItem(Application(Dvbbs.CacheName&"_messanger").createNode(2,"senduser","")).text=senduser
		Node.text=messangertext
	End Sub
	Public Property Let Name(ByVal vNewValue)
		LocalCacheName = LCase(vNewValue)
	End Property
	Public Property Let Value(ByVal vNewValue)
		If LocalCacheName<>"" Then
			Application.Lock
			Application(CacheName & "_" & LocalCacheName &"_-time")=Now()
			Application(CacheName & "_" & LocalCacheName) = vNewValue
			Application.unLock
		End If
	End Property
	Public Property Get Value()
		If LocalCacheName<>"" Then 	
				Value=Application(CacheName & "_" & LocalCacheName)
		End If
	End Property
	Public Function ObjIsEmpty()
		ObjIsEmpty=True	
		If  Not IsDate(Application(CacheName & "_" & LocalCacheName &"_-time")) Then Exit Function
		If DateDiff("s",CDate(Application(CacheName & "_" & LocalCacheName &"_-time")),Now()) < (60*Reloadtime) Then ObjIsEmpty=False		
	End Function
	'取得基本设置数据
	Public Sub loadSetup()
		Dim Rs,locklist,ip,ip1,XMLDom,Node,i
		Application.Lock
		Set XMLDom=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		XMLDom.Load Server.MapPath(MyDbPath &"inc\guest.xml")
		Set Application(Dvbbs.CacheName&"_info_guest")=XMLDom.cloneNode(True)
		Set XMLDom=Nothing
		Application.UnLock
		Name="setup"
		Set Rs = Dvbbs.Execute("Select id, Forum_Setting, Forum_ads, Forum_Badwords, Forum_rBadword, Forum_Maxonline, Forum_MaxonlineDate, Forum_TopicNum, Forum_PostNum, Forum_TodayNum, Forum_UserNum, Forum_YesTerdayNum, Forum_MaxPostNum, Forum_MaxPostDate, Forum_lastUser, Forum_LastPost, Forum_BirthUser, Forum_Sid, Forum_Version, Forum_NowUseBBS, Forum_IsInstall, Forum_challengePassWord, Forum_Ad, Forum_ChanName, Forum_ChanSetting, Forum_LockIP, Forum_Cookiespath, Forum_Boards, Forum_alltopnum, Forum_pack, Forum_Cid, Forum_AvaSiteID, Forum_AvaSign, Forum_AdminFolder, Forum_BoardXML, Forum_Css From [Dv_Setup]")
		Value = Rs.GetRows(1)
		CacheData=value
		Set XMLDom=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		XMLDom.appendChild(XMLDom.createElement("xml"))
		locklist=Trim(CacheData(25,0))
		locklist=Split(locklist,"|")
		For Each Ip in locklist
			Ip1=Split(Ip,".")
			Set Node=XMLDom.documentElement.appendChild(XMLDom.createNode(1,"lockip",""))
			For i=0 to UBound(ip1)
				Node.attributes.setNamedItem(XMLDom.createNode(2,"number"& (i+1),"")).text=ip1(i)
			Next
		Next
		Application.Lock
		Set Application(CacheName & "_forum_lockip")=XMLDom.cloneNode(True)
		Application.UnLock
		Set XMLDom=Nothing
		Dim stylesheet
		Set stylesheet=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		stylesheet.load Server.MapPath(MyDbPath &"inc\GetBrowser.xslt")
		Application.Lock
		Set Application(CacheName & "_getbrowser")=Server.CreateObject("msxml2.XSLTemplate" & MsxmlVersion)
		Application(CacheName & "_getbrowser").stylesheet=stylesheet
		Set Application(CacheName & "_csslist")=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		Application(CacheName & "_csslist").Loadxml CacheData(35,0)
		Set Application(CacheName & "_accesstopic")=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		Application(CacheName & "_accesstopic").Loadxml CacheData(34,0)
		Application.unLock
	End Sub
	Public Sub LoadBoardList()
		Application.Lock
		Dim Rs,boardmaster,master,node,Board_setting
		Set Rs=Execute("select boardid,boardtype,ParentID,depth,rootid,Child,indeximg,parentstr,cid as checkout,cid as hidden,cid as nopost,cid as checklock,cid as mode,cid as simplenesscount,readme From Dv_board Order by rootid,Orders")
		Set Application(CacheName&"_boardlist")=RecordsetToxml(rs,"board","BoardList")
		Rs.Close
		Set Rs=Execute("select boardid From Dv_board Order by Orders")
		Set Application(CacheName&"_boardmaster")=RecordsetToxml(rs,"boardmaster","masterlist")
		Rs.Close
		Set Rs=Execute("select boardmaster,boardid,Board_setting From Dv_board Order by Orders")
		Do While Not Rs.EOF
			boardmaster=split(Rs("boardmaster")&"","|")
			Set Node=Application(CacheName&"_boardmaster").documentElement.selectSingleNode("boardmaster[@boardid='"& Rs(1)&"']")
			For Each Master In boardmaster
				Node.appendChild(Application(CacheName&"_boardmaster").createNode(1,"master","")).text=Master
			Next
			Board_setting=Split(Rs("Board_setting"),",")
			Application(CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@checkout").text=Board_setting(2)
			Application(CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@hidden").text=Board_setting(1)
			Application(CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@nopost").text=Board_setting(43)
			Application(CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@checklock").text=Board_setting(0)
			Application(CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@mode").text=Board_setting(39)
			Application(CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@simplenesscount").text=Board_setting(41)
			Rs.MoveNext
		Loop
		Set Application(CacheName&"_sboardlist")=Application(CacheName&"_boardlist").cloneNode(True)
		For each node in Application(CacheName&"_sboardlist").documentElement.selectNodes("board")
			node.attributes.removeNamedItem("readme")
			node.attributes.removeNamedItem("simplenesscount")
			node.attributes.removeNamedItem("mode")
			node.attributes.removeNamedItem("checklock")
			node.attributes.removeNamedItem("checkout")
			node.attributes.removeNamedItem("parentstr")
			node.attributes.removeNamedItem("indeximg")
		Next
		MakBoardNav 0 ,""
		Application.unLock
		Rs.Close
		Set Rs= Nothing
	End Sub
	Public Sub MakBoardNav(parentid,Node1)
		Dim Node,Dom
		If parentid=0 Then 	
			Set Application(CacheName&"_ssboardlist")=Server.CreateObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion )
			Set Node1=Application(CacheName&"_ssboardlist").appendChild(Application(CacheName&"_ssboardlist").createElement("BoardList"))
		End If
		For Each Node in Application(CacheName&"_sboardlist").documentElement.selectNodes("board[@parentid="&parentid&"]")
			MakBoardNav Node.selectSingleNode("@boardid").text,Node1.appendChild(Node.cloneNode(True))
		Next
	End Sub
	Public Sub LoadPlusMenu()
		Name = "ForumPlusMenu"
		Dim Rs,XMLDom,Node,plus_setting,stylesheet,XMLStyle,proc
		Set Rs=Execute("Select id,plus_type,plus_name,mainpage,plus_copyright,plus_setting,isshowmenu as width,isshowmenu as height From Dv_Plus Where  Isuse=1 Order By ID")
		Set XMLDom=RecordsetToxml(rs,"plus","")
		Set Rs=Nothing
		For Each Node In XMLDom.documentElement.selectNodes("plus")
			plus_setting=Split(Split(node.selectSingleNode("@plus_setting").text,"|||")(0),"|")
			node.selectSingleNode("@plus_setting").text=plus_setting(0)
			node.selectSingleNode("@width").text=plus_setting(1)
			node.selectSingleNode("@height").text=plus_setting(2)
		Next
		Set XMLStyle=Server.CreateObject("msxml2.XSLTemplate" & MsxmlVersion)
		Set stylesheet=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		stylesheet.load Server.MapPath(MyDbPath &"inc\Templates\plusmenu.xslt")
		XMLStyle.stylesheet=stylesheet
		Set proc=XMLStyle.createProcessor()
		proc.input = XMLDom
  	proc.transform()
  	value=proc.output
	End Sub
	Public Sub LoadBoardData(bid)
		Dim Rs
		Set Rs=Execute("select boardid,boarduser,board_ads,board_user,isgroupsetting,rootid,board_setting,sid,cid,Rules From Dv_board Where Boardid="&bid)
		Set Application(CacheName &"_boarddata_" & bid)=RecordsetToxml(rs,"boarddata","")
		Rs.Close
		Set Rs= Nothing
	End Sub
	Public Sub LoadBoardinformation(bid)'加载动态板面信息数据
		Dim Rs,lastpost,i
		Set Rs=Execute("select boardid,boardtopstr,postnum,topicnum,todaynum,lastpost as lastpost_0,boardid as lastpost_1,boardid as lastpost_2,boardid as lastpost_3,boardid as lastpost_4,boardid as lastpost_5,boardid as lastpost_6,boardid as lastpost_7 From Dv_board Where Boardid="&bid)
		Set Application(CacheName &"_information_" & bid)=RecordsetToxml(rs,"information","")
		lastpost=Split(Application(CacheName &"_information_" & bid).documentElement.selectSingleNode("information/@lastpost_0").text,"$")
		For i=0 to UBound(lastpost)
			Application(CacheName &"_information_" & bid).documentElement.selectSingleNode("information/@lastpost_"& i &"").text=lastpost(i)
			If i = 7 Then Exit For
		Next
		Rs.Close
		Set Rs= Nothing
	End Sub
	Public Sub LoadGroupSetting()
		Dim Rs
		Set Rs=Dvbbs.Execute("Select GroupSetting,UserGroupID,ParentGID,IsSetting,UserTitle From Dv_UserGroups")
		Set Application(CacheName &"_groupsetting")=RecordsetToxml(rs,"usergroup","")
		Set Rs=Dvbbs.Execute("Select UserGroupID,usertitle,titlepic,orders From Dv_UserGroups order by orders")
		Set Application(CacheName &"_grouppic")=RecordsetToxml(rs,"usergroup","grouppic")
		Set Rs=Nothing
	End Sub
	Public Sub Loadstyle()
		Dim Rs
		Set Rs=Dvbbs.Execute("Select *  From Dv_style")
		Set Application(CacheName &"_style")=RecordsetToxml(rs,"style","")
		Set Rs=Nothing
		LoadStyleMenu()
	End Sub
	Public Sub LoadStyleMenu()'生成风格选单数据
		Name="style_list"
		Dim XMLDom,stylesheet,XMLStyle,proc
		Set XMLDom=Application(CacheName &"_style").cloneNode(True)
		XMLDom.documentElement.appendChild(Application(CacheName & "_csslist").documentElement.cloneNode(True))
		Set XMLStyle=Server.CreateObject("msxml2.XSLTemplate" & MsxmlVersion)
		Set stylesheet=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		stylesheet.load Server.MapPath(MyDbPath &"inc\Templates\stylemenu.xslt")
		XMLStyle.stylesheet=stylesheet
		Set proc=XMLStyle.createProcessor()
		proc.input = XMLDom
  	proc.transform()
  	value=proc.output
	End Sub
	Public Sub UpdateForum_Info(act)'act=0 不处理缓存,act=1 处理缓存
		If value <> "1900-1-1" Then 
			value="1900-1-1"
			Dim Rs,LastPostInfo,TempStr,i,Board
			Dim Forum_YesterdayNum,Forum_TodayNum,Forum_LastPost,Forum_MaxPostNum,Forum_MaxPostDate
			Set Rs=Execute("Select Top 1 Forum_YesterdayNum,Forum_TodayNum,Forum_LastPost,Forum_MaxPostNum From Dv_Setup")
			Forum_YesterdayNum=Rs(0)
			Forum_TodayNum=Rs(1)
			Forum_LastPost=Rs(2)
			Forum_MaxPostNum=Rs(3)
			Set Rs=Nothing
			LastPostInfo = Split(Forum_LastPost,"$")
			If Not IsDate(LastPostInfo(2)) Then LastPostInfo(2)=Now()	
			If DateDiff("d",CDate(LastPostInfo(2)),Now())<>0 Then'最后发帖时间不是今天，	
				TempStr=LastPostInfo(0)&"$"&LastPostInfo(1)&"$"&Now()&"$"&LastPostInfo(3)&"$"&LastPostInfo(4)&"$"&LastPostInfo(5)&"$"&LastPostInfo(6)&"$"&LastPostInfo(7)
				Execute("Update Dv_Setup Set Forum_YesterdayNum="&Forum_TodayNum&",Forum_LastPost='"&TempStr&"',Forum_TodayNum=0")
				Execute("update Dv_board Set TodayNum=0")
				If act=1 Then
					If not IsObject(Application(CacheName&"_boardlist")) Then LoadBoardList()
					For Each board in Application(CacheName&"_boardlist").documentElement.selectNodes("board/@boardid")
						LoadBoardinformation board.text
					Next
				End If
			End If
			If Forum_TodayNum >Forum_MaxPostNum Then
				Execute("Update Dv_Setup Set Forum_MaxPostNum=Forum_TodayNum,Forum_MaxPostDate="&SqlNowString)
			End If
			If act=1 Then loadSetup()
			Dim xmlhttp
			If IsSqlDataBase =0 Then
				On Error Resume Next
				Set xmlhttp = Server.CreateObject("msxml2.ServerXMLHTTP")
				xmlhttp.setTimeouts 65000, 65000, 65000, 65000
		  	xmlhttp.Open "POST",Get_ScriptNameUrl& "Loadservoces.asp",false
		  	xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		  	xmlhttp.send()
		  	Set xmlhttp = Nothing
			End If
		End If
		Name="Date"
		value=Date()
	End Sub
	Public Sub GetForum_Setting()
		Name="Date"
		If ObjIsEmpty() Then
			UpdateForum_Info(0)
		ElseIf  Cstr(value) <> Cstr(Date()) Then 
			UpdateForum_Info(1)
		End If
		Name="setup"
		If ObjIsEmpty Then loadSetup()
		If Not IsObject(Application(CacheName&"_boardlist")) Then
				LoadBoardList()
		End If
		If Not IsObject(Application(CacheName &"_style")) Then
				Loadstyle()
		End If
		Name="setup"
		CacheData=value
		Dim Setting,OpenTime,ischeck:Setting= Split(CacheData(1,0),"|||"):Forum_Info =  Split (Setting(0),",")
		Forum_Setting = Split (Setting(1),","):Forum_UploadSetting = Split(Forum_Setting(7),"|")
		Forum_user = Setting(2):Forum_user = Split (Forum_user,","):Forum_Copyright = Setting(3)
		Forum_ChanSetting = Split(CacheData(24,0),","):	Forum_Version = CacheData(18,0):BadWords = Split(CacheData(3,0),"|")
		rBadWord = Split(CacheData(4,0),"|"):	Main_Sid=CacheData(17,0):Maxonline = CacheData(5,0):NowUseBBS = CacheData(19,0):Cookiepath = CacheData(26,0)
		If ScriptFolder = Lcase(CacheData(33,0)) Then Page_Admin = True
		Rem 禁止代理服务器访问开始,如需要允许访问，请屏蔽此段代码。
		If Forum_Setting(100)="1" Then
			If actforip <> "" Then
				Session(CacheName & "UserID")=empty
				Set Dvbbs=Nothing
				Response.Status = "302 Object Moved" 
				Response.End 
			End If
			If UBound(Forum_Setting)> 101 Then
				IP_MAX=CLng(Forum_Setting(101))
			Else
				IP_MAX=0
			End If
		End If
		Rem 禁止代理服务器访问结束
		If Forum_Setting(21)="1" And Not Page_Admin Then Set Dvbbs=Nothing:Response.redirect "showerr.asp?action=stop"	
		If BoardID <>0 Then
			If Application(CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"&BoardID&"']") Is Nothing Then
				Set Dvbbs=Nothing
				Response.Write "错误的版面参数"
  			Response.End
			End If
		End If
		If BoardID > 0 Then
			If Not IsObject(Application(CacheName &"_boarddata_" & Boardid)) Then LoadBoardData boardid
			If Not IsObject (Application(CacheName &"_information_" & boardid)) Then LoadBoardinformation BoardID
			Dim Nodelist,node
			Forum_ads = Split(Application(CacheName &"_boarddata_" & Boardid).documentElement.selectSingleNode("boarddata/@board_ads").text,"$")
			Forum_user = Split(Application(CacheName &"_boarddata_" & Boardid).documentElement.selectSingleNode("boarddata/@board_user").text,",")
			board_Setting = Split(Application(CacheName &"_boarddata_" & Boardid).documentElement.selectSingleNode("boarddata/@board_setting").text,",")
			BoardType = Application(CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"&BoardID&"']/@boardtype").text
			BoardRootID = Application(CacheName &"_boarddata_" & Boardid).documentElement.selectSingleNode("boarddata/@rootid").text
			BoardParentID=CLng(Application(CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"&BoardID&"']/@parentid").text)	
			Sid = Application(CacheName &"_boarddata_" & Boardid).documentElement.selectSingleNode("boarddata/@sid").text
			Boardreadme=Application(CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"&BoardID&"']/@readme").text
			If Len(Board_Setting(22))< 24 Then Board_Setting(22)="1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1"
			OpenTime=Split(Board_Setting(22),"|")
			setting=Board_Setting(21)
			ischeck=Clng(Board_Setting(18))
			If Board_Setting(50)<>"0" And Board_Setting(50)<>"" Then 
				tourl=Board_Setting(50)
				Set Dvbbs=Nothing
				Response.Redirect tourl
			End If
		Else 
			Forum_ads = Split(CacheData(2,0),"$")
			If Len(Forum_Setting(70))< 24 Then Forum_Setting(70)="1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1"
			OpenTime=Split(Forum_Setting(70),"|")
			setting=Forum_Setting(69)
			ischeck=Forum_Setting(26)
			If Not IsNumeric(ischeck) Then ischeck=0
			ischeck=CLng(ischeck)		
		End If
		'定时开放判断
		If Not Page_Admin And Cint(setting)=1 Then
			If OpenTime(Hour(Now))="1" Then
				tourl="showerr.asp?action=stop&boardid="&Dvbbs.BoardID&""
				Set Dvbbs=Nothing:
				 Set Dvbbs=Nothing:Response.Redirect tourl
			End If 		
		End If
		'在线人数限制
		If ischeck > 0 And Not Page_Admin Then
			If MyBoardOnline.Forum_Online > ischeck And BoardID=0 Then
				tourl="showerr.asp?action=limitedonline&lnum="&ischeck
				If Not IsONline(Membername,1) Then Set Dvbbs=Nothing:Response.Redirect tourl
			End If
			If BoardID > 0 Then
				tourl="showerr.asp?action=limitedonline&lnum="&ischeck
				If (Not IsONline(Membername,1)) And MyBoardOnline.Board_Online > ischeck Then Set Dvbbs=Nothing:Response.Redirect tourl
			End If
		End If
		Dim CookiesSid
		CookiesSid = Request.Cookies("skin")("skinid_"&BoardID)
		If InStr(CookiesSid,"_")=0  Or CookiesSid = "" Then
			If BoardID = 0 Then 
				SkinID = Main_Sid
				CssID=CacheData(30,0)
			Else
				SkinID = Sid
				CssID=Application(CacheName &"_boarddata_" & Boardid).documentElement.selectSingleNode("boarddata/@cid").text
			End If
		Else
			CookiesSid=Split(CookiesSid,"_")
			CssID=CookiesSid(1)
			SkinID=CookiesSid(0)
		End If
		Setting=empty
	End Sub
	Public Function IsReadonly()
		IsReadonly=False
		Dim TimeSetting
		If Forum_Setting(69)="2" Then
			TimeSetting=split(Forum_Setting(70),"|")
			If TimeSetting(Hour(Now))="1" Then
				IsReadonly=True
				Exit Function
			End If
		End If
		If BoardID>0 Then 
			If Board_Setting(21)="2" Then
				TimeSetting=split(Board_Setting(22),"|")
				If TimeSetting(Hour(Now))="1" Then IsReadonly=True
			End If
		End If 
	End Function
	Public Function IsONline(UserName,action)
		IsONline=False
		If Trim(UserName)="" Then Exit Function
		If IsObject(Session(CacheName & "UserID")) And action=1 Then
				IsONline=True:Exit Function 
		End If
		Dim Rs
		Set Rs =Execute("Select UserID From Dv_Online Where Username='"&UserName&"'")
		If Not Rs.EOF  Then IsONline=True
		Set rs=Nothing  
	End Function  
	Public Sub LoadTemplates(Page_Fields)
		Dim Style_Pic,Main_Style,TempStyle,cssfilepath
		If Application(CacheName &"_style").documentElement.selectSingleNode("style[@id='"& SkinID &"']") Is Nothing Then
			If Not Application(CacheName &"_style").documentElement.selectSingleNode("style/@id") Is Nothing Then
				SkinID=Application(CacheName &"_style").documentElement.selectSingleNode("style/@id").text 
			Else
				Set Dvbbs=Nothing
				Response.Write "模板数据无法提取,请检查模板数据"
				Response.End
			End If
		End If
		Dim hascss
		If Application(CacheName & "_csslist").documentElement.selectSingleNode("css[@id='"& CssID &"' and tid='"& SkinID &"']") Is Nothing Then
			If Not Application(CacheName & "_csslist").documentElement.selectSingleNode("css[tid='"& SkinID &"']/@id") Is Nothing Then
				CssID=Application(CacheName & "_csslist").documentElement.selectSingleNode("css[tid='"& SkinID &"']/@id").text
				hascss=true
			ElseIf Not Application(CacheName & "_csslist").documentElement.selectSingleNode("css/@id") Is Nothing Then
				CssID=Application(CacheName & "_csslist").documentElement.selectSingleNode("css/@id").text
				cssfilepath=Application(CacheName & "_csslist").documentElement.selectSingleNode("@cssfilepath").text
				Forum_PicUrl=cssfilepath & Application(CacheName & "_csslist").documentElement.selectSingleNode("css[@id='"& CssID &"']/@picurl").text
			Else
				SkinID=Application(CacheName &"_style").documentElement.selectSingleNode("style/@id").text
				If Not Application(CacheName & "_csslist").documentElement.selectSingleNode("css[tid='"& SkinID &"']/@id") Is Nothing Then
					CssID=Application(CacheName & "_csslist").documentElement.selectSingleNode("css[tid='"& SkinID &"']/@id").text
					hascss=true
				Else
				CssID=Application(CacheName & "_csslist").documentElement.selectSingleNode("css/@id").text
				hascss=true
				End If
			End If
		Else
			hascss=true
		End If
		If hascss Then
			cssfilepath=Application(CacheName & "_csslist").documentElement.selectSingleNode("@cssfilepath").text
			Forum_PicUrl=cssfilepath & Application(CacheName & "_csslist").documentElement.selectSingleNode("css[@id='"& CssID &"' and tid='"& SkinID &"']/@picurl").text
			StyleName=Application(CacheName &"_style").documentElement.selectSingleNode("style[@id='"& SkinID &"']/@stylename").text
		End If
		Main_Style = Replace(Application(CacheName &"_style").documentElement.selectSingleNode("style[@id='"& SkinID &"']/@main_style").text,"{$PicUrl}",Forum_PicUrl)		'风格图片路径替换
		If Not (Instr(ScriptName,"index")>0 Or Page_Admin) Then
			Style_Pic = Replace(Application(CacheName &"_style").documentElement.selectSingleNode("style[@id='"& SkinID &"']/@style_pic").text,"{$PicUrl}",Forum_PicUrl)		'风格图片路径替换
			Style_Pic = Split(Style_Pic,"@@@")
			Forum_UserFace = Style_Pic(0)
			Forum_PostFace = Style_Pic(1)
			Forum_Emot = Style_Pic(2)
		End If
		If Page_Fields<>"" Then
			Template.value =Application(CacheName &"_style").documentElement.selectSingleNode("style[@id='"& SkinID &"']/@page_"& LCase(Page_Fields)).text
		End If
		Main_Style = Split(Main_Style,"@@@")
		mainhtml = Split(Main_Style(0),"|||")
		lanstr = Split(Main_Style(1),"|||")
		mainpic = Split(Main_Style(2),"|||")
		mainsetting = Split(mainhtml(0),"||")
		If hascss Then
			If Application(CacheName & "_csslist").documentElement.selectSingleNode("css[@id='"& CssID &"' and tid='"& SkinID &"']/@filename").text = "" Then
				Forum_CSS="<style type=""text/css"">" & Application(CacheName & "_csslist").documentElement.selectSingleNode("css[@id='"& CssID &"' and tid='"& SkinID &"']/cssdata").text &"</style>"
				Forum_CSS = Replace(Forum_CSS,"{$width}",mainsetting(0))
				Forum_CSS = Replace(Forum_CSS,"{$PicUrl}",Forum_PicUrl)
			Else
				Forum_CSS="<link rel=""stylesheet"" type=""text/css"" href="""& cssfilepath & Application(CacheName & "_csslist").documentElement.selectSingleNode("css[@id='"& CssID &"' and tid='"& SkinID &"']/@filename").text &".css"" />"
			End If
		Else
			Forum_CSS="<link rel=""stylesheet"" type=""text/css"" href="""& cssfilepath & Application(CacheName & "_csslist").documentElement.selectSingleNode("css[@id='"& CssID &"']/@filename").text &".css"" />"	
		End If
	End Sub
	Rem 判断发言是否来自外部
	Public Function ChkPost()
		Dim server_v1,server_v2
		Chkpost=False 
		server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
		server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
		If Mid(server_v1,8,len(server_v2))=server_v2 Then Chkpost=True 
	End Function
	Public Sub ReloadSetupCache(MyValue,N)'更新总设置表部分缓存数组，入口：更新内容、数组位置
		CacheData(N,0) = MyValue
		Name="setup"
		value=CacheData
	End Sub
	Public Sub NeedUpdateList(username,act)'更新用户资料缓存(缓存用户名,是否需要添加)[0=不添加,只作清理,1=需要添加]
		Dim Tmpstr,TmpUsername
		Name="NeedToUpdate"
		If ObjIsEmpty() Then Value=""
		Tmpstr=Value
		TmpUsername=","&username&","
		Tmpstr=Replace(Tmpstr,TmpUsername,",")
		Tmpstr=Replace(Tmpstr,",,",",")
		If act=1 Then 
			If IsONline(username,0) Then
				If Tmpstr="" Then
					Tmpstr=TmpUsername
				Else
					Tmpstr=Tmpstr&TmpUsername
				End If
			End If
		End If
		Tmpstr=Replace(Tmpstr,",,",",")
		Value=Tmpstr
	End Sub
	Public Sub LetGuestSession()'写入客人session
		Dim StatUserID,UserSessionID
		StatUserID = checkStr(Trim(Request.Cookies(Forum_sn)("StatUserID")))
		If IsNumeric(StatUserID) = 0 or StatUserID = "" Then
			StatUserID = Replace(UserTrueIP,".","")
			UserSessionID = Replace(Startime,".","")
			If IsNumeric(StatUserID) = 0 or StatUserID = "" Then StatUserID = 0
			StatUserID = Ccur(StatUserID) + Ccur(UserSessionID)
		End If
		StatUserID = Ccur(StatUserID)
		Response.Cookies(Forum_sn).Expires=DateAdd("s",3600,Now())
		Response.Cookies(Forum_sn).path=cookiepath
		Response.Cookies(Forum_sn)("StatUserID") = StatUserID
		Set UserSession=Application(Dvbbs.CacheName&"_info_guest").cloneNode(True)
		UserSession.documentElement.selectSingleNode("userinfo/@statuserid").text=StatUserID
		UserSession.documentElement.selectSingleNode("userinfo/@cometime").text=Now()
		UserSession.documentElement.selectSingleNode("userinfo/@activetime").text=DateAdd("s",-3600,Now())
		UserSession.documentElement.selectSingleNode("userinfo/@boardid").text=boardid
		Dim BS
		Set Bs=GetBrowser()
		UserSession.documentElement.appendChild(Bs.documentElement)
		If EnabledSession Then
			Session(CacheName & "UserID")=UserSession.xml
		End If
	End Sub 
	'根据页面来判断是否需要执行TrueCheckUserLogin
	Public Function NeedChecklongin()
		NeedChecklongin=True
		If UserID > 0 Then
			If InStr(ScriptName,"admin_")>0 Then Exit Function
			Dim pagelist
			pagelist=",post.asp,usermanager.asp,mymodify.asp,modifypsw.asp,modifyadd.asp,usersms.asp,"
			pagelist=pagelist & "friendlist.asp,favlist.asp,myfile.asp,friendlist.asp,recycle.asp,"
			pagelist=pagelist & "fileshow.asp,bbseven.asp,dispuser.asp,savepost.asp,plus_tools_pay.asp,joinvipgroup.asp,plus_tools_center.asp"
			If InStr(pagelist,","&ScriptName&",")>0 Then Exit Function
		End If
		NeedChecklongin=False
	End Function 
	'验证用户登陆
	Public Sub CheckUserLogin()
		If EnabledSession Then
			Set UserSession=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If Not UserSession.loadxml(Session(CacheName & "UserID")&"") Then
				If UserID > 0 Then 
					TrueCheckUserLogin
				Else
					Call LetGuestSession()
				End If
			Else
				If UserID >0 Or UserSession.documentElement.selectSingleNode("userinfo/@userid").text<>"0"  Then
					Dim NeedToUpdate,toupdate
					toupdate=False
					Name="NeedToUpdate"
					If Not ObjIsEmpty() Then 
						NeedToUpdate=","&Value&","
						If InStr(NeedToUpdate,","&MemberName&",")>0 Then
							Call NeedUpdateList(MemberName,0)
							toupdate=True
						End If
					End If
					If NeedChecklongin Or  toupdate Then TrueCheckUserLogin
				Else
		
				End If
			End If
		Else
			If UserID > 0 Then 
					TrueCheckUserLogin
				Else
					Call LetGuestSession()
			End If	
		End If
		UserID=CLng(UserSession.documentElement.selectSingleNode("userinfo/@userid").text)
		UserGroupID=CLng(UserSession.documentElement.selectSingleNode("userinfo/@usergroupid").text)
		If UserID > 0 Then
			GetCacheUserInfo
		Else
			UserGroupID = 7
			Lastlogin = Now()
		End If
		Browser=Checkstr(UserSession.documentElement.selectSingleNode("agent/@browser").text)
		version=replace(Checkstr(UserSession.documentElement.selectSingleNode("agent/@version").text),"--","")
		platform=Checkstr(UserSession.documentElement.selectSingleNode("agent/@platform").text)
		If (Browser="unknown" And version="unknown" And platform="unknown") Or Request("IsSearch")="1" Then
			If IsWebSearch Then
				IsSearch = True
			Else
				IsSearch = False
			End If
			If Request("IsSearch") = "1" Then IsSearch = True
			Cls_IsSearch = True
		End If
		'IP锁定
		If UserSession.documentElement.selectSingleNode("agent/@lockip").text="1"  Then
			If Not Page_Admin Then Set Dvbbs=Nothing:Response.Redirect "showerr.asp?action=iplock"
			'If Not Page_Admin Then Session(CacheName & "UserID")=empty:Response.Status = "302 Object Moved" 
		End If	
		GetGroupSetting
	End Sub
	Rem xmlroot跟节点名称 row记录行节点名称
	Public Function RecordsetToxml(Recordset,row,xmlroot)
		Dim i,node,rs,j,DataArray
		If xmlroot="" Then xmlroot="xml"
		If row="" Then row="row"
		Set RecordsetToxml=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		RecordsetToxml.appendChild(RecordsetToxml.createElement(xmlroot))
		If Not Recordset.EOF Then
			DataArray=Recordset.GetRows(-1)
			For i=0 To UBound(DataArray,2)
				Set Node=RecordsetToxml.createNode(1,row,"")
				j=0
				For Each rs in Recordset.Fields
						 node.attributes.setNamedItem(RecordsetToxml.createNode(2,LCase(rs.name),"")).text= DataArray(j,i)& ""
						 j=j+1
				Next
				RecordsetToxml.documentElement.appendChild(Node)
			Next
		End If
		DataArray=Null
	End Function
	Public Function ArrayToxml(DataArray,Recordset,row,xmlroot)
		Dim i,node,rs,j
		If xmlroot="" Then xmlroot="xml"
		Set ArrayToxml=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		ArrayToxml.appendChild(ArrayToxml.createElement(xmlroot))
		If row="" Then row="row"
		For i=0 To UBound(DataArray,2)
			Set Node=ArrayToxml.createNode(1,row,"")
			j=0
			For Each rs in Recordset.Fields
					 node.attributes.setNamedItem(ArrayToxml.createNode(2,LCase(rs.name),"")).text= DataArray(j,i)& ""
					 j=j+1
			Next
			ArrayToxml.documentElement.appendChild(Node)
		Next
	End Function
	Public Function Createpass()'系统分配随机密码
		Dim Ran,i,LengthNum
		LengthNum=16
		Createpass=""
		For i=1 To LengthNum
			Randomize
			Ran = CInt(Rnd * 2)
			Randomize
			If Ran = 0 Then
				Ran = CInt(Rnd * 25) + 97
				Createpass =Createpass& UCase(Chr(Ran))
			ElseIf Ran = 1 Then
				Ran = CInt(Rnd * 9)
				Createpass = Createpass & Ran
			ElseIf Ran = 2 Then
				Ran = CInt(Rnd * 25) + 97
				Createpass =Createpass& Chr(Ran)
			End If
		Next
	End Function
	Public Sub NewPassword()'更新用户验证密码
		If UserID=0 Then Exit Sub
		Response.Write "<iframe style=""border:0px;width:0px;height:0px;""  src=""newpass.asp"" name=""Dvnewpass""></iframe>"
	End Sub
	Public Sub TrueCheckUserLogin()
		Dim Rs,SQL,FoundMyGroupID
		FoundMyGroupID = 0
		Sql="Select UserID,UserName,UserPassword,UserEmail,UserPost,UserTopic,UserSex,UserFace,UserWidth,UserHeight,JoinDate,LastLogin as cometime ,LastLogin,LastLogin as activetime,UserLogins,Lockuser,Userclass,UserGroupID,UserGroup,userWealth,userEP,userCP,UserPower,UserBirthday,UserLastIP,UserDel,UserIsBest,UserHidden,UserMsg,IsChallenge,UserMobile,TitlePic,UserTitle,TruePassWord,UserToday,UserMoney,UserTicket,FollowMsgID,Vip_StarTime,Vip_EndTime,userid as boardid"
		Sql=Sql & " From [Dv_User] Where UserID = " & UserID
		Set Rs = Execute(Sql)
		If Rs.EOF Then
			UserID = 0:LetGuestSession():Exit Sub
		Else
			If Not (LCase(Rs("UserName"))=LCase(Membername) and Rs("TruePassWord")=Memberword) Then
				If EnabledSession Then
					Set UserSession=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
					If UserSession.loadxml(Session(CacheName & "UserID")&"")  Then
						If UserSession.documentElement.selectSingleNode("userinfo/@username") Is Nothing Or UserSession.documentElement.selectSingleNode("userinfo/@userpassword") Is Nothing Then
							UserID = 0:LetGuestSession():Exit Sub
						Else
							If Not (LCase(Rs("UserName"))=LCase(UserSession.documentElement.selectSingleNode("userinfo/@username").text) and Rs("UserPassword")=UserSession.documentElement.selectSingleNode("userinfo/@userpassword").text) Then
									UserID = 0:LetGuestSession():Exit Sub
							End If
						End If
					Else
						UserID = 0:LetGuestSession():Exit Sub
					End If
				Else
							UserID = 0:LetGuestSession():Exit Sub
				End If
			End If
			If Rs("LockUser")=1 Then
				UserID = 0:LetGuestSession():Exit Sub
			End if
		End If
		Set UserSession=RecordsetToxml(rs,"userinfo","xml")
		UserSession.documentElement.selectSingleNode("userinfo/@cometime").text=Now()
		UserSession.documentElement.selectSingleNode("userinfo/@activetime").text=DateAdd("s",-3600,Now())
		UserSession.documentElement.selectSingleNode("userinfo/@boardid").text=boardid
		UserSession.documentElement.selectSingleNode("userinfo").attributes.setNamedItem(UserSession.createNode(2,"isuserpermissionall","")).text=FoundUserPermission_All()
		If Not (UserSession.documentElement.selectSingleNode("userinfo/@usergroupid2") is Nothing )  Then
				FoundMyGroupID =  CLng(UserSession.documentElement.selectSingleNode("userinfo/@usergroupid2").text)
		End If	
		If FoundMyGroupID > 0 Then
			UserSession.documentElement.selectSingleNode("userinfo").attributes.setNamedItem(UserSession.createNode(2,"usergroupid2","")).text=FoundMyGroupID
		End If
		Dim BS
		Set Bs=GetBrowser()
		UserSession.documentElement.appendChild(Bs.documentElement)
		If EnabledSession Then
			Session(CacheName & "UserID")= UserSession.xml
		End If
		Set Rs=Nothing
		GetCacheUserInfo()
	End Sub
	Public Sub GetCacheUserInfo()	'用户登录成功后，采用本函数读取用户数组并判断一些常用信息
		UserID = Clng(UserSession.documentElement.selectSingleNode("userinfo/@userid").text)
		MemberName = UserSession.documentElement.selectSingleNode("userinfo/@username").text
		Lastlogin = UserSession.documentElement.selectSingleNode("userinfo/@lastlogin").text
		If Not IsDate(LastLogin) Then LastLogin = Now()
		UserGroupID = CLng(UserSession.documentElement.selectSingleNode("userinfo/@usergroupid").text)
		If Trim(UserSession.documentElement.selectSingleNode("userinfo/@usertoday").text)="" Then
			Execute("Update [Dv_User] Set UserToday='0|0|0|0|0' Where UserID = " & UserID)
			UserSession.documentElement.selectSingleNode("userinfo/@usertoday").text="0|0|0|0|0"
			UserToday = Split("0|0|0|0|0","|")
		Else
			UserToday = Split(UserSession.documentElement.selectSingleNode("userinfo/@usertoday").text,"|")
			If Ubound(UserToday) <> 4 Then
				Execute("Update [Dv_User] Set UserToday='0|0|0|0|0' Where UserID = " & UserID)
				UserSession.documentElement.selectSingleNode("userinfo/@usertoday").text="0|0|0|0|0"
				UserToday = Split("0|0|0|0|0","|")
			End If
		End If
		'判断是否VIP组成员
			If IsDate(UserSession.documentElement.selectSingleNode("userinfo/@vip_startime").text) Then
				If DateDiff("d",Now(),UserSession.documentElement.selectSingleNode("userinfo/@vip_endtime").text)>0 Then
					VipGroupUser = True
				Else
					Dim tRs
					'将已过期的VIP用户移回注册组并清空有效时间
					If UserGroupID>8 Then
						Set tRs=Execute("Select Top 1 * From Dv_UserGroups Where ParentGID=3 And MinArticle<="& CCur(UserSession.documentElement.selectSingleNode("userinfo/@userpost").text) &" Order By MinArticle Desc")
							If not tRs.Eof Then
								Execute("Update Dv_User Set UserClass='"&tRs("UserTitle")&"',TitlePic='"&tRs("GroupPic")&"',UserGroupID="&tRs("UserGroupID")&",Vip_StarTime=null,Vip_EndTime=null Where UserID="&UserID)
							End If
						Set tRs=Nothing
					Else
						Execute("Update Dv_User Set Vip_StarTime=null,Vip_EndTime=null Where UserID="&UserID)
					End If
					UserSession.documentElement.selectSingleNode("userinfo/@vip_startime").text = ""
					UserSession.documentElement.selectSingleNode("userinfo/@vip_endtime").text =""
				End If
		End If
		Select Case UserGroupID
		Case 8
			Vipuser = True
		Case 3
			If BoardID=0 Then 	Boardmaster = True
		Case 2
			Superboardmaster = True
			Boardmaster = True
		Case 1
			Master = True
			Boardmaster = True
		End Select
		If UserSession.documentElement.selectSingleNode("userinfo/@ischallenge").text  = "1" Then FoundIsChallenge = True
		If DateDiff("d",LastLogin,Now())<>0 Then
			Execute("Update [Dv_User] Set UserToday='0|0|0|0|0',LastLogin = " & SqlNowString & " Where UserID = " & UserID)
			UserSession.documentElement.selectSingleNode("userinfo/@usertoday").text = "0|0|0|0|0"
			LastLogin = Now()
		End If
		If Userhidden = 2 and DateDiff("s",Lastlogin,Now())>Clng(Forum_Setting(8))*60 Then
			Execute("Update [Dv_User] Set UserLastIP = '" & UserTrueIP & "',LastLogin = " & SqlNowString & " Where UserID = " & UserID)
			Lastlogin = Now()
		End If
		sendmsgnum=0:sendmsgid=0:sendmsguser=""
		If UserSession.documentElement.selectSingleNode("userinfo/@usermsg").text<>"" Then
			Dim Usermsg
			Usermsg=Split(UserSession.documentElement.selectSingleNode("userinfo/@usermsg").text,"||")
			If Ubound(Usermsg)=2 Then
				sendmsgnum=Usermsg(0)
				sendmsgid=Usermsg(1)
				sendmsguser=Usermsg(2)
			End If
		End If
		'跟踪用户处理
		Dim FollowMsgID
		Set FollowMsgID=UserSession.documentElement.selectSingleNode("userinfo/@followmsgid")
		If Not ( FollowMsgID Is Nothing) Then
		If FollowMsgID.text <>"" Then
			Dim ToolsFollowUserID,i,Rs,Tools_inceptid,Tools_newincept,Tools_msginfo
		ToolsFollowUserID = Split( FollowMsgID.text,",")
		For i=0 To Ubound(ToolsFollowUserID)
				If Len(ToolsFollowUserID(i))>0 and Len(ToolsFollowUserID(i))<50 and ToolsFollowUserID(i)<>"" Then
					ToolsFollowUserID(i) = CheckStr(ToolsFollowUserID(i))
						Execute("Insert into Dv_Message (incept,sender,title,content,sendtime,flag,issend) values ('"& ToolsFollowUserID(i)&"','系统消息','您跟踪的用户"&Dvbbs.MemberName&"已登录','您使用了论坛道具“狗仔队”，您所跟踪的用户 "&Dvbbs.Membername&" 于 "&Now()&" 登录了论坛，请您及时和该用户取得联系，感谢您采用我们的服务。',"&SqlNowString&",0,1)")
					Set Rs=Execute("Select top 1 id,sender From Dv_Message Where incept ='"& ToolsFollowUserID(i) &"'")
						Tools_inceptid=Rs(0) &"||"& Rs(1)
						Set Rs=Execute("Select Count(id) From Dv_Message Where Flag=0 and issend=1 and delR=0 And incept='"& ToolsFollowUserID(i) &"'")
						Tools_newincept = Rs(0)
						Set Rs=Nothing
					If IsNull(Tools_newincept) Then Tools_newincept=0
						Tools_msginfo=Tools_newincept & "||" & Tools_inceptid
						Execute("update [dv_user] set UserMsg='"&CheckStr(Tools_msginfo)&"' where username='"&ToolsFollowUserID(i)&"'")
				End If
			Next
			 FollowMsgID.text = ""
			Execute("UpDate Dv_User Set FollowMsgID='' Where UserID="&UserID)
		End If
		End If
		FoundUser=True
		UserSession.documentElement.selectSingleNode("userinfo/@lastlogin").text=Lastlogin
		Dim iUserMagicFace'用户头像处理
		iUserMagicFace = Split(UserSession.documentElement.selectSingleNode("userinfo/@userface").text,"|")
		If Ubound(iUserMagicFace) = 1 Then UserSession.documentElement.selectSingleNode("userinfo/@userface").text = iUserMagicFace(1)
	End Sub
	Private Sub GetGroupSetting()
		If Not IsObject(Application(CacheName &"_groupsetting")) Then LoadGroupSetting()
		If Application(CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& UserGroupID &"']/@groupsetting") Is nothing Then UserGroupID=7
		GroupSetting = Split(Application(CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& UserGroupID &"']/@groupsetting").text,",")
		If ScriptName="reg.asp"  or ScriptName ="login.asp" or Page_Admin Then GroupSetting(0)=1
		If Cint(GroupSetting(0))=0  Then AddErrCode "8":Showerr()
		UserGroupParent = Cint(Application(CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& UserGroupID &"']/@parentgid").text)
		UserGroupParentID=Split(Application(CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& UserGroupID &"']/@issetting").text,"|")
		If UserID > 0 Then IsUserPermissionAll = CLng(UserSession.documentElement.selectSingleNode("userinfo/@isuserpermissionall").text)
		If BoardID > 0 And Not ScriptName="showerr.asp" Then CheckBoardInfo()
		If UserID > 0 And BoardID=0 Then
			If IsUserPermissionAll="1" Then LoadUserPermission_All()
		End If
			If Not (UserSession.documentElement.selectSingleNode("userinfo/@usergroupid2") is Nothing )  Then
				If  CLng(UserSession.documentElement.selectSingleNode("userinfo/@usergroupid2").text)	>0 Then
					IsUserPermissionOnly = 1
				End If
			End If
			'If GroupSetting(70)="1"  Then
			'	Master = True
			'Else
			'	Master = False
			'End If
	End Sub
	'用户是否存在论坛全局自定义权限
	Public Function FoundUserPermission_All()
		Dim PerRs
		FoundUserPermission_All = 0
		Set PerRs=Execute("Select Uc_Setting From Dv_UserAccess Where Uc_Boardid=0 And uc_UserID= "& UserID )
		If Not (PerRs.Eof And PerRs.Bof) Then FoundUserPermission_All = 1
		PerRs.Close:Set PerRs=Nothing
	End Function
	Public Sub LoadUserPermission_All()
		Dim Rs
		Set Rs=Dvbbs.execute("Select Uc_Setting From Dv_UserAccess Where Uc_Boardid=0 And uc_UserID="&UserID)
		If Not(Rs.Eof And Rs.Bof) Then
			UserPermission=Split(Rs(0),",")
			GroupSetting = Split(Rs(0),",")
			FoundUserPer=True
		End If
		Set Rs=Nothing
	End Sub
	Public Sub ActiveOnline()
		'当在120秒内刷新同一个页面则不更新online数据
		If Not IsNumeric(UserSession.documentElement.selectSingleNode("userinfo/@boardid").text) Or UserSession.documentElement.selectSingleNode("userinfo/@boardid").text="" Then UserSession.documentElement.selectSingleNode("userinfo/@boardid").text="0"
		If DateDiff("s",UserSession.documentElement.selectSingleNode("userinfo/@activetime").text,Now()) < 120 And CLng(UserSession.documentElement.selectSingleNode("userinfo/@boardid").text) = BoardID  And Not InStr(ScriptName,"showerr")>0 Then Exit Sub
		'更新数组
		UserSession.documentElement.selectSingleNode("userinfo/@activetime").text=Now()
		UserSession.documentElement.selectSingleNode("userinfo/@boardid").text=boardid
		UserActiveOnline
		'新增更新用户最后登录时间，以保证贴子中在线判断的准确性
		If UserSession.documentElement.selectSingleNode("userinfo/@userid").text <> "0" Then
			If UserSession.documentElement.selectSingleNode("userinfo/@userhidden").text="2" Then
				Dvbbs.execute("update [Dv_user] set lastlogin=" & SqlNowString & " where userid="&Dvbbs.userid)
			End If
		End If
	End Sub
	Private Sub UserActiveOnline()
		Dim Actcome,SQl,Rs
		Dim uip,StatsStr
			uip = UserTrueIP
        	StatsStr = Stats
        	StatsStr = Replace(StatsStr, "'", "")
        	StatsStr = Replace(StatsStr, Chr(0), "")
        	StatsStr = Replace(StatsStr, "--", "——")
        	StatsStr = Left(StatsStr, 250)
		If UserID = 0 Then
			Dim StatUserID
			StatUserID = UserSession.documentElement.selectSingleNode("userinfo/@statuserid").text
			SQL = "Select ID,Boardid From [Dv_Online] Where ID = " & Ccur(StatUserID)
			Set Rs = Execute(SQL)
			If Rs.EOF  Then
				If IP_MAX>0 Then
					If Onlineip(UserTrueIP) > IP_MAX Then
						Session(CacheName & "UserID")=empty
						Set Dvbbs=Nothing
						Response.Status = "302 Object Moved" 
						Response.End  	
	  			End If
  			End if
				If CInt(Forum_Setting(36)) = 0 Then
					Actcome = ""
				Else
					Actcome = address(uip)
				End If
				If Cls_IsSearch Then Exit Sub  '不记录搜索引擎的客人 2004-8-30 Dv.Yz
				SQL = "Insert Into [Dv_Online](ID,Username,Userclass,Ip,Startime,Lastimebk,Boardid,Browser,Stats,Usergroupid,Actcome,Userhidden,actforip) Values (" & StatUserID & ",'客人','客人','" & UserTrueIP & "'," & SqlNowString & "," & SqlNowString & "," & Boardid & ",'" & platform&"|"&Browser&version & "','" & StatsStr & "',7,'" & Actcome & "'," & Userhidden & ",'"& checkstr(actforip)&"')"
				'更新缓存总在线数据
				MyBoardOnline.Forum_Online=MyBoardOnline.Forum_Online+1
				Name="Forum_Online"
				value=MyBoardOnline.Forum_Online 
			Else
				SQL = "Update [Dv_Online] Set Lastimebk = " & SqlNowString & ",Boardid = " & Boardid & ",Stats = '" & StatsStr & "' Where ID = " & Ccur(StatUserID)
			End If
			Rs.Close
			Set Rs = Nothing
			Execute(SQL)
		Else
			SQL = "Select ID,Boardid From [DV_Online] Where UserID = " & UserID
			Set Rs = Execute(SQL)
			If Rs.Eof And Rs.Bof Then
				If CInt(forum_setting(36)) = 0 Then
					Actcome = ""
				Else
					Actcome = address(uip)
				End If
				SQL = "Insert Into [Dv_Online](ID,Username,Userclass,Ip,Startime,Lastimebk,Boardid,Browser,Stats,Usergroupid,Actcome,Userhidden,UserID,actforip) Values (" & Session.SessionID & ",'" & Membername & "','" & Memberclass & "','" & UserTrueIP & "'," & SqlNowString & "," & SqlNowString & "," & Boardid & ",'" & platform&"|"&Browser&version & "','" & StatsStr & "'," & UserGroupID & ",'" & Actcome & "'," & Userhidden & "," & UserID & ",'"& checkstr(actforip)&"')"
				'更新缓存总在线数据
				MyBoardOnline.Forum_Online=MyBoardOnline.Forum_Online+1
				Name="Forum_Online"
				Dvbbs.value=MyBoardOnline.Forum_Online
				'更新缓存总用户在线数据
				MyBoardOnline.Forum_UserOnline=MyBoardOnline.Forum_UserOnline+1
				Name="Forum_UserOnline"
				value=MyBoardOnline.Forum_UserOnline
			Else
				SQL = "Update [Dv_Online] Set Lastimebk = " & SqlNowString & ",Boardid = " & Boardid & ",Stats = '" & StatsStr & "' Where UserID = " & UserID
			End If
			Rs.Close
			Set Rs = Nothing
			Execute(SQL)
		End If	
		'更新在线峰值
		If CLng(MyBoardOnline.Forum_Online) > CLng(Maxonline) Then
			Execute("update [Dv_setup] set Forum_Maxonline="&CLng(MyBoardOnline.Forum_Online)&",Forum_MaxonlineDate="& SqlNowString) 
			CacheData(5,0)=MyBoardOnline.Forum_Online
			CacheData(6,0)=Now()
			Name="setup"
			value=CacheData
		End If
		Rem 删除超时用户
		MyBoardOnline.OnlineQuery
	End Sub
	'去掉HTML标记
	Public Function Replacehtml(Textstr)
		Dim Str,re
		Str=Textstr
		Set re=new RegExp
			re.IgnoreCase =True
			re.Global=True
			re.Pattern="<(.[^>]*)>"
			Str=re.Replace(Str, "")
			Set Re=Nothing
			Replacehtml=Str
	End Function
	Function Onlineip(ip)
		Dim SQl
		SQL="Select Count(*) From Dv_online where ip='"&ip&"'"
		Onlineip=Execute(SQL)(0)
		If IsNull(Onlineip) Then Onlineip=0
	End Function
	Public Sub Nav()
		Head()
		ShowTopTable()
		IsTopTable = 1
	End Sub
	Public Sub head()
		Nowstats=stats
		If ScriptName="index.asp" Then
			If BoardType<>"" Then Stats=BoardType & Left(Replacehtml(boardreadme),20)&"...."
		ElseIf ScriptName <> "dispbbs.asp" Then
			If BoardType<>"" Then Stats=BoardType&"-"&Stats
		End If
		Stats=Replace(Stats,Chr(13),"")
		stats=Replacehtml(stats)
		'搜索引擎优化部分
		If Request("IsSearch_a") <> "" Then stats = stats & "-网站地图"
		Nowstats=Replacehtml(Nowstats)
		If IsSearch Then
			Response.Write Replace(Replace(Replace(mainhtml(1),"{$keyword}",Replace(Forum_info(8),"|",",")),"{$description}",Forum_info(10))&vbNewLine,"{$title}",stats &"["& Forum_Info(0) &"] -- Powered By Dvbbs.net," & Now())
		Else
			Response.Write Replace(Replace(Replace(mainhtml(1),"{$keyword}",Replace(Forum_info(8),"|",",")),"{$description}",Forum_info(10))&vbNewLine,"{$title}",stats &"["& Forum_Info(0) &"]")
		End If
		'搜索引擎优化结束
		Response.Write Forum_CSS
		Response.Write Chr(10)
		If Boardid=0 Then
			Response.Write "<link title="""& Forum_Info(0) &"-频道列表"" type=""application/rss+xml"" rel=""alternate"" href=""rssfeed.asp"" />"
			Response.Write Chr(10)
			Response.Write "<link title="""& Forum_Info(0) &"-最新20篇论坛主题"" type=""application/rss+xml"" rel=""alternate"" href=""rssfeed.asp?rssid=4"" />"
		Else
			Response.Write "<link title="""& Replacehtml(BoardType) &"-最新20篇论坛主题"" type=""application/rss+xml"" rel=""alternate"" href=""rssfeed.asp?boardid="&boardid&"&amp;rssid=4"" />"
		End If
		Response.Write Chr(10)
		Response.Write mainhtml(2)
		Dim node,XMLDOM
		If Not Cls_IsSearch Then
			Set XMLDOM=Application(CacheName&"_ssboardlist").cloneNode(True)
			If Dvbbs.GroupSetting(37)="0"  Then'去掉隐藏论坛
				For each node in XMLDOM.documentElement.getElementsByTagName("board")
					If node.attributes.getNamedItem("hidden").text="1" Then
						node.parentNode.removeChild(node)
					End If
				Next
			End If
			Response.Write "<script language=""javascript"" type=""text/javascript"">"
			Response.Write "var boardxml='<?xml version=""1.0"" encoding=""gb2312""?>"& replace(XMLDom.documentElement.XML ,"'","\'")&"';"
			Response.Write "</script>"
		End If
		If Cls_IsSearch Then Exit Sub
		If Forum_Setting(19)="1" And Not Page_Admin Then
			If ( Trim(Forum_Setting(64))<>"" And InStr(LCase(Forum_Setting(64)),ScriptName) >0 And Cint(Forum_Setting(20))>0) Then
				If DateDiff("s",UserSession.documentElement.selectSingleNode("userinfo/@activetime").text,Now())< Cint(Forum_Setting(20)) and boardid=CLng(UserSession.documentElement.selectSingleNode("userinfo/@boardid").text) and InStr(LCase(Cstr(Request.ServerVariables("HTTP_REFERER"))),ScriptName) > 0 Then
					Response.Write "<div style=""margin-top:24px;text-align:left;text-indent :24px;"">本页面启用了防刷新设置,请不要在"& Forum_Setting(20) &"秒内连续刷新本页面.</div>"
					Response.Write "<div style=""text-align:left;text-indent :24px;"">"&(Forum_Setting(20)-DateDiff("s",UserSession.documentElement.selectSingleNode("userinfo/@activetime").text,Now()))&"秒后将重新加载页面请稍等....</div>"
					Response.Write "<script language=""javascript"" type=""text/javascript"">"
					Response.Write "setTimeout('location.reload();',"&(Forum_Setting(20)-DateDiff("s",UserSession.documentElement.selectSingleNode("userinfo/@activetime").text,Now()))*1000&");"
					Response.Write "</script>"
					Response.Write "</body></html>"
					Set Dvbbs=Nothing
					Response.End
				Else
						UserSession.documentElement.selectSingleNode("userinfo/@activetime").text=Now()
						UserSession.documentElement.selectSingleNode("userinfo/@boardid").text=Boardid
				End If
			End If 
		End If
	End Sub 
	Public Sub ShowTopTable()
		Dim TempStr,ForumMenu,Tempstr1
		Dim RayMenuInfo,RayMenu
		If IsSearch Then
			'搜索引擎优化部分
			'加入针对搜索引擎的导航栏,同时增加官方链接(可增加自身网站PageRank以及给官方增加搜索引擎友好度),请勿擅自取消
			sysmenu = mainhtml(20)
			Dim node,XMLDom,iTempStr
			Set XMLDOM=Application(Dvbbs.CacheName&"_sboardlist").cloneNode(True)
			iTempStr = "&nbsp;&nbsp;"
			For each node in XMLDOM.documentElement.selectNodes("board[@parentid=0]")
				iTempStr = iTempStr & "<a href=""index.asp?IsSearch_a=2&BoardID="&Node.selectSingleNode("@boardid").text&""">" & Node.selectSingleNode("@boardtype").text & "</a>&nbsp;&nbsp;"
			Next
			sysmenu = Replace(sysmenu,"{$catlist}",iTempStr & "<a href=""http://www.cndw.com"" title=""论坛,bbs,免费论坛,国内最大的论坛软件服务提供商,blog,boke,博客,防火墙,插件"">动网论坛</a>&nbsp;&nbsp;<a href=""http://bbs.cndw.com"" title=""论坛,bbs,免费论坛,国内最大的论坛软件服务提供商官方讨论区,blog,boke,博客,asp,asp.net,电脑,软件,灌水,防火墙,开发,插件"">官方论坛</a>&nbsp;&nbsp;<a href=http://tool.cndw.com>站长工具</a>")
			'搜索引擎优化结束
		ElseIf UserID = 0 Then 
			sysmenu = mainhtml(7)
		Else
			sysmenu = Replace(mainhtml(6),"{$username}",Membername)
			If UserHidden=2 Then
				sysmenu = Replace(sysmenu,"{$hiddeninfo}",lanstr(3))
			Else
				sysmenu = Replace(sysmenu,"{$hiddeninfo}",lanstr(4))
			End If
			If Master Or GroupSetting(70)="1" Then
				sysmenu = Replace(sysmenu,"{$manageinfo}",mainhtml(10))
			ElseIf Superboardmaster Then
				sysmenu = Replace(sysmenu,"{$manageinfo}",mainhtml(19))
			Else
				sysmenu = Replace(sysmenu,"{$manageinfo}","")
			End If
			'If Forum_ChanSetting(0)="1" Then
			'	RayMenuInfo = Split(mainhtml(11),"||")
			'	RayMenu = Replace(Replace(RayMenuInfo(4),"{$channame}",CacheData(23,0)),"{$forumurl}",Forum_Info(1))
			'	If FoundIsChallenge Then
			'		RayMenu = RayMenu & RayMenuInfo(2)
			'	Else
			'		RayMenu = RayMenu & RayMenuInfo(3)
			'	End If
			'	RayMenu = Replace(RayMenuInfo(1),"{$raymenu}",RayMenu)
			'	sysmenu = Replace(sysmenu,"{$raymenuinfo}",RayMenuInfo(0))
			'Else
			sysmenu = Replace(sysmenu,"{$raymenuinfo}","")
			'End If
			sysmenu = Replace(sysmenu,"{$userid}",UserID)
			'Response.Write RayMenu
		End If
		If Dvbbs.Forum_setting(99) = "1" Then
			sysmenu = Replace(sysmenu,"{$isboke}",mainhtml(21))
		Else
			sysmenu = Replace(sysmenu,"{$isboke}","")
		End If
		If Forum_Setting(90)=0 Then 
			sysmenu = Replace(sysmenu,"{$Plus_Tools}","")
		Else
			sysmenu = Replace(sysmenu,"{$Plus_Tools}",mainhtml(16))
		End If
		If GroupSetting(57) = "1" Then
			Name = "style_list"
			Tempstr1=replace(Value,"$boardid",boardid)
			If Dvbbs.BoardID = 0 Then
				TempStr1 = Replace(TempStr1,"{$dskinid}",CacheData(17,0))
			Else
				TempStr1 = Replace(TempStr1,"{$dskinid}",Sid)
			End If
		Else
			mainhtml(9)=Replace(Replace(Replace(Replace(mainhtml(9),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")
			mainhtml(9) = Split(mainhtml(9),"||")
			Tempstr1=Replace(Replace(mainhtml(9)(0),"{$dskinid}",CacheData(17,0)),"{$csslist}","")
		End If
		sysmenu = Replace(sysmenu,"{$syles}",Tempstr1)
		TempStr = TempStr & Chr(10) & mainhtml(4)
		TempStr = Replace(TempStr,"{$width}",mainsetting(0))
		TempStr = Replace(TempStr,"{$link}",Forum_Info(1))
		If Boardid>0 Then 
			If Board_Setting(51)="" Or Board_Setting(51) = "0"  Then
				TempStr = Replace(TempStr,"{$logo}",Forum_Info(6))
			Else
				TempStr = Replace(TempStr,"{$logo}",Board_Setting(51))
			End If
		Else
			TempStr = Replace(TempStr,"{$logo}",Forum_Info(6))
		End If
		If Trim(Forum_info(7))<>"0" And Trim(Forum_info(7))<>""  Then
			TempStr = Replace(TempStr,"{$mailto}",Forum_Info(7))
		Else
			TempStr = Replace(TempStr,"{$mailto}","mailto:" & Forum_Info(5))
		End If
		TempStr = Replace(TempStr,"{$title}",Forum_Info(0) & "-" & Replace(stats,"'","\'"))
		TempStr = Replace(TempStr,"{$top_ads}",Forum_ads(0))
		TempStr = Replace(TempStr,"{$menu}",Chr(10) & sysmenu)
		TempStr = Replace(TempStr,"{$boardid}",boardid)
		TempStr = Replace(TempStr,"{$alertcolor}",mainsetting(1))
		Name = "ForumPlusMenu"
		If ObjIsEmpty Then LoadPlusMenu()
		ForumMenu = Value
		If ForumMenu <> "" Then
			TempStr = Replace(TempStr,"{$plusmenu}",ForumMenu)
		Else
			TempStr = Replace(TempStr,"{$plusmenu}","")
		End If
		Response.Write TempStr
		TempStr = ""
	End Sub 
	Public Sub Head_var(IsBoard,idepth,GetTitle,GetUrl)
		Dim NavStr,AllBoardList
		If Dvbbs.BoardID=0 Then BoardReadme=lanstr(2) & " <b>" & Forum_Info(0) & "</b>"
		If BoardID>0 Then
			NavStr = " <a href="""&Forum_Info(11)&""" onmouseover=""showmenu(event,BoardJumpList(0),'',0);"" style=""cursor:hand;"">"&Forum_info(0)&"</a> → "
		Else
			NavStr = " <a href="""&Forum_Info(11)&""">"&Forum_info(0)&"</a> → "
		End If
		If IsBoard=1 Then
			If BoardParentID=0 Then
				NavStr = NavStr & " <a href=""index.asp?boardid="&BoardID&""" onMouseOver=""showmenu(event,BoardJumpList("&Dvbbs.Boardid&"),'',0);"">"&BoardType&"</a>"
			Else
				If ScriptName="dispbbs.asp" Then 
					NavStr = NavStr & BoardInfoData & " → <a href=""index.asp?boardid="&BoardID&"&page="&Request("page")&""">"&BoardType&"</a>"
				Else
					NavStr = NavStr & BoardInfoData & " → <a href=""index.asp?boardid="&BoardID&""">"&BoardType&"</a>"
				End If
			End If
			NavStr = NavStr & " → " & Nowstats
		Elseif IsBoard=2 Then
			NavStr = NavStr & Nowstats
		Else
			NavStr = NavStr & "<a href="&GetUrl&">"&GetTitle&"</a> → " & Nowstats
		End If
		NavStr = Replace(mainhtml(5),"{$nav}",NavStr)
		NavStr = Replace(NavStr,"{$width}",mainsetting(0))
		NavStr = Replace(NavStr,"{$boardreadme}",BoardReadme)
		If ScriptName="dispbbs.asp" Then
			If Second(Now()) Mod 2 = 0 Then
				NavStr = Replace(NavStr,"{$SearchStr}","<a href=""query.asp?stype=8&keyword="&Server.HtmlEncode(Nowstats&"")&"&isWeb=2"" target=""_blank"" title=""在更多网站中搜索此类问题,搜索和查看更多相关的精彩主题""><font color=""green""><b>搜一搜更多此类问题</b></font></a>")
			Else
				NavStr = Replace(NavStr,"{$SearchStr}","<a href=""query.asp?stype=8&keyword="&Server.HtmlEncode(Nowstats&"")&"&isWeb=2"" target=""_blank"" title=""在更多网站中搜索此类问题,搜索和查看更多相关的精彩主题""><font class=""redfont""><b>搜一搜相关精彩主题</b></font></a>")
			End If
		Else
			NavStr = Replace(NavStr,"{$SearchStr}","")
		End If
		If UserID>0 Then
			'sendmsgnum,sendmsgid,sendmsguser
			IsBoard = Split(mainhtml(12),"||")
			If Clng(SendMsgNum)>0 Then
				BoardReadme = IsBoard(0)
				If Forum_Setting(10)=1 Then
					BoardReadme = BoardReadme & IsBoard(1) & IsBoard(2)
				Else
					BoardReadme = BoardReadme & IsBoard(2)
				End If
				BoardReadme = Replace(BoardReadme,"{$smsid}",sendmsgid)
				BoardReadme = Replace(BoardReadme,"{$sender}",sendmsguser)
				BoardReadme = Replace(BoardReadme,"{$newmsgnum}",sendmsgnum)
			Else
				BoardReadme = IsBoard(3)
			End If
			Dim i,UserGroupList,iGroupName
			If UserGroupParent = 4 Then
				BoardReadme = BoardReadme & IsBoard(4)
				For i = 0 To Ubound(UserGroupParentID)
					If i = 0 Then
						UserGroupList = "<a href=""cookies.asp?action=ReGroup&amp;GroupID="& UserGroupParentID(i) &""">"& Application(CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& UserGroupParentID(i) &"']/@usertitle").text &"</a><br />"
					Else
						UserGroupList = UserGroupList & "<a href=""cookies.asp?action=ReGroup&amp;GroupID="&UserGroupParentID(i) & """>"& Application(CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& UserGroupParentID(i) &"']/@usertitle").text &"</a>"
					End If
				Next
				BoardReadme = Replace(BoardReadme,"{$UserGroupList}",UserGroupList)
			ElseIf Cint(IsUserPermissionOnly) > 0 Then
				BoardReadme = BoardReadme & IsBoard(4)
				UserGroupList = "<a href=""cookies.asp?action=ReGroup&amp;GroupID="& Usersession.documentElement.selectSingleNode("userinfo/@usergroupid2").text &""">"&Application(CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& usersession.documentElement.selectSingleNode("userinfo/@usergroupid2").text &"']/@usertitle").text&"</a><br />"
				BoardReadme = Replace(BoardReadme,"{$UserGroupList}",UserGroupList)
			End If
			NavStr = Replace(NavStr,"{$umsg}",BoardReadme)
		Else
			NavStr = Replace(NavStr,"{$umsg}","")
		End If
		NavStr = Replace(NavStr,"{$alertcolor}",mainsetting(1))
		NavStr = Replace(NavStr,"{$showstr}","")
		Response.Write vbNewLine & NavStr
	End Sub
	Public Sub AddErrCode(ErrCode)
		If ErrCodes = "" Then
			ErrCodes = ErrCode
		Else
			ErrCodes = ErrCodes & "," & ErrCode
		End If
	End Sub
	Public Property Let ErrType(ByVal Value)
		ShowErrType = Value
	End Property
	Public Sub Showerr()
		If ErrCodes<>"" Then
			If Stats="" Then 
			If BoardID=0 Then
				Stats="访问"& Forum_Info(0)
			Else
				Stats="访问"& BoardType
			End If
			End If
			Dim parameter
			If ShowErrType = 1 Then
				parameter="showerr.asp?BoardID="&Boardid&"&ErrCodes="&ErrCodes&"&action="&server.URLEncode(Stats)&"&ShowErrType=1"
				Set Dvbbs=Nothing
				Response.redirect parameter
			Else
				parameter="showerr.asp?BoardID="&Boardid&"&ErrCodes="&ErrCodes&"&action="&server.URLEncode(Stats)
				Set Dvbbs=Nothing
				Response.redirect parameter
			End If
		End If
	End Sub 
	Public Sub Showmessanger(title,messangertext)
		Response.Write "<div style=""position:absolute;top:220px;right:10px;width:350px;height : 90px;background:#fff;border: 5px solid #e4e8ef; "" id=""dv_msg"">"
		Response.Write "<div class=""th""><div>"&title&"</div></div>"
		Response.Write "<div class=""mainbar3"" style=""height : 36px;overflow :auto;"">"
		Response.Write messangertext
		Response.Write "</div>"
		Response.Write "<div class=""mainbar2"" style=""height : 20px;line-height : 20px;""><input type=""button"" value=""  关闭  "" onclick=""document.getElementById('dv_msg').style.visibility='hidden';"" />"
		Response.Write "</div></div>"
	End Sub
	Public Sub Footer()
		Dim Tmp,node
		If IsTopTable=1 Then
			If UserID>0  Then
				If IsObject( Application(Dvbbs.CacheName&"_messanger")) Then	
					Set Node=Application(Dvbbs.CacheName&"_messanger").documentElement.selectSingleNode("messanger[@touserid="&userid&"]")
					If Not Node is Nothing Then
						Showmessanger Node.selectSingleNode("@senduser").text,Node.text
						Application(Dvbbs.CacheName&"_messanger").documentElement.removeChild(Node)
					End If
				End If
			End If	
			Tmp = mainhtml(17)
			Tmp = Replace(Tmp,"{$boardid}",boardid)
			If (Dvbbs.Forum_ChanSetting(13)="1" And Dvbbs.Forum_ChanSetting(0)="1") Or Dvbbs.Forum_ChanSetting(3)="0" Then
				Tmp = Replace(Tmp,"{$UserTicket}","<br />" & lanstr(11))
			Else
				Tmp = Replace(Tmp,"{$UserTicket}","")
			End If
			Response.Write Tmp
			Tmp = mainhtml(8)
			If Forum_Setting(30) = "1" Then 
				Dim Endtime
				Endtime = Timer()
				Tmp = Replace(Tmp,"{$runtime}","页面执行时间 0"&FormatNumber((Endtime-Startime),5)&" 秒, "&SqlQueryNum&" 次数据查询<br />")
			Else
				Tmp = Replace(Tmp,"{$runtime}","")
			End If
			Dim Alibaba_Ad
			If IsSqlDataBase = 0 Or (IsBuss = 0 And IsSqlDataBase = 1) Or Forum_Info(0)="动网先锋论坛" Then
				Alibaba_Ad = "<div>网上贸易 创造奇迹! <a href = ""http://china.alibaba.com"" title = ""从网民、网友时代进入“网商”时代"" target=""_blank"">阿里巴巴</a> <a href = ""http://www.alibaba.com"" title= ""Online Marketplace of Manufacturers & Wholesalers"" target = ""_blank"">Alibaba</a></div>"
			End If
			Tmp = Replace(Tmp,"{$powered}","Powered By <a href = ""http://www.dvbbs.net/"" target = ""_blank"">Dvbbs</a>  <a href = ""http://www.dvbbs.net/download.asp"" target = ""_blank"">Version " & fVersion & "</a>" )
			If Dvbbs.Forum_ChanSetting(3)="0" Then
				Tmp = Replace(Tmp,"{$alipaymsg}","<a href=""https://www.alipay.com"" target=""_blank"" title=""本论坛采用阿里巴巴支付宝网上银行支付系统，安全、可靠、便捷""><img src="""&Dvbbs_Server_Url&"dvbbs/alipay_icon2.gif"" border=""0"" alt=""""></a>")
			Else
				Tmp = Replace(Tmp,"{$alipaymsg}","")
			End If
			Tmp = Replace(Tmp,"{$Footer_ads}",Forum_ads(1) & Alibaba_Ad)
			Tmp = Replace(Tmp,"{$copyright}",Forum_Copyright)
		Else
			Tmp ="</body></html>"
		End If
		Response.Write Tmp
		CacheData=Null
		Forum_Setting = Null
		Forum_UploadSetting = Null
		Forum_user =Null
		Forum_ChanSetting =Null
		BadWords = Null
		rBadWord = Null
		Forum_ads=Null
		Set Dvbbs=Nothing
	End Sub
	Public Sub Dvbbs_Suc(sucmsg)
		Dim TempStr
		TempStr = mainhtml(13)
		TempStr = Replace(TempStr,"{$sucmsg}",sucmsg)
		TempStr = Replace(TempStr,"{$returnurl}",Request.ServerVariables("HTTP_REFERER"))
		Response.Write TempStr
		TempStr = ""
	End Sub
	Public Function Execute(Command)
		If Not IsObject(Conn) Then ConnectionDatabase		
		If IsDeBug = 0 Then 
			On Error Resume Next
			Set Execute = Conn.Execute(Command)
			If Err Then
				err.Clear
				Set Conn = Nothing
				Response.Write "查询数据的时候发现错误，请检查您的查询代码是否正确。"
				Response.End
			End If
		Else
			If ShowSQL=1 Then
				Response.Write command & "<br>"
			End If
			Set Execute = Conn.Execute(Command)
		End If	
		SqlQueryNum = SqlQueryNum+1
	End Function
	'-----------------------------------------------------------------------------------------------------
	'独立道具查询
	Public Function Plus_Execute(Command)
		If Cint(Forum_Setting(92))=1 Then
			If Not IsObject(Plus_Conn) Then Plus_ConnectionDatabase
		Else
			If Not IsObject(Conn) Then ConnectionDatabase
		End If			
		If IsDeBug = 0 Then 
			On Error Resume Next
			If Cint(Forum_Setting(92))=1 Then
				Set Plus_Execute = Plus_Conn.Execute(Command)
			Else
				Set Plus_Execute = Conn.Execute(Command)
			End If
			If Err Then
				err.Clear
				If Cint(Forum_Setting(92))=1 Then
					Set Plus_Conn = Nothing
				Else
					Set Conn = Nothing
				End If
				Response.Write "查询数据的时候发现错误，请检查您的查询代码是否正确。"
				Response.End
			End If
		Else
			'Response.Write command & "<br>"
			If Cint(Forum_Setting(92))=1 Then
				Set Plus_Execute = Plus_Conn.Execute(Command)
			Else
				Set Plus_Execute = Conn.Execute(Command)
			End If
		End If	
		SqlQueryNum = SqlQueryNum+1
	End Function
	'IP/来源
	Public Function address(sip)
		Dim aConnStr,aConn,adb
		Dim str1,str2,str3,str4
		Dim  num
		Dim country,city
		Dim irs,SQL
		address="未知"
		If IsNumeric(Left(sip,2)) Then
			If sip="127.0.0.1" Then sip="192.168.0.1"
			str1=Left(sip,InStr(sip,".")-1)
			sip=mid(sip,instr(sip,".")+1)
			str2=Left(sip,instr(sip,".")-1)
			sip=Mid(sip,InStr(sip,".")+1)
			str3=Left(sip,instr(sip,".")-1)
			str4=Mid(sip,instr(sip,".")+1)
			If isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0 Then
			Else		
				num=CLng(str1)*16777216+CLng(str2)*65536+CLng(str3)*256+CLng(str4)-1
				adb = "data/ipaddress.mdb"
				aConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(adb)
				Set AConn = Server.CreateObject("ADODB.Connection")
				aConn.Open aConnStr
				country="亚洲"
				city=""
				sql="select top 1 country,city from dv_address where ip1 <="& num &" and ip2 >="& num 
				Set irs=aConn.execute(sql)
				If Not(irs.EOF And irs.bof) Then
					country=irs(0)
					city=irs(1)
				End If
				irs.close
				Set irs=Nothing
				aConn.Close
				Set aConn = Nothing 
				SqlQueryNum = SqlQueryNum+1
			End If
			address=country&city
		End If
	End Function
	'显示验证码
	Public Function GetCode()
			GetCode= Dvbbs.mainhtml(15)&"<img src=""DV_getcode.asp"" alt= ""验证码,看不清楚?请点击刷新验证码"" style=""cursor : pointer;height : 20px;"" onclick=""this.src='DV_getcode.asp'""/> "
	End Function
	'检查验证码是否正确
	Public Function CodeIsTrue()
		Dim CodeStr
		CodeStr=Lcase(Trim(Request("CodeStr")))
		If CStr(Session("GetCode"))=CStr(CodeStr) And CodeStr<>""  Then
			CodeIsTrue=True
			Session("GetCode")=empty
		Else
			CodeIsTrue=False
			Session("GetCode")=empty
		End If	
	End Function
	'用于用户发布的各种信息过滤，带脏话过滤
	Public Function HTMLEncode(fString)
		If Not IsNull(fString) Then
			fString = replace(fString, ">", "&gt;")
			fString = replace(fString, "<", "&lt;")
			fString = Replace(fString, CHR(32), " ")		'&nbsp;
			fString = Replace(fString, CHR(9), " ")			'&nbsp;
			fString = Replace(fString, CHR(34), "&quot;")
			'fString = Replace(fString, CHR(39), "&#39;")	'单引号过滤
			fString = Replace(fString, CHR(13), "")
			fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
			fString = Replace(fString, CHR(10), "<BR> ")
			fString=ChkBadWords(fString)
			HTMLEncode = fString
		End If
	End Function
	'用于论坛本身的过滤，不带脏话过滤
	Public Function iHTMLEncode(fString)
		If Not IsNull(fString) Then
			fString = replace(fString, ">", "&gt;")
			fString = replace(fString, "<", "&lt;")
			fString = Replace(fString, CHR(32), " ")
			fString = Replace(fString, CHR(9), " ")
			fString = Replace(fString, CHR(34), "&quot;")
			'fString = Replace(fString, CHR(39), "&#39;")
			fString = Replace(fString, CHR(13), "")
			fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
			fString = Replace(fString, CHR(10), "<BR> ")
			iHTMLEncode = fString
		End If
	End Function
	Public Function CheckNumeric(Byval CHECK_ID)
		If CHECK_ID<>"" and IsNumeric(CHECK_ID) Then _
			CHECK_ID = cCur(CHECK_ID) _
		Else _
			CHECK_ID = 0
		CheckNumeric = CHECK_ID
	End Function
	Public Function strLength(str)
		If isNull(str) Or Str = "" Then
			StrLength = 0
			Exit Function
		End If
		Dim WINNT_CHINESE
		WINNT_CHINESE=(len("例子")=2)
		If WINNT_CHINESE Then
			Dim l,t,c
			Dim i
			l=len(str)
			t=l
			For i=1 To l
				c=asc(mid(str,i,1))
				If c<0 Then c=c+65536
				If c>255 Then t=t+1
			Next
			strLength=t
		Else 
			strLength=len(str)
		End If
	End Function
	Public Function ChkBadWords(Str)
		If IsNull(Str) Then Exit Function
		Dim i
		For i = 0 To UBound(BadWords)
			If InStr(Str,BadWords(i))>0 Then
				If i > UBound(rBadWord) Then
					Str = Replace(Str,BadWords(i),"*")
				Else
					Str = Replace(Str,BadWords(i),rBadWord(i))
				End If
			End If
		Next
		ChkBadWords = Str
	End Function
	Public Function Checkstr(Str)
		If Isnull(Str) Then
			CheckStr = ""
			Exit Function 
		End If
		Str = Replace(Str,Chr(0),"")
		CheckStr = Replace(Str,"'","''")
	End Function
	Public Sub ReloadBoardCache(lboardid)
		Dim Boardidlist,bid
		Boardidlist=Split(lboardid,",")
		For Each bid in Boardidlist
			If Bid<> 0 and Bid<>"" Then
				LoadBoardData CLng(bid)
			End If
		Next
	End Sub
	Public Sub ReloadBoardInfo(lboardid)
		Dim Boardidlist,bid
		Boardidlist=Split(lboardid,",")
		For Each bid in Boardidlist
			If Bid<> 0 and Bid<>"" Then
				bid=CLng(Trim(bid))
				'是否更新缓存的判断，定义为30秒内不重复更新。
				If Not IsObject(Application(CacheName &"_information_" & bid)) Then
						LoadBoardinformation bid
						Application(CacheName &"_information_" & bid).documentElement.attributes.setNamedItem(Application(CacheName &"_information_" & bid).createNode(2,"lastupdate","")).text=Now()
				Else
					If Application(CacheName &"_information_" & bid).documentElement.selectSingleNode("@lastupdate") is Nothing Then
						Application(CacheName &"_information_" & bid).documentElement.attributes.setNamedItem(Application(CacheName &"_information_" & bid).createNode(2,"lastupdate","")).text=Now()
						LoadBoardinformation bid
						Application(CacheName &"_information_" & bid).documentElement.attributes.setNamedItem(Application(CacheName &"_information_" & bid).createNode(2,"lastupdate","")).text=Now()
					Else
						If DateDiff("s",Application(CacheName &"_information_" & bid).documentElement.selectSingleNode("@lastupdate").text,Now()) > 30 Then
						Application(CacheName &"_information_" & bid).documentElement.selectSingleNode("@lastupdate").text=Now()
						LoadBoardinformation bid
						Application(CacheName &"_information_" & bid).documentElement.attributes.setNamedItem(Application(CacheName &"_information_" & bid).createNode(2,"lastupdate","")).text=Now()
						End If
						End If
				End If
			End If
		Next
	End Sub
	'取得带端口的URL
	Property Get Get_ScriptNameUrl()
		If request.servervariables("SERVER_PORT")="80" Then
			Get_ScriptNameUrl="http://" & request.servervariables("server_name")&replace(lcase(request.servervariables("script_name")),ScriptName,"")
		Else
			Get_ScriptNameUrl="http://" & request.servervariables("server_name")&":"&request.servervariables("SERVER_PORT")&replace(lcase(request.servervariables("script_name")),ScriptName,"")
		End If
	End Property
	Public Function GetBrowser()
		Dim Agent,XSLTemplate,proc
		Set Agent=Application(CacheName&"_forum_lockip").cloneNode(True)
		Agent.documentElement.attributes.setNamedItem(Agent.createNode(2,"ip","")).text=UserTrueIP
		Agent.documentElement.attributes.setNamedItem(Agent.createNode(2,"actforip","")).text=actforip
		Agent.documentElement.appendChild(Agent.createTextNode(Request.ServerVariables("HTTP_USER_AGENT")))
		Set XSLTemplate=Application(CacheName & "_getbrowser")		
		Set proc = XSLTemplate.createProcessor()
		proc.input = Agent
		proc.transform()
		Set Agent=Nothing
		Set GetBrowser=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		GetBrowser.loadxml proc.output
	End Function
	'---------------------------------------------------
	'记录道具操作日志信息(发生数量，记录事件类型，备注内容，用户最后剩余金币和点券（金币|点券）)
	'Log_ID,ToolsID,CountNum,Log_Money,Log_Ticket,AddUserName,AddUserID,Log_IP,Log_Time,Log_Type,BoardID,Conect,HMoney
	'Log_Type类型(0=其它,1=使用,2=转让,3=充值,4=购买,5=奖励,6=vip交易)
	'HMoney最后剩余金币和点券（金币|点券）
	'boardid 记录版面参数，后台为-1
	'---------------------------------------------------
	Public Sub ToolsLog(Log_ToolsID,CountNum,Log_Money,Log_Ticket,Log_Type,Conect,HMoney)
		Dim Sql
		Conect = CheckStr(Conect)
		HMoney = CheckStr(HMoney)
		Sql = "Insert into [Dv_MoneyLog] (ToolsID,CountNum,Log_Money,Log_Ticket,AddUserName,AddUserID,Log_IP,Log_Type,BoardID,Conect,HMoney) values (" &_
			CheckNumeric(Log_ToolsID) &","&_
			CheckNumeric(CountNum) &","&_
			CheckNumeric(Log_Money) &","&_
			CheckNumeric(Log_Ticket) &",'"&_
			MemberName &"',"&_
			UserID &",'"&_
			UserTrueIP &"',"&_
			Log_Type &","&_
			BoardID &",'"&_
			Conect &"','"&_
			HMoney &"'"&_
			")"
		'Response.Write Sql
		Dvbbs.Execute(Sql)
	End Sub
	'是否真正的搜索引擎
	Public Function IsWebSearch()
		IsWebSearch = False
		Dim Botlist,i
		BotList = "Google,Isaac,SurveyBot,Baiduspider,yahoo,yisou,3721,ia_archiver,P.Arthur,FAST-WebCrawler,Java,Microsoft-ATL-Native,TurnitinBot,WebGather,Sleipnir"
		Botlist = Split(Botlist,",")
		For i = 0 To Ubound(Botlist)
			If InStr(Lcase(Request.ServerVariables("HTTP_USER_AGENT")),Lcase(Botlist(i))) > 0 Then
				IsWebSearch = True
				Exit For
			End If
		Next
	End Function
End Class
Class cls_Templates
	Public html,Strings,pic
	Public Property Let Value(ByVal vNewValue)
		Dim TemplateStr,tmpstr,Re
			Set re=new RegExp
			re.IgnoreCase =True
			re.Global=True
			re.Pattern=Chr(10)
			tmpstr=re.Replace(vNewValue, vbnewline)
			re.Pattern="{\$PicUrl}"
			TemplateStr=re.Replace(tmpstr, Dvbbs.Forum_PicUrl)
			Set Re=Nothing
		'TemplateStr = replace(vNewValue,Chr(10),vbnewline)
		'TemplateStr = Replace(TemplateStr,"{$PicUrl}",Dvbbs.Forum_PicUrl)
		tmpstr = Split(TemplateStr,"@@@")
		html = Split(tmpstr(0),"|||"):Strings = Split(tmpstr(1),"|||"):pic = Split(tmpstr(2),"|||")
	End Property
	Private Sub class_terminate()
		html=empty
		Strings=empty 
		pic=empty
	End Sub
End Class
Class cls_UserOnlne
	Public Forum_Online,Forum_UserOnline,Forum_GuestOnline
	Private l_Online,l_GuestOnline
	Private Sub Class_Initialize()
		Dvbbs.Name="Forum_Online"
		Dvbbs.Reloadtime=60
		If Dvbbs.ObjIsEmpty() Then ReflashOnlineNum
		Dvbbs.Name="Forum_Online"
		Forum_Online = Dvbbs.Value
		Dvbbs.Name="Forum_UserOnline"
		If Dvbbs.ObjIsEmpty() Then ReflashOnlineNum
		Forum_UserOnline=Dvbbs.Value
		If Forum_Online < 0  Or Forum_UserOnline < 0 Or Forum_UserOnline > Forum_Online Then ReflashOnlineNum
		Forum_GuestOnline = Forum_Online - Forum_UserOnline
		l_Online=-1:l_GuestOnline=-1
		Dvbbs.Reloadtime=28800
	End Sub
	Public Sub OnlineQuery()
		Dim SQL,SQL1
		Dim TempNum,TempNum1
		Dvbbs.Name="delOnline_time"
		If Dvbbs.ObjIsEmpty() Then Dvbbs.Value=Now()
		If DateDiff("s",Dvbbs.Value,Now()) > Clng(Dvbbs.Forum_Setting(8))*10 Then
			Dvbbs.Value=Now()
			If Not IsObject(Conn) Then ConnectionDatabase
			If IsSqlDataBase = 1 Then
				SQL = "Delete From [DV_Online] Where UserID=0 And Datediff(Mi, Lastimebk, " & SqlNowString & ") > " & Clng(Dvbbs.Forum_Setting(8))
				SQL1 = "Delete From [DV_Online] Where UserID>0 And Datediff(Mi, Lastimebk, " & SqlNowString & ") > " & Clng(Dvbbs.Forum_Setting(8))
			Else
				SQL = "Delete From [Dv_Online] Where UserID=0 And Datediff('s', Lastimebk, " & SqlNowString & ") > " & Dvbbs.Forum_Setting(8) & "*60" 
				SQL1 = "Delete From [Dv_Online] Where UserID>0 And Datediff('s', Lastimebk, " & SqlNowString & ") > " & Dvbbs.Forum_Setting(8) & "*60"
			End If
			Conn.Execute SQL,TempNum
			Conn.Execute SQL1,TempNum1
			Dvbbs.SqlQueryNum = Dvbbs.SqlQueryNum + 2
			'如果删除客人数大于0，则应该更新总数
			If TempNum>0 Then
				'更新缓存总在线数据
				Forum_Online = Forum_Online - TempNum
				Forum_GuestOnline = Forum_GuestOnline - TempNum
			End If
			'如果删除用户数大于0，则应该更新总数和用户数
			If TempNum1>0 Or  TempNum>0 Then
				'更新缓存总在线数据
				Forum_Online = Forum_Online - TempNum1
				Forum_UserOnline = Forum_UserOnline - TempNum1
				
			End If
			Dvbbs.Name="Forum_Online"
			Dvbbs.Value=Forum_Online
			'更新缓存总用户在线数据
			Dvbbs.Name="Forum_UserOnline"
			Dvbbs.Value=Forum_UserOnline
			Forum_Online = Forum_Online - TempNum1
		End If
	End Sub
	'刷新在线数据缓存
	Public Sub ReflashOnlineNum
		Dim Rs
		Dvbbs.Name="Forum_Online"
		Set Rs=Dvbbs.Execute("Select Count(*) From Dv_Online")
		Dvbbs.Value=Rs(0)
		Forum_Online = Dvbbs.Value
		Dvbbs.Name="Forum_UserOnline"
		Set Rs=Dvbbs.Execute("Select Count(*) From Dv_Online Where UserID>0")
		If Not IsNull(Rs(0)) Then
			Dvbbs.Value=Rs(0)
		Else
			Dvbbs.Value=0
		End If
		Forum_UserOnline = Dvbbs.Value
		Set Rs=Nothing
	End Sub
	'查询在某版面的在线总数
	Public Property Get Board_Online
		Board_Online=Board_UserOnline+Board_GuestOnline
	End Property
	Public Property Get Board_GuestOnline
		If l_GuestOnline=-1 Then
			Dim Rs
			Set Rs=Dvbbs.Execute("Select Count(*) From Dv_Online where BoardID="&Dvbbs.BoardID&" and UserID=0")
			l_GuestOnline=Rs(0):Rs.Close:Set Rs= Nothing 
		End If
		If IsNull(l_GuestOnline) Then l_GuestOnline=0
		Board_GuestOnline=l_GuestOnline
	End Property
	Public Property Get Board_UserOnline
		If l_Online=-1 Then
			Dim Rs
			Set Rs=Dvbbs.Execute("Select Count(*) From Dv_Online where BoardID="&Dvbbs.BoardID&" and UserID>0")
			l_Online=Rs(0):Rs.Close:Set Rs= Nothing 
		End If
		Board_UserOnline=l_Online
	End Property
End Class
%>