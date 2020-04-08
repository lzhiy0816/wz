<object runat="server" id="DvStream" progid="ADODB.Stream"></object>
<%
'=========================================================
' File: Dv_ClsMain.asp
' Version:8.2.0
' Date: 2007-3-10
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
Const fversion="8.2.0"
Dim IP_MAX
Const guestxml="<?xml version=""1.0"" encoding=""gb2312""?><xml><userinfo statuserid=""0"" userid=""0"" username=""客人"" userclass=""客人"" usergroupid=""7"" cometime="""" boardid=""0"" activetime="""" statusstr=""""/></xml>"
Class Cls_Forum
	Rem Const
	Public BoardID,SqlQueryNum,Forum_Info,Forum_Setting,Forum_user,Forum_Copyright,Forum_ads,Forum_ChanSetting,Forum_UploadSetting
	Public Forum_sn,Forum_Version,Stats,StyleName,ErrCodes,NowUseBBS,Cookiepath,ScriptFolder,BoardInfoData,UserSession
	Public MainSetting,sysmenu,UserToday,BoardJumpList,BoardList,CacheData,Maxonline
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
	Public IsUserPermissionOnly,IsUserPermissionAll,ShowSQL,actforip,DvRegExp,DvRegExp1
	Public GroupName,ScriptPath,Forum_apis
	Rem Const
	Function iCreateObject(str)
		'iis5创建对象方法Server.CreateObject(ObjectName);
		'iis6创建对象方法CreateObject(ObjectName);
		'默认为iis6，如果在iis5中使用，需要改为Server.CreateObject(str);
		Set iCreateObject=CreateObject(str)
	End Function

	Function CreateXmlDoc(str)
		Set CreateXmlDoc = iCreateObject(str)
		CreateXmlDoc.async=false
	End Function

	Public Function ReadTextFile(fileName)
		On Error Resume Next
			'Response.Write Server.MapPath(ScriptPath&fileName)
			DvStream.charset="gb2312"
			DvStream.Mode = 3
			DvStream.open()
			DvStream.LoadFromFile(Server.MapPath(ScriptPath&fileName))
			ReadTextFile=DvStream.ReadText
			DvStream.close()
		If Err Then
			err.Clear
			PageEnd()
			Response.Clear
			Response.Write ScriptPath&fileName & "文件不存在！请检查,或者恢复官方模板数据！"
			Response.End
		End If
	End Function

	Function writeToFile(fileName,Text)
		DvStream.charset="gb2312"
		DvStream.Mode = 3
		DvStream.open()
		DvStream.WriteText(Text)
		DvStream.SaveToFile Server.MapPath(ScriptPath&fileName),2
		DvStream.close()
	End Function

	Private Sub Class_Initialize()		
	End Sub

	public Sub PageInit()
		ScriptPath="./"
		Forum_sn="DvForum 8.2"'如果一个虚拟目录或站点开多个论坛，则每个要错开，不能定义同一个名称
		Forum_sn=Forum_sn & "_" & Request.servervariables("SERVER_NAME")
		CacheName="DvCache 8.2"'如果一个虚拟目录或站点开多个论坛，则每个要错开，不能定义同一个名称
		IsUserPermissionOnly = 0
		IsUserPermissionAll = 0
		ShowErrType = 0 '错误信息显示模式
		SqlQueryNum = 0
		Reloadtime=600
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
		If InStr(ScriptName,"showerr")>0 Or InStr(ScriptName,"login")>0 Or InStr(ScriptName,"admin_")>0 Or InStr(ScriptName,"ajax")>0 Then Page_Admin=True
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
	End Sub
	'isapi_write
	Public Function ArchiveHtml(Textstr)
		Str=Textstr
		If isUrlreWrite = 1 Then
			Dim Str,re,Matches,Match
			Set re=new RegExp
			re.IgnoreCase =True
			re.Global=True
			re.Pattern = "<a(.[^>]*)index\.asp\?boardid=(\d+)(&|&amp;)topicmode=(\d+)?(&|&amp;)list_type=([\d,]+)?(&|&amp;)page=(\d+)?"
			str = re.Replace(str,"<a$1index_$2_$4_$6_$8.html")
			re.Pattern = "<a(.[^>]*)index\.asp\?boardid=(\d+)(&|&amp;)action=(.[^&]*)?(&|&amp;)topicmode=(\d+)?(&|&amp;)list_type=([\d,]+)?(&|&amp;)page=(\d+)?"
			str = re.Replace(str,"<a$1index_$2_$4_$6_$8_$10.html")
			re.Pattern = "<a(.[^>]*)index\.asp\?boardid=(\d+)(&|&amp;)page=(\d+)?(&|&amp;)action=(.[^<>""\'\s]*)?"
			str = re.Replace(str,"<a$1index_$2_$4_$6.html")
			re.Pattern = "<a(.[^>]*)index\.asp\?boardid=(\d+)(&|&amp;)topicmode=(\d+)?"
			str = re.Replace(str,"<a$1index_$2_$4.html")
			re.Pattern = "<a(.[^>]*)index\.asp\?boardid=(\d+)(&|&amp;)page=(\d+)?"
			str = re.Replace(str,"<a$1index_$2_$4_.html")
			re.Pattern = "<a(.[^>]*)index\.asp\?boardid=(\d+)"
			str = re.Replace(str,"<a$1index_$2.html")
			re.Pattern = "<a(.[^>|_]*)index\.asp"
			str = re.Replace(str,"<a$1index.html")
			re.Pattern = "<a(.[^>]*)dispbbs\.asp\?boardid=(\d+)(&|&amp;)replyid=(\d+)?(&|&amp;)id=(\d+)?(&|&amp;)skin=(\d+)?(&|&amp;)page=(\d+)?(&|&amp;)star=(\d+)?"
			str = re.Replace(str,"<a$1dispbbs_$2_$4_$6_skin$8_$10_$12.html")
			re.Pattern = "<a(.[^>]*)dispbbs\.asp\?boardid=(\d+)(&|&amp;)replyid=(\d+)?(&|&amp;)id=(\d+)?(&|&amp;)skin=(\d+)?(&|&amp;)star=(\d+)?"
			str = re.Replace(str,"<a$1dispbbs_$2_$4_$6_skin$8_$10.html")
			re.Pattern = "<a(.[^>]*)dispbbs\.asp\?boardid=(\d+)(&|&amp;)replyid=(\d+)?(&|&amp;)id=(\d+)?(&|&amp;)skin=(\d+)?"
			str = re.Replace(str,"<a$1dispbbs_$2_$4_$6_skin$8.html")
			re.Pattern = "<a(.[^>]*)dispbbs\.asp\?boardid=(\d+)(&|&amp;)id=(\d+)?(&|&amp;)page=(\d+)?(&|&amp;)(star|move)=([\w\d]+)?"
			str = re.Replace(str,"<a$1dispbbs_$2_$4_$6_$9.html")
			re.Pattern = "<a(.[^>]*)dispbbs\.asp\?boardid=(\d+)(&|&amp;)id=(\d+)?(&|&amp;)page=(\d+)?"
			str = re.Replace(str,"<a$1dispbbs_$2_$4_$6.html")
			re.Pattern = "<a(.[^>]*)dispbbs\.asp\?boardid=(\d+)(&|&amp;)id=(\d+)?"
			str = re.Replace(str,"<a$1dispbbs_$2_$4.html")
			re.Pattern = "<a(.[^>]*)dv_rss\.asp\?s=(.[^<|>|""|\'|\s]*)"
			str = re.Replace(str,"<a$1dv_rss_$2.html")
			re.Pattern = "<a(.[^>]*)dv_rss\.asp"
			str = re.Replace(str,"<a$1dv_rss.html")
			Set Re=Nothing
		End If
		ArchiveHtml = Str
	End Function

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
	End Sub

	Public Sub PageEnd()
		If EnabledSession Then
			If Not UserSession Is Nothing  Then Session(CacheName & "UserID")= UserSession.xml
		End If
		Set UserSession=Nothing 
		If IsObject(Conn) Then Conn.Close : Set Conn = Nothing
		If IsObject(Plus_Conn) Then Plus_Conn.Close : Set Plus_Conn = Nothing
		Set DvRegExp= Nothing
		Set DvRegExp1= Nothing
		CacheData=Null
		Forum_Setting = Null
		Forum_UploadSetting = Null
		Forum_user =Null
		Forum_ChanSetting =Null
		BadWords = Null
		rBadWord = Null
		Forum_ads=Null
	End Sub

	Public Sub Sendmessanger(touserid,senduser,messangertext)
		Dim Node
		If Not IsObject( Application(Dvbbs.CacheName&"_messanger")) Then
			Set  Application(Dvbbs.CacheName&"_messanger")=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
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
		'Response.Write DateDiff("s",CDate(Application(CacheName & "_" & LocalCacheName &"_-time")),Now())&"秒"
		ObjIsEmpty=False
		If  IsDate(Application(CacheName & "_" & LocalCacheName &"_-time")) Then
			If DateDiff("s",CDate(Application(CacheName & "_" & LocalCacheName &"_-time")),Now()) > (60*Reloadtime) Then ObjIsEmpty=True
		Else
			ObjIsEmpty=True
		End If
		If ObjIsEmpty Then RemoveCache()
	End Function

	Public Sub RemoveCache()
		Application.Lock
		Application.Contents.Remove(CacheName & "_" & LocalCacheName)
		Application.Contents.Remove(CacheName & "_" & LocalCacheName &"_-time")
		Application.unLock
	End Sub
	'取得基本设置数据
	Public Sub loadSetup()
		Dim Rs,locklist,ip,ip1,XMLDom,Node,i
		Name="setup"
		Set Rs = Dvbbs.Execute("Select id, Forum_Setting, Forum_ads, Forum_Badwords, Forum_rBadword, Forum_Maxonline, Forum_MaxonlineDate, Forum_TopicNum, Forum_PostNum, Forum_TodayNum, Forum_UserNum, Forum_YesTerdayNum, Forum_MaxPostNum, Forum_MaxPostDate, Forum_lastUser, Forum_LastPost, Forum_BirthUser, Forum_Sid, Forum_Version, Forum_NowUseBBS, Forum_IsInstall, Forum_challengePassWord, Forum_Ad, Forum_ChanName, Forum_ChanSetting, Forum_LockIP, Forum_Cookiespath, Forum_Boards, Forum_alltopnum, Forum_pack, Forum_Cid, Forum_AvaSiteID, Forum_AvaSign, Forum_AdminFolder, Forum_BoardXML, Forum_Css, Forum_apis From [Dv_Setup]")
		Value = Rs.GetRows(1)
		CacheData=value
		Set XMLDom=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
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
			Set Application(CacheName & "_forum_lockip")=XMLDom
			Application.UnLock
		Set XMLDom=Nothing
		If Not isobject(Application(CacheName & "_getbrowser")) Then
			Dim stylesheet
			Set stylesheet=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			stylesheet.load Server.MapPath(MyDbPath &"inc\GetBrowser.xslt")
			Application.Lock
			Set Application(CacheName & "_getbrowser")=Dvbbs.iCreateObject("msxml2.XSLTemplate" & MsxmlVersion)
			Application(CacheName & "_getbrowser").stylesheet=stylesheet
			Application.unLock
		End If
		Application.Lock
		Set Application(CacheName & "_accesstopic")=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		Application(CacheName & "_accesstopic").Loadxml Replace(Replace(CacheData(34,0),Chr(10),""),Chr(13),"")
		Application.unLock
	End Sub

	Public Sub LoadBbsBoard()		
	End Sub

	Public Sub LoadBoardList()
		Dim TempXmlDoc,TempMasterDoc,ChildNode
		Dim Rs,boardmaster,master,node,Board_setting
		Set Rs=Execute("select boardid,boardtype,ParentID,depth,rootid,Child,indeximg,parentstr,cid as checkout,cid as hidden,cid as nopost,cid as checklock,cid as mode,cid as simplenesscount,readme,boardmaster From Dv_board Order by rootid,Orders")
		Set TempXmlDoc = RecordsetToxml(rs,"board","BoardList")
		Rs.Close
		Set TempMasterDoc = Dvbbs.CreateXmlDoc("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		TempMasterDoc.documentElement = TempMasterDoc.createElement("masterlist")
		Set Rs=Execute("select boardmaster,boardid,Board_setting From Dv_board Order by Orders")
		Do While Not Rs.EOF
			Set Node = TempMasterDoc.documentElement.appendChild(TempMasterDoc.createNode(1,"boardmaster",""))
			Node.setAttribute "boardid",Rs(1)
			boardmaster=split(Rs("boardmaster")&"","|")
			For Each Master In boardmaster
				Node.appendChild(TempMasterDoc.createNode(1,"master","")).text=Master
			Next
			Board_setting=Split(Rs("Board_setting"),",")
			TempXmlDoc.documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@checkout").text=Board_setting(2)
			TempXmlDoc.documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@hidden").text=Board_setting(1)
			TempXmlDoc.documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@nopost").text=Board_setting(43)
			TempXmlDoc.documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@checklock").text=Board_setting(0)
			TempXmlDoc.documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@mode").text=Board_setting(39)
			TempXmlDoc.documentElement.selectSingleNode("board[@boardid='"& Rs("Boardid")&"']/@simplenesscount").text=Board_setting(41)
		Rs.MoveNext
		Loop
		Rs.Close
		Set Rs= Nothing
		Application.Lock
		Set Application(CacheName&"_boardmaster") = TempMasterDoc
		Set Application(CacheName&"_boardlist") = TempXmlDoc
		Application(CacheName&"_boardmaster_xml") = TempMasterDoc.xml
		Application(CacheName&"_boardlist_xml") = TempXmlDoc.xml
		Application.unLock
	End Sub

	Public Sub LoadPlusMenu()
		Name = "ForumPlusMenu"
		Dim Rs,XMLDom,Node,plus_setting,stylesheet,XMLStyle,proc
		Set Rs=Execute("Select id,plus_type,plus_name,mainpage,plus_copyright,plus_setting,isshowmenu as width,isshowmenu as height From Dv_Plus Where  Isuse=1 Order By ID")
		Set XMLDom=RecordsetToxml(rs,"plus","")
		Rs.close()
		Set Rs=Nothing
		For Each Node In XMLDom.documentElement.selectNodes("plus")
			plus_setting=Split(Split(node.selectSingleNode("@plus_setting").text,"|||")(0),"|")
			node.selectSingleNode("@plus_setting").text=plus_setting(0)
			node.selectSingleNode("@width").text=plus_setting(1)
			node.selectSingleNode("@height").text=plus_setting(2)
		Next
		Set XMLStyle=Dvbbs.iCreateObject("msxml2.XSLTemplate" & MsxmlVersion)

		Set stylesheet=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		
		stylesheet.load Server.MapPath(MyDbPath &"inc\Templates\plusmenu.xslt")
		XMLStyle.stylesheet=stylesheet
		Set proc=XMLStyle.createProcessor()
		proc.input = XMLDom
  		proc.transform()
  		value=proc.output
	End Sub

	Public Sub LoadBoardData(bid)
		Dim Rs
		Application.Lock
		Set Rs=Execute("select boardid,boarduser,board_ads,board_user,isgroupsetting,rootid,board_setting,sid,cid,Rules From Dv_board Where Boardid="&bid)
		Set Application(CacheName &"_boarddata_" & bid)=RecordsetToxml(rs,"boarddata","")
		Rs.Close
		Set Rs= Nothing
		Application.unLock
	End Sub

	Public Sub LoadBoardinformation(bid)'加载动态板面信息数据
		Dim Rs,lastpost,i
		Application.Lock
		Set Rs=Execute("select boardid,boardtopstr,postnum,topicnum,todaynum,lastpost as lastpost_0 From Dv_board Where Boardid="&bid)
		Set Application(CacheName &"_information_" & bid)=RecordsetToxml(rs,"information","")
		lastpost=Split(Application(CacheName &"_information_" & bid).documentElement.selectSingleNode("information/@lastpost_0").text,"$")
		For i=0 to UBound(lastpost)
			Application(CacheName &"_information_" & bid).documentElement.firstChild.setAttribute "lastpost_"& i,lastpost(i)
			If i = 7 Then Exit For
		Next
		Rs.Close
		Set Rs= Nothing
		Application.unLock
	End Sub

	Public Sub LoadAllBoardinformation()'加载所有板面信息数据
		Dim Rs,lastpost,i
		Dim TempXmlDom,Node,TempNode,TempXmlDom1
		Set Rs=Execute("select boardid,boardtopstr,postnum,topicnum,todaynum,lastpost as lastpost_0 From Dv_board Order by Orders")
		Set TempXmlDom = RecordsetToxml(rs,"information","")
		Rs.Close
		Set Rs = Nothing
		For Each Node In TempXmlDom.documentElement.selectNodes("information")
			lastpost=Split(Node.getAttribute("lastpost_0"),"$")
			For i=0 to UBound(lastpost)
				Node.setAttribute "lastpost_"& i,lastpost(i)
				If i = 7 Then Exit For
			Next
			Set TempXmlDom1=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Set TempNode = TempXmlDom1.appendChild(TempXmlDom1.createNode(1,"xml",""))
			TempNode.appendChild(Node)
			Application.Lock
			Set Application(CacheName &"_information_" & Node.getAttribute("boardid")) = TempXmlDom1
			Application.UnLock
		Next
		If IsObject(TempXmlDom1) Then Set TempXmlDom1 = Nothing
	End Sub

	Public Sub LoadGroupSetting()
		Dim Rs
		Set Rs=Dvbbs.Execute("Select GroupSetting,UserGroupID,ParentGID,IsSetting,UserTitle From Dv_UserGroups")
		Set Application(CacheName &"_groupsetting")=RecordsetToxml(rs,"usergroup","")
		Set Rs=Dvbbs.Execute("Select UserGroupID,usertitle,titlepic,orders From Dv_UserGroups order by orders")
		Set Application(CacheName &"_grouppic")=RecordsetToxml(rs,"usergroup","grouppic")
		Rs.close()
		Set Rs=Nothing
	End Sub

	Public Sub Loadstyle()
		Dim Rs
		Application.Lock
		Set Rs=Dvbbs.Execute("Select *  From Dv_Templates")
		Set Application(CacheName &"_style")=RecordsetToxml(rs,"style","")
		Rs.close()
		Set Rs=Nothing
		Application.UnLock
		LoadStyleMenu()
	End Sub

	Public Sub LoadStyleMenu()'生成风格选单数据
		Name="style_list"
		Dim HTMLstr
		HTMLStr="<a href=""cookies.asp?action=stylemod&amp;boardid=$boardid"" >恢复默认设置</a>"
		Dim Node
		For Each Node in Application(CacheName &"_style").documentElement.selectNodes("style")
			HTMLstr=(HTMLstr&"<br /><a href=""cookies.asp?action=stylemod&amp;skinid="& node.selectSingleNode("@id").text & "&amp;boardid=$boardid"">"& node.selectSingleNode("@type").text& "</a>")
		Next
  	value=HTMLstr
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
			Rs.close()
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
				Execute("Update Dv_Setup Set Forum_MaxPostNum="&Forum_TodayNum&",Forum_MaxPostDate="&SqlNowString)
			End If
			If act=1 Then loadSetup()
			Dim xmlhttp
			If IsSqlDataBase =0 Then
				On Error Resume Next
				Set xmlhttp = Dvbbs.CreateXmlDoc("msxml2.ServerXMLHTTP")
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
		Forum_apis=CacheData(36,0)
		Dim Setting,OpenTime,ischeck:Setting= Split(CacheData(1,0),"|||"):Forum_Info =  Split (Setting(0),",")
		Forum_Setting = Split (Setting(1),","):Forum_UploadSetting = Split(Forum_Setting(7),"|")
		Forum_user = Setting(2):Forum_user = Split (Forum_user,","):Forum_Copyright = Setting(3)
		Forum_ChanSetting = Split(CacheData(24,0),",")
		Forum_Version = fversion 'CacheData(18,0)
		BadWords = Split(CacheData(3,0),"|")
		Set DvRegExp=new RegExp
		DvRegExp.IgnoreCase =True
		DvRegExp.Global=true
		Set DvRegExp1=new RegExp
		DvRegExp1.IgnoreCase =True
		DvRegExp1.Global=false
		if(CacheData(3,0)<>"") Then
			DvRegExp.Pattern="(" & CacheData(3,0) &")"
			DvRegExp1.Pattern="(" & CacheData(3,0) &")"
		End if
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
		Rem Hantg 2007-12-05
		If UBound(Forum_Setting)<107 Then
			Redim Preserve Forum_Setting(106)
		End If
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
			If Board_Setting(50)<>"0" And Board_Setting(50)<>"" Then Set Dvbbs=Nothing:Response.Redirect Board_Setting(50)
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
			If OpenTime(Hour(Now))="1" Then Response.redirect "showerr.asp?action=stop&boardid="&Dvbbs.BoardID&""
		End If
		'在线人数限制
		If ischeck > 0 And Not Page_Admin Then
			If MyBoardOnline.Forum_Online > ischeck And BoardID=0 Then
				If Not IsONline(Membername,1) Then Set Dvbbs=Nothing:Response.Redirect "showerr.asp?action=limitedonline&lnum="&ischeck
			End If
			If BoardID > 0 Then
				If (Not IsONline(Membername,1)) And MyBoardOnline.Board_Online > ischeck Then Set Dvbbs=Nothing:Response.Redirect "showerr.asp?action=limitedonline&lnum="&ischeck
			End If
		End If
		Dim CookiesSid
		CookiesSid = Request.Cookies("skin")("skinid_"&BoardID)
		
		If  CookiesSid = "" Then
			If BoardID = 0 Then 
				SkinID = Main_Sid
			Else
				SkinID = Sid
			End If
		Else
			SkinID = CookiesSid
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
		Rs.close()
		Set rs=Nothing  
	End Function  

	Public Sub LoadTemplates(Page_Fields)
		Dim Style_Pic,Main_Style,TempStyle,cssfilepath
		Forum_PicUrl="skins/Default/"
		Template.TPLname=Page_Fields
		Dim Node
		Set Node=Application(CacheName &"_style").documentElement.selectSingleNode("style[@id='"& Skinid &"']")
		If(Node is Nothing) Then
			Set Node=Application(CacheName &"_style").documentElement.selectSingleNode("style")
			If (Node is Nothing) Then
				Response.Write "没有注册可用风格模板，请到后台设置"
				Response.End
			Else
				Skinid=node.selectSingleNode("@id").text
			End if
		End If
		Template.ChildFolder=CheckStr(node.getAttribute("folder"))
		If Template.ChildFolder="" Then Template.ChildFolder="template_1"

		if (lcase(Page_Fields)="index" or lcase(Page_Fields)="dispbbs" or lcase(Page_Fields)="showerr") Then 
			Template.Cache=True '
		End if
		If Not (Instr(ScriptName,"index")>0 Or Page_Admin) Then
			Dim FolderPath
			If MyDbPath = "../../../" Then
				FolderPath = ""
			Else
				FolderPath = MyDbPath
			End If
			Style_Pic = Dvbbs.ReadTextFile(FolderPath & Template.Folder &"\"& Template.ChildFolder &"\pub_FaceEmot.htm")
			'Style_Pic = Dvbbs.ReadTextFile(Template.Folder &"\"& Template.ChildFolder &"\pub_FaceEmot.htm")
			Style_Pic = Split(Style_Pic,"@@@")
			Forum_UserFace = Style_Pic(0)
			Forum_PostFace = Style_Pic(1)
			Forum_Emot = Style_Pic(2)
		End If
		MainSetting = Split(MainHtml(0),"||")
		if ubound(MainSetting) >12 Then Forum_PicUrl=MainSetting(13)
	End Sub

	Public Function MainPic(Index)
		Dim Tplname,Cache
		Tplname=Template.TPLname
		Cache=Template.Cache
		Template.TPLname="pub"
		Template.Cache=true
		MainPic=Template.Pic(Index)
		Template.TPLname=Tplname
		Template.Cache=Cache
	End Function

	Public Function LanStr(Index)
		Dim Tplname,Cache
		Tplname=Template.TPLname
		Cache=Template.Cache
		Template.TPLname="pub"
		Template.Cache=true
		LanStr=Template.Strings(Index)
		Template.TPLname=Tplname
		Template.Cache=Cache
	End Function

	Public Function MainHtml(Index)
		Dim Tplname,Cache
		Tplname=Template.TPLname
		Cache=Template.Cache
		Template.TPLname="pub"
		Template.Cache=True
		MainHtml=Template.HTML(Index)
		Template.TPLname=Tplname
		Template.Cache=Cache
	End Function

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
		Set UserSession=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		UserSession.Loadxml guestxml
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
			Set UserSession=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
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
					
					If NeedChecklongin Or toupdate Then TrueCheckUserLogin
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
		If UserID=0 Then
			UserToday = Split("0|0|0|0|0","|")
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
		End If	
		Call GetGroupSetting()
	End Sub
	Rem xmlroot跟节点名称 row记录行节点名称
	Public Function RecordsetToxml(Recordset,row,xmlroot)
		Dim i,node,rs,j,DataArray
		If xmlroot="" Then xmlroot="xml"
		If row="" Then row="row"
		Set RecordsetToxml=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
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
		Set ArrayToxml=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
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
	'Response.Write "<iframe style=""border:0px;width:0px;height:0px;""  src=""newpass.asp"" name=""Dvnewpass""></iframe>"
		Dim TruePassWord,usercookies,i
		usercookies=Request.Cookies(Dvbbs.Forum_sn)("usercookies")
		TruePassWord=Dvbbs.Createpass
		If (Isnull(usercookies) or usercookies="") And Not Isnumeric(usercookies) Then usercookies=0

		Call updateCookiesInfo(usercookies,TruePassWord)
		'检查写入是否成功如果成功则更新数据
		i=0
		Do While i<3
			If Dvbbs.checkStr(Trim(Request.Cookies(Dvbbs.Forum_sn)("password")))=TruePassWord Then
				Dvbbs.Execute("Update [Dv_user] Set TruePassWord='"&TruePassWord&"' where UserID="&Dvbbs.UserID)
				Dvbbs.MemberWord = TruePassWord
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@truepassword").text= TruePassWord
				Exit Do
			Else
				Call updateCookiesInfo(usercookies,TruePassWord)
			End If
			i=i+1
		Loop
	End Sub

	'更新用户Cookies信息
	Sub updateCookiesInfo(usercookies,TruePassWord)
		Select Case Cint(usercookies)
			Case 0
				Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
			Case 1
				Response.Cookies(Dvbbs.Forum_sn).Expires=Date+1
				Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
			Case 2
				Response.Cookies(Dvbbs.Forum_sn).Expires=Date+31
				Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
			Case 3
				Response.Cookies(Dvbbs.Forum_sn).Expires=Date+365
				Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
		End Select
		Response.Cookies(Dvbbs.Forum_sn).path=Dvbbs.cookiepath
		Response.Cookies(Dvbbs.Forum_sn)("username") = Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@username").text
		Response.Cookies(Dvbbs.Forum_sn)("UserID") = Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userid").text
		Response.Cookies(Dvbbs.Forum_sn)("userclass") = Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userclass").text 
		Response.Cookies(Dvbbs.Forum_sn)("userhidden") = Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userhidden").text
		Response.Cookies(Dvbbs.Forum_sn)("password") = TruePassWord
		Response.Flush
	End Sub
	
	Public Sub TrueCheckUserLogin()	
		Dim Rs,SQL,FoundMyGroupID
		FoundMyGroupID = 0

		Sql="Select UserID,UserName,UserPassword,UserEmail,UserPost,UserTopic,UserSex,UserFace,UserWidth,UserHeight,JoinDate,LastLogin as cometime ,LastLogin,LastLogin as activetime,UserLogins,Lockuser,Userclass,UserGroupID,UserGroup,userWealth,userEP,userCP,UserPower,UserBirthday,UserLastIP,UserDel,UserIsBest,UserHidden,UserMsg,IsChallenge,UserMobile,TitlePic,UserTitle,TruePassWord,UserToday,UserMoney,UserTicket,FollowMsgID,Vip_StarTime,Vip_EndTime,userid as boardid,Usersetting"
		Sql=Sql & " From [Dv_User] Where UserID = " & UserID
		Set Rs = Execute(Sql)

		If Rs.EOF Then
			UserID = 0:LetGuestSession():Exit Sub
		Else
			If Not (LCase(Rs("UserName"))=LCase(Membername) and Rs("TruePassWord")=Memberword) Then
				If EnabledSession Then
					Set UserSession=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
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
		Set UserSession = RecordsetToxml(rs,"userinfo","xml")
		UserSession.documentElement.selectSingleNode("userinfo/@cometime").text=Now()
		UserSession.documentElement.selectSingleNode("userinfo/@activetime").text=DateAdd("s",-3600,Now())
		UserSession.documentElement.selectSingleNode("userinfo/@boardid").text=boardid
		UserSession.documentElement.selectSingleNode("userinfo").attributes.setNamedItem(UserSession.createNode(2,"isuserpermissionall","")).text=FoundUserPermission_All()
		Dim BS
		Set Bs=GetBrowser()
		UserSession.documentElement.appendChild(Bs.documentElement)
		If EnabledSession Then
			Session(CacheName & "UserID")= UserSession.xml
		End If
		Rs.close()
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
						Rs.close()
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
		If UserID > 0 Then IsUserPermissionAll = CLng(UserSession.documentElement.selectSingleNode("userinfo/@isuserpermissionall").text)
		If BoardID > 0 And Not ScriptName="showerr.asp" Then CheckBoardInfo()
		If UserID > 0 And BoardID=0 Then
			If IsUserPermissionAll="1" Then LoadUserPermission_All()
		End If
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
		Rs.close()
		Set Rs=Nothing
	End Sub

	Public Sub ActiveOnline()
		Response.Write "<script language=""JavaScript"">"
		Response.Write "setTimeout('ActiveOnline("&boardid&")',2000);"
		Response.Write "</script>"
	End Sub

	Public Sub ActiveOnline1()
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
				Execute("update [Dv_user] set lastlogin=" & SqlNowString & " where userid="&Dvbbs.userid)
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
				If Cls_IsSearch Then Exit Sub  
				SQL = "Insert Into [Dv_Online](ID,Username,Userclass,Ip,Startime,Lastimebk,Boardid,Browser,Stats,Usergroupid,Actcome,Userhidden,actforip) Values (" & StatUserID & ",'客人','客人','" & UserTrueIP & "'," & SqlNowString & "," & SqlNowString & "," & Boardid & ",'" & platform&"|"&Browser&version & "','" & StatsStr & "',7,'" & Actcome & "'," & Userhidden & ",'"& checkstr(actforip)&"')"
				'更新缓存总在线数据
				MyBoardOnline.Forum_Online=MyBoardOnline.Forum_Online+1
				Name="Forum_Online"
				value=MyBoardOnline.Forum_Online 
			Else
				Dim TempUserInfo,TempHidden
				Set TempUserInfo=UserSession.documentElement.selectSingleNode("userinfo")
				If TempUserInfo Is Nothing Then
					TempHidden=2
				Else
					TempHidden=Dvbbs.CheckNumeric(TempUserInfo.getAttribute("userhidden"))
				End If
				SQL = "Update [Dv_Online] Set Lastimebk = " & SqlNowString & ",Boardid = " & Boardid & ",Stats = '" & StatsStr & "',Userhidden="&TempHidden&" Where ID = " & Ccur(StatUserID)
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
		Response.Write MainHtml(18)
		Nowstats=stats
		If ScriptName="index.asp" Then
			If BoardType<>"" Then Stats=BoardType & (Left(Replacehtml(boardreadme),20)&"....")
		ElseIf ScriptName <> "dispbbs.asp" Then
			If BoardType<>"" Then Stats=BoardType&("-"&Stats)
		End If
		Stats=Replace(Stats,Chr(13),"")
		stats=Replacehtml(stats)
		'搜索引擎优化部分
		If Request("IsSearch_a") <> "" Then stats = stats & "-网站地图"
		Nowstats=Replacehtml(Nowstats)
		If IsSearch Then
			Response.Write Replace(Replace(Replace(MainHtml(1),"{$keyword}",Replace(Forum_info(8),"|",",")),"{$description}",Forum_info(10)),"{$title}",stats &"["& Forum_Info(0) &"] -- Powered By Dvbbs.net," & Now())
		Else
			Response.Write Replace(Replace(Replace(MainHtml(1),"{$keyword}",Replace(Forum_info(8),"|",",")),"{$description}",Forum_info(10)),"{$title}",(stats &"["& Forum_Info(0) &"]"))
		End If
		'搜索引擎优化结束
		If Boardid=0 Then
			Response.Write "<link title="""& Forum_Info(0) &"-频道列表"" type=""application/rss+xml"" rel=""alternate"" href=""rssfeed.asp"" />"
			Response.Write Chr(10)
			Response.Write "<link title="""& Forum_Info(0) &"-最新20篇论坛主题"" type=""application/rss+xml"" rel=""alternate"" href=""rssfeed.asp?rssid=4"" />"
		Else
			Response.Write "<link title="""& Replacehtml(BoardType) &"-最新20篇论坛主题"" type=""application/rss+xml"" rel=""alternate"" href=""rssfeed.asp?boardid="&boardid&"&amp;rssid=4"" />"
		End If
		Response.Write Chr(10)
		Response.Write MainHtml(2)
		Dim node,XMLDOM
		If Not Cls_IsSearch Then
			Response.Write "<script language=""javascript"" type=""text/javascript"">"
			Response.Write "var boardxml='',ISAPI_ReWrite = "&isUrlreWrite&",forum_picurl='"&Forum_PicUrl&"';"
			Response.Write "</script>"
		Else
			Exit Sub
		End If
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
			sysmenu = MainHtml(20)
			Dim node,XMLDom,iTempStr
			Set XMLDOM=Application(Dvbbs.CacheName&"_boardlist")
			iTempStr = "&nbsp;&nbsp;"
			For each node in XMLDOM.documentElement.selectNodes("board[@parentid=0]")
				iTempStr = iTempStr & ("<a href=""index.asp?IsSearch_a=2&BoardID="&Node.selectSingleNode("@boardid").text&""">" & Node.selectSingleNode("@boardtype").text & "</a>&nbsp;&nbsp;")
			Next
			sysmenu = Replace(sysmenu,"{$catlist}",(iTempStr & "<a href=""http://www.cndw.com"" title=""论坛,bbs,免费论坛,国内最大的论坛软件服务提供商,blog,boke,博客,防火墙,插件"">动网论坛</a>&nbsp;&nbsp;<a href=""http://bbs.cndw.com"" title=""论坛,bbs,免费论坛,国内最大的论坛软件服务提供商官方讨论区,blog,boke,博客,asp,asp.net,电脑,软件,灌水,防火墙,开发,插件"">官方论坛</a>&nbsp;&nbsp;<a href=http://tool.cndw.com>站长工具</a>"))
			'搜索引擎优化结束
		ElseIf UserID = 0 Then 
			sysmenu = MainHtml(7)
		Else
			sysmenu = Replace(MainHtml(6),"{$username}",Membername)
			If UserHidden=2 Then
				sysmenu = Replace(sysmenu,"{$hiddeninfo}",LanStr(3))
			Else
				sysmenu = Replace(sysmenu,"{$hiddeninfo}",LanStr(4))
			End If
			If Master Or GroupSetting(70)="1" Then
				sysmenu = Replace(sysmenu,"{$manageinfo}",MainHtml(10))
			ElseIf Superboardmaster Then
				sysmenu = Replace(sysmenu,"{$manageinfo}",MainHtml(19))
			Else
				sysmenu = Replace(sysmenu,"{$manageinfo}","")
			End If
			sysmenu = Replace(sysmenu,"{$raymenuinfo}","")
	
			sysmenu = Replace(sysmenu,"{$userid}",UserID)

		End If
		Tempstr1 = Replace(MainHtml(17),"{$boardid}",boardid)
		If (Dvbbs.Forum_ChanSetting(13)="1" And Dvbbs.Forum_ChanSetting(0)="1") Or Dvbbs.Forum_ChanSetting(3)="0" Then
			Tempstr1 = Replace(Tempstr1,"{$UserTicket}",("<br />" & LanStr(11)))
		Else
			Tempstr1 = Replace(Tempstr1,"{$UserTicket}","")
		End If
		Tempstr1=Split(Tempstr1&"","||")
		If UBound(Tempstr1)=2 Then
			sysmenu = Replace(sysmenu&"","{$menu_membertools}",Tempstr1(0))
			sysmenu = Replace(sysmenu&"","{$menu_show}",Tempstr1(2))
		Else
			sysmenu = Replace(sysmenu&"","{$menu_membertools}","")
			sysmenu = Replace(sysmenu&"","{$menu_show}","")
		End If
		If Dvbbs.Forum_setting(99) = "1" Then
			sysmenu = Replace(sysmenu,"{$isboke}",MainHtml(21))
		Else
			sysmenu = Replace(sysmenu,"{$isboke}","")
		End If
		If Forum_Setting(90)=0 Then 
			sysmenu = Replace(sysmenu,"{$Plus_Tools}","")
		Else
			sysmenu = Replace(sysmenu,"{$Plus_Tools}",MainHtml(16))
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
			Dim MainHtml_9
			MainHtml_9=MainHtml(9)
			MainHtml_9 = Split(MainHtml_9,"||")
			Tempstr1=Replace(Replace(MainHtml_9(0),"{$dskinid}",CacheData(17,0)),"{$csslist}","")
		End If
		sysmenu = Replace(sysmenu,"{$syles}",Tempstr1)
		TempStr = TempStr & (Chr(10) & MainHtml(4))
		TempStr = Replace(TempStr,"{$width}",MainSetting(0))
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
			TempStr = Replace(TempStr,"{$mailto}",("mailto:" & Forum_Info(5)))
		End If
		TempStr = Replace(TempStr,"{$title}",(Forum_Info(0) & "-" & Replace(stats,"'","\'")))
		TempStr = Replace(TempStr,"{$top_ads}",Forum_ads(0))
		TempStr = Replace(TempStr,"{$menu}",Chr(10) & sysmenu)
		TempStr = Replace(TempStr,"{$boardid}",boardid)
		TempStr = Replace(TempStr,"{$alertcolor}",MainSetting(1))
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
		If Request.Cookies("geturl")<>Request.ServerVariables("PATH_INFO")&"?"&Request.QueryString Then
			Response.Cookies("geturl") = Request.ServerVariables("PATH_INFO")&"?"&Request.QueryString
		End If
	End Sub 

	Public Sub Head_var(IsBoard,idepth,GetTitle,GetUrl)
		Dim NavStr,AllBoardList
		If Dvbbs.BoardID=0 Then BoardReadme=LanStr(2) & " <b>" & Forum_Info(0) & "</b>"
		If BoardID>0 Then
			NavStr = " <a href="""&(Forum_Info(11)&""" onmouseover=""showmenu(event,BoardJumpList(0),'',0);"" style=""cursor:hand;""><b>"&Forum_info(0)&"</b></a> → ")
			'NavStr = "<li class=""mainmenu_li"" onmouseover=""appMenu(this,0,0);""> <a href="""&(Forum_Info(11)&"""><b>"&Forum_info(0)&"</b></a><div class=""submenu"" onmouseout=""hidemenu(1);hidemenu(0)""></div></li> → ")			
		Else
			NavStr = " <a href="""&(Forum_Info(11)&""">"&Forum_info(0)&"</a> → ")
		End If
		If IsBoard=1 Then
			If BoardParentID=0 Then
				NavStr = NavStr & (" <a href=""index.asp?boardid="&BoardID&""" onMouseOver=""showmenu(event,BoardJumpList("&Dvbbs.Boardid&"),'',0);"">"&BoardType&"</a>")
				'NavStr = NavStr & ("<li class=""mainmenu_li"" onmouseover=""appMenu(this,0,"&BoardID&");""><a href=""index.asp?boardid="&BoardID&""">"&BoardType&"</a> <div class=""submenu"" onmouseout=""hidemenu(1);hidemenu(0)""></div></li>")				
			Else
				If ScriptName="dispbbs.asp" Then 
					NavStr = NavStr & (BoardInfoData & " → <a href=""index.asp?boardid="&BoardID&"&page="&Request("page")&""">"&BoardType&"</a>")
				Else
					NavStr = NavStr & (BoardInfoData & " → <a href=""index.asp?boardid="&BoardID&""">"&BoardType&"</a>")
				End If
			End If
			NavStr = NavStr & (" → " & Nowstats)
		Elseif IsBoard=2 Then
			NavStr = NavStr & Nowstats
		Else
			NavStr = NavStr & ("<a href="&GetUrl&">"&GetTitle&"</a> → " & Nowstats)
		End If
		NavStr = Replace(MainHtml(5),"{$nav}",NavStr)
		NavStr = Replace(NavStr,"{$width}",MainSetting(0))
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
			IsBoard = Split(MainHtml(12),"||")
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

			NavStr = Replace(NavStr,"{$umsg}",BoardReadme)
		Else
			NavStr = Replace(NavStr,"{$umsg}","")
		End If
		NavStr = Replace(NavStr,"{$alertcolor}",MainSetting(1))
		NavStr = Replace(NavStr,"{$showstr}","")
		'NavStr =NavStr&"</li>"
		Response.Write ArchiveHtml((vbNewLine & NavStr))
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

	Public Property Get ErrType()
		ErrType = ShowErrType
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
				parameter = MyDbPath& "showerr.asp?BoardID="&Boardid&"&ErrCodes="&ErrCodes&"&action="&server.URLEncode(Stats)&"&ShowErrType=1"
				PageEnd()
				Response.redirect parameter
			Else
				parameter = MyDbPath& "showerr.asp?BoardID="&Boardid&"&ErrCodes="&ErrCodes&"&action="&server.URLEncode(Stats)
				PageEnd()
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
			Tmp = MainHtml(8)
			If Forum_Setting(30) = "1" Then 
				Dim Endtime
				Endtime = Timer()
				Tmp = Replace(Tmp,"{$runtime}","页面执行时间 0"&FormatNumber((Endtime-Startime),5)&" 秒, "&SqlQueryNum&" 次数据查询<br />")
			Else
				Tmp = Replace(Tmp,"{$runtime}","")
			End If
			Dim Alibaba_Ad
			If IsSqlDataBase = 0 Or (IsBuss = 0 And IsSqlDataBase = 1) Or Forum_Info(0)="动网先锋论坛" Then
				Alibaba_Ad = "<div></div>"
			End If
			Tmp = Replace(Tmp,"{$powered}","Powered By <a href = ""http://www.dvbbs.net/"" target = ""_blank"">Dvbbs</a>  <a href = ""http://www.dvbbs.net/download.asp"" target = ""_blank"">Version " & fVersion & "</a>" )
			If Dvbbs.Forum_ChanSetting(3)="0" Then
				Tmp = Replace(Tmp,"{$alipaymsg}","")
			Else
				Tmp = Replace(Tmp,"{$alipaymsg}","")
			End If
			Tmp = Replace(Tmp,"{$Footer_ads}",Forum_ads(1) & Alibaba_Ad)
			Tmp = Replace(Tmp,"{$copyright}",Forum_Copyright)
			Tmp = Replace(Tmp,"{$boardid}",BoardID)
		Else
			Tmp ="</body></html>"
		End If
		Response.Write Tmp
		
	End Sub

	Public Sub Dvbbs_Suc(sucmsg)
		Dim TempStr
		TempStr = MainHtml(13)
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
				Set AConn = Dvbbs.iCreateObject("ADODB.Connection")
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
			'GetCode= Dvbbs.MainHtml(15)&("<img id=""imgid"" src="""&DvCodeFile&""" alt=""验证码,看不清楚?请点击刷新验证码"" style=""cursor:pointer; vertical-align:middle;height:18px;"" onclick=""this.src='"&DvCodeFile&"?t='+Math.random()""/> ")
			GetCode= Dvbbs.MainHtml(15)'o 08.01.18
	End Function

	'检查验证码是否正确
	Public Function CodeIsTrue()
		Dim CodeStr
		CodeStr=Lcase(Trim(Request.Form("CodeStr")))
		If CStr(Session("GetCode"))=CStr(CodeStr) And CodeStr<>""  Then
			CodeIsTrue=True
			'Session("GetCode")=empty
		Else
			CodeIsTrue=False
			'Session("GetCode")=empty
		End If	
	End Function
	
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
			fString = Replace(fString, CHR(10) & CHR(10), "</p><p> ")
			fString = Replace(fString, CHR(10), "<br/> ")
			fString=ChkBadWords(fString)
			HTMLEncode = fString
		End If
	End Function
	'用于用户发布的各种信息过滤，带脏话过滤 不过滤HTML 用于 公告 短信 小字报等
	Public Function TextEnCode(fString)
		If Not IsNull(fString) Then
			fString = Replace(fString, CHR(32), " ")
			fString = Replace(fString, CHR(9), "&nbsp;")
			fString = Replace(fString, CHR(13), "")
			fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
			fString = Replace(fString, CHR(10), "<br/>")
			fString=ChkBadWords(fString)
			TextEnCode = fString
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
			fString = Replace(fString, CHR(10) & CHR(10), "</p><p> ")
			fString = Replace(fString, CHR(10), "<br/> ")
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

	Function strCut(str,strlen)
		Dim l,t,c,i
		'On Error Resume Next
		str = Replacehtml(str) Rem 去掉HTML标记
		l=len(str):t=0
		For i=1 To l
			c=Abs(Asc(Mid(str,i,1)))
			If c>255 Then
				t=t+2
			Else
				t=t+1
			End If
			If t>=strlen Then
				strCut=left(str,i)&"..."
				Exit Function
			Else
				strCut=str
			End If
		Next
		strCut=Left(str,strlen)
	End Function

	Function GetIndex(S)
		Dim i
		GetIndex=-1
		For i = 0 To UBound(BadWords)
			if(BadWords(i)=s) Then
				GetIndex=i
				Exit For
			End If
		Next
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
						Application.Lock
						Application(CacheName &"_information_" & bid).documentElement.attributes.setNamedItem(Application(CacheName &"_information_" & bid).createNode(2,"lastupdate","")).text=Now()
						Application.unLock
				Else
					If Application(CacheName &"_information_" & bid).documentElement.selectSingleNode("@lastupdate") is Nothing Then
						Application.Lock
						Application(CacheName &"_information_" & bid).documentElement.attributes.setNamedItem(Application(CacheName &"_information_" & bid).createNode(2,"lastupdate","")).text=Now()
						Application.unLock
						LoadBoardinformation bid
						Application.Lock
						Application(CacheName &"_information_" & bid).documentElement.attributes.setNamedItem(Application(CacheName &"_information_" & bid).createNode(2,"lastupdate","")).text=Now()
						Application.unLock
					Else
						If DateDiff("s",Application(CacheName &"_information_" & bid).documentElement.selectSingleNode("@lastupdate").text,Now()) > 30 Then
						Application.Lock
						Application(CacheName &"_information_" & bid).documentElement.attributes.setNamedItem(Application(CacheName &"_information_" & bid).createNode(2,"lastupdate","")).text=Now()
						Application.unLock
						LoadBoardinformation bid
						Application.Lock
						Application(CacheName &"_information_" & bid).documentElement.attributes.setNamedItem(Application(CacheName &"_information_" & bid).createNode(2,"lastupdate","")).text=Now()
						Application.unLock
						End If
						End If
				End If
			'Application.Lock
			'Application.Contents.Remove(CacheName &"_information_" & bid)
			'Application.unLock
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
		Set GetBrowser=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
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
		Dvbbs.Plus_Execute(Sql)
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

	Public Sub Cache_GroupName()
		Dim Rs,Sql
		Set Rs = Dvbbs.Execute("Select ID,GroupName,UserNum,Stats From Dv_GroupName order by id")
		If Not Rs.Eof Then
			Sql = Rs.GetRows(-1)
		End If
		Rs.Close
		Set Rs = Nothing
		Application(CacheName & "_GroupName") = Sql
		'Response.Write "更新"
	End Sub

	Rem  Flash 组图插件开始
	Public Sub Load_qcomic_plus()
		If Not IsObject( Application(Dvbbs.CacheName&"_qcomic_plus")) Then
			Dim XMLDom_,rs_,Node
			Set rs_ = Dvbbs.Execute("Select * From Dv_Qcomic")
			Set XMLDom_=Dvbbs.RecordsetToxml(rs_,"setting","qcomic")
			If Not XMLDom_.documentElement.hasChildNodes Then
				Set Node = XMLDom_.documentElement.appendChild(XMLDom_.createNode(1,"setting",""))
				Node.setAttribute "senable",0
			End If
			Application.Lock
			Set Application(Dvbbs.CacheName & "_qcomic_plus") = XMLDom_'.documentElement.cloneNode(True)
			Application.UnLock
		End If
	End Sub
	
	Property Get qcomic_plus()
		Load_qcomic_plus()
		
		Dim str_senable,senable,boardlist,plus_enable
		str_senable = Application(Dvbbs.CacheName & "_qcomic_plus").documentElement.selectSingleNode("setting/@senable").text
		senable = Split(str_senable,"|||")(0)
		boardlist = Split(str_senable,"|||")(1)
		plus_enable = false
		If (Instr(boardlist, ",0,")>0 And senable="1") Then plus_enable = True
		If (Instr(boardlist, ","&Dvbbs.Boardid&",")>0 And senable="1") Then plus_enable = True
		If (plus_enable=True) Then
			qcomic_plus = True
		Else
			qcomic_plus = False
		End If
	End Property
	
	Function qcomic_plus_setting()
		Load_qcomic_plus()
	
		qcomic_plus_setting = Application(Dvbbs.CacheName & "_qcomic_plus").documentElement.selectSingleNode("setting/@semail").text
		qcomic_plus_setting = qcomic_plus_setting & "||||" & Application(Dvbbs.CacheName & "_qcomic_plus").documentElement.selectSingleNode("setting/@sid").text
		qcomic_plus_setting = qcomic_plus_setting & "||||" & Application(Dvbbs.CacheName & "_qcomic_plus").documentElement.selectSingleNode("setting/@spassword").text
		qcomic_plus_setting = qcomic_plus_setting & "||||" & Application(Dvbbs.CacheName & "_qcomic_plus").documentElement.selectSingleNode("setting/@skey").text
		qcomic_plus_setting = qcomic_plus_setting & "||||" & Application(Dvbbs.CacheName & "_qcomic_plus").documentElement.selectSingleNode("setting/@owidth").text
		qcomic_plus_setting = qcomic_plus_setting & "||||" & Application(Dvbbs.CacheName & "_qcomic_plus").documentElement.selectSingleNode("setting/@oheight").text
		qcomic_plus_setting = qcomic_plus_setting & "||||" & Application(Dvbbs.CacheName & "_qcomic_plus").documentElement.selectSingleNode("setting/@iwidth").text
		qcomic_plus_setting = qcomic_plus_setting & "||||" & Application(Dvbbs.CacheName & "_qcomic_plus").documentElement.selectSingleNode("setting/@iheight").text
	End Function
	
	Public Function ReplaceUbb(Textstr)
		Dim Str,re
		Str=Textstr
		Set re=new RegExp
			re.IgnoreCase =True
			re.Global=True
			re.Pattern="\[(.[^]]*)\]"
			Str=re.Replace(Str, "")
			Set Re=Nothing
			ReplaceUbb=Str
	End Function
	Rem  Flash 组图插件结束

End Class

Class cls_Templates
	Public Folder,ChildFolder,tplname,Cache
	Private Sub Class_Initialize()
		Folder="Resource"
		ChildFolder="Template_1"
		Cache=False
	End Sub

	Public Function Html(index)
		if(Not Cache ) Then
			Html=Dvbbs.ReadTextFile(Folder&"\"&ChildFolder&"\"&tplname&"_html"&index&".htm")
		Else
			Dvbbs.Name="tpl" & Folder&"\"&ChildFolder&"\"&tplname&"_html"&index
			if Dvbbs.ObjIsEmpty() Then
				Dvbbs.Value=Dvbbs.ReadTextFile(Folder&"\"&ChildFolder&"\"&tplname&"_html"&index&".htm")
			End if
			Html=Dvbbs.Value
		End If
		html=Replace(html,"{$PicUrl}",Dvbbs.Forum_PicUrl)
	End Function

	Public Function Pic(index)
		if(Not Cache ) Then
			Pic=Dvbbs.Forum_PicUrl&Dvbbs.ReadTextFile((Folder&"\"&ChildFolder&"\")&tplname&"_Pic"&index&".htm")
		Else
			Dvbbs.Name="tpl" & Folder&"\"&ChildFolder&"\"&tplname&"_Pic"&index
			if Dvbbs.ObjIsEmpty() Then
				Dvbbs.Value=Dvbbs.ReadTextFile(Folder&"\"&ChildFolder&"\"&tplname&"_Pic"&index&".htm")
			End If
			If InStr(Dvbbs.Value,"{$PicUrl}")>0 Then
				Pic=Dvbbs.Forum_PicUrl&Dvbbs.Value
			Else
				Pic=Dvbbs.Value
			End if
		End If
		Pic=Replace(Pic,"{$PicUrl}","")
	End Function

	Public Function Strings(index)
		if(Not Cache ) Then
			Strings=Dvbbs.ReadTextFile((Folder&"\"&ChildFolder&"\")&tplname&"_Strings"&index&".htm")
		Else
			Dvbbs.Name="tpl" & Folder&"\"&ChildFolder&"\"&tplname&"_Strings"&index
			if Dvbbs.ObjIsEmpty() Then
				Dvbbs.Value=Dvbbs.ReadTextFile(Folder&"\"&ChildFolder&"\"&tplname&"_Strings"&index&".htm")
			End if
			Strings=Dvbbs.Value
		End If
	End Function

	Private Sub class_terminate()
	End Sub
End Class

Class cls_UserOnlne
	Public Forum_Online,Forum_UserOnline,Forum_GuestOnline
	Private l_Online,l_GuestOnline
	Private Sub Class_Initialize()
		Dim Reloadtime
		Reloadtime=Dvbbs.Reloadtime
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
		Dvbbs.Reloadtime=Reloadtime
	End Sub

	Public Sub OnlineQuery()
		Dim SQL,SQL1,Delflag
		Dim TempNum,TempNum1
		Delflag=False
		Dvbbs.Name="delOnline_time"
		If Dvbbs.ObjIsEmpty() Then
			Delflag=True:Dvbbs.Value=Now()
		Else
			If DateDiff("s",Dvbbs.Value,Now()) > Clng(Dvbbs.Forum_Setting(8))*10 Then Delflag=True
		End If
		If Delflag Then
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
		Rs.close()
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
Dvbbs.PageInit()
%>