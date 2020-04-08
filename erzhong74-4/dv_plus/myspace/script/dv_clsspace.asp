<%
'个性空间类模块
'Writen by Sunwin
Class Cls_Space
Public ScriptPath
Public Sid,Act,ReCache
Public XmlDoc,Space_Info,Forum_info,Forum_User,Space_User,Admin
Public Space_Skinpath
Private Forum_Skinpath

Private Sub Class_Initialize()
	ScriptPath = MyDbPath & "Dv_plus/myspace/script/"
	ReCache = 0 '是否首页缓存，0=是，1=否
	Sid = Dvbbs.CheckNumeric(Request("sid"))
	Act = Request("act")
	Admin = false
	'Forum_Skinpath = Application(Dvbbs.CacheName & "_csslist").documentElement.selectSingleNode("@cssfilepath").text
	Forum_Skinpath=Dvbbs.Forum_PicUrl
	Space_Skinpath = MyDbPath & Replace(Forum_Skinpath,"Default/","") &"myspace/"
	If Sid=0 and Dvbbs.Userid>0 Then
		Sid = Dvbbs.Userid
	End If

	If Dvbbs.userid=0 Then
		'Dvbbs.AddErrCode(6)
		'Dvbbs.Showerr()
	Else
		If Dvbbs.UserId=Sid Or Dvbbs.Master Then
			Admin = true
		End If
	End If
	Set XmlDoc = Dvbbs.CreateXmlDoc("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	LoadData()
End Sub

Private Sub class_terminate()
	Set XmlDoc = Nothing
End Sub

'获取数据
Public Sub LoadData()

	If Act="" and ReCache = 0 Then
		If LoadCache_Data Then
			ReCache=0
		Else
			ReCache=1
		End If
	Else
		ReCache=1
	End If
	If ReCache=0 Then
		'读取文件缓存
		'XmlDoc.Load Server.MapPath(MyDbPath &"space.xml")
		'建立数据对象
		Set Forum_info = XmlDoc.DocumentElement.selectSingleNode("forum_info")
		Set Space_Info = XmlDoc.documentElement.selectSingleNode("space_info")
		Set Space_User = XmlDoc.documentElement.selectSingleNode("space_user")
		If Not (XmlDoc.DocumentElement.selectSingleNode("forum_user") is Nothing) Then
			XmlDoc.DocumentElement.removeChild(XmlDoc.DocumentElement.selectSingleNode("forum_user"))
		End If
		Set Forum_User = XmlDoc.DocumentElement.appendChild(XmlDoc.createNode(1,"forum_user",""))
		Forum_User.appendChild(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo").cloneNode(True))

		If Act<>"" Then
			Space_Info.removeChild(Space_Info.selectSingleNode("//leftchannal"))
			Space_Info.removeChild(Space_Info.selectSingleNode("//rightchannal"))
			If Space_Info.getAttribute("s_style") <>"2"  Then
				SpaceChannal Space_Info,Space_Info.getAttribute("s_left"),"leftchannal"
			End If
			If Space_Info.getAttribute("s_style") ="2" Then
				SpaceChannal Space_Info,Space_Info.getAttribute("s_right"),"rightchannal"
			End If
		End If
	Else
		XmlDoc.Loadxml("<?xml version=""1.0"" encoding=""gb2312""?><dv_space/>")
		'加载论坛基本信息
		SetupForum_Info()
		'加载个人空间基本信息
		LoadUserSpaceData()
		'加载当前访问用户基本数据
		Set Forum_User = XmlDoc.DocumentElement.appendChild(XmlDoc.createNode(1,"forum_user",""))
		Forum_User.appendChild(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo").cloneNode(True))
		'加载空间用户数据
		LoadSpaceUserData()
	End If
	If Act <> "modifyskin" Then
		Space_Info.removeAttribute "s_css"
	End If
	Space_Info.setAttribute "s_stylevalue",Space_Info.getAttribute("s_style")
	If Space_Info.getAttribute("s_style")="3" and Act<>"" Then
		Space_Info.setAttribute "s_style","1"
	End If
	
End Sub

Public Function LoadCache_Data()
	Dim Rs,Sql,UpTime
	LoadCache_Data = False
	If Admin Then
		Sql = "Select Top 1 id,Userid,ownercachedb From [Dv_Space_user] where Userid="&sid
	Else
		Sql = "Select Top 1 id,Userid,cachedb From [Dv_Space_user] where Userid="&sid
	End If
	Set Rs=Dvbbs.Execute(Sql)
	If Rs.Eof Then
		'该用户还没开通个性首页
		Set Space_Info = GetDefaultData.cloneNode(True)
	Else
		If XmlDoc.LoadXml(Rs(2)&"") Then
			'加放更新时间判断 set_5频道更新时间，单位分钟
			Set Space_Info = XmlDoc.documentElement.selectSingleNode("space_info")
			If Not (Space_Info is Nothing) Then
				If IsDate(Space_Info.getAttribute("cacheuptime")) Then
					UpTime = cdate(Space_Info.getAttribute("cacheuptime"))
				Else
					UpTime = cdate(Space_Info.getAttribute("updatetime"))
				End If
				
				If Datediff("n",UpTime,Now())<Int(Space_Info.getAttribute("set_5")) Then
					LoadCache_Data = True
				End If
			End If
		End If
	End If
End Function

'保存首页缓存
Public Sub SaveCache_Data()
	If Sid=0 Then Exit Sub
	If ReCache=1 and Act="" Then
		Space_Info.setAttribute "cacheuptime",now()
		If Admin Then
			Dvbbs.Execute("update [Dv_Space_user] Set ownercachedb='"&Dvbbs.Checkstr(XmlDoc.xml)&"',updatetime = "&SqlNowString&"  where Userid="&sid)
		Else
			Dvbbs.Execute("update [Dv_Space_user] Set cachedb='"&Dvbbs.Checkstr(XmlDoc.xml)&"',updatetime = "&SqlNowString&"  where Userid="&sid)
		End If
		'Response.Write "已更新"
	End If
	'XmlDoc.save Server.MapPath(MyDbPath &"space.xml")
End Sub


'个人空间基本信息
Public Sub LoadUserSpaceData()
	Dim Rs,Sql,Node
	Dim Setting,i
	Sql = "Select Top 1 id,userid,username,title,intro,s_left,s_right,s_center,s_css,s_style,s_path,updatetime,lock,[set],plusdb from [Dv_Space_user] where"
	Sql = Sql & " Userid=" & Sid
	Set Rs=Dvbbs.Execute(Sql)
	If Rs.Eof Then
		Set Space_Info = GetDefaultData.cloneNode(True)
	Else
		Set Node = Dvbbs.RecordsetToxml(rs,"space_info","")
		Set Space_Info = Node.documentElement.selectSingleNode("space_info")
	End If
	Rs.Close
	Set Rs = Nothing
	Setting = Split(Space_Info.getAttribute("set")&"",",")
	If Ubound(Setting)=20 Then
		For i=0 to Ubound(Setting)
			Space_Info.setAttribute "set_"&i,Setting(i)
		Next
		Space_Info.removeAttribute "set"
	End If
	Dim TempXmlDoc
	Set TempXmlDoc = Dvbbs.CreateXmlDoc("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	If TempXmlDoc.LoadXml(Space_Info.getAttribute("plusdb")) Then
		Space_Info.appendChild(TempXmlDoc.DocumentElement.cloneNode(True))
	Else
		Space_Info.appendChild(TempXmlDoc.createNode(1,"modules",""))
	End If
	Space_Info.removeAttribute "plusdb"
	Set TempXmlDoc = Nothing
	SpaceChannal Space_Info,Space_Info.getAttribute("s_left"),"leftchannal"
	SpaceChannal Space_Info,Space_Info.getAttribute("s_center"),"mainchannal"
	SpaceChannal Space_Info,Space_Info.getAttribute("s_right"),"rightchannal"
	Space_Info.setAttribute "skinpath",Space_Skinpath & Space_Info.getAttribute("s_path")
	Space_Info.setAttribute "isadmin",Admin
	XmlDoc.DocumentElement.appendChild(Space_Info)
End Sub

'建立频道列表
Private Sub SpaceChannal(Node,Arr,Name)
	Dim S_channal,i,ChildNode,Childs
	If Act<>"" Then
		If name="mainchannal" Then
			Arr = ""
		End If
		If (Space_Info.getAttribute("s_style") <>"2" and name="leftchannal") or (Space_Info.getAttribute("s_style") ="2" and name="rightchannal") Then
			Arr = "userinfo"
		Else
			Arr = ""
		End If
	End If
	S_channal = Split(Lcase(Arr),",")
	Set ChildNode = XmlDoc.createNode(1,Name,"")
	For i=0 to Ubound(S_channal)
		Set Childs = XmlDoc.createNode(1,"channals","")
		Childs.setAttribute "id",S_channal(i)
		ChildNode.appendChild(Childs)
	Next
	Node.appendChild(ChildNode)
End Sub


Private Function GetDefaultData()
	Dim TempXmlDoc,Node
	Set TempXmlDoc = Dvbbs.CreateXmlDoc("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	If Not TempXmlDoc.Load(Server.MapPath(ScriptPath&"myspace.xml")) Then
		Response.Write "Load MySpace default data failse！"
		Response.End
	End If
	Set GetDefaultData = TempXmlDoc.documentElement.selectSingleNode("space_info")
	Set TempXmlDoc = Nothing
End Function

'个人空间所属用户基本信息
Public Sub LoadSpaceUserData()
	Dim Rs,Sql,Node,ChildNode
	Dim Attribute,UserIM,i,UserInfo,UserFace
	Sql="Select top 1 UserID,UserName,UserEmail,UserPost,UserTopic,UserFav,UserIsBest,UserSex,UserFace,UserWidth,UserHeight,JoinDate,LastLogin,UserLogins,Lockuser,Userclass,UserGroupID,UserGroup,userWealth,userEP,userCP,UserPower,UserBirthday,UserLastIP,UserDel,UserIsBest,UserHidden,UserMsg,IsChallenge,UserMobile,TitlePic,UserTitle,UserToday,UserMoney,UserTicket,FollowMsgID,Vip_StarTime,Vip_EndTime,Userim,Userinfo,UserSetting,Fav_Boards"
	Sql = Sql & " From [Dv_User] Where UserID = " & Sid
	Set Rs=Dvbbs.Execute(Sql)
	
	If Rs.Eof Then
		Set Space_User = XmlDoc.createNode(1,"space_user","")
	Else
		Set Node = Dvbbs.RecordsetToxml(rs,"space_user","")
		Set Space_User = Node.documentElement.selectSingleNode("space_user")
	End If
	Rs.Close
	Set Rs = Nothing

	If Space_User.getAttribute("userface")<>"" Then
		If Instr(Space_User.getAttribute("userface"),"|") Then
			Space_User.setAttribute "userface",Split(Space_User.getAttribute("userface"),"|")(1)
		End If
	End If
	If Space_User.getAttribute("userim")<>"" Then
		UserIM = Split(Space_User.getAttribute("userim"),"|||")
		If not IsArray(UserIM) Then ReDim UserIM(6)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"homepage",""))
		ChildNode.text = UserIM(0)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"qq",""))
		ChildNode.text = UserIM(1)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"icp",""))
		ChildNode.text = UserIM(2)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"msn",""))
		ChildNode.text = UserIM(3)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"aim",""))
		ChildNode.text = UserIM(4)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"yahoo",""))
		ChildNode.text = UserIM(5)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"uc",""))
		ChildNode.text = UserIM(6)
		Space_User.removeAttribute "userim"
	End If
	If Space_User.getAttribute("userinfo")<>"" Then
		UserInfo = Split(Space_User.getAttribute("userinfo"),"|||")
		If not IsArray(UserInfo) Or Ubound(UserInFo)<>14 Then ReDim UserInfo(14)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"realname",""))
		ChildNode.text = UserInfo(0)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"character",""))
		ChildNode.text = UserInfo(1)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"personal",""))
		ChildNode.text = DVbbs.htmlencode(UserInfo(2))
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"contry",""))
		ChildNode.text = UserInfo(3)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"province",""))
		ChildNode.text = UserInfo(4)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"city",""))
		ChildNode.text = UserInfo(5)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"shengxiao",""))
		ChildNode.text = UserInfo(6)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"blood",""))
		ChildNode.text = UserInfo(7)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"belief",""))
		ChildNode.text = UserInfo(8)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"occupation",""))
		ChildNode.text = UserInfo(9)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"marital",""))
		ChildNode.text = UserInfo(10)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"education",""))
		ChildNode.text = UserInfo(11)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"college",""))
		ChildNode.text = UserInfo(12)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"phone",""))
		ChildNode.text = UserInfo(13)
		Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"address",""))
		ChildNode.text = UserInfo(14)
		Space_User.removeAttribute "userinfo"
	End If
	If Space_User.getAttribute("userdel")<>"" Then
		Dim UserDel
		UserDel = Dvbbs.CheckNumeric(Space_User.getAttribute("userdel"))
		UserDel = UserDel * -1
		If Space_User.getAttribute("userpost")>0 then
		Space_User.setAttribute "userdelpercent",FormatPercent(UserDel/(Clng(Space_User.getAttribute("userpost"))+UserDel))
		End if
	Else
		Space_User.setAttribute "userdelpercent","0%"
	End If
	If Space_User.getAttribute("usersetting")="" Or IsNull(Space_User.getAttribute("usersetting")) Then Space_User.setAttribute "usersetting","1|||0|||0|||0"
	'If Space_User.getAttribute("usersetting")<>"" Then
		Dim UserSetting
		UserSetting = Split(Space_User.getAttribute("usersetting"),"|||")
		If Ubound(UserSetting)>1 Then
			Space_User.setAttribute "set1",UserSetting(0)
			Space_User.setAttribute "set2",UserSetting(1)
			Space_User.setAttribute "set3",UserSetting(2)
			If Ubound(UserSetting)<3 Then
				Space_User.setAttribute "set4",0
			Else
				Space_User.setAttribute "set4",UserSetting(3)
			End If
		Else
			Space_User.setAttribute "set1",1
			Space_User.setAttribute "set2",0
			Space_User.setAttribute "set3",0
			Space_User.setAttribute "set4",0
		End If
		Space_User.removeAttribute "usersetting"
	'End If
	If Space_User.getAttribute("userfav")<>"" Then
		Dim FriendMod
		i = 0
		For Each FriendMod in Split(Space_User.getAttribute("userfav"),",")
			Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"userfav",""))
			ChildNode.setAttribute "m",i
			ChildNode.setAttribute "name",FriendMod
			i = i+1
		Next
		Space_User.removeAttribute "userfav"
	End If
	'fav_boards
	If Space_User.getAttribute("fav_boards")<>"" Then
		Dim FavBoardID
		i = 0
		For Each FavBoardID in Split(Space_User.getAttribute("fav_boards"),",")
			Set ChildNode = Space_User.appendChild(XmlDoc.createNode(1,"favbid",""))
			ChildNode.setAttribute "bid",FavBoardID
			i = i+1
		Next
		Space_User.removeAttribute "fav_boards"
	End If
	XmlDoc.DocumentElement.appendChild(Space_User)
End Sub

Private Sub SetupForum_Info()
	Dim Powered
	Powered = "Powered By <a href = ""http://www.dvbbs.net/"" target = ""_blank"">Dvbbs</a>  <a href = ""http://www.dvbbs.net/download.asp"" target = ""_blank"">Version " & fVersion & "</a>"
		Set Forum_info=XmlDoc.DocumentElement.appendChild(XmlDoc.createNode(1,"forum_info",""))
		Forum_info.setAttribute "type",Dvbbs.Forum_info(0)
		Forum_info.setAttribute "maxonline",Dvbbs.CacheData(5,0)
		Forum_info.setAttribute "maxonlinedate",Dvbbs.CacheData(6,0)
		Forum_info.setAttribute "topicnum",Dvbbs.CacheData(7,0)
		Forum_info.setAttribute "postnum",Dvbbs.CacheData(8,0)
		Forum_info.setAttribute "maxonline",Dvbbs.CacheData(5,0)
		Forum_info.setAttribute "maxonlinedate",Dvbbs.CacheData(6,0)
		Forum_info.setAttribute "topicnum",Dvbbs.CacheData(7,0)
		Forum_info.setAttribute "postnum",Dvbbs.CacheData(8,0)
		Forum_info.setAttribute "todaynum",Dvbbs.CacheData(9,0)
		Forum_info.setAttribute "usernum",Dvbbs.CacheData(10,0)
		Forum_info.setAttribute "yesterdaynum",Dvbbs.CacheData(11,0)
		Forum_info.setAttribute "maxpostnum",Dvbbs.CacheData(12,0)
		Forum_info.setAttribute "maxpostdate",Dvbbs.CacheData(13,0)
		Forum_info.setAttribute "lastuser",Dvbbs.CacheData(14,0)
		Forum_info.setAttribute "online",MyBoardOnline.Forum_Online
		Forum_info.setAttribute "useronline",MyBoardOnline.Forum_UserOnline
		Forum_info.setAttribute "guestonline",MyBoardOnline.Forum_GuestOnline
		Forum_info.setAttribute "createtime",FormatDateTime(Dvbbs.Forum_Setting(74),1)
		Forum_info.setAttribute "url",Dvbbs.Get_ScriptNameUrl()
		Forum_info.setAttribute "skinpath",Forum_Skinpath
		Forum_info.setAttribute "copyright",Dvbbs.Forum_Copyright
		Forum_info.setAttribute "powered",Powered
		Forum_info.setAttribute "picurl",Dvbbs.Forum_PicUrl
End Sub

'输出模板页面
Public Sub TranTemplate()
	Forum_info.setAttribute "act",Act
	Forum_info.setAttribute "querynum",Dvbbs.SqlQueryNum
	Forum_info.setAttribute "runtime","0" & FormatNumber((Timer()-Startime),5)
	Dim Xmlskin,Proc,XmlStyle
	Set Xmlskin = Dvbbs.CreateXmlDoc("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	If Not (Xmlskin.load(Server.MapPath(ScriptPath &"myspace.xslt"))) Then
		Response.Write "模板数据出错，请与管理员联系！"
		Response.End
	End If
	Set XMLStyle=Dvbbs.iCreateObject("Msxml2.XSLTemplate" & MsxmlVersion)
	XMLStyle.stylesheet=Xmlskin
	Set Proc=XMLStyle.createProcessor()
	Proc.input = XmlDoc
  	proc.transform()
  	Response.Write proc.output
	Set XmlStyle = Nothing
	Set Xmlskin = Nothing
	SaveCache_Data()
End Sub

Public Sub head()
		Dvbbs.Stats=Replace(Dvbbs.Stats,Chr(13),"")
		Dvbbs.Stats=Dvbbs.Replacehtml(Dvbbs.stats)
		'搜索引擎优化部分
		If Request("IsSearch_a") <> "" Then Dvbbs.stats = Dvbbs.stats & "-网站地图"
		If Dvbbs.IsSearch Then
			Response.Write Replace(Replace(Replace(Dvbbs.mainhtml(1),"{$keyword}",Replace(Dvbbs.Forum_info(8),"|",",")),"{$description}",Dvbbs.Forum_info(10))&vbNewLine,"{$title}",Dvbbs.stats &"["& Dvbbs.Forum_Info(0) &"] -- Powered By Dvbbs.net," & Now())
		Else
			Response.Write Replace(Replace(Replace(Dvbbs.mainhtml(1),"{$keyword}",Replace(Dvbbs.Forum_info(8),"|",",")),"{$description}",Dvbbs.Forum_info(10))&vbNewLine,"{$title}",Dvbbs.stats &"["& Dvbbs.Forum_Info(0) &"]")
		End If
		'搜索引擎优化结束
		Response.Write Chr(10)
		Response.Write Dvbbs.mainhtml(2)
		Response.Write "<script language=""javascript"" type=""text/javascript"">"
		'Response.Write "var boardxml='<?xml version=""1.0"" encoding=""gb2312""?>"& replace(XMLDom.documentElement.XML ,"'","\'")&"';"
		Response.Write "var boardxml='',ISAPI_ReWrite = "&isUrlreWrite&",forum_picurl='"&Dvbbs.Forum_PicUrl&"';"
		Response.Write "</script>"
End Sub

Public Sub Suc(msg)
	Dvbbs.Head()
	Dvbbs.Dvbbs_Suc(msg)

End Sub

'创建用户风格目录，需要FSO支持，若不支持返回空
'目录格式：skins/myspace/users/skin_@Userid
Public Function CreatStylePath()
	Dim SkinPath,Fso
	CreatStylePath = ""
	SkinPath = Space_Skinpath&"userskins/skin_"&Sid&"/"
	On Error Resume Next
	Set Fso = Dvbbs.iCreateObject("Scripting.FileSystemObject")
		If Err Then
			Err.clear
			Exit Function
		End if
		If Fso.FolderExists(Server.MapPath(SkinPath))=False Then
			Fso.CreateFolder(Server.MapPath(SkinPath))
		End If
		If Err.Number = 0 Then
			CreatStylePath = "userskins/skin_"&Sid&"/"
		Else
			Err.Clear
			Exit Function
		End If
	Set Fso = Nothing
End Function

'风格目录复制
Public Sub CopyFolder(Folder1,Folder2)
	On Error Resume Next
	Dim Fso
	Set Fso = Dvbbs.iCreateObject("Scripting.FileSystemObject")
		If Err Then
			Err.clear
			Exit Sub
		End if
		Fso.CopyFolder Server.MapPath(Space_Skinpath&Folder1),Server.MapPath(Space_Skinpath&Folder2),true
		If Err Then
			Err.Clear
			Exit Sub
		End If
	Set Fso = Nothing
End Sub

Public Sub WriteFile(Filepath,Text)
	On error resume Next
	Dim FileName,Fso
	FileName = Server.MapPath(Filepath)
	Set FSO=Dvbbs.iCreateObject("Scripting.FileSystemObject")
	If Err Then
		Response.Write "<br /><li>*您的服务器不支持写文件(*"&Err.Description&"),CSS文件写入失败,请手工操作或把生成文件的内容清空!</li>"
		Err.Clear
		Exit Sub
	End If
	Fso.CreateTextFile(FileName).WriteLine(Text)
	If  Err Then
		Response.Write "<br /><li>*您的服务器不支持写文件(*"&Err.Description&"),CSS文件写入失败,请手工操作或把生成文件的内容清空!</li>"
		err.Clear
		Exit Sub
	End If
	Set Fso = Nothing
End Sub
End Class


%>