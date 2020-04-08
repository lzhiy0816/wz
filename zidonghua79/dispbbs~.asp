<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/dv_ubbcode.asp"-->
<%
If Dvbbs.BoardID < 1 Then
	Response.Write "参数错误"
	Response.End 
End If
Dvbbs.LoadTemplates("dispbbs")
Dim AnnounceID,ReplyID,Star,Skin,replyid_a,AnnounceID_a,RootID_a
Dim CanReply,IsTop,IsVote,TopicCount,PollID,TotaluseTable,mainTopic,TrueMaster
Dim CanRead,TopicInfo
Dim PostBuyUser,abgcolor,bgcolor,UserName,PostUserName
Dim Page,LockTopic,action,n,EmotPath
Dim dv_ubb
Dim T_GetMoneyType,T_GetMoney,T_UseTools
Chk_Topic_Err
Dvbbs.Showerr()
Response.Write Dvbbs.mainhtml(18)
Response.Cookies("Dvbbs")=GetFormID()
Dvbbs.Nav()
Dvbbs.Showerr()
Dvbbs.Head_var 1,"","",""
GetForumTextAd(2)
Dvbbs.ActiveOnline()
EmotPath=Split(Dvbbs.Forum_emot,"|||")(0)		'em心情路径
action=Request("action")
Page=Request("Page")
If isNumeric(Page) = 0 or Page="" Then Page=1
Page=Clng(Page)
Dvbbs.ShowErr()
Set dv_ubb=new Dvbbs_UbbCode
dv_ubb.PostType=1
Show_Topic_HTML()
Set dv_ubb=Nothing
Dvbbs.NewPassword()
Dvbbs.Footer()

Sub Chk_Topic_Err()
	AnnounceID=Request("ID")
	If Dvbbs.GroupSetting(2)="0" Then Dvbbs.AddErrcode(31)
	If AnnounceID="" Or Not IsNumeric(AnnounceID) Then Dvbbs.AddErrCode(30)
	ReplyID=Request("ReplyID")
	If ReplyID="" Or Not IsNumeric(ReplyID) Then ReplyID=AnnounceID
	Star=Request("Star")
	If Star="" Or Not IsNumeric(Star) Then Star=1
	Star=Clng(Star)
	Skin=Request("Skin")
	If Skin="" Or Not IsNumeric(Skin) Then Skin=Dvbbs.Board_setting(42)
	If Dvbbs.ErrCodes<>"" Then Exit Sub
	Dim SQl,Rs
	'浏览购买帖权限
	CanRead=False
	TrueMaster=False
	Rem 为兼顾管理菜单显示,对有管理权限的暂时当版主等级处理,为的是显示管理菜单.
	If Not Dvbbs.BoardMaster Then
		If Dvbbs.UserID > 0 Then
			If Dvbbs.GroupSetting(18) = "1" Then
				Dvbbs.BoardMaster=True
			ElseIf Dvbbs.GroupSetting(19) = "1" Then
				Dvbbs.BoardMaster=True
			ElseIf Dvbbs.GroupSetting(20) = "1" Then
				Dvbbs.BoardMaster=True
			ElseIf Dvbbs.GroupSetting(21) = "1" Then
				Dvbbs.BoardMaster=True
			ElseIf Dvbbs.GroupSetting(22) = "1" Then
				Dvbbs.BoardMaster=True
			ElseIf Dvbbs.GroupSetting(23) = "1" Then
				Dvbbs.BoardMaster=True
			ElseIf Dvbbs.GroupSetting(24) = "1" Then
				Dvbbs.BoardMaster=True
			ElseIf Dvbbs.GroupSetting(25) = "1" Then
				Dvbbs.BoardMaster=True
			ElseIf Dvbbs.GroupSetting(26) = "1" Then
				Dvbbs.BoardMaster=True
			ElseIf Dvbbs.GroupSetting(27) = "1" Then
				Dvbbs.BoardMaster=True
			ElseIf Dvbbs.GroupSetting(28) = "1" Then
				Dvbbs.BoardMaster=True
			ElseIf Dvbbs.GroupSetting(29) = "1" Then
				Dvbbs.BoardMaster=True
			ElseIf Dvbbs.GroupSetting(30) = "1" Then
				Dvbbs.BoardMaster=True
			ElseIf Dvbbs.GroupSetting(31) = "1" Then
				Dvbbs.BoardMaster=True
			End If
		End If
	Else
		TrueMaster=True	
	End If
	If Dvbbs.BoardMaster Then CanRead=True	
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	SQL="Select title,istop,isbest,PostUserName,PostUserid,hits,isvote,child,pollid,LockTopic,PostTable,BoardID,TopicMode,GetMoney,UseTools,GetMoneyType,DateAndTime,Expression,topicID,topicID as ReplyID,topicID as star,topicID as page,topicID as skin  From DV_topic where topicID="&Announceid
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open SQL,Conn,1,3
	Dvbbs.SqlQueryNum=Dvbbs.SqlQueryNum+1
	'Set Rs=Dvbbs.Execute(SQL)
	If Not(Rs.BOF and Rs.EOF) then
		If Rs(11)<>Dvbbs.BoardID Then Dvbbs.AddErrCode(29)
		Rs(5)=Rs(5)+1
		Rs.Update
		mainTopic=Rs(0)
		istop=Rs(1)
		isVote=Rs(6)
		TopicCount=Rs(7)+1
		pollid=Rs(8)
		T_GetMoney = cCur(Rs(13))
		T_UseTools = Rs(14)
		T_GetMoneyType = Cint(Rs(15))
		PostUserName = Rs(3)
		'锁定多少天前的帖子判断 2004-9-16 Dv.Yz
		If Not Ubound(Dvbbs.Board_Setting) > 70 Then
			Locktopic = Rs(9)
		Else
			If (Not Clng(Dvbbs.Board_Setting(71)) = 0) And Datediff("d", Rs(16), Now()) > Clng(Dvbbs.Board_Setting(71)) And Rs(9)=0 Then
				Locktopic = 1
				Rs(9)=Locktopic
				Rs.Update
			Else
				Locktopic = Rs(9)
			End If
		End If
		
		TotalUseTable=rs(10)
		mainTopic=Dvbbs.ChkBadWords(mainTopic)
		If Not Rs("topicmode")=1 Then
			mainTopic=replace(mainTopic,"<","&lt;")
		End If
		Dvbbs.Stats=mainTopic
		Set Topicinfo=Dvbbs.RecordsetToxml(Rs,"postinfo","xml")
	Else
		Dvbbs.AddErrcode(32)
	End If
	Rs.Close
	Set Rs=Nothing
	CanReply=False
	If Dvbbs.UserID =0 And Dvbbs.GroupSetting(3)="0" Then
		CanReply=False
	ElseIf (Not Dvbbs.Board_Setting(0)="1"  And Cint(locktopic)=0) Or (Dvbbs.master Or Dvbbs.superboardmaster Or Dvbbs.boardmaster) Then
		CanReply=True
	End If
	
End Sub

Function Topic_Ads()
		Randomize
		Topic_Ads=Dvbbs.Forum_ads(14)(CInt(UBound(Dvbbs.Forum_ads(14))*Rnd))
End Function
Sub Show_Topic_HTML()
	Dim namestyle,nameglow(7),postbuyinfo,SQL,Rs,i,XMLDom,PageCount,postarray,Node,postuseridlist,postuserlist,UserGroupID,postbody,Topic,cmd,postbuyusers,UserIM
	Dim userface
	nameglow(0)=Dvbbs.mainsetting(9):nameglow(1)=Dvbbs.mainsetting(7):nameglow(2)=Dvbbs.mainsetting(7):nameglow(3)=Dvbbs.mainsetting(5):nameglow(6)="gray":nameglow(7)=Dvbbs.mainsetting(11)
	If Dvbbs.Forum_ads(7)="1" Then
		Dvbbs.Forum_ads(14)=replace(Dvbbs.Forum_ads(14),Chr(13),"")
		Dvbbs.Forum_ads(14)=Split(Dvbbs.Forum_ads(14),Chr(10))
	End If
	If TopicCount mod Cint(Dvbbs.Board_Setting(27))=0 then
			PageCount= TopicCount \ Cint(Dvbbs.Board_Setting(27))
	Else
			PageCount= TopicCount \ Cint(Dvbbs.Board_Setting(27))+1
	End If
	If star > PageCount Then star = PageCount
	If star < 1 Then star = 1
	If Skin=1 Then
		If Clng(replyid)=Clng(Announceid) Then
			SQL="Select Top 1 AnnounceID,UserName,Topic,dateandtime,body,Expression,ip,RootID,signflag,isbest,PostUserid,layer,isagree,GetMoneyType,IsUpload,Ubblist,LockTopic,GetMoney,UseTools,PostBuyUser,ParentID From "& TotalUseTable &" where RootID="& Announceid &" and Boardid="& Dvbbs.Boardid&" "
		Else
			SQL="Select AnnounceID,UserName,Topic,dateandtime,body,Expression,ip,RootID,signflag,isbest,PostUserid,layer,isagree,GetMoneyType,IsUpload,Ubblist,LockTopic,GetMoney,UseTools,PostBuyUser,ParentID From "& TotalUseTable &" where AnnounceID="&replyID&" and Boardid="& Dvbbs.Boardid&" "
		End If
	Else
		Rem 如果是原版论坛,没经过转换的建议使用这行,可以减少消耗
		SQL="Select AnnounceID,UserName,Topic,dateandtime,body,Expression,ip,RootID,signflag,isbest,PostUserid,layer,isagree,GetMoneyType,IsUpload,Ubblist,LockTopic,GetMoney,UseTools,PostBuyUser,ParentID From "& TotalUseTable &" where  RootID="& Announceid &" and Boardid="& Dvbbs.Boardid&" Order By Announceid"
		Rem 如果你的论坛是从别的论坛转换过来的,如出现楼层错误,则可以把下面的注释去掉,避免错误.
		Rem SQL="Select AnnounceID,UserName,Topic,dateandtime,body,Expression,ip,RootID,signflag,isbest,PostUserid,layer,isagree,GetMoneyType,IsUpload,Ubblist,LockTopic,GetMoney,UseTools,PostBuyUser,ParentID From "& TotalUseTable &" where  RootID="& Announceid &" and Boardid="& Dvbbs.Boardid&" Order By dateandtime"
	End If
	Set Rs=Dvbbs.Execute(SQL)
	If Not Rs.EOF Then
		If Not Skin=1 Then 	Rs.Move(Cint(Dvbbs.Board_Setting(27))*(star-1))
		If Not Rs.EOF Then
			TopicInfo.documentElement.firstChild.selectSingleNode("@replyid").text=Rs(0)
			TopicInfo.documentElement.firstChild.selectSingleNode("@star").text=star
			TopicInfo.documentElement.firstChild.selectSingleNode("@page").text=page
			TopicInfo.documentElement.firstChild.selectSingleNode("@skin").text=skin
			If TopicInfo.documentElement.firstChild.selectSingleNode("@topicmode").text <> "1" Then
				TopicInfo.documentElement.firstChild.selectSingleNode("@title").text=replace(TopicInfo.documentElement.firstChild.selectSingleNode("@title").text,"<","&lt;")
			End If
			Rem 如果顶楼是主题贴，并且是购买贴的时候，提取购买信息,讨厌的多层数组分割，烦琐的代码跟踪by 老迷！
			If Rs("ParentID")=0 and TopicInfo.documentElement.firstChild.selectSingleNode("@getmoneytype").text="3" Then
				Set Node=TopicInfo.documentElement.firstChild.appendChild(topicInfo.createNode(1,"postbuyinfo",""))
				postbuyusers=split(Rs("postbuyuser"),"|||")
				postbuyinfo=postbuyusers(0)
				postbuyinfo=Split(postbuyinfo,"@@@") 'Rem postbuyinfo(0) 收入 postbuyinfo(1) 购买限制 postbuyinfo(2) vip是否需要购买 postbuyinfo(3) 允许购买用户列表
				If UBound(postbuyinfo)> 2 Then 
					Node.attributes.setNamedItem(TopicInfo.createNode(2,"money","")).text=postbuyinfo(0)
					Node.attributes.setNamedItem(TopicInfo.createNode(2,"maxbuy","")).text=postbuyinfo(1)
					Node.attributes.setNamedItem(TopicInfo.createNode(2,"notvipbuy","")).text=postbuyinfo(2)
					Node.attributes.setNamedItem(TopicInfo.createNode(2,"buyuser","")).text=","& postbuyinfo(3) &","
				End If
				For i= 2 to UBound(postbuyusers)
					If postbuyusers(i)<>"" Then
						Node.appendChild(topicInfo.createNode(1,"postbuyusers","")).text=postbuyusers(i)
					End If
				Next
			End If
			postarray=Rs.GetRows(Cint(Dvbbs.Board_Setting(27)))
			Set XMLDom=Dvbbs.ArrayToxml(postarray,rs,"row","post")
			postarray=null
			'循环中整理数据,并取出发贴用户ID列表
			i=0
			For Each Node In XmlDom.documentElement.SelectNodes("row")
				If i=0 Then
					postuseridlist=Node.selectSingleNode("@postuserid").text
				Else
					postuseridlist=postuseridlist&","&Node.selectSingleNode("@postuserid").text
				End If
				i=i+1
			Next
			'说明:postuserlist为发贴用户数据
			Set Rs=Dvbbs.Execute("Select userid,useremail,UserIM,UserMobile,Usersign,userclass,Usertitle,Userwidth,Userheight,UserPost,Userface,JoinDate,userWealth,userEP,userCP,Userbirthday,Usersex,UserGroup,LockUser,userPower,titlepic,UserGroupID,LastLogin,UserHidden,IsChallenge,UserMoney,UserTicket,UserAvaSetting,UserIsAva From dv_user Where UserID IN ("& postuseridlist &")")
			If Not Rs.EOF Or postuseridlist<>"" Then
				Set postuserlist=Dvbbs.RecordsetToxml(Rs,"user","userlist")
				For Each Node In XmlDom.documentElement.SelectNodes("row")
					Ubblists=Node.selectSingleNode("@ubblist").text
					If postuserlist.documentElement.selectSingleNode("user[@userid="&Node.selectSingleNode("@postuserid").text&"]/@usergroupid") Is Nothing Then
						UserGroupID=7
					Else
						UserGroupID=postuserlist.documentElement.selectSingleNode("user[@userid="&Node.selectSingleNode("@postuserid").text&"]/@usergroupid").text
					End If
					'过滤标题HTML
					If TopicInfo.documentElement.firstChild.selectSingleNode("@topicmode").text <> "1" Then
						Node.selectSingleNode("@topic").text=replace(Node.selectSingleNode("@topic").text,"<","&lt;")
					Else
						If Node.selectSingleNode("@parentid").text<>"0" Then
							Node.selectSingleNode("@topic").text=replace(Node.selectSingleNode("@topic").text,"<","&lt;")
						End If
					End If
					postbody=Node.selectSingleNode("@body").text
					'过滤脏字
					postbody=Dvbbs.ChkBadWords(postbody)
					Topic=Node.selectSingleNode("@topic").text
					Node.selectSingleNode("@topic").text=Dvbbs.ChkBadWords(Topic)
					UserName=Node.selectSingleNode("@username").text
					
					PostBuyUser=Node.selectSingleNode("@postbuyuser").text
					ReplyID_a=Node.selectSingleNode("@announceid").text
					RootID_a=Node.selectSingleNode("@rootid").text
					AnnounceID_a=ReplyID_a
					Node.selectSingleNode("@username").text=Dvbbs.ChkBadWords(username)
					'Ubb转换
					If InStr(Ubblists,",39,") > 0  Then
						Node.selectSingleNode("@body").text = dv_ubb.Dv_UbbCode(postbody,UserGroupID,1,0)
					Else
						Node.selectSingleNode("@body").text = dv_ubb.Dv_UbbCode(postbody,UserGroupID,1,1) 
					End If
					'利用ubblist节点传送广告数据
					If Dvbbs.Forum_ads(7)="1" Then
						Node.selectSingleNode("@ubblist").text=Topic_Ads
					Else
						Node.selectSingleNode("@ubblist").text=""
					End If
					Node.selectSingleNode("@topic").text=Dvbbs.Replacehtml(Node.selectSingleNode("@topic").text)
				Next
				For Each Node In postuserlist.documentElement.SelectNodes("user")
					Rem 分解userIM数组
					UserIM=Split(Node.selectSingleNode("@userim").text,"|||")
					Node.attributes.setNamedItem(postuserlist.createNode(2,"homepage","")).text=UserIM(0)
					Node.attributes.setNamedItem(postuserlist.createNode(2,"oicq","")).text=UserIM(1)
					Node.attributes.setNamedItem(postuserlist.createNode(2,"icq","")).text=UserIM(2)
					Node.attributes.setNamedItem(postuserlist.createNode(2,"msn","")).text=UserIM(3)
					Node.attributes.setNamedItem(postuserlist.createNode(2,"aim","")).text=UserIM(4)
					Node.attributes.setNamedItem(postuserlist.createNode(2,"yahoo","")).text=UserIM(5)
					Node.selectSingleNode("@userim").text=UserIM(6)
					userface=Dv_FilterJS(Node.selectSingleNode("@userface").text)
					Node.selectSingleNode("@userface").text=userface
					If Dvbbs.forum_setting(42) = "0" Then '关闭签名的判断 2005-5-17 Dv.Yz
						Node.selectSingleNode("@usersign").text=""	
					Else
							If canusersign(Node.selectSingleNode("@usergroupid").text) = 1 Then
								Node.selectSingleNode("@usersign").text=Dvbbs.ChkBadWords(Dv_ubb.Dv_SignUbbCode(Node.selectSingleNode("@usersign").text,Node.selectSingleNode("@usergroupid").text))	
							Else
								Node.selectSingleNode("@usersign").text=""
							End If
					End If
					Node.selectSingleNode("@joindate").text=Formatdatetime(Node.selectSingleNode("@joindate").text,1)
					If Node.selectSingleNode("@userhidden").text = "2" Then
						If IsDate(Node.selectSingleNode("@lastlogin").text) Then
							If DateDiff("s",Node.selectSingleNode("@lastlogin").text,Now())>(cCur(dvbbs.Forum_Setting(8))*60) Then
								Node.selectSingleNode("@userhidden").text = "1"
							End If
						Else
								Node.selectSingleNode("@userhidden").text = "1"
						End If
					Else
						Node.selectSingleNode("@userhidden").text = "1"
					End If
					'获取自己用户组的名字样式前后标记,并生成节点
					namestyle=Split(Application(Dvbbs.CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& Node.selectSingleNode("@usergroupid").text &"']/@groupsetting").text,",")(58)
					Node.attributes.setNamedItem(postuserlist.createNode(2,"namestyle","")).text=namestyle
					'修正显示用户组光晕效果 2005.10.13 By Winder.F
					UserGroupID=Node.selectSingleNode("@usergroupid").text
					If UserGroupID < 9 Then
						Node.attributes.setNamedItem(postuserlist.createNode(2,"nameglow","")).text=nameglow(UserGroupID-1)
					Else
						Node.attributes.setNamedItem(postuserlist.createNode(2,"nameglow","")).text=Dvbbs.mainsetting(5)
					End If
				Next
				'合并数据
				XMLDom.documentElement.appendChild(postuserlist.documentElement)
			End If
		Else
			Set XMLDom=Server.CreateObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
			XMLDom.appendChild(XMLDom.createElement("post"))	
		End If
		'插入门派数据节点，避免不三不四的门派名称出现
		If Not IsObject(Application(Dvbbs.CacheName & "_GroupName")) Then Load_GroupName()
			XMLDom.documentElement.appendChild(Application(Dvbbs.CacheName & "_GroupName").documentElement.cloneNode(True))
		Else
			Set XMLDom=Server.CreateObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
			XMLDom.appendChild(XMLDom.createElement("post"))
		End If
		Rem 主贴信息节点的插入 节点名称postinfo
		Set Node = XMLDom.documentElement.appendChild(TopicInfo.documentElement.firstChild)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"pagecount","")).text=PageCount
		Node.attributes.setNamedItem(XMLDom.createNode(2,"boardtype","")).text=Dvbbs.Replacehtml(Dvbbs.Boardtype)
		Rem 传送客户端浏览器信息节点
		XMLDom.documentElement.appendChild(Dvbbs.UserSession.documentElement.lastChild.cloneNode(True))
		Rem 传送用户信息节点
		Set Node=XMLDom.documentElement.appendChild(Dvbbs.UserSession.documentElement.firstChild.cloneNode(True))
		Rem 附加一些需要用到的用户权限数据
		If CanRead Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"canread","")).text=1
		End If
		If CanReply Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"canreply","")).text=1
		End If
		If Dvbbs.Board_Setting(68)="1" Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"cananony","")).text=1
		End If
		If dvbbs.boardmaster Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"boardmaster","")).text=1
		End If
		If TrueMaster Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"truemaster","")).text=1
		End If
		If Dvbbs.GroupSetting(41)="1" Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"canreadbest","")).text=1
		End If
		If Dvbbs.GroupSetting(30) ="1" Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"canlookip","")).text=1
		End If
		If Dvbbs.VipGroupUser Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"vipgroupuser","")).text=1
		End If
		Rem GroupSetting(10) 可以编辑自己的文章 GroupSetting(11) 可以删除自己的文章 GroupSetting(12) 可以移动自己的主题 GroupSetting(13) 可以打开关闭自己的主题
		If Dvbbs.GroupSetting(10)="1" Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"caneditmypost","")).text=1
		End If
		If Dvbbs.GroupSetting(11)="1" Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"candelmypost","")).text=1
		End If
		If Dvbbs.GroupSetting(12)="1" Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"canmovemypost","")).text=1
		End If
		If Dvbbs.GroupSetting(13)="1" Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"canlockmypost","")).text=1
		End If
		Rem 传送用于快速回复的发贴表情图标数据和em心情图标数据
		If Not IsObject(Application(Dvbbs.CacheName & "_postface")) Then LoadForum_PostFace()
		XMLDom.documentElement.appendChild(Application(Dvbbs.CacheName & "_postface").documentElement.cloneNode(True))
		If Not IsObject(Application(Dvbbs.CacheName & "_emot")) Then LoadForum_emot()
		XMLDom.documentElement.appendChild(Application(Dvbbs.CacheName & "_emot").documentElement.cloneNode(True))
		Rem 相关设置节点,一个个手工添加:(
		Set Node = XMLDom.documentElement.appendChild(XMLDom.createNode(1,"setting",""))
		Node.attributes.setNamedItem(XMLDom.createNode(2,"pagesize","")).text=Dvbbs.Board_Setting(27)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"maxpostlen","")).text=Dvbbs.Board_Setting(16)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"piccheck","")).text=Dvbbs.Board_Setting(4)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"picurl","")).text=Dvbbs.Forum_PicUrl
		Node.attributes.setNamedItem(XMLDom.createNode(2,"fontsize","")).text=Dvbbs.Board_setting(28)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"postalipay","")).text=Dvbbs.Board_setting(67)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"upostalipay","")).text=Dvbbs.Forum_setting(89)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"isboke","")).text=Dvbbs.Forum_setting(99)
		Rem 如果要恢复行距设置请把下面一行屏蔽掉
		Dvbbs.Board_setting(29)="normal"
		Node.attributes.setNamedItem(XMLDom.createNode(2,"lineheight","")).text=Dvbbs.Board_setting(29)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"indent","")).text=Dvbbs.Board_setting(69)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"usertitle","")).text=Dvbbs.forum_setting(6)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"menpai","")).text=Dvbbs.forum_setting(32)
		If Dvbbs.Forum_Setting(90)="1" Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"usetools","")).text=1
		End If
		If Dvbbs.Forum_ChanSetting(0)="1" Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"useray","")).text=1
		End If
		If Dvbbs.forum_setting(2)<>"0" Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"cansendmail","")).text=1
		End If
		Rem 虚拟形象设置插入
		Node.attributes.setNamedItem(XMLDom.createNode(2,"avatarsetting","")).text=Dvbbs.Forum_Setting(82) 
		Node.attributes.setNamedItem(XMLDom.createNode(2,"avatarmode","")).text=Dvbbs.board_Setting(59)
		'插入投票节点
		Dim votexml,votearray,voteitem,votearray1,node1
		If isVote = 1 Then
			Set Rs=Nothing 
			Set Rs=Dvbbs.Execute("Select * From Dv_Vote Where VoteID="&PollID)
			If Not Rs.EOF Then
				Set votexml=Dvbbs.RecordsetToxml(rs,"voteinfo","vote")
				Set Node=votexml.documentElement.firstChild
				votearray1=split(Node.selectSingleNode("@vote").text,"|")
				votearray=split(Node.selectSingleNode("@votenum").text,"|")
				Node.selectSingleNode("@vote").text=""
				Node.selectSingleNode("@votenum").text=""
				i=0
				For Each voteitem in votearray1
					Set node1 = Node.appendChild(votexml.createNode(1,"voteitem",""))
					Node1.text=voteitem
					Node1.attributes.setNamedItem(votexml.createNode(2,"num","")).text=votearray(i)
					i=i+1
				Next
				'投票过期限制判断
				If datediff("d",Node.selectSingleNode("@timeout").text,Now()) > 0 Then
					Node.attributes.setNamedItem(votexml.createNode(2,"istimeout","")).text=1
				Else
					'检查此人是否已经投过票了
					If Dvbbs.UserID > 0  and Locktopic =0 Then
							If Not Dvbbs.Execute("Select * From Dv_voteuser Where voteid="& PollID &" And userid="& Dvbbs.userid).EOF Then
								Node.attributes.setNamedItem(votexml.createNode(2,"alreadyvote","")).text=1
							End If
					End If
				End If
				XMLDom.documentElement.appendChild(node)
			End If
		End If
		Dim xslt,proc,XMLStyle
		If IsObject(Application(Dvbbs.CacheName&"_dispbbsemplate_"&Dvbbs.SkinID)) Then
			Set XSLT=Application(Dvbbs.CacheName&"_dispbbsemplate_"&Dvbbs.SkinID)
		Else
			Set XMLStyle=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			XMLStyle.loadxml template.html(0)
			Set XSLT=Server.CreateObject("Msxml2.XSLTemplate" & MsxmlVersion)
			XSLT.stylesheet=XMLStyle
			Set Application(Dvbbs.CacheName&"_dispbbsemplate_"&Dvbbs.SkinID)=XSLT
		End If	
		Set proc = XSLT.createProcessor()
		proc.input = XMLDom
		proc.transform()
   		Response.Write  proc.output
		Set XMLDOM=Nothing
		Set XSLt=Nothing
		Set proc=Nothing
		
End Sub
Sub LoadForum_emot()
	Dim emot,emotXML,i
	emot=split(Dvbbs.Forum_emot,"|||")
	Set emotXML=Server.CreateObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	emotXML.appendChild(emotXML.createElement("emot"))
	emotXML.documentElement.attributes.setNamedItem(emotXML.createNode(2,"path","")).text=emot(0)
	For i= 1 to UBound(emot)-1
		emotXML.documentElement.appendChild(emotXML.createNode(1,"em","")).text=emot(i)
	Next
	Set Application(Dvbbs.CacheName & "_emot")=emotXML.cloneNode(True)
End Sub
Sub LoadForum_PostFace()
	Dim PostFace,PostFaceXML,i
	PostFace=split(Dvbbs.Forum_PostFace,"|||")
	Set PostFaceXML=Server.CreateObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	PostFaceXML.appendChild(PostFaceXML.createElement("postface"))
	PostFaceXML.documentElement.attributes.setNamedItem(PostFaceXML.createNode(2,"path","")).text=PostFace(0)
	For i= 1 to UBound(PostFace)-1
		PostFaceXML.documentElement.appendChild(PostFaceXML.createNode(1,"face","")).text=PostFace(i)
	Next
	Set Application(Dvbbs.CacheName & "_postface")=PostFaceXML.cloneNode(True)
End Sub
Function GetFormID()
	Dim i,sessionid
	sessionid = Session.SessionID
	For i=1 to Len(sessionid)
		GetFormID=GetFormID&Chr(Mid(sessionid,i,1)+97)
	Next
End Function
Function canusersign(GroupID)
	If Application(Dvbbs.CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& GroupID &"']/@groupsetting") Is Nothing Then GroupID=7
	canusersign = Split(Application(Dvbbs.CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& GroupID &"']/@groupsetting").text,",")(55)
End Function
%>
