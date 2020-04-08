<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/dv_ubbcode.asp"-->
<!--#include file="Dv_plus/IndivGroup/Dv_IndivGroup_Config.asp"-->
<!--#include file="Dv_plus/IndivGroup/Dv_IndivGroup_MainCls.asp"-->
<%
Dim Rs,SQL,TopicInfo
Dim TopicID,ReplyID,Star,Skin,TopicTitle,TopFlag,TopicCount,PostUserName,TopicLockFlag
Dim CanRead,TrueMaster,CanReply

If Dv_IndivGroup_MainClass.ID=0 Or Dv_IndivGroup_MainClass.Name="" Then Response.redirect "showerr.asp?ErrCodes=对不起，你访问的圈子不存在或已经被删除&action=OtherErr"
If Dv_IndivGroup_MainClass.PowerFlag>0 Then
	If Dv_IndivGroup_MainClass.PowerFlag>3 And Dv_IndivGroup_MainClass.GroupStats=3 Then Response.redirect "showerr.asp?ErrCodes=<li>圈子“"&Dv_IndivGroup_MainClass.Name&"”已关闭，只有圈子管理员才能进入。&action=OtherErr"
Else
	Response.redirect "showerr.asp?ErrCodes=<li>抱歉，圈子“"&Dv_IndivGroup_MainClass.Name&"”不公开，只有圈子成员才能进入。&action=OtherErr"
End If
If Dv_IndivGroup_MainClass.BoardID=0 Then Response.Write "错误，栏目ID为0，该帖子不能浏览。"
Dvbbs.LoadTemplates("indivgroup")

Dim replyid_a,AnnounceID_a,RootID_a
Dim PostBuyUser,abgcolor,bgcolor,UserName
Dim Page,action,n,EmotPath,dv_ubb

EmotPath=Split(Dvbbs.Forum_emot,"|||")(0)		'em心情路径

Call Chk_Topic_Err:Dvbbs.Showerr()
Dvbbs.Nav()
Dv_IndivGroup_MainClass.Head_var 1,"",""

If UserFlashGet = 1 Then
%>
<!--#include file="Dv_plus/Flashget/Flashget_base64.asp"-->
<%
	Response.Write "<script src=""http://ufile.kuaiche.com/Flashget_union.php?fg_uid="&FlashGetID&"""></script>"
End If

GetForumTextAd(0)
Dvbbs.ActiveOnline()
Page=Dvbbs.CheckNumeric(Request("Page")):If Page=0 Then Page=1

Set Dv_ubb=new Dvbbs_UbbCode
Dv_ubb.PostType=1
Show_Topic_HTML()
Set Dv_ubb=Nothing

Dvbbs.Footer()
Dvbbs.PageEnd()
Sub Chk_Topic_Err()
	TopicID=Dvbbs.CheckNumeric(Request("ID")):If TopicID=0 Then Dvbbs.AddErrCode(30):Exit Sub
	ReplyID=Dvbbs.CheckNumeric(Request("ReplyID")):If ReplyID=0 Then ReplyID=TopicID
	Star=Dvbbs.CheckNumeric(Request("Star")):If Star=0 Then Star=1
	Skin=Dvbbs.CheckNumeric(Request("Skin"))
	'浏览帖子权限
	If Dv_IndivGroup_MainClass.PowerFlag>0 And Dv_IndivGroup_MainClass.PowerFlag<4 Then TrueMaster=True	
	CanRead=False:CanReply=True

	SQL="Select GroupID,title,istop,isbest,PostUserName,PostUserid,hits,child,pollid,LockTopic,BoardID,TopicMode,DateAndTime,Expression,TopicID,topicID as ReplyID,topicID as star,topicID as page,topicID as skin From Dv_Group_Topic where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&Dv_IndivGroup_MainClass.BoardID&" And TopicID="&TopicID
	Set Rs=Dv_IndivGroup_MainClass.Execute(SQL)
	If Not(Rs.BOF and Rs.EOF) then
		If Rs(10)=Dv_IndivGroup_MainClass.BoardID Then
			Dv_IndivGroup_MainClass.Execute("Update Dv_Group_Topic Set hits=hits+1 Where GroupID="&Dv_IndivGroup_MainClass.ID&" And TopicID="&TopicID)
			TopicTitle=Rs(1)
			TopFlag=Rs(2)
			TopicCount=Rs(7)+1
			PostUserName = Rs(4)
			TopicLockFlag = Rs(9)
			TopicTitle=Replace(Dvbbs.ChkBadWords(TopicTitle),"<","&lt;")
			Dv_IndivGroup_MainClass.Stats = TopicTitle
			Set Topicinfo=Dvbbs.RecordsetToxml(Rs,"postinfo","xml")
		Else
			'Dvbbs.AddErrCode(29)
		End if
	Else
		Dvbbs.AddErrcode(32)
	End If
	Rs.Close:Set Rs=Nothing
End Sub

Function Topic_Ads(n)
	Randomize
	Topic_Ads=Dvbbs.Forum_ads(n)(CInt(UBound(Dvbbs.Forum_ads(n))*Rnd))
End Function

Sub Show_Topic_HTML()
	Dim namestyle,nameglow(7)
	Dim i,XMLDom,Node
	Dim PageCount,postarray,postuseridlist,postuserlist,UserGroupID,postbody,Topic,cmd,postbuyusers,UserIM
	Dim userface
	nameglow(0)=Dvbbs.mainsetting(9):nameglow(1)=Dvbbs.mainsetting(7):nameglow(2)=Dvbbs.mainsetting(7):nameglow(3)=Dvbbs.mainsetting(5):nameglow(6)="gray":nameglow(7)=Dvbbs.mainsetting(11)
	'是否开启帖间广告
	If Dvbbs.Forum_ads(7)="1" Then
		Dvbbs.Forum_ads(14)=replace(Dvbbs.Forum_ads(14),Chr(13),"")
		Dvbbs.Forum_ads(14)=Split(Dvbbs.Forum_ads(14),Chr(10))
	End If
	'帖子顶楼顶部广告位
	If Dvbbs.Forum_ads(18)="1" Then
		Dvbbs.Forum_ads(19)=replace(Dvbbs.Forum_ads(19),Chr(13),"")
		Dvbbs.Forum_ads(19)=Split(Dvbbs.Forum_ads(19),Chr(10))
	End If
	'帖子顶楼底部广告位
	If Dvbbs.Forum_ads(20)="1" Then
		Dvbbs.Forum_ads(21)=replace(Dvbbs.Forum_ads(21),Chr(13),"")
		Dvbbs.Forum_ads(21)=Split(Dvbbs.Forum_ads(21),Chr(10))
	End If
	'帖子顶楼左右广告位
	If Dvbbs.Forum_ads(22)<>"0" Then
		Dvbbs.Forum_ads(23)=replace(Dvbbs.Forum_ads(23),Chr(13),"")
		Dvbbs.Forum_ads(23)=Split(Dvbbs.Forum_ads(23),Chr(10))
	End If

	'插入顶帖广告节点
	Set Node = TopicInfo.documentElement.firstChild.appendChild(TopicInfo.createNode(1,"forum_ads",""))
		Node.attributes.setNamedItem(TopicInfo.createNode(2,"id","")).text = 18
		Node.attributes.setNamedItem(TopicInfo.createNode(2,"set","")).text = Dvbbs.Forum_ads(18)
	If Dvbbs.Forum_ads(18)="1" Then
		Node.attributes.setNamedItem(TopicInfo.createNode(2,"code","")).text = Topic_Ads(19)
	End If
	Set Node = TopicInfo.documentElement.firstChild.appendChild(TopicInfo.createNode(1,"forum_ads",""))
		Node.attributes.setNamedItem(TopicInfo.createNode(2,"id","")).text = 20
		Node.attributes.setNamedItem(TopicInfo.createNode(2,"set","")).text = Dvbbs.Forum_ads(20)
	If Dvbbs.Forum_ads(20)="1" Then
		Node.attributes.setNamedItem(TopicInfo.createNode(2,"code","")).text = Topic_Ads(21)
	End If
	Set Node = TopicInfo.documentElement.firstChild.appendChild(TopicInfo.createNode(1,"forum_ads",""))
		Node.attributes.setNamedItem(TopicInfo.createNode(2,"id","")).text = 22
		Node.attributes.setNamedItem(TopicInfo.createNode(2,"set","")).text = Dvbbs.Forum_ads(22)
	If Dvbbs.Forum_ads(22)<>"0" Then
		Node.attributes.setNamedItem(TopicInfo.createNode(2,"code","")).text = Topic_Ads(23)
	End If

	If TopicCount mod 10=0 then
			PageCount= TopicCount \ 10
	Else
			PageCount= TopicCount \ 10+1
	End If
	If star > PageCount Then star = PageCount
	If star < 1 Then star = 1
	If Skin=1 Then
		If Clng(replyid)=Clng(TopicID) Then
			SQL="Select Top 1 AnnounceID,GroupID,UserName,Topic,dateandtime,body,Expression,ip,RootID,signflag,isbest,PostUserid,layer,isagree,IsUpload,Ubblist,TopicLockFl,ParentID From Dv_Group_BBS where RootID="& TopicID &" and Boardid="&Dv_IndivGroup_MainClass.BoardID&" And GroupID="&Dv_IndivGroup_MainClass.ID
		Else
			SQL="Select AnnounceID,GroupID,UserName,Topic,dateandtime,body,Expression,ip,RootID,signflag,isbest,PostUserid,layer,isagree,IsUpload,Ubblist,LockTopic,ParentID From Dv_Group_BBS where AnnounceID="&replyID&" and Boardid="&Dv_IndivGroup_MainClass.BoardID&" And GroupID="&Dv_IndivGroup_MainClass.ID
		End If
	Else
		SQL="Select AnnounceID,GroupID,UserName,Topic,dateandtime,body,Expression,ip,RootID,signflag,isbest,PostUserid,layer,isagree,IsUpload,Ubblist,LockTopic,ParentID From Dv_Group_BBS where RootID="& TopicID &" and Boardid="&Dv_IndivGroup_MainClass.BoardID&" And GroupID="&Dv_IndivGroup_MainClass.ID&" Order By Announceid"
	End If
	Set Rs=Dvbbs.Execute(SQL)
	If Not Rs.EOF Then
		If Not Skin=1 Then Rs.Move(10*(star-1))
		If Not Rs.EOF Then
			TopicInfo.documentElement.firstChild.selectSingleNode("@replyid").text=Rs(0)
			TopicInfo.documentElement.firstChild.selectSingleNode("@star").text=star
			TopicInfo.documentElement.firstChild.selectSingleNode("@page").text=page
			TopicInfo.documentElement.firstChild.selectSingleNode("@skin").text=skin



			If TopicInfo.documentElement.firstChild.selectSingleNode("@topicmode").text <> "1" Then
				TopicInfo.documentElement.firstChild.selectSingleNode("@title").text=replace(TopicInfo.documentElement.firstChild.selectSingleNode("@title").text,"<","&lt;")
			End If
			postarray=Rs.GetRows(10)
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
					'PostBuyUser=Node.selectSingleNode("@postbuyuser").text
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
						Node.selectSingleNode("@ubblist").text=Topic_Ads(14)
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
			Set XMLDom=Dvbbs.CreateXmlDoc("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
			XMLDom.appendChild(XMLDom.createElement("post"))	
		End If
	Else
		Set XMLDom=Dvbbs.CreateXmlDoc("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
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
		If TrueMaster Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"truemaster","")).text=1
			Node.attributes.setNamedItem(XMLDom.createNode(2,"canlookip","")).text=1
		End If
		Node.attributes.setNamedItem(XMLDom.createNode(2,"canreadbest","")).text=1
		Rem GroupSetting(10) 可以编辑自己的文章 GroupSetting(11) 可以删除自己的文章 GroupSetting(12) 可以移动自己的主题 GroupSetting(13) 可以打开关闭自己的主题
		Node.attributes.setNamedItem(XMLDom.createNode(2,"caneditmypost","")).text=1
		Node.attributes.setNamedItem(XMLDom.createNode(2,"candelmypost","")).text=0
		Node.attributes.setNamedItem(XMLDom.createNode(2,"canmovemypost","")).text=0
		Node.attributes.setNamedItem(XMLDom.createNode(2,"canlockmypost","")).text=0
		Rem 相关设置节点,一个个手工添加:(
		Set Node = XMLDom.documentElement.appendChild(XMLDom.createNode(1,"setting",""))
		Node.attributes.setNamedItem(XMLDom.createNode(2,"pagesize","")).text=10
		Node.attributes.setNamedItem(XMLDom.createNode(2,"maxpostlen","")).text=16240
		Node.attributes.setNamedItem(XMLDom.createNode(2,"picurl","")).text=Dvbbs.Forum_PicUrl
		Node.attributes.setNamedItem(XMLDom.createNode(2,"fontsize","")).text=9

		Node.attributes.setNamedItem(XMLDom.createNode(2,"upostalipay","")).text=Dvbbs.Forum_setting(89)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"isboke","")).text=Dvbbs.Forum_setting(99)
		Rem 如果要恢复行距设置请把下面一行屏蔽掉
		Node.attributes.setNamedItem(XMLDom.createNode(2,"lineheight","")).text="normal"
		Node.attributes.setNamedItem(XMLDom.createNode(2,"indent","")).text=24
		Node.attributes.setNamedItem(XMLDom.createNode(2,"usertitle","")).text=Dvbbs.forum_setting(6)

		Dim xslt,proc,XMLStyle
		Set XMLStyle=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			XMLStyle.loadxml template.html(3)
			'XMLStyle.load Server.MapPath("Dv_plus/IndivGroup/Skin/dispbbs.xslt")
			Set XSLT=Dvbbs.iCreateObject("Msxml2.XSLTemplate" & MsxmlVersion)
			XSLT.stylesheet=XMLStyle
		Set proc = XSLT.createProcessor()
		proc.input = XMLDom
		proc.transform()
   		Response.Write  proc.output
		Set XMLDOM=Nothing
		Set XSLt=Nothing
		Set proc=Nothing
End Sub

Function canusersign(GroupID)
	If Application(Dvbbs.CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& GroupID &"']/@groupsetting") Is Nothing Then GroupID=7
	canusersign = Split(Application(Dvbbs.CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& GroupID &"']/@groupsetting").text,",")(55)
End Function
%>