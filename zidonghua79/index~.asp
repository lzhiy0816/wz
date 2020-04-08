<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<%
Rem 首页页面设置
Const CachePage=false '是否做页面缓存
Const CacheTime=60    '缓存失效时间
Dim XMLDom,page,TopicMode,Cmd

If Request("w") = "1" Then
	Passport_Main()
	Response.End
End If

If (Not Response.IsClientConnected) and Dvbbs.userid=0 Then
	Session(Dvbbs.CacheName & "UserID")=empty
	Response.Clear
	Response.End
Else
	If Request("action")="xml" Then
		Showxml()
	Else
		Main()
	End If
End If
Sub Showxml()
	Dim node
	Set XMLDOM=Application(Dvbbs.CacheName&"_ssboardlist").cloneNode(True)
	If Dvbbs.GroupSetting(37)="0"  Then'去掉隐藏论坛
		For each node in XMLDOM.documentElement.getElementsByTagName("board")
			If node.attributes.getNamedItem("hidden").text="1" Then
				node.parentNode.removeChild(node)
			End If
		Next
	End If
	Response.Clear
	Response.CharSet="gb2312"  
	Response.ContentType="text/xml"
	Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"&vbNewLine
	Response.Write XMLDom.documentElement.XML
	Response.Flush
	Set XMLDOM=Nothing
	Set Dvbbs=Nothing
	Response.End
End Sub
Sub Main()
	Dvbbs.LoadTemplates("index")
	If Dvbbs.BoardID=0 Then
		Dvbbs.Stats=Replace(template.Strings(0),"动网先锋论坛",Dvbbs.Forum_Info(0))
		Response.Write Dvbbs.mainhtml(18)
		Dvbbs.Nav()
		Dvbbs.ActiveOnline()
		GetForumTextAd(0)
		BoardList()
	Else
		Chk_List_Err()
		TopicMode=0
		If Request("topicmode")<>"" and IsNumeric(Request("topicmode")) Then TopicMode=Cint(Request("topicmode"))
		If Dvbbs.Board_Setting(43)="0" Then
			Dvbbs.Stats=Dvbbs.LanStr(7)
		Else
			Dvbbs.Stats=Dvbbs.LanStr(8)
		End If
		Response.Write Dvbbs.mainhtml(18)
		Dvbbs.Nav()
		Dvbbs.ActiveOnline()
		Dvbbs.Head_var 1,Application(Dvbbs.CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"&Dvbbs.BoardID&"']/@depth").text,"",""
		GetForumTextAd(1)
		BoardList()
		Page=Request("Page")
		If ( Not isNumeric(Page) )or Page="" Then Page=1
		Page=Clng(Page)
		If Page <1 Then Page=1
		If Dvbbs.Board_Setting(43)="0" Then
			topicList()
		End If
	End If
	Dvbbs.Footer
End Sub
Sub Chk_List_Err()
	If Dvbbs.Board_Setting(1)="1" and Dvbbs.GroupSetting(37)="0" Then
		Dvbbs.AddErrCode(26)
	ElseIf  Request("action")="batch"  and Dvbbs.GroupSetting(45)<>"1"Then
		Dvbbs.AddErrCode(28)
	End If
	Dvbbs.showerr()
End Sub
Sub topicList()
	Dim Node,modelist,modelistimg,i,cpost,ctopic
	cpost=0
	ctopic=0
	If Application(Dvbbs.CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"&Dvbbs.BoardID&"']/@child").text<>"0" Then
	For Each Node In Application(Dvbbs.CacheName&"_boardlist").documentElement.selectNodes("board[@parentid='"&Dvbbs.BoardID&"']/@boardid")
			ctopic=ctopic+CLng(Application(Dvbbs.CacheName &"_information_" & node.text).documentElement.selectSingleNode("information/@topicnum").text)
			cpost=cpost+CLng(Application(Dvbbs.CacheName &"_information_" & node.text).documentElement.selectSingleNode("information/@postnum").text)
	Next 
	End If
	Set XMLDom=Application(Dvbbs.CacheName &"_boarddata_" & Dvbbs.boardid).cloneNode(True)
	XMLDom.documentElement.firstChild.attributes.removeNamedItem("boarduser")
	XMLDom.documentElement.firstChild.attributes.removeNamedItem("board_ads")
	XMLDom.documentElement.firstChild.attributes.removeNamedItem("board_user")
	XMLDom.documentElement.firstChild.attributes.removeNamedItem("isgroupsetting")
	XMLDom.documentElement.firstChild.attributes.removeNamedItem("rootid")
	XMLDom.documentElement.firstChild.attributes.removeNamedItem("board_setting")
	XMLDom.documentElement.firstChild.attributes.removeNamedItem("sid")
	XMLDom.documentElement.firstChild.attributes.removeNamedItem("cid")
	XMLDom.documentElement.firstChild.attributes.setNamedItem(XMLDom.createNode(2,"boardtype","")).text=Dvbbs.boardtype
	XMLDom.documentElement.firstChild.attributes.setNamedItem(XMLDom.createNode(2,"forum_online","")).text=MyBoardOnline.Forum_Online
	XMLDom.documentElement.firstChild.attributes.setNamedItem(XMLDom.createNode(2,"board_useronline","")).text=MyBoardOnline.Board_UserOnline
	XMLDom.documentElement.firstChild.attributes.setNamedItem(XMLDom.createNode(2,"board_guestonline","")).text=MyBoardOnline.Board_GuestOnline
	XMLDom.documentElement.firstChild.attributes.setNamedItem(XMLDom.createNode(2,"postnum","")).text=CLng(Application(Dvbbs.CacheName &"_information_" & Dvbbs.boardid).documentElement.selectSingleNode("information/@postnum").text)-cpost
	XMLDom.documentElement.firstChild.attributes.setNamedItem(XMLDom.createNode(2,"topicnum","")).text=CLng(Application(Dvbbs.CacheName &"_information_" & Dvbbs.boardid).documentElement.selectSingleNode("information/@topicnum").text)-ctopic
	XMLDom.documentElement.firstChild.attributes.setNamedItem(XMLDom.createNode(2,"todaynum","")).text=CLng(Application(Dvbbs.CacheName &"_information_" & Dvbbs.boardid).documentElement.selectSingleNode("information/@todaynum").text)
	modelist=Split(Dvbbs.Board_Setting(48),"$$")
	modelistimg=Split(Dvbbs.Board_Setting(49),"$$")
	For i= 0 to UBound(modelist) -1
		Set Node = XMLDom.documentElement.firstChild.appendChild(XMLDom.createNode(1,"mode",""))
		Node.text=modelist(i)
		If i < UBound(modelistimg) Then Node.attributes.setNamedItem(XMLDom.createNode(2,"pic","")).text=modelistimg(i)
	Next
	XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"picurl","")).text=Dvbbs.Forum_PicUrl
	If Dvbbs.Forum_Setting(14)="1" Or Dvbbs.Forum_Setting(15)="1" Then 
		XMLDom.documentElement.firstChild.attributes.setNamedItem(XMLDom.createNode(2,"showonline","")).text="1"
	Else
		XMLDom.documentElement.firstChild.attributes.setNamedItem(XMLDom.createNode(2,"showonline","")).text="0"
	End If
	XMLDom.documentElement.appendChild(Application(Dvbbs.CacheName &"_boardmaster").documentElement.selectSingleNode("boardmaster[@boardid='"& Dvbbs.boardid&"']").cloneNode(True))
	Rem ===============传送论坛信息和设置数据到XML===============================================================
	Set Node=XMLDom.documentElement.appendChild(XMLDom.createNode(1,"forum_setting",""))
	Node.attributes.setNamedItem(XMLDom.createNode(2,"logincheckcode","")).text=Dvbbs.forum_setting(79)'登录验证码设置
	If Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(10)=1 Then 	Node.attributes.setNamedItem(XMLDom.createNode(2,"loginmobile","")).text=""'手机会员登录
	Node.attributes.setNamedItem(XMLDom.createNode(2,"rss","")).text=Dvbbs.Forum_ChanSetting(2)'rss订阅
	Node.attributes.setNamedItem(XMLDom.createNode(2,"wap","")).text=Dvbbs.Forum_ChanSetting(1)'wap访问
	Node.attributes.setNamedItem(XMLDom.createNode(2,"ishot","")).text=Dvbbs.Forum_Setting(44)'热贴最少回复
	Node.attributes.setNamedItem(XMLDom.createNode(2,"pagesize","")).text=Dvbbs.Board_Setting(26)'列表分页大小
	Node.attributes.setNamedItem(XMLDom.createNode(2,"postalipay","")).text=Dvbbs.Board_Setting(67)
	Node.attributes.setNamedItem(XMLDom.createNode(2,"dispsize","")).text=Dvbbs.Board_Setting(27) '贴子分页大小
	Node.attributes.setNamedItem(XMLDom.createNode(2,"tools","")).text=Dvbbs.Forum_Setting(90)'道具中心开关
	Node.attributes.setNamedItem(XMLDom.createNode(2,"newfalgpic","")).text=Dvbbs.Board_Setting(60) '显示新贴标志的设置
	Node.attributes.setNamedItem(XMLDom.createNode(2,"ForumUrl","")).text=Dvbbs.Get_ScriptNameUrl()
	If Dvbbs.Board_Setting(3)="1" Or Dvbbs.Board_Setting(57)="1" Then
		Node.attributes.setNamedItem(XMLDom.createNode(2,"auditcount","")).text=auditcount
	End If
	Rem 参数传递
	XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"action","")).text=Request("action")
	XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"page","")).text=Page
	XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"topicmode","")).text=topicmode
	If Dvbbs.Boardmaster Then
		XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"ismaster","")).text="1"
	Else
		XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"ismaster","")).text="0"
	End If
	If Dvbbs.Board_Setting(68)="1" Then
		XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"cananony","")).text="1"
	Else
		XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"cananony","")).text="0"
	End If
	XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"canlookuser","")).text=Dvbbs.GroupSetting(1)
	If Not IsObject(Application(Dvbbs.CacheName & "_smallpaper")) Then LoadBoardNews_Paper()
	For Each Node in Application(Dvbbs.CacheName & "_smallpaper").documentElement.SelectNodes("smallpaper[@s_boardid='"&Dvbbs.Boardid&"']")
		XMLDom.documentElement.appendChild(Node.cloneNode(True))
	Next
	LoadTopiclist()
	Response.Write vbNewLine & "<script language=""javascript"" type=""text/javascript"">" & vbNewLine
	Response.Write LoadToolsInfo & vbNewLine
	Response.Write "</script>" & vbNewLine
	If Cint(TopicMode) <> "0" Then
		XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"modecount","")).text=Dvbbs.Execute("Select  Count(*) From Dv_Topic Where Mode="&TopicMode&" and BoardID="&Dvbbs.BoardID&" And IsTop=0")(0)
	End If
	transform_topicList()
End Sub
Function auditcount()
	Dim Rs
	Set Rs=Dvbbs.Execute("select count(*) from "& Dvbbs.Nowusebbs &" where boardid=777 and locktopic="&Dvbbs.BoardID)
	If IsNull(Rs(0)) Then
		auditcount=0
	Else
		auditcount=Rs(0)
	End If
	Set Rs=Nothing	
End Function
Sub LoadTopiclist()
	If (Not Response.IsClientConnected) and Dvbbs.userid=0 Then
		Session(Dvbbs.CacheName & "UserID")=empty
		Response.Clear
		Response.End
	End If
	Dim Node,nodes,topidlist,Rs,Sql,lastpost,i,PostTime,limitime
	If Page=1 Then
		topidlist=Dvbbs.CacheData(28,0)
		If topidlist="" Then
			topidlist=Application(Dvbbs.CacheName &"_information_" & Dvbbs.boardid).documentElement.selectSingleNode("information/@boardtopstr").text
		ElseIf Trim(Application(Dvbbs.CacheName &"_information_" & Dvbbs.boardid).documentElement.selectSingleNode("information/@boardtopstr").text)<>"" Then
			topidlist=topidlist &","& Application(Dvbbs.CacheName &"_information_" & Dvbbs.boardid).documentElement.selectSingleNode("information/@boardtopstr").text
		End If
		If Trim(topidlist) <>"" Then 
			Set Rs=Dvbbs.Execute("Select topicid,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,expression,topicmode,mode,getmoney,getmoneytype,usetools,issmstopic,hidename from dv_topic Where istop > 0 and topicid in ("& Dvbbs.Checkstr(topidlist) &") Order By istop desc, Lastposttime Desc")
			If Not Rs.EOF Then
				SQL=Rs.GetRows(-1)
				Set topidlist=Dvbbs.ArrayToxml(sql,rs,"row","toptopic")
				SQL=Empty
				For Each Node in topidlist.documentElement.SelectNodes("row")
					Node.selectSingleNode("@title").text=Dvbbs.ChkBadWords(Node.selectSingleNode("@title").text)
					If Not Node.selectSingleNode("@topicmode").text ="1"  Then
						Node.selectSingleNode("@title").text=replace(Node.selectSingleNode("@title").text,"<","&lt;")
					End If
					Node.selectSingleNode("@lastpost").text=Dvbbs.ChkBadWords(Node.selectSingleNode("@lastpost").text)
					Node.selectSingleNode("@postusername").text=Dvbbs.ChkBadWords(Node.selectSingleNode("@postusername").text)
					i=0
					For each lastpost in split(Node.selectSingleNode("@lastpost").text,"$")
						Node.attributes.setNamedItem(topidlist.createNode(2,"lastpost_"& i,"")).text=lastpost
						i=i+1
					Next
					If Dvbbs.Board_Setting(60)<>"" And Dvbbs.Board_Setting(60)<>"0" Then
						If Dvbbs.Board_Setting(38) = "0" Then
								PostTime = Node.selectSingleNode("@lastpost_2").text
						Else
								PostTime = Node.selectSingleNode("@dateandtime").text
						End If
						If DateDiff("n",Posttime,Now)+Cint(Dvbbs.Forum_Setting(0)) < CLng(Dvbbs.Board_Setting(61)) Then
							Node.attributes.setNamedItem(topidlist.createNode(2,"datedifftime","")).text=DateDiff("n",Posttime,Now)+Cint(Dvbbs.Forum_Setting(0))
						End If
					End If
				Next
				XMLDom.documentElement.appendChild(topidlist.documentElement)
				End If
				Set Rs=Nothing
		End If
	End If
	
	If IsSqlDataBase=1 And IsBuss=1 Then
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
		Cmd("@pagesize")=Cint(Dvbbs.Board_Setting(26))
		Cmd("@topicmode")=TopicMode
		Cmd("@tl")=0
		Set Rs=Cmd.Execute
		If Not Rs.EoF Then
			SQL=Rs.GetRows(-1)
			Set topidlist=Dvbbs.ArrayToxml(sql,rs,"row","topic")
		Else
			Set topidlist=Nothing
		End If
	Else
		Set Rs = Server.CreateObject ("adodb.recordset")
		If Cint(TopicMode)=0 Then
			Sql="Select  TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,TopicMode,Mode,GetMoney,GetMoneyType,UseTools,IsSmsTopic,hidename From Dv_Topic Where BoardID="&Dvbbs.BoardID&" And IsTop=0 Order By LastPostTime Desc"
		Else
			Sql="Select  TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,TopicMode,Mode,GetMoney,GetMoneyType,UseTools,IsSmsTopic,hidename From Dv_Topic Where Mode="&TopicMode&" and BoardID="&Dvbbs.BoardID&" And IsTop=0  Order By LastPostTime Desc"
		End If
		Rs.Open Sql,Conn,1,1
		If Page >1 Then
			Rs.Move (page-1) * Clng(Dvbbs.Board_Setting(26))
		End If
		If Not Rs.EoF Then
			SQL=Rs.GetRows(Dvbbs.Board_Setting(26))
			Set topidlist=Dvbbs.ArrayToxml(sql,rs,"row","topic")
		Else
			Set topidlist=Nothing
		End If
	End If
	SQL=Empty
	If Not topidlist Is Nothing Then 
		For Each Node in topidlist.documentElement.SelectNodes("row")
				Node.selectSingleNode("@title").text=Dvbbs.ChkBadWords(Node.selectSingleNode("@title").text)
				If Not Node.selectSingleNode("@topicmode").text ="1"  Then
						Node.selectSingleNode("@title").text=replace(Node.selectSingleNode("@title").text,"<","&lt;")
				End If
				Node.selectSingleNode("@postusername").text=Dvbbs.ChkBadWords(Node.selectSingleNode("@postusername").text)
				i=0
				For each lastpost in split(Node.selectSingleNode("@lastpost").text,"$")
					Node.attributes.setNamedItem(topidlist.createNode(2,"lastpost_"& i,"")).text=lastpost
					i=i+1
				Next
				If Dvbbs.Board_Setting(60)<>"" And Dvbbs.Board_Setting(60)<>"0" Then
					If Dvbbs.Board_Setting(38) = "0" Then
							PostTime = Node.selectSingleNode("@lastpost_2").text
					Else
							PostTime = Node.selectSingleNode("@dateandtime").text
					End If
					If DateDiff("n",Posttime,Now)+Cint(Dvbbs.Forum_Setting(0)) < CLng(Dvbbs.Board_Setting(61)) Then
						Node.attributes.setNamedItem(topidlist.createNode(2,"datedifftime","")).text=DateDiff("n",Posttime,Now)+Cint(Dvbbs.Forum_Setting(0))
					End If
				End If
			Next
		XMLDom.documentElement.appendChild(topidlist.documentElement)
	End If
		Set Rs=Nothing
	Dvbbs.SqlQueryNum = Dvbbs.SqlQueryNum + 1
End Sub
Sub transform_topicList()
	If (Not Response.IsClientConnected) and Dvbbs.userid=0 Then
		Response.Clear
		Session(Dvbbs.CacheName & "UserID")=empty
		Response.End
	End If
	Dim proc,XMLStyle,node,cnode
	If Not IsObject(Application(Dvbbs.CacheName & "_listtemplate_"& Dvbbs.SkinID)) Then
		Set Application(Dvbbs.CacheName & "_listtemplate_"& Dvbbs.SkinID)=Server.CreateObject("Msxml2.XSLTemplate" & MsxmlVersion )
		Set XMLStyle=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion )
		XMLStyle.loadxml template.html(1) ' 
		'XMLStyle.load Server.MapPath("index_list.xslt")
		'插入各种图片的设置数据
		Set Node=XMLStyle.createNode(1,"xsl:variable","http://www.w3.org/1999/XSL/Transform")
		Set CNode=XMLStyle.createNode(2,"name","")
		CNode.text="picurl"
		Node.attributes.setNamedItem(CNode)
		node.text=Dvbbs.Forum_PicUrl
		XMLStyle.documentElement.appendChild(node)
		Set Node=XMLStyle.createNode(1,"xsl:variable","http://www.w3.org/1999/XSL/Transform")
		Set CNode=XMLStyle.createNode(2,"name","")
		CNode.text="pic_nofollow"
		Node.attributes.setNamedItem(CNode)
		node.text=Dvbbs.mainpic(10)
		XMLStyle.documentElement.appendChild(node)
		Set Node=XMLStyle.createNode(1,"xsl:variable","http://www.w3.org/1999/XSL/Transform")
		Set CNode=XMLStyle.createNode(2,"name","")
		CNode.text="pic_follow"
		Node.attributes.setNamedItem(CNode)
		node.text=Dvbbs.mainpic(11)
		XMLStyle.documentElement.appendChild(node)
		Set Node=XMLStyle.createNode(1,"xsl:variable","http://www.w3.org/1999/XSL/Transform")
		Set CNode=XMLStyle.createNode(2,"name","")
		CNode.text="ztopic"
		Node.attributes.setNamedItem(CNode)
		node.text=Dvbbs.mainpic(0)
		XMLStyle.documentElement.appendChild(node)
		Set Node=XMLStyle.createNode(1,"xsl:variable","http://www.w3.org/1999/XSL/Transform")
		Set CNode=XMLStyle.createNode(2,"name","")
		CNode.text="istopic"
		Node.attributes.setNamedItem(CNode)
		node.text=Dvbbs.mainpic(1)
		XMLStyle.documentElement.appendChild(node)
		Set Node=XMLStyle.createNode(1,"xsl:variable","http://www.w3.org/1999/XSL/Transform")
		Set CNode=XMLStyle.createNode(2,"name","")
		CNode.text="opentopic"
		Node.attributes.setNamedItem(CNode)
		node.text=Dvbbs.mainpic(2)
		XMLStyle.documentElement.appendChild(node)
		Set Node=XMLStyle.createNode(1,"xsl:variable","http://www.w3.org/1999/XSL/Transform")
		Set CNode=XMLStyle.createNode(2,"name","")
		CNode.text="hottopic"
		Node.attributes.setNamedItem(CNode)
		node.text=Dvbbs.mainpic(3)
		XMLStyle.documentElement.appendChild(node)
		Set Node=XMLStyle.createNode(1,"xsl:variable","http://www.w3.org/1999/XSL/Transform")
		Set CNode=XMLStyle.createNode(2,"name","")
		CNode.text="ilocktopic"
		Node.attributes.setNamedItem(CNode)
		node.text=Dvbbs.mainpic(4)
		XMLStyle.documentElement.appendChild(node)
		Set Node=XMLStyle.createNode(1,"xsl:variable","http://www.w3.org/1999/XSL/Transform")
		Set CNode=XMLStyle.createNode(2,"name","")
		CNode.text="besttopic"
		Node.attributes.setNamedItem(CNode)
		node.text=Dvbbs.mainpic(5)
		XMLStyle.documentElement.appendChild(node)
		Set Node=XMLStyle.createNode(1,"xsl:variable","http://www.w3.org/1999/XSL/Transform")
		Set CNode=XMLStyle.createNode(2,"name","")
		CNode.text="votetopic"
		Node.attributes.setNamedItem(CNode)
		node.text=Dvbbs.mainpic(6)
		XMLStyle.documentElement.appendChild(node)
		Set Node=XMLStyle.createNode(1,"xsl:variable","http://www.w3.org/1999/XSL/Transform")
		Set CNode=XMLStyle.createNode(2,"name","")
		CNode.text="pic_toptopic1"
		Node.attributes.setNamedItem(CNode)
		node.text=Dvbbs.mainpic(19)
		XMLStyle.documentElement.appendChild(node)
		Application(Dvbbs.CacheName & "_listtemplate_"& Dvbbs.SkinID).stylesheet=XMLStyle
	End If
	Set proc = Application(Dvbbs.CacheName & "_listtemplate_"& Dvbbs.SkinID).createProcessor()
	proc.input = XMLDom
	proc.transform()
	Response.Write  proc.output
	Set XMLDom=Nothing 
	Set proc=Nothing
End Sub
Sub LoadBoardlistData()
	Dim Node,Xpath,LastPost,BoardiD,Xpath1
	Set XMLDom=Application(Dvbbs.CacheName&"_boardlist").cloneNode(True)
	XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"boardid","")).text=Dvbbs.BoardID
	If Dvbbs.Boardid=0 Then
		Xpath="board[@depth=1]"
		Xpath1="board[@depth=0]"
		XMLDom.documentElement.appendChild(Application(Dvbbs.CacheName &"_grouppic").documentElement.cloneNode(True))
		If Not IsObject(Application(Dvbbs.CacheName & "_link")) Then LoadlinkList()
		XMLDom.documentElement.appendChild(Application(Dvbbs.CacheName & "_link").documentElement.cloneNode(True))
		Rem ===============传送论坛信息和设置数据到XML===============================================================
		Set Node=XMLDom.documentElement.appendChild(XMLDom.createNode(1,"forum_info",""))
		Node.attributes.setNamedItem(XMLDom.createNode(2,"forum_type","")).text=Dvbbs.forum_info(0)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"forum_maxonline","")).text=Dvbbs.CacheData(5,0)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"forum_maxonlinedate","")).text=Dvbbs.CacheData(6,0)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"forum_topicnum","")).text=Dvbbs.CacheData(7,0)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"forum_postnum","")).text=Dvbbs.CacheData(8,0)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"forum_todaynum","")).text=Dvbbs.CacheData(9,0)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"forum_usernum","")).text=Dvbbs.CacheData(10,0)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"forum_yesterdaynum","")).text=Dvbbs.CacheData(11,0)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"forum_maxpostnum","")).text=Dvbbs.CacheData(12,0)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"forum_maxpostdate","")).text=Dvbbs.CacheData(13,0)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"forum_lastuser","")).text=Dvbbs.CacheData(14,0)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"forum_online","")).text=MyBoardOnline.Forum_Online
		Node.attributes.setNamedItem(XMLDom.createNode(2,"forum_useronline","")).text=MyBoardOnline.Forum_UserOnline
		Node.attributes.setNamedItem(XMLDom.createNode(2,"forum_guestonline","")).text=MyBoardOnline.Forum_GuestOnline
		Node.attributes.setNamedItem(XMLDom.createNode(2,"forum_createtime","")).text=FormatDateTime(Dvbbs.Forum_Setting(74),1)
		Set Node=XMLDom.documentElement.appendChild(XMLDom.createNode(1,"forum_setting",""))
		Node.attributes.setNamedItem(XMLDom.createNode(2,"logincheckcode","")).text=Dvbbs.forum_setting(79)'登录验证码设置
		If Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(10)=1 Then 	Node.attributes.setNamedItem(XMLDom.createNode(2,"loginmobile","")).text=""'手机会员登录
		Node.attributes.setNamedItem(XMLDom.createNode(2,"rss","")).text=Dvbbs.Forum_ChanSetting(2)'rss订阅
		Node.attributes.setNamedItem(XMLDom.createNode(2,"wap","")).text=Dvbbs.Forum_ChanSetting(1)'wap访问
		Node.attributes.setNamedItem(XMLDom.createNode(2,"pic_0","")).text=template.pic(0)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"pic_1","")).text=template.pic(1)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"pic_2","")).text=template.pic(2)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"pic_3","")).text=template.pic(3)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"issearch_a","")).text=0
		Node.attributes.setNamedItem(XMLDom.createNode(2,"ForumUrl","")).text=Dvbbs.Get_ScriptNameUrl()
		If Dvbbs.Forum_setting(29)="1" Then
			If Not IsObject(Application(Dvbbs.CacheName & "_biruser")) Then
				Forum_BirUser()
			ElseIf Application(Dvbbs.CacheName & "_biruser").documentElement.selectSingleNode("@date").text <> CStr(Date()) Then
				Forum_BirUser()
			End If
			XMLDom.documentElement.appendChild(Application(Dvbbs.CacheName &"_biruser").documentElement.cloneNode(True))
		End If
		Rem ========================================================================================================================================
	Else
		Xpath="board[@parentid="&Dvbbs.BoardID&" and @depth="& CLng(XMLDom.documentElement.selectSingleNode("board[@boardid="& Dvbbs.boardid &"]/@depth").text)+1&"]"
		Xpath1="board[@boardid="& Dvbbs.Boardid&"]"
	End If
	If Dvbbs.BoardID<>0 Then
		Set Node=XMLDom.documentElement.appendChild(XMLDom.createNode(1,"forum_setting",""))
		Node.attributes.setNamedItem(XMLDom.createNode(2,"pic_0","")).text=template.pic(0)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"pic_1","")).text=template.pic(1)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"pic_2","")).text=template.pic(2)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"pic_3","")).text=template.pic(3)
		Node.attributes.setNamedItem(XMLDom.createNode(2,"issearch_a","")).text=1
	End If
	For Each Node In XMLDom.documentElement.selectNodes(Xpath)
		BoardId=Node.selectSingleNode("@boardid").text
		If Not IsObject(Application(Dvbbs.CacheName &"_information_" & BoardID) ) Then Dvbbs.LoadBoardinformation BoardID
		LastPost=Node.appendChild(Application(Dvbbs.CacheName &"_information_" & BoardID).documentElement.firstChild.cloneNode(True)).selectSingleNode("@lastpost_2").text
		If Not IsDate(LastPost) Then LastPost=Now()
		If DateDiff("h",Dvbbs.Lastlogin,LastPost)=0 Then Node.attributes.setNamedItem(XMLDom.createNode(2,"newpost","")).text="1"
		XMLDom.documentElement.appendChild(Application(Dvbbs.CacheName &"_boardmaster").documentElement.selectSingleNode("boardmaster[@boardid='"& boardid &"']").cloneNode(True))
	Next
	XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"picurl","")).text=Dvbbs.Forum_PicUrl
	XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"lastupdate","")).text=Now()
	If CachePage Then 
		Set Application(Dvbbs.CacheName & "_Pagecache_index_" & Dvbbs.BoardID)=XMLDOM.cloneNode(True)
	End If 
End Sub
Sub BoardList()
	If Dvbbs.BoardID=0 Then
		ShowNews()
	ElseIf Application(Dvbbs.CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid="&dvbbs.boardid&"]/@nopost").text<>"1" Then 
		ShowNews()
	End If
	Dim Node,ShowMod,Xpath1,BoardId
	If CachePage Then
		If Not IsObject(Application(Dvbbs.CacheName & "_Pagecache_index_" & Dvbbs.BoardID)) Then
			LoadBoardlistData()
		Else
			If DateDiff("s",Application(Dvbbs.CacheName & "_Pagecache_index_" & Dvbbs.BoardID).documentElement.selectSingleNode("@lastupdate").text,Now()) > CacheTime Then 
				LoadBoardlistData()
			Else
				Set XmlDom=Application(Dvbbs.CacheName & "_Pagecache_index_" & Dvbbs.BoardID).cloneNode(True)
			End If
		End If	
	Else
		LoadBoardlistData()
	End If
	If Dvbbs.GroupSetting(37)="0"  Then
		For each node in XMLDOM.documentElement.selectNodes("board[@hidden=1]")
				XMLDom.documentElement.removeChild(node)
		Next
	End If
	If Dvbbs.BoardID=0 Then
		Xpath1="board[@depth=0]"
	Else
		Xpath1="board[@boardid="& Dvbbs.Boardid&"]"
	End If
		Set Node=XMLDom.documentElement.selectSingleNode("forum_setting")
		If Dvbbs.IsSearch Then
			Node.attributes.setNamedItem(XMLDom.createNode(2,"issearch","")).text=1
		Else
			Node.attributes.setNamedItem(XMLDom.createNode(2,"issearch","")).text=0
		End If
	For Each Node In XMLDom.documentElement.selectNodes(Xpath1)
		BoardId=Node.selectSingleNode("@boardid").text
		ShowMod=Request.Cookies("List")("list"&BoardId)
		If ShowMod<>"" And IsNumeric(ShowMod) Then
			Node.selectSingleNode("@mode").text=ShowMod
		End If		
	Next
	If Dvbbs.BoardID=0  Then
		XMLDom.documentElement.appendChild(Dvbbs.UserSession.documentElement.firstChild.cloneNode(True))
		XMLDom.documentElement.appendChild(Dvbbs.UserSession.documentElement.lastChild.cloneNode(True))
		If Dvbbs.UserID <>0 Then
			'身份切换数据节点
			If UBound(Dvbbs.UserGroupParentID) <> -1 Then
				For Each Node In Dvbbs.UserGroupParentID
					XMLDom.documentElement.appendChild(XMLDom.createNode(1,"myusergroup","")).text=Node
				Next
				ElseIf Dvbbs.IsUserPermissionOnly = 1 Then
				XMLDom.documentElement.appendChild(XMLDom.createNode(1,"myusergroup","")).text=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usergroupid2").text
			End If
		End If
	End If
	If Dvbbs.Forum_ads(2)="1" or Dvbbs.Forum_ads(13)="1" Then Response.Write "<script language=""javascript"" src=""inc/Dv_Adv.js"" type=""text/javascript""></script>"
	transform_BoardList()
	If Dvbbs.Boardid=0 Then 
		If Dvbbs.Forum_Setting(14)="1" Or Dvbbs.Forum_Setting(15)="1" Then 
		Response.Write "<iframe style=""border:0px;width:0px;height:0px;""  src=""Online.asp?action=1&amp;Boardid=0"" name=""hiddenframe""></iframe>"
		Else
			Response.Write "<iframe style=""border:0px;width:0px;height:0px;"" src="""" name=""hiddenframe""></iframe>"
		End If
	End If
	If Dvbbs.Forum_ads(2)="1" or Dvbbs.Forum_ads(13)="1" Then
		Response.Write "<script language=""javascript"" type=""text/javascript"">" & vbNewLine
		If Dvbbs.Forum_ads(2)="1" Then Response.Write Chr(9) & "move_ad('"&Dvbbs.Forum_ads(3)&"','"&Dvbbs.Forum_ads(4)&"','"&Dvbbs.Forum_ads(5)&"','"&Dvbbs.Forum_ads(6)&"');" & vbNewLine
	If Dvbbs.Forum_ads(13)="1" Then Response.Write Chr(9) & "fix_up_ad('"& Dvbbs.Forum_ads(8) & "','" & Dvbbs.Forum_ads(10) & "','" & Dvbbs.Forum_ads(11) & "','" & Dvbbs.Forum_ads(9) & "');"& vbNewLine
		Response.Write vbNewLine&"</script>"
	End If	
End Sub
Sub transform_BoardList()
	Dim proc,XMLStyle
	If (Not Response.IsClientConnected) and Dvbbs.userid=0 Then
		Response.Clear
		Session(Dvbbs.CacheName & "UserID")=empty
		Response.End
	Else
		If Not IsObject(Application(Dvbbs.CacheName & "_indextemplate_"& Dvbbs.SkinID)) Then
			Set Application(Dvbbs.CacheName & "_indextemplate_"& Dvbbs.SkinID)=Server.CreateObject("Msxml2.XSLTemplate" & MsxmlVersion)
			Set XMLStyle=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			XMLStyle.loadxml template.html(0) ' Server.MapPath("index.xslt")
			Application(Dvbbs.CacheName & "_indextemplate_"& Dvbbs.SkinID).stylesheet=XMLStyle
		End If
		Set proc = Application(Dvbbs.CacheName & "_indextemplate_"& Dvbbs.SkinID).createProcessor()
		proc.input = XMLDom
		proc.transform()
		Response.Write  proc.output
		Set XMLDom=Nothing 
		Set proc=Nothing
	End If
End Sub
Sub ShowNews()
	Dim Rs,proc,NewsDom,XMLStyle
	If Not IsObject(Application(Dvbbs.CacheName & "_News")) Then
		Set Rs=Dvbbs.Execute("Select boardid,title,addtime,bgs From Dv_bbsnews order by id desc")
		Set Application(Dvbbs.CacheName & "_News")=Dvbbs.RecordsetToxml(rs,"news","")
	End If
	Set NewsDom=Application(Dvbbs.CacheName & "_News").cloneNode(True)
	NewsDom.documentElement.attributes.setNamedItem(NewsDom.createNode(2,"boardid","")).text=Dvbbs.BoardID
	If not IsObject(Application(Dvbbs.CacheName & "_shownews_"&Dvbbs.SkinID)) Then
		Set Application(Dvbbs.CacheName & "_shownews_"&Dvbbs.SkinID)=Server.CreateObject("Msxml2.XSLTemplate" & MsxmlVersion)
		Set XMLStyle=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		If  UBound(template.html)>3 Then
			XMLStyle.loadxml template.html(3)
		Else
			XMLStyle.load Server.MapPath(MyDbPath &"inc\Templates\Dv_News.xslt")
		End If
		Application(Dvbbs.CacheName & "_shownews_"&Dvbbs.SkinID).stylesheet=XMLStyle
	End If
	Set proc = Application(Dvbbs.CacheName & "_shownews_"&Dvbbs.SkinID).createProcessor()
	proc.input = NewsDom
	proc.transform()
	Response.Write  proc.output
	Set NewsDom=Nothing 
	Set proc=Nothing
End Sub
Sub LoadlinkList()
	Dim rs
	Set Rs=Dvbbs.Execute("select * From Dv_bbslink Order by islogo desc,id ")
	Set Application(Dvbbs.CacheName & "_link")=Dvbbs.RecordsetToxml(rs,"link","bbslink")
	Set Rs=Nothing
End Sub
Sub Forum_BirUser()
	Dim Rs,SQL,NowMonth,NowDate,todaystr0,todaystr1,node
	NowMonth=Month(Date())
	NowDate=Day(Date())
	If NowMonth< 10 Then
		todaystr0="0"&NowMonth
	Else
		todaystr0=CStr(NowMonth)
	End If
	If NowDate < 10 Then
		todaystr0=todaystr0&"-"&"0"&NowDate
	Else
		todaystr0=todaystr0&"-"&NowDate
	End If
	todaystr1=NowMonth&"-"&NowDate
	If todaystr0=todaystr1 Then
		SQL="select username,Userbirthday from [Dv_user] where Userbirthday like '%"&todaystr1&"' Order by UserID"
	Else
		SQL="select username,Userbirthday from [Dv_user] where Userbirthday like '%"&todaystr1&"' Or Userbirthday like '%"&todaystr0&"' Order by UserID"
	End If
	Set Rs=Dvbbs.Execute(SQL)
	Set Application(Dvbbs.CacheName & "_biruser")=Dvbbs.RecordsetToxml(rs,"user","biruser")
	Set Rs=Nothing
	For Each node In Application(Dvbbs.CacheName & "_biruser").documentElement.selectNodes("user")
		todaystr0=Node.selectSingleNode("@userbirthday").text
		If IsDate(todaystr0) Then
			Node.attributes.setNamedItem(Application(Dvbbs.CacheName & "_biruser").createNode(2,"age","")).text=datediff("yyyy",todaystr0,Now())
		Else
			Application(Dvbbs.CacheName & "_biruser").documentElement.removeChild(node)
		End If
	Next
	Application(Dvbbs.CacheName & "_biruser").documentElement.attributes.setNamedItem(Application(Dvbbs.CacheName & "_biruser").createNode(2,"date","")).text=Date()
End Sub
Function LoadToolsInfo()
	Dim Tools_Info,i,ShowTools,TempStr
	Dvbbs.Name="Plus_ToolsInfo"
	If Dvbbs.ObjIsEmpty() Then
		Dim Rs,Sql
		Sql = "Select ID,ToolsName From Dv_Plus_Tools_Info order by ID"
		Set Rs = Dvbbs.Plus_Execute(Sql)
		If Not Rs.Eof Then
			Sql = Rs.GetString(,, "§§§", "@#@", "")
		End If
		Rs.Close : Set Rs = Nothing
		Tools_Info = Split(Sql,"@#@")
		TempStr =  "var ShowTools = new Array();" & vbNewLine
		For i=0 To Ubound(Tools_Info)-1
			ShowTools = Split(Tools_Info(i),"§§§")
			TempStr = TempStr & "ShowTools["&ShowTools(0)&"]='"&Replace(Replace(Replace(ShowTools(1),"\","\\"),"'","\'"),chr(13),"")&"';"
		Next
		Dvbbs.value = TempStr & vbNewLine
	End If
	LoadToolsInfo = Dvbbs.value
End Function

Sub Passport_Main()
	Dim UserID,ForumID,token,t,ForumMsg,toUrl,Passport
	UserID = Request("uid")
	ForumID = Request("fid")
	token = Request("token")
	Passport = Request("passport")
	t = Request("t")
	If UserID = "" Or Not IsNumeric(UserID) Then UserID = 0
	UserID = cCur(UserID)
	If ForumID = "" Or Not IsNumeric(ForumID) Then ForumID = 0
	ForumID = cCur(ForumID)
	If t = "" Or Not IsNumeric(t) Then t = 1
	t = cCur(t)
	If UserID = 0 Or ForumID = 0 Or token = "" Or Passport = "" Then
		Response.Write "非法的参数！"
		Response.End
	End If
	Dim iForumUrl
	Select Case t
	Case "1"
		ForumMsg = "<li>您成功的注册了论坛通行证帐号，请牢记您填写的通行证帐号和密码。"
		toUrl = "reg.asp?action=redir"
	Case "2"
		ForumMsg = "<li>login suc。"
		toUrl = "login.asp?action=redir"
	Case Else
		ForumMsg = "<li>您成功的注册了论坛通行证帐号，请牢记您填写的通行证帐号和密码。"
		toUrl = "index.asp"
	End Select
	iForumUrl = toUrl & "&ErrorCode=1&ErrorMsg="&ForumMsg&"&passport="&Passport&"&token="&token
%>
<html>
<head>
<!--禁止被框架-->
<script type="text/javascript" language="JavaScript">
<!--
if (top.location !== self.location) {
	top.location = "index.asp?w=1&t=<%=t%>&uid=<%=UserID%>&fid=<%=ForumID%>&passport=<%=Passport%>&token=<%=token%>";
}
-->
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>欢迎访问<%=Dvbbs.Forum_Info(0)%></title>
</head>
<frameset border=0 rows=*,79 frameborder=0 framespacing=0>
<frame longdesc="" src="<%=iForumUrl%>" name="MainWin" noresize frameborder="0" marginwidth=0  marginheight=0 scrolling="auto">
<frame longdesc="" src="http://www.dvbbs.net/passport/index.asp?uid=<%=UserID%>&fid=<%=ForumID%>&token=<%=token%>&t=<%=t%>&s=1" name="top"  noresize frameborder="0" marginwidth=0  marginheight=0 scrolling="no">
</frameset>

<noframes>
<a href="http://www.dvbbs.net" target="_top">动网论坛_国内最大的免费论坛软件服务提供</a> 版权所有 2005  
此 html 框架集显示多个 web 页。若要查看此框架集，请使用支持 html 4.0 及更高版本的 web 浏览器。
</noframes>
</html>
<%
End Sub
%>