<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="Dv_plus/IndivGroup/Dv_IndivGroup_Config.asp"-->
<!--#include file="Dv_plus/IndivGroup/Dv_IndivGroup_MainCls.asp"-->
<%
If (Not Response.IsClientConnected) and Dvbbs.userid=0 Then
	Session(Dvbbs.CacheName & "UserID")=empty
	Response.Clear
	Response.End
End If
Dim Rs,SQL
Dim XMLDom,Node,XSLTemplate,XMLStyle,proc

Dvbbs.LoadTemplates("indivgroup")

If Dv_IndivGroup_MainClass.ID=0 Or Dv_IndivGroup_MainClass.Name="" Then
	Response.redirect "showerr.asp?ErrCodes=<li>对不起，你访问的圈子不存在或已经被删除</li>&action=OtherErr"
End If
If Dv_IndivGroup_MainClass.PowerFlag>0 Then
	If Dv_IndivGroup_MainClass.PowerFlag>3 And Dv_IndivGroup_MainClass.GroupStats=3 Then Response.redirect "showerr.asp?ErrCodes=<li>圈子“"&Dv_IndivGroup_MainClass.Name&"”已关闭，只有圈子管理员才能进入。</li>&action=OtherErr"
Else
	Response.redirect "showerr.asp?ErrCodes=<li>抱歉，圈子“"&Dv_IndivGroup_MainClass.Name&"”不公开，只有圈子成员才能进入。&action=OtherErr"
End If

If LCase(Request("action"))="xml" Then
	Showxml()
Else
	Main()
End If
Dvbbs.PageEnd()
Sub Showxml()
	Dim node
	Set XMLDOM=Dv_IndivGroup_MainClass.BoardXMLDom.cloneNode(True)
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
	Dvbbs.Nav()
	Dvbbs.ActiveOnline()
	Dv_IndivGroup_MainClass.LoadGroupBoard
	If Dv_IndivGroup_MainClass.BoardID=0 Then
		Dv_IndivGroup_MainClass.Stats = "栏目列表"
		Dv_IndivGroup_MainClass.Head_var 2,"",""
		GetForumTextAd(0)
		ShowBoardList()
	Else
		If Dv_IndivGroup_MainClass.BoardName="" Then Response.redirect "showerr.asp?ErrCodes=<li>栏目参数错误，该栏目可能已经被删除。"&Dv_IndivGroup_MainClass.BoardID&Dv_IndivGroup_MainClass.BoardName&"&action=OtherErr"
		Dv_IndivGroup_MainClass.Stats = "评论列表"
		Dv_IndivGroup_MainClass.Head_var 1,"",""
		GetForumTextAd(0)
		ShowTopicList()
	End If
	Dvbbs.Footer
End Sub

Sub ShowBoardList()
	Dim i,ShowMod,BoardID,LastPost,LastPostDate
	Set XMLDom=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	XMLDom.appendChild(XMLDom.createElement("IndivGroup"))
	
	XMLDom.documentElement.appendChild(Dv_IndivGroup_MainClass.BoardXMLDom.documentElement.cloneNode(True))
	XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"groupid","")).text=Dv_IndivGroup_MainClass.ID
	XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"groupname","")).text=Dv_IndivGroup_MainClass.Name
	XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"powerflag","")).text=Dv_IndivGroup_MainClass.PowerFlag

	Set Node=XMLDom.documentElement.appendChild(XMLDom.createNode(1,"forum_setting",""))
	Node.attributes.setNamedItem(XMLDom.createNode(2,"pic_0","")).text=template.pic(0)
	Node.attributes.setNamedItem(XMLDom.createNode(2,"pic_1","")).text=template.pic(1)
	Node.attributes.setNamedItem(XMLDom.createNode(2,"pic_2","")).text=template.pic(2)

	For Each Node In XMLDom.documentElement.selectNodes("/IndivGroup/BoardList/Board")
		LastPost=Node.selectSingleNode("@lastpost").text
		If IsNull(LastPost) Then
			LastPost=""
		Else
			LastPost=Split(LastPost,"$")
		End if
		For i=0 To UBound(LastPost)
			Node.attributes.setNamedItem(XMLDom.createNode(2,"lastpost_"&i,"")).text=LastPost(i)
		Next
		If UBound(LastPost)<2 Then
			LastPostDate=Now()
		Else
			If Not IsDate(LastPost(2)) Then LastPost(2)=Now()
			LastPostDate=LastPost(2)
		End if
		If DateDiff("h",Dvbbs.Lastlogin,LastPostDate)=0 Then Node.attributes.setNamedItem(XMLDom.createNode(2,"newpost","")).text="1"
	Next

	XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"picurl","")).text=Dvbbs.Forum_PicUrl
	XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"lastupdate","")).text=Now()
	XMLDom.documentElement.appendChild(Dv_IndivGroup_MainClass.MasterXMLDom.documentElement.cloneNode(True))

	If Dvbbs.Forum_ads(2)="1" or Dvbbs.Forum_ads(13)="1" Then Response.Write "<script language=""javascript"" src=""inc/Dv_Adv.js"" type=""text/javascript""></script>"

	Set XSLTemplate=Dvbbs.iCreateObject("Msxml2.XSLTemplate" & MsxmlVersion)
	Set XMLStyle=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	XMLStyle.loadxml template.html(1)
	'XMLStyle.load Server.MapPath("IndivGroup/Skin/Index.xslt")
	XSLTemplate.stylesheet=XMLStyle
	Set proc = XSLTemplate.createProcessor()
	proc.input = XMLDom
	proc.transform()
	Response.Write  proc.output
	Set XMLDom=Nothing 
	Set proc=Nothing

	If Dvbbs.Forum_ads(2)="1" or Dvbbs.Forum_ads(13)="1" Then
		Response.Write "<script language=""javascript"" type=""text/javascript"">" & vbNewLine
		If Dvbbs.Forum_ads(2)="1" Then Response.Write Chr(9) & "move_ad('"&Dvbbs.Forum_ads(3)&"','"&Dvbbs.Forum_ads(4)&"','"&Dvbbs.Forum_ads(5)&"','"&Dvbbs.Forum_ads(6)&"');" & vbNewLine
		If Dvbbs.Forum_ads(13)="1" Then Response.Write Chr(9) & "fix_up_ad('"& Dvbbs.Forum_ads(8) & "','" & Dvbbs.Forum_ads(10) & "','" & Dvbbs.Forum_ads(11) & "','" & Dvbbs.Forum_ads(9) & "');"& vbNewLine
		Response.Write vbNewLine&"</script>"
	End If	
End Sub

Sub ShowTopicList()
	Dim modelist,modelistimg,i,cpost,ctopic
	Dim nodes,topiclist,lastpost,PostTime,limitime
	Dim cnode
	Const IndivGroup_EachPageTopicCount=20 '每页显示的评论主题行数
	Const IndivGroup_EachPagePostCount=10 '每页显示的评论个数
	Set XMLDom=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	XMLDom.appendChild(XMLDom.createElement("IndivGroup"))

	'插入栏目信息
	Set Node=Dv_IndivGroup_MainClass.BoardXMLDom.documentElement.selectSingleNode("Board[@id='"&Dv_IndivGroup_MainClass.BoardID&"']")
	XMLDom.documentElement.appendChild(Node)

	XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"picurl","")).text=Dvbbs.Forum_PicUrl
	XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"powerflag","")).text=Dv_IndivGroup_MainClass.PowerFlag
	Rem 插入各种图片的设置数据
	XMLDOM.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"pic_nofollow","")).text=Dvbbs.mainpic(10)
	XMLDOM.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"pic_follow","")).text=Dvbbs.mainpic(11)
	XMLDOM.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"istopic","")).text=Dvbbs.mainpic(1)
	XMLDOM.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"opentopic","")).text=Dvbbs.mainpic(2)
	XMLDOM.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"hottopic","")).text=Dvbbs.mainpic(3)
	XMLDOM.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"ilocktopic","")).text=Dvbbs.mainpic(4)
	XMLDOM.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"besttopic","")).text=Dvbbs.mainpic(5)

	Rem ===============传送论坛信息和设置数据到XML===============================================================
	Set Node=XMLDom.documentElement.appendChild(XMLDom.createNode(1,"forum_setting",""))
	Node.attributes.setNamedItem(XMLDom.createNode(2,"rss","")).text=Dvbbs.Forum_ChanSetting(2)'rss订阅
	Node.attributes.setNamedItem(XMLDom.createNode(2,"wap","")).text=Dvbbs.Forum_ChanSetting(1)'wap访问
	Node.attributes.setNamedItem(XMLDom.createNode(2,"ishot","")).text=Dvbbs.Forum_Setting(44)'热贴最少回复
	Node.attributes.setNamedItem(XMLDom.createNode(2,"pagesize","")).text=IndivGroup_EachPageTopicCount'列表分页大小
	Node.attributes.setNamedItem(XMLDom.createNode(2,"dispsize","")).text=IndivGroup_EachPagePostCount '贴子分页大小
	Node.attributes.setNamedItem(XMLDom.createNode(2,"newflagpic","")).text="Images/others/02.gif" '显示新贴标志的设置
	Node.attributes.setNamedItem(XMLDom.createNode(2,"ForumUrl","")).text=Dvbbs.Get_ScriptNameUrl()

	Rem 参数传递
	XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"action","")).text=Request("action")
	XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"page","")).text=Dv_IndivGroup_MainClass.Page

	If Dv_IndivGroup_MainClass.Page=1 Then
		Set topiclist=Dv_IndivGroup_MainClass.TopTopicXMLDom
		For Each Node in topiclist.documentElement.SelectNodes("row")
			Node.selectSingleNode("@title").text=Dvbbs.ChkBadWords(Node.selectSingleNode("@title").text)
			If Not Node.selectSingleNode("@topicmode").text ="1"  Then
				Node.selectSingleNode("@title").text=replace(Node.selectSingleNode("@title").text,"<","&lt;")
			End If
			Node.selectSingleNode("@lastpost").text=Dvbbs.ChkBadWords(Node.selectSingleNode("@lastpost").text)
			Node.selectSingleNode("@postusername").text=Dvbbs.ChkBadWords(Node.selectSingleNode("@postusername").text)
			i=0
			For each lastpost in split(Node.selectSingleNode("@lastpost").text,"$")
				Node.attributes.setNamedItem(topiclist.createNode(2,"lastpost_"& i,"")).text=lastpost
				i=i+1
			Next
			PostTime = Node.selectSingleNode("@lastpost_2").text
			Posttime = DateDiff("n",Posttime,Now)+Cint(Dvbbs.Forum_Setting(0))
			If Posttime < 60 Then Node.attributes.setNamedItem(topiclist.createNode(2,"datedifftime","")).text=Posttime
		Next
		XMLDom.documentElement.appendChild(topiclist.documentElement)
	End If
	
	Set topiclist=Dv_IndivGroup_MainClass.TopicXMLDom(IndivGroup_EachPageTopicCount)
	If Not topiclist Is Nothing Then 
		For Each Node in topiclist.documentElement.SelectNodes("row")
			Node.selectSingleNode("@title").text=Dvbbs.ChkBadWords(Node.selectSingleNode("@title").text)
			If Not Node.selectSingleNode("@topicmode").text ="1"  Then
				Node.selectSingleNode("@title").text=replace(Node.selectSingleNode("@title").text,"<","&lt;")
			End If
			Node.selectSingleNode("@postusername").text=Dvbbs.ChkBadWords(Node.selectSingleNode("@postusername").text)
			i=0
			For each lastpost in split(Node.selectSingleNode("@lastpost").text,"$")
				Node.attributes.setNamedItem(topiclist.createNode(2,"lastpost_"& i,"")).text=lastpost
				i=i+1
			Next
			PostTime = Node.selectSingleNode("@lastpost_2").text
			Posttime = DateDiff("n",Posttime,Now)+Cint(Dvbbs.Forum_Setting(0))
			If Posttime < 60 Then	Node.attributes.setNamedItem(topiclist.createNode(2,"datedifftime","")).text=Posttime
		Next
		XMLDom.documentElement.appendChild(topiclist.documentElement)
	End If	

	Set XSLTemplate=Dvbbs.iCreateObject("Msxml2.XSLTemplate" & MsxmlVersion )
	Set XMLStyle=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion )
	XMLStyle.loadxml template.html(2)
	'XMLStyle.Load Server.MapPath("Dv_plus/IndivGroup/Skin/TopicList.xslt")

	XSLTemplate.stylesheet=XMLStyle
	Set proc = XSLTemplate.createProcessor()
	proc.input = XMLDom
	proc.transform()
	Response.Write  proc.output
	Set XMLDom=Nothing 
	Set proc=Nothing
End Sub
%>