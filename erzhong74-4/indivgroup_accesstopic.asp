<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<!--#include file="Dv_plus/IndivGroup/Dv_IndivGroup_Config.asp"-->
<!--#include file="Dv_plus/IndivGroup/Dv_IndivGroup_MainCls.asp"-->
<%
Dim AdminLockTopic,XMLDom,TableList,paramnode,AccessSetting
Dim IGName
'Dvbbs.ShowSQL=1
AdminLockTopic=False 
If (Dvbbs.master or Dvbbs.superboardmaster) And Cint(Dvbbs.GroupSetting(36))=1 Then
	AdminLockTopic=True 
End If

If Not AdminLockTopic Then Response.redirect "showerr.asp?ErrCodes=<li>您没有在本版面审核帖子的权限。&action=OtherErr"

Call Main
Call ShowHTML
If Request("action")="view" Then
	Response.Write "</body></html>"
Else
	Dvbbs.activeonline()
	Dvbbs.footer()
End If
Dvbbs.PageEnd()

Sub Main()
	Dim Rs
	'LoadTableList()
	Set XMLDom=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	XMLDom.appendChild(XMLDom.createElement("xml"))
	Set paramnode=XMLDom.documentElement.appendChild(XMLDom.createNode(1,"param",""))
	paramnode.attributes.setNamedItem(XMLDom.createNode(2,"action","")).text=Request("action")
	If Dvbbs.Master Then paramnode.attributes.setNamedItem(XMLDom.createNode(2,"master","")).text=1

	Dv_IndivGroup_MainClass.ID=0
	If Not IGName="" Then
		Set Rs=Dvbbs.Execute("Select ID From Dv_GroupName Where GroupName='"&IGName&"'")
		If Not Rs.Eof Then
			Dv_IndivGroup_MainClass.ID=Rs(0)
			Dv_IndivGroup_MainClass.BoardID = Dvbbs.CheckNumeric(Request("boardid"))
		End If
		Rs.Close:Set Rs=Nothing
	End If

	Dvbbs.LoadTemplates("indivgroup")
	If Not Dvbbs.ChkPost() Then Response.Redirect "index.asp"
	Select Case Request("action")
		Case "manage"'批量审核
			Dvbbs.stats="批量审核"
			Dvbbs.Nav
			manage()
			If Dvbbs.BoardID=0 Then
				Dvbbs.Head_var 2,0,"",""
			Else
				Dvbbs.Head_var 1,Application(Dvbbs.CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"&Dvbbs.BoardID&"']/@depth").text,"",""
			End If
		Case "view"'查看单个审核贴
			Dvbbs.stats="审核贴子"
			View()
			Dvbbs.Head()
		Case "save"
			Dvbbs.stats="保存审核设置"
			Dvbbs.Nav
			Dvbbs.Head_var 2,0,"",""
			SaveAccessSetting()
		Case Else
			Dvbbs.stats="帖子审核"
			Dvbbs.Nav
			Dvbbs.Head_var 2,0,"",""

			IGName	= Trim(Request("igname"))
			LoadAccessCount()
			accesslist()
	End Select
End Sub

Sub SaveAccessSetting()
	If Request.form("action")="" Then Exit Sub
	Dim id,node,Dom,node1,queststr,checkuser,node2
	Set Dom=Dvbbs.CreateXmlDoc("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	Dom.appendChild(Dom.createElement("accesspost"))
	For Each id in Request("id")
		Set Node=Dom.documentElement.appendChild(Dom.createNode(1,"setting",""))
		Node.attributes.setNamedItem(Dom.createNode(2,"type","")).text=Request("setting_"&id&"_type")
		queststr=Request("setting_"&id&"_use")
		If queststr="" Then queststr="0"
		Node.attributes.setNamedItem(Dom.createNode(2,"use","")).text=queststr
		Set Node1=Node.appendChild(Dom.createNode(1,"check",""))
		queststr=Request("setting_"&id&"_check_new")
		If queststr="" Then queststr="0"
		Node1.attributes.setNamedItem(Dom.createNode(2,"new","")).text=queststr
		queststr=Request("setting_"&id&"_check_re")
		If queststr="" Then queststr="0"
		Node1.attributes.setNamedItem(Dom.createNode(2,"re","")).text=queststr
		queststr=Request("setting_"&id&"_check_edit")
		If queststr="" Then queststr="0"
		Node1.attributes.setNamedItem(Dom.createNode(2,"edit","")).text=queststr
		For Each checkuser in Request("setting_"&id&"_checkuser")
			Set Node1=Node.appendChild(Dom.createNode(1,"checkuser",""))
			Node1.attributes.setNamedItem(Dom.createNode(2,"usergroupid","")).text=Request("setting_"&id&"_checkuser_"&checkuser&"_usergroupid")
			queststr=Request("setting_"&id&"_checkuser_"&checkuser&"_use")
			If queststr="" Then queststr="0"
			Node1.attributes.setNamedItem(Dom.createNode(2,"use","")).text=queststr
			Set node2=Node1.appendChild(Dom.createNode(1,"usertopic",""))
			Node2.attributes.setNamedItem(Dom.createNode(2,"value","")).text=Request("setting_"&id&"_checkuser_"&checkuser&"_usertopic")
			queststr=Request("setting_"&id&"_checkuser_"&checkuser&"_usertopic_use")
			If queststr="" Then queststr="0"
			Node2.attributes.setNamedItem(Dom.createNode(2,"use","")).text=queststr
			Set node2=Node1.appendChild(Dom.createNode(1,"userpost",""))
			Node2.attributes.setNamedItem(Dom.createNode(2,"value","")).text=Request("setting_"&id&"_checkuser_"&checkuser&"_userpost")
			queststr=Request("setting_"&id&"_checkuser_"&checkuser&"_userpost_use")
			If queststr="" Then queststr="0"
			Node2.attributes.setNamedItem(Dom.createNode(2,"use","")).text=queststr
			Set node2=Node1.appendChild(Dom.createNode(1,"regdate",""))
			Node2.attributes.setNamedItem(Dom.createNode(2,"value","")).text=Request("setting_"&id&"_checkuser_"&checkuser&"_regdate")
			queststr=Request("setting_"&id&"_checkuser_"&checkuser&"_regdate_use")
			If queststr="" Then queststr="0"
			Node2.attributes.setNamedItem(Dom.createNode(2,"use","")).text=queststr
			Set node2=Node1.appendChild(Dom.createNode(1,"userdel",""))
			Node2.attributes.setNamedItem(Dom.createNode(2,"value","")).text=Request("setting_"&id&"_checkuser_"&checkuser&"_userdel")
			queststr=Request("setting_"&id&"_checkuser_"&checkuser&"_userdel_use")
			If queststr="" Then queststr="0"
			Node2.attributes.setNamedItem(Dom.createNode(2,"use","")).text=queststr
			Set node2=Node1.appendChild(Dom.createNode(1,"lockuser",""))
			queststr=Request("setting_"&id&"_checkuser_"&checkuser&"_lockuser_use")
			If queststr="" Then queststr="0"
			Node2.attributes.setNamedItem(Dom.createNode(2,"use","")).text=queststr
			Set node2=Node1.appendChild(Dom.createNode(1,"checkcontent",""))
			queststr=Request("setting_"&id&"_checkuser_"&checkuser&"_checkcontent_use")
			If queststr="" Then queststr="0"
			Node2.attributes.setNamedItem(Dom.createNode(2,"use","")).text=queststr
		Next
		Set Node1=Node.appendChild(Dom.createNode(1,"checkcontent",""))
		Set node2=Node1.appendChild(Dom.createNode(1,"checkpic",""))
		queststr=Request("setting_"&id&"_checkcontent_checkpic_use")
		If queststr="" Then queststr="0"
		Node2.attributes.setNamedItem(Dom.createNode(2,"use","")).text=queststr
		Set node2=Node1.appendChild(Dom.createNode(1,"checklink",""))
		queststr=Request("setting_"&id&"_checkcontent_checklink_use")
		If queststr="" Then queststr="0"
		Node2.attributes.setNamedItem(Dom.createNode(2,"use","")).text=queststr
		Set node2=Node1.appendChild(Dom.createNode(1,"checkflash",""))
		queststr=Request("setting_"&id&"_checkcontent_checkflash_use")
		If queststr="" Then queststr="0"
		Node2.attributes.setNamedItem(Dom.createNode(2,"use","")).text=queststr
		Set node2=Node1.appendChild(Dom.createNode(1,"checkmp",""))
		queststr=Request("setting_"&id&"_checkcontent_checkmp_use")
		If queststr="" Then queststr="0"
		Node2.attributes.setNamedItem(Dom.createNode(2,"use","")).text=queststr
		Set node2=Node1.appendChild(Dom.createNode(1,"checkrm",""))
		queststr=Request("setting_"&id&"_checkcontent_checkrm_use")
		If queststr="" Then queststr="0"
		Node2.attributes.setNamedItem(Dom.createNode(2,"use","")).text=queststr
		queststr=Request("setting_"&id&"_checkcontent_checkword")
		If queststr <>"" Then
			queststr=split(queststr,vbnewline)
			For Each Node2 in queststr
				If node2<>"" Then
					Node1.appendChild(Dom.createNode(1,"checkword","")).attributes.setNamedItem(Dom.createNode(2,"content","")).text=Node2
				End If
			Next
		End If
		Set Node1=Node.appendChild(Dom.createNode(1,"nocheck",""))
		queststr=Request("setting_"&id&"_nocheck")
		If queststr <>"" Then
			queststr=split(queststr,vbnewline)
			For Each Node2 in queststr
				If node2<>"" Then
					Node1.appendChild(Dom.createNode(1,"username","")).text=Node2
				End If
			Next
		End If
	Next
	Dvbbs.Execute("update Dv_setup Set Forum_BoardXML='"&Dvbbs.Checkstr(Replace(Replace(Dom.xml,Chr(13),""),Chr(10),""))&"'")
	Dvbbs.LoadSetup()
	Dvbbs.Execute("Insert Into Dv_Log (l_AnnounceID,l_BoardID,l_touser,l_username,l_content,l_ip,l_type) values (0,"&Dvbbs.BoardID&",'审核设置','" & Dvbbs.MemberName & "','更新','" & Dvbbs.userTrueIP & "',3)")
End Sub
Sub LoadAccessSetting()
	Dim dom,Node,i,position,position1,node1
	Set Dom=Application(Dvbbs.CacheName & "_accesstopic").cloneNode(True)
	If Request("addnew")="1"  and Request("action")="modify" Then
		Set Node=Dom.documentElement.appendChild(Dom.createNode(1,"setting",""))
		Node.attributes.setNamedItem(Dom.createNode(2,"type","")).text="新建设置"
		Node.appendChild(Dom.createNode(1,"checkuser",""))
	ElseIf Request("delsetting")="1" Then
		Set Node=Dom.documentElement.selectNodes("setting")
		position=CLng(Request("position"))
		If position < Node.length +1 Then
			Dom.documentElement.removeChild(Node(position-1))
		End If
	ElseIf Request("delusergroup")="1" Then
		Set Node=Dom.documentElement.selectNodes("setting")
		position=CLng(Request("position"))
		position1=CLng(Request("position1"))
		If position < Node.length +1 Then
			Set Node1=Node(position-1).selectNodes("checkuser")
			If position1 < Node1.length +1 Then
				Node(position-1).removeChild(Node1(position1-1))
			End If
		End If
	ElseIf Request("addusergroup")="1" Then
		Set Node=Dom.documentElement.selectNodes("setting")
		position=CLng(Request("position"))
		If position < Node.length +1 Then
			Node(position-1).appendChild(Dom.createNode(1,"checkuser",""))
		End If
	End If
	XMLDom.documentElement.appendChild(dom.documentElement)
	XMLDom.documentElement.appendChild(Application(Dvbbs.CacheName &"_grouppic").documentElement.cloneNode(True))
End Sub
Sub manage()
	Dim id,passed,replyid,i,node,PostTable,LockTopic,boardid,rs,today,isvote,PollID
	i=1
	For Each id in Request.form("id")
		Set Node = XMLDom.documentElement.appendChild(XMLDom.createNode(1,"result",""))
		If IsNumeric(id) and id<>"" Then
			id		= Dvbbs.CheckNumeric(id)
			replyid	= Request("replyid")(i)
			passed	= Request("pass_"&id&"_"& replyid)
			If replyid="" Then replyid=id
			replyid=Dvbbs.CheckNumeric(replyid)
			Node.attributes.setNamedItem(XMLDom.createNode(2,"rootid","")).text=id
			Node.attributes.setNamedItem(XMLDom.createNode(2,"announceid","")).text=replyid
			Rem 检查主题表是否有记录，并且取得其主贴状态
			Set Rs=Dvbbs.Execute("select GroupID,Boardid,LockTopic,PollID From Dv_Group_Topic Where topicid="& id &"")
			If Not Rs.EOF Then
				Node.setAttribute "groupid", Rs("GroupID")
				LockTopic=Rs("boardid")
				PollID=Rs("PollID")
				today=0
				If Passed="1" Then
					Rem 通过审核
					If replyid=id Then
						Set Rs=Dvbbs.Execute("select dateandtime,PostUserid,LockTopic From Dv_Group_BBs Where RootID="& id &" and Boardid=-2")
						If Not Rs.EOF Then 
							boardid=rs("LockTopic")
							If datediff("d",rs(0),Now()) =0 Then today=1
							Node.attributes.setNamedItem(XMLDom.createNode(2,"topic","")).text=1
							Node.attributes.setNamedItem(XMLDom.createNode(2,"child","")).text=0
							Node.attributes.setNamedItem(XMLDom.createNode(2,"today","")).text=today
							Node.attributes.setNamedItem(XMLDom.createNode(2,"boardid","")).text=boardid
							Node.attributes.setNamedItem(XMLDom.createNode(2,"color","")).text="green"
							Node.attributes.setNamedItem(XMLDom.createNode(2,"stats","")).text="通过审核成功。"
							Dvbbs.Execute("update Dv_Group_BBs set boardid="&boardid&",LockTopic=0 Where RootID="& id &" and ParentID=0 and Boardid=-2")
							Dvbbs.Execute("update Dv_Group_Topic Set boardid="&boardid&",LockTopic=0,Child=0  Where topicid="& id &" and Boardid=-2")
							'UpdatepostUser  rs(1),boardid,1
							If Rs(1)<>0 Then Dvbbs.Sendmessanger Rs(1),"系统[审核]","您发表的贴子已经通过审核，请<a href=""dispbbs.asp?boardid="& boardid&"&amp;id="&id&""" target=""_blank"">点此查看</a>"
						Else
							Node.attributes.setNamedItem(XMLDom.createNode(2,"topic","")).text=0
							Node.attributes.setNamedItem(XMLDom.createNode(2,"child","")).text=0
							Node.attributes.setNamedItem(XMLDom.createNode(2,"today","")).text=0
							Node.attributes.setNamedItem(XMLDom.createNode(2,"stats","")).text="失败，原因：找不到相关记录，数据可能已经被别的管理人员处理了。"
						End If
					Else
						Set Rs=Dvbbs.Execute("select dateandtime,PostUserid,ParentID,LockTopic From Dv_Group_BBs Where RootID="& id &" and AnnounceID="&replyid&" and  Boardid=-2")
						If Not Rs.EOF Then
							If datediff("d",rs(0),Now())=0 Then today=1
							boardid=rs("LockTopic")
							Node.attributes.setNamedItem(XMLDom.createNode(2,"topic","")).text=0
							Node.attributes.setNamedItem(XMLDom.createNode(2,"child","")).text=1
							Node.attributes.setNamedItem(XMLDom.createNode(2,"today","")).text=today
							Node.attributes.setNamedItem(XMLDom.createNode(2,"boardid","")).text=boardid
							Node.attributes.setNamedItem(XMLDom.createNode(2,"color","")).text="green"
							Node.attributes.setNamedItem(XMLDom.createNode(2,"stats","")).text="通过审核成功。"
							If Rs("ParentID")=0 Then
								Dvbbs.Execute("update Dv_Group_BBs set boardid="&boardid&",LockTopic=0 Where RootID="& id &" and ParentID=0")
								Dvbbs.Execute("update Dv_Group_Topic Set boardid="&boardid&",LockTopic=0,Child=0  Where topicid="& id)
								'UpdatepostUser  rs(1),boardid,1
								If Rs(1)<>0 Then Dvbbs.Sendmessanger Rs(1),"系统[审核]","您发表的贴子已经通过审核，请<a href=""dispbbs.asp?boardid="& boardid&"&amp;id="&id&""" target=""_blank"">点此查看</a>"
							Else
								Dvbbs.Execute("update Dv_Group_BBs set boardid="&boardid&",LockTopic=0 Where RootID="& id &" and AnnounceID="&replyid )
								Dvbbs.Execute("update dv_Group_topic Set boardid="&boardid&",LockTopic=0,Child=Child+1  Where topicid="& id)
								'UpdatepostUser  rs(1),boardid,0
								If Rs(1)<>0 Then Dvbbs.Sendmessanger Rs(1),"系统[审核]","您发表的贴子已经通过审核，请<a href=""dispbbs.asp?boardid="& boardid&"&amp;id="&id&"&amp;skin=1&amp;replyid="&replyid&""" target=""_blank"">点此查看</a>"
							End If
						Else
							Node.attributes.setNamedItem(XMLDom.createNode(2,"topic","")).text=0
							Node.attributes.setNamedItem(XMLDom.createNode(2,"child","")).text=0
							Node.attributes.setNamedItem(XMLDom.createNode(2,"today","")).text=0
							Node.attributes.setNamedItem(XMLDom.createNode(2,"stats","")).text="失败，原因：找不到相关记录,数据可能已经被别的管理人员处理了。"
						End If
					End If
				ElseIf Passed="0" Then
					Rem 删除
					If CLng(replyid)=id Then
						Set Rs=Dvbbs.Execute("select PostUserid From Dv_Group_BBs Where RootID="& id &"")
						If Not Rs.EOF Then
							If Rs(0)<>0 Then Dvbbs.Sendmessanger Rs(0),"系统[审核]","您发表的贴子未能通过审核，请注意您发表的内容。"
						End If
						'If isvote=1 Then
						'	Dvbbs.Execute("delete From Dv_vote  Where voteid="& PollID &"")
						'End If
						Dvbbs.Execute("delete From Dv_Group_BBs Where RootID="& id &"")
						Dvbbs.Execute("delete From Dv_Group_Topic Where topicid="& id)
					Else
						Set Rs=Dvbbs.Execute("select ParentID,PostUserid From Dv_Group_BBs Where RootID="& id &" and AnnounceID="&replyid )
						If Not Rs.EOF Then
							If Rs(1)<>0 Then Dvbbs.Sendmessanger Rs(1),"系统[审核]","您发表的贴子未能通过审核，请注意您发表的内容。"
							If  Rs(0) <> 0 Then 
								Dvbbs.Execute("delete From Dv_Group_BBs Where RootID="& id &" and AnnounceID="&replyid)
							Else
								Dvbbs.Execute("delete From Dv_Group_BBs Where RootID="& id &"")
								Dvbbs.Execute("delete From Dv_Group_Topic Where topicid="& id)
							End If
						End If
					End If
					'清除上传附件 2005-12-5 Dv.Yz
					'Dvbbs.Execute("UPDATE Dv_Upfile SET F_Flag = 4 WHERE F_AnnounceID = '" & Id & "|" & Replyid & "'")
					Node.attributes.setNamedItem(XMLDom.createNode(2,"topic","")).text=0
					Node.attributes.setNamedItem(XMLDom.createNode(2,"child","")).text=0
					Node.attributes.setNamedItem(XMLDom.createNode(2,"today","")).text=0
					Node.attributes.setNamedItem(XMLDom.createNode(2,"color","")).text="red"
					Node.attributes.setNamedItem(XMLDom.createNode(2,"stats","")).text="删除待审核贴成功。"
				Else
					Node.attributes.setNamedItem(XMLDom.createNode(2,"topic","")).text=0
					Node.attributes.setNamedItem(XMLDom.createNode(2,"child","")).text=0
					Node.attributes.setNamedItem(XMLDom.createNode(2,"today","")).text=0
					Node.attributes.setNamedItem(XMLDom.createNode(2,"stats","")).text="待审，您没有对该贴进行处理。"
				End If
			Else
				Node.attributes.setNamedItem(XMLDom.createNode(2,"topic","")).text=0
				Node.attributes.setNamedItem(XMLDom.createNode(2,"child","")).text=0
				Node.attributes.setNamedItem(XMLDom.createNode(2,"today","")).text=0
				Node.attributes.setNamedItem(XMLDom.createNode(2,"stats","")).text="失败，原因：找不到相关记录。"
			End If
		Else
			Node.attributes.setNamedItem(XMLDom.createNode(2,"topic","")).text=0
			Node.attributes.setNamedItem(XMLDom.createNode(2,"child","")).text=0
			Node.attributes.setNamedItem(XMLDom.createNode(2,"today","")).text=0
			Node.attributes.setNamedItem(XMLDom.createNode(2,"stats","")).text="失败，原因：参数错误。"
		End If
		i=i+1
	Next
	Dim Results,allpost,alltopic,alltoday,topic,Child,TmpID
	allpost=0
	alltopic=0
	alltoday=0
	'统计一下更新情况

	For Each Results In	XMLDom.documentElement.selectNodes("result")
		If Dv_IndivGroup_MainClass.ID=0 Then
			If Dvbbs.CheckNumeric(Results.getAttribute("groupidflag"))=0 Then
				Set Node =XMLDom.documentElement.selectNodes("result[@groupid="& Results.getAttribute("groupid") &"]")
				If Node.length > 0 Then
					topic=0
					Child=0
					today=0
					For Each TmpID in node
						topic=topic+CLng(tmpid.selectSingleNode("@topic").text)
						Child=Child+CLng(tmpid.selectSingleNode("@child").text)
						today=today+CLng(tmpid.selectSingleNode("@today").text)
						TmpID.setAttribute "groupidflag", 1
					Next
					If topic+Child >0 Then
						alltopic=alltopic+topic
						allpost=allpost+topic+Child
						alltoday=alltoday+today
						UpDate_GroupInfo Results.getAttribute("groupid"),topic,Child,today
					End If
				End If
			End If
		End If
		If Dvbbs.CheckNumeric(Results.getAttribute("boardidflag"))=0 And Dvbbs.CheckNumeric(Results.getAttribute("boardid"))>0 Then
			Set Node =XMLDom.documentElement.selectNodes("result[@boardid="& Results.getAttribute("boardid") &"]")
			If Node.length > 0 Then
				topic=0
				Child=0
				today=0
				For Each TmpID in node
					topic=topic+CLng(tmpid.selectSingleNode("@topic").text)
					Child=Child+CLng(tmpid.selectSingleNode("@child").text)
					today=today+CLng(tmpid.selectSingleNode("@today").text)
					TmpID.setAttribute "boardidflag", 1
				Next
				If topic+Child >0 Then
					alltopic=alltopic+topic
					allpost=allpost+topic+Child
					alltoday=alltoday+today
					UpDate_BoardInfo Results.getAttribute("boardid"),topic,Child,today
				End If
			End If
		End If
	Next

	Tolog("批量审核")
End Sub
Sub View()
	Dim Node,id,replyid,Rs,PostTable,SQL
	id		= Dvbbs.CheckNumeric(Request("id"))
	replyid	= Dvbbs.CheckNumeric(Request("replyid"))
'	If Not IsNumeric(replyid) Or replyid="" Then replyid=0
	If id=0 Then Response.redirect "showerr.asp?ErrCodes=<li>请指定所需参数。&action=OtherErr"

	Set Rs=Dvbbs.Execute("Select boardid From Dv_Group_Topic Where TopicID="&Id)
	If Rs.EOF Then
		Response.redirect "showerr.asp?ErrCodes=<li>记录不存在！&action=OtherErr"
	Else
'		PostTable=Rs(0)
		If replyid=0 Then
			SQL="Select * From Dv_Group_BBS Where rootid="&ID&" And ParentID=0 And Boardid=-2"
		Else
			SQL="Select * From Dv_Group_BBS Where rootid="&ID&" And AnnounceID="&replyid&" and Boardid=-2"
		End If
		Set Rs=Dvbbs.Execute(SQL)
		If Rs.EOF Then
			Response.redirect "showerr.asp?ErrCodes=<li>找不到匹配记录！&action=OtherErr"
		Else
			XMLDom.documentElement.appendChild(Dvbbs.RecordsetToxml(rs,"row","").documentElement.firstChild)
		End If
	End If
End Sub
Sub ShowHTML()
	Dim xslt,proc,XMLStyle
	Set XMLStyle=Dvbbs.CreateXmlDoc("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	XMLStyle.loadxml template.html(6)
	'XMLStyle.load Server.MapPath("inc/AccessTopic.xslt")
	Set XSLT=Dvbbs.iCreateObject("Msxml2.XSLTemplate" & MsxmlVersion)
	XSLT.stylesheet=XMLStyle
	Set proc = XSLT.createProcessor()
	proc.input = XMLDom
	proc.transform()
	Response.Write  proc.output
	'XMLDom.save Server.MapPath("IndivGroup_AccessTopic.xml")
	Set XMLDOM=Nothing
	Set XSLt=Nothing
	Set proc=Nothing	
End Sub
Sub accesslist()
	Dim TableID,PostTable
	Dim SQL,node,TmpSQL,Rs,SQL1,Pagesize,Page,pagecount
	Dim Keyword

	Pagesize=30		'手工设置每页最大显示30条
	Keyword	= Trim(Request("keyword"))
	TableID=Dvbbs.CheckNumeric(Request("tableid"))
	'传送参数到xml
	paramnode.setAttribute "igname", IGName
	paramnode.setAttribute "boardid", Dv_IndivGroup_MainClass.BoardID
	paramnode.setAttribute "tableid", TableID
	paramnode.setAttribute "keyword", Keyword
	paramnode.setAttribute "pagesize", Pagesize

	Page=Dvbbs.CheckNumeric(Request("Page"))
	If Page=0 Then Page=1

	'根据页面参数产生查询代码
	IGName=Dvbbs.CheckStr(IGName)
	Keyword=Dvbbs.Checkstr(Keyword)

	If TableID=0 Then
		PostTable="Dv_Group_Topic"
		SQL = "topicid as id,GroupID,Title as topic,LockTopic as bid,PostUsername as username,PostUserid as userid,DateAndTime"
		TmpSQL=" And (Title like '%"&Keyword&"%' or PostUsername='"&Keyword&"')"
	Else
		PostTable="Dv_Group_BBS"
		SQL ="rootid as id,topic,GroupID,body,LockTopic as bid,username,PostUserid as userid,AnnounceID as replyID,DateAndTime,ParentID"
		TmpSQL=" And (Topic like '%"&Keyword&"%' or Username='"&Keyword&"')"
	End If

	SQL ="Select "&SQL&" From "&PostTable&" Where Boardid=-2"
	SQL1 ="Select Count(*) as Length From "&PostTable&" Where Boardid=-2"

	If Dv_IndivGroup_MainClass.ID > 0 Then
		SQL= SQL & " And GroupID="&Dv_IndivGroup_MainClass.ID
		SQL1= SQL1 & " And GroupID="&Dv_IndivGroup_MainClass.ID
		If Dv_IndivGroup_MainClass.BoardID > 0 Then
			SQL= SQL &" And LockTopic="&Dv_IndivGroup_MainClass.BoardID
			SQL1= SQL1 &" And LockTopic="&Dv_IndivGroup_MainClass.BoardID
		End If
	End If

	If Not Keyword="" Then
		SQL= SQL & TmpSQL
		SQL1= SQL1 & TmpSQL
	End If
	If TableID="0" Then
		SQL = SQL&" Order By TopicID"
	Else
		SQL = SQL&" Order By AnnounceID"
	End If
	Set Rs=Dvbbs.Execute(SQL1)
	paramnode.setAttribute "searchcount", Rs(0)
	'计算一下当前Page参数是否合法。如果超出范围，强制为最后一页
	If Rs(0) mod Pagesize =0 then
			PageCount= Rs(0) \ Pagesize
	Else
			PageCount= Rs(0) \ Pagesize+1
	End If
	If Page > PageCount Then Page=PageCount
	paramnode.setAttribute "page", Page
	If Not (Rs(0)=0 Or IsNull(Rs(0)))Then
		Set Rs=Dvbbs.Execute(SQL)
		If Not page=1 Then 	Rs.Move(pagesize*(page))
		SQL=RS.GetRows(Pagesize)
		Set Node=Dvbbs.ArrayToxml(SQL,rs,"row","datarows")
		Rs.Close:Set Rs=Nothing
		XMLDom.documentElement.appendChild(node.documentElement)

		Dim TempNode,TempGroupIDList,TempBoardIDList
		TempGroupIDList="" : TempBoardIDList=""
		If Dv_IndivGroup_MainClass.ID=0 Then
			For Each Node In XMLDom.documentElement.selectNodes("/xml/datarows/row")
				If TempGroupIDList="" Then
					TempGroupIDList=Node.getAttribute("groupid")
				Else
					TempGroupIDList=TempGroupIDList &","& Node.getAttribute("groupid")
				End If
				If TempBoardIDList="" Then
					TempBoardIDList=Node.getAttribute("bid")
				Else
					TempBoardIDList=TempBoardIDList &","& Node.getAttribute("bid")
				End If
			Next
			Set Rs=Dvbbs.Execute("Select ID,GroupName From Dv_GroupName Where ID In ("&TempGroupIDList&")")
			If Not Rs.Eof Then
				Set Node=Dvbbs.RecordsetToxml(Rs,"row","grouplist")
				XMLDom.documentElement.appendChild(node.documentElement)
			End If
			Set Rs=Dvbbs.Execute("Select ID,BoardName From Dv_Group_Board Where ID In ("&TempBoardIDList&")")
			If Not Rs.Eof Then
				Set Node=Dvbbs.RecordsetToxml(Rs,"row","boardlist")
				XMLDom.documentElement.appendChild(node.documentElement)
			End If
		Else
			Set Rs=Dvbbs.Execute("Select ID,GroupName From Dv_GroupName Where ID="&Dv_IndivGroup_MainClass.ID)
			If Not Rs.Eof Then
				Set Node=Dvbbs.RecordsetToxml(Rs,"row","grouplist")
				XMLDom.documentElement.appendChild(node.documentElement)
			End If
			Set Rs=Dvbbs.Execute("Select ID,BoardName From Dv_Group_Board Where RootID="&Dv_IndivGroup_MainClass.ID)
			If Not Rs.Eof Then
				Set Node=Dvbbs.RecordsetToxml(Rs,"row","boardlist")
				XMLDom.documentElement.appendChild(node.documentElement)
			End If
		End If
	End If
	Rs.Close:Set Rs=Nothing
End Sub

'Sub LoadTableList()
'	Dim Rs
'	Set Rs=Dvbbs.Execute("select * from [Dv_TableList]")
'	TableList=Rs.GetRows()
'	Set XMLDom=Dvbbs.ArrayToxml(TableList,Rs,"PostTable","xml")
'End Sub

Sub LoadAccessCount()
	Dim SQL
	SQL = ""
	If Dv_IndivGroup_MainClass.ID > 0 Then
		paramnode.setAttribute "igid", Dv_IndivGroup_MainClass.ID
		paramnode.setAttribute "boardid", Dv_IndivGroup_MainClass.Boardid
		SQL = " And GroupID="& Dv_IndivGroup_MainClass.ID
		If Dv_IndivGroup_MainClass.BoardID>0 Then SQL = SQL &" And LockTopic="&Dv_IndivGroup_MainClass.BoardID
	End If
	XMLDom.documentElement.setAttribute "topiccount", Dvbbs.Execute("select Count(*) From Dv_Group_Topic Where boardid=-2"&SQL)(0)
	XMLDom.documentElement.setAttribute "postcount", Dvbbs.Execute("select Count(*) From Dv_Group_bbs Where boardid=-2"&SQL)(0)
End Sub

Sub UpDate_GroupInfo(GroupID,topic,Child,today)
	Dim SQL
	SQL="update Dv_GroupName set PostNum=PostNum+"&topic+Child&",TopicNum=TopicNum+"&topic&" where ID="&GroupID
	Dvbbs.Execute(sql)
End Sub

Sub UpDate_BoardInfo(BoardID,topic,Child,today)
	Dim SQL
	SQL="update Dv_Group_Board set PostNum=PostNum+"&topic+Child&",TopicNum=TopicNum+"&topic&" where ID="&BoardID
	Dvbbs.Execute(sql)
End Sub

Sub Tolog(Info)
	Dvbbs.Execute("Insert Into Dv_Log (l_AnnounceID,l_BoardID,l_touser,l_username,l_content,l_ip,l_type) values (0,"&Dvbbs.BoardID&",'审核贴子','" & Dvbbs.MemberName & "','" & Dvbbs.CheckStr(Info) & "','" & Dvbbs.userTrueIP & "',3)")
End Sub
%>