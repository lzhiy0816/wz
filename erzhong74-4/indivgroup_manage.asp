<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="inc/dv_clsother.asp" -->
<!--#include file="Dv_plus/IndivGroup/Dv_IndivGroup_Config.asp"-->
<!--#include file="Dv_plus/IndivGroup/Dv_IndivGroup_MainCls.asp"-->
<%
Dim Rs,SQL
Dim XMLDom,Node,XSLTemplate,XMLStyle,proc
Dvbbs.LoadTemplates("indivgroup")

If Dv_IndivGroup_MainClass.ID=0 Or Dv_IndivGroup_MainClass.Name="" Then Response.redirect "showerr.asp?ErrCodes=对不起，你访问的圈子不存在或已经被删除1&action=OtherErr"
If Dv_IndivGroup_MainClass.PowerFlag=0 Or Dv_IndivGroup_MainClass.PowerFlag>3 Then Response.redirect "showerr.asp?ErrCodes=<li>只有该圈子管理员或论坛管理员才能进入圈子管理页面。&action=OtherErr"

Dv_IndivGroup_MainClass.stats="圈子管理"
Dvbbs.Nav()
Dv_IndivGroup_MainClass.Head_var 2,"",""
Select Case LCase(Request("managetype"))
	Case "info"
		InfoManage()
	Case "board"
		BoardManage()
	Case "user"
		UserManage()
	Case "updatedata"
		UpdateData()
	Case Else
		InfoManage()
End Select
Dvbbs.Footer()
Dvbbs.PageEnd()
Sub InfoManage()
	If LCase(Request("action"))="infosave" Then
		Dim GroupName,GroupInfo,AppUserID,AppUserName,GroupStats,GroupSetting,GroupViewflag

		GroupName = Dvbbs.CheckStr(Dvbbs.Replacehtml(Request("groupname")))
		GroupInfo = Dvbbs.CheckStr(NewlineEnCode(Request("groupinfo")))
		AppUserID = Dvbbs.CheckNumeric(Request("appuserid"))
		Response.write "<script language=""javascript"">"
		If Dv_IndivGroup_MainClass.PowerFlag=1 or Dv_IndivGroup_MainClass.PowerFlag=2 Then
			Set Rs=Dv_IndivGroup_MainClass.Execute("Select UserName From Dv_GroupUser Where GroupID="&Dv_IndivGroup_MainClass.ID&" And UserID="&AppUserID&" And IsLock=2")
			If Not Rs.Eof Then
				AppUserName = Rs(0)
			Else
				Response.write "parent.document.getElementById('infoForm').innerHTML='<font color=""red"">错误：ID为”"&AppUserID&"“的成员不是圈子管理员，不能把创建权转让给该成员。</font>'"
			End If
		End If
		GroupStats = Dvbbs.CheckNumeric(Request("groupstats"))
		GroupSetting = Dvbbs.CheckNumeric(Request("groupsetting"))
		GroupViewflag = Dvbbs.CheckNumeric(Request("Viewflag"))
		Set Rs = Dvbbs.iCreateObject ("adodb.recordset")
		SQL = "Select ClassId,GroupName,GroupLogo,GroupInfo,AppUserID,AppUserName,Stats,Locked,ViewFlag From Dv_GroupName Where ID="&Dv_IndivGroup_MainClass.ID
		Rs.Open SQL,Conn,1,3
		If Not Rs.Eof Then
			Rs("ClassId") = Dvbbs.CheckNumeric(Request("ClassId"))
			Rs("GroupName") = GroupName
			Rs("GroupInfo") = GroupInfo
			Rs("GroupLogo") = Dvbbs.CheckStr(Request("GroupLogo"))
			If Dv_IndivGroup_MainClass.PowerFlag<3 Then
				Rs("AppUserID") = AppUserID
				Rs("AppUserName") = AppUserName
			End If
			Rs("Stats") = GroupStats
			Rs("Locked") = GroupSetting
			Rs("ViewFlag") = GroupViewflag
			Rs.Update
			Dvbbs.Execute("Update Dv_Group_Class Set GroupCount=(Select Count(ClassId) From Dv_GroupName Where Stats=1 And ClassId=Dv_Group_Class.Id)")
			Dv_IndivGroup_MainClass.LoadGroupClass()
			Response.write "parent.document.getElementById('infoForm').innerHTML='<li>圈子信息更新成功。[<a href=\""?managetype=info&action=info&groupid="&Dv_IndivGroup_MainClass.ID&"\"">查看基本信息</a>]'"
		Else
			Response.write "parent.document.getElementById('infoForm').innerHTML='<font color=""red"">错误:GroupID参数错误，不能编辑圈子信息，可能是该圈子被删除了。</font>'"
		End If
		Response.write "</script>"
	Else
		Set XMLDom=Dv_IndivGroup_MainClass.InfoXMLDom
		If LCase(Request("action"))="infoedit" Then XMLDom.documentElement.firstChild.selectSingleNode("@groupinfo").text = CodeEnNewline(Dv_IndivGroup_MainClass.Info)

		XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"groupid","")).text=Dv_IndivGroup_MainClass.ID
		XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"powerflag","")).text=Dv_IndivGroup_MainClass.PowerFlag
		XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"action","")).text=LCase(Request("action"))
		XMLDom.documentElement.appendChild(Dv_IndivGroup_MainClass.MasterXMLDom.documentElement.cloneNode(True))
		Rem 插入圈子分类数据
		XMLDom.documentElement.appendChild(Application(Dvbbs.CacheName & "_Dv_Group_Class").documentElement.cloneNode(True))

		Set XSLTemplate=Dvbbs.iCreateObject("Msxml2.XSLTemplate" & MsxmlVersion )
		Set XMLStyle=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion )
		'Response.clear:Response.write XMLDom.xml:Response.end
		XMLStyle.loadxml template.html(5)
		'XMLStyle.Load Server.MapPath("Dv_Plus/IndivGroup/Skin/Manage.xslt")

		XSLTemplate.stylesheet=XMLStyle
		Set proc = XSLTemplate.createProcessor()
		proc.input = XMLDom
		proc.transform()
		Response.Write  proc.output
		Set XMLDom=Nothing
		Set proc=Nothing
	End If
End Sub

Sub BoardManage()
	Dim BoardID
	If Request("action")="boardsave" Then
		Dim boardname,boardinfo,indeximg,boardrules,boardstats
		boardid = Dvbbs.CheckNumeric(Request("groupboardid"))
		boardname = Dvbbs.CheckStr(Dvbbs.Replacehtml(Request("boardname")))
		boardinfo = Dvbbs.CheckStr(NewlineEncode(Request("boardinfo")))
		indeximg = Dvbbs.CheckStr(Request("indeximg"))
		boardrules = Dvbbs.CheckStr(NewlineEncode(Request("boardrules")))
		boardstats = Dvbbs.CheckNumeric(Request("boardstats"))
		Response.write "<script language=""javascript"">"
		If boardname="" Then Response.write "parent.document.getElementById('boardnamestr').innerHTML='←<font color=""red"">错误：栏目不能为空</font>'":Exit Sub
		If boardid>0 Then
			Dv_IndivGroup_MainClass.Execute("Update Dv_Group_Board Set boardname='"&boardname&"',boardinfo='"&boardinfo&"',indeximg='"&indeximg&"',founddate='"&now()&"',rules='"&boardrules&"',boardstats="&boardstats&" Where rootid="&Dv_IndivGroup_MainClass.ID&" And ID="&boardid)
			Response.write "parent.document.getElementById('boardForm').innerHTML='<li>栏目编辑成功。[<a href=\""javascript:history.go(-1)\"">返回</a>]'"
		Else
			Dv_IndivGroup_MainClass.Execute("insert into Dv_Group_Board(boardname,boardinfo,indeximg,rootid,founddate,rules,boardstats) values('"&boardname&"','"&boardinfo&"','"&indeximg&"',"&Dv_IndivGroup_MainClass.ID&",'"&now()&"','"&boardrules&"',"&boardstats&")")
			Response.write "parent.document.getElementById('boardForm').innerHTML='<li>栏目增加成功。[<a href=\""javascript:history.go(-1)\"">返回</a>]'"
		End If
		Response.write "</script>"
	ElseIf Request("action")="boarddelete" Then
		BoardID = Dv_IndivGroup_MainClass.BoardID
		Set Rs = Dv_IndivGroup_MainClass.Execute("Select PostNum,TopicNum,TodayNum From Dv_Group_Board Where RootID="&Dv_IndivGroup_MainClass.ID&" And ID="&BoardID)
		If Not Rs.Eof Then
			Dv_IndivGroup_MainClass.Execute("Update Dv_GroupName Set PostNum=PostNum-"&Rs(0)&",TopicNum=TopicNum-"&Rs(1)&",TodayNum=TodayNum-"&Rs(2)&" Where ID="&Dv_IndivGroup_MainClass.ID)
			Dv_IndivGroup_MainClass.Execute("Delete From Dv_Group_Board Where RootID="&Dv_IndivGroup_MainClass.ID&" And ID="&BoardID)
			Dv_IndivGroup_MainClass.Execute("Delete From Dv_Group_Topic Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&BoardID)
			Dv_IndivGroup_MainClass.Execute("Delete From Dv_Group_BBS Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&BoardID)
		End IF
		Rs.Close:Set Rs=Nothing
		Response.write "<script language=""javascript"">"
		Response.write "alert('删除成功');"
		Response.write "parent.document.getElementById('board_"&BoardID&"').style.display='none';"
		Response.write "</script>"
	Else
		Set XMLDom=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		XMLDom.appendChild(XMLDom.createElement("IndivGroup"))
		XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"groupid","")).text=Dv_IndivGroup_MainClass.ID
		XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"powerflag","")).text=Dv_IndivGroup_MainClass.PowerFlag
		XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"action","")).text=LCase(Request("action"))

		If LCase(Request("action"))="boardmanage" Then
			Set Node=Dv_IndivGroup_MainClass.BoardXMLDom.documentElement.selectSingleNode("Board[@id='"&Dv_IndivGroup_MainClass.BoardID&"']")
			If Not Node Is Nothing Then
				Node.selectSingleNode("@boardinfo").text = CodeEnNewline(Node.selectSingleNode("@boardinfo").text)
				Node.selectSingleNode("@rules").text = CodeEnNewline(Node.selectSingleNode("@rules").text)
				XMLDom.documentElement.appendChild(Node)
			End If
		Else
			XMLDom.documentElement.appendChild(Dv_IndivGroup_MainClass.BoardXMLDom.documentElement.cloneNode(True))
			For Each Node In XMLDom.documentElement.selectNodes("BoardList/Board")
				Node.setAttribute "boardname",Replace(Node.getAttribute("boardname"),"'","%27")
			Next
		End If
		Set XSLTemplate=Dvbbs.iCreateObject("Msxml2.XSLTemplate" & MsxmlVersion )
		Set XMLStyle=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		'Response.clear:Response.write XMLDom.xml:Response.end
		XMLStyle.loadxml template.html(5)
		'XMLStyle.Load Server.MapPath("Dv_Plus/IndivGroup/Skin/Manage.xslt")

		XSLTemplate.stylesheet=XMLStyle
		Set proc = XSLTemplate.createProcessor()
		proc.input = XMLDom
		proc.transform()
		Response.Write  proc.output
		Set XMLDom=Nothing
		Set proc=Nothing
	End If
End Sub

Sub UserManage()
	Dim GroupUserID,loadlink
	Select Case LCase(Request("action"))
		Case "usersave"
			Dim UserStats,UserIntro,UserName
			GroupUserID = Dvbbs.CheckNumeric(Request("GroupUserID"))
			UserStats = Dvbbs.CheckNumeric(Request("userstats"))
			UserIntro = Dvbbs.CheckStr(NewlineEncode(Request("userintro")))
			Response.write "<script language=""javascript"">"

			Set Rs=Dv_IndivGroup_MainClass.Execute("Select UserID,UserName,IsLock From Dv_GroupUser Where ID="&GroupUserID&" And GroupID="&Dv_IndivGroup_MainClass.ID)
			If Not Rs.Eof Then
                '当前操作员有权利权限并且被修改的人不是管理员,创始人有权修改所有人,管理员有权修改所有人
				If (Dv_IndivGroup_MainClass.PowerFlag=3 And Rs(2)<>2) Or Dv_IndivGroup_MainClass.PowerFlag=2 Or Dvbbs.Master  Then
					If Rs(2)=0 And UserStats>0 And Dv_IndivGroup_MainClass.UserNum = Dv_IndivGroup_MainClass.LimitMemberNum Then
						Response.write "parent.document.getElementById('userForm').innerHTML='<li>修改成员资料失败，圈子已经达到成员数上限，不能修改审核中成员的状态。'"
					Else
						Dv_IndivGroup_MainClass.Execute("Update Dv_GroupUser Set IsLock="&UserStats&",Intro='"&UserIntro&"' Where ID="&GroupUserID&" And GroupID="&Dv_IndivGroup_MainClass.ID)

						SQL = ""
						If Rs(2)=0 And UserStats>0 Then
							SQL = "UserNum=UserNum+1"
							If Dv_IndivGroup_MainClass.UserNum+1 = Dv_IndivGroup_MainClass.LimitMemberNum Then SQL = SQL & ",Stats=2"
						End If
						If Rs(2)>1 And UserStats=0 Then SQL = "UserNum=UserNum-1"
						If SQL <> "" Then Dv_IndivGroup_MainClass.Execute("Update Dv_GroupName Set "&SQL&" Where ID="&Dv_IndivGroup_MainClass.ID)
					End If
					Response.write "parent.document.getElementById('userForm').innerHTML='<li>修改成员资料成功'"
				Else
					UserName=Rs(1)
					Response.write "parent.document.getElementById('userForm').innerHTML='<li>对不起，成员“"&UserName&"”是管理员，不能修改同等级成员信息。'"
				End If
			Else
				Response.write "parent.document.getElementById('userForm').innerHTML='错误：修改失败，成员已经退出或被踢出“"&Dv_IndivGroup_MainClass.Name&"”圈子。';"
			End If
			Rs.Close:Set Rs=Nothing

			Response.write "</script>"
		Case "deleteuser"
			GroupUserID = Dvbbs.CheckNumeric(Request("GroupUserID"))
			Response.write "<script language=""javascript"">"

			Set Rs=Dv_IndivGroup_MainClass.Execute("Select UserID,UserName,IsLock From Dv_GroupUser Where ID="&GroupUserID&" And GroupID="&Dv_IndivGroup_MainClass.ID)
			If Not Rs.Eof Then
                '操作者有管理全并且被修改的人不是管理员且不是自己
				If Dv_IndivGroup_MainClass.PowerFlag<3 And Rs(2)<>2 And Rs(0)<>Dvbbs.UserID Then
                    '处理删除用户不干净的问题
                    Dim RsTemp,UserGroupTemp
                    Set RsTemp = Dv_IndivGroup_MainClass.Execute("Select UserGroup From Dv_user where UserId="&Rs(0))
                    If Not RsTemp.EOF Then
                        UserGroupTemp = RsTemp(0)
                        UserGroupTemp = Replace(UserGroupTemp,Dv_IndivGroup_MainClass.ID&",","")
                        Dv_IndivGroup_MainClass.Execute("Update Dv_User Set UserGroup='"&UserGroupTemp&"' where UserId="&Rs(0))
                    End If
                    Set RsTemp = Nothing
                    '以上部分处理删除用户不干净的问题
                    If Rs(2)<>0 Then '如果是还没有通过审核的用户，则不需要减少用户总数
					    Dv_IndivGroup_MainClass.Execute("Update Dv_GroupName Set UserNum=UserNum-1 Where ID="&Dv_IndivGroup_MainClass.ID)
                    End If
					Dv_IndivGroup_MainClass.Execute("Delete From Dv_GroupUser Where ID="&GroupUserID&" And GroupID="&Dv_IndivGroup_MainClass.ID)

					Response.write "alert('删除成员资料成功');"
					loadlink = Split(Request.ServerVariables("HTTP_REFERER"))
					Response.write "parent.location.href='"&loadlink(UBound(loadlink))&"';"
				Else
					UserName=Rs(1)
					Response.write "alert('对不起，成员“"&UserName&"”是管理员，不能删除同等级成员信息。');"
				End If
			Else
				Response.write "alert('错误：删除失败，成员不存在。');"
			End If
			Rs.Close:Set Rs=Nothing

			Response.write "</script>"
		Case "passuser"
			GroupUserID = Dvbbs.CheckNumeric(Request("GroupUserID"))
			Response.write "<script language=""javascript"">"

			If Dv_IndivGroup_MainClass.UserNum < Dv_IndivGroup_MainClass.LimitMemberNum Then
				Dv_IndivGroup_MainClass.Execute("Update Dv_GroupUser Set IsLock=1 Where ID="&GroupUserID&" And IsLock=0 And GroupID="&Dv_IndivGroup_MainClass.ID)
				If Dv_IndivGroup_MainClass.UserNum+1 = Dv_IndivGroup_MainClass.LimitMemberNum Then
					Dv_IndivGroup_MainClass.Execute("Update Dv_GroupName Set UserNum=UserNum+1,Stats=2 Where ID="&Dv_IndivGroup_MainClass.ID)
				Else
					Dv_IndivGroup_MainClass.Execute("Update Dv_GroupName Set UserNum=UserNum+1 Where ID="&Dv_IndivGroup_MainClass.ID)
				End If

				Response.write "alert('审核通过成功');"
				loadlink = Split(Request.ServerVariables("HTTP_REFERER"))
				Response.write "parent.location.href='"&loadlink(UBound(loadlink))&"';"
			Else
				Response.write "alert('审核通过失败，已经达到圈子成员数上限。');"
				loadlink = Split(Request.ServerVariables("HTTP_REFERER"))
				Response.write "parent.location.href='"&loadlink(UBound(loadlink))&"';"
			End If

			Response.write "</script>"
		Case "setadmin"
			GroupUserID = Dvbbs.CheckNumeric(Request("GroupUserID"))
			Response.write "<script language=""javascript"">"

			Dv_IndivGroup_MainClass.Execute("Update Dv_GroupUser Set IsLock=2 Where ID="&GroupUserID&" And IsLock=1 And GroupID="&Dv_IndivGroup_MainClass.ID)

			Response.write "alert('设置管理员成功');"
			loadlink = Split(Request.ServerVariables("HTTP_REFERER"))
			Response.write "parent.location.href='"&loadlink(UBound(loadlink))&"';"

			Response.write "</script>"
		Case Else
			Dim Query,UserXMLDom
			Query = Dvbbs.CheckNumeric(Request("query"))
			Set XMLDom=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			XMLDom.appendChild(XMLDom.createElement("IndivGroup"))
			XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"groupid","")).text=Dv_IndivGroup_MainClass.ID
			XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"powerflag","")).text=Dv_IndivGroup_MainClass.PowerFlag
			XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"action","")).text=LCase(Request("action"))

			If LCase(Request("action"))="usermanage" Then
				Dim Node
				GroupUserID = Dvbbs.CheckNumeric(Request("groupuserid"))
				SQL = "Select ID,GroupID,UserID,UserName,Islock,Intro From [Dv_GroupUser] Where GroupID="&Dv_IndivGroup_MainClass.ID&" And ID="&GroupUserID
				Set Rs=Dv_IndivGroup_MainClass.Execute(SQL)
				Set UserXMLDom = Dvbbs.RecordsetToxml(Rs,"User","UserList")
				Set Node=UserXMLDom.documentElement.selectSingleNode("User")
				If Not Node Is Nothing Then
					Node.selectSingleNode("@intro").text = CodeEnNewline(Node.selectSingleNode("@intro").text)
					XMLDom.documentElement.appendChild(Node)
				End If
			Else
				SQL = "ID,GroupID,UserID,UserName,Islock,Intro"
				Select Case Query
					Case 1:SQL="Select "&SQL&" From [Dv_GroupUser] Where GroupID="&Dv_IndivGroup_MainClass.ID&" And IsLock=0 Order By ID Desc"
					Case 2:SQL="Select "&SQL&" From [Dv_GroupUser] Where GroupID="&Dv_IndivGroup_MainClass.ID&" And IsLock=1 Order By ID Desc"
					Case 3:SQL="Select "&SQL&" From [Dv_GroupUser] Where GroupID="&Dv_IndivGroup_MainClass.ID&" And IsLock=2 Order By ID Desc"
					Case Else
						SQL="Select "&SQL&" From [Dv_GroupUser] Where GroupID="&Dv_IndivGroup_MainClass.ID&" Order By ID Desc"
				End Select

				Dim i,MaxRows,Endpage,CountNum,PageSearch
				Endpage=0:MaxRows=20:CountNum=0
				PageSearch = "managetype=user&query="&Query&"&groupid="&Dv_IndivGroup_MainClass.ID

				Set Rs = Dvbbs.iCreateObject ("adodb.recordset")
				If Not IsObject(Dv_IndivGroup_Conn) Then Dv_IndivGroup_ConnectionDatabase:
				Rs.Open SQL,Dv_IndivGroup_Conn,1,1
				If Not Rs.eof Then
					CountNum = Rs.RecordCount
					If CountNum Mod MaxRows=0 Then
						Endpage = CountNum \ MaxRows
					Else
						Endpage = CountNum \ MaxRows+1
					End If
					Rs.MoveFirst
					If Dv_IndivGroup_MainClass.Page > Endpage Then Dv_IndivGroup_MainClass.Page = Endpage
					If Dv_IndivGroup_MainClass.Page >1 Then Rs.Move (Dv_IndivGroup_MainClass.Page-1) * MaxRows
					SQL=Rs.GetRows(MaxRows)
					Set UserXMLDom=Dvbbs.ArrayToxml(SQL,Rs,"User","UserList")
					Rs.close:Set Rs = Nothing
					XMLDom.documentElement.appendChild(UserXMLDom.documentElement.cloneNode(True))
				End If
				'插入分页信息
				XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"MaxRows","")).text=MaxRows
				XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"CountNum","")).text=CountNum
				XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"PageSearch","")).text=PageSearch
			End If
			XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"Page","")).text=Dv_IndivGroup_MainClass.Page

			Set XSLTemplate=Dvbbs.iCreateObject("Msxml2.XSLTemplate" & MsxmlVersion )
			Set XMLStyle=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			'Response.clear:Response.write XMLDom.xml:Response.end
			XMLStyle.loadxml template.html(5)
			'XMLStyle.Load Server.MapPath("Dv_Plus/IndivGroup/Skin/Manage.xslt")

			XSLTemplate.stylesheet=XMLStyle
			Set proc = XSLTemplate.createProcessor()
			proc.input = XMLDom
			proc.transform()
			Response.Write  proc.output
			Set XMLDom=Nothing
			Set proc=Nothing
	End Select
End Sub

Sub UpdateData()
	Dim i,InfoStr
	Dim TopicNum,PostNum,TodayNum,LastPost
	Dim GroupUserNum,GroupTopicNum,GroupPostNum,GroupTodayNum

	Set XMLDom=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	XMLDom.appendChild(XMLDom.createElement("IndivGroup"))
	XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"groupid","")).text=Dv_IndivGroup_MainClass.ID
	XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"powerflag","")).text=Dv_IndivGroup_MainClass.PowerFlag
	XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"action","")).text=LCase(Request("action"))

	InfoStr=""
	GroupUserNum=0:GroupTopicNum=0:GroupPostNum=0:GroupTodayNum=0

	SQL = "Select ID From Dv_Group_Board Where RootID="&Dv_IndivGroup_MainClass.ID
	Set Rs=Dv_IndivGroup_MainClass.Execute(SQL)
	If Not Rs.Eof Then
		SQL = Rs.GetRows(-1)
		Rs.Close:Set Rs=Nothing
		For i=0 To UBound(SQL,2)
			Set Rs=Dv_IndivGroup_MainClass.Execute("Select Count(*) From Dv_Group_Topic Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&SQL(0,i)&" And Istop=0")
			TopicNum = Rs(0):If IsNull(TopicNum) Then TopicNum=0
			GroupTopicNum = GroupTopicNum + TopicNum

			Set Rs=Dv_IndivGroup_MainClass.Execute("Select Count(*) From Dv_Group_BBS Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&SQL(0,i))
			PostNum = Rs(0):If IsNull(PostNum) Then PostNum=0
			GroupPostNum = GroupPostNum + PostNum

			If IsSqlDataBase=1 then
				Set Rs=Dv_IndivGroup_MainClass.Execute("Select Count(*) From Dv_Group_BBS Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&SQL(0,i)&" And Datediff(d,Dateandtime,"&SqlNowString&")=0")
			Else
				Set Rs=Dv_IndivGroup_MainClass.Execute("Select Count(*) From Dv_Group_BBS Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&SQL(0,i)&" And Datediff('d',Dateandtime,"&SqlNowString&")=0")
			End if
			TodayNum = Rs(0):If IsNull(TodayNum) Then TodayNum=0
			GroupTodayNum = GroupTodayNum + TodayNum

			Set Rs=Dv_IndivGroup_MainClass.Execute("Select Top 1 T.title,B.Announceid,B.Dateandtime,B.Username,B.Postuserid,B.Rootid From Dv_Group_BBS b Inner Join Dv_Group_Topic T On b.rootid=T.TopicID Where B.BoardID="&SQL(0,i)&" Order By B.Announceid Desc")
			If Not(Rs.Eof And Rs.Bof) Then
				LastPost = Rs(3)&"$"&Rs(1)&"$"&Rs(2)&"$"&Dv_IndivGroup_MainClass.cutStr(Replace(Rs(0),"$","&#36;"),25)&"$0$"&Rs(4)&"$"&Rs(5)&"$"&SQL(0,i)
			Else
				LastPost = "无$0$"&now()&"$无$0$0$0$"&SQL(0,i)
			End If
			LastPost = Dvbbs.CheckStr(LastPost)
			Dv_IndivGroup_MainClass.Execute("Update Dv_Group_Board Set TopicNum="&TopicNum&",PostNum="&PostNum&",TodayNum="&TodayNum&",LastPost='"&LastPost&"' Where RootID="&Dv_IndivGroup_MainClass.ID&" And ID="&SQL(0,i))
		Next
	End If
	Rs.Close:Set Rs=Nothing

	SQL = "Select Count(*) From Dv_GroupUser Where GroupID="&Dv_IndivGroup_MainClass.ID&" And IsLock>0"
	Set Rs=Dv_IndivGroup_MainClass.Execute(SQL)
	GroupUserNum = Rs(0):If IsNull(GroupUserNum) Then GroupUserNum=0
	Dv_IndivGroup_MainClass.Execute("Update Dv_GroupName Set UserNum="&GroupUserNum&",TopicNum="&GroupTopicNum&",PostNum="&GroupPostNum&",TodayNum="&GroupTodayNum&" Where ID="&Dv_IndivGroup_MainClass.ID)
	InfoStr="圈子“"&Dv_IndivGroup_MainClass.Name&"”数据更新成功。"
	XMLDOM.documentElement.attributes.setNamedItem(XMLDOM.createNode(2,"InfoStr","")).text=InfoStr

	Set XSLTemplate=Dvbbs.iCreateObject("Msxml2.XSLTemplate" & MsxmlVersion )
	Set XMLStyle=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	XMLStyle.loadxml template.html(5)
	'XMLStyle.Load Server.MapPath("Dv_Plus/IndivGroup/Skin/Manage.xslt")

	XSLTemplate.stylesheet=XMLStyle
	Set proc = XSLTemplate.createProcessor()
	proc.input = XMLDom
	proc.transform()
	Response.Write  proc.output
	Set XMLDom=Nothing
	Set proc=Nothing
End Sub

Function NewlineEnCode(Str)
	If Not IsNull(Str) Then
		Str = Replace(Str, CHR(13), "")
		Str = Replace(Str, CHR(10), "<BR/>")
		NewlineEnCode = Str
	End If
End Function

Function CodeEnNewline(Str)
	If Not IsNull(Str) Then
		Str = Replace(Str, "<BR/>", CHR(10))
		CodeEnNewline = Str
	End If
End Function
%>