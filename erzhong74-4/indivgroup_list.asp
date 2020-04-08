<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="Dv_plus/IndivGroup/Dv_IndivGroup_Config.asp"-->
<!--#include file="Dv_plus/IndivGroup/Dv_IndivGroup_MainCls.asp"-->
<%
Dim Rs,Sql,i,Action
Const IndivGroup_EachPageCount=20 '每页显示圈子个数
Dvbbs.LoadTemplates("indivgroup")
Action=Lcase(Request("action"))
Select Case Action
	Case "grouplist"
		Show_GroupList()
	Case "appgroup"
		saveappgroup()
	Case "appjiongroup"
		AppJionGroup()
	Case Else
		Show_GroupList()
End Select
Dvbbs.PageEnd()

'圈子查询列表显示
Sub Show_GroupList()
	Dim XMLDom,ForumNode,IndivGroupXMLDom,AppPowerFlag,AppJionFlag,OrderName
	Dim IndivGroupTotal,TotalRec,Page,PageCount,Orders,keyword,SQLQueryStr
	Dim	XSLTemplate,XMLStyle,proc
	Set XMLDom=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	XMLDom.appendChild(XMLDom.createElement("xml"))
	XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"action","")).text=Action
	
	Dim ClassId
	ClassId=Dvbbs.CheckNumeric(Request("ClassId"))

	Page=Dvbbs.CheckNumeric(Request("page"))
	If Page=0 Then Page=1
	Orders = Dvbbs.CheckNumeric(Request("orders"))
	If Orders=0 Then Orders=3
	keyword=Dvbbs.CheckStr(Request("keyword"))

	Set ForumNode=XMLDom.createNode(1,"Forum","")
	ForumNode.attributes.setNamedItem(XMLDom.createNode(2,"Page","")).text=Page
	ForumNode.attributes.setNamedItem(XMLDom.createNode(2,"pagesize","")).text=IndivGroup_EachPageCount
	ForumNode.attributes.setNamedItem(XMLDom.createNode(2,"Orders","")).text=Orders
	ForumNode.attributes.setNamedItem(XMLDom.createNode(2,"keyword","")).text=keyword
	ForumNode.attributes.setNamedItem(XMLDom.createNode(2,"classid","")).text=ClassId Rem 分类ID

	If Dvbbs.UserID=0 Then 
		AppPowerFlag=1
		AppJionFlag=1
	Else
		If Dvbbs.GroupSetting(73)=1 Then AppPowerFlag=0 Else AppPowerFlag=2
		AppJionFlag=0
	End if
	ForumNode.attributes.setNamedItem(XMLDom.createNode(2,"AppPowerFlag","")).text=AppPowerFlag
	ForumNode.attributes.setNamedItem(XMLDom.createNode(2,"AppJionFlag","")).text=AppJionFlag

	If Action="mygrouplist" Then
		'查询我参与的圈子
		SQL = "Select G.id,G.groupname,G.appuserid,G.appusername,G.usernum,G.stats,G.postnum,G.topicnum,G.todaynum,G.limituser,G.PassDate,U.IsLock From Dv_GroupUser U Inner Join Dv_GroupName G On U.GroupID=G.ID Where U.UserID="&Dvbbs.UserID&" Order By U.ID Desc"
	ElseIf Action="usergrouplist" Then
		Dim UserID:UserID=Dvbbs.CheckNumeric(Request("userid"))
		SQL = "Select G.id,G.groupname,G.appuserid,G.appusername,G.usernum,G.stats,G.postnum,G.topicnum,G.todaynum,G.limituser,G.PassDate From Dv_GroupUser U Inner Join Dv_GroupName G On U.GroupID=G.ID Where U.UserID="&UserID&" Order By U.ID Desc"
	Else
		SQL="id,ClassId,groupname,appuserid,appusername,usernum,stats,postnum,topicnum,todaynum,limituser,PassDate"
		If keyword<>"" Then
			SQLQueryStr = "Where Stats>0 And GroupName like '%"&keyword&"%'"
		Else
			SQLQueryStr = "Where Stats>0"
		End If
		If ClassId>0 Then SQLQueryStr = SQLQueryStr & " And ClassId="&ClassId
		Select Case Orders
			Case 1
				SQL="Select "&SQL&" From [Dv_GroupName] "&SQLQueryStr&" Order By PostNum Desc"
			Case 2
				SQL="Select "&SQL&" From [Dv_GroupName] "&SQLQueryStr&" Order By UserNum Desc"
			Case 3
				SQL="Select "&SQL&" From [Dv_GroupName] "&SQLQueryStr&" Order By PassDate Desc"
		End Select
	End If
	Dv_IndivGroup_MainClass.Stats="圈子列表"
	Dvbbs.Nav()
	Dv_IndivGroup_MainClass.Head_var 2,template.Strings(0),"?action=grouplist"

	'Set Rs = Dv_IndivGroup_MainClass.Execute("Select Count(*) From Dv_GroupName Where Stats>0")
	'If IsNull(Rs(0)) Then IndivGroupTotal=0 Else IndivGroupTotal=Rs(0)
	'Rs.Close:Set Rs=Nothing


	Set Rs=Dvbbs.iCreateObject("ADODB.RecordSet")
	If Not IsObject(Dv_IndivGroup_Conn) Then Dv_IndivGroup_ConnectionDatabase
	Rs.Open SQL,Dv_IndivGroup_Conn,1,1
	Dvbbs.SqlQueryNum = Dvbbs.SqlQueryNum + 1

	TotalRec=0 : PageCount=0
	If Not Rs.Eof Then
		TotalRec=Rs.RecordCount
		If TotalRec Mod IndivGroup_EachPageCount=0 Then
			PageCount= TotalRec\IndivGroup_EachPageCount
		Else
			PageCount= TotalRec\IndivGroup_EachPageCount + 1
		End If
		If Page > PageCount Then Page = PageCount
		If Page > 1 Then Rs.Move (page-1) * IndivGroup_EachPageCount
		
		SQL=Rs.GetRows(IndivGroup_EachPageCount)
		Set IndivGroupXMLDom=Dvbbs.ArrayToxml(SQL,Rs,"row","IndivGroup")
		XMLDom.documentElement.appendChild(IndivGroupXMLDom.documentElement)
	End If
	ForumNode.attributes.setNamedItem(XMLDom.createNode(2,"IndivGroupTotal","")).text=TotalRec
	SQL=Empty:Rs.Close

	If Dvbbs.UserID>0 Then
		Dim AudUserGroup,UserGroup,Node
		AudUserGroup = ""
		UserGroup = Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usergroup").text
		Set Rs=Dv_IndivGroup_Conn.Execute("Select IsLock From Dv_GroupUser Where UserID="&Dvbbs.UserID&" And IsLock=0")
		If Not Rs.Eof Then AudUserGroup=","&Rs.GetString(, -1, "", ",", "")&","
		Rs.Close:Set Rs=Nothing
		For Each Node In XMLDom.DocumentElement.SelectNodes("/xml/IndivGroup/row")
			If Instr(UserGroup,","&Node.GetAttribute("id")&",")>0 Then Node.SetAttribute "islock",2
			If AudUserGroup<>"" Then
				If Instr(AudUserGroup,","&Node.GetAttribute("id")&",")>0 Then Node.SetAttribute "islock",1
			End If
		Next
	End If
	XMLDom.documentElement.appendChild(ForumNode)
	Rem 插入圈子分类数据
	XMLDom.documentElement.appendChild(Application(Dvbbs.CacheName & "_Dv_Group_Class").documentElement.cloneNode(True))

	Set XSLTemplate=Dvbbs.iCreateObject("Msxml2.XSLTemplate" & MsxmlVersion )
	Set XMLStyle=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion )
	XMLStyle.loadxml template.html(0)
	'XMLStyle.Load Server.MapPath("Dv_Plus/IndivGroup/Skin/List.xslt")
	XSLTemplate.stylesheet=XMLStyle
	Set proc = XSLTemplate.createProcessor()
	proc.input = XMLDom
	proc.transform()
	Response.Write  proc.output
	'XMLDom.save Server.MapPath("IndivGroup_List.xml")
	Set XMLDom=Nothing
	Set proc=Nothing

	Dvbbs.ActiveOnline()
	Dvbbs.Footer()
End Sub

'保存圈子信息
Sub saveappgroup()
	Dim UserID,UserName,GroupName,GroupInfo,GroupSetting,viewflag,ClassId
	Dim i,errflag
	UserID = Dvbbs.UserID
	UserName = Dvbbs.MemberName
	errflag = ""
	If UserID=0 Or UserName="" Then	errflag = errflag & "$你还没有登录，不能申请圈子"
	If Dvbbs.GroupSetting(73)=0 Then errflag = errflag & "$你没有申请圈子的权限，请联系管理员咨询"
	GroupName = Dvbbs.CheckStr(Dvbbs.Replacehtml(Request("GroupName")))
	GroupInfo = Dvbbs.CheckStr(NewlineEnCode(Request("GroupInfo")))
	GroupSetting = Dvbbs.CheckNumeric(Request("groupsetting"))
	viewflag = Dvbbs.CheckNumeric(Request("viewflag"))
	ClassId = Dvbbs.CheckNumeric(Request("ClassId"))
	If GroupName="" Then
		errflag = errflag & "$圈子名称没有填写"
	End if
	Response.write "<script langauge=""javascript"">"
	If errflag<>"" Then
		errflag = split(errflag,"$")
		Response.write "parent.information('错误：\n"
		For i=1 To UBound(errflag)
			Response.write "1、"&errflag(i)&"\n"
		Next
		Response.write "');"
	Else
		Set Rs=Dv_IndivGroup_MainClass.Execute("Select GroupName,AppUserName From Dv_GroupName Where AppUserName='"&UserName&"' Or GroupName='"&GroupName&"'")
		If Not Rs.Eof Then
			If UserName=Rs(1) Then
				Response.write "parent.information('你已经申请过圈子了，一个用户只能申请一个圈子。');"
			Else
				Response.write "parent.information('圈子名称已经有人申请，请重新填写。');"
			End If
		Else
			Rs.Close
			If Dvbbs.forum_setting(102)=0 Then 
				Dv_IndivGroup_MainClass.Execute("Insert Into Dv_GroupName(ClassId,GroupName,GroupInfo,AppUserID,AppUserName,UserNum,Stats,LimitUser,AppDate,PassDate,Locked,viewflag) Values("&ClassId&",'"&GroupName&"','"&GroupInfo&"','"&UserID&"','"&UserName&"',1,1,"&Dvbbs.GroupSetting(74)&","&SqlNowString&","&SqlNowString&","&GroupSetting&","&viewflag&")")
				
				Dim GroupID
				Set Rs = Dv_IndivGroup_MainClass.Execute("Select Top 1 ID,AppUserID,AppUserName From Dv_GroupName Order By ID Desc")
				GroupID = Dvbbs.CheckNumeric(Rs(0))
				Rs.Close
				Dv_IndivGroup_MainClass.Execute("Insert Into Dv_GroupUser(GroupID,UserID,UserName,IsLock) Values("&GroupID&","&UserID&",'"&UserName&"',2)")

				Set Rs=Dvbbs.Execute("Select UserGroup From Dv_User Where UserID="&UserID)
				If Not Rs.Eof Then
					Dim UserGroup
					UserGroup=Rs(0)
					If IsNull(UserGroup) Or UserGroup="" Then
						UserGroup = ","&GroupID&","
					Else
						UserGroup = UserGroup&GroupID&","
					End If
					Dvbbs.Execute("Update Dv_User Set UserGroup='"&UserGroup&"' Where UserID="&UserID)
				End If
				Response.write "parent.location.href='IndivGroup_index.asp?groupid="&GroupID&"';"
				Rs.Close:Set Rs=Nothing
			Else
				Dv_IndivGroup_MainClass.Execute("Insert Into Dv_GroupName(ClassId,GroupName,GroupInfo,AppUserID,AppUserName,UserNum,Stats,LimitUser,AppDate,PassDate,Locked,viewflag) Values("&ClassId&",'"&GroupName&"','"&GroupInfo&"','"&UserID&"','"&UserName&"',0,0,"&Dvbbs.GroupSetting(74)&","&SqlNowString&","&SqlNowString&","&GroupSetting&","&viewflag&")")
				Response.write "parent.information('申请信息已经保存，请等待管理员通过');"
			End If
		End If
	End If
	Response.write "</script>"
End Sub

Sub AppJionGroup()
	Dim GroupID,GroupName,Setting,IsLock
	GroupID = Dvbbs.CheckNumeric(Request("groupid"))
	Set Rs=Dv_IndivGroup_MainClass.Execute("Select GroupName,Locked From Dv_GroupName Where ID="&GroupID)
	If Rs.Eof Then Response.write "<script langauge=""javascript"">parent.information('你要加入的圈子不存在，或已经被删除，');</script>":Exit Sub
	GroupName = Rs(0)
	Setting = Rs(1)
	Rs.Close:Set Rs=Nothing

	Response.write "<script langauge=""javascript"">"
	Set Rs=Dv_IndivGroup_MainClass.Execute("Select IsLock From Dv_GroupUser Where GroupID="&GroupID&" And UserID="&Dvbbs.UserID)
	If Not Rs.Eof Then
		If Rs(0)=0 Then
			Response.write "parent.information('您已经申请加入过 """&GroupName&""" 圈子，请等待管理人员审核。');"
		Else
			Response.write "parent.information('您已经是 """&GroupName&""" 圈子的成员了。');"
		End If
	Else
		If Setting=1 Then IsLock=0 Else IsLock=1
		Dv_IndivGroup_MainClass.Execute("Insert Into Dv_GroupUser(GroupID,UserID,UserName,IsLock) Values("&GroupID&","&Dvbbs.UserID&",'"&Dvbbs.MemberName&"',"&IsLock&")")
		If Setting=0 Then
			Dv_IndivGroup_MainClass.Execute("Update Dv_GroupName Set UserNum=UserNum+1 Where ID="&GroupID)
			Response.write "parent.information('您成功加入了 """&GroupName&""" 圈子。');"
		Else
			Response.write "parent.information('您申请加入 """&GroupName&""" 圈子的信息已经保存，请等待管理人员受理。');"
		End If
	End If
	Rs.Close:Set Rs=Nothing
	Response.write "</script>"
End Sub

Function NewlineEnCode(Str)
	If Not IsNull(Str) Then
		Str = Replace(Str, CHR(13), "")
		Str = Replace(Str, CHR(10), "<BR/>")
		NewlineEnCode = Str
	End If
End Function
%>