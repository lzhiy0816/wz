<!-- #include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<!--#include file="../inc/dv_clsother.asp"-->
<%
Head()
Dim Admin_Flag
Dim NewsConfigFile
Dim XmlDoc,Node
Dim NewsName,NewsType,Updatetime,Skin_Head,Skin_Main,Skin_Footer,NewsSql
Admin_flag=",1,"
NewsConfigFile = MyDbPath & "Dv_ForumNews/Dv_NewsSetting.config"
NewsConfigFile = Server.MapPath(NewsConfigFile)

If Not Dvbbs.master Or InStr(","&Session("flag")&",",admin_flag)=0 Then
	Errmsg=ErrMsg + "<li>本页面为管理员专用，请<a href=../admin_login.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。</li>"
	Dvbbs_Error()
Else
	Main()
	If FoundErr Then Call Dvbbs_Error()
	Footer()
End If

Sub Main()
%>
<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align="center">
<tr><th colspan="2" height="23">论坛首页调用管理</th></tr>
<tr>
<td width="20%" class="forumrow" align="center">
<button Style="width:80;height:50;border: 1px outset ;">注意事项</button>
</td>
<td width="80%" class="forumrowHighlight">
	①添加调用后，在列表中点击相应的预览可以看到效果，将调用代码复制到你的首页就可以了。
	<br>②如果你的首页是和论坛程序分开，在填写调用模板时建议用上绝对地址路径。
	<br>③若需要设置外部调用限制和设置临时文件名，修改Dv_News.asp文件，文件里附有说明。
	<br>④建议根据不同的调用设定更新时间间隔，如不是经常更新的版块调用可以设置长一些时间间隔，这样可以有效地减低消耗。
</td>
</tr>
<tr><td colspan="2" class="forumrowHighlight">
<a href="?Act=AddSetting">添加首页调用</a> | <a href="?Act=NewsList">首页调用列表</a> | <a href="<%=MyDbPath%>Dv_News_Demo.asp" target="_blank">查看所有调用演示</a>
</td></tr>
</table>
<%
	Select Case Request("Act")
		Case "NewsList": Call NewsList()
		Case "AddSetting" , "EditNewsInfo" : Call AddSetting()
		Case "SaveSetting" , "SaveEditSetting" : Call SaveSetting()
		Case "DelNewsInfo" : Call DelNewsInfo()
		Case Else
		Call NewsList()
	End Select
End Sub

'删除记录
Sub DelNewsInfo()
	Dim DelNodes,DelChildNodes
	Set XmlDoc = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	If Not XmlDoc.load(NewsConfigFile) Then
		ErrMsg = "调用列表中为空，请填写调用后再执行本操作!"
		Dvbbs_Error()
		Exit Sub
	End If
	'Response.Write Request.Form("DelNodes").count
	For Each DelNodes in Request.Form("DelNodes")
		Set DelChildNodes = XmlDoc.DocumentElement.selectSingleNode("NewsCode[@AddTime='"&DelNodes&"']")
		If Not (DelChildNodes is nothing) Then
			XmlDoc.DocumentElement.RemoveChild(DelChildNodes)
		End If
	Next
	Call SaveXml()
	Dv_suc("所选的记录已删除!")
End Sub

Sub SaveSetting()
	NewsName	= Replace(Request.Form("NewsName"),"""","")
	NewsType	= Replace(Request.Form("NewsType"),"""","")
	Updatetime	= Dvbbs.CheckNumeric(Request.Form("Updatetime"))
	Skin_Head	= Request.Form("Skin_Head")
	Skin_Main	= Request.Form("Skin_Main")
	Skin_Footer	= Request.Form("Skin_Footer")

	If NewsName="" Then
		Errmsg=ErrMsg + "<li>请填写调用标识！</li>"
	Else
		NewsName = Lcase(NewsName)
	End If
	If NewsType < "1" Then
		Errmsg=ErrMsg + "<li>选取调用类型！</li>"
	End If
	If Skin_Main = "" Then
		Errmsg=ErrMsg + "<li>模板_主体循环标记部分不能为空！</li>"
	End If
	If Errmsg<>"" Then Dvbbs_Error() : Exit Sub
	Call LoadXml()

	If FoundNewsName(NewsName) and Request("Act") <> "SaveEditSetting" Then
		Errmsg=ErrMsg + "<li>调用标识已存在，不能重复添加！</li>"
		Dvbbs_Error()
		Exit Sub
	End If
	Select Case NewsType
		Case "1"		'帖子调用
			Call NewsType_1()
		Case "2"		'信息调用
			Call NewsType_2()
		Case "3"		'版块调用
			Call NewsType_3()
		Case "4"		'会员调用
			Call NewsType_4()
		Case "5"		'公告调用
			Call NewsType_5()
		Case "6"		'展区调用
			Call NewsType_6()
		Case Else
			Errmsg=ErrMsg + "<li>请正确选取调用类型！</li>"
			Dvbbs_Error()
	End Select
	Call CreateXmlLog()
	Call SaveXml()
	Dv_suc("调用设置成功!")
End Sub

Sub LoadXml()
	Set XmlDoc = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	If Not XmlDoc.load(NewsConfigFile) Then
		XmlDoc.loadxml "<?xml version=""1.0"" encoding=""gb2312""?><NewscodeInfo/>"
	End If
End Sub

'检查是否存在相同的标识
Function FoundNewsName(NewsName)
	Dim Test
	Set Test = XmlDoc.DocumentElement.selectSingleNode("NewsCode[@NewsName="""&NewsName&"""]")
	FoundNewsName = not (Test is nothing)
End Function

Sub SaveXml()
	XmlDoc.save NewsConfigFile
	Set XmlDoc = Nothing
End Sub

'公共记录
Sub CreateXmlLog()
	Dim attributes,createCDATASection,ChildNode
	Dim FormName,NoAttFormName
	Dim Addtime
	AddTime = Now()
	If Request("Act") = "SaveEditSetting" and Request.Form("AddTime")<>"" Then
		Set Node = XmlDoc.DocumentElement.selectSingleNode("NewsCode[@AddTime='"&Request.Form("AddTime")&"']")
		If Not (Node is nothing) Then
			AddTime = Node.getAttribute("AddTime")
			XmlDoc.DocumentElement.RemoveChild(Node)
		End If
	End If
	'创建节点
	Set Node=XmlDoc.createNode(1,"NewsCode","")
	NoAttFormName = ",Skin_Head,Skin_Main,Skin_Footer,Act,AddTime,Board_Input0,Board_Input1,Board_Input2,Board_Input3,Board_Input4,"
	For Each FormName In Request.Form
		If Instr(NoAttFormName,","&FormName&",")=0 Then
			Set attributes=XmlDoc.createAttribute(FormName)
			If FormName="NewsName" Then
				attributes.text = Lcase(Replace(Request.Form(FormName),"""",""))
			Else
				attributes.text = Replace(Request.Form(FormName),"""","")
			End If
			node.attributes.setNamedItem(attributes)
		End If
	Next
	Set attributes=XmlDoc.createAttribute("MasterName")
	attributes.text = Dvbbs.Membername
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDoc.createAttribute("MasterUserID")
	attributes.text = Dvbbs.UserID
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDoc.createAttribute("MasterIP")
	attributes.text = Dvbbs.UserTrueIP
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDoc.createAttribute("AddTime")
	attributes.text = AddTime
	node.attributes.setNamedItem(attributes)
	Set attributes=XmlDoc.createAttribute("LastTime")
	attributes.text = DateAdd("s", -Updatetime,now())
	node.attributes.setNamedItem(attributes)
	Set ChildNode = XmlDoc.createNode(1,"Search","")
	Set createCDATASection=XmlDoc.createCDATASection(replace(NewsSql,"]]>",""))
	ChildNode.appendChild(createCDATASection)
	node.appendChild(ChildNode)
	Set ChildNode = XmlDoc.createNode(1,"Skin_Head","")
	Set createCDATASection=XmlDoc.createCDATASection(replace(Skin_Head,"]]>","]]&gt;"))
	ChildNode.appendChild(createCDATASection)
	node.appendChild(ChildNode)
	Set ChildNode = XmlDoc.createNode(1,"Skin_Main","")
	Set createCDATASection=XmlDoc.createCDATASection(replace(Skin_Main,"]]>","]]&gt;"))
	ChildNode.appendChild(createCDATASection)
	node.appendChild(ChildNode)
	Set ChildNode = XmlDoc.createNode(1,"Skin_Footer","")
	Set createCDATASection=XmlDoc.createCDATASection(replace(Skin_Footer,"]]>","]]&gt;"))
	ChildNode.appendChild(createCDATASection)
	node.appendChild(ChildNode)
	''特殊版面增加
	If NewsType = "3" Then
		Set ChildNode = XmlDoc.createNode(1,"Board_Input0","")
		Set createCDATASection=XmlDoc.createCDATASection(Replace(Request.Form("Board_Input0"),"]]>","]]&gt;"))
		ChildNode.appendChild(createCDATASection)
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Board_Input1","")
		Set createCDATASection=XmlDoc.createCDATASection(Replace(Request.Form("Board_Input1"),"]]>","]]&gt;"))
		ChildNode.appendChild(createCDATASection)
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Board_Input2","")
		Set createCDATASection=XmlDoc.createCDATASection(Replace(Request.Form("Board_Input2"),"]]>","]]&gt;"))
		ChildNode.appendChild(createCDATASection)
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Board_Input3","")
		Set createCDATASection=XmlDoc.createCDATASection(Replace(Request.Form("Board_Input3"),"]]>","]]&gt;"))
		ChildNode.appendChild(createCDATASection)
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Board_Input4","")
		Set createCDATASection=XmlDoc.createCDATASection(Replace(Request.Form("Board_Input4"),"]]>","]]&gt;"))
		ChildNode.appendChild(createCDATASection)
		node.appendChild(ChildNode)
	ElseIf NewsType = "6" Then
		Set ChildNode = XmlDoc.createNode(1,"Board_Input0","")
		Set createCDATASection=XmlDoc.createCDATASection(Replace(Request.Form("Board_Input0"),"]]>","]]&gt;"))
		ChildNode.appendChild(createCDATASection)
		node.appendChild(ChildNode)
	End If
	XmlDoc.documentElement.appendChild(node)
End Sub


'帖子调用
Sub NewsType_1()
	Dim News_Total,Topiclen,Orders,TopicType,Boardid,BoardLimit,BoardType,UserIDList,Sdate
	News_Total = Dvbbs.CheckNumeric(Request.Form("Total"))
	Topiclen = Dvbbs.CheckNumeric(Request.Form("Topiclen"))
	Orders = Request.Form("Orders")
	Sdate = Dvbbs.CheckNumeric(Request.Form("Sdate"))
	TopicType = Request.Form("TopicType")
	Boardid = Dvbbs.CheckNumeric(Request.Form("Boardid"))
	BoardLimit = Dvbbs.CheckNumeric(Request.Form("BoardLimit"))
	BoardType = Request.Form("BoardType")
	UserIDList = Request.Form("UserIDList")
	If News_Total = 0 Then News_Total = 10
	Dim OrderBy,Searchstr,SearchBoard,Tempstr
	NewsSql = "SELECT TOP "& News_Total
	If Orders = "3" Then
		OrderBy = " Hits Desc, "
	ElseIf orders = "1" or orders = "2" Then
		OrderBy = " Dateandtime Desc, "
	End If
	'指定版面
	If Boardid>0 Then
		SearchBoard = " AND Boardid = " & Boardid
		If BoardType > "0" Then
			Tempstr = GetChildBoardID(Boardid)
			If BoardType = "2" Then
				Tempstr = Boardid & "," &Tempstr
			End If
			If Tempstr<>"" Then
				Tempstr = Left(Tempstr,InStrRev(Tempstr, ",")-1)
				SearchBoard = " AND Boardid in (" & Tempstr &") "
			End If
		End If
	Else
		Tempstr = Cstr(Boardid)
	End If
	'限制不显示特列版面
	If BoardLimit="1" and Tempstr<>"" Then
		Tempstr = GetBoardid(Tempstr)
		If Boardid<>0 Then
			SearchBoard = " AND Boardid in (" & Tempstr &") "
		Else
			If Tempstr<>"" Then
				SearchBoard = " AND Boardid not in (" & Tempstr &") "
			End If
		End If
	End If
	If SearchBoard<>"" Then
		Searchstr = SearchBoard
	End If
	If UserIDList<>"" Then
		If Instr(UserIDList,",") Then
			If IsNumeric(Replace(UserIDList,",","")) Then
				Searchstr = Searchstr & " AND PostUserID IN ("&UserIDList&")"
			End If
		Else
			UserIDList = Dvbbs.CheckNumeric(UserIDList)
			If UserIDList > 0 Then
				Searchstr = Searchstr & " AND PostUserID = "&UserIDList
			End If
		End If
	End If
	If Sdate>0 Then
		If IsSqlDataBase=1 Then
			Searchstr = Searchstr & " AND Datediff(day,DateAndTime,"&SqlNowString&") < " & Sdate
		Else
			Searchstr = Searchstr & " AND Datediff('d',DateAndTime,"&SqlNowString&") < " & Sdate
		End If
	End If
	If TopicType = 1 Then		'显示精华主题
		If Searchstr<>"" Then
			Searchstr = " Where "&  Mid(Searchstr,InStr(Searchstr, "AND")+3)
		End If
		NewsSql = NewsSql & " PostUserName,Title,Rootid,Boardid,Dateandtime,Announceid,Id,Expression From [Dv_BestTopic] " & Searchstr & "  ORDER BY " & OrderBy & " Id Desc"
	ElseIf TopicType=2 Then		'显示主题和回复
		NewsSql = NewsSql & " UserName,Topic,Rootid,Boardid,Dateandtime,Announceid,Body,Expression From "&Dvbbs.NowUseBBS&" Where not (Boardid in (444,777)) "& Searchstr &" ORDER BY "& OrderBy &" AnnounceID Desc"
	Else		'显示主题
		If Orders = 2 Then OrderBy = " Lastposttime Desc, "
		NewsSql = NewsSql & " PostUserName,Title,Topicid,Boardid,Dateandtime,Topicid,Hits,Expression,LastPost From [Dv_topic] Where not (Boardid in (444,777)) "& Searchstr & " ORDER BY "& OrderBy &" Topicid Desc"
	End If
End Sub

'信息调用
Sub NewsType_2()
End Sub

'版块调用
Sub NewsType_3()
End Sub

'会员调用
Sub NewsType_4()
	Dim News_Total,Orders
	News_Total = Dvbbs.CheckNumeric(Request.Form("Total"))
	Orders = Request.Form("Orders")
	Dim OrderBy
	If News_Total = 0 Then News_Total = 10
	NewsSql = "SELECT TOP "& News_Total &" UserID,UserName,UserTopic,UserPost,UserIsBest,UserWealth,UserCP,UserEP,UserDel,UserSex,JoinDate,UserLogins From [Dv_user] "
	Select Case Request.Form("UserOrders")
	Case "0"
		'OrderBy = " JoinDate desc, "
		OrderBy = ""
	Case "1"
		OrderBy = " UserPost desc, "
	Case "2"
		OrderBy = " UserTopic desc, "
	Case "3"
		OrderBy = " UserIsBest desc, "
	Case "4"
		OrderBy = " UserWealth desc, "
	Case "5"
		OrderBy = " UserEP desc, "
	Case "6"
		OrderBy = " UserCP desc, "
	Case "7"
		OrderBy = " UserDel desc, "
	Case "8"
		OrderBy = " UserLogins desc, "
	End Select
	NewsSql = NewsSql & " ORDER BY " & OrderBy & " UserID desc "
End Sub

'公告调用
Sub NewsType_5()
	Dim News_Total,Boardid
	News_Total = Dvbbs.CheckNumeric(Request.Form("Total"))
	Boardid = Dvbbs.CheckNumeric(Request.Form("Boardid"))
	If News_Total = 0 Then News_Total = 10
	NewsSql = "SELECT TOP "& News_Total &" ID,Boardid,Title,UserName,AddTime FROM [Dv_bbsnews] "
	If Boardid > 0 Then
		NewsSql = NewsSql & " WHERE Boardid="& Boardid
	End If
	NewsSql = NewsSql & " ORDER BY ID DESC"
End Sub

'展区调用
Sub NewsType_6()
	Dim News_Total,Boardid,FileOrders,BoardLock,FileType,BoardLimit
	Dim Searchstr,OrderBy
	News_Total = Dvbbs.CheckNumeric(Request.Form("Total"))
	Boardid = Dvbbs.CheckNumeric(Request.Form("Boardid"))
	FileOrders = Request.Form("FileOrders")
	BoardLock = Dvbbs.CheckNumeric(Request.Form("BoardLock"))
	FileType = Request.Form("FileType")
	BoardLimit = Dvbbs.CheckNumeric(Request.Form("BoardLimit"))
	If News_Total = 0 Then News_Total = 8
	If FileType<>"all" Then
		FileType = Dvbbs.CheckNumeric(FileType)
		Searchstr = " AND F_Type = "&FileType
	End If

	'指定版面
	Dim SearchBoard
	Dim Rs,Tempstr
	If Boardid > 0 Then
		Select Case BoardLock
		Case 1
			SearchBoard = " AND F_BoardID <> " & Boardid
			Tempstr = "0"
		Case 3,4
			Tempstr = GetChildBoardID(Boardid)
			If BoardLock = 4 Then
				Tempstr = Boardid & "," &Tempstr
			End If
			If TempStr<>"" Then
				Tempstr = Left(Tempstr,InStrRev(Tempstr, ",")-1)
				SearchBoard = " AND F_BoardID in (" & Tempstr &") "
			End If
		Case Else
			SearchBoard = " AND F_BoardID = " & Boardid
		End Select
	Else
		Tempstr = Cstr(Boardid)
	End If

	'限制不显示特列版面
	If BoardLimit="1" and Tempstr<>"" Then
		Tempstr = GetBoardid(Tempstr)
		If Boardid<>0 Then
			If BoardLock = 1 Then
				SearchBoard = " AND F_BoardID in (" & Boardid &","& Tempstr &") "
			Else
				SearchBoard = " AND F_BoardID in (" & Tempstr &") "
			End If
		Else
			If Tempstr<>"" Then
				SearchBoard = " AND F_BoardID not in (" & Tempstr &") "
			End If
		End If
	End If
	Select Case FileOrders
	Case 1
		OrderBy = " F_ViewNum DESC, "
	Case 2
		OrderBy = " F_DownNum DESC, "
	Case 3
		OrderBy = " F_FileSize DESC, "
	Case Else
		OrderBy = ""
	End Select
	Searchstr = Searchstr & SearchBoard
	NewsSql = "SELECT TOP "& News_Total &" F_ID,F_AnnounceID,F_BoardID,F_Username,F_Filename,F_Readme,F_Type,F_FileType,F_AddTime,F_Viewname,F_ViewNum,F_DownNum,F_FileSize FROM [DV_Upfile] WHERE F_Flag<>4 "
	NewsSql = NewsSql & Searchstr & " ORDER BY "& OrderBy &" F_ID DESC"
End Sub

'BoardidVal<>0 取出调用的版面ID，当BoardidVal=0 取出不被调用的版面ID
Function GetBoardid(BoardidVal)
	Dim TempData,Nodelist,Nodes
	If BoardidVal<>"0" Then
		BoardidVal = "," & BoardidVal & ","
	End If

	Set Nodelist = Application(Dvbbs.CacheName&"_boardlist").cloneNode(True).documentElement.getElementsByTagName("board")
	For Each Nodes in Nodelist
		If BoardidVal<>"0" Then
			If Instr(BoardidVal,","&Nodes.attributes.getNamedItem("boardid").text&",") and Nodes.attributes.getNamedItem("hidden").text="0" and Nodes.attributes.getNamedItem("checkout").text="0" Then
				TempData = TempData & Nodes.attributes.getNamedItem("boardid").text &","
			End If
		Else
			If Nodes.attributes.getNamedItem("hidden").text="1" or Nodes.attributes.getNamedItem("checkout").text="1" Then
				TempData = TempData & Nodes.attributes.getNamedItem("boardid").text &","
			End If
		End If
	Next
	If TempData<>"" Then
		GetBoardid = Left(TempData,InStrRev(TempData, ",")-1)
	End If
End Function

'获取下属版块ID
Private Function GetChildBoardID(BoardIDVal)
		Dim TempData,Nodelist,Node
		Set Nodelist = Application(Dvbbs.CacheName&"_boardlist").cloneNode(True).documentElement.getElementsByTagName("board")
		For Each Node in Nodelist
			If Instr(","&Node.attributes.getNamedItem("parentstr").text&",",","&BoardIDVal&",")>0 Then
				TempData = TempData & Node.attributes.getNamedItem("boardid").text &","
			End If
		Next
		GetChildBoardID = TempData
End Function

Sub AddSetting()
	Dim ChildNode,attributes,Action
	Call LoadXml()
	If Request("Act") = "EditNewsInfo" Then
		Set Node = XmlDoc.DocumentElement.selectSingleNode("NewsCode[@AddTime='"&Request("DelNodes")&"']")
		If (Node is nothing) Then
			ErrMsg = "<li>所选取的调用已不存在!</li>"
			Dvbbs_Error()
			Exit Sub
		End If
		Action = "SaveEditSetting"
	Else
		Set Node=XmlDoc.createNode(1,"NewsCode","")
		Set ChildNode = XmlDoc.createNode(1,"Skin_Head","")
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Skin_Main","")
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Skin_Footer","")
		node.appendChild(ChildNode)
		Action = "SaveSetting"
	End If
	'当不是编辑版面调用时创建临时节点
	If NewsType <> "3" or NewsType <> "6" Then
		Set ChildNode = XmlDoc.createNode(1,"Board_Input0","")
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Board_Input1","")
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Board_Input2","")
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Board_Input3","")
		node.appendChild(ChildNode)
		Set ChildNode = XmlDoc.createNode(1,"Board_Input4","")
		node.appendChild(ChildNode)
	End If
	Set XmlDoc = Nothing
	Dim Boardid
	Boardid = "0"
	If Node.getAttribute("Boardid") <> "" Then
		Boardid = Node.getAttribute("Boardid")
	End If
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function BoardJumpListSelect_Admin(boardid,selectname,fristoption,fristvalue,checknopost){
		if (!xslDoc){GetBoardXml(boardxml);
	}
	var sel = 0;
	var sObj = document.getElementById(selectname);
	if (sObj)
	{
		sObj.options[0] =  new Option(fristoption, fristvalue);
		if (xslDoc.parseError){
				if (xslDoc.parseError.errorCode!=0){
					return;
			}
		}
		var nodes = xslDoc.documentElement.getElementsByTagName("board")
		if (nodes)
		{
			for (var i = 0,k = 1;i<nodes.length;i++) {
				var t = nodes[i].getAttribute("boardtype");
				var v = nodes[i].getAttribute("boardid");
				if (v==boardid)
				{
					sel = k;
				}
				if (nodes[i].getAttribute("depth")==0){
					var outtext="╋";
				}
				else
				{
					var outtext="";
					for (var j=0;j<(nodes[i].getAttribute("depth"));j++)
					{
						if (j>0){outtext+=" |"}
						outtext+="  "
					}
					outtext+="├"
				}
				t = outtext + t
				t = t.replace(/<[^>]*>/g, "")
				t = t.replace(/&[^&]*;/g, "")
				if(checknopost==1 && nodes[i].getAttribute("nopost")=='1')
				{
						t+="(不许转移)"
				}
				sObj.options[k++] = new Option(t, v);
			}
			sObj.options[sel].selected = true;
		}
	}
}
//-->
</SCRIPT>
<br>
<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align="center">
<form METHOD=POST ACTION="?Act=<%=Action%>" name="TheForm">
<tr><th colspan="2" height="23">首页调用管理</th></tr>
<tr>
<td width="30%" class="forumrowHighlight" align="right">
调用标识名称：
</td>
<td width="70%" class="forumrow">
<INPUT TYPE="text" NAME="NewsName" size="10" Maxlength="10" onkeyup="OutputNewsCode(this.value);" value="<%=Node.getAttribute("NewsName")%>">(请使用英文或数字设定调用名称,并且是唯一标识.不能超出10个字符)
</td>
</tr>
<tr>
<td width="15%" class="forumrowHighlight" align="right">
调用代码：
</td>
<td width="85%" class="forumrow">
<INPUT TYPE="text" NAME="Newscode" size="70" disabled value="<script src=&quot;Dv_News.asp?GetName=<%=Node.getAttribute("NewsName")%>&quot;></script>">
</td>
</tr>
<tr>
<td class="forumrowHighlight" align="right">
调用说明：
</td>
<td class="forumrow">
<INPUT TYPE="text" NAME="Intro" size="30" Maxlength="30" value="<%=Node.getAttribute("Intro")%>">(提示说明,以作管理区分.不能超出30个字符)
</td>
</tr>
<tr>
<td class="forumrowHighlight" align="right">
调用类型：
</td>
<td class="forumrow">
	<SELECT NAME="NewsType" ID="NewsType" onchange="NewsTypeSel(this.selectedIndex)">
	<option value="0">选取调用类型</option>
	<option value="1">帖子调用</option>
	<option value="2">信息调用</option>
	<option value="3">版块调用</option>
	<option value="4">会员调用</option>
	<option value="5">公告调用</option>
	<option value="6">展区调用</option>
	</SELECT>
</td>
</tr>
<tr>
<td class="forumrowHighlight" align="right">
数据更新间隔：
</td>
<td class="forumrow"><INPUT TYPE="text" NAME="Updatetime" value="<%=Node.getAttribute("Updatetime")%>">(单位：秒)</td>
</tr>
<tr>
<td class="forumrowHighlight" align="right">
时间显示格式：
</td>
<td class="forumrow">
<SELECT NAME="FormatTime" ID="FormatTime">
	<option value="0" SELECTED>YYYY-M-D H:M:S(长格式)</option>
	<option value="1">YYYY年M月D</option>
	<option value="2">YYYY-M-D</option>
	<option value="3">H:M:S</option>
	<option value="4">hh:mm</option>
</SELECT>
(按服务器时间区域格式显示。)
</td>
</tr>

<tr>
<td class="forumrowHighlight" align="right" valign="top">调用设置：</td>
<td class="forumrowHighlight">
<div id="News"></div>
</td>
</tr>
<!-- 调用模板设置 -->
<tr><th colspan="2" height="23">调用模板设置(请用HTML语法填写)</th></tr>
<tr>
<td class="forumrowHighlight" align="right" valign="top">模板_开始标记部分
</td>
<td class="forumrowHighlight">
	<textarea name="Skin_Head" ID="Skin_Head" style="width:100%;" rows="3"><%=Server.Htmlencode(Node.selectSingleNode("Skin_Head").text&"")%></textarea>
	<br><a href="javascript:admin_Size(-3,'Skin_Head')"><img src="images/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(3,'Skin_Head')"><img src="images/plus.gif" unselectable="on" border='0'></a>
</td>
</tr>
<tr>
<td class="forumrowHighlight" align="right" valign="top">
模板_主体循环标记部分
<fieldset title="模板变量">
<legend>&nbsp;模板变量说明&nbsp;</legend>
<div id="skin_info" align="left"></div>
</fieldset>
</td>
<td class="forumrowHighlight" valign="top">
	<div id="DisInput"></div>
	<textarea name="Skin_Main" ID="Skin_Main" style="width:100%;" rows="10"><%=Server.Htmlencode(Node.selectSingleNode("Skin_Main").text&"")%></textarea>
	<br><a href="javascript:admin_Size(-3,'Skin_Main')"><img src="images/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(3,'Skin_Main')"><img src="images/plus.gif" unselectable="on" border='0'></a>
</td>
</tr>
<tr>
<td class="forumrowHighlight" align="right" valign="top">模板_结束标记部分
</td>
<td class="forumrowHighlight">
	<textarea name="Skin_Footer" ID="Skin_Footer" style="width:100%;" rows="3"><%=Server.Htmlencode(Node.selectSingleNode("Skin_Footer").text&"")%></textarea>
	<br><a href="javascript:admin_Size(-3,'Skin_Footer')"><img src="images/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(3,'Skin_Footer')"><img src="images/plus.gif" unselectable="on" border='0'></a>
</td>
</tr>
<!-- 调用模板设置 -->
<tr>
<td class="forumrowHighlight" align="right">
</td>
<td class="forumrowHighlight" align="center">
<INPUT TYPE="submit" value="提交">&nbsp;&nbsp;&nbsp;<INPUT TYPE="reset" value="重填">
<INPUT TYPE="hidden" name="AddTime" value="<%=Node.getAttribute("AddTime")%>">
</td>
</tr>
</form>
</table>
<!-- 设置信息部分 -->
<div id="News_1" style="display:none">
<!-- 帖子调用 -->
<table border="0" cellpadding="3" cellspacing="1" style="width:100%" >
<tr>
<td class="forumrow">
显示记录数：<INPUT TYPE="text" NAME="Total" size="3" value="<%=Node.getAttribute("Total")%>">
</td><td class="forumrow">
标题长度：<INPUT TYPE="text" NAME="Topiclen" size="4" value="<%=Node.getAttribute("Topiclen")%>">
</td>
<td class="forumrow">
帖子排序：<SELECT NAME="Orders" ID="Orders">
	<option value="0" SELECTED>默认最新排序(推荐使用)</option>
	<option value="1">按照时间(按最新主题时间)</option>
	<option value="2">按照时间(按最新回复时间)</option>
	<option value="3">按照点击(最热帖)</option>
	</SELECT>
</td>
</tr>
<tr><td class="forumrow" colspan="3">
天数的限制：<INPUT TYPE="text" NAME="Sdate" value="<%=Node.getAttribute("Sdate")%>" size="3">(查询多少天内帖子，1为当天。若为空则日期不限，建议为空。)
</td></tr>
<tr><td class="forumrow" colspan="3">
显示的类型：<SELECT NAME="TopicType" ID="TopicType">
	<option value="0" SELECTED>显示主题</option>
	<option value="1">显示精华主题</option>
	<option value="2">显示主题和回复</option>
	</SELECT>
	(不推荐数据量大的用户使用调用主题和回复。)
</td>
</tr>
<tr><td class="forumrow" colspan="3">
调用的版面：<SELECT id="Boardid0" NAME="Boardid"></SELECT>
<BR>
版面&nbsp;&nbsp;设置：<SELECT NAME="BoardType" ID="BoardType">
	<option value="0" SELECTED>只显示该版面的数据</option>
	<option value="1">显示该版面的下级所有版面的数据</option>
	<option value="2">显示该版面和下级所有版面的数据</option>
	</SELECT>
<BR>版面的限制：<SELECT NAME="BoardLimit" ID="BoardLimit">
	<option value="0" SELECTED>显示所有数据</option>
	<option value="1">不显示特殊版面数据</option>
	</SELECT>（特殊版面指隐藏版面和认证版面）
</td>
</tr>
<tr><td class="forumrow" colspan="3">
单独用户ID：<INPUT TYPE="text" NAME="UserIDList" value="<%=Node.getAttribute("UserIDList")%>">(请填写用户会员ID,用英文逗号分隔)
</td>
</tr>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
BoardJumpListSelect_Admin('<%=Boardid%>',"Boardid0","选取所有版面","",0);
//-->
</SCRIPT>
</div>
<div id="News_2" style="display:none">
<!-- 信息调用 -->
<table border="0" cellpadding="3" cellspacing="1" style="width:100%" >
<tr>
<td></td>
</tr>
</table>
</div>
<div id="News_3" style="display:none">
<!-- 版块调用 -->
<table border="0" cellpadding="3" cellspacing="1" style="width:100%" >
<tr>
<td class="forumrow">
显示模式：<SELECT NAME="Orders" ID="Orders">
	<option value="0" SELECTED>树型结构</option>
	<option value="1">地图结构</option>
	</SELECT>
</td>
<td class="forumrow">
<input type="text" name="BoardTab" value="<%=Node.getAttribute("BoardTab")%>" size="2">(地图结构时，限制每行显示数量)
</td>
</tr>
<tr>
<td class="forumrow" colspan="2">
限制调用版块的层数：<input type="text" name="Depth" size="2" value="<%=Node.getAttribute("Depth")%>"><BR>(如0,表示只调用第一级分类
;为空则表示调用所有，当地图结构模式时，层数不能超过1;)
</td>
</tr>
<tr>
<td class="forumrow">
调用的版面：<SELECT id="Boardid1" NAME="Boardid"></SELECT>
</td>
<td class="forumrow">
<input type="radio" name="Stats" value="0">显示所有版块
<input type="radio" name="Stats" value="1" checked>不显示隐藏版块
</td>
</tr>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
BoardJumpListSelect_Admin('<%=Boardid%>',"Boardid1","选取所有版面","",0);
//-->
</SCRIPT>
</div>
<div id="News_4" style="display:none">
<!-- 会员调用 -->
<table border="0" cellpadding="3" cellspacing="1" style="width:100%">
<tr>
<td class="forumrow">
显示记录数：<INPUT TYPE="text" NAME="Total" size="3" value="<%=Node.getAttribute("Total")%>">
</td>
<td class="forumrow">
会员排序：<SELECT NAME="UserOrders" ID="UserOrders">
	<option value="0" SELECTED>按注册时间</option>
	<option value="1">按用户文章</option>
	<option value="2">按用户主题</option>
	<option value="3">按用户精华</option>
	<option value="4">按用户金钱</option>
	<option value="5">按用户经验</option>
	<option value="6">按用户魅力</option>
	<option value="7">按用户被删帖数</option>
	<option value="8">按用户登陆次数</option>
	</SELECT>
</td>
</tr>
</table>
</div>
<div id="News_5" style="display:none">
<!-- 公告调用 -->
<table border="0" cellpadding="3" cellspacing="1" style="width:100%">
<tr>
<td class="forumrow">
显示记录数：<INPUT TYPE="text" NAME="Total" value="<%=Node.getAttribute("Total")%>" size="3">
</td><td class="forumrow">
标题长度：<INPUT TYPE="text" NAME="Topiclen" value="<%=Node.getAttribute("Topiclen")%>" size="4">
</td>
</tr>
<tr>
<td class="forumrow" colspan="2">
调用的版面：<SELECT id="Boardid2" NAME="Boardid"></SELECT>
</td>
</tr>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
BoardJumpListSelect_Admin('<%=Boardid%>',"Boardid2","选取所有版面","",0);
//-->
</SCRIPT>
</div>
<div id="News_6" style="display:none">
<!-- 展区调用 -->
<table border="0" cellpadding="3" cellspacing="1" style="width:100%">
<tr>
<td>
显示记录数：<INPUT TYPE="text" NAME="Total" value="<%=Node.getAttribute("Total")%>" size="3">
&nbsp;&nbsp;&nbsp;&nbsp;
每行显示个数：<INPUT TYPE="text" NAME="Tab" value="<%=Node.getAttribute("Tab")%>" size="3">
&nbsp;&nbsp;&nbsp;&nbsp;标题长度：<INPUT TYPE="text" NAME="Topiclen" value="<%=Node.getAttribute("Topiclen")%>" size="4">
<br>
调用的版面：<SELECT id="Boardid3" NAME="Boardid"></SELECT>
版面限制设置：
	<SELECT NAME="BoardLock" ID="BoardLock">
	<option value="0">不限制</option>
	<option value="1">该版面不被调用</option>
	<option value="2">只调用该版面</option>
	<option value="3">该版面的下级版面</option>
	<option value="4">该版及下级所有版面</option>
	</SELECT>
<BR>版面的限制：<SELECT NAME="BoardLimit" ID="BoardLimit">
	<option value="0" SELECTED>显示所有数据</option>
	<option value="1">不显示特殊版面数据</option>
	</SELECT>（特殊版面指隐藏版面和认证版面）
<br>
调用文件类型 ： <SELECT NAME="FileType" ID="FileType">
	<option value="all" SELECTED>所有文件</option>
	<option value="0">文件集</option>
	<option value="1">图片集</option>
	<option value="2">FLASH集</option>
	<option value="3">音乐集</option>
	<option value="4">电影集</option>
	</SELECT>
<br>
显示排序：<SELECT NAME="FileOrders" ID="FileOrders">
	<option value="0" SELECTED>默认</option>
	<option value="1">按浏览次数</option>
	<option value="2">按下载次数</option>
	<option value="3">按文件大小</option>
	</SELECT>
</td>
</tr>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
BoardJumpListSelect_Admin('<%=Boardid%>',"Boardid3","选取所有版面","",0);
//BoardJumpListSelect_Admin(<%=Boardid%>,"Boardid3","选取所有版面","",0);
//-->
</SCRIPT>
</div>
<!-- 变量说明 -->
<div id="skininfo_0" style="display:none"></div>
<div id="skininfo_1" style="display:none">
	<ol>
	
	<li>标题：{$Topic}</li>
	<li>作者：{$UserName}</li>
	<li>发表时间：{$PostTime}</li>
	<li>回复者：{$ReplyName}</li>
	<li>回复时间：{$ReplyTime}</li>
	<li>版块名称：{$BoardName}</li>
	<li>版块说明：{$BoardInfo}</li>
	<li>心情图标：{$Face}</li>
	<li>帖子ID：{$ID}</li>
	<li>帖子ReplyID：{$ReplyID}</li>
	<li>版面ID：{$Boardid}</li>
	</ol>
</div>
<div id="skininfo_2" style="display:none">
	<ol>
	<li>□- 主题总数 ：{$TopicNum}</li>
	<li>□- 论坛贴数 ：{$PostNum}</li>
	<li>□- 注册人数 ：{$JoinMembers}</li>
	<li>□- 论坛在线 ：{$AllOnline}</li>
	<li>□- 新进会员 ：{$LastUser}</li>
	<li>□- 今日帖数 ：{$TodayPostNum}</li>
	<li>□- 昨日贴数 ：{$YesterdayPostNum}</li>
	<li>□- 高峰贴数 ：{$TopPostNum}</li>
	<li>□- 最高在线 ：{$TopOnline}</li>
	<li>□- 建站时间 ：{$BuildDay}</li>
	</ol>
</div>
<div id="skininfo_3" style="display:none">
<ol>
<li>版块ID：{$BoardID}</li>
<li>版块名称：{$BoardName}</li>
<li>版块说明：{$BoardInfo}</li>
<li>版块下级分版数：{$BoardChild}</li>
<li>版块帖子数：{$PostNum}</li>
<li>版块主题数：{$TopicNum}</li>
<li>版块当天发帖数：{$TodayNum}</li>
<li>版块规则说明：{$Rules}</li>
</ol>
</div>
<div id="Board_Input" style="display:none">
<fieldset title="模板变量" style="padding:5px">
<legend>&nbsp;版块特殊设置&nbsp;</legend>
区版块前标识符：<input type="text" name="Board_Input0" value="<%=Server.Htmlencode(Node.selectSingleNode("Board_Input0").text&"")%>" size="40">
<br>
子版块前标识符：<input type="text" name="Board_Input1" value="<%=Server.Htmlencode(Node.selectSingleNode("Board_Input1").text&"")%>" size="40">
<br>
上下级版块间隔：<input type="text" name="Board_Input2" value="<%=Server.Htmlencode(Node.selectSingleNode("Board_Input2").text&"")%>" size="40">
<br>
同级版块间隔：<input type="text" name="Board_Input4" value="<%=Server.Htmlencode(Node.selectSingleNode("Board_Input4").text&"")%>" size="40">
<br>
版块换行：<input type="text" name="Board_Input3" value="<%=Server.Htmlencode(Node.selectSingleNode("Board_Input3").text&"")%>" size="40">
</fieldset>
</div>
<div id="skininfo_4" style="display:none">
	<ol>
	<li>用户ID ：{$UserID}</li>
	<li>用户名 ：{$UserName}</li>
	<li>用户主题数 ：{$UserTopic}</li>
	<li>用户帖子数 ：{$UserPost}</li>
	<li>用户精华数 ：{$UserBest}</li>
	<li>用户金钱 ：{$UserWealth}</li>
	<li>用户魅力 ：{$UserCP}</li>
	<li>用户经验 ：{$UserEP}</li>
	<li>用户被删帖数 ：{$UserDel}</li>
	<li>用户性别 ：{$UserSex}</li>
	<li>用户注册时间 ：{$JoinDate}</li>
	<li>用户登陆次数 ：{$UserLogins}</li>
	</ol>
</div>
<div id="skininfo_5" style="display:none">
	<ol>
	<li>公告ID：{$ID}</li>
	<li>标题：{$Topic}</li>
	<li>作者：{$UserName}</li>
	<li>版块名称：{$BoardName}</li>
	<li>版块ID：{$Boardid}</li>
	<li>时间：{$PostTime}</li>
	</ol>
</div>
<div id="skininfo_6" style="display:none">
	<ol>
	<li>作者：{$UserName}</li>
	<li>版块名称：{$BoardName}</li>
	<li>版块ID：{$Boardid}</li>
	<li>时间：{$AddTime}</li>
	<li>文件ID：{$ID}</li>
	<li>文件名：{$Filename}</li>
	<li>文件说明：{$Readme}</li>
	<li>文件类型：{$FileType}</li>
	<li>文件预览文件名：{$ViewFilename}</li>
	<li>浏览数：{$ViewNum}</li>
	<li>下载数：{$DownNum}</li>
	<li>文件大小：{$FileSize}</li>
	<li>帖子主题ID：{$RootID}</li>
	<li>帖子对应ID：{$ReplyID}</li>
	<li>交替颜色：{$TColor}</li>
	</ol>
</div>
<div id="Show_Input" style="display:none">
<fieldset title="模板变量" style="padding:5px">
<legend>&nbsp;展区特殊设置&nbsp;</legend>
图片宽度：<input type="text" name="PicWidth" value="<%=Node.getAttribute("PicWidth")%>" size="10"> 象素
<br>
图片高度：<input type="text" name="PicHeight" value="<%=Node.getAttribute("PicHeight")%>" size="10"> 象素
<br>
交替颜色1：<input type="text" name="TColor1" id="TColor1" value="<%=Node.getAttribute("TColor1")%>" size="10">
<img border=0 src="../images/post/rect.gif" style="cursor:pointer;background-Color:<%=Node.getAttribute("TColor1")%>;" onclick="Getcolor(this,'TColor1');" title="选取颜色!">
<br>
交替颜色2：<input type="text" name="TColor2" id="TColor1" value="<%=Node.getAttribute("TColor2")%>" size="10">
<img border=0 src="../images/post/rect.gif" style="cursor:pointer;background-Color:<%=Node.getAttribute("TColor2")%>;" onclick="Getcolor(this,'TColor2');" title="选取颜色!">
<br>
换行标记：<input type="text" name="Board_Input0" value="<%=Server.Htmlencode(Node.selectSingleNode("Board_Input0").text&"")%>" size="40">
</fieldset>
</div>
<iframe width="260" height="165" id="colourPalette" src="../images/post/nc_selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;border:1px gray solid" frameborder="0" scrolling="no" ></iframe>
<!-- 变量说明 -->
<!-- 默认模式 -->

<!-- 默认模式 -->
<SCRIPT LANGUAGE="JavaScript">
<!--
function NewsTypeSel(Val){
	var skininfo = document.getElementById("skininfo_"+Val);
	if (skininfo){
		document.getElementById("skin_info").innerHTML = skininfo.innerHTML;
	}
	if (Val>0){
		var News = document.getElementById("News_"+Val);
		document.getElementById("News").innerHTML = News.innerHTML;
	}else{
		document.getElementById("News").innerHTML = "";
	}
	if (Val==3){
	document.getElementById("DisInput").innerHTML = document.getElementById("Board_Input").innerHTML;
	}
	else if(Val==6){
	document.getElementById("DisInput").innerHTML = document.getElementById("Show_Input").innerHTML;
	}
	else
	{
	document.getElementById("DisInput").innerHTML = "";
	}
}

function OutputNewsCode(Val){
	var obj = TheForm.Newscode;
	if (obj){
		if (Val!=''){
			obj.value = "<scr"+"ipt src=\"Dv_News.asp?GetName="+Val.toLowerCase()+"\"><\/scr"+"ipt>";
		}else{
			obj.value = "";
		}
	}
}
//默认值
CheckSel("FormatTime",'<%=Node.getAttribute("FormatTime")%>');
CheckSel("NewsType",'<%=Node.getAttribute("NewsType")%>');
CheckSel("Orders",'<%=Node.getAttribute("Orders")%>');
CheckSel("UserOrders",'<%=Node.getAttribute("UserOrders")%>');
NewsTypeSel(<%=Node.getAttribute("NewsType")%>);
CheckSel("TopicType",'<%=Node.getAttribute("TopicType")%>');
CheckSel("BoardLimit",'<%=Node.getAttribute("BoardLimit")%>');
chkradio(TheForm.Stats,'<%=Node.getAttribute("Stats")%>');
CheckSel("BoardLock",'<%=Node.getAttribute("BoardLock")%>');
CheckSel("FileType",'<%=Node.getAttribute("FileType")%>');BoardType
CheckSel("FileOrders",'<%=Node.getAttribute("FileOrders")%>');
CheckSel("BoardType",'<%=Node.getAttribute("BoardType")%>');
//-->
</SCRIPT>
<%
End Sub

Sub NewsList()
	Set XmlDoc = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	If Not XmlDoc.load(NewsConfigFile) Then
		ErrMsg = "首页调用列表为空，请添加首页调用后再执行本操作!"
		Dvbbs_Error()
		Exit Sub
	End If
	Dim SendLogNode,Childs
	Set SendLogNode = XmlDoc.DocumentElement.SelectNodes("NewsCode")
	Childs = SendLogNode.Length	'列表数
	%>
	<br>
	<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align="center">
	<tr><th colspan="7" height="23">首页调用列表</th></tr>
	<tr>
	<td width="1%" height="23" class=bodytitle align=center nowrap>选取</td>
	<td width="10%" class=bodytitle align=center>类别</td>
	<td width="10%" class=bodytitle align=center>名称</td>
	<td width="*" class=bodytitle align=center nowrap>说明</td>
	<td width="20%" class=bodytitle align=center>添加时间/更新时间</td>
	<td width="20%" class=bodytitle align=center>添加者</td>
	<td width="1%" class=bodytitle align=center>操作</td>
	</tr>
	<form action="?" method="post" name="TheForm">
	<%
	Dim SearchStr,Topic,i
	i=0
	For Each Node in SendLogNode
	%>
	<tr>
	<td class="forumrowHighlight" align=center>
	<INPUT TYPE="checkbox" NAME="DelNodes" value="<%=Node.getAttribute("AddTime")%>">
	</td>
	<td class="forumrow" align=center><%=NewsCodeType(Node.getAttribute("NewsType"))%></td>
	<td class="forumrow" align=center><%=Node.getAttribute("NewsName")%></td>
	<td class="forumrow">
	<%=Node.getAttribute("Intro")%>
	<br><font color="gray">更新时间间隔为：(<font color="red"><%=Node.getAttribute("Updatetime")%></font>) 秒。</font>
	</td>
	<td class="forumrow"><%=Node.getAttribute("AddTime")%><br><font color="red"><%=Node.getAttribute("LastTime")%></font></td>
	<td class="forumrow" align=center><%=Node.getAttribute("MasterName")%><br><font color="gray"><%=Node.getAttribute("MasterIP")%></font></td>
	<td class="forumrow" align=center><input type="submit"  onclick="this.form.Act.value='EditNewsInfo';Selchecked(this.form.DelNodes,<%=i%>);" value="编辑">
	<input type="button" value="预览" onclick="runscript(viewcode,'<%=Node.getAttribute("NewsName")%>');">
	</td>
	</tr>
	<%
	i=i+1
	Next
	%>
	<tr>
		<td colspan="7" class="forumRowHighlight">
		<input type="hidden" name="viewcode" value="">
		<input type="hidden" name="Act" value="DelNewsInfo">
		<input type="submit" name="Submit" value="删除记录"  onclick="{if(confirm('注意：所删除的模版将不能恢复！')){this.form.submit();return true;}return false;}">  <input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">全选</td>
	</tr>
	</form>
	</table>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function Selchecked(obj,n){
		if (obj[n]){
			obj[n].checked=true;
		}else{
			obj.checked=true;
		}
	}
	function runscript(n,Val){
	TheForm.viewcode.value = "<scr"+"ipt src=\"Dv_News.asp?GetName="+Val.toLowerCase()+"\"><\/scr"+"ipt>";
	txtRun=n;	window.open("../Dv_NewsView.asp",'Dv_ViewNews','toolbar=no,location=no,directories=no,status=yes,menubar=yes,scrollbars=yes,resizable=yes,copyhistory=no,width=780,height=450,left=200,top=150');
	}
	//-->
	</SCRIPT>
	</table>
	<%
	Set XmlDoc = Nothing
End Sub


Function NewsCodeType(TypeVal)
	NewsCodeType = "未知"
	Select Case Cstr(TypeVal)
	Case "1"
		NewsCodeType = "帖子"
	Case "2"
		NewsCodeType = "信息"
	Case "3"
		NewsCodeType = "版块"
	Case "4"
		NewsCodeType = "会员"
	Case "5"
		NewsCodeType = "公告"
	Case "6"
		NewsCodeType = "展区"
	End Select
	NewsCodeType = NewsCodeType & "调用"
End Function
%>
