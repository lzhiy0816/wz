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
	Errmsg=ErrMsg + "<li>��ҳ��Ϊ����Աר�ã���<a href=../admin_login.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�</li>"
	Dvbbs_Error()
Else
	Main()
	If FoundErr Then Call Dvbbs_Error()
	Footer()
End If

Sub Main()
%>
<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align="center">
<tr><th colspan="2" height="23">��̳��ҳ���ù���</th></tr>
<tr>
<td width="20%" class="forumrow" align="center">
<button Style="width:80;height:50;border: 1px outset ;">ע������</button>
</td>
<td width="80%" class="forumrowHighlight">
	����ӵ��ú����б��е����Ӧ��Ԥ�����Կ���Ч���������ô��븴�Ƶ������ҳ�Ϳ����ˡ�
	<br>����������ҳ�Ǻ���̳����ֿ�������д����ģ��ʱ�������Ͼ��Ե�ַ·����
	<br>������Ҫ�����ⲿ�������ƺ�������ʱ�ļ������޸�Dv_News.asp�ļ����ļ��︽��˵����
	<br>�ܽ�����ݲ�ͬ�ĵ����趨����ʱ�������粻�Ǿ������µİ����ÿ������ó�һЩʱ����������������Ч�ؼ������ġ�
</td>
</tr>
<tr><td colspan="2" class="forumrowHighlight">
<a href="?Act=AddSetting">�����ҳ����</a> | <a href="?Act=NewsList">��ҳ�����б�</a> | <a href="<%=MyDbPath%>Dv_News_Demo.asp" target="_blank">�鿴���е�����ʾ</a>
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

'ɾ����¼
Sub DelNewsInfo()
	Dim DelNodes,DelChildNodes
	Set XmlDoc = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	If Not XmlDoc.load(NewsConfigFile) Then
		ErrMsg = "�����б���Ϊ�գ�����д���ú���ִ�б�����!"
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
	Dv_suc("��ѡ�ļ�¼��ɾ��!")
End Sub

Sub SaveSetting()
	NewsName	= Replace(Request.Form("NewsName"),"""","")
	NewsType	= Replace(Request.Form("NewsType"),"""","")
	Updatetime	= Dvbbs.CheckNumeric(Request.Form("Updatetime"))
	Skin_Head	= Request.Form("Skin_Head")
	Skin_Main	= Request.Form("Skin_Main")
	Skin_Footer	= Request.Form("Skin_Footer")

	If NewsName="" Then
		Errmsg=ErrMsg + "<li>����д���ñ�ʶ��</li>"
	Else
		NewsName = Lcase(NewsName)
	End If
	If NewsType < "1" Then
		Errmsg=ErrMsg + "<li>ѡȡ�������ͣ�</li>"
	End If
	If Skin_Main = "" Then
		Errmsg=ErrMsg + "<li>ģ��_����ѭ����ǲ��ֲ���Ϊ�գ�</li>"
	End If
	If Errmsg<>"" Then Dvbbs_Error() : Exit Sub
	Call LoadXml()

	If FoundNewsName(NewsName) and Request("Act") <> "SaveEditSetting" Then
		Errmsg=ErrMsg + "<li>���ñ�ʶ�Ѵ��ڣ������ظ���ӣ�</li>"
		Dvbbs_Error()
		Exit Sub
	End If
	Select Case NewsType
		Case "1"		'���ӵ���
			Call NewsType_1()
		Case "2"		'��Ϣ����
			Call NewsType_2()
		Case "3"		'������
			Call NewsType_3()
		Case "4"		'��Ա����
			Call NewsType_4()
		Case "5"		'�������
			Call NewsType_5()
		Case "6"		'չ������
			Call NewsType_6()
		Case Else
			Errmsg=ErrMsg + "<li>����ȷѡȡ�������ͣ�</li>"
			Dvbbs_Error()
	End Select
	Call CreateXmlLog()
	Call SaveXml()
	Dv_suc("�������óɹ�!")
End Sub

Sub LoadXml()
	Set XmlDoc = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	If Not XmlDoc.load(NewsConfigFile) Then
		XmlDoc.loadxml "<?xml version=""1.0"" encoding=""gb2312""?><NewscodeInfo/>"
	End If
End Sub

'����Ƿ������ͬ�ı�ʶ
Function FoundNewsName(NewsName)
	Dim Test
	Set Test = XmlDoc.DocumentElement.selectSingleNode("NewsCode[@NewsName="""&NewsName&"""]")
	FoundNewsName = not (Test is nothing)
End Function

Sub SaveXml()
	XmlDoc.save NewsConfigFile
	Set XmlDoc = Nothing
End Sub

'������¼
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
	'�����ڵ�
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
	''�����������
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


'���ӵ���
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
	'ָ������
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
	'���Ʋ���ʾ���а���
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
	If TopicType = 1 Then		'��ʾ��������
		If Searchstr<>"" Then
			Searchstr = " Where "&  Mid(Searchstr,InStr(Searchstr, "AND")+3)
		End If
		NewsSql = NewsSql & " PostUserName,Title,Rootid,Boardid,Dateandtime,Announceid,Id,Expression From [Dv_BestTopic] " & Searchstr & "  ORDER BY " & OrderBy & " Id Desc"
	ElseIf TopicType=2 Then		'��ʾ����ͻظ�
		NewsSql = NewsSql & " UserName,Topic,Rootid,Boardid,Dateandtime,Announceid,Body,Expression From "&Dvbbs.NowUseBBS&" Where not (Boardid in (444,777)) "& Searchstr &" ORDER BY "& OrderBy &" AnnounceID Desc"
	Else		'��ʾ����
		If Orders = 2 Then OrderBy = " Lastposttime Desc, "
		NewsSql = NewsSql & " PostUserName,Title,Topicid,Boardid,Dateandtime,Topicid,Hits,Expression,LastPost From [Dv_topic] Where not (Boardid in (444,777)) "& Searchstr & " ORDER BY "& OrderBy &" Topicid Desc"
	End If
End Sub

'��Ϣ����
Sub NewsType_2()
End Sub

'������
Sub NewsType_3()
End Sub

'��Ա����
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

'�������
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

'չ������
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

	'ָ������
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

	'���Ʋ���ʾ���а���
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

'BoardidVal<>0 ȡ�����õİ���ID����BoardidVal=0 ȡ���������õİ���ID
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

'��ȡ�������ID
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
			ErrMsg = "<li>��ѡȡ�ĵ����Ѳ�����!</li>"
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
	'�����Ǳ༭�������ʱ������ʱ�ڵ�
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
					var outtext="��";
				}
				else
				{
					var outtext="";
					for (var j=0;j<(nodes[i].getAttribute("depth"));j++)
					{
						if (j>0){outtext+=" |"}
						outtext+="  "
					}
					outtext+="��"
				}
				t = outtext + t
				t = t.replace(/<[^>]*>/g, "")
				t = t.replace(/&[^&]*;/g, "")
				if(checknopost==1 && nodes[i].getAttribute("nopost")=='1')
				{
						t+="(����ת��)"
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
<tr><th colspan="2" height="23">��ҳ���ù���</th></tr>
<tr>
<td width="30%" class="forumrowHighlight" align="right">
���ñ�ʶ���ƣ�
</td>
<td width="70%" class="forumrow">
<INPUT TYPE="text" NAME="NewsName" size="10" Maxlength="10" onkeyup="OutputNewsCode(this.value);" value="<%=Node.getAttribute("NewsName")%>">(��ʹ��Ӣ�Ļ������趨��������,������Ψһ��ʶ.���ܳ���10���ַ�)
</td>
</tr>
<tr>
<td width="15%" class="forumrowHighlight" align="right">
���ô��룺
</td>
<td width="85%" class="forumrow">
<INPUT TYPE="text" NAME="Newscode" size="70" disabled value="<script src=&quot;Dv_News.asp?GetName=<%=Node.getAttribute("NewsName")%>&quot;></script>">
</td>
</tr>
<tr>
<td class="forumrowHighlight" align="right">
����˵����
</td>
<td class="forumrow">
<INPUT TYPE="text" NAME="Intro" size="30" Maxlength="30" value="<%=Node.getAttribute("Intro")%>">(��ʾ˵��,������������.���ܳ���30���ַ�)
</td>
</tr>
<tr>
<td class="forumrowHighlight" align="right">
�������ͣ�
</td>
<td class="forumrow">
	<SELECT NAME="NewsType" ID="NewsType" onchange="NewsTypeSel(this.selectedIndex)">
	<option value="0">ѡȡ��������</option>
	<option value="1">���ӵ���</option>
	<option value="2">��Ϣ����</option>
	<option value="3">������</option>
	<option value="4">��Ա����</option>
	<option value="5">�������</option>
	<option value="6">չ������</option>
	</SELECT>
</td>
</tr>
<tr>
<td class="forumrowHighlight" align="right">
���ݸ��¼����
</td>
<td class="forumrow"><INPUT TYPE="text" NAME="Updatetime" value="<%=Node.getAttribute("Updatetime")%>">(��λ����)</td>
</tr>
<tr>
<td class="forumrowHighlight" align="right">
ʱ����ʾ��ʽ��
</td>
<td class="forumrow">
<SELECT NAME="FormatTime" ID="FormatTime">
	<option value="0" SELECTED>YYYY-M-D H:M:S(����ʽ)</option>
	<option value="1">YYYY��M��D</option>
	<option value="2">YYYY-M-D</option>
	<option value="3">H:M:S</option>
	<option value="4">hh:mm</option>
</SELECT>
(��������ʱ�������ʽ��ʾ��)
</td>
</tr>

<tr>
<td class="forumrowHighlight" align="right" valign="top">�������ã�</td>
<td class="forumrowHighlight">
<div id="News"></div>
</td>
</tr>
<!-- ����ģ������ -->
<tr><th colspan="2" height="23">����ģ������(����HTML�﷨��д)</th></tr>
<tr>
<td class="forumrowHighlight" align="right" valign="top">ģ��_��ʼ��ǲ���
</td>
<td class="forumrowHighlight">
	<textarea name="Skin_Head" ID="Skin_Head" style="width:100%;" rows="3"><%=Server.Htmlencode(Node.selectSingleNode("Skin_Head").text&"")%></textarea>
	<br><a href="javascript:admin_Size(-3,'Skin_Head')"><img src="images/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(3,'Skin_Head')"><img src="images/plus.gif" unselectable="on" border='0'></a>
</td>
</tr>
<tr>
<td class="forumrowHighlight" align="right" valign="top">
ģ��_����ѭ����ǲ���
<fieldset title="ģ�����">
<legend>&nbsp;ģ�����˵��&nbsp;</legend>
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
<td class="forumrowHighlight" align="right" valign="top">ģ��_������ǲ���
</td>
<td class="forumrowHighlight">
	<textarea name="Skin_Footer" ID="Skin_Footer" style="width:100%;" rows="3"><%=Server.Htmlencode(Node.selectSingleNode("Skin_Footer").text&"")%></textarea>
	<br><a href="javascript:admin_Size(-3,'Skin_Footer')"><img src="images/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(3,'Skin_Footer')"><img src="images/plus.gif" unselectable="on" border='0'></a>
</td>
</tr>
<!-- ����ģ������ -->
<tr>
<td class="forumrowHighlight" align="right">
</td>
<td class="forumrowHighlight" align="center">
<INPUT TYPE="submit" value="�ύ">&nbsp;&nbsp;&nbsp;<INPUT TYPE="reset" value="����">
<INPUT TYPE="hidden" name="AddTime" value="<%=Node.getAttribute("AddTime")%>">
</td>
</tr>
</form>
</table>
<!-- ������Ϣ���� -->
<div id="News_1" style="display:none">
<!-- ���ӵ��� -->
<table border="0" cellpadding="3" cellspacing="1" style="width:100%" >
<tr>
<td class="forumrow">
��ʾ��¼����<INPUT TYPE="text" NAME="Total" size="3" value="<%=Node.getAttribute("Total")%>">
</td><td class="forumrow">
���ⳤ�ȣ�<INPUT TYPE="text" NAME="Topiclen" size="4" value="<%=Node.getAttribute("Topiclen")%>">
</td>
<td class="forumrow">
��������<SELECT NAME="Orders" ID="Orders">
	<option value="0" SELECTED>Ĭ����������(�Ƽ�ʹ��)</option>
	<option value="1">����ʱ��(����������ʱ��)</option>
	<option value="2">����ʱ��(�����»ظ�ʱ��)</option>
	<option value="3">���յ��(������)</option>
	</SELECT>
</td>
</tr>
<tr><td class="forumrow" colspan="3">
���������ƣ�<INPUT TYPE="text" NAME="Sdate" value="<%=Node.getAttribute("Sdate")%>" size="3">(��ѯ�����������ӣ�1Ϊ���졣��Ϊ�������ڲ��ޣ�����Ϊ�ա�)
</td></tr>
<tr><td class="forumrow" colspan="3">
��ʾ�����ͣ�<SELECT NAME="TopicType" ID="TopicType">
	<option value="0" SELECTED>��ʾ����</option>
	<option value="1">��ʾ��������</option>
	<option value="2">��ʾ����ͻظ�</option>
	</SELECT>
	(���Ƽ�����������û�ʹ�õ�������ͻظ���)
</td>
</tr>
<tr><td class="forumrow" colspan="3">
���õİ��棺<SELECT id="Boardid0" NAME="Boardid"></SELECT>
<BR>
����&nbsp;&nbsp;���ã�<SELECT NAME="BoardType" ID="BoardType">
	<option value="0" SELECTED>ֻ��ʾ�ð��������</option>
	<option value="1">��ʾ�ð�����¼����а��������</option>
	<option value="2">��ʾ�ð�����¼����а��������</option>
	</SELECT>
<BR>��������ƣ�<SELECT NAME="BoardLimit" ID="BoardLimit">
	<option value="0" SELECTED>��ʾ��������</option>
	<option value="1">����ʾ�����������</option>
	</SELECT>���������ָ���ذ������֤���棩
</td>
</tr>
<tr><td class="forumrow" colspan="3">
�����û�ID��<INPUT TYPE="text" NAME="UserIDList" value="<%=Node.getAttribute("UserIDList")%>">(����д�û���ԱID,��Ӣ�Ķ��ŷָ�)
</td>
</tr>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
BoardJumpListSelect_Admin('<%=Boardid%>',"Boardid0","ѡȡ���а���","",0);
//-->
</SCRIPT>
</div>
<div id="News_2" style="display:none">
<!-- ��Ϣ���� -->
<table border="0" cellpadding="3" cellspacing="1" style="width:100%" >
<tr>
<td></td>
</tr>
</table>
</div>
<div id="News_3" style="display:none">
<!-- ������ -->
<table border="0" cellpadding="3" cellspacing="1" style="width:100%" >
<tr>
<td class="forumrow">
��ʾģʽ��<SELECT NAME="Orders" ID="Orders">
	<option value="0" SELECTED>���ͽṹ</option>
	<option value="1">��ͼ�ṹ</option>
	</SELECT>
</td>
<td class="forumrow">
<input type="text" name="BoardTab" value="<%=Node.getAttribute("BoardTab")%>" size="2">(��ͼ�ṹʱ������ÿ����ʾ����)
</td>
</tr>
<tr>
<td class="forumrow" colspan="2">
���Ƶ��ð��Ĳ�����<input type="text" name="Depth" size="2" value="<%=Node.getAttribute("Depth")%>"><BR>(��0,��ʾֻ���õ�һ������
;Ϊ�����ʾ�������У�����ͼ�ṹģʽʱ���������ܳ���1;)
</td>
</tr>
<tr>
<td class="forumrow">
���õİ��棺<SELECT id="Boardid1" NAME="Boardid"></SELECT>
</td>
<td class="forumrow">
<input type="radio" name="Stats" value="0">��ʾ���а��
<input type="radio" name="Stats" value="1" checked>����ʾ���ذ��
</td>
</tr>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
BoardJumpListSelect_Admin('<%=Boardid%>',"Boardid1","ѡȡ���а���","",0);
//-->
</SCRIPT>
</div>
<div id="News_4" style="display:none">
<!-- ��Ա���� -->
<table border="0" cellpadding="3" cellspacing="1" style="width:100%">
<tr>
<td class="forumrow">
��ʾ��¼����<INPUT TYPE="text" NAME="Total" size="3" value="<%=Node.getAttribute("Total")%>">
</td>
<td class="forumrow">
��Ա����<SELECT NAME="UserOrders" ID="UserOrders">
	<option value="0" SELECTED>��ע��ʱ��</option>
	<option value="1">���û�����</option>
	<option value="2">���û�����</option>
	<option value="3">���û�����</option>
	<option value="4">���û���Ǯ</option>
	<option value="5">���û�����</option>
	<option value="6">���û�����</option>
	<option value="7">���û���ɾ����</option>
	<option value="8">���û���½����</option>
	</SELECT>
</td>
</tr>
</table>
</div>
<div id="News_5" style="display:none">
<!-- ������� -->
<table border="0" cellpadding="3" cellspacing="1" style="width:100%">
<tr>
<td class="forumrow">
��ʾ��¼����<INPUT TYPE="text" NAME="Total" value="<%=Node.getAttribute("Total")%>" size="3">
</td><td class="forumrow">
���ⳤ�ȣ�<INPUT TYPE="text" NAME="Topiclen" value="<%=Node.getAttribute("Topiclen")%>" size="4">
</td>
</tr>
<tr>
<td class="forumrow" colspan="2">
���õİ��棺<SELECT id="Boardid2" NAME="Boardid"></SELECT>
</td>
</tr>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
BoardJumpListSelect_Admin('<%=Boardid%>',"Boardid2","ѡȡ���а���","",0);
//-->
</SCRIPT>
</div>
<div id="News_6" style="display:none">
<!-- չ������ -->
<table border="0" cellpadding="3" cellspacing="1" style="width:100%">
<tr>
<td>
��ʾ��¼����<INPUT TYPE="text" NAME="Total" value="<%=Node.getAttribute("Total")%>" size="3">
&nbsp;&nbsp;&nbsp;&nbsp;
ÿ����ʾ������<INPUT TYPE="text" NAME="Tab" value="<%=Node.getAttribute("Tab")%>" size="3">
&nbsp;&nbsp;&nbsp;&nbsp;���ⳤ�ȣ�<INPUT TYPE="text" NAME="Topiclen" value="<%=Node.getAttribute("Topiclen")%>" size="4">
<br>
���õİ��棺<SELECT id="Boardid3" NAME="Boardid"></SELECT>
�����������ã�
	<SELECT NAME="BoardLock" ID="BoardLock">
	<option value="0">������</option>
	<option value="1">�ð��治������</option>
	<option value="2">ֻ���øð���</option>
	<option value="3">�ð�����¼�����</option>
	<option value="4">�ð漰�¼����а���</option>
	</SELECT>
<BR>��������ƣ�<SELECT NAME="BoardLimit" ID="BoardLimit">
	<option value="0" SELECTED>��ʾ��������</option>
	<option value="1">����ʾ�����������</option>
	</SELECT>���������ָ���ذ������֤���棩
<br>
�����ļ����� �� <SELECT NAME="FileType" ID="FileType">
	<option value="all" SELECTED>�����ļ�</option>
	<option value="0">�ļ���</option>
	<option value="1">ͼƬ��</option>
	<option value="2">FLASH��</option>
	<option value="3">���ּ�</option>
	<option value="4">��Ӱ��</option>
	</SELECT>
<br>
��ʾ����<SELECT NAME="FileOrders" ID="FileOrders">
	<option value="0" SELECTED>Ĭ��</option>
	<option value="1">���������</option>
	<option value="2">�����ش���</option>
	<option value="3">���ļ���С</option>
	</SELECT>
</td>
</tr>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
BoardJumpListSelect_Admin('<%=Boardid%>',"Boardid3","ѡȡ���а���","",0);
//BoardJumpListSelect_Admin(<%=Boardid%>,"Boardid3","ѡȡ���а���","",0);
//-->
</SCRIPT>
</div>
<!-- ����˵�� -->
<div id="skininfo_0" style="display:none"></div>
<div id="skininfo_1" style="display:none">
	<ol>
	
	<li>���⣺{$Topic}</li>
	<li>���ߣ�{$UserName}</li>
	<li>����ʱ�䣺{$PostTime}</li>
	<li>�ظ��ߣ�{$ReplyName}</li>
	<li>�ظ�ʱ�䣺{$ReplyTime}</li>
	<li>������ƣ�{$BoardName}</li>
	<li>���˵����{$BoardInfo}</li>
	<li>����ͼ�꣺{$Face}</li>
	<li>����ID��{$ID}</li>
	<li>����ReplyID��{$ReplyID}</li>
	<li>����ID��{$Boardid}</li>
	</ol>
</div>
<div id="skininfo_2" style="display:none">
	<ol>
	<li>��- �������� ��{$TopicNum}</li>
	<li>��- ��̳���� ��{$PostNum}</li>
	<li>��- ע������ ��{$JoinMembers}</li>
	<li>��- ��̳���� ��{$AllOnline}</li>
	<li>��- �½���Ա ��{$LastUser}</li>
	<li>��- �������� ��{$TodayPostNum}</li>
	<li>��- �������� ��{$YesterdayPostNum}</li>
	<li>��- �߷����� ��{$TopPostNum}</li>
	<li>��- ������� ��{$TopOnline}</li>
	<li>��- ��վʱ�� ��{$BuildDay}</li>
	</ol>
</div>
<div id="skininfo_3" style="display:none">
<ol>
<li>���ID��{$BoardID}</li>
<li>������ƣ�{$BoardName}</li>
<li>���˵����{$BoardInfo}</li>
<li>����¼��ְ�����{$BoardChild}</li>
<li>�����������{$PostNum}</li>
<li>�����������{$TopicNum}</li>
<li>��鵱�췢������{$TodayNum}</li>
<li>������˵����{$Rules}</li>
</ol>
</div>
<div id="Board_Input" style="display:none">
<fieldset title="ģ�����" style="padding:5px">
<legend>&nbsp;�����������&nbsp;</legend>
�����ǰ��ʶ����<input type="text" name="Board_Input0" value="<%=Server.Htmlencode(Node.selectSingleNode("Board_Input0").text&"")%>" size="40">
<br>
�Ӱ��ǰ��ʶ����<input type="text" name="Board_Input1" value="<%=Server.Htmlencode(Node.selectSingleNode("Board_Input1").text&"")%>" size="40">
<br>
���¼��������<input type="text" name="Board_Input2" value="<%=Server.Htmlencode(Node.selectSingleNode("Board_Input2").text&"")%>" size="40">
<br>
ͬ���������<input type="text" name="Board_Input4" value="<%=Server.Htmlencode(Node.selectSingleNode("Board_Input4").text&"")%>" size="40">
<br>
��黻�У�<input type="text" name="Board_Input3" value="<%=Server.Htmlencode(Node.selectSingleNode("Board_Input3").text&"")%>" size="40">
</fieldset>
</div>
<div id="skininfo_4" style="display:none">
	<ol>
	<li>�û�ID ��{$UserID}</li>
	<li>�û��� ��{$UserName}</li>
	<li>�û������� ��{$UserTopic}</li>
	<li>�û������� ��{$UserPost}</li>
	<li>�û������� ��{$UserBest}</li>
	<li>�û���Ǯ ��{$UserWealth}</li>
	<li>�û����� ��{$UserCP}</li>
	<li>�û����� ��{$UserEP}</li>
	<li>�û���ɾ���� ��{$UserDel}</li>
	<li>�û��Ա� ��{$UserSex}</li>
	<li>�û�ע��ʱ�� ��{$JoinDate}</li>
	<li>�û���½���� ��{$UserLogins}</li>
	</ol>
</div>
<div id="skininfo_5" style="display:none">
	<ol>
	<li>����ID��{$ID}</li>
	<li>���⣺{$Topic}</li>
	<li>���ߣ�{$UserName}</li>
	<li>������ƣ�{$BoardName}</li>
	<li>���ID��{$Boardid}</li>
	<li>ʱ�䣺{$PostTime}</li>
	</ol>
</div>
<div id="skininfo_6" style="display:none">
	<ol>
	<li>���ߣ�{$UserName}</li>
	<li>������ƣ�{$BoardName}</li>
	<li>���ID��{$Boardid}</li>
	<li>ʱ�䣺{$AddTime}</li>
	<li>�ļ�ID��{$ID}</li>
	<li>�ļ�����{$Filename}</li>
	<li>�ļ�˵����{$Readme}</li>
	<li>�ļ����ͣ�{$FileType}</li>
	<li>�ļ�Ԥ���ļ�����{$ViewFilename}</li>
	<li>�������{$ViewNum}</li>
	<li>��������{$DownNum}</li>
	<li>�ļ���С��{$FileSize}</li>
	<li>��������ID��{$RootID}</li>
	<li>���Ӷ�ӦID��{$ReplyID}</li>
	<li>������ɫ��{$TColor}</li>
	</ol>
</div>
<div id="Show_Input" style="display:none">
<fieldset title="ģ�����" style="padding:5px">
<legend>&nbsp;չ����������&nbsp;</legend>
ͼƬ��ȣ�<input type="text" name="PicWidth" value="<%=Node.getAttribute("PicWidth")%>" size="10"> ����
<br>
ͼƬ�߶ȣ�<input type="text" name="PicHeight" value="<%=Node.getAttribute("PicHeight")%>" size="10"> ����
<br>
������ɫ1��<input type="text" name="TColor1" id="TColor1" value="<%=Node.getAttribute("TColor1")%>" size="10">
<img border=0 src="../images/post/rect.gif" style="cursor:pointer;background-Color:<%=Node.getAttribute("TColor1")%>;" onclick="Getcolor(this,'TColor1');" title="ѡȡ��ɫ!">
<br>
������ɫ2��<input type="text" name="TColor2" id="TColor1" value="<%=Node.getAttribute("TColor2")%>" size="10">
<img border=0 src="../images/post/rect.gif" style="cursor:pointer;background-Color:<%=Node.getAttribute("TColor2")%>;" onclick="Getcolor(this,'TColor2');" title="ѡȡ��ɫ!">
<br>
���б�ǣ�<input type="text" name="Board_Input0" value="<%=Server.Htmlencode(Node.selectSingleNode("Board_Input0").text&"")%>" size="40">
</fieldset>
</div>
<iframe width="260" height="165" id="colourPalette" src="../images/post/nc_selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;border:1px gray solid" frameborder="0" scrolling="no" ></iframe>
<!-- ����˵�� -->
<!-- Ĭ��ģʽ -->

<!-- Ĭ��ģʽ -->
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
//Ĭ��ֵ
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
		ErrMsg = "��ҳ�����б�Ϊ�գ��������ҳ���ú���ִ�б�����!"
		Dvbbs_Error()
		Exit Sub
	End If
	Dim SendLogNode,Childs
	Set SendLogNode = XmlDoc.DocumentElement.SelectNodes("NewsCode")
	Childs = SendLogNode.Length	'�б���
	%>
	<br>
	<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align="center">
	<tr><th colspan="7" height="23">��ҳ�����б�</th></tr>
	<tr>
	<td width="1%" height="23" class=bodytitle align=center nowrap>ѡȡ</td>
	<td width="10%" class=bodytitle align=center>���</td>
	<td width="10%" class=bodytitle align=center>����</td>
	<td width="*" class=bodytitle align=center nowrap>˵��</td>
	<td width="20%" class=bodytitle align=center>���ʱ��/����ʱ��</td>
	<td width="20%" class=bodytitle align=center>�����</td>
	<td width="1%" class=bodytitle align=center>����</td>
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
	<br><font color="gray">����ʱ����Ϊ��(<font color="red"><%=Node.getAttribute("Updatetime")%></font>) �롣</font>
	</td>
	<td class="forumrow"><%=Node.getAttribute("AddTime")%><br><font color="red"><%=Node.getAttribute("LastTime")%></font></td>
	<td class="forumrow" align=center><%=Node.getAttribute("MasterName")%><br><font color="gray"><%=Node.getAttribute("MasterIP")%></font></td>
	<td class="forumrow" align=center><input type="submit"  onclick="this.form.Act.value='EditNewsInfo';Selchecked(this.form.DelNodes,<%=i%>);" value="�༭">
	<input type="button" value="Ԥ��" onclick="runscript(viewcode,'<%=Node.getAttribute("NewsName")%>');">
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
		<input type="submit" name="Submit" value="ɾ����¼"  onclick="{if(confirm('ע�⣺��ɾ����ģ�潫���ָܻ���')){this.form.submit();return true;}return false;}">  <input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">ȫѡ</td>
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
	NewsCodeType = "δ֪"
	Select Case Cstr(TypeVal)
	Case "1"
		NewsCodeType = "����"
	Case "2"
		NewsCodeType = "��Ϣ"
	Case "3"
		NewsCodeType = "���"
	Case "4"
		NewsCodeType = "��Ա"
	Case "5"
		NewsCodeType = "����"
	Case "6"
		NewsCodeType = "չ��"
	End Select
	NewsCodeType = NewsCodeType & "����"
End Function
%>
