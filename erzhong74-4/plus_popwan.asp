<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="Plus_popwan/cls_setup.asp"-->
<%
	Dim Action
	Dvbbs.LoadTemplates("")
	Dvbbs.Stats = "������Ϸ����"
	Dvbbs.Nav()
	Dvbbs.Head_var 0,0,Plus_Popwan.Program,"plus_popwan.asp"
	Dvbbs.ActiveOnline()
	action = Request("action")
	Page_main()

	If action<>"frameon" Then
		Dvbbs.Footer
	End If
	Dvbbs.PageEnd()

'ҳ���Ҳ����ݲ���
Sub Page_Center()
	Dim view
	view = Lcase(Request("view"))
	Select Case view
		Case "service"	: Service()
		Case "popmng"	:  PopManange()
		Case "regpopwan":  RegPopwan()
		Case "playgame"	:  PlayGame()
		Case "mygames"	:  MyGames()
		Case "nav"		:  PlusNav()
		Case "updategames"	: UpdateGames()
		Case "addplus"		:  addplus()
		Case "editpopwan"	:  Editpopwan()
		Case "loadgamelist"	:  LoadGameList()
		Case "defaultupdate"	:  DefaultUpdate()
		
		Case Else
			GameList()
	End Select

End Sub

Sub MyGames()
	'վ���ѿ�ͨ��Ϸ�б�
	'GameList()

		Dim XmlDom,Node,PUrl,DataToSend,Sign

		Sign = MD5(Dvbbs.membername&Plus_Popwan.ConfigNode.getAttribute("siteid")&Plus_Popwan.ConfigNode.getAttribute("key"),32)
		PUrl = "http://union.popwan.com/PlayGames/?"
		DataToSend = "siteid="&Plus_Popwan.ConfigNode.getAttribute("siteid")&"&username="&Server.urlencode(Dvbbs.membername)&"&sign="&Sign&"&encode=gb2312"
		PUrl = PUrl&DataToSend

%>
<div class="pw_tb0">
���ѽ������Ϸ�б�
	<div class="pw_par1" id="usergames">
	
<%
		'Response.Write PUrl
		Set XmlDom = Plus_Popwan.HttpPost(PUrl)
		If XmlDom Is Nothing Then
			Response.Write "��δ�н�����Ϸ����Ϣ���Ͽ�����������Ϊ���ṩ����Ϸ��Ȥ�ɣ�</div></div>"
			Exit Sub
		End If
		Set Node = XmlDom.selectSingleNode("masterselgames")
		If Node Is Nothing Then
			Response.Write "��δ�н�����Ϸ����Ϣ���Ͽ�����������Ϊ���ṩ����Ϸ��Ȥ�ɣ�</div></div>"
			Exit Sub
		End If
		'Plus_Popwan.Gamelist
		'XmlDom.documentElement.SelectNodes("masterselgames/game")
		Response.Write "<ul>"
		For Each Node In XmlDom.documentElement.SelectNodes("//masterselgames/game")
		%>
		<li>
		<a href="plus_popwan.asp?view=playgame&gid=<%=Server.URlencode(Node.selectSingleNode("entergameurl").text)%>" target="_blank"><%=Node.selectSingleNode("gamename").text%></a><br/>
		<span class="img"><a href="plus_popwan.asp?view=playgame&gid=<%=Server.URlencode(Node.selectSingleNode("entergameurl").text)%>" target="_blank"><img src="<%=Node.selectSingleNode("bigpic").text%>" alt="" border="0"/></a></span>
		</li>
		<%
		Next
		Response.Write "</ul>"
%><div style="clear:both;"></div>
	</div>
</div>
<%
Set XmlDom = Nothing
End Sub

Sub GameList()
	Dim Node,i
	i = 0
	If Not (Plus_Popwan.Gamelist Is Nothing) Then
		If Plus_Popwan.Gamelist.length < 1 Then
			Exit Sub
		End If
	Else
		Exit Sub
	End If
%>
<div class="pw_tb0">
��վ�ѿ�ͨ����Ϸ�б�
	<div class="pw_par1">
	<%
		For Each Node In Plus_Popwan.Gamelist
		If (Node.selectSingleNode("gamename") Is Nothing) Then Exit For
	%>
		<div class="gamebox1">
			<div><b class="font14"><%=Node.selectSingleNode("gamename").text%></b></div>
			<div>
				<div class="p1">
				<a href="<%=Plus_Popwan.ConfigNode.getAttribute("gamesite")&"/"&Node.selectSingleNode("shortname").text%>/" target="_blank"><img src="<%=Node.selectSingleNode("smallpic").text%>" alt="" border="0" /></a>
				</div>
				<div class="p2">
				<a href="plus_popwan.asp?view=playgame&gid=<%=Server.URlencode(Node.selectSingleNode("entergameurl").text)%>" target="_blank"><img src="<%=Plus_Popwan.Folder%>/images/gogame.gif" alt="" /></a> <a href="<%=Plus_Popwan.ConfigNode.getAttribute("gamesite")&"/"&Node.selectSingleNode("shortname").text%>/" target="_blank">�ٷ���վ</a>
				<%If Plus_Popwan.IsMaster Then%>
				<font class="green">������Ϸ�����ӣ�</font><input id="c_<%=i%>" name="gourl" value="<%=Plus_Popwan.ConfigNode.getAttribute("bbsurl")%>plus_popwan.asp?view=playgame&gid=<%=Node.selectSingleNode("entergameurl").text%>" size="40"/><button onclick="copyText(document.getElementById('c_<%=i%>'))">����</button>
				<%End If%>
				<br/>
				<%=Left(Node.selectSingleNode("gameintro").text,100)&"..."%>
				</div>
				<div style="clear:both"></div>
			</div>
		</div>

	<%	i =  i + 1
		Next
	%>
	</div>
</div>
<%
End Sub

'��ϸ��ϵ
Sub Service()
%>
<iframe id="pw_frame" src="<%=Plus_Popwan.ConfigNode.getAttribute("unionurl")%>/CustomerService/" scrolling="yes" frameborder="0" allowtransparency="true"></iframe>
<%
End Sub

'�޸İ�ע����Ϣ
Sub Editpopwan()
	If Not Plus_Popwan.IsMaster Then
		Response.Write "��û�иò���Ȩ�ޣ�"
		Exit Sub
	End If
	Dim BbsName,GameSite,BbsUrl,Userkey
	If LCase(Request("react"))="save" Then
		GameSite = Replace(LCase(Request.form("gamesite")),"'","")
		If Left(Request.form("gamesite"),7)<>"http://" Then
			GameSite = "http://"&GameSite
		End If
		Plus_Popwan.ConfigNode.setAttribute "bbsname",Dvbbs.Forum_info(0)
		Plus_Popwan.ConfigNode.setAttribute "bbsurl",Dvbbs.Forum_info(1)
		Plus_Popwan.ConfigNode.setAttribute "key",Request.form("userkey")
		Plus_Popwan.ConfigNode.setAttribute "siteid",Request.form("siteid")
		Plus_Popwan.ConfigNode.setAttribute "gamesite",GameSite
		Plus_Popwan.ConfigNode.setAttribute "joined","1"
		Plus_Popwan.Update_Config()
		Response.Write "�󶨲����ɹ��������µ����ȡ��Ϸ�б�����ظ��¡�"
		Exit Sub
	End If
%>
<div class="pw_join2">
	<table width="100%" border="0" cellpadding="4" cellspacing="2" class="pw_tb1">
	<form action="?view=editpopwan&react=save" method="post">
	<input type="hidden" size="40" name="bbsname" value="<%=Plus_Popwan.ConfigNode.getAttribute("bbsname")%>" />
	<input type="hidden" size="40" name="domain" value="<%=Plus_Popwan.ConfigNode.getAttribute("bbsurl")%>" />
	<tr>
		<th colspan="3"><font class="green">*�ѿ�ͨ</font>���޸İ���ע����������ϣ�*��ϸ�����������˹���ƽ̨���޸ġ���</th>
	</tr>
	<tr>
		<td align="right">����վ��������</td>
		<td><input type="text" size="40" name="gamesite" value="<%=Plus_Popwan.ConfigNode.getAttribute("gamesite")%>" /></td>
		<td rowspan="3"><input type="submit" name="submit" value="���°���������"/></td>
	</tr>
	<tr>
		<td align="right">����UserKEY</td>
		<td><input type="text" size="40" name="userkey" value="<%=Plus_Popwan.ConfigNode.getAttribute("key")%>" /></td>
		
	</tr>
	<tr>
		<td align="right">����SiteID</td>
		<td><input type="text" size="10" name="siteid" value="<%=Plus_Popwan.ConfigNode.getAttribute("siteid")%>" /></td>
	</tr>
	</form>
</table>
</div>

<%
End Sub

'ע�Ὺͨ��
Sub JoinForm()
	If Not Plus_Popwan.IsMaster Then
		Response.Write "��û�иò���Ȩ�ޣ�"
		Exit Sub
	End If
	Dim BbsName,GameSite,BbsUrl,Userkey
%>
<div class="pw_join2">
	<table width="100%" border="0" cellpadding="4" cellspacing="2" class="pw_tb1">
	<form action="?view=regpopwan" method="post">
	<input type="hidden" size="40" name="bbsname" value="<%=Plus_Popwan.ConfigNode.getAttribute("bbsname")%>" />
	<input type="hidden" size="40" name="domain" value="<%=Plus_Popwan.ConfigNode.getAttribute("bbsurl")%>" />
	<tr>
		<th><font class="yellow">*δ��ͨ</font>�����Ͽ�ͨ�ҵ���Ϸ����</th>
	</tr>
	<tr>
		<td>
		<input type="submit" name="submit" value="һ����ͨ�ҵ���Ϸ����"/>
		</td>
	</tr>
	</form>
</table>
</div>
<%
Call Editpopwan()
%>

<div class="pw_join1"></div>
<div class="pw_join0">
	<div class="pw_readme">
		<ul>
		<li>������ߵİ��������ģʽ������Ϊ���е�ѩ��ʽ�������鿴����</li>
		<li>վ�����е���ϷƵ��������ά������ӵ�����ݺͿ�����Ϸ��ƽ̨</li>
		<li>�ɰ���վ�����������ʣ��������û�ȫ������վ����</li>
		<li>�����Ϸͬʱ���û�ѡ���û�ճ�Ը��ߣ�����Ҳ����</li>
		<li>��������վ�������û����û�������ע�ἴ������Ϸ</li>
		<li>��ע�����ƣ�����վ�����ɼ��룬Сվ��һ����ֵ�û��ɱ������й������</li>
		<li>��ϵͳ�������˹���Ԥ�����վ��������٣��������漰����10������</li>
		</ul>
		
	</div>
</div>
<%
End Sub


Sub RegPopwan()
	Dim WebName,Domain,ValidCode,Sign,XmlDom,xmlhttp,DateTime,DataToSend,DataBack
	WebName=Server.Urlencode(Request.form ("bbsname"))
	Domain=Server.Urlencode(Replace(Request.form ("domain"),"'",""))
	ValidCode=Server.Urlencode(Request.form ("validcode"))
	DateTime=Server.Urlencode(CstrDateTime(Now()))
	DataToSend="domain=" & Domain & "&webname=" & WebName & "&time=" & DateTime & "&validcode=" & ValidCode &"&encode=gb2312"
	Sign=MD5(DataToSend,32)
	DataToSend=DataToSend & "&sign=" & Sign
	Set xmlhttp = Server.CreateObject("msxml2.ServerXMLHTTP")
	xmlhttp.setTimeouts 10000, 10000, 10000, 10000
'	xmlhttp.Open "POST","����ע��ҳ���ַ",False
	xmlhttp.Open "POST",Plus_Popwan.ConfigNode.getAttribute("unionurl") & "?" & DataToSend,False
	xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xmlhttp.send()
'Response.Write Plus_Popwan.ConfigNode.getAttribute("unionurl") & "?" & DataToSend
'	Response.Write Plus_Popwan.ConfigNode.getAttribute("unionurl") & "?" & DataToSend
		Set XmlDom = Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		If XmlDom.Load(xmlhttp.responseXML) Then
			If Not (XmlDom.documentElement.selectSingleNode("success") Is Nothing) Then
				If XmlDom.documentElement.selectSingleNode("success").text="1" Then
					Plus_Popwan.ConfigNode.setAttribute "key",XmlDom.documentElement.selectSingleNode("responseresult/userkey").text
					Plus_Popwan.ConfigNode.setAttribute "siteid",XmlDom.documentElement.selectSingleNode("responseresult/siteid").text
					Plus_Popwan.ConfigNode.setAttribute "gamesite","http://u.popwan.com/"&XmlDom.documentElement.selectSingleNode("responseresult/domain").text
					Plus_Popwan.ConfigNode.setAttribute "joined","1"
					Plus_Popwan.ConfigNode.setAttribute "addtime","1"					
					Plus_Popwan.Update_Config()
					Response.Write XmlDom.documentElement.selectSingleNode("message").text
				Else
					Response.Write XmlDom.documentElement.selectSingleNode("message").text
				End If
			Else				
				Response.Write "�����ύ��������!"
			End If
		Else
			'Response.Write PostUrl
			Response.Write "ͨ�ų�ʱ���������ύ��"
			'Response.Write xmlhttp.responsetext
		End If
		Set XmlHttp = Nothing
		Set XmlDom = Nothing
End Sub


Sub UpdateGames()
	If Not Plus_Popwan.IsMaster Then
		Response.Write "��û�иò���Ȩ�ޣ�"
		Exit Sub
	End If

	If Request("react") = "updategames" Then
		'������Ϸ�б�
		Plus_Popwan.UpdateGamesInfo()
	Else
		%>
		<div class="pw_tb0">
		��������Ϸƽ̨��������Ϸ�б�����ִ�����¸��£�
		<div class="pw_par1">
		<input type="submit" name="submit" value="��ȡ���¿�ͨ��Ϸ�б�" onclick="window.location='plus_popwan.asp?view=updategames&react=updategames'"/>
		˵��������ȡ�ڹ���ƽ̨�ѹ�ѡ����Ϸ�б�
		</div>
		</div>
		<%
	End If
End Sub

'���뿪ͨPOPWAN/����
Sub PopManange()
	If Not Plus_Popwan.IsMaster Then
		Response.Write "��û�иò���Ȩ�ޣ�"
		Exit Sub
	End If
	Dim IsJoin
	IsJoin = Plus_Popwan.ConfigNode.getAttribute("joined")
	If IsJoin="0" Then
		JoinForm()
		Exit Sub
	End If

	PopMainTop()
	UpdateGames()
	PlusNav()
	DefaultUpdate()
	GameList()

End Sub

Sub PlusNav()
%>
<div class="pw_tb0">
�����ֱ�ӵ������ͨ��Ϸ�������ӣ�
<div class="pw_par1">
<input type="submit" name="submit" value="�����Ϸ��������" onclick="window.location='plus_popwan.asp?view=addplus&plus_id=POPWANID'"/>
</div>
</div>
<%
End Sub

Sub Addplus_main(iType,iName,iUrl,iPlus_id,iCopyright,isblank)
	Dim plus_setting
	plus_setting = isblank&"|0|0|||0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0|||�����ֶ�1,�����ֶ�2,�����ֶ�3,�����ֶ�4,�����ֶ�5,�����ֶ�6,�����ֶ�7,�����ֶ�8,�����ֶ�9,�����ֶ�10,�����ֶ�11,�����ֶ�12,�����ֶ�13,�����ֶ�14,�����ֶ�15,�����ֶ�16,�����ֶ�17,�����ֶ�18,�����ֶ�19|||0,0|24,,,0,0,0,0,0,0,0,0,0,"

	Dvbbs.Execute("Insert into dv_plus (Plus_Type,Plus_Name,Isuse,Plus_Setting,Mainpage,IsShowMenu,plus_id,plus_Copyright) values ("&Dvbbs.Checkstr(iType)&",'"&Dvbbs.Checkstr(iName)&"',1,'"&Dvbbs.Checkstr(plus_setting)&"','"&Dvbbs.Checkstr(iUrl)&"',1,'"&Dvbbs.Checkstr(iPlus_id)&"','"&Dvbbs.Checkstr(iCopyright)&"')")
End Sub

Sub addplus()
	If Not Plus_Popwan.IsMaster Then
		Response.Write "��û�иò���Ȩ�ޣ�"
		Exit Sub
	End If
	If Plus_Popwan.ConfigNode.getAttribute("joined")="0" Then
		Response.Write "<font class=""yellow"">*��ͨ���������վ��ſ��Խ��иò�����</font>"
		JoinForm()
		Exit Sub
	End If
	Dim SiteID,Plus_Name,Mainpage,i,iid
	Dim Node
		SiteID = "PopWanID"
		Plus_Name = "������Ϸ����"

	Dvbbs.Execute("delete from Dv_Plus where Plus_id like '"& SiteID &"%'")
	Addplus_main 0,Plus_Name,"plus_popwan.asp",SiteID,Plus_Name,0
	iid = Getid(SiteID)
	If Plus_Popwan.Gamelist.length>0 Then
		For Each Node In Plus_Popwan.Gamelist
			Addplus_main iid,Node.selectSingleNode("gamename").text,"plus_popwan.asp?view=playgame&gid="&Server.URlencode(Node.selectSingleNode("entergameurl").text),SiteID&"_"&Node.selectSingleNode("shortname").text,Node.selectSingleNode("gamename").text,1
		Next
	End If
		
		LoadForumPlusMenuCache
		%>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
			alert("��ӳɹ���");
			location.href='plus_popwan.asp?view=nav';
		//-->
		</SCRIPT>
		<%
End Sub

Function Getid(id)
	Dim iRs
	Set iRs = Dvbbs.Execute("Select Top 1 [id] From Dv_Plus Where Plus_Type='0' and Plus_id='"&Dvbbs.Checkstr(id)&"'")
	If Not iRs.Eof Then
		Getid = iRs(0)
	Else
		Getid = 0
	End If
	iRs.Close:Set iRs = Nothing
End Function

Sub LoadForumPlusMenuCache()
	Dvbbs.Name="Plus_Settingts"
	Dim Rs,SQL
	SQL = "select plus_ID,Plus_Setting,Plus_Name,plus_Copyright from [Dv_plus] Order By ID"
	Set Rs = Dvbbs.Execute(SQL)	
	If Not Rs.Eof Then
		Dvbbs.Name="Plus_Settingts"
		Dvbbs.value = Rs.GetRows(-1)
	End If
	Set Rs = Nothing
	Dvbbs.LoadPlusMenu()
End Sub


Sub PopMainTop()
%>
<div class="pw_tb0">
����ע�����ˣ����ע����Ϣ�ǣ�
<div class="pw_par1">
<span style="float:right;">[<a href="plus_popwan.asp?view=editpopwan">�޸İ���Ϣ</a>] | [<a href="<%=Plus_Popwan.p_configdb%>" target="_blank">�������ƽ̨</a>]</span>
��̳���ƣ�<span class="bluefont"><%=Plus_Popwan.ConfigNode.getAttribute("bbsname")%></span><br/>
����������<a href="" target="_blank"><%=Plus_Popwan.ConfigNode.getAttribute("gamesite")%></a><br/>
����SiteID��<span class="green"><%=Plus_Popwan.ConfigNode.getAttribute("siteid")%></span><br/>
����UserKey��<span class="yellow"><%=Plus_Popwan.ConfigNode.getAttribute("key")%></span><br/>
	<%
		Dim XmlDom,Node
		Set XmlDom = Plus_Popwan.HttpPost(Plus_Popwan.ConfigNode.getAttribute("apiurl")&"/webinfo/?siteid="&Plus_Popwan.ConfigNode.getAttribute("siteid"))
		If XmlDom Is Nothing Then
			Exit Sub
		End If
		Set Node = XmlDom.selectSingleNode("webinfo")
		If Not (Node Is Nothing) Then
			Response.Write "����ע��<span  class=""bluefont"">"&Node.selectSingleNode("yestoday").text&"</span>"
			Response.Write "������ע��<span  class=""yellow"">"&Node.selectSingleNode("today").text&"</span>"
		End If
		Set XmlDom = Nothing
	%>
</div>
</div>
<%
End Sub

'�û�������Ϸ��ת
Sub PlayGame()
	Dim GoGame,UniUrl,SiteID,UserKey,Sign,DateTime,UserEmail,Userid
	SiteID = Plus_Popwan.ConfigNode.getAttribute("siteid")
	UserKey = Plus_Popwan.ConfigNode.getAttribute("key")
	UniUrl = Plus_Popwan.ConfigNode.getAttribute("gamesite")
	GoGame = Request.QueryString ("gid")
	DateTime = CstrDateTime(Now()) 'FormatDateTime(Now(),0)
	
	UserEmail = ""
	If Dvbbs.Userid>0 Then
		Sign = MD5("userid="&Dvbbs.Userid&"&username=" & Dvbbs.MemberName & "&siteid=" & siteid & "&time=" & DateTime & "&userkey=" & UserKey,32)
		UserEmail = Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@useremail").text
		Response.Redirect UniUrl & "/webuserreg/?userid="&Dvbbs.Userid&"&username=" & Server.Urlencode(Dvbbs.MemberName) & "&email="&Server.Urlencode(UserEmail)&"&siteid=" & SiteID & "&time=" & Server.Urlencode(DateTime) & "&sign=" & Sign & "&encode=gb2312&go=" & Server.Urlencode(Plus_Popwan.ConfigNode.getAttribute("gamesite")&"/"&GoGame)
	Else
		Response.Redirect "login.asp?f="&Server.UrlEncode("plus_popwan.asp?gid="&GoGame)
	End If
End Sub

Function CstrDateTime(t)
	Dim y,m,d,h,mi,s
	y=CStr(Year(t))
	m=CStr(Month(t))
	d=CStr(Day(t))
	h=CStr(Hour(t))
	mi=CStr(Minute(t))
	s=CStr(Second(t))
	CstrDateTime =  y &"-" & m &"-" &d & " "& h &":"& mi &":" &s
End Function

Sub LoadGameList()
	If Not Plus_Popwan.IsMaster Then
		Response.Write "��û�иò���Ȩ�ޣ�"
		Exit Sub
	End If
	If Request("react") = "updategames" Then
		'����Ĭ����Ϸ�б�
		Plus_Popwan.UpdateDefaultGamesInfo(1)
	Else
		Call UpdateGames()
		%>
		<div class="pw_tb0">
		��ȡĬ����Ϸ�б�ִ�����¸��£�
			<div class="pw_par1">
				<input type="submit" name="submit" value="��ȡĬ����Ϸ�б�" onclick="window.location='plus_popwan.asp?view=loadgamelist&react=updategames'"/>
				˵��������ȡ��Ϸƽ̨Ĭ����Ϸ�б�
		</div>
		</div>
		<%
	End If
End Sub

Sub DefaultUpdate()
	If Not Plus_Popwan.IsMaster Then
		Response.Write "��û�иò���Ȩ�ޣ�"
		Exit Sub
	End If
	If Request("react") = "updat" Then
		If Plus_Popwan.ConfigNode.getAttribute("isupdate")="0" Then
			Plus_Popwan.ConfigNode.setAttribute "isupdate","1"
		Else
			Plus_Popwan.ConfigNode.setAttribute "isupdate","0"
		End If
		Plus_Popwan.Update_Config()
		Response.Write "�޸ĳɹ���"
		Exit Sub
	Else
		%>
		<div class="pw_tb0">
		�Ƿ��Զ�����Ĭ����Ϸ�б�ִ�����¸��£�
			<div class="pw_par1">
			��ǰ״̬Ϊ��
			<%If Plus_Popwan.ConfigNode.getAttribute("isupdate")=0 Then%>
				<span class="bluefont">�ֶ�����</span>
				<input type="submit" name="submit" value="�Զ�����" onclick="window.location='plus_popwan.asp?view=defaultupdate&react=updat'"/>
			<%Else%>
				<span class="green">�Զ�����</span>
				<input type="submit" name="submit" value="�ֶ�����" onclick="window.location='plus_popwan.asp?view=defaultupdate&react=updat'"/>
			<%End If%>
			</div>
		</div>
		<%
	End If
End Sub
%>