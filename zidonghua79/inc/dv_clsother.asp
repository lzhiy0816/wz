<%
Rem ����ҳ��ͨ�ú���
'Dvbbs.Board_Setting(40)�Ƿ�̳��ϼ�������˳��ȡ���ϼ���̳������Ϣ
'���ֻȡ���ϵ�10��������Ϣ
'��������˵��ִ�
Sub CheckBoardInfo()
	Dim parentstr,parentboard,XpathSQL,i,Maxdepth,Node,NavStr
	parentstr = Application(Dvbbs.CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"&Dvbbs.BoardID&"']/@parentstr").text
	parentboard=Split(parentstr,",")
	If Dvbbs.UserID > 0 and Not Dvbbs.BoardMaster And CLng(Dvbbs.UserGroupID)=3 Then
		If Dvbbs.Board_Setting(40)="1" And Dvbbs.BoardParentID > 0 Then
			For i=0 to UBound(parentboard)
				If parentboard(i) <> "0" Then XpathSQL=XpathSQL &" or @boardid = " & parentboard(i)
			Next
		End If
		XpathSQL="boardmaster[@boardid = "& Dvbbs.Boardid & XpathSQL &"]/master[ . ='"& Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@username").text &"']"
		Dvbbs.BoardMaster=Not Application(Dvbbs.CacheName&"_boardmaster").documentElement.selectSingleNode(XpathSQL) Is Nothing
	End If
	If Dvbbs.BoardParentID > 0 Then
		Maxdepth=9
		If Ubound(parentboard) < Maxdepth+1 Then
			Maxdepth=Ubound(parentboard)
		End If
		For i=0 to Maxdepth
			If parentboard(i) <> "0"  Then
				Set Node=Application(Dvbbs.CacheName&"_boardlist").documentElement.selectSingleNode("board[@boardid='"& parentboard(i) &"']")
				If Node Is Nothing Then Exit For
				If i=0 Then
						NavStr=" <a href=""index.asp?boardid="& parentboard(i) &""" onmouseover=""showmenu(event,BoardJumpList("&Node.selectSingleNode("@boardid").text&"),'',0);"">"& Node.selectSingleNode("@boardtype").text &"</a> "
					Else
						NavStr=NavStr& "�� <a href=""index.asp?boardid="& parentboard(i) &""">"& Node.selectSingleNode("@boardtype").text &"</a> "
					End If
			End If
		Next
	End If
	Dvbbs.BoardInfoData=NavStr
	GetBoardPermission()
End Sub
Sub GetBoardPermission()
	Dim Rs,IsGroupSetting
	If Not Application(Dvbbs.CacheName &"_boarddata_" & Dvbbs.Boardid).documentElement.selectSingleNode("boarddata/@isgroupsetting") is nothing Then IsGroupSetting = Application(Dvbbs.CacheName &"_boarddata_" & Dvbbs.Boardid).documentElement.selectSingleNode("boarddata/@isgroupsetting").text
	If IsGroupSetting <> ""  Then
		IsGroupSetting = "," & IsGroupSetting & ","
		If InStr(IsGroupSetting,"," & Dvbbs.UserGroupID & ",")>0 Then
			Set Rs=Dvbbs.Execute("Select PSetting From Dv_BoardPermission Where Boardid="&Dvbbs.Boardid&" And GroupID="&Dvbbs.UserGroupID)
			If Not (Rs.Eof And Rs.Bof) Then
				Dvbbs.GroupSetting = Split(Rs(0),",")
			End If
			Set Rs=Nothing
		End If
		If Dvbbs.UserID>0 And InStr(IsGroupSetting,",0_"&Dvbbs.UserID&",")>0 Then
			Set Rs=Dvbbs.execute("Select Uc_Setting From Dv_UserAccess Where Uc_Boardid="&Dvbbs.BoardID&" And uc_UserID="&Dvbbs.Userid)
			If Not(Rs.Eof And Rs.Bof) Then
				Dvbbs.UserPermission=Split(Rs(0),",")
				Dvbbs.GroupSetting = Split(Rs(0),",")
				Dvbbs.FoundUserPer=True
			End If
			Set Rs=Nothing
		If Dvbbs.GroupSetting(70)="1"  Then
				Dvbbs.Master = True
			Else
				Dvbbs.Master = False
			End If
		End If
	End If
	If Not Dvbbs.BoardMaster Then Chkboardlogin()
End Sub
Rem �ܷ������̳���ж�
Public Sub Chkboardlogin()
	If Dvbbs.Board_Setting(1)="1" And Dvbbs.GroupSetting(37)="0" Then Dvbbs.AddErrCode(26)
	If Dvbbs.GroupSetting(0)="0"  Then Dvbbs.AddErrCode(27)
	'������̳����(�������¡����֡���Ǯ����������������������ɾ����ע��ʱ��)
	Dim BoardUserLimited
	BoardUserLimited = Split(Dvbbs.Board_Setting(54),"|")
	If Ubound(BoardUserLimited)=8 Then
		'����
		If Trim(BoardUserLimited(0))<>"0" And IsNumeric(BoardUserLimited(0)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(0)&"</B> ���ܽ���&action=OtherErr"
			If Clng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text)<Clng(BoardUserLimited(0)) Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(0)&"</B> ���ܽ���&action=OtherErr"
		End If
		'����
		If Trim(BoardUserLimited(1))<>"0" And IsNumeric(BoardUserLimited(1)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(1)&"</B> ���ܽ���&action=OtherErr"
			If Clng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text)<Clng(BoardUserLimited(1)) Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(1)&"</B> ���ܽ���&action=OtherErr"
		End If
		'��Ǯ
		If Trim(BoardUserLimited(2))<>"0" And IsNumeric(BoardUserLimited(2)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���Ǯ����Ϊ <B>"&BoardUserLimited(2)&"</B> ���ܽ���&action=OtherErr"
			If Clng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text)<Clng(BoardUserLimited(2)) Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���Ǯ����Ϊ <B>"&BoardUserLimited(2)&"</B> ���ܽ���&action=OtherErr"
		End If
		'����
		If Trim(BoardUserLimited(3))<>"0" And IsNumeric(BoardUserLimited(3)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(3)&"</B> ���ܽ���&action=OtherErr"
			If Clng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text)<Clng(BoardUserLimited(3)) Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(3)&"</B> ���ܽ���&action=OtherErr"
		End If
		'����
		If Trim(BoardUserLimited(4))<>"0" And IsNumeric(BoardUserLimited(4)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(4)&"</B> ���ܽ���&action=OtherErr"
			If Clng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpower").text)<Clng(BoardUserLimited(4)) Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(4)&"</B> ���ܽ���&action=OtherErr"
		End If
		'����
		If Trim(BoardUserLimited(5))<>"0" And IsNumeric(BoardUserLimited(5)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(5)&"</B> ���ܽ���&action=OtherErr"
			If Clng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userisbest").text)<Clng(BoardUserLimited(5)) Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(5)&"</B> ���ܽ���&action=OtherErr"
		End If
		'ɾ��
		If Trim(BoardUserLimited(6))<>"0" And IsNumeric(BoardUserLimited(6)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���ɾ������ <B>"&BoardUserLimited(6)&"</B> ���ܽ���&action=OtherErr"
			If Clng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userdel").text)>Clng(BoardUserLimited(6)) Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���ɾ������ <B>"&BoardUserLimited(6)&"</B> ���ܽ���&action=OtherErr"
		End If
		'ע��ʱ��
		If Trim(BoardUserLimited(7))<>"0" And IsNumeric(BoardUserLimited(7)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û�ע��ʱ����� <B>"&BoardUserLimited(7)&"</B> ���Ӳ��ܽ���&action=OtherErr"
			If DateDiff("s",Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@joindate").text,Now)<Clng(BoardUserLimited(7))*60 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û�ע��ʱ����� <B>"&BoardUserLimited(7)&"</B> ���Ӳ��ܽ���&action=OtherErr"
		End If
		
	End If
	'��֤����ж�Board_Setting(2)
	If Dvbbs.Board_Setting(2)="1" Then
		Dim Get_BoardUser_Money,Canlogin,Boarduser,i,BoardUser_Money
		Get_BoardUser_Money = False
		If Clng(Dvbbs.Board_Setting(62))>0 Or Clng(Dvbbs.Board_Setting(63))>0 Then Get_BoardUser_Money = True
			Canlogin = False
		If Dvbbs.UserID=0 Then
			Dvbbs.AddErrCode(24)
			Dvbbs.showerr()
		Else
			If Not Application(Dvbbs.CacheName &"_boarddata_" & Dvbbs.Boardid).documentElement.selectSingleNode("boarddata/@boarduser") Is Nothing Then boarduser=Application(Dvbbs.CacheName &"_boarddata_" & Dvbbs.Boardid).documentElement.selectSingleNode("boarddata/@boarduser").text
			If Boarduser=""  Then	
				Canlogin = False
			Else
				Boarduser=Split(Boarduser,",")
				For i = 0 To Ubound(Boarduser)
					If Get_BoardUser_Money Then
						BoardUser_Money = Split(Boarduser(i),"=")
						If Trim(Lcase(BoardUser_Money(0))) = Trim(Lcase(Dvbbs.MemberName)) Then
							'�޸��ж�֧����һ��ȯ����������Ч�� 2004-8-29 Dv.Yz
							If Not DateDiff("d",BoardUser_Money(1),Now()) > Cint(Dvbbs.Board_Setting(64))*30 Then
								Canlogin = True
								Exit For
							End If
						End If
					Else
						If Trim(Lcase(Boarduser(i))) = Trim(Lcase(Dvbbs.MemberName)) Then
							Canlogin = True
							Exit For
						End If
					End If
				Next
			End If
		End If
		If Get_BoardUser_Money And Instr(Lcase(Dvbbs.ScriptName),"pay_boardlimited")=0 And (Not Canlogin) Then Response.Redirect "pay_boardlimited.asp?boardid=" & Dvbbs.BoardID
		If Instr(Lcase(Dvbbs.ScriptName),"pay_boardlimited")=0 And (Not Canlogin) Then
			Dvbbs.AddErrCode(25)	
		End If
	End If
	Dvbbs.showerr()
End Sub
'�õ���̳���ֹ��λ��������,PageID=0Ϊ��ҳ,=1Ϊ�����б�ҳ��,=2Ϊ��������ҳ��
Sub GetForumTextAd(PageID)
	If Dvbbs.Forum_ads(12) = "1" Then
			 Select Case PageID
			 	Case 0
			 		
			 	Case 1
			 		If Not(Dvbbs.Forum_ads(15) = "0" Or Dvbbs.Forum_ads(15) = "2") Then Exit Sub
			 	Case 2
			 	 If Not((Dvbbs.Forum_ads(15) = "1" Or Dvbbs.Forum_ads(15) = "2")) Or (Dvbbs.Forum_ads(15) = "3") Then Exit Sub
			 	Case Else
			 		Exit Sub	
			 End Select
			 Dvbbs.Name = "Text_ad_"&Dvbbs.BoardID&"_"& Dvbbs.SkinID
			If Dvbbs.ObjIsEmpty() Then LoardTextAd()
			Response.Write Dvbbs.Value
	End If	
End Sub
Sub LoardTextAd()
	Dim XmlAds,Adstext,textvalue,XMLStyle,proc
		If IsObject(Application(Dvbbs.CacheName & "_TextAdservices")) Then
			Set XmlAds=Application(Dvbbs.CacheName & "_TextAdservices").cloneNode(True)
		Else
			Set XmlAds=Server.CreateObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
			XmlAds.appendChild(XmlAds.createElement("xml"))
		End If
	Adstext=split(Dvbbs.Forum_ads(16),Chr(10))
	If Dvbbs.Forum_ads(17)="" Or Not IsNumeric(Dvbbs.Forum_ads(17)) Then  Dvbbs.Forum_ads(17)=4
	Dvbbs.Forum_ads(17)=Abs(Dvbbs.Forum_ads(17))
	If Dvbbs.Forum_ads(17)=0 Then Dvbbs.Forum_ads(17)=4
	XmlAds.documentElement.attributes.setNamedItem(XmlAds.createNode(2,"tdcount","")).text=Dvbbs.Forum_ads(17)
	For Each textvalue In adstext
		XmlAds.documentElement.appendChild(XmlAds.createNode(1,"text","")).text=textvalue
	Next
	If not IsObject(Application(Dvbbs.CacheName & "_showtextads_"& Dvbbs.SkinID)) Then
		Set Application(Dvbbs.CacheName & "_showtextads_"& Dvbbs.SkinID)=Server.CreateObject("Msxml2.XSLTemplate" & MsxmlVersion)
		Set XMLStyle=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"  & MsxmlVersion)
		If UBound(Dvbbs.mainhtml) > 22 Then
			XMLStyle.loadxml Dvbbs.mainhtml(22)
		Else
			XMLStyle.load Server.MapPath(MyDbPath &"inc\Templates\Dv_textads.xslt")
		End If
		Application(Dvbbs.CacheName & "_showtextads_"& Dvbbs.SkinID).stylesheet=XMLStyle
	End If
	Set proc = Application(Dvbbs.CacheName & "_showtextads_"& Dvbbs.SkinID).createProcessor()
	proc.input = XmlAds
  proc.transform()
  Dvbbs.Name = "Text_ad_"&Dvbbs.BoardID&"_"& Dvbbs.SkinID
  Dvbbs.Value=proc.output
  Set XmlAds=Nothing 
  Set proc=Nothing
End Sub
Sub LoadBoardNews_Paper()
	Dim Rs,node
	Set Rs=Dvbbs.Execute("select S_ID,S_BoardID,S_UserName,S_Title From Dv_Smallpaper")
	If Not Rs.EOF Then
		Set Application(Dvbbs.CacheName & "_smallpaper")=Dvbbs.RecordsetToxml(rs,"smallpaper","xml")
		For Each Node in Application(Dvbbs.CacheName & "_smallpaper").documentElement.SelectNodes("smallpaper")
			Node.selectSingleNode("@s_username").text=Dvbbs.ChkBadWords(Node.selectSingleNode("@s_username").text)
			Node.selectSingleNode("@s_title").text=Dvbbs.ChkBadWords(Node.selectSingleNode("@s_title").text)
		Next
	Else
		Set Application(Dvbbs.CacheName & "_smallpaper")=Server.CreateObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
		Application(Dvbbs.CacheName & "_smallpaper").appendChild(Application(Dvbbs.CacheName & "_smallpaper").createElement("xml"))
	End If
End Sub

Sub Load_GroupName()
	Dim Rs,node
	Set Rs=Dvbbs.Execute("select * From Dv_GroupName")
	If Not Rs.EOF Then
		Set Application(Dvbbs.CacheName & "_GroupName")=Dvbbs.RecordsetToxml(rs,"groupname","group")
	Else
		Set Application(Dvbbs.CacheName & "_GroupName")=Server.CreateObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
		Application(Dvbbs.CacheName & "_GroupName").appendChild(Application(Dvbbs.CacheName & "_GroupName").createElement("groupname"))
	End If
End Sub
%>