<%
'=========================================================
' Copyright (C) 2003,2004 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
' File: Class_Mobile.asp
' Date: 2004-8-3
' Author: Dv Dever ,www.aspsky.net
' �ļ���;���ֻ��ƶ���̳����
'=========================================================
Class Mobile_Forum
	Rem ====================�������ֿ�ʼ==============
	Public Path,StartID,Number,Mobile,Stype,OP,Child,Self
	Public PathCount
	Public OtherContent
	Private ViewIpLimited
	Dim Re


	Rem ====================�������ֽ���==============
	
	Rem ======================���̲���================
	'Class����ʱ�Զ�ִ�еĴ���
	Private Sub Class_Initialize()
		Dvbbs.UserID = 0
		OtherContent = ""
		Path = Trim(Checkstr(Request("Path")))
		StartID = ChkNumeric(Trim(Request("StartID")))
		Number = ChkNumeric(Trim(Request("Number")))
		Mobile = ChkNumeric(Trim(Request("Mobile")))
		Stype = ChkNumeric(Trim(Request("Stype")))
		OP = ChkNumeric(Trim(Request("OP")))
		Child = ChkNumeric(Trim(Request("Child")))
		Self = Child
		Path = Split(Path,"/")
		PathCount = Ubound(Path)
		ViewIpLimited = ",219.238.232.59,219.153.18.230,219.153.18.162,,"
		'ViewIpLimited = ",61.132.138.120,"
		'ChkIpLimited
		ChkWapUser
	End Sub

	Private Sub ChkIpLimited()
		Dim ReServerIp
		ReServerIp = Trim(Request.ServerVariables("REMOTE_ADDR"))
		If ReServerIp = "" Or Instr(ViewIpLimited,ReServerIp) = 0 Then
			ShowMobileErr("����IP��"& ReServerIp &" �����������Ƶĵ�ַ��")
		End If
	End Sub

	'��֤�û�
	Private Sub ChkWapUser
		Mobile = Ccur(Mobile)
		If Mobile=0 or Len(Mobile)<11 Then
			If InStr(Dvbbs.ScriptName,"wap_userlogin.asp")=0 Then
				ShowMobileErr("�����ֻ����룺"& Mobile &" ���ܷ��ʱ���̳��")
			End If
		End If
		If Mobile=0 Then
			Dvbbs.UserID = 0
			Dvbbs.UserGroupID = 7
			Dvbbs.LetGuestSession()
		Else
			'Session(CacheName & "UserID")�û�����=0dvbbs+1ˢ��ʱ��+2����ʱ��+3���ڰ���ID+4�û�ID+5�û���+6�û�����+7�û�����+8�û�������+9�û�������+10�û��Ա�+11�û�ͷ��+12�û�ͷ���+13�û�ͷ���+14�û�ע��ʱ��+15�û�����½ʱ��+16�û���½����+17�û�״̬+18�û��ȼ�+19�û���ID+20�û�����+21�û���Ǯ+22�û�����UserEp+23�û�����UserCp+24�û�����+25�û�����+26����½IP+27�û���ɾ����+28�û�������+29�û�����״̬+30�û��������+31�û������Ա+32�û��ֻ�+33�û���ͼ��+34�û�ͷ��+35��֤����+36�û�������Ϣ+37�û����+38�û���ȯ+	39�����û�ID+40VIP�Ǽ�ʱ��+41VIP��ֹʱ��+42�Ƿ����ȫ���Զ���Ȩ��IsUserPermissionAll+43�Ƿ��ж������û���IsUserPermissionOnly+44��ʱ����+45Dvbbs
			Dim Rs,SQL
			Sql="Select UserID,UserName,UserPassword,UserEmail,UserPost,UserTopic,UserSex,UserFace,UserWidth,UserHeight,JoinDate,LastLogin as cometime ,LastLogin,LastLogin as activetime,UserLogins,Lockuser,Userclass,UserGroupID,UserGroup,userWealth,userEP,userCP,UserPower,UserBirthday,UserLastIP,UserDel,UserIsBest,UserHidden,UserMsg,IsChallenge,UserMobile,TitlePic,UserTitle,TruePassWord,UserToday,UserMoney,UserTicket,FollowMsgID,Vip_StarTime,Vip_EndTime,userid as boardid"
			Sql=Sql & " From [Dv_User] Where UserMobile = '" & Mobile &"'"
			Set Rs = Dvbbs.Execute(Sql)
			If Rs.Eof Then
				Rs.Close:Set Rs = Nothing
				Dvbbs.UserID = 0
				Dvbbs.UserGroupID = 7
				Dvbbs.LetGuestSession()
			Else
				Set Dvbbs.UserSession=RecordsetToxml(rs,"userinfo","xml")
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@cometime").text=Now()
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@activetime").text=DateAdd("s",-3600,Now())
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@boardid").text=boardid
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo").attributes.setNamedItem(Dvbbs.UserSession.createNode(2,"isuserpermissionall","")).text=Dvbbs.FoundUserPermission_All()
				If Not (Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usergroupid2") is Nothing )  Then
						Dvbbs.FoundMyGroupID =  CLng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usergroupid2").text)
				End If	
				If Dvbbs.FoundMyGroupID > 0 Then
					Dvbbs.UserSession.documentElement.selectSingleNode("userinfo").attributes.setNamedItem(Dvbbs.UserSession.createNode(2,"usergroupid2","")).text=Dvbbs.FoundMyGroupID
				End If
				Dim BS
				Set Bs=Dvbbs.GetBrowser()
				Dvbbs.UserSession.documentElement.appendChild(Bs.documentElement)
				If EnabledSession Then
					Session(Dvbbs.CacheName & "UserID")= Dvbbs.UserSession.xml
				End If
				Dvbbs.GetCacheUserInfo()
			End If
		End If
		'IP����
		If Dvbbs.UserSession.documentElement.selectSingleNode("agent/@lockip").text="1"  Then
			If Not Dvbbs.Page_Admin Then Response.Redirect "showerr.asp?action=iplock"
		End If
		Dvbbs.GetGroupSetting()
			If CInt(Dvbbs.GroupSetting(0))=0 And Not Dvbbs.Page_Admin Then
			ShowXMLStar
			AddErrCode(8)
			ShowXMLEnd
			Response.End
		End If
	End Sub
	'���XML��ʼ�ı��
	Public Sub ShowXMLStar()
		Response.Clear
		Response.CharSet="gb2312"
		Response.ContentType="text/xml" 
		Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"
		Response.Write vbNewLine
		Response.Write "<bbs_query>"
		Response.Write vbNewLine
		Response.Write "<forumname>"&ForMatHtmlTitle(Dvbbs.Forum_Info(0))&"</forumname>"
	End Sub 
	Public Sub ShowErr(ErrCode,ErrMsg)
		'Call ShowXMLStar
		'If Dvbbs.ScriptName="wap_board.asp" Then
			'ShowCodes "" ,4 ,0 ,4 ,ErrMsg ,"" ,Dvbbs.Forum_Info(0) ,Now ,Now
		'Else
			Response.Write "<errcode>"
			Response.Write ErrCode
			Response.Write "</errcode>"
			Response.Write vbNewLine
			Response.Write "<errmsg>"
			Response.Write ForMatHtml(ErrMsg)
			Response.Write "</errmsg>"
		'End If
	End Sub

	'���ģ��
	Public Sub ShowCodes(S_Self ,S_Child ,S_Sid ,S_Stype ,S_Name ,S_Content ,S_OtherContent,S_Author ,S_Createtime ,S_Modifytime)
		Dim CodesString
		CodesString = "<query_result>" & vbNewLine
		CodesString = CodesString & "<self>" & S_Self & "</self>" & vbNewLine
		CodesString = CodesString & "<child>" & S_Child & "</child>" & vbNewLine
		CodesString = CodesString & "<sid>" & S_Sid & "</sid>" & vbNewLine
		CodesString = CodesString & "<stype>" & S_Stype & "</stype>" & vbNewLine
		CodesString = CodesString & "<name><![CDATA[" & ForMatHtmlTitle(S_Name) & "]]></name>" & vbNewLine
		CodesString = CodesString & S_Content
		CodesString = CodesString & S_OtherContent & vbNewLine
		CodesString = CodesString & "<author>" & S_Author & "</author>" & vbNewLine
		CodesString = CodesString & "<createtime>" & S_Createtime & "</createtime>" & vbNewLine
		CodesString = CodesString & "<modifytime>" & S_Modifytime & "</modifytime>" & vbNewLine
		CodesString = CodesString & "</query_result>" & vbNewLine
		Response.Write CodesString
	End Sub

	Public Function Format_Content(sType,sBody)
		Dim CodesString
		If sType = 1 Then
			CodesString = "<content type=""other"" src="""&sBody&"""></content>" & vbNewLine
		Else
			CodesString = "<content type=""text""><![CDATA[" & sBody &"]]></content>" & vbNewLine
		End if
		Format_Content = CodesString
	End Function

	'���XML�����ı��
	Public Sub ShowXMLEnd()
		Response.Write vbNewLine
		Response.Write "</bbs_query>"
	End Sub
	Rem ====================���̲��ֽ���==============
	
	Rem ======================��������================
	'ͨ�ò������˺���
	Public Function Checkstr(Str)
		If Isnull(Str) Then
			CheckStr = ""
			Exit Function 
		End If
		Str = Replace(Str,Chr(0),"")
		CheckStr = Replace(Str,"'","''")
	End Function

	'�жϲ����Ƿ������Ͳ��Ҳ�������
	Public Function IsTrueNumeric(Str)
		Dim Numeric
		Numeric=Str & ""
		If IsNumeric(Numeric) And InStr(Numeric,",")=0 Then 
			IsTrueNumeric=True
		Else
			IsTrueNumeric=False
		End If
	End Function

	'�жϲ����Ƿ������Ͳ��Ҳ�������
	Public Function ChkNumeric(Str)
		ChkNumeric = 0
		If Str = Null Then Exit Function
		If IsNumeric(Str) And InStr(Str,",")=0 Then 
			ChkNumeric = cCur(Str)
		End If
	End Function

	'����HTML��ǣ��������з���	����
	Function ForMatHtml(str)
		OtherContent = ""
		Dim Tempstr,RegFound
		RegFound = False
		If Str<>"" And Not IsNull(Str) Then
			Set Re = new RegExp
			re.IgnoreCase =True
			re.Global=True
			re.Pattern = "(<br>)"
			Str = re.Replace(Str , CHR(13) & CHR(10))
			re.Pattern="(</p><p>)"
			Str=re.Replace(Str, CHR(13) & CHR(10))
			re.Pattern="(<[|\/]p>)"
			Str=re.Replace(Str, CHR(13) & CHR(10))
			re.Pattern="<(.[^>]*)>"
			Str=re.Replace(Str,"")
			re.Pattern = "(&nbsp;)"
			Str = re.Replace(Str,Chr(9))
			're.Pattern = "\[(i|b|u|center)\]((.|\n)*)\[\/(\1)\]"
			'Str = re.Replace(Str,"<$1>$2</$4>")
			re.Pattern = "\[(fly|move)\]((.|\n)*)\[\/(\1)\]"
			Str = re.Replace(Str,CHR(10)&"$2"&CHR(10))
			re.Pattern = "\[(size|color|face|glow|shadow)=(.[^\[]*)\]((.|\n)*)\[\/(\1)\]"
			Str = re.Replace(Str,"$3")
			re.Pattern = "\[align=(center|left|right)\]((.|\n)*)\[\/align\]"
			Str = re.Replace(Str,"[$1]$2[/$1]")
			re.Pattern = "\[(point|post|power|usercp|money|usemoney|username)=(.[^\[]*)\]((.|\n)*)\[\/(\1)\]"
			Str = re.Replace(Str,CHR(10)&"�������������ݲ����������"&CHR(10))
			re.Pattern="(\[replyview\])(.|\n)*(\[\/replyview\])"
			Str = re.Replace(Str,CHR(10)&"�������������ݲ����������"&CHR(10))
			re.Pattern = "\[(html|code)\]((.|\n)*)\[\/(\1)\]"
			Str = re.Replace(Str, CHR(10) & "����Ϊ����Դ���룺"& CHR(10)&CHR(10) & "$2" & CHR(10))
			re.Pattern = "\[(quote)\]((.|\n)*)\[\/(\1)\]"
			Str = re.Replace(Str, CHR(10) & "���ģ�"& CHR(10)&CHR(10) & "$2" & CHR(10))
			re.Pattern = "\[(url)=(.[^\[]*)\]((.|\n)*)\[\/(\1)\]"
			Str = re.Replace(Str,"($3 :$2)")
			If InStr(Lcase(Str),"[em")>0 Then
				re.Pattern="\[em(.[^\[]*)\]"
				Str=re.Replace(Str,"[img src="""&EmotPath&"em$1.gif""]")
			End If
			
			Dim Matches,Match
			're.Pattern = "\[(mp|rm|qt|flash)(.[^\[]*)\]((.|\n)([^\[\]]+)*)\[\/\1\]"
			re.Pattern = "\[(mp|rm|qt|flash|img)(.[^\[]*)\]([^\[\]]+)\[\/\1\]"
			If re.Test(Str) Then
				Set Matches = re.Execute(Str)
				For Each Match in Matches
					OtherContent = OtherContent & Match.Value
				Next
				OtherContent = re.Replace(OtherContent,"<content type=""$1"" src=""$3""></content>")
				Str = re.Replace(Str,"")
			End If
			If InStr(Lcase(Str),"[img]")>0 Then
				re.Pattern = "\[(img)\]([^\[\]]+)\[\/\1\]"
				If re.Test(Str) Then
					Set Matches = re.Execute(Str)
					For Each Match in Matches
						OtherContent = OtherContent & Match.Value
					Next
					OtherContent = re.Replace(OtherContent,"<content type=""image"" src=""$2""></content>")
					Str = re.Replace(Str,"")
				End If
			End If
			If InStr(Lcase(Str),"[upload=")>0 Then
				re.Pattern = "\[(upload)=(gif|jpg|jpeg|bmp|png)\]([^\[\]]+)\[\/\1\]"
				If re.Test(Str) Then
					Set Matches = re.Execute(Str)
					For Each Match in Matches
						OtherContent = OtherContent & Match.Value
					Next
					OtherContent = re.Replace(OtherContent,"<content type=""image"" src="""&Forum_Url&"$3""></content>")
					Str = re.Replace(Str,"")
				End If
			End If
			If InStr(Lcase(Str),"[upload=")>0 Then
				re.Pattern = "\[(upload)=(swf|swi)\]([^\[\]]+)\[\/\1\]"
				If re.Test(Str) Then
					Set Matches = re.Execute(Str)
					For Each Match in Matches
						OtherContent = OtherContent & Match.Value
					Next
					OtherContent = re.Replace(OtherContent,"<content type=""flash"" src="""&Forum_Url&"$3""></content>")
					Str = re.Replace(Str,"")
				End If
			End If
			If InStr(Lcase(Str),"[upload=")>0 Then
				re.Pattern = "\[(upload)=(.[^\[]*)\]([^\[\]]+)\[\/\1\]"
				If re.Test(Str) Then
					Set Matches = re.Execute(Str)
					For Each Match in Matches
						OtherContent = OtherContent & Match.Value
					Next
					OtherContent = re.Replace(OtherContent,"<content type=""other"" src="""&Forum_Url&"$3""></content>")
					Str = re.Replace(Str,"")
				End If
			End If
			
			re.Pattern="(\[(|\/)B\])"
			Str=re.Replace(Str, "[$2b]")
			re.Pattern="(\[(|\/)i\])"
			Str=re.Replace(Str, "[$2i]")
			re.Pattern="(\[(|\/)u\])"
			Str=re.Replace(Str, "[$2u]")
			Set Re=Nothing
			Str = Replace(Str,"]]>","]]&gt;")
			ForMatHtml = Str
		End If
	End Function

	'����
	Function ForMatHtmlTitle(str)
		If Str<>"" And Not IsNull(Str) Then
			Str = Replace(Str, CHR(10), "")
			Str = Replace(Str, CHR(13), "")
			Dim Re
			Set Re = new RegExp
			re.IgnoreCase =True
			re.Global=True
			re.Pattern="<(.[^>]*)>"
			Str=re.Replace(Str,"")
			re.Pattern = "(&nbsp;)"
			Str = re.Replace(Str,Chr(9))
			Set Re=Nothing
			Str = Replace(Str,"]]>","]]&gt;")
			ForMatHtmlTitle = Str
		End If
	End Function

	Rem ====================�������ֽ���==============
	
	Rem ======================���Բ���================
	
	Rem ====================���Բ��ֽ���==============

	'������Ϣ��ErrCodeΪ������Ϣ����
	Public Sub AddErrCode(ByVal ErrCode)
		If ErrCode<>"" and IsNumeric(ErrCode) Then
			Dvbbs.LoadTemplates("showerr")
			ShowErr 0,Template.Strings(ErrCode)
		End If
	End Sub

	Public Sub LoadBoardPass()
		If Dvbbs.BoardID < 1 Then Exit Sub
		'Dvbbs.Name="BoardInfo_" & Dvbbs.BoardID
		'If Dvbbs.ObjIsEmpty() Then Dvbbs.ReloadBoardInfo(Dvbbs.BoardID)
		'Dvbbs.Board_Data = Dvbbs.Value
		'Dvbbs.Board_Setting = Split(Dvbbs.Board_Data(16,0),",")
		'Dvbbs.sid = Dvbbs.Board_Data(15,0)
		If Dvbbs.UserID > 0 Then
			If Dvbbs.Master Then
				Dvbbs.Boardmaster=True
			ElseIf Dvbbs.Superboardmaster Then
				Dvbbs.Boardmaster=True
			ElseIf Dvbbs.UserGroupID =3 And Not Trim(Dvbbs.BoardMasterList) = "" Then
				If Instr("|"&lcase(Dvbbs.BoardMasterList)&"|","|"&lcase(Dvbbs.Membername)&"|")>0 Then
					Dvbbs.Boardmaster=True
				End If
			End If
		End If
		GetBoardPermission()
	End Sub

	Rem ��ð����û���Ȩ������
	Public Sub GetBoardPermission()
		Dim Rs,IsGroupSetting
		IsGroupSetting = Dvbbs.IsGroupSetting
		If IsGroupSetting<>"" And Not IsNull(IsGroupSetting) Then
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
			End If
		End If
		Call Chkboardlogin()
	End Sub
	Rem �ܷ������̳���ж�
	Public Sub Chkboardlogin()
		Dvbbs.GetForum_Setting
		If Dvbbs.Board_Setting(1)="1" And Dvbbs.GroupSetting(37)="0" Then ShowMobileErr("��û��Ȩ�޽���������̳��")
		If Dvbbs.GroupSetting(0)="0"  Then ShowMobileErr("��û��Ȩ�޽��뱾��̳��")
		If Dvbbs.Boardmaster Then Exit Sub
		'������̳����(�������¡����֡���Ǯ����������������������ɾ����ע��ʱ��)
		Dim BoardUserLimited
		BoardUserLimited = Split(Dvbbs.Board_Setting(54),"|")
		If Ubound(BoardUserLimited)=8 Then
			'����
			If Trim(BoardUserLimited(0))<>"0" And IsNumeric(BoardUserLimited(0)) Then
				If Dvbbs.UserID = 0 Then ShowMobileErr("�������������û���������Ϊ "&BoardUserLimited(0)&" ���ܽ���")
				If Clng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text)<Clng(BoardUserLimited(0)) Then ShowMobileErr("�������������û���������Ϊ "&BoardUserLimited(0)&" ���ܽ���")
			End If
			'����
			If Trim(BoardUserLimited(1))<>"0" And IsNumeric(BoardUserLimited(1)) Then
				If Dvbbs.UserID = 0 Then ShowMobileErr("�������������û���������Ϊ "&BoardUserLimited(1)&" ���ܽ���")
				If Clng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text)<Clng(BoardUserLimited(1)) Then ShowMobileErr("�������������û���������Ϊ "&BoardUserLimited(1)&" ���ܽ���")
			End If
			'��Ǯ
			If Trim(BoardUserLimited(2))<>"0" And IsNumeric(BoardUserLimited(2)) Then
				If Dvbbs.UserID = 0 Then ShowMobileErr("�������������û���Ǯ����Ϊ "&BoardUserLimited(2)&" ���ܽ���")
				If Clng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text)<Clng(BoardUserLimited(2)) Then ShowMobileErr("�������������û���Ǯ����Ϊ "&BoardUserLimited(2)&" ���ܽ���")
			End If
			'����
			If Trim(BoardUserLimited(3))<>"0" And IsNumeric(BoardUserLimited(3)) Then
				If Dvbbs.UserID = 0 Then ShowMobileErr("�������������û���������Ϊ "&BoardUserLimited(3)&" ���ܽ���")
				If Clng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text)<Clng(BoardUserLimited(3)) Then ShowMobileErr("�������������û���������Ϊ "&BoardUserLimited(3)&" ���ܽ���")
			End If
			'����
			If Trim(BoardUserLimited(4))<>"0" And IsNumeric(BoardUserLimited(4)) Then
				If Dvbbs.UserID = 0 Then ShowMobileErr("�������������û���������Ϊ "&BoardUserLimited(4)&" ���ܽ���")
				If Clng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpower").text)<Clng(BoardUserLimited(4)) Then ShowMobileErr("�������������û���������Ϊ "&BoardUserLimited(4)&" ���ܽ���")
			End If
			'����
			If Trim(BoardUserLimited(5))<>"0" And IsNumeric(BoardUserLimited(5)) Then
				If Dvbbs.UserID = 0 Then ShowMobileErr("�������������û���������Ϊ "&BoardUserLimited(5)&" ���ܽ���")
				If Clng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userisbest").text)<Clng(BoardUserLimited(5)) Then ShowMobileErr("�������������û���������Ϊ "&BoardUserLimited(5)&" ���ܽ���")
			End If
			'ɾ��
			If Trim(BoardUserLimited(6))<>"0" And IsNumeric(BoardUserLimited(6)) Then
				If Dvbbs.UserID = 0 Then ShowMobileErr("�������������û���ɾ������ "&BoardUserLimited(6)&" ���ܽ���")
				If Clng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userdel").text)>Clng(BoardUserLimited(6)) Then ShowMobileErr("�������������û���ɾ������ "&BoardUserLimited(6)&" ���ܽ���")
			End If
			'ע��ʱ��
			If Trim(BoardUserLimited(7))<>"0" And IsNumeric(BoardUserLimited(7)) Then
				If Dvbbs.UserID = 0 Then ShowMobileErr("�������������û�ע��ʱ����� "&BoardUserLimited(7)&" ���Ӳ��ܽ���")
				If DateDiff("s",Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@joindate").text,Now)<Clng(BoardUserLimited(7))*60 Then ShowMobileErr("�������������û�ע��ʱ����� "&BoardUserLimited(7)&" ���Ӳ��ܽ���")
			End If
		
		End If
		'��֤����ж�Board_Setting(2)
		If Dvbbs.Board_Setting(2)="1" Then
			Dim Get_BoardUser_Money,Canlogin
			Get_BoardUser_Money = False
			If Clng(Dvbbs.Board_Setting(62))>0 Or Clng(Dvbbs.Board_Setting(63))>0 Then Get_BoardUser_Money = True
			Canlogin = False
			If Dvbbs.UserID=0 Then
				ShowMobileErr("����̳Ϊ��֤��̳����ȷ�������û����Ѿ��õ�����Ա����֤����롣")
			Else
				Dim Boarduser,i,BoardUser_Money
				BoardUser = Dvbbs.boarduser
				If Ubound(Boarduser)=-1 Then	'Ϊ��ʱֵ����-1
					Canlogin = False
				Else
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
			End	If
			If Get_BoardUser_Money And Instr(Lcase(Dvbbs.ScriptName),"pay_boardlimited")=0 And (Not Canlogin) Then
				'Response.Redirect "pay_boardlimited.asp?boardid=" & Dvbbs.BoardID
				ShowMobileErr("����̳Ϊ��֤��̳����ȷ�������û����Ѿ��õ�����Ա����֤����롣")
			End If
			If Instr(Lcase(Dvbbs.ScriptName),"pay_boardlimited")=0 And (Not Canlogin) Then
				ShowMobileErr("����̳Ϊ��֤��̳����ȷ�������û����Ѿ��õ�����Ա����֤����롣")
			End If
		End If
	End Sub
	Public Sub ShowMobileErr(ErrStr)
		ShowXMLStar
		ShowErr 0,ErrStr
		ShowXMLEnd
		Response.End
	End Sub

End Class

Dim DvbbsWap
Set DvbbsWap = New Mobile_Forum
%>