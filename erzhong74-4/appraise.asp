<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<%
Const GuestPost=True  Rem �Ƿ�������֧�ַ���  True False

Dim Action

If Dvbbs.BoardID=0 Then Msg("�������ݴ����벻Ҫ�Ƿ��ύ��")
If Dvbbs.GroupSetting(2)="0" Then Msg("��û��Ȩ�޽������ۣ�")

Action = LCase(Request("action"))
Select Case Action
	Case "post"	:	Post()
	Case "list"	:	list()
End Select

Sub Post()
	Dim Rs,SQL,IsAgree,URs,ErrInfo
	Dim AType,AContent,ATitle,TopicID,PostID,PostTable
	Dim T_LockTopic,P_LockTopic,P_LockUser

	TopicID = Dvbbs.CheckNumeric(Request.Form("topicid"))
	PostID = Dvbbs.CheckNumeric(Request.Form("announceid"))
	AType = Dvbbs.CheckNumeric(Request.Form("atype"))
	'Msg(TopicID&","&PostID&","&AType)
	If TopicID=0 Or PostID=0 Then
		Msg("�������ݴ����벻Ҫ�Ƿ��ύ��")
		Exit Sub
	End If

	'AType:0Ϊ����, 1Ϊ֧��, 2Ϊ����
	If AType=2 Then
		AContent="--����--"
	ElseIf AType=0 Then
		AContent="--����--"
	Else
		AContent="--֧��--"
	End If

	If Not GuestPost Then
		If Dvbbs.UserId=0 Then
			Msg("��û�е�½�����½����֧�֣�лл��")
			Exit Sub
		End If
		Set Rs=Dvbbs.Execute("Select Count(1) From Dv_Appraise Where UserID="&Dvbbs.UserId&" And PostID="&PostID)
	Else
		Dvbbs.MemberName = Dvbbs.UserTrueIP
		Set Rs=Dvbbs.Execute("Select Count(1) From Dv_Appraise Where UserName='"&Dvbbs.UserTrueIP&"' And PostID="&PostID)
	End If
	If Rs(0)>0 Then
		Msg("���Ѿ�����������ӵ����ۣ������ٴ����ۣ�")
		Exit Sub
	End If

	Set Rs=Dvbbs.Execute("Select PostTable,LockTopic,Title From Dv_Topic Where BoardID="&Dvbbs.BoardID&" And TopicID="&TopicID)
	If Rs.Eof Then
		Msg("����ID������Ҫ���۵����ⲻ���ڻ��Ѿ���ɾ��")
		Exit Sub
	End If
	PostTable=Rs(0):T_LockTopic = Rs(1):ATitle="Re:"&Dvbbs.CheckStr(Rs(2))
	Rs.Close
	If CInt(T_LockTopic)>0 Then
		Msg("���ⱻ���������������ۣ�")
		Exit Sub
	End If

	Set Rs=Dvbbs.Execute("Select IsAgree,LockTopic,PostUserID From "&PostTable&" Where BoardID="&Dvbbs.BoardID&" And AnnounceID="&PostID)
	If Rs.Eof Then
		Msg("����ID������Ҫ���۵����Ӳ����ڻ��Ѿ���ɾ��")
		Exit Sub
	End If
	IsAgree = Split(Dvbbs.CheckStr(Rs(0)),"|")
	P_LockTopic=Rs(1)
	If CInt(P_LockTopic)>0 Then
		Msg("���ӱ������������Σ����������ۣ�")
		Exit Sub
	End If
	
	Dvbbs.Execute("Insert Into Dv_Appraise (TopicID,PostID,AType,ATitle,AContent,UserID,UserName,[DateTime],Ip,BoardID) Values ("&TopicID&","&PostID&","&AType&",'"&ATitle&"','"&AContent&"',"&Dvbbs.UserID&",'"&Dvbbs.MemberName&"',"&SqlNowString&",'"&Dvbbs.UserTrueIP&"',"&Dvbbs.BoardID&")")

	If Ubound(IsAgree)<6 Then
		If Ubound(IsAgree)<1 Then
			IsAgree = Split("0|0|||0|0|0","|")
		Else
			IsAgree = Split(Join(IsAgree,"|")&"|||0|0|0","|")
		End If
	End If
	IsAgree(AType+4)=IsAgree(AType+4)+1
	Dvbbs.Execute("UpDate "&PostTable&" Set IsAgree='"&Join(IsAgree,"|")&"' Where BoardID="&Dvbbs.BoardID&" And AnnounceID="&PostID)
	Rem �ύ���
	Msg(TopicID&","&PostID&","&AType)
End Sub

Sub Msg(Str)
	Response.Clear
	Response.write Str
	Response.End
End Sub
%>