<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="Dv_plus/IndivGroup/Dv_IndivGroup_Config.asp"-->
<!--#include file="Dv_plus/IndivGroup/Dv_IndivGroup_MainCls.asp"-->
<%
Dim PostManage
Dvbbs.LoadTemplates("indivgroup")
'Dvbbs.SHOWSQL = 1
If Dv_IndivGroup_MainClass.ID=0 Or Dv_IndivGroup_MainClass.Name="" Then Response.redirect "showerr.asp?ErrCodes=�Բ�������ʵ�Ȧ�Ӳ����ڻ��Ѿ���ɾ��&action=OtherErr"
If Dv_IndivGroup_MainClass.BoardID=0 Then Response.redirect "showerr.asp?ErrCodes=�������ĿID��������ȷ�����Ǵ���Ч�����ӽ��롣&action=OtherErr"

Dv_IndivGroup_MainClass.Stats="���ӹ���"
Dvbbs.Nav()
Dv_IndivGroup_MainClass.Head_var 1,"",""
Set PostManage=New Dv_PostManage_Forum
PostManage.main
Dvbbs.ActiveOnline()
Set PostManage=Nothing
Dvbbs.Footer()
Dvbbs.PageEnd()
Class Dv_PostManage_Forum
	Public TopicID,Topic,TopicUserName,TopicUserID,IsBest,IsTop,PostUserID
	Public IP,ID,ReplyID,ActionInfo,Content,AllMsg
	Public doWealth,douserCP,douserEP,UpdateBoardID,UpdateBoardID_1
	Public Rs,SQL,i
	Private LocalCanLockTopic,LocalCanDelTopic,LocalCanMoveTopic,LocalCanTopTopic,LocalCanBestTopic,LocalCanAwardTopic,LocalCanTopTopic_a,LocalCanTopTopic_m,LocalCanTopicMode
	Public title,sucmsg,LogType
	Public Lasttopic,Lastpost
	Public lastrootID,lastpostuser

	Private Sub Class_Initialize()
		'Response.write Request("ID")&"":Response.end
		TopicID=Dvbbs.CheckNumeric(Request("ID"))
		If TopicID=0 Then Response.redirect "showerr.asp?ErrCodes=<li>�Ƿ������Ӳ�����&action=OtherErr"
		ReplyID=Dvbbs.CheckNumeric(Request("replyID"))

		If ReplyID = 0 Then
			Set Rs=Dvbbs.Execute("Select Title,IsBest,IsTop,PostuserID From Dv_Group_Topic Where GroupID="&Dv_IndivGroup_MainClass.ID&" and BoardID="&Dv_IndivGroup_MainClass.BoardID&" and TopicID="&TopicID)
		Else
			Set Rs=Dvbbs.Execute("Select T.Title,T.IsBest,T.IsTop,B.PostUserID From Dv_Group_Topic T Inner Join Dv_Group_BBS B On T.TopicID = B.RootID Where B.GroupID="&Dv_IndivGroup_MainClass.ID&" and B.BoardID="&Dv_IndivGroup_MainClass.BoardID&" and B.AnnounceID="&ReplyID)
		End If
		If Rs.Eof Then
			Response.redirect "showerr.asp?ErrCodes=<li>�����ҵ������ݲ����ڣ�&action=OtherErr"
		Else
			Topic=rs(0)
			IsBest=Rs(1)
			IsTop=Rs(2)
			PostUserID = Rs(3)
		End If
		Rs.Close:Set Rs=Nothing			
	End Sub

	Public Function Main()
		If Dv_IndivGroup_MainClass.PowerFlag>0 And Dv_IndivGroup_MainClass.PowerFlag<4 Then
			Select Case LCase(Request("action"))
				Case "�޸�"				: fixtopic()		: ActionInfo="�޸�����"
				Case "lock"				: lock()			: ActionInfo="��������"
				Case "unchainlock"		: unchainlock()		: ActionInfo="�������"
				Case "settop"			: SetTop()			: ActionInfo="�̶�����"
				Case "unchaintop"		: UnchainTop()		: ActionInfo="����̶�"
				Case "lockpage"			: lockpage()		: ActionInfo="��������"
				Case "unchainlockpage"	: unchainlockpage()	: ActionInfo="�������"
				Case "setbest"			: setbest()			: ActionInfo="��������"
				Case "unchainbest"		: unchainbest()		: ActionInfo="�������"
				Case "deletetopic"		: DeleteTopic()		: ActionInfo="ɾ����������"
				Case "deletepost"		: DeletePost(1)		: ActionInfo="ɾ������"
				Case "delre"			: Call delre()		: ActionInfo="����ɾ������"
			End Select
		ElseIf Dv_IndivGroup_MainClass.PowerFlag=7 And PostUserID = Dvbbs.UserID Then
			Select Case LCase(Request("action"))
				Case "deletetopic"		: DeleteTopic()		: ActionInfo="ɾ����������"
				Case "deletepost"		: DeletePost(1)		: ActionInfo="ɾ������"
			End Select
		Else
			Response.redirect "showerr.asp?ErrCodes=<li>��û�й���Ȧ�����ӵ�Ȩ�ޡ�&action=OtherErr"
		End If
		If Dvbbs.ErrCodes<>"" Then Dvbbs.ShowErr()
	End Function
	'����ɾ������
	Private Sub delre()
		Check_topicInfo()
		If Dvbbs.ErrCodes<>"" Then Exit Sub
		'�ж��û��Ƿ���ɾ������Ȩ��
		If Not CanDelTopic Then Dvbbs.AddErrCode(28)
		Dim DelID,j,i
		j=0
		DelID=Request("DelID")
		If delid="" Then
			Dvbbs.AddErrCode(35)
			Exit Sub
		End If
		delid=Split(delid,",")
		For i = 0 to UBound(delid)
			If Trim(delid(i))<>"" and IsNumeric(Trim(delid(i))) Then
				j=j+1
				replyID=Ccur(Trim(delid(i)))
				dele(0)
			End If
		Next
		If j>0 Then
			Dvbbs.Dvbbs_Suc(SucMsgInfo("����ɾ��"&j&"������,���Ĳ����Ѿ���¼"))
		Else
			Dvbbs.AddErrCode(35)
		End If
	End Sub

	Public Function Check_topicInfo()
		Set Rs=Dvbbs.Execute("Select topic,username,postuserID From Dv_Group_BBS Where ParentID=0 and boardid="&Dv_IndivGroup_MainClass.BoardID&" and RootID="&TopicID)
		If Rs.Eof And Rs.Bof Then
			Dvbbs.AddErrCode(32)
			Exit Function
		End If
		Topic=rs(0)
		TopicUsername=rs(1)
		TopicUserID=Clng(rs(2))
		Rs.close
	End Function
	'�ж��Ƿ�Ϊ�������ظ�
	Public Function isLastPost()
		Dim LastTopic,body,LastRootID,LastPostTime,LastPostUser
		Dim LastPost,uploadpic_n,LastPostUserID,LastID
		isLastPost=False
		'ȡ�õ�ǰ�������ظ�ID
		Set Rs=Dvbbs.Execute("select LastPost from Dv_Group_topic where topicID="&TopicID)
		If not (rs.eof and rs.bof) Then
			If not isnull(rs(0)) and rs(0)<>"" Then
				If Clng(split(rs(0),"$")(1))=Clng(replyID) Then isLastPost=True
			End If
		End If
		If isLastPost Then
			Set Rs=Dvbbs.Execute("select top 1 topic,body,AnnounceID,dateandtime,username,PostUserID,rootID,boardID from Dv_Group_BBS where BoardID="&Dv_IndivGroup_MainClass.BoardID&" And rootID="&TopicID&" order by AnnounceID desc")
			If not(rs.eof and rs.bof) Then
				body=rs(1)
				LastRootID=rs(2)
				LastPostTime=rs(3)
				LastPostUser=replace(rs(4),"$","")
				LastTopic=left(replace(body,"$",""),20)
				LastPostUserID=rs(5)
				LastID=rs(6)
				Dv_IndivGroup_MainClass.BoardID=rs(7)
			Else
				LastTopic="��"
				LastRootID=0
				LastPostTime=now()
				LastPostUser="��"
				LastPostUserID=0
				LastID=0
				Dv_IndivGroup_MainClass.BoardID=0
			End If
			set rs=nothing
			LastPost=LastPostUser & "$" & LastRootID & "$" & LastPostTime & "$" & replace(left(Dvbbs.Replacehtml(LastTopic),20),"$","") & "$" & uploadpic_n & "$" & LastPostUserID & "$" & LastID & "$" & Dv_IndivGroup_MainClass.BoardID
			Dvbbs.Execute("update Dv_Group_topic set LastPost='"&LastPost&"' where topicID="&TopicID)
		End If
	End Function
	'�����������ظ���Ϣ 2005-1-12 Dv.Yz
	Public Function FixLastPost()
		Dim LastTopic,body,LastRootID,LastPostTime,LastPostUser
		Dim LastPost,uploadpic_n,LastPostUserID,LastID
		Set Rs = Dvbbs.Execute("SELECT TOP 1 Topic, Body, AnnounceID, Dateandtime, Username, PostUserID, RootID, BoardID ,signflag FROM Dv_Group_BBS WHERE GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID = "&Dv_IndivGroup_MainClass.BoardID&" AND RootID = " & TopicID & " ORDER BY AnnounceID DESC")
		If Not Rs.Eof Then
			Body = Rs(1)
			LastRootID = Rs(2)
			LastPostTime = Rs(3)
			LastPostUser = Replace(Rs(4),"$","")
			LastTopic = Left(Replace(Body,"$",""),20)
			LastPostUserID = Rs(5)
			LastID = Rs(6)
		Else
			LastTopic = "��"
			LastRootID = 0
			LastPostTime = Now()
			LastPostUser = "��"
			LastPostUserID = 0
			LastID = 0
		End If
		Rs.Close:Set Rs = Nothing
		LastPost = LastPostUser&"$"&LastRootID&"$"&LastPostTime&"$"&Replace(left(Dvbbs.Replacehtml(LastTopic),20),"$","")&"$$"&LastPostUserID&"$"&LastID&"$"&Dv_IndivGroup_MainClass.BoardID
		LastPost = Dvbbs.CheckStr(LastPost)
		Dvbbs.Execute("UPDATE Dv_Group_Topic SET LastPost = '" & LastPost & "' WHERE GroupID="&Dv_IndivGroup_MainClass.ID&" And TopicID = "&TopicID)
	End Function
	'����ָ����̳��Ϣ
	Public Function LastCount(boardID)
		Dim LastTopic,body,LastRootID,LastPostTime,LastPostUser
		Dim LastPost,uploadpic_n,LastpostuserID,LastID
		Set Rs=Dvbbs.Execute("Select Top 1 Announceid,Dateandtime,Username,Postuserid,Rootid From Dv_Group_BBS Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&BoardID&" Order By Announceid Desc")
		If not(rs.eof and rs.bof) Then
			LastID=rs(0)
			LastPostTime=rs(1)
			LastPostUser=rs(2)
			LastPostUserID=rs(3)
			LastRootID=rs(4)
			Set Rs=Dvbbs.Execute("Select Title From Dv_Group_Topic Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&BoardID&" And TopicID="&LastRootid)
			Lasttopic=Replace(Left(Rs(0),15),"$","")
		Else
			LastTopic="��"
			LastRootID=0
			LastPostTime=now()
			LastPostUser="��"
			LastPostUserID=0
			LastID=0
		End If
		Rs.Close:Set Rs=Nothing

		LastPost=LastPostUser & "$" & LastRootID & "$" & LastPostTime & "$" & LastTopic & "$" & uploadpic_n & "$" & LastPostUserID & "$" & LastID & "$" & BoardID
		Dvbbs.Execute("update Dv_Group_board set LastPost='"&LastPost&"' where RootID="&Dv_IndivGroup_MainClass.ID&" And ID="&Dv_IndivGroup_MainClass.BoardID)
	End Function

	'���淢��������
	Public Sub BoardNumAdd(boardID,topicNum,postNum,todayNum)
		Dvbbs.Execute("update Dv_Group_board set postnum=postnum+"&postNum&",topicNum=topicNum+"&topicNum&",todayNum=todayNum+"&todayNum&" where RootID="&Dv_IndivGroup_MainClass.ID&" And ID="&Dv_IndivGroup_MainClass.BoardID)
	End Sub
	
	'���淢��������
	Public Sub BoardNumSub(boardID,topicNum,postNum,todayNum)
		Dvbbs.Execute("update Dv_Group_board set postnum=postnum-"&postNum&",topicNum=topicNum-"&topicNum&",todayNum=todayNum-"&todayNum&" where RootID="&Dv_IndivGroup_MainClass.ID&" And ID="&Dv_IndivGroup_MainClass.BoardID)
	End Sub
	
	'������̳����������
	Public Function AllboardNumAdd(todayNum,postNum,topicNum)
		Dvbbs.Execute("Update dv_GroupName Set TodayNum=todayNum+"&todaynum&",PostNum=PostNum+"&postNum&",TopicNum=topicNum+"&TopicNum&" where ID="&Dv_IndivGroup_MainClass.ID)
	End Function

	'������̳����������
	Public Function AllboardNumSub(todayNum,postNum,topicNum)
		Dvbbs.Execute("Update dv_GroupName Set TodayNum=todayNum-"&todaynum&",PostNum=PostNum-"&postNum&",TopicNum=topicNum-"&TopicNum&" where ID="&Dv_IndivGroup_MainClass.ID)
	End Function

	Private Function SucMsgInfo(GetMsg)
		SucMsgInfo="<li>"+GetMsg
		SucMsgInfo=SucMsgInfo+"<li>"+"<a href=IndivGroup_index.asp?GroupID="&Dv_IndivGroup_MainClass.ID&"&groupboardid="&Dv_IndivGroup_MainClass.BoardID&">���������б�</a>"
		If LCase(Request("action"))<>"deletetopic" Then
			SucMsgInfo=SucMsgInfo+"<li>"+"<a href=IndivGroup_Dispbbs.asp?GroupID="&Dv_IndivGroup_MainClass.ID&"&groupboardid="&Dv_IndivGroup_MainClass.BoardID&"&id="&TopicID&" >�������ۣ���"&server.htmlencode(Topic)&"��</a>"
		End If
	End Function
	'��������
	Public Sub lock()
		Dvbbs.Execute("Update Dv_Group_topic Set locktopic=1 where GroupID="&Dv_IndivGroup_MainClass.ID&" And boardID="&Dv_IndivGroup_MainClass.BoardID&" And topicID="&TopicID)
		sucmsg=ActionInfo&"��"&Server.htmlencode(topic)&"��"
		Dvbbs.Dvbbs_Suc(SucMsgInfo(sucmsg))
	End Sub
	'�����������
	Public Sub unchainlock()
		Dvbbs.Execute("Update Dv_Group_topic Set locktopic=0 where GroupID="&Dv_IndivGroup_MainClass.ID&" And boardID="&Dv_IndivGroup_MainClass.BoardID&" And topicID="&TopicID)
		sucmsg=ActionInfo&"��"&Server.htmlencode(topic)&"��"
		Dvbbs.Dvbbs_Suc(SucMsgInfo(sucmsg))
	End Sub
	'�̶�����
	Public Sub SetTop()
		Dvbbs.Execute("update Dv_Group_topic set istop=1  where GroupID="&Dv_IndivGroup_MainClass.ID&" And boardID="&Dv_IndivGroup_MainClass.BoardID&" and topicID="&TopicID)
		BoardNumSub Dv_IndivGroup_MainClass.BoardID,1,0,0
		sucmsg=ActionInfo&"��"&Server.htmlencode(topic)&"��"
		Dvbbs.Dvbbs_Suc(SucMsgInfo(sucmsg))
	End Sub
	'����̶�
	Public Sub UnchainTop()
		Dvbbs.Execute("update Dv_Group_topic set istop=0  where GroupID="&Dv_IndivGroup_MainClass.ID&" And boardID="&Dv_IndivGroup_MainClass.BoardID&" and topicID="&TopicID)
		BoardNumAdd Dv_IndivGroup_MainClass.BoardID,1,0,0
		sucmsg=ActionInfo&"��"&Server.htmlencode(topic)&"��"
		Dvbbs.Dvbbs_Suc(SucMsgInfo(sucmsg))
	End Sub

	'������������
	Public Sub lockpage()
		Dvbbs.Execute("Update Dv_Group_BBS Set LockTopic=2 where GroupID="&Dv_IndivGroup_MainClass.ID&" And boardID="&Dv_IndivGroup_MainClass.BoardID&" and announceID="&replyID)
		sucmsg=ActionInfo&"��"&Server.htmlencode(topic)&"��"
		Dvbbs.Dvbbs_Suc(SucMsgInfo(sucmsg))
	End Sub

	'���������������
	Public Sub unchainlockpage()
		Dvbbs.Execute("Update Dv_Group_BBS set LockTopic=0 Where GroupID="&Dv_IndivGroup_MainClass.ID&" And boardID="&Dv_IndivGroup_MainClass.BoardID&" and announceID="&replyID)
		sucmsg=ActionInfo&"��"&Server.htmlencode(topic)&"��"
		Dvbbs.Dvbbs_Suc(SucMsgInfo(sucmsg))
	End Sub

	Public Sub fixtopic()
		Dim UseTools
		Rem Set Rs=dvbbs.Execute("select UseTools from Dv_Group_topic where BoardID="&Dv_IndivGroup_MainClass.BoardID&" And topicid="&TopicID)
		UseTools=""

		Set Rs = Dvbbs.Execute("SELECT COUNT(*), MAX(DateAndTime) FROM Dv_Group_BBS WHERE GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID = " & Dv_IndivGroup_MainClass.BoardID & " AND RootID = " & TopicID)
		If Not IsNull(rs(0)) And Not IsNull(rs(1)) Then
			'If  InStr("," & UseTools & ",",",13,")>0 Or InStr("," & UseTools & ",",",14,")>0 Then
			'	Dvbbs.Execute("update Dv_Group_topic set child="&Rs(0)-1&" where GroupID="&Dv_IndivGroup_MainClass.ID&" And topicID="&TopicID)
			'Else
				Dvbbs.Execute("update Dv_Group_topic set child="&Rs(0)-1&",LastPostTime='"&rs(1)&"' where GroupID="&Dv_IndivGroup_MainClass.ID&" And topicID="&TopicID)
			'End If
			Set Rs=Nothing
		End If
		FixLastPost
		sucmsg=ActionInfo&"��"&Server.htmlencode(topic)&"��"
		Dvbbs.Dvbbs_Suc(SucMsgInfo(sucmsg))
	End Sub
	'��������
	Public Sub setbest()
		Dvbbs.Execute("Update Dv_Group_BBS Set isbest=1 where GroupID="&Dv_IndivGroup_MainClass.ID&" And boardID="&Dv_IndivGroup_MainClass.BoardID&" and announceID="&replyID)
		Dvbbs.Execute("Update Dv_Group_topic Set isbest=1 where GroupID="&Dv_IndivGroup_MainClass.ID&" And boardID="&Dv_IndivGroup_MainClass.BoardID&" and topicID="&TopicID)
		sucmsg=ActionInfo&"��"&Server.htmlencode(topic)&"��"
		Dvbbs.Dvbbs_Suc(SucMsgInfo(sucmsg))
	End Sub
	'�����������
	Public Sub unchainbest()
		Dvbbs.Execute("Update Dv_Group_BBS set isbest=0 Where GroupID="&Dv_IndivGroup_MainClass.ID&" And boardID="&Dv_IndivGroup_MainClass.BoardID&" and announceID="&replyID)
		Dvbbs.Execute("Update Dv_Group_topic set isbest=0 Where GroupID="&Dv_IndivGroup_MainClass.ID&" And boardID="&Dv_IndivGroup_MainClass.BoardID&" and topicID="&TopicID)
		sucmsg=ActionInfo&"��"&Server.htmlencode(topic)&"��"
		Dvbbs.Dvbbs_Suc(SucMsgInfo(sucmsg))
	End Sub

	'ɾ������
	Public Sub DeletePost(md)
		Dim todaynum
		Dim isbest,IsUpload
		todaynum=0
		Set Rs=Dvbbs.Execute("select topic,username,postuserID,DateAndTime,isbest,IsUpload from Dv_Group_BBS where GroupID="&Dv_IndivGroup_MainClass.ID&" And boardid="&Dv_IndivGroup_MainClass.BoardID&" and AnnounceID="&replyID)
		If Not Rs.Eof Then
			Topic=Dvbbs.CheckStr(Rs(0))
			topicusername=rs(1)
			topicuserID=rs(2)
			isbest=rs(4)
			If topic="" Then topic="������Ϊ�ظ�����"
			If datediff("d",rs(3),now())=0 Then
				todaynum=1
			Else
				todaynum=0
			End If
		Else
			If md=1 Then
				Dvbbs.AddErrCode(32)
				Exit Sub
			End If
		End If
		Set Rs=Nothing
		
		Dim LastPostime,istop
		'ɾ��ʱ�Զ�ɾ�������ظ���
		If IsBest=1 Then Dvbbs.Execute("update Dv_Group_topic set isbest=0 where GroupID="&Dv_IndivGroup_MainClass.ID&" And boardid="&Dv_IndivGroup_MainClass.BoardID&" and topicid="&TopicID)
		Dvbbs.Execute("Update Dv_Group_BBS Set BoardID=-1,locktopic="&Dv_IndivGroup_MainClass.BoardID&" Where GroupID="&Dv_IndivGroup_MainClass.ID&" And BoardID="&Dv_IndivGroup_MainClass.BoardID&" And AnnounceID="&replyID)
		Set Rs=Dvbbs.Execute("select Max(dateandtime) from Dv_Group_BBS where boardID="&Dv_IndivGroup_MainClass.BoardID&" and rootID="&TopicID)
		LastPostime=rs(0)
		Rs.Close:Set Rs=Nothing
		isLastPost
		call LastCount(Dv_IndivGroup_MainClass.BoardID)
		call BoardNumSub(Dv_IndivGroup_MainClass.BoardID,0,1,todaynum)
		call AllboardNumSub(todaynum,1,0)
		Dvbbs.ReloadBoardInfo(UpdateBoardID)

		sql="update Dv_Group_topic set child=child-1,LastPostTime='"&LastPostime&"' where boardID="&Dv_IndivGroup_MainClass.BoardID&" and topicID="&TopicID
		'Response.Write sql
		Dvbbs.Execute(sql)
		If md=1 Then
			sucmsg=ActionInfo&"��"&Server.htmlencode(topic)&"��"
			Dvbbs.Dvbbs_Suc(SucMsgInfo(sucmsg))
		End If 
	End Sub
	'ɾ������
	Public Sub DeleteTopic()
		Dim todaynum,postnum
		Set Rs=Dvbbs.Execute("select count(*) from Dv_Group_BBS where GroupID="&Dv_IndivGroup_MainClass.ID&" And rootID="&TopicID)
		postNum=Rs(0)
		If IsSqlDataBase=1 Then
			sql="select count(*) from Dv_Group_BBS where GroupID="&Dv_IndivGroup_MainClass.ID&" And rootID="&TopicID&" and dateandtime>'"&date()&"'"
		else
			sql="select count(*) from Dv_Group_BBS where GroupID="&Dv_IndivGroup_MainClass.ID&" And rootID="&TopicID&" and dateandtime>#"&date()&"#"
		end if
		Set Rs=Dvbbs.Execute(sql)
		todayNum=rs(0)
	
		'�������վ������վboardidΪ-1��locktopicΪԭ����ID
		Dvbbs.Execute("update Dv_Group_BBS set BoardID=-1,locktopic="&Dv_IndivGroup_MainClass.BoardID&" where GroupID="&Dv_IndivGroup_MainClass.ID&" And RootID="&TopicID)
		Dvbbs.Execute("update Dv_Group_topic set BoardID=-1,locktopic="&Dv_IndivGroup_MainClass.BoardID&",isbest=0,istop=0 where GroupID="&Dv_IndivGroup_MainClass.ID&" And topicid="&TopicID)

		call LastCount(Dv_IndivGroup_MainClass.BoardID)
		call BoardNumSub(Dv_IndivGroup_MainClass.BoardID,1,postNum,todayNum)
		call AllboardNumSub(todayNum,postNum,1)
		sucmsg=ActionInfo&"��"&Server.htmlencode(topic)&"��"
		Dvbbs.Dvbbs_Suc(SucMsgInfo(sucmsg))
	End Sub
End Class
%>
