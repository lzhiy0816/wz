<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/dv_ubbcode.asp"-->
<!--#include file="inc/dv_template.inc"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/code_encrypt.asp"-->

<%if not session("checked")="ez74yes" then
response.Redirect "ez74login.asp"

else

If Dvbbs.BoardID < 1 Then Response.Write "��������":Response.End
If Request("page") <> "" And CStr(Dvbbs.CheckNumeric(Request("page"))) <> Request("page") Then
    Response.Write "��������"
    Response.End
End If
If Dvbbs.GroupSetting(2)="0" Then Dvbbs.AddErrcode(31):Dvbbs.ShowErr():response.End
Dim PostUserid, G_TopicTitle, G_IsVote, G_Childs, G_PollID, G_LockTopic, G_Hits, G_Expression,FlashId
Dim G_ItemList, G_ItemsPerPage, G_CurrentPage, G_Pages, G_Moved
Dim G_UserList
Dim G_UserItemQuery
Dim G_Floor
Dim G_CanReply
Dim Dv_ubb
Dim CanRead,TrueMaster,Skin
'���¶���ı�����Dv_ubbcode.aspҳ����õ�
Dim EmotPath
Dim TotalUsetable
Dim PostBuyUser
Dim UserName
Dim T_GetMoneyType
Dim AnnounceID, ReplyID, Replyid_a, AnnounceID_a, RootID_a
Dim IsThisBoardMaster 'ȷ����ǰ�û��Ƿ񱾰��������ֹ����Ĳ���Ӱ�쵽 Dvbbs.BoardMaster���³���
IsThisBoardMaster = Dvbbs.BoardMaster
'���������Ȩ��
CanRead=False
TrueMaster=False
Rem Ϊ��˹����˵���ʾ,���й���Ȩ�޵���ʱ�������ȼ�����,Ϊ������ʾ�����˵�.
If Not Dvbbs.BoardMaster Then
	If Dvbbs.UserID > 0 Then
		If Dvbbs.GroupSetting(18) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(19) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(20) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(21) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(22) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(23) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(24) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(25) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(26) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(27) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(28) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(29) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(30) = "1" Then
			Dvbbs.BoardMaster=True
		ElseIf Dvbbs.GroupSetting(31) = "1" Then
			Dvbbs.BoardMaster=True
		End If
	End If
Else
	TrueMaster=True
End If
If Dvbbs.BoardMaster Then CanRead=True

'��ʼ����
AnnounceID		= 0			'����ID
G_UserItemQuery = "userid,username,useremail,userpost,usertopic,usersign,usersex,userface,userwidth,userheight,joindate,lastlogin,userlogins,lockuser,userclass,userwealth,userep,usercp,userpower,userdel,userisbest,usertitle,userhidden,usermoney,userticket,titlepic,usergroupid,userim,useremail" '��ѯ�û����ֶ��б�
'0-userid,1-username,2-useremail,3-userpost,4-usertopic,5-usersign,6-usersex,7-userface,8-userwidth,9-userheight,10-joindate,11-lastlogin,12-userlogins,13-lockuser,14-userclass,15-userwealth,16-userep,17-usercp,18-userpower,19-userdel,20-userisbest,21-usertitle,22-userhidden,23-usermoney,24-userticket,25-titlepic,26-UserGroupID,27-userim
LoadTopicInfo
LoadBBSListData
LoadUserListData
Dvbbs.LoadTemplates("dispbbs")
Dvbbs.Nav()
Dvbbs.Head_var 1,"","",""
If UserFlashGet = 1 Then
	%><!--#include file="Dv_plus/Flashget/Flashget_base64.asp"--><%
	Response.Write "<script src=""http://ufile.kuaiche.com/Flashget_union.php?fg_uid="&FlashGetID&"""></script>"
End If
Response.Write GetForumTextAd(2)
Dvbbs.ActiveOnline()

EmotPath=Split(Dvbbs.Forum_emot,"|||")(0)		'em����·��
Set Dv_ubb=new Dvbbs_UbbCode
Dv_ubb.PostType=1
TPL_Scan	Template.html(0)'Dvbbs.ReadTextFile("dispbbsnew.tpl")'
TPL_Flush
Set Dv_ubb=Nothing

Sub LoadTopicInfo()
	AnnounceID	= Dvbbs.CheckNumeric(Request("ID"))
	If 0=AnnounceID Then Dvbbs.AddErrCode(30):Dvbbs.Showerr()
	G_CurrentPage = Dvbbs.CheckNumeric(Request("star"))
	If 0=G_CurrentPage Then G_CurrentPage=1
	Skin=Request("Skin")
	If Skin="" Or Not IsNumeric(Skin) Then Skin=Dvbbs.Board_setting(42)
	Dim Rs, SQL, iLockSet, iTopicMode, sMove
	sMove		= request("move")
	iLockSet	= Dvbbs.CheckNumeric(Dvbbs.Board_Setting(71))
	SQL="Select top 1 TopicID,boardid,title,hits,isvote,child,pollid,LockTopic,PostTable,TopicMode,DateAndTime,Expression,GetMoneyType,PostUserid From Dv_topic where "
	If ""=sMove Then
		SQL	= SQL & ("topicID=" & AnnounceID)
	Else
		SQL	= SQL & "BoardID=" & Dvbbs.BoardID & " and topicID"
		If "next"=sMove Then
			SQL	= SQL & ("<" & AnnounceID & " order by topicID desc")
		Else
			SQL	= SQL & (">" & AnnounceID & " order by topicID")
		End If
	End If
	If Not IsObject(Conn) Then ConnectionDatabase
	Set Rs=Dvbbs.iCreateObject("Adodb.RecordSet")
	Rs.Open SQL,Conn,1,3
	Dvbbs.SqlQueryNum=Dvbbs.SqlQueryNum+1
	If Rs.eof Or Rs.bof Then
		If ""<>sMove Then
			Response.Write "<script language=""javascript"">alert(""�Ѿ������һ�������ˣ�"");history.go(-1);</script>"
			Rs.Close
			Set Rs=Nothing
			Dvbbs.PageEnd
			Response.End
		Else
			Dvbbs.AddErrcode(32)
		End If
	Else
		If CStr(Rs("BoardID"))<>CStr(Dvbbs.BoardID) Then Dvbbs.AddErrCode(29)
		G_Hits			= Dvbbs.CheckNumeric(Rs("hits"))
		Rs("hits")		= G_Hits+1
		G_LockTopic		= Rs("LockTopic")
		If 0=G_LockTopic And iLockSet<>0 And Datediff("d", Rs("DateAndTime"),Now())>iLockSet Then
			G_LockTopic = 1
			Rs("LockTopic")	= G_LockTopic
		End If
		On Error Resume Next
		Rs.Update
		If Err Then Err.Clear
		G_TopicTitle	= Rs("title")
		G_IsVote		= Rs("isvote")
		G_Childs		= Rs("child")
		G_PollID		= Rs("pollid")
		G_Expression	= Rs("Expression")
		T_GetMoneyType	= Rs("GetMoneyType")
		TotalUsetable	= Rs("PostTable")
		iTopicMode		= Rs("topicmode")
		AnnounceID		= Rs("TopicID")
		PostUserid		= Rs("PostUserid")
	End If
	Rs.Close
	Set Rs=Nothing
	Dvbbs.Showerr()
	G_TopicTitle		= Dvbbs.ChkBadWords(G_TopicTitle)
	ReplyID			= Dvbbs.CheckNumeric(Request("ReplyID"))
	If 0=ReplyID Then ReplyID=AnnounceID
	If iTopicMode<>1 Then G_TopicTitle=replace(G_TopicTitle, "<", "&lt;")
	Dvbbs.Stats			= G_TopicTitle
	G_Childs			= G_Childs+1
	Select Case iTopicMode
		Case 2	G_TopicTitle	= "<font color=""red"">"&G_TopicTitle&"</font>"
		Case 3	G_TopicTitle	= "<font color=""blue"">"&G_TopicTitle&"</font>"
		Case 4	G_TopicTitle	= "<font color=""green"">"&G_TopicTitle&"</font>"
		Case Else
	End Select
End Sub

Sub LoadBBSListData()
	G_ItemsPerPage	= Dvbbs.CheckNumeric(Dvbbs.Board_Setting(27))
	G_Pages	= G_Childs \ G_ItemsPerPage
	If (G_Childs Mod G_ItemsPerPage)<>0 Then G_Pages = G_Pages + 1
	If G_Pages<=0 Then G_Pages = 1
	If G_CurrentPage > G_Pages Then G_CurrentPage = G_Pages
	G_Moved	= G_ItemsPerPage*(G_CurrentPage-1)
	Dim Rs, SQL, Cmd
	If 1=Skin Then
		If ReplyID=AnnounceID Then
			SQL="Select Top 1 AnnounceID,UserName,Topic,dateandtime,body,Expression,ip,RootID,signflag,isbest,PostUserid,layer,isagree,GetMoneyType,IsUpload,Ubblist,LockTopic,GetMoney,UseTools,PostBuyUser,ParentID,FlashId From "& TotalUseTable &" where RootID="& AnnounceID &" and Boardid="& Dvbbs.Boardid&" "
		Else
			SQL="Select AnnounceID,UserName,Topic,dateandtime,body,Expression,ip,RootID,signflag,isbest,PostUserid,layer,isagree,GetMoneyType,IsUpload,Ubblist,LockTopic,GetMoney,UseTools,PostBuyUser,ParentID,FlashId From "& TotalUseTable &" where AnnounceID="&ReplyID&" and Boardid="& Dvbbs.Boardid&" "
		End If
	Else
		SQL="Select AnnounceID,UserName,Topic,dateandtime,body,Expression,ip,RootID,signflag,isbest,PostUserid,layer,isagree,GetMoneyType,IsUpload,Ubblist,LockTopic,GetMoney,UseTools,PostBuyUser,ParentID,FlashId From "& TotalUsetable &" where  RootID="& ReplyID &" and Boardid="& Dvbbs.Boardid&" Order By Announceid" '0-AnnounceID,1-UserName,2-Topic,3-dateandtime,4-body,5-Expression,6-ip,7-RootID,8-signflag,9-isbest,10-PostUserid,11-layer,12-isagree,13-GetMoneyType,14-IsUpload,15-Ubblist,16-LockTopic,17-GetMoney,18-UseTools,19-PostBuyUser,20-ParentID
	End If
	If IsSqlDataBase=1 And IsBuss=1 And Skin=0 Then
		Set Cmd = Dvbbs.iCreateObject("ADODB.Command")
		Set Cmd.ActiveConnection=conn
		Cmd.CommandText="dv_disp"
		Cmd.CommandType=4
		Cmd.Parameters.Append cmd.CreateParameter("@boardid",3)
		Cmd.Parameters.Append cmd.CreateParameter("@rootid",3)
		Cmd.Parameters.Append cmd.CreateParameter("@pagenow",3)
		Cmd.Parameters.Append cmd.CreateParameter("@pagesize",3)
		Cmd.Parameters.Append cmd.CreateParameter("@totalusetable",200,1,20)
		Cmd("@boardid")=Dvbbs.BoardID
		Cmd("@rootid")=ReplyID
		Cmd("@pagenow")=G_CurrentPage
		Cmd("@pagesize")=G_ItemsPerPage
		Cmd("@totalusetable")=TotalUsetable
		Set Rs=Cmd.Execute
		If Not Rs.EoF Then
			G_ItemList=Rs.GetRows(-1)
		Else
			Dvbbs.AddErrCode(29)
		End If
		Rs.close()
		Set Rs=Nothing
		Set Cmd =  Nothing
	Else
		Set Rs=Dvbbs.Execute(SQL)
		If Rs.eof Or Rs.bof Then
			Dvbbs.AddErrCode(29)
		Else
			On Error Resume Next
			If 1<>Skin Then Rs.Move(G_Moved)
			If Err Then Err.clear
			If Not Rs.eof Then
				G_ItemList = Rs.GetRows(G_ItemsPerPage)
			Else
				Dvbbs.AddErrCode(29)
			End If
		End If
		Rs.Close
		Set Rs=Nothing
	End If
	Dvbbs.Showerr()
	G_CanReply=False '�Ƿ������ظ�
	If Not Dvbbs.Board_Setting(0)="1"  And Cint(G_LockTopic)=0 Then
		If Dvbbs.GroupSetting(5)="1" Then
			G_CanReply=True
		ElseIf Dvbbs.UserID = PostUserid and  Dvbbs.GroupSetting(4)="1" Then
			G_CanReply=True
		ElseIf Dvbbs.master Or Dvbbs.superboardmaster Or Dvbbs.boardmaster Then
			G_CanReply=True
		End If
	End If
End Sub

Sub LoadUserListData()
	If IsArray(G_UserList) Or Not IsArray(G_ItemList) Then Exit Sub
	Dim Rs, i, j, iTempUserID, iUbd, sUserIDList
	iUbd		= UBound(G_ItemList,2)
	sUserIDList	= G_ItemList(10, 0)
	For i=0 To iUbd
		sUserIDList	= sUserIDList & ("," & G_ItemList(10,i))
	Next
	Set Rs		= Dvbbs.Execute("Select " & G_UserItemQuery & " From dv_user Where UserID IN ("& sUserIDList &")")
	If Rs.Eof Or Rs.Bof Then
		'ȫ���ǿ���
	Else
		G_UserList	= Rs.GetRows(-1)
	End If
	Rs.Close
	Set Rs		= Nothing
	'�����û�����
	For i=0 To iUbd
		iTempUserID			= G_ItemList(10, i)
		G_ItemList(10, i)	= 0	'��ʼΪ�ο�
		If IsArray(G_UserList)	Then
			For j=UBound(G_UserList,2) To 0 Step -1
				If G_UserList(0, j)=iTempUserID Then
					G_ItemList(10, i)	= j+1	'�������1��ʵ����ʱҪ��1
					Exit For
				End If
			Next
		End If
	Next
End Sub

Sub LoadAndParseVote(sTemplate)
	Dim Rs,aVote,s,a1,a2,u1,u2,i,j,t,sLoop
	Dim votetype,votchilds,votchilds_title,votchilds_ep
	Set Rs=Dvbbs.Execute("Select voteid,vote,votenum,votetype,lockvote,voters,timeout,uarticle,uwealth,uep,ucp,upower From Dv_Vote Where VoteID="&G_PollID)
	If Not Rs.eof Then
		aVote=Rs.GetRows(-1)
	Else
		Exit Sub
	End If
	Set Rs=Nothing
	s=sTemplate
	s=Replace(s,"{$showvote.voteid}",aVote(0,0))
	votetype=aVote(3,0)
	s=Replace(s,"{$showvote.lockvote}",aVote(4,0))
	s=Replace(s,"{$showvote.voters}",aVote(5,0))
	s=Replace(s,"{$showvote.timeout}",aVote(6,0))
	s=Replace(s,"{$showvote.uarticle}",aVote(7,0))
	s=Replace(s,"{$showvote.uwealth}",aVote(8,0))
	s=Replace(s,"{$showvote.uep}",aVote(9,0))
	s=Replace(s,"{$showvote.ucp}",aVote(10,0))
	s=Replace(s,"{$showvote.upower}",aVote(11,0))
	If 0=Dvbbs.userid Then
		s=Replace(s,"{$showvote.input}","����δ��¼�����ܲ��롣")
	Else
		If datediff("d",aVote(6,0),Now()) > 0 Then
			s=Replace(s,"{$showvote.input}","�ѹ��ڣ����ܲ��롣")
		Else
			If G_LockTopic Then
				s=Replace(s,"{$showvote.input}","��������Ѿ����������ܲ��롣")
			Else
				If Not Dvbbs.Execute("Select * From Dv_voteuser Where voteid="& G_PollID &" And userid="& Dvbbs.userid).EOF Then
					s=Replace(s,"{$showvote.input}","���Ѿ�Ͷ��Ʊ�ˣ�������ɣ�")
				Else
					s=Replace(s,"{$showvote.input}","<input type=""submit"" name=""VoteSubmit"" value=""Ͷ Ʊ"" style=""margin:5px;""/>")
				End If
			End If
		End If
	End If
	a1=Split(Dvbbs.ChkBadWords(aVote(1,0)),"|")
	a2=Split(aVote(2,0),"|")
	u1=UBound(a1)
	t=0
	For i=0 To u1
		sLoop=sLoop&"<tr>"
		votchilds = Split(a1(i),"@@")
		If votetype = 2 and Ubound(votchilds)>=3 Then
			sLoop=sLoop&("<td width=""40%"">"&(i+1)&". "&votchilds(0)&"</td>") 'title
			votchilds_title = Split(votchilds(2),"$$")
			votchilds_ep = Split(votchilds(3),"$$")
			sLoop=sLoop&"<td>"
			For j=0 To UBound(votchilds_title)
				If ""<>votchilds_title(j) Then
					Select Case votchilds(1)
						Case "1"
							sLoop=sLoop&" <input type=""checkbox"" name=""postvote_"&i&""" value="""&j&""" class=""chkbox""/>"&votchilds_title(j)&" "
						Case "2"
							sLoop=sLoop&" �ش�<textarea name=""postvote_"&i&""" style=""width:70%;height:80px;""></textarea> "
						Case Else
							sLoop=sLoop&" <input type=""radio"" name=""postvote_"&i&""" value="""&j&""" style=""border:none;""/>"&votchilds_title(j)&" "
					End Select
				End If
			Next
			sLoop=sLoop&"</td>"
		Else
			sLoop=sLoop&"<td width=""20%"">"
			If 0=aVote(3,0) Then
				sLoop=sLoop&"<input type=""radio"" name=""postvote"" value="""&i&""" style=""border:none;""/>"
			Else
				sLoop=sLoop&"<input type=""checkbox"" name=""postvote_"&i&""" value="""&i&""" style=""border:none;""/>"
			End If
			sLoop=sLoop&(a1(i)&"</td><td width=""80%"" valign=""top"">") 'title
			sLoop=sLoop&"<script language=""javascript"">try{ShowVoteList("&a2(i)&",{$total});}catch(e){}</script>"
			sLoop=sLoop&("</td>")
			t=t+a2(i)
		End If
		sLoop=sLoop&"</tr>"
	Next
	sLoop=Replace(sLoop,"{$total}",t)
	s=Replace(s,"{$showvote.list}",sLoop)
	TPL_Scan s
End Sub

Sub ParsePageNode(sToken)
	Dim a,i,s
	Select Case sToken
		Case "topicid"
			TPL_Echo AnnounceID
		Case "announceid"
			TPL_Echo G_ItemList(0,0)
		Case "topic"
			TPL_Echo G_TopicTitle
		Case "hits"
			TPL_Echo G_Hits
		Case "currentpage"
			TPL_Echo G_CurrentPage
		Case "boardpage"
			TPL_Echo Dvbbs.CheckNumeric(request("page"))
		Case "bbstable"
			TPL_Echo TotalUsetable
		Case "postfacelist"
			a=split(Dvbbs.Forum_PostFace,"|||")
			s=s&"<div style=""float:left;""><input id=""face_1_gif"" type=""radio"" value=""face1.gif"" name=""Expression"" checked=""checked"" style=""border:none"" /><img src=""Skins/default/topicface/face1.gif"" alt="""" /></div>"
			For i=2 To UBound(a)-1
				s=s&"<div style=""float:left;""><input type=""radio"" value=""face"&i&".gif"" name=""Expression"" style=""border:none"" /><img src="""&a(0)&"face"&i&".gif"" alt="""" /></div>"
				If 1=(i-2) Mod 3 Then s=s&"<br style=""clear:both"" />"
			Next
			TPL_Echo s
		Case "modelink"
			If 1=Skin Then
				TPL_Echo "<a href=""dispbbs.asp?BoardID="&Dvbbs.Boardid&"&ID="&AnnounceID&"&skin=0"">ƽ��</a>"
			Else
				TPL_Echo "<a href=""dispbbs.asp?BoardID="&Dvbbs.Boardid&"&replyID="&G_ItemList(0,0)&"&ID="&AnnounceID&"&skin=1"">����</a>"
			End If
		Case "treemode"
			If 1=Skin Then TPL_Echo "<div id=""postlist"" style=""margin-top : 10px; margin-bottom : 10px; ""> </div><span id=""showpagelist""></span><iframe style=""border:0px;width:0px;height:0px;"" src=""loadtree.asp?boardid="&Dvbbs.Boardid&"&amp;star="&G_CurrentPage&"&amp;replyid="&ReplyID&"&amp;id="&AnnounceID&"&amp;openid="&ReplyID&""" name=""hiddenframe""></iframe>"
		Case "topicadminlist"
			s=""
			If Dvbbs.Boardmaster Then
				If 1=T_GetMoneyType Then
					s=s& "		<a href=""BuyPost.asp?Action=Close&amp;BoardID="&Dvbbs.boardid&"&amp;PostTable="&TotalUseTable&"&amp;ID="&AnnounceID&"&amp;ReplyID="&ReplyID&""" title=""��������"">��������</a><br />"
				End If
				s=s& "	<a href=""admin_postings.asp?action=ר�����&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""ר�����"">ר�����</a><br />"
				If 1=G_LockTopic Then
					s=s& "		<a href=""admin_postings.asp?action=����&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""��������⿪����"">�������</a><br />"
				Else
					s=s& "		<a href=""admin_postings.asp?action=����&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""����������"">��������</a><br />"
				End If
				s=s& "	<a href=""admin_postings.asp?action=����&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""�������������������б���ǰ��"">��������</a><br />"
				s=s& "	<a href=""admin_postings.asp?action=����&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""��������ŵ������б��Ͽ���λ��"">��������</a><br />"
				s=s& "	<a href=""admin_postings.asp?action=��������&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""����ɾ��������ĸ���"">��������</a><br />"
				s=s& "	<a href=""admin_postings.asp?action=ɾ������&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""ע�⣺��������ɾ���������������ӣ����ָܻ�"">ɾ������</a><br />"
				s=s& "	<a href=""admin_postings.asp?action=�ƶ�&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&"&amp;replyID="&ReplyID&""" title=""�ƶ�����"">�ƶ�����</a><br />"
				s=s& "	<a href=""admin_postings.asp?action=���ù̶�&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""�����������ù̶�"">���ù̶�</a><br />"
				TPL_Echo "<div class=""m_li_top"" style=""display:inline;"" onmouseover=""showmenu1('Menu_0',0);""><a href=""#"">�������</a>"
				TPL_Echo "	<div class=""submenu submunu_popup"" id=""Menu_0"" onmouseout=""hidemenu1()"">"
				If ""<>s Then TPL_Echo s
				TPL_Echo "	</div></div>"
			ElseIf IsSelfPost() Then
				If 1=T_GetMoneyType Then
					s=s& "		<a href=""BuyPost.asp?Action=Close&amp;BoardID="&Dvbbs.boardid&"&amp;PostTable="&TotalUseTable&"&amp;ID="&AnnounceID&"&amp;ReplyID="&ReplyID&""" title=""��������"">��������</a><br />"
				End If
				If "1"=Dvbbs.GroupSetting(13) Then
					If 1=G_LockTopic Then
						s=s& "		<a href=""admin_postings.asp?action=����&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""��������⿪����"">�������</a><br />"
					Else
						s=s& "		<a href=""admin_postings.asp?action=����&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""����������"">��������</a><br />"
					End If
				End If
				If "1"=Dvbbs.GroupSetting(11) Then
					s=s& "<a href=""admin_postings.asp?action=��������&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""����ɾ��������ĸ���"">��������</a><br />"
					s=s& "<a href=""admin_postings.asp?action=ɾ������&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&""" title=""ע�⣺��������ɾ���������������ӣ����ָܻ�"">ɾ������</a><br />"
				End If
				If "1"=Dvbbs.GroupSetting(12) Then
					s=s& "<a href=""admin_postings.asp?action=�ƶ�&amp;BoardID="&Dvbbs.boardid&"&amp;ID="&AnnounceID&"&amp;replyID="&ReplyID&""" title=""�ƶ�����"">�ƶ�����</a><br />"
				End If
				If ""<>s Then
					TPL_Echo "<div class=""m_li_top"" style=""display:inline;"" onmouseover=""showmenu1('Menu_0',0);""><a href=""jascript:;"">�������</a><div class=""submenu submunu_popup"" id=""Menu_0"" onmouseout=""hidemenu1()"">"
					TPL_Echo s
					TPL_Echo "</div></div>"
				End If
			End If
		Case Else
	End Select
End Sub

Sub ParseUserNode(sToken)
	Dim i, p, s, s2, bShowAll
	i		= G_ItemList(10, G_Floor)
	bShowAll	= False
	If i>0 Then bShowAll = 2<>G_ItemList(8, G_Floor) Or Dvbbs.BoardMaster Or Dvbbs.UserID=G_UserList(0, G_ItemList(10, G_Floor)-1)
	'			  �ǿ����� ���� (��������			����	 �ǹ���Ա	����	���Լ�)
	If bShowAll Then
		Select Case sToken
			Case "userid"		TPL_Echo G_UserList(0, i-1)
			Case "username"		TPL_Echo UserName
			Case "richname"
				s	= G_UserList(26, i-1)
				Select Case s
					'Case 0	s2= Dvbbs.mainsetting(9)
					Case 1	s2= Dvbbs.mainsetting(9)
					Case 2	s2= Dvbbs.mainsetting(7)
					Case 3	s2= Dvbbs.mainsetting(7)
					Case 4	s2= Dvbbs.mainsetting(5)
					Case 5	s2= "gray"
					Case 6	s2= "gray"
					Case 7	s2= Dvbbs.mainsetting(5)
					Case 8	s2= Dvbbs.mainsetting(11)
					Case Else s2= Dvbbs.mainsetting(5)
				End Select
				s=Split(Split(Application(Dvbbs.CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& G_UserList(26, i-1) &"']/@groupsetting").text,",")(58),"��")
				TPL_Echo "<span style=""width:105px;filter:glow(color='"&s2&"',strength='2');"">"&(s(0)&replace(UserName,chr(255),"")&s(1))
				If 2=G_ItemList(8, G_Floor) Then TPL_Echo "&nbsp;&nbsp;[������]"
				TPL_Echo "</span>"
			Case "useremail"	TPL_Echo G_UserList(2, i-1)
			Case "userpost"		TPL_Echo G_UserList(3, i-1)
			Case "usertopic"	TPL_Echo G_UserList(4, i-1)
			Case "usersign"
				s	= G_UserList(5, i-1)
				p	= G_UserList(26, i-1)
				If ""<>s And 1=G_ItemList(8, G_Floor) And Dvbbs.forum_setting(42)="1" Then
					Set s2	= Application(Dvbbs.CacheName &"_groupsetting")
					If Not s2 Is Nothing Then
						If Application(Dvbbs.CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& p &"']/@groupsetting") Is Nothing Then p	= 7
						If Split(Application(Dvbbs.CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& p &"']/@groupsetting").text,",")(55)	Then TPL_Echo	"<img src=""images/sigline.gif"" alt="""" /><br />" & Dvbbs.ChkBadWords(Dv_ubb.Dv_SignUbbCode(s, p))
					End If
					Set s2	= Nothing
				End If
			Case "usersex"
				s2=G_UserList(11, i-1)
				If "2"=G_UserList(22, i-1) Then
					If IsDate(s2) Then
						If DateDiff("s",s2,Now())>(cCur(dvbbs.Forum_Setting(8))*60) Then
							G_UserList(22, i-1)="1"
						End If
					Else
							G_UserList(22, i-1)="1"
					End If
				Else
					G_UserList(22, i-1)="1"
				End If
				s	= G_UserList(6, i-1)
				If "1"=G_UserList(22, i-1) Then
					Select Case s
						Case "1"
							TPL_Echo	"<img src=""Skins/Default/ofMale.gif"" alt=""˧��Ӵ�����ߣ�����������"" />"
						Case Else
							TPL_Echo	"<img src=""Skins/Default/ofFeMale.gif"" alt=""��Ůѽ�����ߣ����Ը��Ұɣ�"" />"
					End Select
				Else
					Select Case s
						Case "1"
							TPL_Echo	"<img src=""Skins/Default/Male.gif"" alt=""˧�磬�����ޣ�"" />"
						Case Else
							TPL_Echo	"<img src=""Skins/Default/FeMale.gif"" alt=""��Ůѽ�����ߣ��������Ұɣ�"" />"
					End Select
				End If
			Case "userface"
				s2	= s2 &	"<img"
				s2	= s2 &	(" width="""&G_UserList(8, i-1)&"""")
				s2	= s2 &	(" height="""&G_UserList(9, i-1)&"""")
				s2	= s2 &	" src="""
				s	= Dv_FilterJS(G_UserList(7, i-1))
				p	= InStr(s, "|")
				If p>0 Then
					s2	= s2 &	Mid(s, p+1)
					s	= Left(s, p-1)
				Else
					s2	= s2 &	s
					s	= "0"
				End If
				s2	= s2 &	""" alt="""" />"
				If "0"<>s Then	s2	= s2 &	("<br/><div><a href=""javascript:DispMagicEmot('"&s&"',350,500)"">�鿴ħ��ͷ��</a></div>")
				TPL_Echo	s2
			Case "joindate"		TPL_Echo G_UserList(10, i-1)
			Case "lastlogin"	TPL_Echo G_UserList(11, i-1)
			Case "userlogins"	TPL_Echo G_UserList(12, i-1)
			Case "lockuser"		TPL_Echo G_UserList(13, i-1)
			Case "userclass"	TPL_Echo G_UserList(14, i-1)
			Case "userwealth"	TPL_Echo G_UserList(15, i-1)
			Case "userep"		TPL_Echo G_UserList(16, i-1)
			Case "usercp"		TPL_Echo G_UserList(17, i-1)
			Case "userpower"
				s	= G_UserList(18, i-1)
				If "0"<>s Then
					TPL_Echo "<b><font color=""red"">" & s & "</font></b>"
				Else
					TPL_Echo s
				End If
			Case "userdel"		TPL_Echo G_UserList(19, i-1)
			Case "userisbest"	TPL_Echo G_UserList(20, i-1)
			Case "usertitle"	TPL_Echo G_UserList(21, i-1)
			Case "usermoney"	TPL_Echo G_UserList(23, i-1)
			Case "userticket"	TPL_Echo G_UserList(24, i-1)
			Case "titlepic"
				s2	= s2 &	"<img src=""skins/Default/star/"
				s2	= s2 &	G_UserList(25, i-1)
				s2	= s2 &	""" alt="""" />"
				TPL_Echo	s2
			Case "qq"
				If Not IsArray(G_UserList(27, i-1)) Then	G_UserList(27, i-1)=Split(G_UserList(27, i-1), "|||")
				TPL_Echo	G_UserList(27, i-1)(1)
			Case "link_qq"
				If Not IsArray(G_UserList(27, i-1)) Then	G_UserList(27, i-1)=Split(G_UserList(27, i-1), "|||")
				If ""<>G_UserList(27, i-1)(1) Then TPL_Echo "	 | <a href=""tencent://message/?uin="&G_UserList(27, i-1)(1)&""" title=""�������QQ��Ϣ��"&UserName&""">QQ</a>"
			Case "email"
				TPL_Echo	G_UserList(28, i-1)
			Case "homepage"
				If Not IsArray(G_UserList(27, i-1)) Then	G_UserList(27, i-1)=Split(G_UserList(27, i-1), "|||")
				TPL_Echo	G_UserList(27, i-1)(0)
			Case "uc"
				If Not IsArray(G_UserList(27, i-1)) Then	G_UserList(27, i-1)=Split(G_UserList(27, i-1), "|||")
				TPL_Echo	G_UserList(27, i-1)(6)
		End Select
	Else	'�ο� �� ������
		Select Case sToken
			Case "username"
				s	= Split(G_ItemList(6, G_Floor), ".")
				If i>0 Then
					s2	= s2 &	"����"
				Else
					s2	= s2 &	"����"
				End If
				s2	= s2 &	("(" & s(0) & "." & s(1) & ".*.*)")
				TPL_Echo	s2
			Case "richname"
				s	= Split(G_ItemList(6, G_Floor), ".")
				If i>0 Then
					s2	= s2 &	"����"
				Else
					s2	= s2 &	"����"
				End If
				s2	= s2 &	("(" & s(0) & "." & s(1) & ".*.*)")
				s=Split(Split(Application(Dvbbs.CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='7']/@groupsetting").text,",")(58),"��")
				TPL_Echo "<span style=""width:130px;filter:glow(color='gray',strength='2');"">"&(s(0)&s2&s(1)&"</span>")
			Case "userface"
				If i>0 Then
					TPL_Echo	"<img src=""Skins/Default/anyon.gif"" width=""111"" height=""111"" border=""0"" />"
				Else
					TPL_Echo	"<img src=""Skins/Default/guest.gif"" width=""111"" height=""111"" border=""0"" />"
				End If
			Case Else
		End Select
	End If
End Sub

Function IsSelfPost()
	IsSelfPost=False
	If G_ItemList(10, G_Floor)>0 Then
		If G_UserList(0, G_ItemList(10, G_Floor)-1)=Dvbbs.UserID Then
			IsSelfPost=True
		End If
	End If
End Function

Function GetPostUserID()
	If G_ItemList(10, G_Floor)>0 Then
		GetPostUserID=G_UserList(0, G_ItemList(10, G_Floor)-1)
	Else
		GetPostUserID=0
	End If
End Function

Sub ParseBBSListNode(sToken)
	Dim i, a, postbuyusers, postbuyinfo, j
	Select Case sToken
		Case "announceid"
			TPL_Echo G_ItemList(0, G_Floor)
		Case "title"
			TPL_Echo Dvbbs.Replacehtml(Dvbbs.ChkBadWords(G_ItemList(2, G_Floor)))
		Case "body"
			i	= G_ItemList(10, G_Floor)
			If i>0 Then
				i	= G_UserList(26, i-1)
			Else
				i	= 7	'����
			End If
			If 0=G_Floor And 1=G_CurrentPage And CLng(G_ItemList(20, G_Floor))=0 Then	'��¥��Ҫ�жϹ�����
				If G_LockTopic Then TPL_Echo "<div class=""limitinfo"">�����ѱ�����</div><br/>"
				If 3=T_GetMoneyType Then
					TPL_Echo "<div class=""limitinfo""><font color=""gray"">����������Ҫ֧�� <b><font color=""red"">"&G_ItemList(17, G_Floor)&"</font></b> ����ҷ��ɲ鿴��"
					If IsSelfPost() Then
						TPL_Echo "��������������"
					ElseIf TrueMaster Then
						TPL_Echo "�������ǹ�����Ա������Կ������ݡ�"
					Else
						If Trim(PostBuyUser)="" Then PostBuyUser="0@@@-1@@@0@@@|||$PayMoney|||"
						postbuyusers=split(PostBuyUser,"|||")
						postbuyinfo=postbuyusers(0)
						postbuyinfo=Split(postbuyinfo,"@@@") 'Rem postbuyinfo(0) ����money   postbuyinfo(1) ��������maxbuy   postbuyinfo(2) vip�Ƿ���Ҫ����notvipbuy   postbuyinfo(3) ���������û��б�buyuser
						If UBound(postbuyinfo)<=2 Then Exit Sub
						a=False
						For j=2 to UBound(postbuyusers)
							If postbuyusers(j)<>"" And postbuyusers(j)=Dvbbs.MemberName Then
								a=True
							End If
						Next
						If a Then
							TPL_Echo "���Ѿ�����"
						ElseIf Dvbbs.VipGroupUser And "1"=postbuyinfo(2) Then
							TPL_Echo "��������vip�û���������Ϊ������vip�û����⹺��鿴��������ֱ�Ӳ鿴��"
						ElseIf (""=postbuyinfo(3) Or InStr(","&postbuyinfo(3)&",", ","&Dvbbs.MemberName&",")>0) And Dvbbs.userid>0 Then
							TPL_Echo "����Ҫ���򷽿ɿ������ݡ�</font><br /><input type=""button"" value=""��Ҫ�鿴���ݣ���������"" onclick=""location.href='BuyPost.asp?action=buy&amp;boardid="&Dvbbs.BoardID&"&amp;id="&AnnounceID&"&amp;ReplyID="&G_ItemList(0, G_Floor)&"&amp;PostTable="&TotalUsetable&"'""/>"
							TPL_Echo "<br/></div>"
							Exit Sub
						Else
							TPL_Echo "����Ҫ���򷽿ɿ������ݡ�"
							If Dvbbs.userid>0 Then
								TPL_Echo "¥���������������Թ���"
							Else
								TPL_Echo "����δ��¼�����ܹ���"
							End If
							TPL_Echo "</font><br/></div>"
							Exit Sub
						End If
					End If
					TPL_Echo "</font><br/></div>"
				End If
			End If
			Ubblists=G_ItemList(15, G_Floor)
			If InStr(Ubblists,",39,") > 0  Then
				TPL_Echo	Dvbbs.ChkBadWords( Dv_ubb.Dv_UbbCode(G_ItemList(4, G_Floor),i,1,0) )
			Else
				TPL_Echo	Dvbbs.ChkBadWords( Dv_ubb.Dv_UbbCode(G_ItemList(4, G_Floor),i,1,1) )
			End If
		Case "bodystyle"
			TPL_Echo ("font-size:"&Dvbbs.Board_setting(28)&"pt;text-indent:"&Dvbbs.Board_setting(69)&"px;")
		Case "floor"
			i	= G_Moved+G_Floor
			TPL_Echo i+1
		Case "dateandtime" TPL_Echo G_ItemList(3, G_Floor)
		Case "showpage"
			TPL_ShowPage	G_CurrentPage, G_Childs, Dvbbs.CheckNumeric(Dvbbs.Board_Setting(27)), 10, "dispbbs.asp?boardid="&Dvbbs.BoardID&"&id="&AnnounceID&"&page="&Dvbbs.CheckNumeric(request("page"))&"&star="
		Case "bestinfo"
			If 1=G_ItemList(9, G_Floor) Then
				TPL_Echo "<div class=""info""><img src="""&Dvbbs.Forum_PicUrl&"jing.gif"" border=""0"" title=""��������Ϊ����"" align=""absmiddle""/>[��������Ϊ����]</div>"
			End If
		Case "bestpic"
			If 1=G_ItemList(9, G_Floor) Then
				TPL_Echo "<img src=""images/best.gif"" border=""0"" style=""position:absolute;z-index:1;"" title=""����������֤"" align=""absmiddle"" />"
			End If
		Case "appraise"
			If IsNull(G_ItemList(12, G_Floor)) Then Exit Sub
			SplitIsAgree
			a = G_ItemList(12, G_Floor)
			If a(1)>0 Then
				TPL_Echo	"<div class=""info"">����������<img src="""&Dvbbs.Forum_PicUrl&"agree.gif"" border=""0"" alt=""���������"&a(1)&"����ҽ���"" align=""absmiddle""/>���������<font color=""red"">"&a(1)&"</font>����ҽ���</div>"
				If ""<>a(3) Then TPL_Echo "(" & a(3) & ")"
			ElseIf a(0)>0 Then
				TPL_Echo	"<div class=""info"">����������<img src="""&Dvbbs.Forum_PicUrl&"disagree.gif"" border=""0"" alt=""�������۳�"&a(0)&"�����"" align=""absmiddle""/>�������۳�<font color=""red"">"&a(0)&"</font>�����</div>"
				If ""<>a(2) Then TPL_Echo "(" & a(2) & ")"
			End If
		Case "agree"
			If IsNull(G_ItemList(12, G_Floor)) Then
				TPL_Echo "0"
			Else
				SplitIsAgree
				TPL_Echo G_ItemList(12, G_Floor)(5)
			End If
		Case "against"
			If IsNull(G_ItemList(12, G_Floor)) Then
				TPL_Echo "0"
			Else
				SplitIsAgree
				TPL_Echo G_ItemList(12, G_Floor)(6)
			End If
		Case "neutral"
			If IsNull(G_ItemList(12, G_Floor)) Then
				TPL_Echo "0"
			Else
				SplitIsAgree
				TPL_Echo G_ItemList(12, G_Floor)(4)
			End If
		Case "moneytype"
			i = T_GetMoneyType
			If 0=i Then Exit Sub
			If 0=G_Floor And G_CurrentPage=1 Then
				Select Case i
					Case 1 TPL_Echo "<div class=""info"">���ͽ������Ҫ���� <font color=""red"">" & G_ItemList(17, G_Floor) & "</font> �����</div>"
					Case 5 TPL_Echo "<div class=""info"">���ͽ������Ҫ���� <font color=""red"">" & G_ItemList(17, G_Floor) & "</font> �����[�ѽ���]</div>"
					Case 2 TPL_Echo "<div class=""info"">���������������� <font color=""red"">" & G_ItemList(17, G_Floor) & "</font> �����</div>"
					'��������ڶ�body�ļ�����Ѿ����ˡ�
					'Case 3 TPL_Echo "�鿴������֧�� <font color=""red"">" & G_ItemList(17, G_Floor) & "</font> �����"
					Case Else
				End Select
			Else
				If (1=i Or 5=i) And G_ItemList(10, G_Floor)>0 Then
					If G_UserList(0, G_ItemList(10, G_Floor)-1)<>PostUserid Then
						TPL_Echo "<div class=""info"">���<font color=""red"">" & G_ItemList(17, G_Floor) & "</font>�����</div>"
					End If
				End If
				If 1=i And PostUserid=Dvbbs.UserID And G_ItemList(10, G_Floor)>0 Then
					If G_UserList(0, G_ItemList(10, G_Floor)-1)<>Dvbbs.UserID Then
						TPL_Echo "<div class=""info""><a href=""BuyPost.asp?Action=Send&PostTable="&TotalUsetable&"&BoardId="&Dvbbs.BoardID&"&ID="&AnnounceID&"&ReplyID="&ReplyID_a&"&UserName="&UserName&""" title=""���ͽ�Ҹ���Ա"&UserName&""" target=""_blank"">���ͽ��</a></div>"
					End If
				End If
				If 2=i Then TPL_Echo "<span class=""info"">����¥��:<font color=""red"">" & G_ItemList(17, G_Floor) & "</font>�����</span>"
			End If
		Case "usetools"
			If ""<>G_ItemList(18, G_Floor) Or (0<T_GetMoneyType And (G_Moved+G_Floor)=0) Then
				TPL_Echo "<hr /><a href=""javascript:;"" onclick=""openScript('ViewInfo.asp?t=2&amp;action=View&amp;PostTable="&TotalUsetable&"&amp;BoardId="&Dvbbs.BoardID&"&amp;ID="&AnnounceID&"&amp;ReplyID="&ReplyID_a&"',600,450)""><img src="""&Dvbbs.Forum_PicUrl&"mini_query.gif"" border=""0"" alt=""�鿴ʹ�õ�����ϸ��Ϣ""  class=""imgonclick"" /></a>"
			End If
		Case "topicface"
			TPL_Echo G_ItemList(5, G_Floor)
		Case "ip"
			If "1"=Dvbbs.GroupSetting(30) And (TrueMaster Or 3<>Dvbbs.UserGroupID) Then TPL_Echo "&nbsp;&nbsp;&nbsp;Post IP��<a href=""TopicOther.asp?t=1&amp;boardid="&Dvbbs.Boardid&"&amp;userid="&GetPostUserID()&"&amp;ip="&G_ItemList(6, G_Floor)&"&amp;action=lookip"" title=""����鿴�û���Դ������"">"&G_ItemList(6, G_Floor)&"</a>"
		Case "magicface"
			If 0=G_Floor And G_CurrentPage=1 And Skin=0 Then
				i=InStr(G_Expression,"|")
				If i>0 Then
					a=Left(G_Expression, i-1)
					If "0"<>a Then
						TPL_Echo "<div style=""float:right;margin-right:20px;""><a href=""javascript:DispMagicEmot("&a&",350,500)""><img src=""dv_plus/tools/magicface/gif/"&a&".gif"" border=""0"" alt=""""/><br /><br/></a></div><script type=""text/javascript"" language=""javascript"">LoadMagicEmot("&a&","&AnnounceID&");</script>"
					End If
				End If
			End If
		Case "adminlist"
			i=GetPostUserID()
			TPL_Echo "<span class=""m_li_top"" style=""display:inline;"" onmouseover=""showmenu1('Menu_"&ReplyID_a&"',0);""><a href=""#"">��������</a>"
			TPL_Echo "	<div class=""submenu submunu_popup"" id=""Menu_"&ReplyID_a&""" onmouseout=""hidemenu1()"">"
			TPL_Echo "	<a href=""TopicOther.asp?t=6&amp;BoardID="&Dvbbs.boardid&"&amp;id="&AnnounceID&"&amp;ReplyID="&ReplyID_a&""">�ٱ�����</a><br />"
			If Dvbbs.Boardmaster Then
				If G_Floor>0 Or G_CurrentPage>1 Then
					TPL_Echo "	<a href=""admin_postings.asp?action=dele_a&amp;BoardID="&Dvbbs.boardid&"&amp;replyID="&ReplyID_a&"&amp;ID="&AnnounceID&"&amp;star=1&amp;userid="&i&""">ɾ������</a><br />"
				End If
				TPL_Echo "	<a href=""admin_postings.asp?action=copy_a&amp;BoardID="&Dvbbs.boardid&"&amp;replyID="&ReplyID_a&"&amp;ID="&AnnounceID&"&amp;star=1&amp;userid="&i&""">��������</a><br />"
				If 0=G_ItemList(9,G_Floor) Then
					TPL_Echo "<a href=""admin_postings.asp?action=isbest_a&amp;BoardID="&Dvbbs.boardid&"&amp;replyID="&ReplyID_a&"&amp;ID="&AnnounceID&"&amp;star=1&amp;userid="&i&""">��Ϊ����</a><br />"
				Else
					TPL_Echo "<a href=""admin_postings.asp?action=nobest_a&amp;BoardID="&Dvbbs.boardid&"&amp;replyID="&ReplyID_a&"&amp;ID="&AnnounceID&"&amp;star="&G_CurrentPage&"&amp;userid="&GetPostUserID()&""">�������</a><br/>"
				End If
				Select Case G_ItemList(16,G_Floor)
					Case 2
						TPL_Echo "<a href=""admin_postings.asp?action=nolockpage_a&amp;BoardID="&Dvbbs.boardid&"&amp;replyID="&ReplyID_a&"&amp;ID="&AnnounceID&"&amp;star="&G_CurrentPage&"&amp;userid="&GetPostUserID()&""">�������</a><br />"
					Case 3
						TPL_Echo "<a href=""AccessTopic.asp?action=unlock&amp;BoardID="&Dvbbs.boardid&"&amp;replyID="&ReplyID_a&"&amp;ID="&AnnounceID&""">���ͨ��</a><br />"
					Case Else
						TPL_Echo "<a href=""admin_postings.asp?action=islockpage_a&amp;BoardID="&Dvbbs.boardid&"&amp;replyID="&ReplyID_a&"&amp;ID="&AnnounceID&"&amp;star="&G_CurrentPage&"&amp;userid="&GetPostUserID()&""">��������</a><br />"
				End Select
				TPL_Echo "	<a href=""admin_postings.asp?action=RewardMoney&amp;BoardID="&Dvbbs.boardid&"&amp;replyID="&ReplyID_a&"&amp;ID="&AnnounceID&"&amp;star=1"" title=""������������ɽ�����۳������û���ط�ֵ"">��������</a><br />"
			Else
				If ((G_Floor>0 Or G_CurrentPage>1) And IsSelfPost() And "1"=Dvbbs.GroupSetting(11)) Then
					TPL_Echo "	<a href=""admin_postings.asp?action=dele_a&amp;BoardID="&Dvbbs.boardid&"&amp;replyID="&ReplyID_a&"&amp;ID="&AnnounceID&"&amp;star=1&amp;userid="&i&""">ɾ������</a><br />"
				End If
			End If
			If "1"=Dvbbs.Forum_Setting(90) Then TPL_Echo "	<a title=""�Ա���ʹ����̳����"" href=""javascript:openScript('plus_Tools_InfoSetting.asp?action=0&amp;BoardID="&Dvbbs.boardid&"&amp;TopicID="&AnnounceID&"&amp;ReplyID="&ReplyID_a&"',500,400)"">ʹ�õ���</a><br />"
			TPL_Echo "	</div></span>"
		Case Else
	End Select
End Sub

Sub SplitIsAgree()
	Dim a : a = G_ItemList(12, G_Floor)
	If Not IsArray(a) Then
		a = Split(G_ItemList(12, G_Floor), "|")
		If UBound(a)<6 Then
			If UBound(a)<1 Then
				a = Split("0|0|||0|0|0","|")
			Else
				a = Split(G_ItemList(12, G_Floor)&"|||0|0|0","|")
			End If
		End If
		G_ItemList(12, G_Floor) = a
	End If
End Sub

Function Topic_Ads(n)
		Randomize
		Topic_Ads=Dvbbs.Forum_ads(n)(CInt(UBound(Dvbbs.Forum_ads(n))*Rnd))
End Function

Sub ParseADNode(sToken)
	If 0=G_Floor And UBound(Dvbbs.Forum_ads)>23 Then
		Select Case sToken
			Case "first_body_top"
				If "1"=Dvbbs.Forum_ads(18) Then
					Dvbbs.Forum_ads(19)=Split(Dvbbs.Forum_ads(19), "#####")
					TPL_Echo	"<div class=""first_body_top"">"&Topic_Ads(19)&"</div>"
				End If
			Case "first_body_bottom"
				If "1"=Dvbbs.Forum_ads(20) Then
					Dvbbs.Forum_ads(21)=Split(Dvbbs.Forum_ads(21), "#####")
					TPL_Echo	"<div class=""first_body_bottom"">"&Topic_Ads(21)&"</div>"
				End If
			Case "first_body_left"
				If "1"=Dvbbs.Forum_ads(22) Then
					Dvbbs.Forum_ads(23)=Split(Dvbbs.Forum_ads(23), "#####")
					TPL_Echo	"<div class=""first_body_left"">"&Topic_Ads(23)&"</div>"
				End If
			Case "first_body_right"
				If "2"=Dvbbs.Forum_ads(22) Then
					Dvbbs.Forum_ads(23)=Split(Dvbbs.Forum_ads(23), "#####")
					TPL_Echo	"<div class=""first_body_right"">"&Topic_Ads(23)&"</div>"
				End If
			Case Else
		End Select
	End If
	Select Case sToken
		Case "bbslist_bottom"
			If "1"=Dvbbs.Forum_ads(7) Then	TPL_Echo	Topic_Ads(14)
		Case Else
	End Select
End Sub

Sub TPL_ParseNode(sTokenType, sTokenName)
	Select Case sTokenType
		Case "page"
			ParsePageNode		sTokenName
		Case "user"
			ParseUserNode	sTokenName
		Case "bbslist"
			ParseBBSListNode	sTokenName
		Case "ad"
			ParseADNode			sTokenName
		Case "qcomic_plus"
			DIM Qcomic_setting, codeStr
			Qcomic_setting = Split(Dvbbs.qcomic_plus_setting(), "||||")
			codeStr = "phid="&Trim(G_ItemList(21, G_Floor))&"&spassword="&Qcomic_setting(2)&"&ctime="&Now()
			Select Case sTokenName
				Case "qcomic_enable"	:	TPL_Echo "true"
				Case "qcomic_sid"		:	TPL_Echo Qcomic_setting(1)
				Case "qcomic_sid_phid"	:	TPL_Echo Qcomic_setting(1)&"_"&Trim(G_ItemList(21, G_Floor))
				Case "qcomic_auto"		:	TPL_Echo "0"
				Case "qcomic_code"		:	TPL_Echo Server.UrlEncode(AuthCode(codeStr, "ENCODE",Qcomic_setting(3)))
				Case "qcomic_owidth"	:	TPL_Echo Qcomic_setting(4)
				Case "qcomic_oheight"	:	TPL_Echo Qcomic_setting(5)
			End Select
	End Select
End Sub

Sub TPL_ParseArea(sTokenName, sTemplate)
	Dim iUbd,sTemp
	Select Case sTokenName
		Case "bbslist"
			iUbd = UBound(G_ItemList, 2)
			If "1"=Dvbbs.Forum_ads(7) Then	Dvbbs.Forum_ads(14)=Split(Dvbbs.Forum_ads(14),"#####")
			For G_Floor=0 To iUbd
				'���漸����ֵ����dv_ubbcode.asp
				RootID_a	= G_ItemList(7, G_Floor)
				AnnounceID_a= RootID_a
				ReplyID_a	= G_ItemList(0, G_Floor)
				UserName	= G_ItemList(1, G_Floor)
				PostBuyUser = G_ItemList(19, G_Floor)
				TPL_Scan	sTemplate
			Next
		Case "userinfo"
			If G_ItemList(10, G_Floor)>0 Then
				If 2<>G_ItemList(8, G_Floor) Or Dvbbs.BoardMaster Or _
				Dvbbs.UserID=G_UserList(0, G_ItemList(10, G_Floor)-1) Then TPL_Scan	sTemplate
			End If
		Case "boke"
			If 1=Dvbbs.Forum_setting(99) Then	TPL_Scan	sTemplate
		Case "bbslimit"
			sTemp=""
			If 1=G_ItemList(9, G_Floor) Then
				If "1"<>Dvbbs.GroupSetting(41) Then sTemp="<div class=""limitinfo"">����Ȩ�鿴��������</div>"
			Else
				Select Case G_ItemList(16, G_Floor)
					Case 2 sTemp="<div class=""limitinfo"">���ݱ�����</div>"
					Case 3 sTemp="<div class=""limitinfo"">���ݴ����</div>"
					Case Else
						If G_ItemList(10, G_Floor)>0 Then
							Select Case G_UserList(13, G_ItemList(10, G_Floor)-1)
								Case 1 sTemp="<div class=""limitinfo"">�û��ѱ�����</div>"
								Case 2 sTemp="<div class=""limitinfo"">�û��Ѿ�������</div>"
								Case Else
							End Select
						End If
				End Select
			End If
			If ""<>sTemp Then
				TPL_Echo sTemp
			End If
			If ""=sTemp Or TrueMaster Or (Dvbbs.Boardmaster And 3<>Dvbbs.UserGroupID) Then
				TPL_Scan	sTemplate
			End If
		Case "qcomic_plus"
			If Dvbbs.qcomic_plus Then
				If Trim(G_ItemList(21, G_Floor))<>"" And Trim(G_ItemList(21, G_Floor))<>"0" Then
					TPL_Scan	sTemplate
				End If
			End If
		Case "logined"
			If Dvbbs.userid>0 Then TPL_Scan	sTemplate
		Case "showvote"
			If G_IsVote>0 And 1=G_CurrentPage Then LoadAndParseVote sTemplate
		Case "canreply"
			If G_CanReply Then TPL_Scan	sTemplate
		Case "canedit"
			If (IsSelfPost() And Dvbbs.GroupSetting(10)="1") Or Dvbbs.boardmaster Then TPL_Scan	sTemplate
		Case "tenpay"
			If Dvbbs.Board_Setting(67)=1 Then	TPL_Scan sTemplate
	End Select
End Sub
Dvbbs.Footer
Dvbbs.PageEnd
end if
%>