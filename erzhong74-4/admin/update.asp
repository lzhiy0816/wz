<!--#include file="../conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="../inc/dv_clsother.asp"-->
<!--#include file="../inc/ubblist.asp"-->
<%
Head()
Server.ScriptTimeout=9999999
dim admin_flag 
admin_flag=Split("14,20",",")
CheckAdmin(","&admin_flag(0)&",")
CheckAdmin(","&admin_flag(1)&",")

Dim tmprs,body
Call main()
Footer()

sub main()
Dim i
%>
<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
<tr>
<th style="text-align:center;" colspan=2>��̳���ݴ���</th>
</tr>
<tr>
<td width="20%" class="td1" height=25>ע������</td>
<td width="80%" class="td1">�����еĲ������ܽ��ǳ����ķ�������Դ�����Ҹ���ʱ��ܳ�������ϸȷ��ÿһ��������ִ�С�</td>
</tr>
<%
	If request("action")="updat" Then
		If request("submit")="���·ְ�������" Or request("submit")="������̳����" Then
			call updateboard()
		ElseIf request("submit")="�� ��" Then
			call fixtopic()
		ElseIf request("submit")="��������û�" Then
			call Delallonline()
		ElseIf request("submit")="�����ղؼ�" Then
			call Updatebm()
		Else
			call updateall()
		End If
		If founderr Then
			response.write errmsg
		Else
			response.write body
		End If
	ElseIf request("action")="fix"  Then
		Call Fixbbs()
		If founderr Then
			response.write errmsg
		Else
			response.write body
		End If
	ElseIf request("action")="delboard" then
		if isnumeric(request("boardid")) then
		Dvbbs.Execute("update dv_topic set boardid=444 where boardid="&request("boardid"))
		for i=0 to ubound(AllPostTable)
		Dvbbs.Execute("update "&AllPostTable(i)&" set boardid=444 where boardid="&request("boardid"))
		next
		end if
		response.write "<tr><td align=left colspan=2 height=23 class=td1>�����̳���ݳɹ����뷵�ظ����������ݣ�</td></tr>"
	elseif request("action")="updateuser" then
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th style="text-align:center;" colspan=2>�����û�����</th>
</tr>
<tr>
<td width="20%" class="td1">���¼����û�����</td>
<td width="80%" class="td1">ִ�б�����������<font color=red>��ǰ��̳���ݿ�</font>�������¼��������û���������������</td>
</tr>
<tr>
<td width="20%" class="td1">��ʼ�û�ID</td>
<td width="80%" class="td1"><input type=text name="beginID" value="1" size=10>&nbsp;�û�ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="td1">�����û�ID</td>
<td width="80%" class="td1"><input type=text name="endID" value="100" size=10>&nbsp;�����¿�ʼ������ID֮����û����ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="td1">&nbsp;</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="���¼����û�����"></td>
</tr>
</form>

<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr>
<td width="20%" class="td1" valign=top>�����û��ȼ�</td>
<td width="80%" class="td1">ִ�б�����������<font color=red>��ǰ��̳���ݿ�</font>�û�������������̳�ĵȼ��������¼����û��ȼ�����������Ӱ��ȼ�Ϊ������������ܰ��������ݡ�</td>
</tr>
<tr>
<td width="20%" class="td1">��ʼ�û�ID</td>
<td width="80%" class="td1"><input type=text name="beginID" value="1" size=10>&nbsp;�û�ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="td1">�����û�ID</td>
<td width="80%" class="td1"><input type=text name="endID" value="100" size=10>&nbsp;�����¿�ʼ������ID֮����û����ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="td1">&nbsp;</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="�����û��ȼ�"></td>
</tr>
</form>

<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr>
<td width="20%" class="td1" valign=top>�����û���Ǯ/����/����</td>
<td width="80%" class="td1">ִ�б�����������<font color=red>��ǰ��̳���ݿ�</font>�û��ķ�����������̳������������¼����û��Ľ�Ǯ/����/������������Ҳ�����¼��������������ܰ���������<BR>ע�⣺���Ƽ��û����б������������������ݺܶ��ʱ���뾡����Ҫʹ�ã����ұ������Ը�������ɾ�����ӵ�������Ӧ��ֵ�������㣬ֻ�ǰ��շ������ܵ���̳��ֵ���ý������㣬�������ز�����<font color=red>���ұ�������������û���Ϊ�������ͷ���ԭ�����Ա���û���ֵ���޸ġ�</font></td>
</tr>
<tr>
<td width="20%" class="td1">��ʼ�û�ID</td>
<td width="80%" class="td1"><input type=text name="beginID" value="1" size=10>&nbsp;�û�ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="td1">�����û�ID</td>
<td width="80%" class="td1"><input type=text name="endID" value="100" size=10>&nbsp;�����¿�ʼ������ID֮����û����ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="td1">&nbsp;</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="�����û���Ǯ/����/����"></td>
</tr>
</FORM>
<%
	elseif request("action")="updateuserinfo" then
		if request("submit")="���¼����û�����" then
		call updateTopic()
		elseif request("submit")="�����û��ȼ�" then
		call updategrade()
		else
		call updatemoney()
		end if
		if founderr then
		response.write errmsg
		else
		response.write body
		end if
	else
	'������,������,�û���,������,������,�̶ܹ�,���ע��
%>
<tr> 
<th style="text-align:center;" colspan=2>������̳����</th>
</tr>

<form action="update.asp?action=updat" method=post>
<tr>
<td width="20%" class="td2">��������̳����</td>
<td width="80%" class="td2">
<input type="checkbox" class="checkbox" name="u1" value="1">
������
<input type="checkbox" class="checkbox" name="u2" value="1">
������
<input type="checkbox" class="checkbox" name="u3" value="1">
�û���
<input type="checkbox" class="checkbox" name="u4" value="1" checked>
������
<input type="checkbox" class="checkbox" name="u5" value="1" checked>
������
<input type="checkbox" class="checkbox" name="u6" value="1">
�̶ܹ�
<input type="checkbox" class="checkbox" name="u7" value="1">
���ע��
<BR><BR><input type="submit" class="button" name="Submit" value="������̳������"><BR><BR>���ｫ���¼���������̳����������ͻظ������������ӣ��������û��ȣ�����ÿ��һ��ʱ������һ�Ρ�<hr size=1></td>
</tr>
<tr>
<td width="20%" class="td1">���·ְ�������</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="���·ְ�������"><BR><BR>���ｫ���¼���ÿ���������������ͻظ������������ӣ����ظ���Ϣ�ȣ�����ÿ��һ��ʱ������һ�Ρ�<hr size=1>
</td>
</tr>
<tr>
<td width="20%" class="td2">������̳�ղؼ�</td>
<td width="80%" class="td2"><input type="submit" class="button" name="Submit" value="�����ղؼ�"><BR><BR>���ｫ����������̳���ղؼУ�ɾ���������û����ղؼ�¼������ָ���ƶ��������ղص�ַ��ɾ���ѱ�ɾ���������ղؼ�¼��
</td>
</tr>
<tr> 
<th style="text-align:center;" colspan=2>�޸�����(�޸�ָ����Χ�����ӵ����ظ�����)</th>
</tr>
<tr>
<td width="20%" class="td1">��ʼ��ID��</td>
<td width="80%" class="td1"><input type=text name="beginID" value="1" size=10>&nbsp;��������ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="td2">������ID��</td>
<td width="80%" class="td2"><input type=text name="EndID" value="1000" size=10>&nbsp;�����¿�ʼ������ID֮����������ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="td1">&nbsp;</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="�� ��"></td>
</tr>
</form>
<form name=Fix action="update.asp?action=fix" method=post>
<tr> 
<th style="text-align:center;" colspan=2>��������UBB��ǩ(�޸�ָ����Χ����UBB��ǩ)</th>
</tr>
<tr>
<td width="20%" class="td1">��ʼ��ID��</td>
<td width="80%" class="td1"><input type=text name="beginID" value="1" size=10>&nbsp;��������ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="td2">������ID��</td>
<td width="80%" class="td2"><input type=text name="EndID" value="1000" size=10>&nbsp;�����¿�ʼ������ID֮����������ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="td1">�������ı�ʶ����</td>
<td width="80%" class="td1"><input type="text" name="updatedate" value="2003-12-1">(��ʽ��YYYY-M-D) ������̳������v7.0�����ڣ��������д��һ�ɰ���������</td>
</tr>
<tr>
<td width="20%" class="td2">ȥ�������е�HTML���</td>
<td width="80%" class="td2">�� <input type="radio" class="radio" name="killhtml" value="1">
  �� <input type="radio" class="radio" name="killhtml" value="0" checked>&nbsp;<br>ѡ�ǵĻ��������е�HTML��ǽ����Զ�������������ڼ������ݿ�Ĵ�С�����ǻ�ʧȥԭ����HTMLЧ����</td>
</tr>
<tr>
<td width="20%" class="td1">&nbsp;</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="�� ��"></td>
</tr>
</form>

<%
	end if
%>
</table><BR><BR>
<%
	end sub

Sub updateboard()
	'�Ȱ������а���ID�ó���������Ȼ����������������̳�������ܺ�
	Dim allarticle
	Dim alltopic
	Dim alltoday
	Dim allboard
	Dim trs,Esql,ars,rs,i
	Dim Maxid
	Dim LastTopic,LastRootid,LastPostTime,LastPostUser
	Dim LastPost,uploadpic_n,Lastpostuserid,Lastid
	Dim ParentStr
	Dim C,C1,C2
	Dim reBoard_Setting,BoardTopStr,IsGroupSetting
	Dim UserAccessCount,UpGroupSetting,ii
	Dim Slastpost
	ii=0
	'���ô�����ʱ��
	conn.CommandTimeout=3600

	'���Ҫ���µ�����
	If IsNumeric(request("boardid")) And request("boardid")<>"" Then
		Set Rs=Dvbbs.Execute("Select Count(*) From [Dv_board] Where BoardID="&request("boardid"))
		C1=rs(0)
		If Isnull(C1) Then C1=0
	Else
		Set Rs=Dvbbs.Execute("Select Count(*) From [Dv_board]")
		C1=rs(0)
		If Isnull(C1) Then C1=0
	End If
	Set Rs=Nothing
%>
</table><BR>
<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
<tr><td colspan=2 class=td1>
���濪ʼ������̳�������ϣ�����<%=C1%>��������Ҫ����
<table width="400" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=#000000>
<table width="400" border="0" cellspacing="0" cellpadding="1">
<tr> 
<td bgcolor=#ffffff height=9><img src="../skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
</td></tr></table> <span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td></tr>
</table>
<%
	Response.Flush

	'������Child��Orders���Ա��ȸ����¼���̳�����ݲ�ѭ�����ϼ����棬��ʱ�ϼ������ȡ�ľ����¼��������������
	If IsNumeric(request("boardid")) And request("boardid")<>"" Then
		Set Rs=Dvbbs.Execute("Select BoardID,BoardType,Child,ParentStr,RootID,Board_Setting,BoardTopStr,IsGroupSetting From Dv_Board Where BoardID="&Request("BoardID"))
	Else
		Call Boardchild()	'ͳ�Ƹ���������̳���� YZ-2004-2-26ע
		Set Rs=Dvbbs.Execute("Select BoardID,BoardType,Child,ParentStr,RootID,Board_Setting,BoardTopStr,IsGroupSetting From Dv_Board Order by Child,RootID,Orders Desc")
	End If
	Dim SQL, LastPostArr 'LastPostArr XG����2007-04-11
	If Not Rs.EOF Then 
		SQL=Rs.GetRows(-1)
		Set Rs=Nothing
		For i=0 to UBound(SQL,2)
	'Do While Not Rs.Eof
		
		reBoard_Setting=Split(SQL(5,i)&"",",")
		AllBoard = 0
		'�������������
		Set Trs=Dvbbs.Execute("Select Count(*),Sum(Child) From Dv_Topic Where BoardID="&SQL(0,i))
		AllTopic=Trs(0)
		AllArticle=Trs(1)
		If IsNull(AllTopic) Then AllTopic = 0
		If IsNull(AllArticle) Then AllArticle = 0
		AllArticle = AllArticle + AllTopic
		Set Trs=Nothing
		'���н�����
		If IsSqlDataBase = 1 Then
			Set Trs=Dvbbs.Execute("Select Count(*) From "&Dvbbs.NowUseBBS&" Where BoardID="&SQL(0,i)&" and datediff(d,dateandtime,"&SqlNowString&")=0")
		Else
			Set Trs=Dvbbs.Execute("Select Count(*) From "&Dvbbs.NowUseBBS&" Where BoardID="&SQL(0,i)&" and datediff('d',dateandtime,"&SqlNowString&")=0")
		End If
		AllToday=Trs(0)
		Set Trs=Nothing
		If IsNull(AllToday) Then AllToday=0
		'���ظ���Ϣ
		Set Trs=Dvbbs.Execute("Select Top 1 LastPost,TopicID,Title,PostTable From Dv_Topic Where BoardID="&SQL(0,i)&" Order by LastPostTime Desc")
		If Not (Trs.Eof And Trs.Bof) Then
			LastPostArr = Split(Trs(0)&"","$")
			If UBound(LastPostArr)<7 Then
				ReDim LastPostArr(7)
				LastPostArr(1)=Trs(3)
				LastPostArr(3)= Replace(cutStr(Dvbbs.Replacehtml(Trs(2)),20),"$","&#36;")
				LastPostArr(6)=Trs(1)
				Trs.Close:Set Trs=Nothing
				Set Trs = Dvbbs.Execute("Select Top 1 AnnounceID,BoardID,UserName,DateAndTime,PostUserID From "&LastPostArr(1)&" Where RootID="&LastPostArr(6)&" Order by DateAndTime")
				If Not (Trs.Eof And Trs.Bof) Then
					LastPostArr(0) = Trs(2)
					LastPostArr(1) = Trs(0)
					LastPostArr(2) = Trs(3)
					LastPostArr(4) = ""
					LastPostArr(5) = Trs(4)
					LastPostArr(7) = Trs(1)
				Else
					LastPostArr(0) = "��"
					LastPostArr(1) = 0
					LastPostArr(2) = now()
					LastPostArr(3) = "��"
					LastPostArr(4) = ""
					LastPostArr(5) = ""
					LastPostArr(6) = ""
					LastPostArr(7) = ""
				End if
			Else
				LastPostArr(3)= Replace(cutStr(Trs(2),20),"$","&#36;")
			End if
			Trs.Close:Set Trs=Nothing
			LastPost=Join(LastPostArr,"$")
			LastPost=Replace(Dvbbs.Replacehtml(LastPost&""),"'","''")
			'LastPost=Replace(Dvbbs.Replacehtml(Trs(0)&""),"'","''")
		Else
			Trs.Close:Set Trs=Nothing
			LastPost="��$0$"&Now()&"$��$$$$"
		End If
		'���µ�ǰ��������
		'SLastPost = Split(LastPost,"$")
		'If Ubound(SLastPost) < 7 Then LastPost = LastPost & "$"
		Dvbbs.Execute("Update [Dv_board] Set PostNum="&AllArticle&",TopicNum="&AllTopic&",TodayNum="&AllToday&",LastPost='"&Dvbbs.ChkBadWords(LastPost)&"' Where BoardID="&SQL(0,i))
		'�����ǰ������������̳�������������Ϊ������̳����
		If SQL(2,i)>0 Then
			'������������������������������������������
			If SQL(3,i)=0 Then
				ParentStr=SQL(0,i)
				'Set Trs=Dvbbs.Execute("Select Sum(PostNum),Sum(TopicNum),Sum(TodayNum),Count(*) From Dv_board Where (Not BoardID="&SQL(0,i)&") And RootID="&SQL(0,i))
			Else
				ParentStr=SQL(3,i) & "," & SQL(0,i)
				'Set Trs=Dvbbs.Execute("Select Sum(PostNum),Sum(TopicNum),Sum(TodayNum),Count(*) From Dv_board Where ParentStr Like '%"&ParentStr&"%'")
			End If
			Set Trs=Dvbbs.Execute("Select Sum(PostNum),Sum(TopicNum),Sum(TodayNum),Count(*) From Dv_board Where ParentStr Like '%"&ParentStr&"%'")

			If Not (Trs.Eof And Trs.Bof) Then
				'����ð���������������������Ӧ���Ǹð�������+��������������
				If reBoard_Setting(43)="0" Then
					If Not IsNull(Trs(0)) Then AllArticle = Trs(0) + AllArticle
					If Not IsNull(Trs(1)) Then AllTopic = Trs(1) + AllTopic
					If Not IsNull(Trs(2)) Then AllToday = Trs(2) + AllToday
					If Not IsNull(Trs(3)) Then AllBoard = Trs(3) + AllBoard
				Else
					AllArticle=Trs(0)
					AllTopic=Trs(1)
					AllToday=Trs(2)
					AllBoard=Trs(3)
					If IsNull(AllArticle) Then AllArticle=0
					If IsNull(AllTopic) Then AllTopic=0
					If IsNull(AllToday) Then AllToday=0
					If IsNull(AllBoard) Then AllBoard=0
				End If
			End If
			Set Trs=Nothing
			'�������ID
			ParentStr = Sql(0,i)
			Set Trs = Dvbbs.Execute("SELECT Boardid FROM Dv_Board WHERE ParentID = "&Sql(0,i))
			If Not (Trs.Eof And Trs.Bof) Then
				Do While Not Trs.Eof
					ParentStr = ParentStr & "," & Trs(0)
					Trs.Movenext
				Loop
			End If
			Set Trs=Nothing
			'���ظ���Ϣ
			Set Trs=Dvbbs.Execute("Select Top 1 LastPost,TopicID,Title,PostTable From Dv_Topic Where BoardID In ("&ParentStr&") Order by LastPostTime Desc")
			If Not (Trs.Eof And Trs.Bof) Then
				LastPostArr = Split(Trs(0)&"","$")
				If UBound(LastPostArr)<>7 Then
					ReDim LastPostArr(7)
					LastPostArr(1)=Trs(3)
					LastPostArr(3)= Replace(cutStr(Dvbbs.Replacehtml(Trs(2)),20),"$","&#36;")
					LastPostArr(6)=Trs(1)
					Trs.Close:Set Trs=Nothing
					Set Trs = Dvbbs.Execute("Select Top 1 AnnounceID,BoardID,UserName,DateAndTime,PostUserID From "&LastPostArr(1)&" Where RootID="&LastPostArr(6)&" Order by DateAndTime")
					If Not (Trs.Eof And Trs.Bof) Then
						LastPostArr(0) = Trs(2)
						LastPostArr(1) = Trs(0)
						LastPostArr(2) = Trs(3)
						LastPostArr(4) = ""
						LastPostArr(5) = Trs(4)
						LastPostArr(7) = Trs(1)
					Else
						LastPostArr(0) = "��"
						LastPostArr(1) = 0
						LastPostArr(2) = now()
						LastPostArr(3) = "��"
						LastPostArr(4) = ""
						LastPostArr(5) = ""
						LastPostArr(6) = ""
						LastPostArr(7) = ""
					End if
				Else
					LastPostArr(3)= Replace(cutStr(Trs(2),20),"$","&#36;")
				End if
				Trs.Close:Set Trs=Nothing
				LastPost=Join(LastPostArr,"$")
				LastPost=Replace(Dvbbs.Replacehtml(LastPost&""),"'","''")
				'LastPost=Replace(Dvbbs.Replacehtml(Trs(0)&""),"'","''")
			Else
				Trs.Close:Set Trs=Nothing
				LastPost="��$0$"&Now()&"$��$$$$"
			End If
			'���°�������
			'SLastPost = Split(LastPost,"$")
			'If Ubound(SLastPost) < 7 Then LastPost = LastPost & "$"
			Dvbbs.Execute("Update [Dv_board] Set PostNum="&AllArticle&",TopicNum="&AllTopic&",TodayNum="&AllToday&",LastPost='"&Dvbbs.ChkBadWords(LastPost)&"' Where BoardID="&SQL(0,i))
		End If
		'����IsGroupSetting
		'IsGroupSetting=SQL(7,i)
		'Set Trs=Dvbbs.Execute("Select Count(*) From Dv_UserAccess Where uc_BoardID="&SQL(0,i))
		'UserAccessCount = Trs(0)
		'If IsNull(UserAccessCount) Or UserAccessCount="" Then UserAccessCount=0
		'If UserAccessCount>0 Then UpGroupSetting="0"
		'Set Trs=Dvbbs.Execute("Select GroupID From Dv_BoardPermission Where BoardID="&SQL(0,i))
		'If Not Trs.Eof Then
		'	Do While Not Trs.Eof
		'		If UpGroupSetting="" Then
		'			UpGroupSetting = Trs(0)
		'		Else
		'			UpGroupSetting = UpGroupSetting & "," & Trs(0)
		'		End If
		'	Trs.MoveNext
		'	Loop
		'End If
		'���º������̶�������(�̶�������̶�)
		'Set Trs=Dvbbs.Execute("Select TopicID From Dv_Topic Where BoardID="&Rs(0)&" And IsTop In (1,2)")
		If Not IsNull(SQL(6,i)) And SQL(6,i)<>"" Then
		Set Trs=Dvbbs.Execute("Select TopicID,BoardID,IsTop From Dv_Topic Where TopicID In ("&SQL(6,i)&")")
		If tRs.Eof And tRs.Bof Then
			BoardTopStr=""
		Else
			Do While Not Trs.Eof
				If Trs(1)<>444 And Trs(1)<>777 And Trs(2)>0 And Trs(2)<>3 Then
					If BoardTopStr="" Then
						BoardTopStr = Trs(0)
					Else
						BoardTopStr = BoardTopStr & "," & Trs(0)
					End If
				End If
			Trs.MoveNext
			Loop
		End If
		End If
		Dvbbs.Execute("Update Dv_Board Set BoardTopStr='"&BoardTopStr&"' Where BoardID="&SQL(0,i))
		UserAccessCount=""
		IsGroupSetting=""
		UpGroupSetting=""
		BoardTopStr=""
		ii=ii+1
		'If (i mod 100) = 0 Then
			Response.Write "<script>img2.width=" & Fix((ii/C1) * 400) & ";" & VbCrLf
			Response.Write "txt2.innerHTML=""" & FormatNumber(ii/C1*100,4,-1) & """;" & VbCrLf
			Response.Write "img2.title=""" & SQL(0,i) & "(" & ii & ")"";</script>" & VbCrLf
			Response.Flush
		'End If
		body="<table cellpadding=0 cellspacing=0 border=0 width=100% align=center><tr><td colspan=2 class=td1>������̳���ݳɹ���"&SQL(1,i)&"����"&AllArticle&"ƪ���ӣ�"&AllTopic&"ƪ���⣬������"&AllToday&"ƪ���ӡ�</td></tr></table>"
		Response.Write body
		Response.Flush
	'Rs.MoveNext
	'Loop
	Next
	Set Trs=Nothing
	
	End If 
	body=""
	Response.Write "<script>img2.width=400;txt2.innerHTML=""100"";</script>"
	Dvbbs.loadSetup()
	Dim Board
	Dvbbs.LoadBoardList()
	For Each board in Application(Dvbbs.CacheName&"_boardlist").documentElement.selectNodes("board/@boardid")
		Dvbbs.LoadBoardData board.text
		Dvbbs.LoadBoardinformation board.text
	Next
End Sub

Rem ͳ��������̳���� 2004-5-3 Dvbbs.YangZheng
Sub Boardchild()
	Dim cBoardNum, cBoardid
	Dim Trs,rs,Sql,i
	Dim Bn
	Dvbbs.Execute("UPDATE Dv_Board SET Child = 0")
	Set Rs = Dvbbs.Execute("SELECT Boardid, Rootid, ParentID, Depth, Child, ParentStr FROM Dv_Board ORDER BY Boardid DESC")
	If Not (Rs.Eof And Rs.Bof) Then
		Sql = Rs.GetRows(-1)
		Rs.Close:Set Rs = Nothing
		For Bn = 0 To Ubound(Sql,2)
			If Isnull(Sql(4,Bn)) And Cint(Sql(3,Bn)) > 0 Then
				Dvbbs.Execute("UPDATE Dv_Board SET Child = 0 WHERE Boardid = " & Sql(0,Bn))
			End If
			If Cint(Sql(2,Bn)) = 0 And Cint(Sql(3,Bn)) = 0 Then
				Set Trs = Dvbbs.Execute("SELECT COUNT(*) FROM Dv_Board WHERE RootID = " & Sql(1,Bn))
				Cboardnum = Trs(0) - 1
				Trs.Close:Set Trs = Nothing
				If Isnull(Cboardnum) Or Cboardnum < 0 Then Cboardnum = 0
				Dvbbs.Execute("UPDATE Dv_Board SET Child = " & Cboardnum & " WHERE Boardid = " & Sql(0,Bn))
			Elseif Cint(Sql(3,Bn)) > 1 Then
				cBoardid = Split(Sql(5,Bn)&"",",")
				For i = 1 To Ubound(cBoardid)
					Dvbbs.Execute("UPDATE Dv_Board SET Child = Child + 1 WHERE Boardid = " & cBoardid(i))
				Next
			End If
		Next
	End If
End Sub

Sub Updateall()
	'������,������,�û���,������,������,�̶ܹ�,���ע��
	Body = "<tr><td colspan=2 class=td1>��������̳���ݳɹ���"
	Dim AllTopNum,PostNum,TopicNum,LastUser
	Dim TodayNum,UserNum, YesterdayNum,SqlStr,Sql_a,sql
	If Request.Form("u1") = "1" Or Request("index")="1" Then
		TopicNum	= GetTopicnum()
		If Sql_a = "" Then
			Sql_a = "Forum_TopicNum = " & TopicNum & ""
		Else
			Sql_a = Sql_a & ",Forum_TopicNum = " & TopicNum & ""
		End If
		If SqlStr = "" Then
			SqlStr = "��̳���� " & TopicNum & " ƪ����"
		Else
			SqlStr = SqlStr & "��" & TopicNum & " ƪ����"
		End If
	End If
	If Request.Form("u2") = "1" Or Request("index")="1" Then
		PostNum		= Announcenum()
		If Sql_a = "" Then
			Sql_a = "Forum_PostNum = " & PostNum & ""
		Else
			Sql_a = Sql_a & ",Forum_PostNum = " & PostNum & ""
		End If
		If SqlStr = "" Then
			SqlStr = "��̳���� " & PostNum & " ƪ����"
		Else
			SqlStr = SqlStr & "��" & PostNum & " ƪ����"
		End If
	End If
	If Request.Form("u3") = "1" Or Request("index")="1" Then
		UserNum		= Allusers()
		If Sql_a = "" Then
			Sql_a = "Forum_UserNum = " & UserNum & ""
		Else
			Sql_a = Sql_a & ",Forum_UserNum = " & UserNum & ""
		End If
		If SqlStr = "" Then
			SqlStr = "��̳���� " & UserNum & " ���û�"
		Else
			SqlStr = SqlStr & "��" & UserNum & " ���û�"
		End If
	End If
	If Request.Form("u4") = "1" Or Request("index")="1"  Then
		TodayNum	= Alltodays()
		If Sql_a = "" Then
			Sql_a = "Forum_TodayNum = " & TodayNum & ""
		Else
			Sql_a = Sql_a & ",Forum_TodayNum = " & TodayNum & ""
		End If
		If SqlStr = "" Then
			SqlStr = "��̳���� " & TodayNum & " ƪ������"
		Else
			SqlStr = SqlStr & "��" & TodayNum & " ƪ������"
		End If
	End If
	If Request.Form("u5") = "1" Or Request("index")="1"  Then
		YesterdayNum	= Allyesterdays()
		If Sql_a = "" Then
			Sql_a = "Forum_YesterdayNum = " & YesterdayNum & ""
		Else
			Sql_a = Sql_a & ",Forum_YesterdayNum = " & YesterdayNum & ""
		End If
		If SqlStr = "" Then
			SqlStr = "��̳���� " & YesterdayNum & " ƪ������"
		Else
			SqlStr = SqlStr & "��" & YesterdayNum & " ƪ������"
		End If
	End If
	If Request.Form("u6") = "1" Or Request("index")="1" Then
		AllTopNum	= Forum_AllTopNum()
		If Sql_a = "" Then
			Sql_a = "Forum_AllTopNum = '" & AllTopNum & "'"
		Else
			Sql_a = Sql_a & ",Forum_AllTopNum = '" & AllTopNum & "'"
		End If
		If SqlStr = "" Then
			SqlStr = "��̳���� " & UBound(Split(AllTopNum&"", ",")) + 1 & " ���̶�����"
		Else
			SqlStr = SqlStr & "��" & UBound(Split(AllTopNum&"", ",")) + 1 & " ���̶�����"
		End If
	End If
	If Request.Form("u7") = "1" Or Request("index")="1" Then
		LastUser	= Newuser()
		If Sql_a = "" Then
			Sql_a = "Forum_lastUser = '" & Dvbbs.CheckStr(Dvbbs.HtmlEncode(LastUser)) & "'"
		Else
			Sql_a = Sql_a & ",Forum_lastUser = '" & Dvbbs.CheckStr(Dvbbs.HtmlEncode(LastUser)) & "'"
		End If
		If SqlStr = "" Then
			SqlStr = "��̳���¼����û�Ϊ " & LastUser & ""
		Else
			SqlStr = SqlStr & "�����¼����û�Ϊ " & LastUser & ""
		End If
	End If

	Body = Body & SqlStr
	Body = Body & "</td></tr>"

	If Sql_a = "" Then Exit Sub

	Sql = "UPDATE Dv_Setup SET " & Sql_a
	Dvbbs.Execute(Sql)

	Dvbbs.Name="setup"
	Dvbbs.loadSetup()
End sub

Sub fixtopic()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=td1>����Ŀ�ʼ������</td></tr>"
	exit sub
End If
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=td1>����Ľ���������</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=td1>��ʼIDӦ�ñȽ���IDС��</td></tr>"
	exit sub
end if
dim TotalUseTable,Ers,sql,rs,i
dim username,dateandtime,rootid,announceid,postuserid,lastpost,topic
'set rs=Dvbbs.iCreateObject("adodb.recordset")
'Dvbbs.Execute("update Dv_topic set PostTable='dv_bbs1'")
Dim C1
C1=clng(request.form("endid"))-clng(request.form("beginid"))
%>
</table>
&nbsp;<BR>
<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
<tr><td colspan=2 class=td1>
���濪ʼ������̳�������ϣ�Ԥ�Ʊ��ι���<%=C1%>��������Ҫ����
<table width="400" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=#000000>
<table width="400" border="0" cellspacing="0" cellpadding="1">
<tr> 
<td bgcolor=#ffffff height=9><img src="../skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
</td></tr></table> <span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td>
</tr>
</table>

<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
<%
Response.Flush
sql="select topicid,PostTable from Dv_topic where topicid>="&request.form("beginid")&" and topicid<="&request.form("endid")

set rs=Dvbbs.Execute(sql)
if rs.eof and rs.bof then
	body="<tr><td colspan=2 class=td1>�Ѿ�����¼����β�ˣ���������£�</td></tr>"
	exit sub
end if
do while not rs.eof
	sql="select top 1 username,dateandtime,topic,Announceid,PostUserID,rootid,body,boardid from "&rs(1)&" where rootid="&rs(0)&" order by Announceid desc"
	set ers=Dvbbs.Execute(sql)
	if not (ers.eof and ers.bof) then
		username=Ers("username")
		dateandtime=Ers("dateandtime")
		rootid=Ers("rootid")
		topic=left(Ers("body"),20)
		Announceid=ers("Announceid")
		postuserid=ers("postuserid")
		LastPost=username & "$" & Announceid & "$" & dateandtime & "$" & replace(topic,"$","") & "$$" & postuserid & "$" & rootid & "$" & ers("BoardID") & "$"
		LastPost=Dvbbs.Checkstr(LastPost)
		Dvbbs.Execute("update [DV_topic] set LastPost='"&replace(LastPost,"'","")&"',LastPostTime='"&dateandtime&"' where topicid="&rs(0))
		i=i+1
		'If (i mod 100) = 0 Then
		Response.Write "<script>img2.width=" & Fix((i/C1) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""������"&server.htmlencode(ers(2)&"")&"�����ݣ����ڸ�����һ���������ݣ�" & FormatNumber(i/C1*100,4,-1) & """;" & VbCrLf
		Response.Write "img2.title=""" & server.htmlencode(eRs(2)&"") & "(" & i & ")"";</script>" & VbCrLf
		Response.Flush
		'End If
	end if
	'��������� 2004-8-2
	Sql = "SELECT COUNT(*) FROM " & Rs(1) & " WHERE Rootid = " & Rs(0) & " AND Boardid <> 444 AND Boardid <> 777"
	Set Ers = Dvbbs.Execute(Sql)
	Dvbbs.Execute("UPDATE Dv_Topic SET Child = " & Ers(0)-1 & " WHERE Topicid = " & Rs(0) & "")
	Rs.Movenext
loop
set ers=nothing
set rs=nothing
Response.Write "<script>img2.width=400;txt2.innerHTML=""100"";</script>"
%>
<form action="update.asp?action=updat" method=post>
<tr> 
<th style="text-align:center;" colspan=2>�����޸�����(�޸�ָ����Χ�����ӵ����ظ�����)</th>
</tr>
<tr>
<td width="20%" class="td1">��ʼ��ID��</td>
<td width="80%" class="td1"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=5>&nbsp;��������ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="td1">������ID��</td>
<td width="80%" class="td1"><input type=text name="EndID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=5>&nbsp;�����¿�ʼ������ID֮����������ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="td1">&nbsp;</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="�� ��"></td>
</tr>
</form>
<%
end sub

'����̳��������
REM �޸Ĳ�ѯ�������ӱ����� 2004-8-26.Dv.Yz
Function Todays(Boardid)
	Todays = 0
	If IsSqlDataBase = 1 Then
		For i = 0 To Ubound(AllPostTable)
			Set Tmprs = Dvbbs.Execute("SELECT COUNT(Announceid) FROM " & AllPostTable(i) & " WHERE Boardid = " & Boardid & " AND DATEDIFF(day, Dateandtime, " & SqlNowString & ") = 0")
			Todays = Todays + Tmprs(0)
		Next
	Else
		For i = 0 To Ubound(AllPostTable)
			Set Tmprs = Dvbbs.Execute("SELECT COUNT(Announceid) FROM " & AllPostTable(i) & " WHERE Boardid = " & Boardid & " AND DATEDIFF('d', Dateandtime, " & SqlNowString & ") = 0")
			Todays = Todays + Tmprs(0)
		Next
	End If
	Set Tmprs = Nothing
End Function

'ȫ����̳��������
REM �޸Ĳ�ѯ�������ӱ����� 2004-8-26.Dv.Yz
Function Alltodays()
	Dim i
	Alltodays = 0
	If IsSqlDataBase = 1 Then
		For i = 0 To Ubound(AllPostTable)
			Set Tmprs = Dvbbs.Execute("SELECT COUNT(Announceid) FROM " & AllPostTable(i) & " WHERE NOT Boardid IN (444,777) AND DATEDIFF(day, Dateandtime, " & SqlNowString & ") = 0")
			Alltodays = Alltodays + Tmprs(0)
		Next
	Else
		For i = 0 To Ubound(AllPostTable)
			Set Tmprs = Dvbbs.Execute("SELECT COUNT(Announceid) FROM " & AllPostTable(i) & " WHERE NOT Boardid IN (444,777) AND DATEDIFF('d', Dateandtime, " & SqlNowString & ") = 0")
			Alltodays = Alltodays + Tmprs(0)
		Next
	End If
	Set Tmprs = Nothing
End Function

'��̳��������� 2004-8-31.Dv.Yz
Function Allyesterdays()
	Dim i
	Allyesterdays = 0
	If IsSqlDataBase = 1 Then
		For i = 0 To Ubound(AllPostTable)
			Set Tmprs = Dvbbs.Execute("SELECT COUNT(Announceid) FROM " & AllPostTable(i) & " WHERE NOT Boardid IN (444,777) AND DATEDIFF(day, Dateandtime, " & SqlNowString & ") = 1")
			Allyesterdays = Allyesterdays + Tmprs(0)
		Next
	Else
		For i = 0 To Ubound(AllPostTable)
			Set Tmprs = Dvbbs.Execute("SELECT COUNT(Announceid) FROM " & AllPostTable(i) & " WHERE NOT Boardid IN (444,777) AND DATEDIFF('d', Dateandtime, " & SqlNowString & ") = 1")
			Allyesterdays = Allyesterdays + Tmprs(0)
		Next
	End If
	Set Tmprs = Nothing
End Function

'����ע���û�����
function allusers() 
	allusers=Dvbbs.Execute("Select count(userid) from [Dv_user]")(0) 
	If IsNull(allusers) Then allusers=0 
End function
'����ע���û�
Function newuser()
	Dim sql
	sql="Select top 1 username from [Dv_user] order by userid desc"
	Set tmprs=Dvbbs.Execute(sql)
	If tmprs.eof and tmprs.bof Then
		newuser="û�л�Ա"
	Else
   		newuser=tmprs("username")
	End If
	Set tmprs=Nothing 
End function 

'������̳����
function AnnounceNum()
	dim AnnNum,i
	AnnNum=0
	AnnounceNum=0
	For i=0 to ubound(AllPostTable)
		AnnNum=Dvbbs.Execute("Select Count(announceID) from "&AllPostTable(i)&" where not boardid in (444,777)")(0) 
		if isnull(AnnNum) then AnnNum=0
		AnnounceNum=AnnounceNum + AnnNum
	next
end function
'����̳����
function BoardAnnounceNum(boardid)
	dim BoardAnnNum
	BoardAnnNum=0
	BoardAnnounceNum=0
	For i=0 to ubound(AllPostTable)
		BoardAnnNum=Dvbbs.Execute("Select Count(announceID) from "&AllPostTable(i)&" where boardid="&boardid)(0) 
		if isnull(BoardAnnNum) then BoardAnnNum=0
		BoardAnnounceNum=BoardAnnounceNum + BoardAnnNum
	next
end function

'������̳����
function GetTopicnum()
	Dim TopicNum
	TopicNum=Dvbbs.Execute("Select Count(topicid) from DV_topic where not boardid in (444,777)")(0)
	if isnull(TopicNum) then TopicNum=0
	GetTopicnum = TopicNum
end function

'����̳����
function BoardTopicNum(boardid) 
	BoardTopicNum=Dvbbs.Execute("Select Count(topicid) from [Dv_topic] where boardid="&boardid)(0) 
	if isnull(BoardTopicNum) then BoardTopicNum=0 
end function

'��̳�̶ܹ�������
function Forum_AllTopNum()
	Set tmprs=Dvbbs.Execute("Select TopicID From Dv_Topic Where Not BoardID In (444,777) And IsTop=3")
	If tmprs.eof and tmprs.bof Then
		Forum_AllTopNum=""
	Else
		Do While Not tmprs.Eof
			If Forum_AllTopNum="" Then
				Forum_AllTopNum = tmprs(0)
			Else
				Forum_AllTopNum = Forum_AllTopNum & "," & tmprs(0)
			End If
		tmprs.MoveNext
		Loop
	End If
	Set tmprs=Nothing
end function

'�����û�������
sub updateTopic()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=td1>����Ŀ�ʼ������</td></tr>"
	exit sub
end if
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=td1>����Ľ���������</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=td1>��ʼIDӦ�ñȽ���IDС��</td></tr>"
	exit sub
end if
Dim C1
C1=clng(request.form("endid"))-clng(request.form("beginid"))
%>
</table>
&nbsp;<BR>
<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
<tr><td colspan=2 class=td1>
���濪ʼ������̳�û����ϣ�Ԥ�Ʊ��ι���<%=C1%>���û���Ҫ����
<table width="400" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=#000000>
<table width="400" border="0" cellspacing="0" cellpadding="1">
<tr> 
<td bgcolor=#ffffff height=9><img src="../skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
</td></tr></table> <span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td></tr>
</table>

<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
<%
Response.Flush
dim userTopic,UserPost,rs,sql,i
sql="select userid,username from [Dv_user] where userid>="&request.form("beginid")&" and userid<="&request.form("endid")
set rs=Dvbbs.Execute(sql)
if rs.eof and rs.bof then
	body="<tr><td colspan=2 class=td1>�Ѿ�����¼����β�ˣ���������£�</td></tr>"
	exit sub
end if
do while not rs.eof
	UserTopic=UserallTopicnum(rs(0))
	userPost=Userallnum(rs(0))
	Dvbbs.Execute("update [Dv_user] set UserPost="&userPost&",UserTopic="&UserTopic&" where userid="&rs(0))
	i=i+1
	'If (i mod 100) = 0 Then
		Response.Write "<script>img2.width=" & Fix((i/C1) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""������"&rs(1)&"�����ݣ����ڸ�����һ���û����ݣ�" & FormatNumber(i/C1*100,4,-1) & """;" & VbCrLf
		Response.Write "img2.title=""" & Rs(1) & "(" & i & ")"";</script>" & VbCrLf
		Response.Flush
	'End If
rs.movenext
loop
set rs=nothing
Response.Write "<script>img2.width=400;txt2.innerHTML=""100"";</script>"
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th style="text-align:center;" colspan=2>���������û�����</th>
</tr>
<tr>
<td width="20%" class="td1">���¼����û�����</td>
<td width="80%" class="td1">ִ�б�����������<font color=red>��ǰ��̳���ݿ�</font>�������¼��������û���������������</td>
</tr>
<tr>
<td width="20%" class="td1">��ʼ�û�ID</td>
<td width="80%" class="td1"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=10>&nbsp;�û�ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="td1">�����û�ID</td>
<td width="80%" class="td1"><input type=text name="endID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=10>&nbsp;�����¿�ʼ������ID֮����û����ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="td1">&nbsp;</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="���¼����û�����"></td>
</tr>
</form>
<%
end sub

'�����û���Ǯ/����/����
sub updatemoney()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=td1>����Ŀ�ʼ������</td></tr>"
	exit sub
end if
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=td1>����Ľ���������</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=td1>��ʼIDӦ�ñȽ���IDС��</td></tr>"
	exit sub
end if
dim userTopic,userReply,userWealth
dim userEP,userCP

Dim C1,sql,rs,i
C1=clng(request.form("endid"))-clng(request.form("beginid"))
%>
</table>
&nbsp;<BR>
<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
<tr><td colspan=2 class=td1>
���濪ʼ������̳�û����ϣ�Ԥ�Ʊ��ι���<%=C1%>���û���Ҫ����
<table width="400" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=#000000>
<table width="400" border="0" cellspacing="0" cellpadding="1">
<tr> 
<td bgcolor=#ffffff height=9><img src="../skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
</td></tr></table> <span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td></tr>
</table>

<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
<%
Response.Flush
sql="select userlogins,userid,userpost,usertopic,username from [Dv_user] where userid>="&request.form("beginid")&" and userid<="&request.form("endid")
set rs=Dvbbs.Execute(sql)
'shinzeal�����Զ���ʾ���
if rs.eof and rs.bof then
	body="<tr><td colspan=2 class=td1>�Ѿ�����¼����β�ˣ���������£�</td></tr>"
	exit sub
end if
do while not rs.eof
	'userTopic=UserTopicNum(rs(1))
	'userreply=UserReplyNum(rs(1))
	userwealth=rs(0)*Dvbbs.Forum_user(4) + rs("usertopic")*Dvbbs.Forum_user(1) + (rs("userpost")-rs("usertopic"))*Dvbbs.Forum_user(2)
	userEP=rs(0)*Dvbbs.Forum_user(9) + rs("usertopic")*Dvbbs.Forum_user(6) + (rs("userpost")-rs("usertopic"))*Dvbbs.Forum_user(7)
	userCP=rs(0)*Dvbbs.Forum_user(14) + rs("usertopic")*Dvbbs.Forum_user(11) + (rs("userpost")-rs("usertopic"))*Dvbbs.Forum_user(12)
	if isnull(UserWealth) or not isnumeric(userwealth) then userwealth=0
	if isnull(Userep) or not isnumeric(userep) then userep=0
	if isnull(Usercp) or not isnumeric(usercp) then usercp=0
	Dvbbs.Execute("update [Dv_user] set userWealth="&userWealth&",userep="&userep&",usercp="&usercp&" where userid="&rs(1))
	i=i+1
	'If (i mod 100) = 0 Then
		Response.Write "<script>img2.width=" & Fix((i/C1) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""������"&rs(4)&"�����ݣ����ڸ�����һ���û����ݣ�" & FormatNumber(i/C1*100,4,-1) & """;" & VbCrLf
		Response.Write "img2.title=""" & Rs(4) & "(" & i & ")"";</script>" & VbCrLf
		Response.Flush
	'End If
rs.movenext
loop
set rs=nothing
Response.Write "<script>img2.width=400;txt2.innerHTML=""100"";</script>"
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th style="text-align:center;" colspan=2>���������û�����</th>
</tr>
<tr>
<td width="20%" class="td1" valign=top>�����û���Ǯ/����/����</td>
<td width="80%" class="td1">ִ�б�����������<font color=red>��ǰ��̳���ݿ�</font>�û��ķ�����������̳������������¼����û��Ľ�Ǯ/����/������������Ҳ�����¼��������������ܰ���������<BR>ע�⣺���Ƽ��û����б������������������ݺܶ��ʱ���뾡����Ҫʹ�ã����ұ������Ը�������ɾ�����ӵ�������Ӧ��ֵ�������㣬ֻ�ǰ��շ������ܵ���̳��ֵ���ý������㣬�������ز�����<font color=red>���ұ�������������û���Ϊ�������ͷ���ԭ�����Ա���û���ֵ���޸ġ�</font></td>
</tr>
<tr>
<td width="20%" class="td1">��ʼ�û�ID</td>
<td width="80%" class="td1"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=10>&nbsp;�û�ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="td1">�����û�ID</td>
<td width="80%" class="td1"><input type=text name="endID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=10>&nbsp;�����¿�ʼ������ID֮����û����ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="td1">&nbsp;</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="�����û���Ǯ/����/����"></td>
</tr>
</form>
<%
end sub

'�����û��ȼ�
sub updategrade()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=td1>����Ŀ�ʼ������</td></tr>"
	exit sub
end if
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=td1>����Ľ���������</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=td1>��ʼIDӦ�ñȽ���IDС��</td></tr>"
	exit sub
end if

Dim oldMinArticle,Rss,sql,rs
oldMinArticle=0
Set Rss=Dvbbs.Execute("Select UserID From [Dv_User] Where UserID>="&Request.Form("beginid"))
If Rss.Eof And Rss.Bof Then
	body="<tr><td colspan=2 class=td1>�Ѿ�����¼����β�ˣ���������£�</td></tr>"
	Exit Sub
End If
Rss.Close

SQL = "Select UserGroupID From Dv_UserGroups Where Not ParentGID In (0,3)"
Set Rss = Dvbbs.Execute(SQL)
	SQL = Rss.GetString(,, "", ",", "")
Rss.close
Set Rss = Nothing
SQL = SQL & "1"

Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where ParentGID=3 Order By MinArticle Desc")
Do While Not Rs.Eof
	Dvbbs.Execute("Update [Dv_User] Set UserClass='"&Rs("UserTitle")&"',TitlePic='"&Rs("GroupPic")&"',UserGroupID="&Rs("UserGroupID")&" Where (Not UserGroupID In ("&SQL&")) And (UserID>="&Request.Form("beginid")&" And UserID<="&Request.Form("endid")&") And (UserPost<"&oldMinArticle&" And UserPost>="&Rs("MinArticle")&" )")
	oldMinArticle=Rs("MinArticle")
Rs.MoveNext
Loop
Rs.Close
Set Rs=Nothing
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th style="text-align:center;" colspan=2>���������û�����</th>
</tr>
<tr>
<td width="20%" class="td1" valign=top>�����û��ȼ�</td>
<td width="80%" class="td1">ִ�б�����������<font color=red>��ǰ��̳���ݿ�</font>�û�������������̳�ĵȼ��������¼����û��ȼ�����������Ӱ��ȼ�Ϊ������������ܰ��������ݡ�</td>
</tr>
<tr>
<td width="20%" class="td1">��ʼ�û�ID</td>
<td width="80%" class="td1"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=10>&nbsp;�û�ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="td1">�����û�ID</td>
<td width="80%" class="td1"><input type=text name="endID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=10>&nbsp;�����¿�ʼ������ID֮����û����ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="td1">&nbsp;</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="�����û��ȼ�"></td>
</tr>
</form>
<%
end sub

'�û�����������
function UserTopicNum(userid)
	dim topicnum
	topicnum=0
	usertopicnum=0
	set tmprs=Dvbbs.Execute("select count(*) from dv_topic where not boardid in (444,777) and PostUserID="&userid)
	TopicNum=tmprs(0)
	if isnull(TopicNum) then TopicNum=0
	UserTopicNum=UserTopicNum + TopicNum
	set tmprs=nothing
end function
'�û����лظ���
Function UserReplyNum(userid)
	dim replynum,i
	replynum=0
	userreplynum=0
	For i=0 to ubound(AllPostTable)
		set tmprs=Dvbbs.Execute("select count(announceid) from "&AllPostTable(i)&" where not boardid in (444,777) and ParentID>0 and PostUserID="&userid)
		replyNum=tmprs(0)
		if isnull(replyNum) then replyNum=0
		UserReplyNum=UserReplyNum + replynum
	next
	set tmprs=nothing
end function
'�û���������
function Userallnum(userid)
	dim allnum,i
	allnum=0
	userallnum=0
	For i=0 to ubound(AllPostTable)
		set tmprs=Dvbbs.Execute("select count(announceid) from "&AllPostTable(i)&" where not boardid in (444,777) and PostUserID="&userid)
		allnum=tmprs(0)
		if isnull(allnum) then allnum=0
		userallnum=userallnum+allnum
	Next
	Set tmprs=nothing
End function

function UserallTopicnum(userid)
	dim allnum,i
	allnum=0
	UserallTopicnum=0
	For i=0 to ubound(AllPostTable)
		set tmprs=Dvbbs.Execute("select count(*) from Dv_Topic where not boardid in (444,777) and PostUserID="&userid)
		allnum=tmprs(0)
		if isnull(allnum) then allnum=0
		UserallTopicnum=UserallTopicnum+allnum
	Next
	Set tmprs=nothing
End function

Sub fixbbs()
Dim killhtml,updatedate
updatedate=Request("updatedate")
killhtml=Request("killhtml")
If Not IsNumeric(request.form("beginid")) Then
	body="<tr><td colspan=2 class=td1>����Ŀ�ʼ������</td></tr>"
	Exit Sub
End If
If Not IsNumeric(request.form("endid")) Then
	body="<tr><td colspan=2 class=td1>����Ľ���������</td></tr>"
	Exit Sub
End If
If CLng(request.form("beginid"))>clng(request.form("endid")) Then
	body="<tr><td colspan=2 class=td1>��ʼIDӦ�ñȽ���IDС��</td></tr>"
	Exit Sub
End If
Dim C1, D1
C1 = clng(request.form("endid"))-clng(request.form("beginid"))
D1 = Clng(Request.Form("beginid"))
%>
</table>
&nbsp;<BR>
<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
<tr><td colspan=2 class=td1>
���濪ʼ������̳�������ϣ�Ԥ�Ʊ��ι���<%=C1%>��������Ҫ����
<table width="400" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=#000000>
<table width="400" border="0" cellspacing="0" cellpadding="1">
<tr> 
<td bgcolor=#ffffff height=9><img src="../skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
</td></tr></table>
<span id=txt3 name=txt3 style="font-size:9pt;color:red;"></span>
<span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td></tr>
</table>

<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
<%
Response.Flush
Dim TotalUseTable,Ers,SQL1
Dim vBody,isagree,re
Dim Maxid,Rs,Sql,i
Set Rs=Dvbbs.Execute("select Max(topicid) from dv_topic")
Maxid=Rs(0)
Set Rs=Nothing
If Maxid< CLng(request.form("beginid")) Then
	body="<tr><td colspan=2 class=td1>�Ѿ�����¼����β�ˣ���������£�</td></tr>"
	set rs=nothing
	Exit Sub
End If
Sql = "SELECT Topicid, PostTable FROM Dv_Topic WHERE Topicid >= " & CLng(Request.Form("beginid")) & " AND Topicid <= " & CLng(Request.Form("endid")) & " ORDER BY TopicID"
Set Rs=Dvbbs.Execute(sql)
Set  ERs=Dvbbs.iCreateObject("adodb.recordset")
Do While Not rs.eof
	SQl1 ="select Body,isagree,Ubblist,topic,DateAndTime,Announceid from "&Rs(1)&" where Rootid="&CLng(Rs(0))&""
	Ers.open SQL1,Conn,1,3
	If Not(eRs.eof OR ers.BOF) Then
		Do While Not ers.eof
			vbody=eRs(0)
			isagree=eRs(1)
			If IsNull(isagree) Then isagree=""
			If isagree <>""  Then
				isagree=Replace(isagree,"[isubb]","")
			Else
				isagree=ers(1)
			End If
			If killhtml="1" Then
				eRs(0)=vbody
			End If
			ers(1)=	isagree&""
			If IsDate(updatedate) And IsDate(ers(4)) Then
				If updatedate > ers(4) Then
					ers(2)=Ubblist(vbody)
				Else
					If killhtml="1" Then
						Set re=new RegExp
						re.IgnoreCase =true
						re.Global=True
						re.Pattern="<br>"
						vbody=re.Replace(vbody,"[br]")
						re.Pattern="<(.[^>]*)>"
						vbody=re.Replace(vbody,"")
						Set re=Nothing
					End If 
					ers(2)=UbblistOLD(vbody)
				End If
			Else
				ers(2)=UbblistOLD(vbody)
			End If
		ers.update
			Response.Write "<script>txt3.innerHTML=""��������Ϊ"&eRs(5)&"���������ݣ�"";</script>"
			Response.Flush
		eRs.movenext 
		Loop
	End If
	ERs.close
	i = CLng(Rs(0)) - D1 + 1	'����iΪ���λ��ֵ 2005-12-3 Dv.Yz
	'If (i mod 100) = 0 Then
		Response.Write "<script>img2.width=" & Fix((i/C1) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""��������Ϊ"&Rs(0)&"�����ݣ����ڸ�����һ���������ݣ�" & FormatNumber(i/C1*100,4,-1) & """;" & VbCrLf
		Response.Write "img2.title=""" & Rs(0) & "(" & i & ")"";</script>" & VbCrLf
		Response.Flush
	'End If
Rs.movenext 
Loop
Set ers=nothing
set rs=nothing
Response.Write "<script>img2.width=400;txt3.innerHTML='';txt2.innerHTML=""100"";</script>"
%>
<form name=Fix action="update.asp?action=fix" method=post>
<tr> 
<th style="text-align:center;" colspan=2>�����޸�����(�޸�ָ����Χ����UBB��ǩ)</th>
</tr>
<tr>
<td width="20%" class="td1">��ʼ��ID��</td>
<td width="80%" class="td1"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=5>&nbsp;��������ID��������д�������һ��ID�ſ�ʼ�����޸�</td>
</tr>
<tr>
<td width="20%" class="td1">������ID��</td>
<td width="80%" class="td1"><input type=text name="EndID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=5>&nbsp;�����¿�ʼ������ID֮����������ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
</tr>
<tr>
<td width="20%" class="td1">�������ı�ʶ����</td>
<td width="80%" class="td1"><input type="text" name="updatedate" value="<%=updatedate%>">(��ʽ��YYYY-M-D) ������̳������v7.0�����ڣ��������д��һ�ɰ���������</td>
</tr>
<tr>
<td width="20%" class="td1">ȥ�������е�HTML���</td>
<td width="80%" class="td1"><input type="radio" class="radio" name="killhtml" value="1" 
 <%
  If killhtml="1" Then 
  %>
  checked 
  <%
  End If 
  %>
> ��
  <input type="radio" class="radio" name="killhtml" value="0" 
  <%
  If killhtml="0" Then 
  %>
  checked 
  <%
  End If 
  %>
  > �� &nbsp;<br>ѡ�ǵĻ��������е�HTML��ǽ����Զ�������������ڼ������ݿ�Ĵ�С�����ǻ�ʧȥԭ����HTMLЧ����</td>
</tr>
<tr>
<td width="20%" class="td1">&nbsp;</td>
<td width="80%" class="td1"><input type="submit" class="button" name="Submit" value="�� ��"></td>
</tr>
</form>
<%
End Sub

'��������û� 2004-10-11 Dv.Yz
Sub Delallonline()
	Dim Sql
	Sql = "DELETE FROM Dv_Online"
	Dvbbs.Execute(Sql)
	Body = "<tr><td colspan=2 class=td1>��������û����ݳɹ����� <a href=""ReloadForumCache.asp""><font color=""red"">���·���������</font> </a>��"
End Sub

'������̳�ղؼ� 2005-11-30 Dv.Yz
Sub Updatebm()
	Dim bi,Rs,Sql
	Set Rs = Dvbbs.Execute("SELECT Id, Url, UserName FROM Dv_BookMark ORDER BY Id")
	If Not (Rs.Eof And Rs.Bof) Then
		Sql = Rs.GetRows(-1)
		Set Rs = Nothing
		Dim Url, RootID, BoardID, Topic
		Dim C1
		C1 = Ubound(Sql,2) + 1
%>
<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
<tr><td colspan=2 class=td1>
���濪ʼ�����ղؼ����ϣ�Ԥ�Ʊ��ι���<%=C1%>���ղ�������Ҫ����
<table width="400" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=#000000>
<table width="400" border="0" cellspacing="0" cellpadding="1">
<tr> 
<td bgcolor=#ffffff height=9><img src="../skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
</td></tr></table>
<span id=txt3 name=txt3 style="font-size:9pt;color:red;"></span>
<span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td></tr>
</table>
<%
		For bi = 0 To Ubound(Sql,2)
			Set Rs = Dvbbs.Execute("SELECT UserName From Dv_User WHERE UserName = '" & Sql(2,bi) & "'")
			If Rs.Eof And Rs.Bof Then
				Dvbbs.Execute("DELETE FROM Dv_BookMark WHERE Id = " & Sql(0,bi))
				Response.Write "<script>txt3.innerHTML=""ɾ������Ϊ" & Sql(0,bi) & "���ղؼ����ݣ�"";</script>"
				Response.Flush
			Else
				Set Rs = Nothing
				RootID = Split(Split(Sql(1,bi),"&")(1),"=")(1)
				If Isnumeric(RootID) Then
					RootID = Clng(Rootid)
				Else
					RootID = 0
				End If
				Set Rs = Dvbbs.Execute("SELECT BoardID, TopicID, Title FROM Dv_Topic WHERE NOT BoardID IN (444,777) AND TopicID = " & RootID)
				If Rs.Eof And Rs.Bof Then
					Dvbbs.Execute("DELETE FROM Dv_BookMark WHERE Id = " & Sql(0,bi))
					Response.Write "<script>txt3.innerHTML=""ɾ������Ϊ" & Sql(0,bi) & "���ղؼ����ݣ�"";</script>"
					Response.Flush
				Else
					Boardid = Rs(0)
					RootID = Rs(1)
					Topic = Rs(2)
					Topic = Left(Dvbbs.checkStr(trim(topic)),100)
					Url = "dispbbs.asp?boardid=" & Boardid & "&id=" & RootID
					Dvbbs.Execute("UPDATE Dv_BookMark SET Url = '" & Url & "', Topic = '" & Topic & "' WHERE Id = " & Sql(0,bi))
					Response.Write "<script>txt3.innerHTML=""��������Ϊ" & Sql(0,bi) & "���ղؼ����ݣ�"";</script>"
					Response.Flush
				End If
			End If
			Response.Write "<script>img2.width=" & Fix((bi/C1) * 400) & ";" & VbCrLf
			Response.Write "txt2.innerHTML=""��������Ϊ" & Sql(0,bi) & "�����ݣ����ڸ�����һ���ղؼ����ݣ�" & FormatNumber(bi/C1*100,4,-1) & """;" & VbCrLf
			Response.Write "img2.title=""" & Sql(2,bi) & "(" & bi & ")"";</script>" & VbCrLf
			Response.Flush
		Next
		Response.Write "<script>img2.width=400;txt3.innerHTML='������ɣ�<a href=" & Request.ServerVariables("HTTP_REFERER") & "><<������һҳ</a>';txt2.innerHTML=""100"";</script>"
	Else
	%>
		<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
			<tr><td colspan=2 class=td1>�ղؼ�����Ϊ�գ�����Ҫ���£�</td></tr>
		</table>
	<%
	End If
	
End Sub

Function cutStr(str,strlen)
	Str=Dvbbs.Replacehtml(Str)
	Dim l,t,c,i
	l=Len(str)
	t=0
	For i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		If c>255 Then
			t=t+2
		Else
			t=t+1
		End If
		If t>=strlen Then
			cutStr=left(str,i)&"..."
			Exit For
		Else
			cutStr=str
		End If
	Next
	cutStr=Replace(cutStr,chr(10),"")
	cutStr=Replace(cutStr,chr(13),"")
End Function
%>