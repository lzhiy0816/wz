<!--#include file="../conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
	Head()
	dim Str
	dim admin_flag
	admin_flag=",11,"
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=../admin_login.asp target=_top>��¼</a>����롣<br><li>��û�й�����ҳ���Ȩ�ޡ�"
		dvbbs_error()
	else
		if Request("action") = "unite" then
		call unite()
		else
		call boardinfo()
		end if
	Footer()
	end if

sub boardinfo()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr>
	<th height=25>�ϲ���̳����
	</th>
	</tr>
	<form action=boardunite.asp?action=unite method=post>
	<tr>
	<td class=forumrow>
	<B>�ϲ���̳ѡ��</B>��<BR>
<B>������̳����������������Ӷ�ת����Ŀ����̳����ɾ������̳������������</B><BR><BR>
<%
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select boardid,boardtype,depth from dv_board order by rootid,orders"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "û����̳"
	else
		response.write " ����̳ "
		response.write "<select name=oldboard size=1>"
		do while not rs.eof
%>
<option value="<%=rs("boardid")%>"><%if rs("depth")>0 then%>
<%for i=1 to rs("depth")%>
��
<%next%>
<%end if%><%=rs("boardtype")%></option>
<%
		rs.movenext
		loop
		response.write "</select>"
	end if
	rs.close
	sql="select boardid,boardtype,depth from dv_board order by rootid,orders"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "û����̳"
	else
		response.write " �ϲ��� "
		response.write "<select name=newboard size=1>"
		do while not rs.eof
%>
<option value="<%=rs("boardid")%>"><%if rs("depth")>0 then%>
<%for i=1 to rs("depth")%>
��
<%next%>
<%end if%><%=rs("boardtype")%></option>
<%
		rs.movenext
		loop
		response.write "</select>"
	end if
	rs.close
	set rs=nothing
	response.write " <BR><BR><input type=submit name=Submit value=�ϲ���̳><BR><BR>"
%>
	</td>
	</tr>
	<tr>
	<td class=forumrow><B>ע������</B>��<BR><FONT COLOR="red">���в��������棬�����ز���</FONT><BR> ������ͬһ�������ڽ��в��������ܽ�һ������ϲ�����������̳�С�<BR>�ϲ�������ָ������̳�����߰�����������̳������ɾ�����������ӽ�ת�Ƶ�����ָ����Ŀ����̳��
	</td>
	</tr></form>
	</table>
<%
end sub

Sub Unite()
	Dim Newboard
	Dim Oldboard
	Dim ParentStr, iParentStr
	Dim Depth, iParentID, Child
	Dim ParentID, RootID
	If Clng(Request("newboard")) = Clng(Request("oldboard")) Then
		Errmsg = "�벻Ҫ����ͬ�����ڽ��в�����"
		dvbbs_error()
		Exit Sub
	End If
	Newboard = Clng(Request("newboard"))
	Oldboard = Clng(Request("oldboard"))
	'������̳����������������Ӷ�ת����Ŀ����̳����ɾ������̳������������
	'�õ���ǰ����������̳
set rs=Dvbbs.Execute("select ParentStr,Boardid,depth,ParentID,child,RootID from dv_board where boardid="&oldboard)
if rs(0)="0" then
	ParentStr=rs(1)
	iParentID=rs(1)
	ParentID=0
else
	ParentStr=rs(0) & "," & Rs(1)
	iParentID=rs(3)
	ParentID=rs(3)
end if
iParentStr=rs(1)
depth=rs(2)
child=rs(4)+1
RootID=rs(5)
i=0
If ParentID=0 Then
set rs=Dvbbs.Execute("select Boardid from dv_board where boardid="&newboard&" and RootID="&RootID)
Else
set rs=Dvbbs.Execute("select Boardid from dv_board where boardid="&newboard&" and ParentStr like '%"&ParentStr&"%'")
End If
if not (rs.eof and rs.bof) then
	response.write "���ܽ�һ������ϲ�����������̳��"
	exit sub
end if
'�õ���ǰ����������̳ID
i=0
set rs=Dvbbs.Execute("select Boardid from dv_board where RootID="&RootID&" And ParentStr like '%"&ParentStr&"%'")
if not (rs.eof and rs.bof) then
do while not rs.eof
	if i=0 then
		iParentStr=rs(0)
	else
		iParentStr=iParentStr & "," & rs(0)
	end if
	i=i+1
rs.movenext
loop
end if
if i>0 then
	ParentStr=iParentStr & "," & oldboard
else
	ParentStr=oldboard
end if
'������ԭ��������̳������
if depth>0 then
Dvbbs.Execute("update dv_board set child=child-"&child&" where boardid="&iparentid)
'������ԭ��������̳���ݣ������൱�ڼ�֦�����迼��
for i=1 to depth
	'�õ��丸��ĸ���İ���ID
	set rs=Dvbbs.Execute("select parentid from dv_board where boardid="&iparentid)
	if not (rs.eof and rs.bof) then
		iparentid=rs(0)
		Dvbbs.Execute("update dv_board set child=child-"&child&" where boardid="&iparentid)
	end if
next
end if
Conn.CommandTimeOut = 0
'������̳��������
For i=0 to ubound(AllPostTable)
	Dvbbs.Execute("update "&AllPostTable(i)&" set boardid="&newboard&" where boardid in ("&ParentStr&")")
	'���»���վ��������
	Dvbbs.Execute("Update "&AllPostTable(i)&" Set LockTopic="&newboard&" Where BoardID=444 And LockTopic In ("&ParentStr&")")
	Response.Write "������̳���ӱ�" & AllPostTable(i) & "���ݳɹ���<br>"
	Response.Flush
Next
Dvbbs.Execute("update dv_topic set boardid="&newboard&",mode=0 where boardid in ("&ParentStr&")")
Response.Write "������������ݳɹ���<br>"
Response.Flush
Dvbbs.Execute("update dv_besttopic set boardid="&newboard&" where boardid in ("&ParentStr&")")
Response.Write "���¾��������ݳɹ���<br>"
Response.Flush
'���»���վ��������
Dvbbs.Execute("Update Dv_Topic Set LockTopic="&newboard&" Where BoardID=444 And LockTopic In ("&ParentStr&")")
Response.Write "���»���վ���ݳɹ���<br>"
Response.Flush
'shinzeal��������ϴ��ļ�����
Dvbbs.Execute("update DV_Upfile set F_boardid="&newboard&" where F_boardid in ("&ParentStr&")")
Response.Write "�����ϴ������ݳɹ���<br>"
Response.Flush
'ɾ�����ϲ���̳
set rs=Dvbbs.Execute("select sum(postnum),sum(topicnum),sum(todayNum) from dv_board where RootID="&RootID&" And boardid in ("&ParentStr&")")
Dvbbs.Execute("delete from dv_board where RootID="&RootID&" And boardid in ("&ParentStr&")")
	'ɾ�����ϲ���̳���Զ����û�Ȩ��
	Dvbbs.Execute("DELETE FROM Dv_UserAccess WHERE NOT Uc_BoardID IN (SELECT BoardID FROM Dv_Board)")
'��������̳���Ӽ���
dim trs
set trs=Dvbbs.Execute("select ParentStr,boardid from dv_board where boardid="&newboard)
if trs(0)="0" then
ParentStr=trs(1)
else
ParentStr=trs(0)
end if
	'���ºϲ��������������Ϣ
	Dvbbs.Execute("UPDATE Dv_Board SET Postnum = Postnum + " & Rs(0) & ", Topicnum = Topicnum + " & Rs(1) & ", Todaynum = Todaynum + " & Rs(2) & " WHERE Boardid = " & NewBoard)
	Response.Write "�ϲ��ɹ����Ѿ������ϲ���̳��������ת�������ϲ�����̳�������һ����̳���ݡ�"
set rs=nothing
set trs=nothing
RestoreBoardCache()
End Sub
Sub RestoreBoardCache()
	Dim Board
	Dvbbs. LoadBoardList()
	For Each board in Application(Dvbbs.CacheName&"_boardlist").documentElement.selectNodes("board/@boardid")
		Dvbbs.LoadBoardData(board.text)
	Next
	If Request("action")="RestoreBoardCache" Then dv_suc("�ؽ����а��滺��ɹ���")
End Sub
%>