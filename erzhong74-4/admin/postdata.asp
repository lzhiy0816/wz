<!--#include file=../conn.asp-->
<!--#include file="inc/const.asp"-->
<%
Head()
Server.ScriptTimeOut=9999999
Dim admin_flag
admin_flag=Split("26,27",",")
CheckAdmin(","&admin_flag(0)&",")
CheckAdmin(","&admin_flag(1)&",")
Select Case LCase(Request("action"))
	Case "nowused" : Call NowUsed()
	Case "update" : Call Update()
	Case "del" : Call Del()
	Case "creattable" : Call CreatTable()
	Case "search" : Call Search()
	Case "update2" : Call Update2()
	Case "update3" : Call Update3()
	Case Else
		Call main()
End Select
Footer()

Sub main()
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
	<tr>
		<td height="23" colspan="2" class=td1><B>˵��</B>��<BR>������ѡ����������֮һ��ģʽ�������������ڲ�ͬ��֮���ת����</td>
	</tr>
	<tr>
		<th colspan="2">ģʽһ������Ҫת�Ƶ�����</th>
	</tr>
	<FORM METHOD=POST ACTION="?action=search">
	<tr>
		<td height="23" width="20%" class=td1><B>��������</B></td>
		<td height="23" width="80%" class=td1>
			<input type="text" name="keyword">&nbsp;
			<select name="tablename">
			<%=TableSelectForm%>
			</select>&nbsp;<!--shinzeal edit username and topic to searchWhat in 2004/7/4 ���ֵ�ѡ�ı굥��radio�ȽϺ���-->
			<input type="radio" class="radio" name="searchWhat" value="username" checked>�û�&nbsp;
			<input type="radio" class="radio" name="searchWhat" value="topic">����&nbsp;
			<input type="submit" class="button" name="submit" value="����">
		</td>
	</tr>
	</FORM>
	<tr>
		<td height="23" colspan="2" class=td1><B>ע��</B>��������������ڱ�������ͷ����û����ݣ������������������в���</td>
	</tr>
	<tr>
		<th colspan="2">ģʽ�����ڲ�ͬ��ת������</th>
	</tr>
	<FORM METHOD=POST ACTION="?action=update2">
	<tr>
		<td height="23" width="100%" class=td1 colspan="2">&nbsp;
		<select name="OutTablename">
		<%=TableSelectForm%>
		</select><!--shinzeal edit top1000 and end1000 to TopOrEnd in 2004/7/4 ���ֵ�ѡ�ı굥��radio�ȽϺ���-->
		 <input type="radio" class="radio" name="TopOrEnd" value="Top" checked>ǰ <input type="radio" class="radio" name="TopOrEnd" value="End">�� <input type=text name="selnum" value="100" size=3>�� ��¼ת�Ƶ�
		<select name="InTablename">
		<%=TableSelectForm%>
		</select>
		&nbsp;<input type="submit" class="button" name="submit" value="�ύ">
		</td>
	</tr>
	</FORM>
	<tr>
		<td height="23" colspan="2" class=td1><B>ע��</B>����ǰN����¼ָ���ݿ������緢�������ӣ����ƽ��ÿ��������5���ظ�����ô100������������ĸ���������500����¼������ͨ��Ҫ���ܳ���ʱ�䣬���µ��ٶ�ȡ�������ķ����������Լ��������ݵĶ��١�ִ�б����轫���Ĵ����ķ�������Դ���������ڷ����������ٵ�ʱ����߱��ؽ��и��²�����</td>
	</tr>
</table>
<%
end sub

sub nowused()
%>
<form method="POST" action="?action=update">
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr>
<td height="23" colspan="5" class=td1><B>˵��</B>��<BR>�������ݱ���ѡ�е�Ϊ��ǰ��̳��ʹ���������������ݵı���һ�������ÿ�����е�����Խ����̳������ʾ�ٶ�Խ�죬�������е����������ݱ��е������г������������ʱ��������һ�����ݱ��������������ݣ�SQL�汾�û�����ÿ�������ݴﵽ20���Ժ�������ӱ������������ᷢ����̳�ٶȿ�ܶ�ܶࡣ<BR>��Ҳ���Խ���ǰ��ʹ�õ����ݱ����������ݱ����л�����ǰ��ʹ�õ��������ݱ�����ǰ��̳�û�����ʱĬ�ϵı����������ݱ�</td>
</tr>
<tr>
<th colspan="5">��ǰ���ݱ��趨</th>
</tr>
<tr>
<td width="20%"><b>����<B></td>
<td width="20%"><B>˵��</B></td>
<td width="20%"><B>��ǰ����</B></td>
<td width="20%"><B>��ǰĬ��</B></td>
<td width="20%"><B>ɾ��</B></td>
</tr>
<%
Dim i,rs
for i=0 to ubound(AllPostTable)
	On Error Resume Next
	Set Rs=Dvbbs.Execute("select count(*) from "&AllPostTable(i)&"")
	If Err Then
		Err.Clear
%>
	<tr>
		<td width="20%" class=td1><%=AllPostTable(i)%></td>
		<td width="20%" class=td1>�ü�¼��Ӧ���������ݱ������ڻ��Ѿ���ɾ��</td>
		<td width="20%" class=td1>-1</td>
		<td width="20%" class=td1><input value="<%=AllPostTable(i)%>" name="TableName" type="radio" class="radio" disabled="true"></td>
		<td width="20%" class=td1><a href="?action=del&tablename=<%=AllPostTable(i)%>">ɾ��</a></td>
	</tr>
<%	Else%>
	<tr>
		<td width="20%" class=td1><%=AllPostTable(i)%></td>
		<td width="20%" class=td1><%=AllPostTableName(i)%></td>
		<td width="20%" class=td1><%=Rs(0)%></td>
		<td width="20%" class=td1><input value="<%=AllPostTable(i)%>" name="TableName" type="radio" class="radio" <%if Trim(Lcase(Dvbbs.NowUseBBS))=Lcase(AllPostTable(i)) then%>checked<%end if%>></td>
		<td width="20%" class=td1><a href="?action=del&tablename=<%=AllPostTable(i)%>"  onclick="{if(confirm('ɾ�������������ݱ��������ӣ���������ɾ�������ݽ����ɻָ���ȷ��ɾ����?')){return true;}return false;}">ɾ��</a></td>
	</tr>
<%
	End IF
next
%>
<tr>
<td width="100%" colspan=5 class=td1>
<input type="submit" class="button" name="Submit" value="�� ��">
</td>
</tr>
</form>
<FORM METHOD=POST ACTION="?action=CreatTable">
<tr>
<th colspan="5">�������ݱ�</th>
</tr>
<tr>
<td width="20%" class=td1>���ӵı���</td>
<td width="80%" class=td1 colspan=4><input type=text name="tablename" value="Dv_bbs<%=ubound(AllPostTable)+2%>">&nbsp;ֻ����Dv_bbs+���ֱ�ʾ����Dv_bbs5������������಻�ܳ���9</td>
</tr>
<tr>
<td width="20%" class=td1>���ӱ���˵��</td>
<td width="80%" class=td1 colspan=4><input type=text name="tablereadme">&nbsp;�������ñ�����;�����������Ӻ�������ز���������ʾ</td>
</tr>
<tr>
<td width="100%" colspan=5 class=td1>
<input type="submit" class="button" name="Submit" value="�� ��">
</td>
</tr>
</FORM>
</table>
<%
end sub

sub update()
	Dvbbs.Execute("update Dv_setup set Forum_NowUseBBS='"&Dvbbs.CheckStr(request.form("TableName"))&"'")
	Dvbbs.loadSetup()
	Dv_suc("���³ɹ���")
end sub

sub del()
	dim nAllPostTable,nAllPostTableName,ii,TableName
	TableName = Dvbbs.CheckStr(trim(request("tablename")))
	if TableName=Trim(Dvbbs.NowUseBBS) then
		Errmsg=ErrMsg + "<BR><li>��ǰ����ʹ�õı�����ɾ����"
		dvbbs_error()
		exit sub
	end if
	Dvbbs.Execute("delete from dv_Tablelist where TableName='"&TableName&"'")
	Dvbbs.Execute("delete from dv_BestTopic where RootID in (select TopicID from dv_topic where PostTable='"&request("tablename")&"')")
	Dvbbs.Execute("delete from dv_Topic where PostTable='"&TableName&"'")
	On Error Resume Next
	Dvbbs.Execute("drop table "&TableName&"")
	If Err Then Err.Clear
	Dv_suc("ɾ���ɹ���")
end sub

sub CreatTable()
if request.form("tablename")="" then
	Errmsg=ErrMsg + "<BR><li>�����������"
	dvbbs_error()
	exit sub
elseif len(request.form("tablename"))<>7 then
	Errmsg=ErrMsg + "<BR><li>����ı������Ϸ���"
	dvbbs_error()
	exit sub
elseif not isnumeric(right(request.form("tablename"),1)) then
	Errmsg=ErrMsg + "<BR><li>����ı������Ϸ���"
	dvbbs_error()
	exit sub
elseif cint(right(request.form("tablename"),1))>9 or cint(right(request.form("tablename"),1))<0 then
	Errmsg=ErrMsg + "<BR><li>����ı������Ϸ���"
	dvbbs_error()
	exit sub
end if
if request.form("tablereadme")="" then
	Errmsg=ErrMsg + "<BR><li>���������˵����"
	dvbbs_error()
	exit sub
end If
Dim i,sql
for i=0 to ubound(AllPostTable)
	if AllPostTable(i)=request.form("tablename") then
		Errmsg=ErrMsg + "<BR><li>������ı����Ѿ����ڣ����������롣"
		dvbbs_error()
		exit sub
	end if
next

Dim NewAllPostTable,NewAllPostTableName
'�������ݱ��б�

Dvbbs.Execute("insert into dv_TableList(TableName,TableType)Values('"&request.form("tablename")&"','"&request.form("tablereadme")&"') ")
'NewAllPostTable=rs(0) & "|" & request.form("tablename")
'NewAllPostTableName=rs(1) & "|" & request.form("tablereadme")

'Set conn = Dvbbs.iCreateObject("ADODB.connection")
'connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("dvbbs5.mdb")
'conn.open connstr
'�����±�
If IsSqlDataBase=1 Then
	sql="CREATE TABLE ["&request.form("tablename")&"] (AnnounceID int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_"&request.form("tablename")&" PRIMARY KEY,"&_
		"ParentID int default 0,"&_
		"BoardID int default 0,"&_
		"UserName varchar(50),"&_
		"PostUserID int default 0,"&_
		"Topic varchar(250),"&_
		"Body text,"&_
		"DateAndTime smalldatetime default "&SqlNowString&","&_
		"length int Default 0,"&_
		"RootID int Default 0,"&_
		"layer int Default 0,"&_
		"orders int Default 0,"&_
		"isbest tinyint Default 0,"&_
		"ip varchar(40) NULL,"&_
		"Expression varchar(100) NULL,"&_
		"locktopic int Default 0,"&_
		"signflag tinyint Default 0,"&_
		"emailflag tinyint Default 0,"&_
		"isagree varchar(50) NULL,"&_
		"isupload tinyint default 0,"&_
		"isaudit tinyint default 0,"&_
		"PostBuyUser text,"&_
		"UbbList varchar(255),"&_
		"GetMoney int not null Default 0,"&_
		"UseTools varchar(255),"&_
		"GetMoneyType tinyint not null Default 0,"&_
		"FlashId varchar(250) NULL Default 0"&_
		")"
Else
	sql="CREATE TABLE "&request.form("tablename")&" (AnnounceID int IDENTITY (1, 1) NOT NULL CONSTRAINT PrimaryKey PRIMARY KEY,"&_
		"ParentID int default 0,"&_
		"BoardID int default 0,"&_
		"UserName varchar(50),"&_
		"PostUserID int default 0,"&_
		"Topic varchar(250),"&_
		"Body text,"&_
		"DateAndTime smalldatetime default Now(),"&_
		"length int Default 0,"&_
		"RootID int Default 0,"&_
		"layer int Default 0,"&_
		"orders int Default 0,"&_
		"isbest tinyint Default 0,"&_
		"ip varchar(40) NULL,"&_
		"Expression varchar(100) NULL,"&_
		"locktopic int Default 0,"&_
		"signflag tinyint Default 0,"&_
		"emailflag tinyint Default 0,"&_
		"isagree varchar(50) NULL,"&_
		"isupload tinyint default 0,"&_
		"isaudit tinyint default 0,"&_
		"PostBuyUser text,"&_
		"UbbList varchar(255),"&_
		"GetMoney int not null Default 0,"&_
		"UseTools varchar(255),"&_
		"GetMoneyType tinyint not null Default 0,"&_
		"FlashId varchar(250) NULL Default 0"&_
		")"
End If
Dvbbs.Execute(sql)

'��������
Dvbbs.Execute("create index dispbbs on "&request.form("tablename")&" (boardid,rootid)")
Dvbbs.Execute("create index save_1 on "&request.form("tablename")&" (rootid,orders)")
Dvbbs.Execute("create index disp on "&request.form("tablename")&" (boardid)")
Dvbbs.Execute("create index PostUserID on "&request.form("tablename")&" (PostUserID)")
'Dvbbs.Execute("update config set AllPostTable='"&NewAllPostTable&"',AllPostTableName='"&NewAllPostTableName&"'")
Dv_suc("���ӱ��ɹ����뷵�ء�")
end sub

'ģʽ2����
sub update2()
dim trs
dim ForNum,TopNum
Dim orderby,PostUserID,OutTableName,InTableName
OutTableName = Dvbbs.CheckStr(request.form("outtablename"))
InTableName = Dvbbs.CheckStr(request.form("intablename"))
if OutTableName=InTableName then
	Errmsg=ErrMsg + "<BR><li>��������ͬ���ݱ���ת�����ݡ�"
	dvbbs_error()
	exit sub
end if
if (not isnumeric(request.form("selnum"))) or request.form("selnum")="" then
	Errmsg=ErrMsg + "<BR><li>����д��ȷ�ĸ���������"
	dvbbs_error()
	exit sub
end if
if request.form("TopOrEnd")="Top" then	'shinzeal edit this to TopOrEnd in 2004/7/4
orderby=""
else
orderby=" desc"
end if
TopNum=Clng(request.form("selnum"))
if TopNum>100 then
	ForNum=int(TopNum/100)+1
	TopNum=100
else
	ForNum=1
end if

Dim C1
C1=TopNum
%>
&nbsp;<BR>
<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
<tr><td colspan=2 class=td1>
���濪ʼת����̳�������ϣ�Ԥ�Ʊ��ι���<%=C1%>��������Ҫ����
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

dim myrs,maxannid,i,rs
for i=1 to ForNum
set rs=Dvbbs.Execute("select top "&TopNum&" topicid,title from dv_topic where PostTable='"&OutTableName&"' order by topicid "&orderby&"")
if rs.eof and rs.bof then
	Errmsg=ErrMsg + "<BR><li>����ѡ�񵼳������ݱ��Ѿ�û���κ�����"
	dvbbs_error()
	exit sub
else
	do while not rs.eof
		'��ȡ�����������ݱ�
		set trs=Dvbbs.Execute("select * from "&OutTableName&" where rootid="&rs("topicid")&" order by Announceid")
		if not (trs.eof and trs.bof) then
		do while not trs.eof
		'���뵼���������ݱ�
		If IsNull(trs("postuserid")) Or trs("postuserid")="" Then
			PostUserID=0
		Else
			PostUserID=trs("postuserid")
		End If
		Dvbbs.Execute("insert into "&InTableName&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID,isagree,isupload,isaudit,PostBuyUser,UbbList,GetMoney,UseTools,GetMoneyType,FlashId) values ("&trs("boardid")&","&trs("parentid")&",'"&Dvbbs.checkstr(trs("username"))&"','"&Dvbbs.checkstr(trs("topic"))&"','"&Dvbbs.checkstr(trs("body"))&"','"&trs("dateandtime")&"',"&trs("length")&","&trs("rootid")&","&trs("layer")&","&trs("orders")&",'"&trs("ip")&"','"&trs("Expression")&"',"&trs("locktopic")&","&trs("signflag")&","&trs("emailflag")&","&trs("isbest")&","&PostUserID&",'"&trs("isagree")&"',"&trs("isupload")&","&trs("isaudit")&",'"&trs("PostBuyUser")&"','"&Dvbbs.checkstr(trs("UbbList"))&"',"&trs("GetMoney")&",'"&Dvbbs.checkstr(trs("UseTools"))&"',"&trs("GetMoneyType")&",'"&Dvbbs.checkstr(trs("FlashId"))&"')")
		'���¾���'�����ϴ�	'shinzeal add this in 2004/7/4
		If ( Not IsNull(Trs("isbest")) And Trs("isbest")<>"" ) Or ( Not IsNull(Trs("isupload")) And Trs("isupload")<>"" ) Then
			If Trs("isbest")=1 Or Trs("isupload")=1 Then
				Set myrs=Dvbbs.Execute("select max(announceid) from "&Request.Form("intablename")&" where boardid="&Trs("BoardID"))
				maxannid=myrs(0)
				myrs.close
				Set myrs=Nothing
				If Trs("isbest")=1 Then	Dvbbs.Execute("update dv_besttopic set AnnounceID="&maxannid&" where rootid="&rs("topicid"))
				If Trs("isupload")=1 Then Dvbbs.Execute("update Dv_Upfile set F_AnnounceID='" & Rs("TopicID") & "|" & maxannid & "' where F_AnnounceID='" & Rs("TopicID") & "|" & Trs("AnnounceID") & "'")
			End If
		End If
		trs.movenext
		loop
		end if
		'ɾ�������������ݱ���Ӧ����
		Dvbbs.Execute("delete from "&OutTableName&" where RootID="&rs("TopicID"))
		'��������ָ�����ӱ�
		Dvbbs.Execute("update dv_topic set PostTable='"&InTableName&"' where TopicID="&rs("topicid"))
		i=i+1
		'If (i mod 100) = 0 Then
		Response.Write "<script>img2.width=" & Fix((i/C1) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""������"&Server.HtmlEncode(rs(1))&"�����ݣ����ڸ�����һ���������ݣ�" & FormatNumber(i/C1*100,4,-1) & """;" & VbCrLf
		Response.Write "img2.title=""" & Server.HtmlEncode(Rs(1)) & "(" & i & ")"";</script>" & VbCrLf
		Response.Flush
		'End If
	rs.movenext
	loop
end if
next
set trs=nothing
set rs=nothing
Response.Write "<script>img2.width=400;txt2.innerHTML=""100"";</script>"
dv_suc("ת�����ݸ��³ɹ���")
end sub

sub search()
dim keyword
dim totalrec
dim n,rs,sql
dim currentpage,page_count,Pcount,PostUserID
currentPage=request("page")
if currentpage="" or not IsNumeric(currentpage) then
	currentpage=1
else
	currentpage=clng(currentpage)
end if
if request("keyword")="" then
	Errmsg=ErrMsg + "<BR><li>��������Ҫ��ѯ�Ĺؼ��֡�"
	dvbbs_error()
	exit sub
else
	keyword=replace(request("keyword"),"'","")
end if
if request("searchWhat")="username" then
Set Rs=Dvbbs.Execute("Select UserID From Dv_User Where UserName='"&keyword&"'")
If Rs.Eof And Rs.Bof Then
	Errmsg=ErrMsg + "<BR><li>Ŀ���û��������ڣ����������롣"
	dvbbs_error()
	exit sub
Else
	PostUserID=Rs(0)
End If
sql="select * from dv_topic where PostTable='"&Dvbbs.CheckStr(request("tablename"))&"' and PostUserID="&PostUserID&" order by LastPostTime desc"
elseif request("topic")="yes" then
sql="select * from dv_topic where PostTable='"&Dvbbs.CheckStr(request("tablename"))&"' and title like '%"&keyword&"%' order by LastPostTime desc"
else
	Errmsg=ErrMsg + "<BR><li>��ѡ������ѯ�ķ�ʽ��"
	dvbbs_error()
	exit sub
end if
%>
<form method="POST" action="?action=update3">
<!--<input type=hidden name="topic" value="<%=request("topic")%>">
<input type=hidden name="username" value="<%=request("username")%>"> shinzeal add searchWhat in 2004/7/4-->
<input type=hidden name="searchWhat" value="<%=request("searchWhat")%>">

<input type=hidden name="keyword" value="<%=keyword%>">
<input type=hidden name="tablename" value="<%=request("tablename")%>">
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr>
<td height="23" colspan="6" class=td1><B>˵��</B>��<BR>�����Զ����е������������ת�����ݱ��Ĳ�������������ͬ���ڽ���ת��������</td>
</tr>
<tr>
<th colspan="6">����<%=request("tablename")%>���</th>
</tr>
<tr>
<td width="6%" align=center><b>״̬<B></td>
<td width="45%" align=center><B>����</B></td>
<td width="15%" align=center><B>����</B></td>
<td width="6%" align=center><B>�ظ�</B></td>
<td width="22%" align=center><B>ʱ��</B></td>
<td width="6%" align=center><B>����</B></td>
</tr>
<%
set rs=Dvbbs.iCreateObject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
	response.write "<tr> <td class=td1 colspan=6 height=25>û��������������ݡ�</td></tr>"
else
	rs.PageSize = Dvbbs.Forum_Setting(11)
	rs.AbsolutePage=currentpage
	page_count=0
	totalrec=rs.recordcount
	while (not rs.eof) and (not page_count = rs.PageSize)
%>
<tr>
<td width="6%" class=td1 align=center>
<%
if rs("locktopic")=1 then
	response.write "����"
elseif rs("isvote")=1 then
	response.write "ͶƱ"
elseif rs("isbest")=1 then
	response.write "����"
else
	response.write "����"
end if
%>
</td>
<td width="45%" class=td1><%=dvbbs.htmlencode(rs("title"))%></td>
<td width="15%" class=td1 align=center><a href="user.asp?action=modify&userid=<%=rs("postuserid")%>"><%=dvbbs.htmlencode(rs("postusername"))%></a></td>
<td width="6%" class=td1 align=center><%=rs("child")%></td>
<td width="22%" class=td1><%=rs("dateandtime")%></td>
<td width="6%" class=td1 align=center><input type="checkbox" class="checkbox" name="topicid" value="<%=rs("topicid")%>"></td>
</tr>
<%
  	page_count = page_count + 1
	rs.movenext
	wend
	dim endpage
	Pcount=rs.PageCount
	response.write "<tr><td valign=middle nowrap colspan=2 class=td1 height=25>&nbsp;&nbsp;��ҳ�� "

	if currentpage > 4 then
	response.write "<a href=""?page=1&action=search&keyword="&keyword&"&searchWhat="&Request("searchWhat")&"&tablename="&request("tablename")&""">[1]</a> ..."	'shinzeal add searchWhat="&Request("searchWhat")&" in 2004/7/4
	end if
	if Pcount>currentpage+3 then
	endpage=currentpage+3
	else
	endpage=Pcount
	end if
	Dim i
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color="&Dvbbs.mainsetting(1)&">["&i&"]</font>"
		else
        response.write " <a href=""?page="&i&"&action=search&keyword="&keyword&"&searchWhat="&Request("searchWhat")&"&tablename="&request("tablename")&""">["&i&"]</a>"		'shinzeal add searchWhat="&Request("searchWhat")&" in 2004/7/4
		end if
	end if
	next
	if currentpage+3 < Pcount then
	response.write "... <a href=""?page="&Pcount&"&action=search&keyword="&keyword&"&searchWhat="&Request("searchWhat")&"&tablename="&request("tablename")&""">["&Pcount&"]</a>"	'shinzeal add searchWhat="&Request("searchWhat")&" in 2004/7/4
	end if
	response.write "</td>"
	response.write "<td colspan=3 class=td1>���в�ѯ���<input type=checkbox class=checkbox name=allsearch value=yes>"
	response.write "&nbsp;<select name=toTablename>"

	for i=0 to ubound(AllPostTable)
		response.write "<option value="""&AllPostTable(i)&""">"&AllPostTableName(i)& "--" &AllPostTable(i)&"</option>"
	next

	response.write "</select>&nbsp;<input type=submit class=button name=submit value=ת��>"
	response.write "</td>"
	response.write "<td class=td1 align=center><input type=checkbox class=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">"
	response.write "</td></tr>"
end if
rs.close
set rs=nothing
response.write "</table></form><BR><BR>"
end sub

'���������������
sub update3()
dim keyword,trs,PostUserID,TableName,TopicIdStr
Dim Rs,SQL,i
TableName = Dvbbs.CheckStr(request("tablename"))

if request.form("tablename")=request.form("totablename") then
	Errmsg=ErrMsg + "<BR><li>��������ͬ���ݱ��ڽ�������ת����"
	dvbbs_error()
	exit sub
end if
if request.form("allsearch")="yes" then
	if request("keyword")="" then
		Errmsg=ErrMsg + "<BR><li>��������Ҫ��ѯ�Ĺؼ��֡�"
		dvbbs_error()
		exit sub
	else
		keyword=replace(request("keyword"),"'","")
	end if
	if request("searchWhat")="username" then	'shinzeal add searchWhat in 2004/7/4
		Set Rs=Dvbbs.Execute("Select UserID From Dv_User Where UserName='"&keyword&"'")
		If Rs.Eof And Rs.Bof Then
			Errmsg=ErrMsg + "<BR><li>Ŀ���û��������ڣ����������롣"
			dvbbs_error()
			exit sub
		Else
			PostUserID=Rs(0)
		End If
		sql="select topicid,title from dv_topic where PostTable='"&TableName&"' and PostUserID="&PostUserID&" order by LastPostTime desc"
	elseif request("topic")="yes" then
		sql="select topicid,title from dv_topic where PostTable='"&TableName&"' and title like '%"&keyword&"%' order by LastPostTime desc"
	else
		Errmsg=ErrMsg + "<BR><li>��ѡ������ѯ�ķ�ʽ��"
		dvbbs_error()
		exit sub
	end if
else
	if request.form("topicid")="" then
		Errmsg=ErrMsg + "<BR><li>��ѡ��Ҫת�Ƶ����ӡ�"
		dvbbs_error()
		exit sub
	end if

	For i = 1 To request.form("TopicID").Count
		If isNumeric(request.form("TopicID")(i)) Then
			If TopicIdStr = "" Then
				TopicIdStr = request.form("TopicID")(i)
			Else
				TopicIdStr = TopicIdStr & ","& request.form("TopicID")(i)
			End If
		End If
	Next
	sql="select topicid,title from dv_topic where PostTable='"&TableName&"' and TopicID in ("&TopicIdStr&")"
end if

'set rs=Dvbbs.Execute(sql)
Set Rs=Dvbbs.iCreateObject("adodb.recordset")
Rs.Open SQL,Conn,1,1
Dim C1,myrs,maxannid
C1=Rs.ReCordCount
%>
&nbsp;<BR>
<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
<tr><td colspan=2 class=td1>
���濪ʼת����̳�������ϣ�Ԥ�Ʊ��ι���<%=C1%>��������Ҫ����
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
if rs.eof and rs.bof then
	Errmsg=ErrMsg + "<BR><li>û���κμ�¼��ת����"
	dvbbs_error()
	exit sub
else
	do while not rs.eof
	'ȡ��ԭ������
	set trs=Dvbbs.Execute("select * from "&TableName&" where rootid="&rs("topicid")&" order by Announceid")
	if not (trs.eof and trs.bof) then
	'�����±�
	do while not trs.eof
		If IsNull(trs("postuserid")) Or trs("postuserid")="" Then
			PostUserID=0
		Else
			PostUserID=trs("postuserid")
		End If
	Dvbbs.Execute("insert into "&Dvbbs.CheckStr(request("totablename"))&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID,isagree,isupload,isaudit,PostBuyUser,UbbList,GetMoney,UseTools,GetMoneyType,FlashId) values ("&trs("boardid")&","&trs("parentid")&",'"&Dvbbs.checkstr(trs("username"))&"','"&Dvbbs.checkstr(trs("topic"))&"','"&Dvbbs.checkstr(trs("body"))&"','"&trs("dateandtime")&"',"&trs("length")&","&trs("rootid")&","&trs("layer")&","&trs("orders")&",'"&trs("ip")&"','"&trs("Expression")&"',"&trs("locktopic")&","&trs("signflag")&","&trs("emailflag")&","&trs("isbest")&","&PostUserID&",'"&trs("isagree")&"',"&trs("isupload")&","&trs("isaudit")&",'"&trs("PostBuyUser")&"','"&Dvbbs.checkstr(trs("UbbList"))&"',"&trs("GetMoney")&",'"&Dvbbs.checkstr(trs("UseTools"))&"',"&trs("GetMoneyType")&",'"&Dvbbs.checkstr(trs("FlashId"))&"')")
	'���¾���'�����ϴ�	'shinzeal add this in 2004/7/4
	If ( Not IsNull(Trs("isbest")) And Trs("isbest")<>"" ) Or ( Not IsNull(Trs("isupload")) And Trs("isupload")<>"" ) Then
		If Trs("isbest")=1 Or Trs("isupload")=1 Then
			Set myrs=Dvbbs.Execute("select max(announceid) from "&Request.Form("totablename")&" where boardid="&Trs("BoardID"))
			maxannid=myrs(0)
			myrs.close
			Set myrs=Nothing
			If Trs("isbest")=1 Then	Dvbbs.Execute("update dv_besttopic set AnnounceID="&maxannid&" where rootid="&Rs("TopicID"))
			If Trs("isupload")=1 Then Dvbbs.Execute("update Dv_Upfile set F_AnnounceID='" & Rs("TopicID") & "|" & maxannid & "' where F_AnnounceID='" & Rs("TopicID") & "|" & Trs("AnnounceID") & "'")
		End If
	end if
	trs.movenext
	loop
	end if
	'ɾ��ԭ������������
	Dvbbs.Execute("delete from "&TableName&" where rootid="&rs("topicid"))
	'���¸��������
	Dvbbs.Execute("update dv_topic set PostTable='"&Dvbbs.CheckStr(request("totablename"))&"' where topicid="&rs("topicid"))
		i=i+1
		'If (i mod 100) = 0 Then
		Response.Write "<script>img2.width=" & Fix((i/C1) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""������"&Server.HtmlEncode(rs(1))&"�����ݣ����ڸ�����һ���������ݣ�" & FormatNumber(i/C1*100,4,-1) & """;" & VbCrLf
		Response.Write "img2.title=""" & Server.HtmlEncode(Rs(1)) & "(" & i & ")"";</script>" & VbCrLf
		Response.Flush
		'End If
	rs.movenext
	loop
end if
set trs=nothing
set rs=nothing
Response.Write "<script>img2.width=400;txt2.innerHTML=""100"";</script>"
dv_suc("ת�����ݸ��³ɹ���")
end sub

Function TableSelectForm()
	Dim i,Rs
	TableSelectForm = ""
	For i=0 To UBound(AllPostTable)
		On Error Resume Next
		Set Rs=Dvbbs.Execute("Select Top 1 * From "&AllPostTable(i))
		If Err Then
			Err.Clear
		Else
			TableSelectForm = TableSelectForm & "<option value="""&AllPostTable(i)&""">"&AllPostTableName(i)& "--" &AllPostTable(i)&"</option>"
		End If
	next
End Function
%>