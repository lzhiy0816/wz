<!--#include file="../conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="../inc/dv_clsother.asp" -->
<!-- #include file="../inc/GroupPermission.asp" -->
<!--#include file=../inc/md5.asp-->
<%
	Head()
	Server.ScriptTimeout=999999
	dim Str
	dim admin_flag
	admin_flag="9,10"
	founderr=False 
	if not Dvbbs.master or instr(","&session("flag")&",",",9,")=0 or instr(","&session("flag")&",",",10,")=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=../admin_login.asp target=_top>��¼</a>����롣<br><li>��û�й�����ҳ���Ȩ�ޡ�"
		dvbbs_error()
	else
		call main()
		footer()
	end if

	sub main()
%>
<table width="95%" border="0" cellspacing="0" cellpadding="0"  align=center class="tableBorder">
<tr> 
<th width="100%" class="tableHeaderText" colspan=2 height=25>��̳����
</th>
</tr>
<tr>
<td class="forumRowHighlight" colspan=2>
<p><B>ע��</B>��<BR>��ɾ����̳ͬʱ��ɾ������̳���������ӣ�ɾ������ͬʱɾ��������̳���������ӣ� ����ʱ��������д������Ϣ��<BR>�����ѡ��<B>��λ���а���</B>�������а��涼����Ϊһ����̳�����ࣩ����ʱ����Ҫ���¶Ը���������й����Ļ������ã�<B>��Ҫ����ʹ�øù���</B>�����������˴�������ö��޷���ԭ����֮��Ĺ�ϵ�������ʱ��ʹ�ã���������Ҳ����ֻ���ĳ��������и�λ����(������ĸ�����������˵�)�������뿴����˵��<BR><font color=blue>ÿ������ĸ��������������˵�������ǰ����ϸ�Ķ�˵�������������˵��бȱ�İ��������˷�������ͷ��ิλ����</font><BR>
<font color=red>�����ϣ��ĳ��������Ҫ��Ա����һ�����ۣ����ң����ܽ��룬�����ڰ���߼�������������Ӧ�����������Ľ�һ��ȯ���Լ��ܷ��ʵ�ʱ���Ƕ���</font>
</td>
</tr>
<tr>
<td class="forumRowHighlight" height=25>
<B>��̳����ѡ��</B></td>
<td class="forumRowHighlight"><a href="board.asp">��̳������ҳ</a> | <a href="board.asp?action=add">�½���̳����</a> | <a href="?action=settemplates">ģ������������</a> | <a href="?action=orders">һ����������</a> | <a href="?action=boardorders">N����������</a> | <a href="?action=RestoreBoard" onclick="{if(confirm('��λ���а��潫�����а���ָ���Ϊһ������࣬��λ��Ҫ�����а������½��й����Ļ������ã������ز�����ȷ����λ��?')){return true;}return false;}">��λ���а���</a> | <a href="?action=RestoreBoardCache" onclick="{if(confirm('��ʱ��������̳������޸���ǰ̨�������޸�Ч������ܿ�������Ӧ����Ļ���û����Ч���£������ｫ�ؽ����а���Ļ��棬������İ���ܶ࣬�⽫������һ����ʱ�䣬ȷ����?')){return true;}return false;}">�ؽ����滺��</a>
</td>
</tr>
</table>
<p></p>
<%
select case Request("action")
case "add"
	call add()
case "edit"
	call edit()
case "savenew"'������̳
	call savenew()
Case "savedit"
	call savedit()
Case "del"
	call del()
Case "orders"
	call orders()
case "updatorders"
	call updateorders()
Case "boardorders"
	call boardorders()
case "updatboardorders"
	call updateboardorders()
Case "mode"
	call mode()
case "savemod"
	call savemod()
Case "permission"
	call boardpermission()
case "editpermission"
	call editpermission()
case "RestoreBoard"
	call RestoreBoard()
Case "RestoreBoardCache"
	Call RestoreBoardCache()
Case "clearDate"
	Call clearDate
Case "delDate"
	Call delDate
Case "RestoreClass"
	Call RestoreClass
Case "handorders"
	Call handorders
Case "savehandorders"
	Call savehandorders
Case "savesid"
	Call savesid
Case "upallsid"
	Call upallsid
Case "settemplates"
	Call Settemplates
Case Else
	Call boardinfo()
end select
end Sub
Sub upallsid()
	Dim Sid,cid,board
	SID= Request("Sid")
	cid=Request("CID")
	If SID="" Then SID=1
	If CID="" Then CID=1
	SID=CLng(SID)
	CID=CLng(CID)
	Dvbbs.Execute("Update Dv_Setup Set Forum_Sid="& SID &",Forum_CID="& CID )
	Dvbbs.Execute("Update Dv_Board Set Sid="& SID &",CID="& CID )
	Dvbbs.loadSetup()
	For Each board in Application(Dvbbs.CacheName&"_boardlist").documentElement.selectNodes("board/@boardid")
	Dvbbs.LoadBoardData(board.text)
	Next
	Dv_suc("��̳��������ʽͳһ���óɹ�!")
End Sub
Sub savesid
	Dim i,boardid,TempStr
	Dim Templateslist,sid,j,bid,cid
	sid=""
	For Each TempStr in Request.form("upboardid")
		If Bid="" Then
			Bid=TempStr
		Else
			Bid=Bid&","&TempStr
		End If 
	Next
	Bid=split(Bid,",")
	For i=0 to UBound(bid)
		If sid="" Then
			sid=Request("sid"&bid(i))
		Else
			sid=sid&","&Request("sid"&bid(i))
		End If
		If cid="" Then
			cid=Request("cid"&bid(i))
		Else
			cid=cid&","&Request("cid"&bid(i))
		End If
	Next
	sid=split(sid,",")
	cid=split(cid,",")
	For i=0 to UBound(bid)	
		Dvbbs.Execute("Update Dv_board set sid="&CLng(sid(i))&" ,cid="&CLng(cid(i))&" Where BoardId="&bid(i)&" ")
		Dvbbs.LoadBoardData(bid(i))
	Next 
	Dv_suc("��̳ģ���������óɹ�!")
End Sub

Sub Settemplates
'Application(Dvbbs.CacheName &"_style")
Dim reBoard_Setting,MoreMenu,i
Dim Templateslist

%>
<form action ="board.asp?action=upallsid" method=post name="dv">
<table cellspacing="0" cellpadding="0" align=center Class="tableBorder" style="width:98%" >
<tr> 
<th colspan="2" class="tableHeaderText" align=center height=25>ģ �� ͳ һ �� ��
</th>
</tr>
<tr>
<td width="70%" align=Left  class="forumRowHighlight" ><B>������̳������ģ������Ϊ��</b>&nbsp; 
<%
	Dim forum_sid,iCssName,iCssID,iStyleName
	Dim Forum_cid
	set rs=dvbbs.execute("select forum_sid,Forum_CID from dv_setup")
	Forum_sid=rs(0)
	Forum_CID=Rs(1)
	Rs.close:Set Rs=Nothing
%>
<Select Size=1 Name="sid">
<%
For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
	Response.Write "<Option value="""& Templateslist.selectSingleNode("@id").text &""""
	If Forum_sid = CLng(Templateslist.selectSingleNode("@id").text) Then Response.Write " selected "
	Response.Write ">"& Templateslist.selectSingleNode("@stylename").text &" </Option>"
Next
%>
</Select>
Css��ʽ:<Select Size=1 Name="cid">
<%
For Each Templateslist in Application(Dvbbs.CacheName & "_csslist").documentElement.selectNodes("css[tid="& Forum_sid &"]")
	Response.Write "<Option value="""& Templateslist.selectSingleNode("@id").text &""""
	If Forum_cid = CLng(Templateslist.selectSingleNode("@id").text)  Then Response.Write " selected "
	Response.Write ">"& Templateslist.selectSingleNode("@type").text &" </Option>"
Next
%>
</Select>
</td>
<td width="30%" align=Left  class="forumRowHighlight" >
<Input type="submit" name="Submit" value="�� ��"></td>
</tr>
</table><BR>
</form>
<form action ="board.asp?action=savesid" method=post name="dv1">
<table cellspacing="0" cellpadding="0" align=center Class="tableBorder" style="width:98%" >
<tr> 
<th width="70%" class="tableHeaderText" align=Left height=25>&nbsp;��̳����
</th>
<th width="30%" class="tableHeaderText" align=Left height=25>���÷����ʽ
</th>
</tr>
<%
dim classrow
sql="select * from dv_board order by rootid,orders"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof
reBoard_Setting=split(rs("Board_setting"),",")
if classrow="forumRowHighlight" then
	classrow="forumRow"
else
	classrow="forumRowHighlight"
end if
%>
<tr> 
<td height="25"  class="<%=classrow%>">
<%if rs("depth")>0 then%>
<%for i=1 to rs("depth")%>
&nbsp;
<%next%>
<%end if%>
<%if rs("child")>0 then%><img src="../skins/default/plus.gif"><%else%><img src="../skins/default/nofollow.gif"><%end if%>
<%if rs("parentid")=0 then%><b><%end if%><%=rs("boardtype")%><%if rs("child")>0 then%>(<%=rs("child")%>)<%end if%>
<%if rs("parentid")=0 then%></b><%end if%>
</td>
<td align=Left  class="<%=classrow%>">
<Select Size=1 Name="sid<%=Rs("BoardID")%>">
<%
For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
	Response.Write "<Option value="""& Templateslist.selectSingleNode("@id").text &""""
	If Rs("SID") = CLng(Templateslist.selectSingleNode("@id").text) Then Response.Write " selected "
	Response.Write ">"& Templateslist.selectSingleNode("@stylename").text &" </Option>"
Next
%>
</Select>Css��ʽ:<Select Size=1 Name="Cid<%=Rs("BoardID")%>">
<%
For Each Templateslist in Application(Dvbbs.CacheName & "_csslist").documentElement.selectNodes("css[tid="& Rs("SID") &"]")
	Response.Write "<Option value="""& Templateslist.selectSingleNode("@id").text &""""
	If Rs("cid") = CLng(Templateslist.selectSingleNode("@id").text)  Then Response.Write " selected "
	Response.Write ">"& Templateslist.selectSingleNode("@type").text &" </Option>"
Next
%>
</Select>
<Input type="hidden" name="upboardid" value="<%=rs("boardid")%>">
</td></tr>
<%
Rs.movenext
loop
set rs=nothing
%>
<tr>
<td width=300 align=Left  class="forumRowHighlight" ></td>
<td width=300 align=Left  class="forumRowHighlight" ><input type="submit" name="Submit" value="�� ��"></td>
</tr>
</table><BR><BR>
</form>

<%
End Sub 
sub boardinfo()
Dim reBoard_Setting,MoreMenu
Dim classrow,iii
%>
<table width="95%" cellspacing="0" cellpadding="0" align=center class="tableBorder">
<tr> 
<th width="35%" class="tableHeaderText" height=25>��̳����
</th>
<th width="35%" class="tableHeaderText" height=25>����
</th>
</tr>
<%
SQL="select boardid,boardtype,parentid,depth,child,Board_setting from dv_board order by rootid,orders"
SET Rs = Conn.Execute(SQL)
If Rs.eof Then
	Rs.close:Set Rs = Nothing
Else

SQL=Rs.GetRows(-1)
Rs.close:Set Rs = Nothing
For iii=0 To Ubound(SQL,2)
	reBoard_Setting=split(SQL(5,iii),",")
	if classrow="forumRowHighlight" then
		classrow="forumRow"
	else
		classrow="forumRowHighlight"
	end if
	Response.Write "<tr>"
	Response.Write "<td height=""25"" width=""55%""  class="
	Response.Write classrow 
	Response.Write ">"
	if SQL(3,iii)>0 then
		for i=1 to SQL(3,iii)
			Response.Write "&nbsp;&nbsp;"
		next
	end if
	if SQL(4,iii)>0 then
		Response.Write "<img src=""../skins/default/plus.gif"">"
	else
		Response.Write "<img src=""../skins/default/nofollow.gif"">"
	end if
	if SQL(2,iii)=0 then
		Response.Write "<b>"
	end if
	Response.Write SQL(1,iii)
	if SQL(4,iii)>0 then
		Response.Write "("
		Response.Write SQL(4,iii)
		Response.Write ")"
	end if
%>
</td>
<td width="45%" class="<%=classrow%>">
<a href="board.asp?action=add&editid=<%=SQL(0,iii)%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>���Ӱ���</U></font></a> | <a href="board.asp?action=edit&editid=<%=SQL(0,iii)%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>��������</U></font></a> | <a href="BoardSetting.asp?editid=<%=SQL(0,iii)%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>�߼�����</U></font></a>
<%
'if reBoard_Setting(2)=0 then
'	MoreMenu=MoreMenu & "<div class=menuitems><a href=vipboard.asp?boardid="&SQL(0,iii)&"><font color="&Dvbbs.mainsetting(3)&"><U>VIP��̳����</U></font></a></div>"
'elseif reBoard_Setting(2)=0 and reBoard_Setting(46)>0 then
'	MoreMenu=MoreMenu & "<div class=menuitems><a href=vipboard.asp?boardid="&SQL(0,iii)&"&action=reinstall><font color="&Dvbbs.mainsetting(3)&"><U>����VIP��̳</U></font></a></div>"
'elseif reBoard_Setting(2)=1 and reBoard_Setting(46)>0 then
'	MoreMenu=MoreMenu & "<div class=menuitems><a href=vipboard.asp?action=showvipuser&boardid="&SQL(0,iii)&"><font color="&Dvbbs.mainsetting(3)&"><U>�鿴VIP�û�</U></font></a></div>"
'end if

MoreMenu=MoreMenu & "<div class=menuitems><a href=update.asp?action=updat&submit=������̳����&boardid="&SQL(0,iii)&" title=�������ظ������������ظ���><font color="&Dvbbs.mainsetting(3)&"><U>��������</U></font></a></div><div class=menuitems><a href=# onclick=alertreadme(\'��ս���������̳�����������ڻ���վ��ȷ�������?\',\'update.asp?action=delboard&boardid="&SQL(0,iii)&"\')><font color="&Dvbbs.mainsetting(3)&"><U>��հ�������</U></font></a></div>"

if SQL(4,iii)=0 then
MoreMenu=MoreMenu & "<div class=menuitems><a href=# onclick=alertreadme(\'ɾ������������̳���������ӣ�ȷ��ɾ����?\',\'board.asp?action=del&editid="&SQL(0,iii)&"\')><font color="&Dvbbs.mainsetting(3)&"><U>ɾ������</U></font></a></div>"
else
MoreMenu=MoreMenu & "<div class=menuitems><a href=# onclick=alertreadme(\'����̳����������̳��������ɾ����������̳����ɾ������̳��\',\'#\')><font color="&Dvbbs.mainsetting(3)&"><U>ɾ������</U></font></a></div>"
end if
MoreMenu=MoreMenu & "<div class=menuitems><a href=Board.asp?action=clearDate&boardid="&SQL(0,iii)&"><font color="&Dvbbs.mainsetting(3)&"><u>��������</u></font></a></div>"
If SQL(2,iii)=0 Then
	MoreMenu=MoreMenu & "<div class=menuitems><a href=# onclick=alertreadme(\'��λ�÷��ཫ��Ѹ÷����µ����а��涼��λ�ɶ������棬����ԭ���Ķ༶���඼����λ�ɶ������棬�����ز�����ȷ����λ��?\',\'?action=RestoreClass&classid="&SQL(0,iii)&"\')><font color="&Dvbbs.mainsetting(3)&"><u>��λ�÷���</u></font></a></div><div class=menuitems><a href=?action=handorders&classid="&SQL(0,iii)&"><font color="&Dvbbs.mainsetting(3)&"><u>��������(�ֶ�)</u></font></a></div>"
End If
%>
 | <a href="#" onMouseOver="showmenu(event,'<%=MoreMenu%>')" style="CURSOR:hand"><font color=<%=Dvbbs.mainsetting(3)%>><u>�������</u></font></a>
<%
if reBoard_Setting(2)=1 then
	Response.Write "<a href=board.asp?action=mode&boardid="&SQL(0,iii)&"><font color="&Dvbbs.mainsetting(3)&"><U>��֤�û�</U></font></a>"
end if
%>
</td></tr>
<%
MoreMenu=""
Next
End If
%>
</table><BR><BR>
<SCRIPT LANGUAGE="JavaScript">
<!--
function alertreadme(str,url){
{if(confirm(str)){
location.href=url;
return true;
}return false;}
}
//-->
</SCRIPT>
<%
end sub

sub add()
dim rs_c
Dim forum_sid,forum_cid,Style_Option,TempOption
Dim iCssName,iCssID,iStyleName
set rs_c= server.CreateObject ("adodb.recordset")
sql = "select * from dv_board order by rootid,orders"
rs_c.open sql,conn,1,1
	dim boardnum
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select Max(boardid) from dv_board"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	boardnum=1
	else
	boardnum=rs(0)+1
	end if
	if isnull(boardnum) then boardnum=1
	if boardnum=444 then boardnum=445
	if boardnum=777 then boardnum=778
	rs.close
%>
<form action ="board.asp?action=savenew" method=post name=theform>
<input type="hidden" name="newboardid" value=<%=boardnum%>>
<table width="85%" border="0" cellspacing="1" cellpadding="0"  align=center class="tableBorder">
<tr> 
<th height=24 colspan=2><B>��������̳</th>
</tr>
<tr> 
<td width="100%" height=30 class="forumrowHighLight" colspan=2>
˵����<BR>1��������̳�������ص����þ�ΪĬ�����ã��뷵����̳���������ҳ�����б��ĸ߼����������ø���̳����Ӧ���ԣ��������Ը���̳���������Ȩ�����ã��뵽<A HREF="board.asp?action=permission"><font color=blue>��̳Ȩ�޹���</font></A>��������Ӧ�û����ڸð����Ȩ�ޡ�<BR>
2��<font color=blue>��������ӵ�����̳����</font>��ֻ��Ҫ������������ѡ����Ϊ��̳���༴�ɣ�<font color=blue>��������ӵ�����̳����</font>����Ҫ������������ȷ����ѡ�����̳������ϼ�����
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow">��̳����</td>
<td width="60%" class="forumrow"> 
<input type="text" name="boardtype" size="35">
</td>
</tr>
<tr> 
<td width="40%" height=24 class="forumrow">����˵��<BR>����ʹ��HTML����</td>
<td width="60%" class="forumrow"> 
<textarea name="Readme" cols="50" rows="5"></textarea>
</td>
</tr>
<tr> 
<td width="40%" height=24 class="forumrow">�������<BR>����ʹ��HTML����</td>
<td width="60%" class="forumrow"> 
<textarea name="Rules" cols="50" rows="5"></textarea>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow"><U>�������</U></td>
<td width="60%" class="forumrow"> 
<select name=class>
<option value="0">��Ϊ��̳����</option>
<% do while not rs_c.EOF%>
<option value="<%=rs_c("boardid")%>" <%if request("editid")<>"" and clng(request("editid"))=rs_c("boardid") then%>selected<%end if%>>
<%if rs_c("depth")>0 then%>
<%for i=1 to rs_c("depth")%>

<%next%>
<%end if%><%=rs_c("boardtype")%></option>
<%
rs_c.MoveNext 
loop
rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow"><U>ʹ����ʽ���</U><BR>�����ʽ����а�����̳��ɫ��ͼƬ<BR>����Ϣ</td>
<td width="60%" class="forumrow">
<%
	set rs_c=dvbbs.execute("select forum_sid,forum_cid from dv_setup")
	Forum_sid=rs_c(0)
	Forum_cid=rs_c(1)
	rs_c.close:Set rs_c=Nothing
%>
<Select Size=1 Name="sid">
<%
Dim Templateslist
For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
	Response.Write "<Option value="""& Templateslist.selectSingleNode("@id").text &""""
	If Forum_sid = CLng(Templateslist.selectSingleNode("@id").text) Then Response.Write " selected "
	Response.Write ">"& Templateslist.selectSingleNode("@stylename").text &" </Option>"
Next
%>
</Select>Css��ʽ:<Select Size=1 Name="cid">
<%
For Each Templateslist in Application(Dvbbs.CacheName & "_csslist").documentElement.selectNodes("css")
	Response.Write "<Option value="""& Templateslist.selectSingleNode("@id").text &""""
	If Forum_cid = CLng(Templateslist.selectSingleNode("@id").text)  Then Response.Write " selected "
	Response.Write ">"& Templateslist.selectSingleNode("@type").text &" </Option>"
Next
%>
</Select>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow"><U>��̳����</U><BR>�������������|�ָ����磺ɳ̲С��|wodeail</td>
<td width="60%" class="forumrow"> 
<input type="text" name="boardmaster" size="35">
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow"><U>��ҳ��ʾ��̳ͼƬ</U><BR>��������ҳ��̳����������<BR>��ֱ����дͼƬURL</td>
<td width="60%" class="forumrow">
<input type="text" name="indexIMG" size="35">
</td>
</tr>
<tr> 
<td width="40%" height=24 class="forumRow">&nbsp;</td>
<td width="60%" class="forumRow"> 
<input type="submit" name="Submit" value="������̳">
</td>
</tr>
</table>
</form>
<%
set rs_c=nothing
set rs=nothing
end sub

sub edit()
dim rs_c,reBoard_Setting
Dim forum_sid,forum_cid,Style_Option,TempOption
Dim iCssName,iCssID,iStyleName
sql = "select * from dv_board order by rootid,orders"
set rs_c=Dvbbs.Execute(sql)
sql = "select * from dv_board where boardid="&request("editid")
set rs=Dvbbs.Execute(sql)
reBoard_Setting=split(rs("Board_setting"),",")

forum_sid=rs("sid")
forum_cid=rs("cid")
%>        
<form action ="board.asp?action=savedit" method=post name=theform>       
<input type="hidden" name=editid value="<%=Request("editid")%>">
<table width="85%" border="0" cellspacing="1" cellpadding="0"  align=center class="tableBorder">
<tr> 
<th height=24 colspan=2>�༭��̳��<%=rs("boardtype")%></th>
</tr>
<tr> 
<td width="100%" height=30 class="forumrowHighLight" colspan=2>
˵����<BR>1��������̳�������ص����þ�ΪĬ�����ã��뷵����̳���������ҳ�����б��ĸ߼����������ø���̳����Ӧ���ԣ��������Ը���̳���������Ȩ�����ã��뵽<A HREF="board.asp?action=permission"><font color=blue>��̳Ȩ�޹���</font></A>��������Ӧ�û����ڸð����Ȩ�ޡ�<BR>
2��<font color=blue>��������ӵ�����̳����</font>��ֻ��Ҫ������������ѡ����Ϊ��̳���༴�ɣ�<font color=blue>��������ӵ�����̳����</font>����Ҫ������������ȷ����ѡ�����̳������ϼ�����
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow">��̳����</td>
<td width="60%" class="forumrow"> 
<input type="text" name="boardtype" size="35"  value="<%=Server.htmlencode(rs("boardtype"))%>" >
</td>
</tr>
<tr> 
<td width="40%" height=24 class="forumrow">����˵��<BR>����ʹ��HTML����</td>
<td width="60%" class="forumrow"> 
<textarea name="Readme" cols="50" rows="5"><%=server.HTMLEncode(Rs("readme")&"")%></textarea>
</td>
</tr>
<tr> 
<td width="40%" height=24 class="forumrow">�������<BR>����ʹ��HTML����</td>
<td width="60%" class="forumrow"> 
<textarea name="Rules" cols="50" rows="5"><%=server.HTMLEncode(Rs("Rules")&"")%></textarea>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow"><U>�������</U><BR>������̳����ָ��Ϊ��ǰ����<BR>������̳����ָ��Ϊ��ǰ�����������̳</td>
<td width="60%" class="forumrow"> 
<select name=class>
<option value="0">��Ϊ��̳����</option>
<% do while not rs_c.EOF%>
<option value="<%=rs_c("boardid")%>" <% if cint(rs("parentid")) = rs_c("boardid") then%> selected <%end if%>><%if rs_c("depth")>0 then%>
<%for i=1 to rs_c("depth")%>
��
<%next%>
<%end if%><%=rs_c("boardtype")%></option>
<%
rs_c.MoveNext 
loop
rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow"><U>ʹ����ʽ���</U><BR>�����ʽ����а�����̳��ɫ��ͼƬ<BR>����Ϣ</td>
<td width="60%" class="forumrow">
<%
	set rs_c=dvbbs.execute("select forum_sid,forum_cid from dv_setup")
	Forum_sid=rs_c(0)
	Forum_cid=rs_c(1)
	rs_c.close:Set rs_c=Nothing
%>
<Select Size=1 Name="sid">
<%
Dim Templateslist
For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
	Response.Write "<Option value="""& Templateslist.selectSingleNode("@id").text &""""
	If rs("sid") = CLng(Templateslist.selectSingleNode("@id").text) Then Response.Write " selected "
	Response.Write ">"& Templateslist.selectSingleNode("@stylename").text &" </Option>"
Next
%>
</Select>Css��ʽ:<Select Size=1 Name="cid">
<%
For Each Templateslist in Application(Dvbbs.CacheName & "_csslist").documentElement.selectNodes("css[tid="& rs("sid")&"]")
	Response.Write "<Option value="""& Templateslist.selectSingleNode("@id").text &""""
	If rs("cid") = CLng(Templateslist.selectSingleNode("@id").text)  Then Response.Write " selected "
	Response.Write ">"& Templateslist.selectSingleNode("@type").text &" </Option>"
Next
%>
</Select>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow"><U>��̳����</U><BR>�������������|�ָ����磺ɳ̲С��|wodeail</td>
<td width="60%" class="forumrow"> 
<input type="text" name="boardmaster" size="35"  value='<%=rs("boardmaster")%>'>
<input type="hidden" name="oldboardmaster" value='<%=rs("boardmaster")%>'>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow"><U>��ҳ��ʾ��̳ͼƬ</U><BR>��������ҳ��̳����������<BR>��ֱ����дͼƬURL</td>
<td width="60%" class="forumrow">
<input type="text" name="indexIMG" size="35" value="<%=rs("indexIMG")%>">
</td>
</tr>
<tr> 
<td width="40%" height=24 class="forumrow">&nbsp;</td>
<td width="60%" class="forumrow"> 
<input type="submit" name="Submit" value="�ύ�޸�">
</td>
</tr>
<tr> 
<td width="100%" height=30 class="forumrowHighLight" colspan=2 align=right>
<a href="board.asp?action=add&editid=<%=Request("editid")%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>���Ӱ���</U></font></a> | <a href="board.asp?action=edit&editid=<%=Request("editid")%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>��������</U></font></a> | <a href="BoardSetting.asp?editid=<%=Request("editid")%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>�߼�����</U></font></a>
<%if reBoard_Setting(2)=1 then%>
| <a href="board.asp?action=mode&boardid=<%=Request("editid")%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>��֤�û�</U></font></a>
<%end if%>
| <a href="update.asp?action=updat&submit=������̳����&boardid=<%=Request("editid")%>" title="�������ظ������������ظ���"><font color="<%=Dvbbs.mainsetting(3)%>"><U>��������</U></font></a> | <a href="update.asp?action=delboard&boardid=<%=Request("editid")%>" onclick="{if(confirm('��ս���������̳�����������ڻ���վ��ȷ�������?')){return true;}return false;}"><font color="<%=Dvbbs.mainsetting(3)%>"><U>���</U></font></a> | <%if rs("child")=0 then%><a href="board.asp?action=del&editid=<%=Request("editid")%>" onclick="{if(confirm('ɾ������������̳���������ӣ�ȷ��ɾ����?')){return true;}return false;}"><font color="<%=Dvbbs.mainsetting(3)%>"><U>ɾ��</U></a><%else%><a href="#" onclick="{if(confirm('����̳����������̳��������ɾ����������̳����ɾ������̳��')){return true;}return false;}"><font color="<%=Dvbbs.mainsetting(3)%>"><U>ɾ��</U></a><%end if%>
| <a href="Board.asp?action=clearDate&boardid=<%=Request("editid")%>"> <font color="<%=Dvbbs.mainsetting(3)%>"><u>��������</u></a>
</td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
set rs_c=nothing
end sub

Sub Mode()
	Dim Boarduser
	Dim BoarduserNum
%>
<form action ="board.asp?action=savemod" method=post>
<table width="95%" class="tableBorder" cellspacing="1" cellpadding="1" align="center">
<tr> 
<th width="52%" height=22>˵����</th>
<th width="48%">������</th>
</tr>
<tr> 
<td width="52%" height=22 class=forumrow><B>��̳����</B></td>
<td width="48%" class=forumrow> 
<%
Sql = "SELECT Boardid, Boardtype, Boarduser FROM Dv_Board WHERE Boardid = " & Request("boardid")
Set Rs = Dvbbs.Execute(Sql)
If Rs.Eof And Rs.Bof Then
	Response.Write "�ð��沢�����ڻ��߸ð��治�Ǽ��ܰ��档"
Else
	Response.Write Rs(1)
	Response.Write "<input type=hidden value=" & Rs(0) & " name=boardid>"
	Boarduser = Rs(2)
End If
Set Rs = Nothing
%>
</td>
</tr>
<tr> 
<td width="52%" class=forumrow valign=top><B>��֤�û�</B>��
<%
	If Not Isnull(Boarduser) Or Boarduser <> "" Then
		BoarduserNum = Split(Boarduser,",")
		Response.Write "�����湲��<font color=red>" & Ubound(BoarduserNum)+1 & "</font>λ��֤�û���"
	Else
		Response.Write "��������ʱû����֤�û���"
	End If
%>
<br>
ֻ���趨Ϊ��֤��̳����̳��Ҫ��д�ܹ�����ð�����û���ÿ����һ���û���ȷ���û�������̳�д��ڣ�ÿ���û�����<B>�س�</B>�ֿ�</font>
<%
If Clng(Dvbbs.Board_Setting(62))>0 Or Clng(Dvbbs.Board_Setting(63))>0 Then Response.Write "<BR><font color=blue>�˰���������֧����һ��ȯ���ܽ��룬��Ч��Ϊ<font color=red>" & Clng(Dvbbs.Board_Setting(64)) & "</font>���£�����ÿ���û���������ϣ�=��ǰʱ�䣬ÿ��Ч���磺admin="&Now&"</font>"
%>
</td>
<td width="48%" class=forumrow> 
<textarea cols="50" rows="3" name="vipuser" id="vipuser">
<%if not isnull(boarduser) or boarduser<>"" then
	response.write Replace(boarduser,",",Chr(10))
end if%></textarea>
<br><a href="javascript:admin_Size(-3,'vipuser')"><img src="images/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(3,'vipuser')"><img src="images/plus.gif" unselectable="on" border='0'></a>
</td>
</tr>
<tr> 
<td width="52%" height=22 class=forumrow>&nbsp;</td>
<td width="48%" class=forumrow> 
<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</table>
</form>
<%
End Sub 

'����༭��̳��֤�û���Ϣ
'��ڣ��û��б��ַ���
sub savemod()
	dim boarduser
	dim boarduser_1
	dim userlen
	dim updateinfo
	'����������̳���ڵ���֤�û� 2005-3-10 Dv.Yz
	Dim Get_BoardUser_Money, BoardUser_Money
	Get_BoardUser_Money = False
	If Clng(Dvbbs.Board_Setting(62))>0 Or Clng(Dvbbs.Board_Setting(63))>0 Then Get_BoardUser_Money = True
	
	If trim(request("vipuser"))<>"" then
		boarduser=request("vipuser")
		boarduser=split(boarduser,chr(13)&chr(10))
		For i = 0 To Ubound(Boarduser)
			If Not (Boarduser(i) = "" Or Boarduser(i) = " ") Then
				If Get_BoardUser_Money Then
					BoardUser_Money = Split(Boarduser(i),"=")
					If Not DateDiff("d",BoardUser_Money(1),Now()) > Cint(Dvbbs.Board_Setting(64))*30 Then
						Boarduser_1 = "" & Boarduser_1 & "" & Boarduser(i) & ","
					End If
				Else
					Boarduser_1 = "" & Boarduser_1 & "" & Boarduser(i) & ","
				End If
			End If
		Next
		userlen=len(boarduser_1)
		if boarduser_1<>"" then
			boarduser=left(boarduser_1,userlen-1)
			updateinfo=" boarduser='"&boarduser&"' "
			Dvbbs.Execute("update dv_board set "&updateinfo&" where boardid="&request("boardid"))
			Dv_suc("��̳���óɹ�!<LI>�ɹ�������֤�û���"&boarduser&"<LI><a href=""?action=RestoreBoardCache"" >��ִ���ؽ����滺�������Ч</a><br>")
			RestoreBoardCache()
		else
			response.write "<p><font color=red>��û��������֤�û�</font><br><br>"
			Exit Sub
		end if
	Else
		response.write "<p><font color=red>��û��������֤�û�</font><br><br>"
	End If
	
End Sub

'����������̳��Ϣ
Sub savenew()
	If request("boardtype")="" Then
		Errmsg=Errmsg+"<br>"+"<li>��������̳���ơ�"
		founderr=true
	End If
	If request("class")="" Then
		Errmsg=Errmsg+"<br>"+"<li>��ѡ����̳���ࡣ"
		founderr=true
	End If
	If request("readme")="" Then
		Errmsg=Errmsg+"<br>"+"<li>��������̳˵����"
		founderr=true
	End If
	If founderr=true Then
		dvbbs_error()
		exit sub
	End If
	Dim boardid,rootid,parentid,depth,orders,Fboardmaster,maxrootid,parentstr
	If request("class")<>"0" Then
		Set rs=Dvbbs.Execute("select rootid,boardid,depth,orders,boardmaster,ParentStr from dv_board where boardid="&request("class"))
		rootid=rs(0)
		parentid=rs(1)
		depth=rs(2)
		orders=rs(3)
		If depth+1>20 Then
			Errmsg="����̳�������ֻ����20������"
		  dvbbs_error()
		  Exit Sub
		 End If 
		parentstr=rs(5)
	Else
		Set rs=Dvbbs.Execute("select max(rootid) from dv_board")
	  maxrootid=rs(0)+1
		If  IsNull(MaxRootID) Then MaxRootID=1
	End If
	sql="select boardid from dv_board where boardid="&request("newboardid")
	Set rs=Dvbbs.Execute(sql)
	If not (rs.eof and rs.bof) then
		Errmsg="������ָ���ͱ����̳һ������š�"
		dvbbs_error()
		exit sub
	Else
		boardid=request("newboardid")
	End If
	Dim trs,forumuser,setting
	Set trs=Dvbbs.Execute("select * from dv_setup")
	Setting=Split(trs("Forum_Setting"),"|||")
	forumuser=Setting(2)
	set rs = server.CreateObject ("adodb.recordset")
	sql = "select * from dv_board"
	rs.Open sql,conn,1,3
	rs.AddNew
	If request("class")<>"0" Then
		rs("depth")=depth+1
		rs("rootid")=rootid
		rs("orders") = Request.form("newboardid")
		rs("parentid") = Request.Form("class")
		if ParentStr="0" then
		rs("ParentStr")=Request.Form("class")
	Else
	 rs("ParentStr")=ParentStr & "," & Request.Form("class")
	End If
	Else
		rs("depth")=0
		rs("rootid")=maxrootid
		rs("orders")=0
		rs("parentid")=0
		rs("parentstr")=0
		end if
		rs("boardid") = Request.form("newboardid")
		rs("boardtype") = request.form("boardtype")
		rs("readme") = Request.form("readme")
		rs("Rules") = Request.form("Rules")
		rs("TopicNum") = 0
		rs("PostNum") = 0
		rs("todaynum") = 0
		rs("child")=0
		rs("LastPost")="$0$"&Now()&"$$$$$"
		rs("Board_Setting")="0,0,0,0,0,1,1,1,1,1,1,1,1,1,1,1,16240,3,0,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,0,1,100,20,10,9,normal,1,10,10,0,0,0,0,1,0,0,1,4,0,0,0,200,0,0,,$$,0,0,0,1,0|0|0|0|0|0|0|0|0,0|0|0|0|0|0|0|0|0,0,0,0,0,0,0,0,0,0,��ˮ|���|����|�ͷ�|������|���ݲ���|�ظ�����,0,1,0,24,0,0"
		rs("sid")=request.form("sid")
		rs("cid")=request.form("cid")
		rs("board_ads")=trs("forum_ads")
		rs("board_user")=forumuser
		If Request("boardmaster")<>"" Then 
			rs("boardmaster") = Request.form("boardmaster")
		End If
	If request.form("indexIMG")<>"" Then
		rs("indexIMG")=request.form("indexIMG")
	End If
	rs.Update 
	rs.Close
	If Request("boardmaster")<>"" Then Call addmaster(Request("boardmaster"),"none",0)
	dv_suc("��̳���ӳɹ���<br>����̳Ŀǰ�߼�����ΪĬ��ѡ�������������̳���������������ø���̳�ĸ߼�ѡ�<A HREF=BoardSetting.asp?editid="&Request.form("newboardid")&">����˴�����ð���߼�����</A><br>" & str)
	set rs=nothing
	trs.close
	set trs=nothing
	CheckAndFixBoard 0,1
	RestoreBoardCache()
End Sub

'����༭��̳��Ϣ
Sub savedit()
	if clng(request("editid"))=clng(request("class")) then
		Errmsg="������̳����ָ���Լ�"
		dvbbs_error()
		exit sub
	end if
	dim newboardid,maxrootid
	dim parentid,boardmaster,depth,child,ParentStr,rootid,iparentid,iParentStr
	dim trs,brs,mrs
	Dim iii
	set rs = server.CreateObject ("adodb.recordset")
	sql = "select * from dv_board where boardid="&request("editid")
	rs.Open sql,conn,1,3
	newboardid=rs("boardid")
	parentid=rs("parentid")
	iparentid=rs("parentid")
	boardmaster=rs("boardmaster")
	ParentStr=rs("ParentStr")
	depth=rs("depth")
	child=rs("child")
	rootid=rs("rootid")
	'�ж���ָ������̳�Ƿ���������̳
	if ParentID=0 then
		if clng(request("class"))<>0 then
		set trs=Dvbbs.Execute("select rootid from dv_board where boardid="&request("class"))
		if rootid=trs(0) then
			errmsg="������ָ���ð����������̳��Ϊ������̳1"
			dvbbs_error()
			exit sub
		end if
		end if
	else
		set trs=Dvbbs.Execute("select boardid from dv_board where ParentStr like '%"&ParentStr&","&newboardid&"%' and boardid="&request("class"))
		if not (trs.eof and trs.bof) then
			errmsg="������ָ���ð����������̳��Ϊ������̳2"
			dvbbs_error()
			exit sub
		end if
	end if
	if parentid=0 then
		parentid=rs("boardid")
		iparentid=0
	end if
	rs("boardtype") = Request.Form("boardtype")	'ȡ��JS���ˡ�
	rs("parentid") = Request.Form("class")
	rs("boardmaster") = Request("boardmaster")
	rs("readme") = Request("readme")
	rs("Rules") = Request.form("Rules")
	rs("indexIMG")=request.form("indexIMG")
	rs("sid")=Cint(request.form("sid"))
	rs("cid")=Cint(request.form("cid"))
	rs.Update 
	rs.Close
	set rs=nothing
	if request("oldboardmaster")<>Request("boardmaster") then call addmaster(Request("boardmaster"),request("oldboardmaster"),1)
	
	set mrs=Dvbbs.Execute("select max(rootid) from dv_board")
	Maxrootid=mrs(0)+1
	mrs.close:set mrs=nothing
	dv_suc("��̳�޸ĳɹ���<br>" & str)
	CheckAndFixBoard 0,1
	Boardchild()
	RestoreBoardCache()
End sub

'ɾ�����棬ɾ���������ӣ���ڣ�����ID
Sub Del()
	Dim Trs
	'�������ϼ�������̳�����������̳�����¼���̳������ɾ��
	Set tRs = Dvbbs.Execute("SELECT RootID FROM Dv_Board WHERE BoardID = " & Request("editid"))
	Dim UpdateRootID
	UpdateRootID = tRs(0)
	Set Rs = Dvbbs.Execute("SELECT ParentStr, Child, Depth FROM Dv_Board WHERE BoardID = " & Request("editid"))
	If Not (Rs.Eof And Rs.Bof) Then
		If Rs(1) > 0 Then
			Response.Write "����̳����������̳����ɾ����������̳���ٽ���ɾ������̳�Ĳ���"
			Exit Sub
		End If
		'������ϼ����棬���������
		If Rs(2) > 0 Then
			Dvbbs.Execute("UPDATE Dv_Board SET Child = Child - 1 WHERE BoardID IN (" & Rs(0) & ")")
		End If
		Sql = "DELETE FROM Dv_Board WHERE Boardid = " & Request("editid")
		Dvbbs.Execute(Sql)
		For i = 0 To Ubound(AllPostTable)
			Sql = "DELETE FROM " & AllPostTable(i) & " WHERE BoardID = " & Request("editid")
			Dvbbs.Execute(Sql)
		Next
		Dvbbs.Execute("DELETE FROM Dv_Topic WHERE BoardID = " & Request("editid"))
		Dvbbs.Execute("DELETE FROM Dv_BestTopic WHERE BoardID = " & Request("editid"))
		Dvbbs.Execute("DELETE FROM Dv_Upfile WHERE F_BoardID = " & Request("editid"))
		'ɾ����ɾ����̳���Զ����û�Ȩ�� 2004-11-15 Dv.Yz
		Dvbbs.Execute("DELETE FROM Dv_UserAccess WHERE NOT Uc_BoardID IN (SELECT BoardID FROM Dv_Board)")
	End If
	Set Rs = Nothing
	CheckAndFixBoard 0,1
	RestoreBoardCache()
	Dv_suc("<p>��̳ɾ���ɹ���")
End Sub

sub orders()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr> 
	<th height="22">��̳һ���������������޸�(������Ӧ��̳��������������������Ӧ���������)
	</th>
	</tr>
	<tr>
	<td class="Forumrow"><table width="50%">
<%
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select * from dv_Board where ParentID=0 order by RootID"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "��û����Ӧ����̳���ࡣ"
	else
		do while not rs.eof
		response.write "<form action=board.asp?action=updatorders method=post><tr><td width=""50%"">"&rs("boardtype")&"</td>"
		response.write "<td width=""50%""><input type=text name=""OrderID"" size=4 value="""&rs("rootid")&"""><input type=hidden name=""cID"" value="""&rs("rootid")&""">&nbsp;&nbsp;<input type=submit name=Submit value=�޸�></td></tr></form>"
		rs.movenext
		loop
%>
</table>
<BR>&nbsp;<font color=red>��ע�⣬����ϵͳ��<B>�Զ��޸�</B>����ȷ����ţ�</font>
<%
	end if
	rs.close
	set rs=nothing
%>
	</td>
	</tr>
</table>
<%
end sub

sub updateorders()
	dim cID,OrderID,ClassName
	cID=replace(request.form("cID"),"'","")
	OrderID=replace(request.form("OrderID"),"'","")
	set rs=Dvbbs.Execute("select boardid from dv_board where rootid="&orderid)
	If rs.eof and rs.bof Then
		Dvbbs.Execute("update dv_board set rootid="&OrderID&" where rootid="&cID)
		Dv_suc("���óɹ�")
	Else
		response.write "�벻Ҫ��������̳������ͬ�����"
	end if
	RestoreBoardCache()
end sub

Sub Boardorders()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr> 
	<th height="22">��̳N���������������޸�(������Ӧ��̳��������������������Ӧ���������)
	</th>
	</tr>
	<tr>
	<td class="Forumrow"><table width="90%">
<%
	Dim Trs, Uporders, Doorders
	Set Rs = Server.CreateObject ("Adodb.recordset")
	Sql = "SELECT Depth, Child, Parentid, Boardtype, Orders, BoardId FROM Dv_Board ORDER BY RootID, Orders"
	Set Rs = Dvbbs.Execute(Sql)
	If Rs.Eof And Rs.Bof Then
		Response.Write "��û����Ӧ����̳���ࡣ"
	Else
		Sql = Rs.GetRows(-1)
		Dim Bn
		Rs.Close:Set Rs = Nothing
		For Bn = 0 To Ubound(Sql,2)
			Response.Write "<form action=board.asp?action=updatboardorders method=post><tr><td width=""50%"">"
			If Sql(0,Bn) > 0 Then
				For i = 1 To Sql(0,Bn)
					Response.Write "&nbsp;"
				Next
			End If
			If Sql(1,Bn) > 0 Then
				Response.Write "<img src=../skins/default/plus.gif>"
			Else
				Response.Write "<img src=../skins/default/nofollow.gif>"
			End If
			If Sql(2,Bn) = 0 Then
				Response.Write "<b>"
			End If
			Response.Write Sql(3,Bn)
			If Sql(1,Bn) > 0 Then
				Response.Write "(" & Sql(1,Bn) & ")"
			End If
			Response.Write "</td><td width=""50%"">"
			If Sql(2,Bn) > 0 Then
				'�����ͬ��ȵİ�����Ŀ���õ��ð�������ͬ��ȵİ���������λ�ã�֮�ϻ���֮�µİ�������
				'��������������ӦΪFor i=1 to �ð�֮�ϵİ�����
				Set Trs = Dvbbs.Execute("SELECT COUNT(*) FROM Dv_Board WHERE ParentID = " & Sql(2,Bn) & " AND ORDERS < " & Sql(4,Bn) &"")
				Uporders = Trs(0)
				If Isnull(Uporders) Then Uporders = 0
				'���ܽ���������ӦΪFor i=1 to �ð�֮�µİ�����
				Set Trs = Dvbbs.Execute("SELECT COUNT(*) FROM Dv_Board WHERE ParentID = " & Sql(2,Bn) &" AND ORDERS > " & Sql(4,Bn) &"")
				Doorders = Trs(0)
				If Isnull(doorders) Then Doorders = 0
				If Uporders > 0 Then
					Response.Write "<select name=uporders size=1><option value=0>�����ƶ�</option>"
					For i = 1 To Uporders
						Response.Write "<option value=" & i & ">" & i & "</option>"
					Next
					Response.Write "</select>"
				End If
				If Doorders > 0 Then
					If uporders > 0 Then Response.Write "&nbsp;"
					Response.Write "<select name=doorders size=1><option value=0>�����ƶ�</option>"
					For i = 1 To Doorders
						Response.Write "<option value=" & i & ">" & i & "</option>"
					Next
					Response.Write "</select>"
				End If
				If Doorders > 0 Or Uporders > 0 Then
					Response.Write "<input type=hidden name=""editID"" value=""" & Sql(5,Bn) & """>&nbsp;<input type=submit name=Submit value=�޸�>"
				End If
			End If
			Response.Write "</td></tr></form>"
			Uporders = 0
			Doorders = 0
		Next
		Response.Write "</table>"
	End If
%>
	</td>
	</tr>
</table>
<%
End Sub

'N�������ƶ� 2004-10-18 Dv.Yz
Sub Updateboardorders()
	Dim Orders, tRs, Parentid
	Dim Uporders, Doorders
	Dim Frontorders, Nextorders, Lastorders
	Dim Rootid, Depth, Child, Parentstr
	If Not Isnumeric(Request("EditID")) Then
		Response.Write "�Ƿ��Ĳ�����"
		Exit Sub
	End If
	If Request("Uporders") <> "" And Not Cint(Request("Uporders")) = 0 Then
		If Not Isnumeric(Request("Uporders")) Then
			Response.Write "�Ƿ��Ĳ�����"
			Exit Sub
		Elseif Cint(Request("Uporders")) = 0 Then
			Response.Write "��ѡ��Ҫ���������֣�"
			Exit Sub
		End If
		Uporders = Cint(Request("Uporders"))
		'�����ƶ�
		Set Rs = Dvbbs.Execute("SELECT Orders, Rootid, Depth, ParentID, Child, ParentStr FROM Dv_Board WHERE Boardid = " & Request("EditID"))
		Orders = Rs(0)
		Rootid = Rs(1)
		Depth = Rs(2)
		Parentid = Rs(3)
		Child = Rs(4)
		ParentStr = Rs(5) & "," & Request("EditID")
		Set Rs = Nothing

		'ȡLastordersֵ
		Sql = "SELECT Top 1 Orders FROM Dv_Board WHERE Rootid = " & Rootid & " ORDER BY Orders DESC"
		Set Rs = Dvbbs.Execute(Sql)
		Lastorders = Rs(0)
		Set Rs = Nothing

		'ȡȡ�ƶ�����ORDERSֵ
		If Child > 0 Then
			Sql = "SELECT COUNT(*) FROM Dv_Board WHERE ParentStr LIKE '%" & ParentStr & "%' AND Rootid = " & Rootid
			Set Rs = Dvbbs.Execute(Sql)
			Nextorders = Orders + Rs(0)
		Else
			Nextorders = Orders
		End If
		Doorders = Nextorders
		Set Rs = Nothing

		'ȡͬ���������ϵİ���ORDERSֵ
		Sql = "SELECT Top " & Uporders & " Orders FROM Dv_Board WHERE Rootid = " & Rootid & " AND Depth = " & Depth & " AND ParentID = " & Parentid & " AND Orders < " & Orders & " ORDER BY Orders Desc"
		Set Rs = Dvbbs.Execute(Sql)
		If Rs.Eof And Rs.Bof Then
			Frontorders = 0
		Else
			Sql = Rs.GetRows(-1)
			Frontorders = Sql(0,Ubound(Sql,2))
		End If

		'һ�θ���Orders
		Sql = "UPDATE Dv_Board SET Orders = Orders + " & Doorders & " WHERE Rootid = " & Rootid & " AND (Orders >= " & Frontorders & " AND Orders < " & Orders & " OR Orders > " & Nextorders & ")"
'		Response.Write Sql
		Dvbbs.Execute(Sql)

	Elseif Request("Doorders") <> "" Then
		If Not Isnumeric(Request("Doorders")) Then
			Response.Write "�Ƿ��Ĳ�����"
			Exit Sub
		Elseif Cint(Request("doorders")) = 0 Then
			Response.Write "��ѡ��Ҫ�½������֣�"
			Exit Sub
		End If
		Uporders = Cint(Request("doorders"))
		'�����ƶ�
		Set Rs = Dvbbs.Execute("SELECT Orders, Rootid, Depth, ParentID, Child, ParentStr FROM Dv_Board WHERE Boardid = " & Request("EditID"))
		Orders = Rs(0)
		Rootid = Rs(1)
		Depth = Rs(2)
		Parentid = Rs(3)
		Child = Rs(4)
		ParentStr = Rs(5) & "," & Request("EditID")
		Set Rs = Nothing

		'ȡLastordersֵ
		Sql = "SELECT Top 1 Orders FROM Dv_Board WHERE Rootid = " & Rootid & " ORDER BY Orders DESC"
		Set Rs = Dvbbs.Execute(Sql)
		Lastorders = Rs(0)
		Set Rs = Nothing

		'ȡȡ�ƶ�����ORDERSֵ
		If Child > 0 Then
			Sql = "SELECT COUNT(*) FROM Dv_Board WHERE ParentStr LIKE '%" & ParentStr & "%' AND Rootid = " & Rootid
			Set Rs = Dvbbs.Execute(Sql)
			Nextorders = Orders + Rs(0)
		Else
			Nextorders = Orders
		End If
		Set Rs = Nothing

		'ȡͬ���������º���һ������ORDERSֵ
		Sql = "SELECT Top " & Uporders + 1 & " Orders, Child, ParentStr, BoardID FROM Dv_Board WHERE Rootid = " & Rootid & " AND Depth = " & Depth & " AND ParentID = " & Parentid & " AND Orders > " & Orders & " ORDER BY Orders"
		Set Rs = Dvbbs.Execute(Sql)
		If Rs.Eof And Rs.Bof Then
			Frontorders = Lastorders
		Else
			Sql = Rs.GetRows(-1)
			Frontorders = Sql(0,Ubound(Sql,2)) - 1
			If Not Ubound(Sql,2) = Uporders Then
				If Sql(1,Ubound(Sql,2)) > 0 Then
					ParentStr = Sql(2,Ubound(Sql,2)) & "," & Sql(3,Ubound(Sql,2))
					Set Rs = Dvbbs.Execute("SELECT COUNT(*) FROM Dv_Board WHERE ParentStr LIKE '%" & ParentStr & "%' AND Rootid = " & Rootid)
					Frontorders = Sql(0,Ubound(Sql,2)) + Rs(0)
				Else
					Frontorders = Sql(0,Ubound(Sql,2))
				End If
			End If
		End If
		Doorders = Frontorders

		'һ�θ���Orders
		Sql = "UPDATE Dv_Board SET Orders = Orders + " & Doorders & " WHERE Rootid = " & Rootid & " AND Orders >= " & Orders & " AND Orders <= " & Nextorders & " OR Orders > " & Frontorders
'		Response.Write Sql
		Dvbbs.Execute(Sql)

	End If
	CheckAndFixBoard 0,1
	RestoreBoardCache()
	Response.Redirect "board.asp?action=boardorders"
End Sub

Sub Addmaster(s,o,n)
	Dim Arr, Pw, Oarr
	Dim Classname, Titlepic
	Set Rs = Dvbbs.Execute("SELECT Usertitle, GroupPic FROM Dv_UserGroups WHERE Usergroupid = 3")
	If Not (Rs.Eof And Rs.Bof) Then
		Classname = Rs(0)
		Titlepic = Rs(1)
	End If
	Randomize
	Pw = Cint(Rnd * 9000) + 1000
	Arr = Split(s,"|")
	Oarr = Split(o,"|")
	Set Rs = Server.Createobject("Adodb.Recordset")
	For i = 0 To Ubound(Arr)
		Sql = "SELECT * FROM [Dv_User] WHERE Username = '" & Arr(i) & "'"
		Rs.Open Sql,Conn,1,3
		If Rs.Eof And Rs.Bof Then
			Rs.Addnew
			Rs("Username") = Arr(i)
			Rs("Userpassword") = Md5(Pw,16)
			Rs("Userclass") = Classname
			Rs("UserGroupID") = 3
			Rs("Titlepic") = Titlepic
			Rs("UserWealth") = 100
			Rs("Userep") = 30
			Rs("Usercp") = 30
			Rs("Userisbest") = 0
			Rs("Userdel") = 0
			Rs("Userpower") = 0
			Rs("Lockuser") = 0
			'�������ϸ����ʹ��¼����ʾ���ϲ��������
			Rs("UserSex") = 1
			Rs("UserEmail") = Arr(i) & "@aspsky.net"
			Rs("UserFace") = "Images/userface/image1.gif"
			Rs("UserWidth") = 32
			Rs("UserHeight") = 32
			Rs("UserIM") = "||||||||||||||||||"
			Rs("UserFav") = "İ����,�ҵĺ���,������"
			Rs("LastLogin") = Now()
			Rs("JoinDate") = Now()
			Rs("Userpost") = 0
			Rs("Usertopic") = 0
			Rs.Update
			Str = Str & "�������������û���<b>" & Arr(i) & "</b> ���룺<b>" & Pw & "</b><br><br>"
			Dvbbs.Execute("UPDATE Dv_Setup SET Forum_Usernum = Forum_Usernum + 1, Forum_Lastuser = '" & Arr(i) & "'")
		Else
			'�������Ӱ������ı�ȼ��Ĵ��� 2005-3-7 Dv.Yz
			If Rs("UserGroupID") > 3 Then
				Rs("Userclass") = Classname
				Rs("UserGroupID") = 3
				Rs("Titlepic") = Titlepic
				Rs.Update
			End If
		End If
		Rs.Close
	Next
	'�ж�ԭ���������������Ƿ񻹵��ΰ�������û�е����򳷻����û�ְλ��
	If n = 1 Then
		Dim Iboardmaster
		Dim UserGrade, Article
		Iboardmaster = False
		For i = 0 To Ubound(Oarr)
			Set Rs = Dvbbs.Execute("SELECT Boardmaster FROM Dv_Board")
			Do While Not Rs.Eof
				If Instr("|" & Trim(Rs("Boardmaster")) & "|","|" & Trim(Oarr(i)) & "|") > 0 Then
					Iboardmaster = True
					Exit Do
				End If
				Rs.Movenext
			Loop
			If Not Iboardmaster Then
				Set Rs = Dvbbs.Execute("SELECT Userid, UserGroupID, UserPost FROM [Dv_User] WHERE Username = '" & Trim(Oarr(i)) & "'")
				If Not (Rs.Eof And Rs.Bof) Then
					If Rs(1) > 2 Then
						If Not Isnumeric(Rs(2)) Or Rs(2) = "" Then
							Article = 0
						Else
							Article = Cstr(Rs(2))
						End If
						'ȡ��Ӧע���Ա�ĵȼ�
						Set UserGrade = Dvbbs.Execute("SELECT TOP 1 Usertitle, Grouppic,UserGroupID FROM Dv_Usergroups WHERE Minarticle <= " & Article & " AND NOT MinArticle = -1 AND ParentGID = 4 ORDER BY MinArticle DESC")
						If Not (UserGrade.Eof And UserGrade.Bof) Then
							Dvbbs.Execute("UPDATE [Dv_User] SET UserGroupID = 10, Titlepic = '" & UserGrade(1) & "', Userclass = '" & UserGrade(0) & "' WHERE Userid = " & Rs(0))
						End If
						UserGrade.Close:Set UserGrade = Nothing
					End If
				End If
			End If
			Iboardmaster = False
		Next
	End If
	Set Rs = Nothing
End Sub

Rem �ְ����û�Ȩ������ ��д2004-5-2 Dvbbs.YangZheng
Sub BoardPerMission()
	Dim iUserGroupID(100), UserTitle(100),iParentID(100)
	Dim Trs, Ars, k, ii
	Dim Bn
	Set Trs = Dvbbs.Execute("SELECT Usertitle,Usergroupid,ParentGID FROM Dv_UserGroups WHERE Not ParentGID=0 ORDER BY ParentGID,UserGroupId")
	If Not (Trs.Eof And Trs.Bof) Then
		Sql = Trs.GetRows(-1)
		Trs.Close:Set Trs = Nothing
		For ii = 0 To Ubound(Sql,2)
			UserTitle(ii) = Sql(0,ii)
			iUserGroupID(ii) = Sql(1,ii)
			iParentID(ii) = Sql(2,ii)
		Next
	End If
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr>
	<th height="25">�༭��̳Ȩ��</th>
	</tr>
	<tr>
	<td class=forumrow>
	�����������ò�ͬ�û����ڲ�ͬ��̳�ڵ�Ȩ�ޣ���ɫ��ʾΪ����̳���û���ʹ�õ����û���������<BR>
	�ڸ�Ȩ�޲��ܼ̳У�������������һ�������¼���̳�İ��棬��ôֻ�������õİ�����Ч��������������̳��Ч<BR>
	���������������Ч������������ҳ��<B>ѡ���Զ�������</B>��ѡ�����Զ������ú��������õ�Ȩ�޽�<B>����</B>���û������ã������û���Ĭ�ϲ��ܹ������ӣ������������˸��û���ɹ������ӣ���ô���û������������Ϳ��Թ�������
	</td>
	</tr>
</table><BR>
<table width="95%" cellspacing="1" cellpadding="1" align=center class="tableBorder">
<tr> 
<th width="35%" class="tableHeaderText" height=25>��̳����
</th>
<th width="35%" class="tableHeaderText" height=25>�����û���Ȩ��
</th>
</tr>
<%
	Dim Percount
	Sql = "SELECT Depth, Child, Parentid, BoardType, Boardid,IsGroupSetting FROM Dv_Board ORDER BY Rootid, Orders"
	Set Rs = Dvbbs.Execute(Sql)
	If Not (Rs.Eof And Rs.Bof) Then
		Sql = Rs.GetRows(-1)
		Set Rs = Nothing
		For Bn = 0 To Ubound(Sql,2)
			Response.Write "<tr><td height=25 width=40% class=forumrow>"
			If Sql(0,Bn) > 0 Then
				For i = 1 To Sql(0,Bn)
					Response.Write "&nbsp;"
				Next
			End If
			If Sql(1,Bn) > 0 Then
				Response.Write "<img src=../skins/default/plus.gif>"
			Else
				Response.Write "<img src=../skins/default/nofollow.gif>"
			End If
			If Sql(2,Bn) = 0 Then
				Response.Write "<b>"
			End If
			Response.Write Sql(3,Bn)
			If Sql(1,Bn) > 0 Then
				Response.Write "(" & Sql(1,Bn) & ")"
			End If
			'Percount = Dvbbs.Execute("SELECT COUNT(*) FROM Dv_BoardPermission WHERE Boardid = " & Sql(4,Bn))(0)
%>
</td>
<FORM METHOD=POST ACTION="?action=editpermission">
<td width=60% class="forumrow">&nbsp;
<select name="groupid" size=1>
<%
Dim hasc
	
			For k = 0 To ii-1
				Response.Write "<option value=""" & iUserGroupID(k) & """>" & SysGroupName(iParentID(k)) & UserTitle(k)
				If Sql(5,Bn)<>"" Then
					Set Ars = Dvbbs.Execute("SELECT Pid FROM Dv_BoardPerMission WHERE BoardID = " & Sql(4,Bn) & " AND GroupID = " & iUserGroupID(k))
					If Not Ars.Eof Then
						Response.Write "(�Զ���)"
						hasc=1
					End If
				End If
				Response.Write "</option>"
			Next
			Response.Write "</select><input type=hidden value="
			Response.Write Sql(4,Bn)
			Response.Write " name=reboardid><input type=submit name=submit value=����>"
			If hasc=1 Then
				Response.Write "(���Զ������)"
			End If
			Response.Write "</td></FORM></tr>"
		Next
	End If
	Response.Write "</table><BR><BR>"
	Set Ars = Nothing
	Set Trs = Nothing
End Sub

Sub editpermission()
	if not isnumeric(request("groupid")) Or request("groupid")="" Or request("reBoardID")="" Or not isnumeric(request("reBoardID"))  then
	response.write "����Ĳ������뷵�طְ���Ȩ��������ҳѡ����ȷ�����ã�"
	exit sub
	end if
	if request("groupaction")="yes" then
		dim GroupSetting,rspid,SaveGroupid,NewGroupid
		Dim IsGroupSetting,MyIsGroupSetting
		Dim Sql,i,k
		Dim UpdateStr,OldStr,NewStr
		GroupSetting=GetGroupPermission
		If Not IsNumeric(Replace(Request.Form("GroupID"),",","")) or Request.Form("GroupID")="" Then
			Errmsg = ErrMsg + "<BR><li>��ѡ���Ӧ���û��顣"
			Dvbbs_Error()
			Exit Sub
		End If
		SaveGroupid = Request.Form("groupid")
		Set rs= Server.CreateObject("ADODB.Recordset")
		if Request("isdefault")=1 then
			'����ID
			Dvbbs.Execute("Delete from dv_BoardPermission where BoardID="&request.Form("reBoardID")&" and GroupID in ("&SaveGroupid&")")
		Else
			'ʹ���Զ���
			sql="Select Pid,GroupID,PSetting from dv_BoardPermission where BoardID="&request("reBoardID")&" And GroupID in ("&SaveGroupid&") "
			NewGroupid = ","&SaveGroupid&","
			Set Rs=Dvbbs.Execute(sql)
			If Not Rs.eof And Not Rs.bof Then
				If Instr(SaveGroupid,",")=0 Then
					sql="update dv_BoardPermission set PSetting='"&GroupSetting&"' where pid="&Rs(0)
					Dvbbs.Execute(Sql)
					NewGroupid = ""
				Else
					Do while Not Rs.Eof
						NewStr = Split(GroupSetting,",")
						OldStr = Split(Rs(2),",")
						UpdateStr = ""
						For K = 0 To 90
							If Request.Form("CheckGroupSetting("&K&")")="on" Then
								UpdateStr = UpdateStr & NewStr(k)
							Else
								UpdateStr = UpdateStr & OldStr(k)
							End If
							If K<90 Then
								UpdateStr = UpdateStr & ","
							End If
						Next
						sql="update dv_BoardPermission set PSetting='"&UpdateStr&"' where pid="&Rs(0)
						Dvbbs.Execute(Sql)
						NewGroupid = Replace(NewGroupid,","&Rs(1)&",",",")
					Rs.MoveNext
					Loop
				End If
			Else
				Dim iSaveGroupID
				iSaveGroupID = Split(SaveGroupid,",")
				For i = 0 To Ubound(iSaveGroupID)
					sql="insert into dv_BoardPermission (BoardID,GroupID,PSetting) values ("&request("reBoardID")&","&iSaveGroupID(i)&",'"&GroupSetting&"')"
					Dvbbs.Execute(Sql)
				Next
				NewGroupid = ""
			End If
			Set Rs=Nothing


			If Replace(NewGroupid,",","")<>"" Then
				'����������
				NewGroupid = Split(NewGroupid,",")
				For i=1 to Ubound(NewGroupid)-1
					Sql = Dvbbs.Execute("select GroupSetting From Dv_UserGroups where UserGroupID="&NewGroupid(i))(0)
					If Sql<>"" Then
						NewStr = Split(GroupSetting,",")
						OldStr = Split(Sql,",")
						UpdateStr = ""
						For K = 0 To 90
							If Request.Form("CheckGroupSetting("&K&")")="on" Then
								UpdateStr = UpdateStr & NewStr(k)
							Else
								UpdateStr = UpdateStr & OldStr(k)
							End If
							If K<90 Then
								UpdateStr = UpdateStr & ","
							End If
						Next
						sql="insert into dv_BoardPermission (BoardID,GroupID,PSetting) values ("&request("reBoardID")&","&NewGroupid(i)&",'"&UpdateStr&"')"
						Dvbbs.Execute(Sql)
					End If
				Next
			End If
		End If

		'������ȡ���Զ����ID
		Dim IsGroupSetting1
		IsGroupSetting=Get_board_AccUserList(CLng(request("reBoardID")))
		IsGroupSetting1=Get_Board_GroupSetting(CLng(request("reBoardID")))
		If IsGroupSetting="" Then
			IsGroupSetting=IsGroupSetting1
		ElseIf IsGroupSetting1<>"" Then
			IsGroupSetting=IsGroupSetting&","& IsGroupSetting1
		End If
		Dvbbs.Execute("update dv_Board set IsGroupSetting='"&IsGroupSetting&"' Where BoardID="&CLng(request("reBoardID")))
		RestoreBoardCache()
		Set Rs=Nothing
		Dv_suc("�޸ĳɹ���")
Else
	Dim reGroupSetting,reBoardID,groupid
	Dim Groupname,Boardname,founduserper
	founduserper=false
	if request("GroupID")<>"" then
	set rs=Dvbbs.Execute("select * from dv_BoardPermission where boardid="&request("reBoardID")&" and GroupID="&request("GroupID"))
	if rs.eof and rs.bof then
		founduserper=false
	else
	groupid=rs("groupid")
	reGroupSetting=rs("PSetting")
	reBoardID=rs("boardid")
	set rs=Dvbbs.Execute("select usertitle from dv_UserGroups where usergroupid="&groupid)
	groupname=rs("usertitle")
	founduserper=true
	end if
	if not founduserper then
	set rs=Dvbbs.Execute("select * from dv_usergroups where usergroupid="&request("groupid"))
	if rs.eof and rs.bof then
	response.write "δ�ҵ����û��飡"
	exit sub
	end if
	groupid=request("groupid")
	reGroupSetting=rs("GroupSetting")
	reBoardID=request("reBoardID")
	Groupname=rs("usertitle")
	end if
	end if
	set rs=Dvbbs.Execute("select boardtype from dv_board where boardid="&reBoardID)
	Boardname=rs("boardtype")
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<FORM METHOD=POST ACTION="?action=editpermission">
<input type=hidden name="groupid" value="<%=groupid%>">
<input type=hidden name="reBoardID" value="<%=reBoardID%>">
<input type=hidden name="pID" value="<%=request("pid")%>">
<tr> 
<th height="23" colspan="4" >�༭��̳�û���Ȩ��&nbsp;>> <%=boardname%>&nbsp;>> <%=groupname%></th>
</tr>
<tr> 
<td height="23" colspan="4" class=forumrow><input type=radio name="isdefault" value="1" <%if not founduserper then%>checked<%end if%>><B>ʹ���û���Ĭ��ֵ</B> (ע��: �⽫ɾ���κ�֮ǰ�������Զ�������)</td>
</tr>
<tr> 
<td height="23" colspan="4"  class=forumrow><input type=radio name="isdefault" value="0" <%if founduserper then%>checked<%end if%>><B>ʹ���Զ�������</B>&nbsp;(ѡ���Զ������ʹ����������Ч) </td>
</tr>
<tr> 
<td height="23" colspan="4"  class=forumrow>
��<font color="red">�ڸ��¶���û�������ʱ����ѡȡ����ߵĸ�ѡ������ֻ��ѡȡ��������Ŀ�Ż���£�</font>;
<BR>�ڲ�ִ�ж��û������ʱ������Ҫѡȡ��ߵĸ�ѡ������
<br><b>���������û�������</b>��<input type="button" value="ѡ���û���" onclick="getGroup('Select_Group');">
 <INPUT TYPE="checkbox" NAME="chkall" onclick="CheckAll(this.form);">[ȫѡ]
</td>
</tr>
<%
GroupPermission(reGroupSetting)
%>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
<%
Call Select_Group(Request("groupid"))
end if
end sub

sub RestoreBoard()
	'����Ŀǰ������ѭ��i��ֵ����rootid
	'��ԭ���а����depth,orders,parentid,parentstr,childΪ0
	i=0
	set rs=Dvbbs.Execute("select boardid from dv_board order by rootid,orders")
	do while not rs.eof
	i=i+1
	Dvbbs.Execute("update dv_board set rootid="&i&",depth=0,orders=0,ParentID=0,ParentStr='0',child=0 where boardid="&rs(0))
	rs.movenext
	loop
	Set Rs=Nothing
	Dv_suc("�뷵������̳�������á���λ")
	RestoreBoardCache()
End sub

Sub clearDate
	If Dvbbs.Boardid=0 Then
		errmsg=errmsg+"<br><li>��ѡ����̳����"
		dvbbs_error()
		Exit Sub
	End If
	Dim Rs,str1,str2,str3,str4
	Set Rs=Dvbbs.Execute("Select Count(*) from dv_topic where Boardid="& Dvbbs.boardid &"")
	str1=Rs(0)
	str3=0
	str4=0
	For i= 0 to UBound(AllPostTable)
		Set Rs=Dvbbs.Execute("Select Count(*) from "&AllPostTable(i)&" where Boardid="& Dvbbs.boardid &"")
		str2=str2&"������"&AllPostTable(i)&"��"&Rs(0)&"ƪ���£�"
		str3=str3+Rs(0)
		Set Rs=Dvbbs.Execute("Select Count(*) from "&AllPostTable(i)&" where Boardid="& Dvbbs.boardid &" and isbest=1")
		str4=str4+Rs(0)
	Next
	Response.Write"<br>"
	Response.Write"<table cellpadding=0 cellspacing=0 align=center class=""tableBorder"" style=""width:90%"">"
	Response.Write"<tr align=center>"
	Response.Write"<th width=""100%"" height=25 colspan=2>"
	Response.Write Dvbbs.BoardType
	Response.Write "-������Ϣ"
	Response.Write"</td>"
	Response.Write"</tr>"
	Response.Write"<tr>"
	Response.Write"<td width=""100%"" class=""ForumrowHighLight"" colspan=2>"
	Response.Write "<li>��������:<b>"
	Response.Write str1
	Response.Write "</b><li>��������:<b>"
	Response.Write str3
	Response.Write "</b><li>"
	Response.Write str2
	Response.Write "<li>��<B>"&str4&"</B>ƪ��������"
	Response.Write"</td></tr>"
	Response.Write "<form action =""?action=delDate&boardid="&Dvbbs.boardid&""" method=post>"
	Response.Write"<tr>"
	Response.Write"<td class=""ForumrowHighLight"" valign=middle colspan=2 align=left><li>  ���<b>"
	Response.Write Dvbbs.BoardType
	Response.Write "</b>�� "
	Response.Write "<select name=""tablelist""><option value=""all"">�������ݱ�</option>"
	For i= 0 to UBound(AllPostTable)
		Response.Write "<option value="""&AllPostTable(i)&""">"
		Response.Write 	AllPostTableName(i)
		Response.Write "</option>"
	Next 
	Response.Write "</select>"
	Response.Write " �� <input type=text name=dd value=365 size=5 > ��ǰ������"
	Response.Write " <input type=""submit"" name=""Submit"" value=""ִ ��""> <b>ע��:�˲������ɻָ���</b>���о��������ᱻɾ����<BR><BR>���������̳�����ڶִ࣬�д˲��������Ĵ����ķ�������Դ��ִ�й��������ĵȺ����ѡ��ҹ���������ٵ�ʱ����¡�"
	Response.Write "</td></tr>"
	Response.Write "</form>"
	Response.Write"</table>"
End Sub
Sub delDate
	If Dvbbs.Boardid=0 Then
		errmsg=errmsg+"<br><li>��ѡ����̳����"
		dvbbs_error()
		Exit Sub
	End If
	Dim tablelist
	If request.form("tablelist")<>"all" Then
		tablelist=Dvbbs.checkstr(request.form("tablelist"))
	Else
		For i= 0 to UBound(AllPostTable)
		If i=0 Then
			tablelist=AllPostTable(i)
		Else
			tablelist=tablelist&","&AllPostTable(i)
		End If
		Next
	End If
	tablelist=split(tablelist,",")
	Dim SqlTopic
	Dim k
	k=0
	For i= 0 to UBound(tablelist)
		'ɾ�����ݱ���¼
		If IsSqlDataBase=1 Then
		SqlTopic="Select TopicID,isvote,PollID from dv_Topic where Boardid="&Dvbbs.boardid&" and isbest=0 and PostTable='"&tablelist(i)&"' and Datediff(d,LastPostTime,"&SqlNowString&") > "& CLng(request.form("dd"))&" "
		Else
		SqlTopic="Select TopicID,isvote,PollID from dv_Topic where Boardid="&Dvbbs.boardid&" and isbest=0 and PostTable='"&tablelist(i)&"' and Datediff('d',LastPostTime,"&SqlNowString&") > "& CLng(request.form("dd"))&" "
		End If
		Set rs=Dvbbs.Execute(SqlTopic)
		Do While Not Rs.Eof
			Sql="Delete from "&tablelist(i)&" where Boardid="&Dvbbs.boardid&" and rootid="&RS(0)&""
			Dvbbs.Execute(Sql) 
			If Rs(1)=1 And Not IsNull(Rs(2)) Then
				Sql="Delete from dv_vote where voteid="&RS(2)&""
				Dvbbs.Execute(Sql)
			End If 
			Rs.movenext
		k=k+1
		Loop 
		'ɾ���������¼
		If IsSqlDataBase=1 Then
		SqlTopic="Delete from dv_Topic where Boardid="&Dvbbs.boardid&" and isbest=0 and PostTable='"&tablelist(i)&"' and Datediff(d,LastPostTime,"&SqlNowString&") > "& CLng(request.form("dd"))&" "
		Else
		SqlTopic="Delete from dv_Topic where Boardid="&Dvbbs.boardid&" and isbest=0 and PostTable='"&tablelist(i)&"' and Datediff('d',LastPostTime,"&SqlNowString&") > "& CLng(request.form("dd"))&" "
		End If
		Dvbbs.Execute(SqlTopic) 
		Set rs=Nothing 	
	Next
	Response.Write "ɾ����"&k&"�����⡣"
End Sub

Sub RestoreClass()
	Dim ClassID,RootID,RootIDNum,ParentID
	ClassID=Request("ClassID")
	If Not IsNumeric(ClassID) Or ClassID="" Then
		Response.Write "����İ��������"
		Exit Sub
	Else
		ClassID=Clng(ClassID)
	End If
	Set Rs=Dvbbs.Execute("Select RootID,BoardID From Dv_Board Where BoardID="&ClassID)
	If Rs.Eof And Rs.Bof Then
		Response.Write "����İ��������"
		Exit Sub
	Else
		RootID=Rs(0)
		ParentID=Rs(1)
	End If
	i=0
	Set Rs=Dvbbs.Execute("Select BoardID,ParentID From Dv_Board Where RootID="&RootID&" Order By ParentID,Orders,Depth")
	Do While Not Rs.Eof
		If Rs(1)=0 Then
			Dvbbs.Execute("UpDate Dv_Board Set Orders="&i&" Where BoardID="&Rs(0))
		Else
			Dvbbs.Execute("UpDate Dv_Board Set Orders="&i&",ParentID="&ParentID&",ParentStr='"&ParentID&"',Depth=1,child=0 Where BoardID="&Rs(0))
		End If
		i=i+1
	Rs.MoveNext
	Loop
	Set Rs=Dvbbs.Execute("Select Count(*) From Dv_Board Where RootID="&RootID)
	RootIDNum=Rs(0)
	If IsNull(RootIDNum) Or RootIDNum="" Then
		RootIDNum=0
	Else
		RootIDNum=RootIDNum-1
	End If
	Dvbbs.Execute("UpDate Dv_Board Set Child="&RootIDNum&" Where BoardID="&ClassID)
	dv_suc("��λ����ɹ���")
	RestoreBoardCache()
	Set Rs=Nothing
End Sub

Sub handorders()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr> 
	<th height="22">��̳�������������޸�(������Ӧ��̳��������������������Ӧ���������)
	</th>
	</tr>
	<tr>
	<td class="Forumrow">
	<B>ע��</B>��<BR>
1�����ڱ���̳�����㷨���ǵݹ飬��������ȷ�����������ţ�����������̳��ʾ��������<font color=red>�����δ��ȷ�˽�˵�����벻Ҫ�������</font><BR>
2��һ�������������Ϊ0������ȷ���룬<font color=blue>�������������������</font><BR>
3���������Ϊ�����������ں��棬<font color=blue>�����ﲻ����������ָ��ĳ�������������������</font>������Ϊ��ȷ���������뷽ʽ��<BR>
<B>����</B> 0<BR>
--��������A 1<BR>
--��������B 2<BR>
----��������A 3<BR>
----��������B 4<BR>
----��������C 5<BR>
--��������C 6<BR>
A.<font color=blue>Ҫ����������C�ᵽ��������A����</font>�����������룺����(0)-����A(1)-����B(2)-����A(<font color=red>4</font>)-����B(<font color=red>5</font>)-����C(<font color=red>3</font>)-����C(6)<BR>
B.<font color=blue>Ҫ�Ѷ�������C�ᵽ��������B����</font>�����������룺����(0)-����A(1)-����B(<font color=red>3</font>)-����A(<font color=red>4</font>)-����B(<font color=red>5</font>)-����C(<font color=red>6</font>)-����C(<font color=red>2</font>)<BR>
B.<font color=blue>Ҫ�Ѷ�������B�ᵽ��������A����</font>�����������룺����(0)-����A(<font color=red>5</font>)-����B(<font color=red>1</font>)-����A(<font color=red>2</font>)-����B(<font color=red>3</font>)-����C(<font color=red>4</font>)-����C(6)
	</td></tr>
<form action="board.asp?action=savehandorders" method=post>
	<tr>
	<td class="Forumrow"><table width="90%">
<%
dim trs,uporders,doorders,RootID
Set Rs=Dvbbs.Execute("Select RootID From Dv_Board Where BoardID="&Request("classid"))
If Rs.eof And Rs.bof Then
	response.write "��û����Ӧ����̳���ࡣ"
	exit sub
Else
	RootID=Rs(0)
End If
set rs = server.CreateObject ("Adodb.recordset")
sql="select * from dv_Board Where RootID="&RootID&" order by RootID,orders"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	response.write "��û����Ӧ����̳���ࡣ"
else
	do while not rs.eof
	response.write "<tr><td width=""50%"">"
	if rs("depth")>0 then
	for i=1 to rs("depth")
		response.write "&nbsp;"
	next
	end if
	if rs("child")>0 then
		response.write "<img src=../skins/default/plus.gif>"
	else
		response.write "<img src=../skins/default/nofollow.gif>"
	end if
	if rs("parentid")=0 then
		response.write "<b>"
	end if
	response.write rs("boardtype")
	if rs("child")>0 then
		response.write "("&rs("child")&")"
	end if
	response.write "</td><td width=""50%"">"
	Response.Write "<input type=hidden value="""&rs("boardid")&""" name=getboard>"
	Response.Write "<input type=text size=5 value="""&rs("orders")&""" name=orders>"
	response.write "</td></tr>"
	uporders=0
	doorders=0
	rs.movenext
	loop
	Response.Write "<tr><td class=Forumrow><input type=submit name=submit value=�ύ></td></tr>"
	response.write "</table>"
end if
rs.close
set rs=nothing
%>
	</td>
	</tr></form>
</table>
<%
End Sub

Sub savehandorders()
	dim cID,OrderID,ClassName
	cID=replace(request.form("getboard"),"'","")
	OrderID=replace(request.form("Orders"),"'","")
	For i=1 to request.form("getboard").count
		cID=request.form("getboard")(i)
		OrderID=request.form("Orders")(i)
		Dvbbs.Execute("Update Dv_Board Set Orders="&OrderID&" Where BoardID="&cID)
	next
	Dv_suc("���ķ�������ɹ���")
	CheckAndFixBoard 0,1
	RestoreBoardCache()
End Sub

Sub RestoreBoardCache()
	
	Dim Board
	Dvbbs.LoadBoardList()
	For Each board in Application(Dvbbs.CacheName&"_boardlist").documentElement.selectNodes("board/@boardid")
		Dvbbs.LoadBoardData board.text
		Dvbbs.LoadBoardinformation board.text
	Next
	If Request("action")="RestoreBoardCache" Then dv_suc("�ؽ����а��滺��ɹ���")
End Sub
Sub CheckAndFixBoard(ParentID,orders)
	Dim Rs,Child,ParentStr
	If ParentID=0 Then
		Dvbbs.Execute("update dv_board set Depth=0,ParentStr='0' where ParentID=0")
	End If
	Set Rs=Dvbbs.Execute("Select BoardID,rootid,ParentStr,Depth From Dv_Board  where ParentID="&ParentID&" order by Rootid,orders")
	Do while Not Rs.EOF
		If Rs(2)<>"0" Then
			ParentStr=Rs(2)&","&Rs(0)
		Else
			ParentStr=Rs(0)
		End If
		Conn.Execute "update dv_board set Depth="&Rs(3)+1&",ParentStr='"&ParentStr&"',rootid="&rs(1)&" where ParentID="&Rs(0)&"",Child
		Dvbbs.Execute("update dv_board set Child="&Child&",orders="&orders&" Where BoardID="&Rs(0)&"")
		orders=orders+1
		CheckAndFixBoard Rs(0),orders
		Rs.MoveNext
	Loop
	Set Rs=Nothing
End Sub
Function Get_board_AccUserList(bid)
	Dim Rs,tmp
	Set Rs=Dvbbs.Execute("Select uc_userid from dv_UserAccess where uc_boardid="&bid&"")
	tmp=""
	If Not Rs.EOF Then
		Do while Not Rs.EOF
			If tmp="" Then
				tmp="0_"&rs(0)
			Else
				tmp=tmp&",0_"&rs(0)
			End If
		Rs.MoveNext
		Loop
		Get_board_AccUserList=tmp
	Else
		Get_board_AccUserList=""
	End If
	Set Rs=Nothing
End Function
Function Get_Board_GroupSetting(bid)
	Dim Rs,tmp
	Set Rs=Dvbbs.Execute("select GroupID From Dv_BoardPermission Where BoardID="&bid)
	tmp=""
	If Not Rs.EOF Then
		Do while Not Rs.EOF
			If tmp="" Then
				tmp=rs(0)
			Else
				tmp=tmp&","&rs(0)
			End If
		Rs.MoveNext
		Loop
		Get_Board_GroupSetting=tmp
	Else
		Get_Board_GroupSetting=""
	End If
	Set Rs=Nothing
End Function
Rem ͳ��������̳���� 2004-5-3 Dvbbs.YangZheng
Sub Boardchild()
	Dim cBoardNum, cBoardid
	Dim Trs
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
				cBoardid = Split(Sql(5,Bn),",")
				For i = 1 To Ubound(cBoardid)
					Dvbbs.Execute("UPDATE Dv_Board SET Child = Child + 1 WHERE Boardid = " & cBoardid(i))
				Next
			End If
		Next
	End If
End Sub
%>