<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="Plus_popwan/cls_setup.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/GroupPermission.asp"-->
<%
	Dim Action

	Dim Errmsg,founderr,str
	Dim AllPostTableName,AllPostTable

	AllPostTable1
	AllPostTableName=Split(AllPostTableName,"|")
	AllPostTable=Split(AllPostTable,"|")

	Dvbbs.LoadTemplates("")
	Dvbbs.Stats = "��Ϸ��������"
	Dvbbs.Nav()
	Dvbbs.Head_var 0,0,Plus_Popwan.Program,"plus_popwan_forum.asp"
	Dvbbs.ActiveOnline()
	action = Request("action")
	Page_main()

	If action<>"frameon" Then
		Dvbbs.Footer
	End If
	Dvbbs.PageEnd()

'ҳ���Ҳ����ݲ���
Sub Page_Center()
	If Not Plus_Popwan.IsMaster Then
		Dvbbs.AddErrcode(28)
		Dvbbs.ShowErr()
	End If
	
	info()
	Select Case action
		case "add" : add()
		case "edit" : edit()
		case "savenew" : savenew()
		Case "savedit" : savedit()
		Case "del" : del()
		Case Else
			Call boardinfo()
	End Select
End Sub

'=============
'������� ��ʼ
'=============
Sub info()
%>
<style type="text/css">
table {width:100%;}
td {padding-left:5px;}
</style>
<table cellspacing="0" cellpadding="0" class="pw_tb1">
<tr><th>��Ϸ��������</th></tr>
<tr><td>
<ul>
<li>˵���٣���ҳ��������Ϸ���棬���Ȩ�����ã��ɵ���̨��������н������ã�</li>
<li>˵���ڣ������б�����У��ɷ���̶����͹��档</li>
</ul>
</td></tr>
<tr><td>
<a href="?action=add">�����Ϸ����</a>
| <a href="plus_popwan_forum.asp">��Ϸ��������</a>
</td></tr>
</table><br/>
<%
End Sub

sub boardinfo()
Dim reBoard_Setting
Dim Rs,classrow,iii,SQL,i
%>
<table cellspacing="0" cellpadding="0" class="pw_tb1">
<tr> 
<th width="35%">��̳����
</th>
<th width="*">����
</th>
</tr>
<%
SQL="select boardid,boardtype,parentid,depth,child,Board_setting from dv_board order by rootid,orders"
SET Rs = Dvbbs.Execute(SQL)
If Rs.eof Then
	Rs.close:Set Rs = Nothing
Else
SQL=Rs.GetRows(-1)
Rs.close:Set Rs = Nothing
For iii=0 To Ubound(SQL,2)
	reBoard_Setting=split(SQL(5,iii),",")
	Response.Write "<tr>"
	Response.Write "<td height=""25"" width=""55%"">"
	if SQL(3,iii)>0 then
		for i=1 to SQL(3,iii)
			Response.Write "&nbsp;&nbsp;"
		next
	end if
	if SQL(4,iii)>0 then
		Response.Write "<img src=""skins/default/plus.gif"">"
	else
		Response.Write "<img src=""skins/default/nofollow.gif"">"
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
<td width="45%">
<a href="?action=add&editid=<%=SQL(0,iii)%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>��Ӱ���</U></font></a> | <a href="?action=edit&editid=<%=SQL(0,iii)%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>��������</U></font></a> | <%if SQL(4,iii)=0 then%><a href="?action=del&editid=<%=SQL(0,iii)%>" onclick="{if(confirm('ɾ������������̳���������ӣ�ȷ��ɾ����?')){return true;}return false;}"><font color="<%=Dvbbs.mainsetting(3)%>"><U>ɾ��</U></a><%else%><a href="#" onclick="{if(confirm('����̳����������̳��������ɾ����������̳����ɾ������̳��')){return true;}return false;}"><font color="<%=Dvbbs.mainsetting(3)%>"><U>ɾ��</U></font></a><%end if%> | <a href="plus_popwan_post.asp?boardid=<%=Sql(0,iii)%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>����̶���</U></font></a> | <a href="plus_popwan_ann.asp?action=addann&boardid=<%=Sql(0,iii)%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>������</U></font></a>
</td></tr>
<%
Next
End If
%>
</table>
<%
end sub

sub add()
dim rs_c,sql,Rs
Dim forum_sid,forum_cid,Style_Option,TempOption
Dim iCssName,iCssID,iStyleName
set rs_c= server.CreateObject ("adodb.recordset")
sql = "select * from dv_board order by rootid,orders"
If Not IsObject(Conn) Then ConnectionDatabase
rs_c.open sql,conn,1,1
	dim boardnum,i
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select Max(boardid) from dv_board"
	If Not IsObject(Conn) Then ConnectionDatabase
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
<table cellspacing="0" cellpadding="0" class="pw_tb1">
<form action ="?action=savenew" method="post" name="theform">
<input type="hidden" name="newboardid" value="<%=boardnum%>">
<tr> 
<th colspan="2" style="text-align:center;"><B>�����Ϸר����</th>
</tr>
<tr> 
<td width="100%" height="30" colspan="2">
˵����<BR>1�������̳�������ص����þ�ΪĬ�����ã��뷵����̳���������ҳ�����б�ĸ߼����������ø���̳����Ӧ���ԣ��������Ը���̳���������Ȩ�����ã��뵽<font color=blue>��̨��̳Ȩ�޹���</font>��������Ӧ�û����ڸð����Ȩ�ޡ�<BR>
2��<font color=blue>�������ӵ�����̳����</font>��ֻ��Ҫ������������ѡ����Ϊ��̳���༴�ɣ�<font color=blue>�������ӵ�����̳����</font>����Ҫ������������ȷ����ѡ�����̳������ϼ����档
</td>
</tr>
<tr> 
<td width="40%" height="30">��̳����</td>
<td width="60%"> 
<input type="text" name="boardtype" size="35">&nbsp;&nbsp;��� (<a title="��д��" href="javascript:;" onclick="DvWnd.open('��д��','plus_popwan_boardtinput.asp?sh=520&amp;sw=500',800,520,1,{bgc:'black',opa:0.5});">��д��</a>) �ɿ�����д����
</td>
</tr>
<tr> 
<td width="40%" height="24">����˵��<BR>����ʹ��HTML����</td>
<td width="60%"> 
<textarea name="Readme" cols="50" rows="5"></textarea>
</td>
</tr>
<tr> 
<td width="40%" height="24">�������<BR>����ʹ��HTML����</td>
<td width="60%"> 
<textarea name="Rules" cols="50" rows="5"></textarea>
</td>
</tr>
<tr> 
<td width="40%" height="30"><U>�������</U></td>
<td width="60%"> 
<select name="class" id="Boardid"></select>
<SCRIPT LANGUAGE="JavaScript">
<!--
BoardJumpListSelect('<%=Dvbbs.CheckNumeric(Request("editid"))%>','Boardid','��Ϊ��̳����','0','0');
//-->
</SCRIPT>
</td>
</tr>
<tr> 
<td width="40%" height="30"><U>ʹ����ʽ���</U><BR>�����ʽ����а�����̳��ɫ��ͼƬ<BR>����Ϣ</td>
<td width="60%">
<%
	set rs_c=dvbbs.execute("select forum_sid,forum_cid from dv_setup")
	Forum_sid=rs_c(0)
'	Forum_cid=rs_c(1)
	rs_c.close:Set rs_c=Nothing
%>
<Select Size="1" Name="sid">
<%
Dim Templateslist
For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
	Response.Write "<Option value="""& Templateslist.selectSingleNode("@id").text &""""
	If Forum_sid = CLng(Templateslist.selectSingleNode("@id").text) Then Response.Write " selected "
	Response.Write ">"& Templateslist.selectSingleNode("@type").text &" </Option>"
Next
%>
</td>
</tr>
<tr> 
<td width="40%" height="30"><U>��̳����</U><BR>������������|�ָ����磺ɳ̲С��|wodeail</td>
<td width="60%"> 
<input type="text" name="boardmaster" size="35">
</td>
</tr>
<tr> 
<td width="40%" height="30"><U>��ҳ��ʾ��̳ͼƬ</U><BR>��������ҳ��̳����������<BR>��ֱ����дͼƬURL</td>
<td width="60%">
<input type="text" name="indexIMG" size="35">
</td>
</tr>
<!-- URL�ⲿ���� ��ʼ -->
<tr> 
<td width="40%" height="30"><U>URL�ⲿ����</U><BR>��д�����ݺ�����̳�б����˰��潫�Զ��л�������ַ<BR>����дURL����·��</td>
<td width="60%">
<input type="radio" class="radio" name="tempradio" value="0" checked onclick="javascript:document.theform.Board_Setting_50.value='0';">�ر�&nbsp;
<input type="radio" class="radio" name="tempradio" value="1" >ʹ��URL�ⲿ����&nbsp;
<br/><input type=text name="Board_Setting_50" value="0" size="50">
</td>
</tr>
<!-- URL�ⲿ���� ���� -->
<tr> 
<td width="40%" height="24">&nbsp;</td>
<td width="60%"> 
<input type="submit" name="Submit" value="�����̳" class="button">
</td>
</tr>
</form>
</table>
<%
set rs_c=nothing
set rs=nothing
end sub

sub edit()
dim rs_c,reBoard_Setting,rs,sql
Dim forum_sid,forum_cid,Style_Option,TempOption
Dim iCssName,iCssID,iStyleName,i
sql = "select * from dv_board order by rootid,orders"
set rs_c=Dvbbs.Execute(sql)
sql = "select * from dv_board where boardid="&Dvbbs.CheckNumeric(request("editid"))
set rs=Dvbbs.Execute(sql)
reBoard_Setting=split(rs("Board_setting"),",")

forum_sid=rs("sid")
'forum_cid=rs("cid")
%>
<table cellspacing="0" cellpadding="0" class="pw_tb1">
<form action ="?action=savedit" method="post" name="theform">
<input type="hidden" name="editid" value="<%=Request("editid")%>">
<tr> 
<th colspan="2" style="text-align:center;">�༭��Ϸר�⣺<%=rs("boardtype")%></th>
</tr>
<tr> 
<td width="100%" height="30" colspan="2">
˵����<BR>1�������̳�������ص����þ�ΪĬ�����ã��뷵����̳���������ҳ�����б�ĸ߼����������ø���̳����Ӧ���ԣ��������Ը���̳���������Ȩ�����ã��뵽<font color=blue>��̨��̳Ȩ�޹���</font>��������Ӧ�û����ڸð����Ȩ�ޡ�<BR>
2��<font color=blue>�������ӵ�����̳����</font>��ֻ��Ҫ������������ѡ����Ϊ��̳���༴�ɣ�<font color=blue>�������ӵ�����̳����</font>����Ҫ������������ȷ����ѡ�����̳������ϼ����档
</td>
</tr>
<tr> 
<td width="40%" height="30">��̳����</td>
<td width="60%"> 
<input type="text" name="boardtype" size="35" value="<%=Server.htmlencode(rs("boardtype"))%>" >
</td>
</tr>
<tr> 
<td width="40%" height="24">����˵��<BR>����ʹ��HTML����</td>
<td width="60%"> 
<textarea name="Readme" cols="50" rows="5"><%=server.HTMLEncode(Rs("readme")&"")%></textarea>
</td>
</tr>
<tr> 
<td width="40%" height="24">�������<BR>����ʹ��HTML����</td>
<td width="60%"> 
<textarea name="Rules" cols="50" rows="5"><%=server.HTMLEncode(Rs("Rules")&"")%></textarea>
</td>
</tr>
<tr> 
<td width="40%" height="30"><U>�������</U><BR>������̳����ָ��Ϊ��ǰ����<BR>������̳����ָ��Ϊ��ǰ�����������̳</td>
<td width="60%"> 
<select name="class">
<option value="0">��Ϊ��̳����</option>
<% do while not rs_c.EOF%>
<option value="<%=rs_c("boardid")%>" <% if cint(rs("parentid")) = rs_c("boardid") then%> selected <%end if%>>
<%if rs_c("depth")>0 then%>
<%for i=1 to rs_c("depth")%>&nbsp;&nbsp;|<%next%>
<%end if%>&nbsp;��&nbsp;<%=rs_c("boardtype")%></option>
<%
rs_c.MoveNext 
loop
rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="40%" height="30"><U>ʹ����ʽ���</U><BR>�����ʽ����а�����̳��ɫ��ͼƬ<BR>����Ϣ</td>
<td width="60%">
<%
	set rs_c=dvbbs.execute("select forum_sid,forum_cid from dv_setup")
	Forum_sid=rs_c(0)
	Forum_cid=rs_c(1)
	rs_c.close:Set rs_c=Nothing
%>
<Select Size="1" Name="sid">
<%
Dim Templateslist
For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
	Response.Write "<Option value="""& Templateslist.selectSingleNode("@id").text &""""
	If rs("sid") = CLng(Templateslist.selectSingleNode("@id").text) Then Response.Write " selected "
	Response.Write ">"& Templateslist.selectSingleNode("@type").text &" </Option>"
Next
%>
</td>
</tr>
<tr> 
<td width="40%" height="30"><U>��̳����</U><BR>������������|�ָ����磺ɳ̲С��|wodeail</td>
<td width="60%"> 
<input type="text" name="boardmaster" size="35" value='<%=rs("boardmaster")%>'>
<input type="hidden" name="oldboardmaster" value='<%=rs("boardmaster")%>'>
</td>
</tr>
<tr> 
<td width="40%" height="30"><U>��ҳ��ʾ��̳ͼƬ</U><BR>��������ҳ��̳����������<BR>��ֱ����дͼƬURL</td>
<td width="60%">
<input type="text" name="indexIMG" size="35" value="<%=rs("indexIMG")%>">
</td>
</tr>
<!-- URL�ⲿ���� ��ʼ -->
<tr> 
<td width="40%" height="30"><U>URL�ⲿ����</U><BR>��д�����ݺ�����̳�б����˰��潫�Զ��л�������ַ<BR>����дURL����·��</td>
<td width="60%">
<input type="radio" class="radio" name="tempradio" value="0" onclick="javascript:document.theform.Board_Setting_50.value='0';">�ر�&nbsp;
<input type="radio" class="radio" name="tempradio" value="1" >ʹ��URL�ⲿ����&nbsp;
<br/><input type=text name="Board_Setting_50" value="<%=reBoard_Setting(50)%>" size=50>
<script type="text/javascript" language="javascript">
<!--
	var a = document.theform.Board_Setting_50.value;
	if (a == '0'){
		document.theform.tempradio[0].checked=true;
	}else{
		document.theform.tempradio[1].checked=true;
	}
//-->
</script>
</td>
</tr>
<!-- URL�ⲿ���� ���� -->
<tr> 
<td width="40%" height="24">&nbsp;</td>
<td width="60%"> 
<input type="submit" name="Submit" value="�ύ�޸�" class="button">
</td>
</tr>
<tr> 
<td width="100%" height="30" colspan="2" align="right">
<a href="?action=add&editid=<%=Request("editid")%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>��Ӱ���</U></font></a> | <a href="?action=edit&editid=<%=Request("editid")%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>��������</U></font></a> | <%if rs("child")=0 then%><a href="?action=del&editid=<%=Request("editid")%>" onclick="{if(confirm('ɾ������������̳���������ӣ�ȷ��ɾ����?')){return true;}return false;}"><font color="<%=Dvbbs.mainsetting(3)%>"><U>ɾ��</U></a><%else%><a href="#" onclick="{if(confirm('����̳����������̳��������ɾ����������̳����ɾ������̳��')){return true;}return false;}"><font color="<%=Dvbbs.mainsetting(3)%>"><U>ɾ��</U></a><%end if%>
</td>
</tr>
</form>
</table>
<%
rs.close
set rs=nothing
set rs_c=nothing
end sub

'���������̳��Ϣ
Sub savenew()
	If request("boardtype")="" Then
		Errmsg=Errmsg+"<li>��������̳���ơ�"
		founderr=true
	End If
	If request("class")="" Then
		Errmsg=Errmsg+"<li>��ѡ����̳���ࡣ"
		founderr=true
	End If
	If request("readme")="" Then
		Errmsg=Errmsg+"<li>��������̳˵����"
		founderr=true
	End If
	If founderr=true Then
		Response.Redirect "showerr.asp?ErrCodes="& Errmsg &"&action=OtherErr"
		exit sub
	End If
	Dim boardid,rootid,parentid,depth,orders,Fboardmaster,maxrootid,parentstr,rs,SQL
	If request("class")<>"0" Then
		Set rs=Dvbbs.Execute("select rootid,boardid,depth,orders,boardmaster,ParentStr from dv_board where boardid="&Dvbbs.CheckNumeric(request("class")))
		rootid=rs(0)
		parentid=rs(1)
		depth=rs(2)
		orders=rs(3)
		If depth+1>20 Then
			Errmsg="����̳�������ֻ����20������"
		  Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
		  Exit Sub
		 End If 
		parentstr=rs(5)
	Else
		Set rs=Dvbbs.Execute("select max(rootid) from dv_board")
	  maxrootid=rs(0)+1
		If IsNull(MaxRootID) Then MaxRootID=1
	End If
	sql="select boardid from dv_board where boardid="&Dvbbs.CheckNumeric(request("newboardid"))
	Set rs=Dvbbs.Execute(sql)
	If not (rs.eof and rs.bof) then
		Errmsg="������ָ���ͱ����̳һ������š�"
		Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
		exit sub
	Else
		boardid=request("newboardid")
	End If
	Dim trs,forumuser,setting,iBoard_Setting
	Set trs=Dvbbs.Execute("select * from dv_setup")
	Setting=Split(trs("Forum_Setting"),"|||")
	forumuser=Setting(2)
	If Request.form("Board_Setting_50")="" Then
		iBoard_Setting = "0"
	Else
		iBoard_Setting = Dvbbs.CheckStr(Replace(Request.form("Board_Setting_50"),",",""))
	End If
	set rs = server.CreateObject ("adodb.recordset")
	sql = "select * from dv_board"
	If Not IsObject(Conn) Then ConnectionDatabase
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
		rs("Board_Setting")="0,0,0,0,0,1,1,1,1,1,1,1,1,1,1,1,16240,3,0,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,0,1,100,20,10,9,normal,1,10,10,0,0,0,0,1,0,0,1,4,0,0,0,200,0,0,,$$,"&iBoard_Setting&",0,0,1,0|0|0|0|0|0|0|0|0,0|0|0|0|0|0|0|0|0,0,0,0,0,0,0,0,0,0,��ˮ|���|����|�ͷ�|������|���ݲ���|�ظ�����,0,1,0,24,0,0"
		rs("sid")=Dvbbs.CheckNumeric(request.form("sid"))
'		rs("cid")=Dvbbs.CheckNumeric(request.form("cid"))
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
	Dvbbs.dvbbs_suc("��̳��ӳɹ���" & str)
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
		Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
		exit sub
	end if
	dim newboardid,maxrootid,readme,Rules
	dim parentid,boardmaster,depth,child,ParentStr,rootid,iparentid,iParentStr,iBoard_Setting,temp_iBoard_Setting
	dim trs,brs,mrs
	Dim iii,rs,sql
	set rs = server.CreateObject ("adodb.recordset")
	sql = "select * from dv_board where boardid="&request("editid")
	If Not IsObject(Conn) Then ConnectionDatabase
	rs.Open sql,conn,1,3
	newboardid=rs("boardid")
	parentid=rs("parentid")
	iparentid=rs("parentid")
	boardmaster=rs("boardmaster")
	ParentStr=rs("ParentStr")
	depth=rs("depth")
	child=rs("child")
	rootid=rs("rootid")
	temp_iBoard_Setting = Split(Rs("board_setting"),",")
	If Request.form("Board_Setting_50")="" Then
		temp_iBoard_Setting(50) = "0"
	Else
		temp_iBoard_Setting(50) = Dvbbs.CheckStr(Replace(Request.form("Board_Setting_50"),",",""))
	End If
	For iii=0 To 71
		If iii=0 Then
			iBoard_Setting = temp_iBoard_Setting(iii)
		Else
			iBoard_Setting = iBoard_Setting & "," & temp_iBoard_Setting(iii)
		End If
	Next
	'�ж���ָ������̳�Ƿ���������̳
	if ParentID=0 then
		if clng(request("class"))<>0 then
		set trs=Dvbbs.Execute("select rootid from dv_board where boardid="&request("class"))
		if rootid=trs(0) then
			errmsg="������ָ���ð����������̳��Ϊ������̳1"
			Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
			exit sub
		end if
		end if
	else
		set trs=Dvbbs.Execute("select boardid from dv_board where ParentStr like '%"&ParentStr&","&newboardid&"%' and boardid="&request("class"))
		if not (trs.eof and trs.bof) then
			errmsg="������ָ���ð����������̳��Ϊ������̳"
			Response.Redirect "showerr.asp?ErrCodes=<li>"& Errmsg &"&action=OtherErr"
			exit sub
		end if
	end if

	If parentid=0 then
		parentid=rs("boardid")
		iparentid=0
	Else
		Set mrs=Dvbbs.Execute("select max(rootid) from dv_board")
			Maxrootid=mrs(0)+1
		mrs.close:Set mrs=Nothing
		rs("rootid")=Maxrootid
	End if
	rs("boardtype") = Request.Form("boardtype")	'ȡ��JS���ˡ�
	rs("parentid") = Request.Form("class")
	rs("boardmaster") = Request("boardmaster")
	rs("readme") = Request("readme")
	rs("Rules") = Request.form("Rules")
	Rs("Board_Setting") = iBoard_Setting
	rs("indexIMG")=request.form("indexIMG")
	rs("sid")=Cint(request.form("sid"))
'	rs("cid")=Cint(request.form("cid"))
	rs.Update 
	rs.Close:set rs=nothing
	if request("oldboardmaster")<>Request("boardmaster") then call addmaster(Request("boardmaster"),request("oldboardmaster"),1)
	
	Dvbbs.dvbbs_suc("��̳�޸ĳɹ���<br>" & str)
	CheckAndFixBoard 0,1
	Boardchild()
	RestoreBoardCache()
End sub

'ɾ�����棬ɾ���������ӣ���ڣ�����ID
Sub Del()
	Dim Trs,EditId
	EditId = Dvbbs.CheckNumeric(Request("editid"))
	'�������ϼ�������̳�����������̳�����¼���̳������ɾ��
	Set tRs = Dvbbs.Execute("SELECT RootID FROM Dv_Board WHERE BoardID = " & EditId)
	Dim UpdateRootID,Rs,sql,i
	UpdateRootID = tRs(0)
	Set Rs = Dvbbs.Execute("SELECT ParentStr, Child, Depth FROM Dv_Board WHERE BoardID = " &  EditId)
	If Not (Rs.Eof And Rs.Bof) Then
		If Rs(1) > 0 Then
			Response.Write "����̳����������̳����ɾ����������̳���ٽ���ɾ������̳�Ĳ���"
			Exit Sub
		End If
		'������ϼ����棬���������
		If Rs(2) > 0 Then
			Dvbbs.Execute("UPDATE Dv_Board SET Child = Child - 1 WHERE BoardID IN (" & Rs(0) & ")")
		End If
		Sql = "DELETE FROM Dv_Board WHERE Boardid = " & EditId
		Dvbbs.Execute(Sql)
		For i = 0 To Ubound(AllPostTable)
			Sql = "DELETE FROM " & AllPostTable(i) & " WHERE BoardID = " & EditId
			Dvbbs.Execute(Sql)
		Next
		Dvbbs.Execute("DELETE FROM Dv_Topic WHERE BoardID = " & EditId)
		Dvbbs.Execute("DELETE FROM Dv_BestTopic WHERE BoardID = " & EditId)
		Dvbbs.Execute("DELETE FROM Dv_Upfile WHERE F_BoardID = " & EditId)
		Dvbbs.Execute("DELETE FROM Dv_Appraise WHERE BoardID = " & EditId)
		'ɾ����ɾ����̳���Զ����û�Ȩ�� 2004-11-15 Dv.Yz
		Dvbbs.Execute("DELETE FROM Dv_UserAccess WHERE NOT Uc_BoardID IN (SELECT BoardID FROM Dv_Board)")
	End If
	Set Rs = Nothing
	CheckAndFixBoard 0,1
	RestoreBoardCache()
	Dvbbs.Dvbbs_suc("��̳ɾ���ɹ���")
End Sub

Sub Addmaster(s,o,n)
	Dim Arr, Pw, Oarr
	Dim Classname, Titlepic, Rs, Sql, i
	Set Rs = Dvbbs.Execute("SELECT Usertitle, GroupPic FROM Dv_UserGroups WHERE Usergroupid = 3")
	If Not (Rs.Eof And Rs.Bof) Then
		Classname = Rs(0)
		Titlepic = Rs(1)
	End If
	Randomize
	Pw = CInt(Rnd * 9000) + CInt(Rnd * 9000) + 100000
	Arr = Split(s,"|")
	Oarr = Split(o,"|")
	Set Rs = Server.Createobject("Adodb.Recordset")
	For i = 0 To Ubound(Arr)
		Sql = "SELECT * FROM [Dv_User] WHERE Username = '" & Arr(i) & "'"
		If Not IsObject(Conn) Then ConnectionDatabase
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
			'�������ϸ����ʹ��¼����ʾ���ϲ������
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
			Str = Str & "������������û���<b>" & Arr(i) & "</b> ���룺<b>" & Pw & "</b><br><br>"
			Dvbbs.Execute("UPDATE Dv_Setup SET Forum_Usernum = Forum_Usernum + 1, Forum_Lastuser = '" & Arr(i) & "'")
		Else
			'������Ӱ������ı�ȼ��Ĵ��� 2005-3-7 Dv.Yz
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

Sub RestoreBoardCache()	
	Dvbbs.Name="jsonboardlist0"
	Dvbbs.RemoveCache
	Dvbbs.Name="jsonboardlist1"
	Dvbbs.RemoveCache
	Dim Board
	Dvbbs.Name="0"
	Dvbbs.RemoveCache
	Dvbbs.LoadBoardList()
	'Dvbbs.Temp_LoadBoardList()
	For Each board in Application(Dvbbs.CacheName&"_boardlist").documentElement.selectNodes("board/@boardid")
		Dvbbs.LoadBoardData board.text
		Dvbbs.LoadBoardinformation board.text
	Next
	If Request("action")="RestoreBoardCache" Then Dvbbs.Dvbbs_suc("�ؽ����а��滺��ɹ���")
End Sub
Sub CheckAndFixBoard(ParentID,orders)
	Dim Rs,SQL,Child,ParentStr,i
	If ParentID=0 Then
		Dvbbs.Execute("update dv_board set Depth=0,ParentStr='0' where ParentID=0")
	End If
	Set Rs=Dvbbs.Execute("Select BoardID,rootid,ParentStr,Depth From Dv_Board where ParentID="&ParentID&" order by Rootid,orders")
	If Not Rs.Eof Then
		SQL = Rs.GetRows(-1)
		Rs.Close:Set Rs=Nothing
		For i=0 To Ubound(SQL,2)
			If SQL(2,i)<>"0" Then
				ParentStr=SQL(2,i)&","&SQL(0,i)
			Else
				ParentStr=SQL(0,i)
			End If
			If Not IsObject(Conn) Then ConnectionDatabase
			Conn.Execute "update dv_board set Depth="&SQL(3,i)+1&",ParentStr='"&ParentStr&"',rootid="&SQL(1,i)&" where ParentID="&SQL(0,i)&"",Child
			Dvbbs.Execute("update dv_board set Child="&Child&",orders="&orders&" Where BoardID="&SQL(0,i)&"")
			orders=orders+1
			CheckAndFixBoard SQL(0,i),orders
		Next
	End If
End Sub

Rem ͳ��������̳���� 2004-5-3 Dvbbs.YangZheng
Sub Boardchild()
	Dim cBoardNum, cBoardid
	Dim Trs,rs,Sql
	Dim Bn,i
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


Sub AllPostTable1()
	Dim Trs
	Set Trs=Dvbbs.Execute("select * from [Dv_TableList]")
	AllPostTable=""
	Do While Not TRs.EOF
		If AllPostTable=""  Then 
			AllPostTable=TRs("TableName")
			AllPostTableName=TRs("TableType")
		Else
			AllPostTable=AllPostTable&"|"&TRs("TableName")
			AllPostTableName=AllPostTableName&"|"&TRs("TableType")
		End If
	TRs.MoveNext
	Loop 
	Trs.Close
End Sub
'=============
'������� ����
'=============
%>