<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<%
Head()
Dim admin_flag,Act
admin_flag=",45,"
CheckAdmin(admin_flag)
Act = Request.QueryString("action")
Top_Nav()
Call main()
Footer()

Sub Top_Nav()
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<th width="100%" colspan="2">��̳���˿ռ�ϵͳ���� - - <a href="?action=skins">���ģ�����</a>
</th>
</tr>
<!-- <tr>
<td height="23" colspan="2" class="td1">&nbsp;
</td>
</tr>
<tr><td height="23" colspan="2" class="td1">
<button class="button" onclick="window.location='?action=skins'">���ģ�����</button>
</td>
</tr> -->
</table>
<br/>
<%
End Sub

Sub main()
	Select Case Act
		Case "skins"
			Skin_admin()
		Case "add_skin"
			SaveSkin_admin()
		Case "edit_skin"
			Edit_Skin()
		Case "saveedit_skin"
			Saveedit_skin()
		Case "delskins"
			DelSkins()
		Case Else
			Skin_admin()
	End Select
End Sub

Sub Addgroup()
'\Skins\myspace
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
	<form method="POST" action="?action=add_skin">
	<tr> 
	<th colspan=2><b>���ӷ��ģ��</b>
	</th>
	</tr>
	<tr> 
	<td width="20%" class=td1><b>�������</b></td>
	<td width="80%" class=td1>
		<input type="text" name="skinname" size="35" value="Ĭ�Ϸ��">
	</td>
	</tr>
	<tr> 
	<td class=td1><b>�ṩ��</b></td>
	<td class=td1>
		<input type="text" name="skinusername" value="ϵͳĬ��" size="35">
	</td>
	</tr>

	<tr>
	<td class=td1><b>���Ŀ¼</b></td>
	<td class=td1>
		<input type="text" name="skinpath" value="" size="35">
		<br/>���뽫�����ļ��д�����̳/Skins/myspace/Ŀ¼�£�������д��Ӧ���ļ�Ŀ¼���ƣ�
	</td>
	</tr>
	<tr><td height="23" colspan="2" class=td1><input type="submit" name="Submit" value="�� ��" class=button></tr>
	</form>
</table>
<%
End Sub

Sub Edit_Skin()
	Dim SID,CssText
	Sid = Dvbbs.CheckNumeric(Request("id"))
	If Sid=0 Then
		Errmsg=ErrMsg + "<BR/><li>��������ȷ��"
		dvbbs_error()
		exit sub
	End If
	Dim Rs,Sql
	Sql = "Select id,s_name,s_username,s_userid,s_style,s_path,s_lock,s_addtime,s_css From Dv_Space_skin where id="&Sid
	Set Rs = Dvbbs.Execute(Sql)
	If Rs.Eof Then
		Errmsg=ErrMsg + "<BR/><li>���ҵ����ݲ����ڡ�"
		dvbbs_error()
		exit sub
	Else
		CssText = Rs("s_css")
		If Trim(Rs("s_css"))="" or IsNull(Rs("s_css")) Then
			CssText = GetFromFile("../skins/myspace/"&Rs(5)&"style.css")
		End If
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
	<form method="POST" action="?action=saveedit_skin" name="EditSkin">
	<input type="hidden" name="id" value="<%=Rs(0)%>"/>
	<tr> 
	<th colspan="2"><b>�༭���ģ��--<%=Rs(1)%></b>
	</th>
	</tr>
	<tr> 
	<td width="20%" class=td1><b>�������</b></td>
	<td width="80%" class=td1>
		<input type="text" name="skinname" size="35" value="<%=Rs(1)%>">
	</td>
	</tr>
	<tr> 
	<td class=td1><b>�ṩ��</b></td>
	<td class=td1>
		<input type="text" name="skinusername" value="<%=Rs(2)%>" size="35">
	</td>
	</tr>
	<tr> 
	<td class=td1><b>�ṩ��ID</b></td>
	<td class=td1>
		<input type="text" name="skinuserid" value="<%=Rs(3)%>" size="35">(ϵͳĬ��Ϊ��0��)
	</td>
	</tr>
	<tr>
	<td class=td1><b>���Ŀ¼</b></td>
	<td class=td1>
		<input type="text" name="skinpath" value="<%=Rs(5)%>" size="35">
		<br/>���뽫�����ļ��д�����̳/Skins/myspace/Ŀ¼�£�������д��Ӧ���ļ�Ŀ¼���ƣ�
	</td>
	</tr>
	<tr> 
	<td class=td1><b>״̬</b></td>
	<td class=td1>
		<input type="radio" class="radio" name="s_lock" value="0"/>��� <input type="radio" class="radio" name="s_lock" value="1"/>����
	</td>
	</tr>
	<tr> 
	<td class=td1 valign="top"><b>���CSS��ʽ</b></td>
	<td class=td1>
		<textarea name="s_css" id="s_css" style="width:98%;height:80px;"><%=CssText%></textarea>
		<img src="Skins/images/minus.gif" unselectable="on"  onclick="textarea_size(-200,'s_css');" />
		<img src="Skins/images/plus.gif" unselectable="on" onclick="textarea_size(200,'s_css');" />
	</td>
	</tr>

	<tr><td height="23" colspan="2" class=td1><input type="submit" name="Submit" value="�� ��" class=button></tr>
	</form>
</table>
<script language="JavaScript">
<!--
chkradio(document.EditSkin.s_lock,"<%=Rs(6)%>");
//-->
</script>
<%
End If
Rs.Close
Set Rs = Nothing
End Sub

Function GetFromFile(Filepath)
	On error resume Next
	GetFromFile = ""
	Dim FileName,Fso,FileText,ReadAllTextFile
	FileName = Server.MapPath(Filepath)
	Set FSO=Dvbbs.iCreateObject("Scripting.FileSystemObject")
	If Err Then
		err.Clear
		Errmsg=ErrMsg + "<br /><li>���ķ�������֧��д�ļ�,CSS�ļ�д��ʧ��,���ֹ�������������ļ����������!</li>"
		Dvbbs_error()
		Exit Function
	End If
	If Not Fso.FileExists(FileName) Then
		Errmsg=ErrMsg + "<br /><li>ϵͳ�Ҳ�����ʽ�ļ���<a href="""&Filepath&""" target=_blank>"&Filepath&"</a>!</li>"
		Dvbbs_error()
		Exit Function
	Else
		Set FileText = Fso.OpenTextFile(FileName, 1)
		GetFromFile = FileText.ReadAll
		If GetFromFile = "" Then
			Errmsg=ErrMsg + "<br /><li>��ʽ�ļ���<a href="""&Filepath&""" target=_blank>"&Filepath&"</a> ����Ϊ�գ�����¸��ļ�!</li>"
			Dvbbs_error()
			Exit Function
		End If
	End If
	Set Fso = Nothing
End Function

Sub WriteFile(Filepath,Text)
	On error resume Next
	Dim FileName,Fso,FileText,ReadAllTextFile
	FileName = Server.MapPath(Filepath)
	Set FSO=Dvbbs.iCreateObject("Scripting.FileSystemObject")
	If Err Then
		err.Clear
		Errmsg=ErrMsg + "<br /><li>���ķ�������֧��д�ļ�,CSS�ļ�д��ʧ��,���ֹ�������������ļ����������!</li>"
		Dvbbs_error()
		Exit Sub
	End If
	If Not Fso.FileExists(FileName) Then
		Errmsg=ErrMsg + "<br /><li>ϵͳ�Ҳ�����ʽ�ļ���<a href="""&Filepath&""" target=_blank>"&Filepath&"</a>!</li>"
		Dvbbs_error()
		Exit Sub
	Else
		Fso.CreateTextFile(FileName).WriteLine(Text)
		If  Err Then
			err.Clear
			Errmsg=ErrMsg + "<br /><li>���ķ�������֧��д�ļ�,CSS�ļ�д��ʧ��,���ֹ�������������ļ����������!</li>"
			Dvbbs_error()
			Exit Sub
		End If
	End If
	Set Fso = Nothing
End Sub



'�޸ķ��
Sub Saveedit_skin()
	Dim skinname,skinpath,skinusername,sid,skinuserid
	Dim Rs,Sql
	Sid = Dvbbs.CheckNumeric(Request("id"))
	If Sid=0 Then
		Errmsg=ErrMsg + "<BR/><li>��������ȷ��"
		dvbbs_error()
		exit sub
	End If
	skinname = Dvbbs.Checkstr(trim(request.form("skinname")))
	skinpath = Dvbbs.Checkstr(trim(request.form("skinpath")))
	skinusername = Dvbbs.Checkstr(trim(request.form("skinusername")))
	skinuserid = Dvbbs.CheckNumeric(skinuserid)
	If skinname = "" Then
		Errmsg=ErrMsg + "<BR><li>������Ʋ���Ϊ�ա�"
		dvbbs_error()
		exit sub
	End If
	If skinpath = "" Then
		Errmsg=ErrMsg + "<BR><li>���Ŀ¼����Ϊ�ա�"
		dvbbs_error()
		exit sub
	Else
		If Right(skinpath,1)<>"/" Then
			skinpath = skinpath & "/"
		End If
	End If

	Set Rs = Dvbbs.Execute("select top 1 id from Dv_Space_skin where id<>"&sid&" and s_name='"&skinname&"'")
	if not rs.eof  then
		Errmsg=ErrMsg + "<BR><li>�÷�������Ѵ��ڣ��벻Ҫ�ظ����ӡ�"
		dvbbs_error()
		exit sub
	end if

	Sql = "select top 1 s_name,s_username,s_userid,s_style,s_path,s_lock,s_css from Dv_Space_skin where id="&sid
	Set Rs=Dvbbs.iCreateObject("Adodb.RecordSet")
	Rs.Open Sql,Conn,1,3
	If Rs.Eof Then
		Errmsg=ErrMsg + "<BR/><li>���ҵ����ݲ����ڡ�"
		dvbbs_error()
		exit sub
	Else
		Rs(0) = skinname
		Rs(1) = skinusername
		Rs(2) = skinuserid
		Rs(4) = skinpath
		Rs(5) = Dvbbs.CheckNumeric(request.form("s_lock"))
		Rs(6) = request.form("s_css")
		Rs.update
	End If
	Rs.Close
	Set Rs = Nothing
	WriteFile "../skins/myspace/"&skinpath&"style.css",Request.form("s_css")
	Dv_suc("<b>�����޸ĳɹ���</b>")
End Sub

'���ӷ��
Sub SaveSkin_admin()
	Dim skinname,skinpath,skinusername
	Dim Rs,Sql
	skinname = Dvbbs.Checkstr(trim(request.form("skinname")))
	skinpath = Dvbbs.Checkstr(trim(request.form("skinpath")))
	skinusername = Dvbbs.Checkstr(trim(request.form("skinusername")))
	If skinname = "" Then
		Errmsg=ErrMsg + "<BR><li>������Ʋ���Ϊ�ա�"
		dvbbs_error()
		exit sub
	End If

	If skinpath = "" Then
		Errmsg=ErrMsg + "<BR><li>���Ŀ¼����Ϊ�ա�"
		dvbbs_error()
		exit sub
	Else
		If Right(skinpath,1)<>"/" Then
			skinpath = skinpath & "/"
		End If
	End If

	Set Rs = Dvbbs.Execute("select top 1 id from Dv_Space_skin where s_name='"&skinname&"'")
	if rs.eof and rs.bof then
		Dvbbs.Execute("insert into Dv_Space_skin (s_name,s_username,s_userid,s_style,s_path,s_lock) values ('"&skinname&"','"&skinusername&"',0,1,'"&skinpath&"',1)")
	else
		Errmsg=ErrMsg + "<BR><li>�÷�������Ѵ��ڣ��벻Ҫ�ظ����ӡ�"
		dvbbs_error()
		exit sub
	end if
	set rs=nothing
	Dv_suc("<b>���ӳɹ���</b>")
End Sub

Sub Skin_admin()
	Addgroup()
	if Request("react")="lock" Then
		Dim Val
		Val = Dvbbs.CheckNumeric(Request("v"))
		Dvbbs.Execute("update Dv_Space_skin set s_lock="&Val&" where id="&Dvbbs.CheckNumeric(Request("id")))

	End If


	Dim Rs,Sql
	%>
	<script language="JavaScript" src="../inc/Pagination.js"></script>
	<br/>
	<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
	<form name="theform" method="post" action="?action=delskins">
	<tr>
	<th height="23" colspan="7" ><b>���Կռ������</b></th>
	</tr>
	<tr>
	<td class="td2" colspan="7" >
		<ol>
		<li>��Ա����ѡȡ����״̬�ķ���ģ�壻</li>
		<li>�������ģ��״̬�����л���</li>
		<li>�������������������ʽ��</li>
		</ol>
	</td>
	</tr>
	<tr align="center">
	<td width="10%"><b>���ID</b></td>
	<td width="10%"><b>״̬</b></td>
	<td width="20%"><b>�������</b></td>
	<td width="10%"><b>�ṩ��</b></td>
	<td width="20%"><b>���Ŀ¼</b></td>
	<td width="10%"><b>¼��ʱ��</b></td>
	<td width="20%"><b>����</b></td>
	</tr>
	<%
		Sql = "Select id,s_name,s_username,s_userid,s_style,s_path,s_lock,s_addtime From Dv_Space_skin order by id desc"
		's_lock
		'0=��ˣ�1=����
		Dim Page,MaxRows,Endpage,CountNum,PageSearch,SqlString,i
		PageSearch = "?action=skins"
		Endpage = 0
		MaxRows = 20
		CountNum = 0
		Page = Request("Page")
		If IsNumeric(Page) = 0 or Page="" Then Page=1
		Page = Clng(Page)
		Set Rs = Dvbbs.iCreateObject ("adodb.recordset")
		If Not IsObject(Conn) Then ConnectionDatabase
		Rs.Open Sql,Conn,1,1
		If Not Rs.eof Then
			CountNum = Rs.RecordCount
			If CountNum Mod MaxRows=0 Then
				Endpage = CountNum \ MaxRows
			Else
				Endpage = CountNum \ MaxRows+1
			End If
			Rs.MoveFirst
			If Page > Endpage Then Page = Endpage
			If Page < 1 Then Page = 1
			If Page >1 Then 				
				Rs.Move (Page-1) * MaxRows
			End if
			SQL=Rs.GetRows(MaxRows)
		Else
			Response.Write "<tr><td class=""td1"" colspan=""7"" align=""center"">��ʱû���κη�����ݣ�</td></tr>"
		End If
		Rs.close:Set Rs = Nothing
		
		If IsArray(Sql) Then
		For i=0 To Ubound(SQL,2)
			Response.Write "<tr>"
			Response.Write "<td class=""td1"" align=center>"&Sql(0,i)&"</td>"
			Response.Write "<td class=""td1"" align=center>"
			If Sql(6,i)=1 Then
				Response.Write "<a href=""?action=skins&react=lock&v=0&id="&Sql(0,i)&"&page="&page&""" title=""��Ϊ���״̬"">����</a>"
			ElseIf Sql(6,i)=0 Then
				Response.Write "<a href=""?action=skins&react=lock&v=1&id="&Sql(0,i)&"&page="&page&""" title=""��Ϊ����״̬""><font color='red'>���</font></a>"	
			End If
			
			
			Response.Write "</td>"
			Response.Write "<td class=""td1"" align=center><a href=""../skins/myspace/"&Sql(5,i)&"demo.htm"" title=""����÷����ʽ"" target=""_blank"">"&Sql(1,i)&"</a></td>"
			Response.Write "<td class=""td1"" align=center>"&Sql(2,i)&"</td>"
			Response.Write "<td class=""td1"">"&Sql(5,i)&"</td>"
			Response.Write "<td class=""td1"" align=center>"&Sql(7,i)&"</td>"
			Response.Write "<td class=""td1"" align=center><input type=checkbox class=checkbox name=did value="&Sql(0,i)&">ɾ�� | <a href=""?action=edit_skin&id="&Sql(0,i)&""">�༭</a></td>"
			Response.Write "</tr>"
			
		Next
		Response.Write "<tr><td class=td2 colspan=7>��ѡ��Ҫɾ���ķ��<input type=checkbox class=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">ȫѡ <input type=button name=act value=ɾ�� class='button' onclick=""if (confirm('��ȷ��ִ�еĲ�����?')){document.theform.submit()}; "">"
		Response.Write "</td></tr>"
		End If


		Response.Write "<tr><td class=""td1"" colspan=""7""><SCRIPT>PageList("&Page&",3,"&MaxRows&","&CountNum&","""&PageSearch&""",1);</SCRIPT></td></tr>"

	%>
	</form>
	</table>
	<%

End Sub

Sub DelSkins()
	Dim Sql,Rs,i,did
	If Dvbbs.Checkstr(request("did"))="" Then
		Errmsg=ErrMsg + "��ѡ��Ҫɾ���ķ��"
		dvbbs_error()
		exit Sub
	End If
	
	For i = 1 To Request("did").Count
		If isNumeric(Request("did")(i)) Then
			If did = "" Then
				did = Request("did")(i)
			Else
				did = did & "," & Request("did")(i)
			End If
		End If
	Next

	Set Rs = Dvbbs.Execute("Select s_path From Dv_Space_skin where id in ("&did&")")
	Do while not Rs.Eof
		If Rs(0)<>"" Then
			DelFolder("../skins/myspace/"&Rs(0))
		End If
	Rs.MoveNext
	loop
	Rs.Close
	Set Rs = Nothing
	Dvbbs.Execute("delete from Dv_Space_skin where id in ("&did&")")
	Dv_suc("<b>ɾ���ɹ���</b>")
End Sub

Sub DelFolder(Filepath)
	On error resume Next
	Dim FileName,Fso,FileText,ReadAllTextFile
	FileName = Server.MapPath(Filepath)
	Set FSO=Dvbbs.iCreateObject("Scripting.FileSystemObject")
	If Err Then
		err.Clear
		Response.Write  "<br /><li>���ķ�������֧��д�ļ�����,���ֹ�����ɾ��"&Filepath&"!</li>"
		Exit Sub 
	End If
	If Not Fso.FolderExists(FileName) Then
		Response.Write  "<br /><li>ϵͳ�Ҳ���Ҫɾ������ʽ�ļ��У�"&Filepath&"!</li>"
	Else
		Fso.DeleteFolder(FileName)
	End If
	Set Fso = Nothing
End Sub
%>