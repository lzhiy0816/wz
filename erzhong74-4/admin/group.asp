<!--#include file=../conn.asp-->
<!--#include file="inc/const.asp"-->
<!--#include file="../inc/GroupPermission.asp"-->
<%
Head()
Dim admin_flag
admin_flag=",17,"
CheckAdmin(admin_flag)
Main()
Footer()

Sub Main()
	Select Case Request("action")
	Case "editgroup"
		EditGroup()
	Case "saveusergroup"
		SaveUserGroup()
	Case "savesysgroup"
		SaveSysGroup()
	Case "delusergroup"
		DelUserGroup()
	Case "online"
		GroupOnline()
	Case "saveonline"
		SaveGroupOnline()
	Case Else
		UserGroup()
	End Select
End Sub

Sub UserGroup()
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" align=left>&nbsp;������ʾ</th>
</tr>
<tr align=left>
<td height="23" class="td1" style="LINE-HEIGHT: 140%">
<li>������̳�û����Ϊϵͳ�û��顢�����û��顢ע���û��顢�������û�����������
<li>ϵͳ�û���Ϊ���ù̶��û��飬�������ӣ�����̳����֮�ã�����������ģ���ɾ�����������̳�����쳣
<li>�����û��鲻���û��ȼ������������ͨ�������������һЩ����̳�����⹱�׻��������Ա
<li>�������û��鲻���û��ȼ�����������������û�<U>���������ж����ͬ�û����Ȩ��</U>��ͨ�������������һЩ����̳�����⹱�׻��������Ա
<li>ע���û��鼴Ϊ��ͳ���û��ȼ���ÿ����(�ȼ�)���趨��ͬ��Ȩ��
<li>Ĭ��Ȩ��Ϊ�����µ��û���ʱʹ������һЩ����õ�Ȩ�����ã�ͨ���������û����Ҫ�ٴζ�����Ȩ��
</td>
</tr>
<tr align=left>
<td height="23" class="td2" style="LINE-HEIGHT: 140%">
<B>��ݲ���</B>��<a href="#1">ϵͳ��</a> | <a href="#2">������</a> | <a href="#3">��������</a>
 | <a href="#4">Vip�û������</a>
| <a href="?action=editgroup&groupid=4">�༭Ĭ��������</a> | <a href="?action=online">����ͼ������</a>
</td>
</tr>
</table>
<BR>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" colspan="6">ע���û���(�ȼ�)����</th>
</tr>
<tr><td colspan=6 height=25 class="td1">
С��ʾ�����Ȩ�������Էֱ��趨ÿ��ע���û���(�ȼ�)�ֱ�ӵ�в�ͬ����̳Ȩ��
</td></tr>
<tr>
<td height="23" width="5%" class=bodytitle><B>��ID</B></td>
<td height="23" width="20%" class=bodytitle><B>�û���(�ȼ�)����</B></td>
<td width="15%" class=bodytitle><B>���ٷ���</B></td>
<td width="30%" class=bodytitle><B>��(�ȼ�)ͼƬ</B></td>
<td height="23" width="10%" class=bodytitle><B>�û���</B></td>
<td width="20%" class=bodytitle><B>����</B></td>
</tr>
<FORM METHOD=POST ACTION="?action=saveusergroup">
<%
Dim Trs,rs
Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where ParentGID=3 Order By MinArticle")
Do While Not Rs.Eof
%>
<input type=hidden value="<%=rs("UserGroupID")%>" name="usertitleid">
<tr>
<td class=td1 align=center><%=Rs("UserGroupID")%></td>
<td height="23" class=td1><input size=15 value="<%=rs("usertitle")%>" name="usertitle" type=text></td>
<td class=td1><input size=5 value="<%=rs("MinArticle")%>" name="minarticle" type=text></td>
<td class=td1><input size=15 value="<%=rs("grouppic")%>" name="titlepic" type=text><img src="../<%=Dvbbs.Forum_PicUrl%>star/<%=rs("grouppic")%>" border="0"></td>
<td class=td1>
<B><%
Set Trs=Dvbbs.Execute("Select Count(*) From [Dv_User] Where UserGroupID="&Rs("UserGroupID"))
Response.Write Trs(0)
%></B>
</td>
<td class=td1><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">�༭</a> | <a href="user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">�г��û�</a> | <a href="?action=delusergroup&id=<%=rs("UserGroupID")%>" onclick="{if(confirm('ɾ�����������Զ�����һ�����û��ĵȼ������Ҳ��ɻָ���ȷ����?')){return true;}return false;}">ɾ��</a></td>
</tr>
<%
Rs.MoveNext
Loop
Set Rs=Nothing
Set Trs=Nothing
%>
<input type=hidden value="0" name="usertitleid">
<tr>
<td class=td1 align=center><font color=blue>��</font></td>
<td height="23" class=td1><input size=15 value="" name="usertitle" type=text></td>
<td class=td1><input size=5 value="0" name="minarticle" type=text></td>
<td class=td1><input size=15 value="level0.gif" name="titlepic" type=text></td>
<td class=td1>
<B>0</B>
</td>
<td width="20%" class=td1>&nbsp;</td>
</tr>
<tr align=center>
<td colspan=6 height=25 class="td2">
<input type=submit class="button" name=submit value="�ύ����">
</td></tr>
</FORM>
</table>
<BR>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" colspan="6">ϵͳ�û������<a name="1"></a></th>
</tr>
<tr><td colspan=6 height=25 class="td1">
С��ʾ�����Ȩ�������Էֱ��趨ÿ��ϵͳ�û���ֱ�ӵ�в�ͬ����̳Ȩ�ޣ�ϵͳ��ͷ�κ�ͼ����ʾ��ǰ̨�û���Ϣ��
</td></tr>
<tr>
<td height="23" width="5%" class=bodytitle><B>��ID</B></td>
<td height="23" width="20%" class=bodytitle><B>ϵͳ��ͷ��</B></td>
<td width="15%" class=bodytitle><B>ϵͳ������</B></td>
<td height="23" width="30%" class=bodytitle><B>ϵͳ��ͼ��</B></td>
<td height="23" width="10%" class=bodytitle><B>�û���</B></td>
<td width="20%" class=bodytitle><B>����</B></td>
</tr>
<FORM METHOD=POST ACTION="?action=savesysgroup">
<input type=hidden value="1" name="ParentID">
<%
Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where ParentGID=1 Order By UserGroupID")
Do While Not Rs.Eof
%>
<input type=hidden value="<%=rs("UserGroupID")%>" name="usertitleid">
<input value="<%=rs("title")%>" name="title" type=hidden>
<input type=hidden value="<%=rs("IsSetting")%>" name="issetting">
<tr>
<td class=td1 align=center><%=Rs("UserGroupID")%></td>
<td height="23" class=td1><input size=15 value="<%=rs("usertitle")%>" name="usertitle" type=text></td>
<td class=td1><%=Rs("Title")%></td>
<td class=td1><input size=15 value="<%=rs("grouppic")%>" name="titlepic" type=text>
<img src="../<%=Dvbbs.Forum_PicUrl%>star/<%=rs("grouppic")%>" border="0">
</td>
<td class=td1>
<B><%
Set Trs=Dvbbs.Execute("Select Count(*) From [Dv_User] Where UserGroupID="&Rs("UserGroupID"))
Response.Write Trs(0)
%></B>
</td>
<td class=td1><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">�༭</a> | <a href="user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">�г��û�</a></td>
</tr>
<%
Rs.MoveNext
Loop
Set Rs=Nothing
Set Trs=Nothing
%>
<tr align=center>
<td colspan=6 height=25 class="td2">
<input type=submit class="button" name=submit value="�ύ����">
</td></tr>
</FORM>
</table>

<BR>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" colspan="6">�����û������<a name="2"></a></th>
</tr>
<tr><td colspan=6 height=25 class="td1">
С��ʾ�����Ȩ�������Էֱ��趨ÿ�������û���ֱ�ӵ�в�ͬ����̳Ȩ�ޣ�ͨ���������������̳�ϱȽ�������û�Ⱥ�壬������ͷ�κ�ͼ����ʾ��ǰ̨�û���Ϣ��
</td></tr>
<tr>
<td width="5%" class=bodytitle><B>��ID</B></td>
<td height="23" width="15%" class=bodytitle><B>������ͷ��</B></td>
<td width="15%" class=bodytitle><B>ϵͳ������</B></td>
<td width="30%" class=bodytitle><B>������ͼƬ</B></td>
<td height="23" width="10%" class=bodytitle><B>�û���</B></td>
<td width="20%" class=bodytitle><B>����</B></td>
</tr>
<FORM METHOD=POST ACTION="?action=savesysgroup">
<input type=hidden value="2" name="ParentID">
<%
Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where ParentGID=2 Order By UserGroupID")
Do While Not Rs.Eof
%>
<input type=hidden value="<%=rs("UserGroupID")%>" name="usertitleid">
<input type=hidden value="<%=rs("IsSetting")%>" name="issetting">
<tr>
<td class=td1 align=center><%=Rs("UserGroupID")%></td>
<td height="23" class=td1><input size=15 value="<%=rs("usertitle")%>" name="usertitle" type=text></td>
<td class=td1><input size=15 value="<%=rs("title")%>" name="title" type=text></td>
<td class=td1><input size=15 value="<%=rs("grouppic")%>" name="titlepic" type=text>
<img src="../<%=Dvbbs.Forum_PicUrl%>star/<%=rs("grouppic")%>" border="0">
</td>
<td class=td1>
<B><%
Set Trs=Dvbbs.Execute("Select Count(*) From [Dv_User] Where UserGroupID="&Rs("UserGroupID"))
Response.Write Trs(0)
%></B>
</td>
<td class=td1><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">�༭</a> | <a href="user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">�г��û�</a> | <a href="?action=delusergroup&id=<%=rs("UserGroupID")%>" onclick="{if(confirm('ɾ�����������Զ�����һ�����û��ĵȼ������Ҳ��ɻָ���ȷ����?')){return true;}return false;}">ɾ��</a></td>
</tr>
<%
Rs.MoveNext
Loop
Set Rs=Nothing
Set Trs=Nothing
%>
<input type=hidden value="0" name="usertitleid">
<input type=hidden value="" name="issetting">
<tr>
<td class=td1 align=center><font color=blue>��</font></td>
<td height="23" class=td1><input size=15 value="" name="usertitle" type=text></td>
<td class=td1><input size=15 value="" name="title" type=text ></td>
<td class=td1><input size=15 value="level0.gif" name="titlepic" type=text></td>
<td class=td1>
<B>0</B>
</td>
<td class=td1>&nbsp;</td>
</tr>
<tr align=center>
<td colspan=6 height=25 class="td2">
<input type=submit class="button" name=submit value="�ύ����">
</td></tr>
</FORM>
</table>
<!--
<BR>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" colspan="7">�������û������<a name="3"></a></th>
</tr>
<tr><td colspan=7 height=25 class="td1">
С��ʾ�����Ȩ�������Էֱ��趨ÿ���������û����Ĭ����̳Ȩ�ޣ�ͨ���������������̳�ϱȽ�������û�Ⱥ�壬��������ͷ�κ�ͼ����ʾ��ǰ̨�û���Ϣ�У��������û�����û���ͬʱӵ�ж���û����Ȩ�ޡ�<BR><font color=blue>������ID��������д����ID�Ļ�ȡ�ɲο�����ĸ������б������������߷ָ�(�磺2|8)��������ֲ��ܸ��£�����ϸ�������д��ID�Ƿ���ȷ</font>
</td></tr>
<tr>
<td width="5%" class=bodytitle><B>��ID</B></td>
<td height="23" width="13%" class=bodytitle><B>��������ͷ��</B></td>
<td width="10%" class=bodytitle><B>ϵͳ������</B></td>
<td width="17%" class=bodytitle><B>������ID</B></td>
<td width="28%" class=bodytitle><B>��������ͼƬ</B></td>
<td height="23" width="8%" class=bodytitle><B>�û���</B></td>
<td width="25%" class=bodytitle><B>����</B></td>
</tr>
<FORM METHOD=POST ACTION="?action=savesysgroup">
<input type=hidden value="4" name="ParentID">
<%
Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where ParentGID=4 Order By UserGroupID")
Do While Not Rs.Eof
%>
<input type=hidden value="<%=rs("UserGroupID")%>" name="usertitleid">
<tr>
<td class=td1 align=center><%=Rs("UserGroupID")%></td>
<td height="23" class=td1><input size=15 value="<%=rs("usertitle")%>" name="usertitle" type=text></td>
<td class=td1><input size=15 value="<%=rs("title")%>" name="title" type=text></td>
<td class=td1><input size=15 value="<%=rs("IsSetting")%>" name="issetting" type=text> *</td>
<td class=td1><input size=15 value="<%=rs("grouppic")%>" name="titlepic" type=text>
<img src="../<%=Dvbbs.Forum_PicUrl%>star/<%=rs("grouppic")%>" border="0">
</td>
<td class=td1>
<B><%
Set Trs=Dvbbs.Execute("Select Count(*) From [Dv_User] Where UserGroupID="&Rs("UserGroupID"))
Response.Write Trs(0)
%></B>
</td>
<td class=td1><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">�༭</a> | <a href="user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">�г��û�</a> | <a href="?action=delusergroup&id=<%=rs("UserGroupID")%>" onclick="{if(confirm('ɾ�����������Զ�����һ�����û��ĵȼ������Ҳ��ɻָ���ȷ����?')){return true;}return false;}">ɾ��</a></td>
</tr>
<%
Rs.MoveNext
Loop
Set Rs=Nothing
Set Trs=Nothing
%>
<input type=hidden value="0" name="usertitleid">
<tr>
<td class=td1 align=center><font color=blue>��</font></td>
<td height="23" class=td1><input size=15 value="" name="usertitle" type=text></td>
<td class=td1><input size=15 value="" name="title" type=text ></td>
<td class=td1><input size=15 value="" name="issetting" type=text> *</td>
<td class=td1><input size=15 value="level0.gif" name="titlepic" type=text></td>
<td class=td1>
<B>0</B>
</td>
<td class=td1>&nbsp;</td>
</tr>
<tr align=center>
<td colspan=7 height=25 class="td2">
<input type=submit class="button" name=submit value="�ύ����">
</td></tr>
</FORM>
</table>
<BR>
//-->
<%
Dim FoundVipGroup
FoundVipGroup = False
%>
<BR>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" colspan="6">Vip�û������<a name="4"></a></th>
</tr>
<tr><td colspan=6 height=25 class="td1">
С��ʾ��VIP�û�����Ȩ�����ڿ��ƣ������û���ʹ��Ȩ�޹��ڣ�ϵͳ�����Զ�����Աת��Ĭ��ע���顣
</td></tr>
<tr>
<td width="5%" class=bodytitle><B>��ID</B></td>
<td height="23" width="15%" class=bodytitle><B>������ͷ��</B></td>
<td width="15%" class=bodytitle><B>ϵͳ������</B></td>
<td width="30%" class=bodytitle><B>������ͼƬ</B></td>
<td height="23" width="10%" class=bodytitle><B>�û���</B></td>
<td width="20%" class=bodytitle><B>����</B></td>
</tr>
<FORM METHOD=POST ACTION="?action=savesysgroup">
<input type=hidden value="5" name="ParentID">
<%
Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where ParentGID=5 Order By UserGroupID")
Do While Not Rs.Eof
FoundVipGroup = True
%>
<input type=hidden value="<%=rs("UserGroupID")%>" name="usertitleid">
<input type=hidden value="<%=rs("IsSetting")%>" name="issetting">
<tr>
<td class=td1 align=center><%=Rs("UserGroupID")%></td>
<td height="23" class=td1><input size=15 value="<%=rs("usertitle")%>" name="usertitle" type=text></td>
<td class=td1><input size=15 value="<%=rs("title")%>" name="title" type=text></td>
<td class=td1><input size=15 value="<%=rs("grouppic")%>" name="titlepic" type=text>
<img src="../<%=Dvbbs.Forum_PicUrl%>star/<%=rs("grouppic")%>" border="0">
</td>
<td class=td1>
<B><%
Set Trs=Dvbbs.Execute("Select Count(*) From [Dv_User] Where UserGroupID="&Rs("UserGroupID"))
Response.Write Trs(0)
%></B>
</td>
<td class=td1><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">�༭</a> | <a href="user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">�г��û�</a> | <a href="?action=delusergroup&id=<%=rs("UserGroupID")%>" onclick="{if(confirm('ɾ����VIP�û���ʧȥ��ص�VIPȨ�ޣ����Ҳ��ɻָ���ȷ����?')){return true;}return false;}">ɾ��</a></td>
</tr>
<%
Rs.MoveNext
Loop
Set Rs=Nothing
Set Trs=Nothing
%>
<input type=hidden value="0" name="usertitleid">
<input type=hidden value="" name="issetting">
<%
If Not FoundVipGroup Then
%>
<tr>
<td class=td1 align=center><font color=blue>��</font></td>
<td height="23" class=td1><input size=15 value="" name="usertitle" type=text></td>
<td class=td1><input size=15 value="Vip�û���" name="title" type=text ></td>
<td class=td1><input size=15 value="level0.gif" name="titlepic" type=text></td>
<td class=td1>
<B>0</B>
</td>
<td class=td1>&nbsp;</td>
</tr>
<%
Else
%>
<input size=15 value="" name="usertitle" type=hidden>
<input size=15 value="Vip�û���" name="title" type=hidden>
<input size=15 value="level0.gif" name="titlepic" type=hidden>
<%
End If
%>
<tr align=center>
<td colspan=6 height=25 class="td2">
<input type=submit class="button" name=submit value="�ύ����">
</td></tr>
</FORM>
</table>
<br/>
<%
End Sub

'����ע���û���(�ȼ�)����������Ϣ
Sub SaveUserGroup()
	Server.ScriptTimeout=99999999
	Dim UserTitleID,UserTitle,MinArticle,TitlePic,i,rs
	For i=1 To Request.Form("usertitleid").Count
		UserTitleID=Replace(Request.Form("usertitleid")(i),"'","")
		UserTitle=Replace(Request.Form("usertitle")(i),"'","")
		MinArticle=Replace(Request.Form("minarticle")(i),"'","")
		TitlePic=Replace(Request.Form("titlepic")(i),"'","")
		If IsNumeric(UserTitleID) And UserTitle<>"" And IsNumeric(MinArticle) And TitlePic<>"" Then
			Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where ParentGID=3 And UserGroupID="&UserTitleID)
			If Not (Rs.Eof And Rs.Bof) Then
				If Rs("UserTitle")<>Trim(UserTitle) Or Rs("GroupPic")<>Trim(TitlePic) Then
					Dvbbs.Execute("Update [Dv_User] Set UserClass='"&UserTitle&"',TitlePic='"&TitlePic&"' Where UserGroupID="&UserTitleID)
				End If
				Dvbbs.Execute("Update Dv_UserGroups Set UserTitle='"&UserTitle&"',MinArticle="&MinArticle&",GroupPic='"&TitlePic&"' Where UserGroupID="&UserTitleID)
			End If
			'�¼����û���(�ȼ�)
			If Clng(UserTitleID) = 0 Then
				Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where UserGroupID=4")
				Dvbbs.Execute("Insert Into Dv_UserGroups (Title,UserTitle,GroupSetting,Orders,MinArticle,TitlePic,GroupPic,ParentGID) Values ('"&Rs("Title")&"','"&UserTitle&"','"&Rs("GroupSetting")&"',0,"&MinArticle&",'"&Rs("TitlePic")&"','"&TitlePic&"',3)")
			End If
		End If
	Next
	Dv_suc("���������û��飨�ȼ������ϳɹ���")
	Set Rs=Nothing
	Dvbbs.LoadGroupSetting
End Sub

'����ϵͳ�����⡢�������û�������������Ϣ
Sub SaveSysGroup()
	Server.ScriptTimeout=99999999
	Dim UserTitleID,UserTitle,TitlePic,ParentID,Title,IsSetting,FoundIsSetting,mIsSetting,GroupIDList,k,rs,sql,i
	SQL = "Select UserGroupID From Dv_UserGroups"
	Set Rs = Dvbbs.Execute(SQL)
		GroupIDList = Rs.GetString(,, "", ",", "")
	Rs.close
	Set Rs = Nothing
	GroupIDList = "," & GroupIDList
	GroupIDList = Replace(GroupIDList,",","|")
	ParentID = Request.Form("ParentID")
	If Not IsNumeric(ParentID) Or ParentID="" Then
		Errmsg = ErrMsg + "<BR><li>�Ƿ����û��������"
		Dvbbs_Error()
		Exit Sub
	End If
	ParentID = Cint(ParentID)
	FoundIsSetting = True
	For i=1 To Request.Form("usertitleid").Count
		UserTitleID=Replace(Request.Form("usertitleid")(i),"'","")
		UserTitle=Replace(Request.Form("usertitle")(i),"'","")
		Title=Replace(Request.Form("title")(i),"'","")
		TitlePic=Replace(Request.Form("titlepic")(i),"'","")
		IsSetting=Replace(Request.Form("issetting")(i),"'","")
		If IsNumeric(UserTitleID) And UserTitle<>"" And TitlePic<>"" Then
			Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where ParentGID="&ParentID&" And UserGroupID="&UserTitleID)
			If Not (Rs.Eof And Rs.Bof) Then
				If Rs("UserTitle")<>Trim(UserTitle) Or Rs("GroupPic")<>Trim(TitlePic) Then
					Dvbbs.Execute("Update [Dv_User] Set UserClass='"&UserTitle&"',TitlePic='"&TitlePic&"' Where UserGroupID="&UserTitleID)
				End If
			End If
			If ParentID = 4 And Trim(IsSetting)<>"" Then
				mIsSetting = Split(IsSetting,"|")
				For k = 0 To Ubound(mIsSetting)
					'�������û��飬��д��UserGroupID�������򲻸���
					If InStr(GroupIDList,"|" & mIsSetting(k) & "|") = 0 Then
						FoundIsSetting = False
						Exit For
					End If
				Next
				If FoundIsSetting Then
					Dvbbs.Execute("Update Dv_UserGroups Set Title='"&Title&"',UserTitle='"&UserTitle&"',GroupPic='"&TitlePic&"',IsSetting='"&IsSetting&"' Where UserGroupID="&UserTitleID)
					'�¼����û���
					If Clng(UserTitleID) = 0 Then
						Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where UserGroupID=4")
						Dvbbs.Execute("Insert Into Dv_UserGroups (Title,UserTitle,GroupSetting,Orders,MinArticle,TitlePic,GroupPic,ParentGID,IsSetting) Values ('"&Title&"','"&UserTitle&"','"&Rs("GroupSetting")&"',0,0,'"&Rs("TitlePic")&"','"&TitlePic&"',"&ParentID&",'"&IsSetting&"')")
					End If
				Else
					Dvbbs.Execute("Update Dv_UserGroups Set Title='"&Title&"',UserTitle='"&UserTitle&"',GroupPic='"&TitlePic&"' Where UserGroupID="&UserTitleID)
				End If
				FoundIsSetting = True
			Else
				Dvbbs.Execute("Update Dv_UserGroups Set Title='"&Title&"',UserTitle='"&UserTitle&"',GroupPic='"&TitlePic&"' Where UserGroupID="&UserTitleID)
				'�¼����û���
				If Clng(UserTitleID) = 0 Then
					Dim tGroupSetting	'�����±�Խ�磬��ƮƮ
					Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where UserGroupID=4")
					tGroupSetting=Rs("GroupSetting")
					tGroupSetting=Split(tGroupSetting,",")
					tGroupSetting(71)="0��0��0��0"
					tGroupSetting=Join(tGroupSetting,",")
					Dvbbs.Execute("Insert Into Dv_UserGroups (Title,UserTitle,GroupSetting,Orders,MinArticle,TitlePic,GroupPic,ParentGID) Values ('"&Title&"','"&UserTitle&"','"&Rs("GroupSetting")&"',0,0,'"&Rs("TitlePic")&"','"&TitlePic&"',"&ParentID&")")
				End If
			End If
		End If
	Next
	Dvbbs.LoadGroupSetting():iGroupSetting_UserName()
	Dv_suc("�����û������ϳɹ���")
	Set Rs=Nothing
End Sub

'ɾ��ע���û���(�ȼ�)��Ϣ
Sub DelUserGroup()
	Dim UserTitleID,tRs,rs
	UserTitleID = Request("id")
	If Not IsNumeric(UserTitleID) Or UserTitleID = "" Then
		Errmsg = ErrMsg + "<BR><li>��ָ��Ҫɾ�����û���(�ȼ�)��"
		Dvbbs_Error()
		Exit Sub
	End If
	UserTitleID = Clng(UserTitleID)
	'����û����Ƿ�����Լ�ȡ���ٽ��û������Ϣ
	'����û���Ϊ���⡢�������飬��������û���ϢΪ��͵ȼ����û���½����Զ����¸���
	Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where (Not ParentGID=1) And UserGroupID = " & UserTitleID)
	If Rs.Eof And Rs.Bof Then
		Errmsg = ErrMsg + "<BR><li>ָ��Ҫɾ�����û���(�ȼ�)�����ڡ�"
		Dvbbs_Error()
		Exit Sub
	ElseIf Not Rs("UserGroupID") = 8 And Rs("ParentGID") = 2 Then
		'ɾ�������û��飨�ȼ���֮�ж� 2005-4-9 Dv.Yz
		Set tRs = Dvbbs.Execute("SELECT TOP 1 * FROM Dv_UserGroups WHERE ParentGID = 3 ORDER BY MinArticle Desc")
		If tRs.Eof And tRs.Bof Then
			Errmsg = ErrMsg + "<BR><li>ע���û���(�ȼ�)Ϊ�գ�����ɾ����������������һ��ע���û���(�ȼ�)��"
			Dvbbs_Error()
			Exit Sub
		Else
			Dvbbs.Execute("UPDATE Dv_User SET UserClass = '" & tRs("UserTitle") & "', TitlePic = '" & tRs("GroupPic") & "', UserGroupID = " & tRs("UserGroupID") & " WHERE UserGroupID = " & UserTitleID)
			Dvbbs.Execute("DELETE FROM Dv_UserGroups WHERE UserGroupID = " & UserTitleID)
		End If
	Else
		Set tRs=Dvbbs.Execute("Select Top 1 * From Dv_UserGroups Where ParentGID=3 And (Not UserGroupID="&UserTitleID&") And MinArticle<="&Rs("MinArticle")&" Order By MinArticle Desc")
		If tRs.Eof And tRs.Bof Then
			Errmsg = ErrMsg + "<BR><li>���û���(�ȼ�)Ϊ���һ��ע���û��飬����ɾ����"
			Dvbbs_Error()
			Exit Sub
		Else
			Dvbbs.Execute("Update Dv_User Set UserClass='"&tRs("UserTitle")&"',TitlePic='"&tRs("GroupPic")&"',UserGroupID="&tRs("UserGroupID")&" Where UserGroupID="&UserTitleID)
			Dvbbs.Execute("Delete From Dv_UserGroups Where UserGroupID = " & UserTitleID)
		End If
		Set tRs=Nothing
	End If
Dvbbs.LoadGroupSetting():iGroupSetting_UserName()
	Dv_suc("�û��飨�ȼ�������ɾ���ɹ���")
	Set Rs=Nothing
End Sub

Sub EditGroup()
	If Not IsNumeric(Replace(Request("groupid"),",","")) Then
		Errmsg = ErrMsg + "<BR><li>��ѡ���Ӧ���û��顣"
		Dvbbs_Error()
		Exit Sub
	End If
	
	If Request("groupaction")="yes" Then
		Dim GroupSetting,SaveGroupid
		Dim UpdateStr,OldStr,NewStr,k,rs,sql
		If Request.Form("title")="" Then
			Errmsg = ErrMsg + "<BR><li>�������û������ƣ�"
			Dvbbs_Error()
			Exit Sub
		End If
		SaveGroupid = Dvbbs.Checkstr(Request.Form("groupid"))
		GroupSetting=GetGroupPermission
		UpdateStr = ""
		Set Rs = Dvbbs.iCreateObject("ADODB.Recordset")
		Sql="Select UserTitle,GroupPic,GroupSetting From Dv_UserGroups Where UserGroupID in ("&SaveGroupid&") "
		Rs.Open Sql,Conn,1,3
		Do While Not Rs.Eof
			If Instr(SaveGroupid,",")=0 Then
				Rs("UserTitle") = Request.Form("title")
				Rs("GroupPic") = Request.Form("grouppic")
				Rs("GroupSetting") = GroupSetting
				'Response.Write GroupSetting
			Else
				NewStr = Split(GroupSetting,",")
				OldStr = Split(Rs("GroupSetting"),",")
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
				If Request.Form("CheckGroupPic")="on" Then
					Rs("GroupPic") = Request.Form("grouppic")
				End If
				Rs("GroupSetting") = UpdateStr
				UpdateStr = ""
			End If
			'Rs("isdisp")=Request("isdisp")
			Rs.update
		Rs.MoveNext
		Loop
		Rs.close
		Set Rs=Nothing
		Dv_suc("�û��飨�ȼ��������޸ĳɹ���")
		Dvbbs.LoadGroupSetting():iGroupSetting_UserName()
	Else
		Dim reGroupSetting
		Set Rs=Dvbbs.Execute("Select * From Dv_Usergroups Where UserGroupID="&Request("groupid"))
		If Rs.Eof And Rs.Bof Then
			Errmsg = ErrMsg + "<BR><li>δ�ҵ����û��飡"
			Dvbbs_Error()
			Exit Sub
		End If
		reGroupSetting=Split(Rs("GroupSetting"),",")
%>
<FORM METHOD=POST ACTION="?action=editgroup" name="TheForm">
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr>
<th style="text-align:center;" colspan="4">
��̳�û���Ȩ������
</th>
</tr>
<tr><td colspan=4 height=25 class="td1"><B>˵��</B>��
<BR>�����������������ø����û��飨�ȼ�������̳�е�Ĭ��Ȩ�ޣ�
<BR>�ڿ���ɾ���ͱ༭�����ӵ��û��飻
<BR>��<font color="red">�ڸ��¶���û�������ʱ����ѡȡ����ߵĸ�ѡ������ֻ��ѡȡ��������Ŀ�Ż���£�</font>
<BR>�ܲ�ִ�ж��û������ʱ������Ҫѡȡ��ߵĸ�ѡ������
</td></tr>
<tr><td colspan=4 height=25 class="td1">
<b>���ù���</b>��
[<a href="#setting1">�༭�û���(�ȼ�)������Ϣ</a>] 
[<a href="#setting2">������ѡ��</a>] 
[<a href="#setting3">����Ȩ��</a>] 
[<a href="#setting4">����/����༭Ȩ��</a>] 
[<a href="#setting5">�ϴ�Ȩ������</a>] 
[<a href="#setting6">����Ȩ��</a>] 
[<a href="#setting7">����Ȩ��</a>] 
[<a href="#setting8">����Ȩ��</a>] 
[<a href="#setting9">��ҪȨ������</a>] 
</td></tr>
<tr><td colspan=4 height=25 class="td1">
<b>���������û�������</b>��<input type="button" class="button" value="ѡ���û���" onclick="getGroup('Select_Group');">
</td></tr>
<tr>
<th style="text-align:center;" colspan="4" align=left>
<a name="setting1"></a>
<INPUT TYPE="checkbox" class="checkbox" NAME="chkall" onclick="CheckAll(this.form);">[ȫѡ]
�༭�û���(�ȼ�)������Ϣ
</th>
</tr>
<tr><td colspan=4 height=25 class="bodytitle">
<B><a href="?">�û���(�ȼ�)����</a> >> <%=SysGroupName(Rs("ParentGID"))%></B>
<%=Rs("UserTitle")%>
<input name="groupid" type="hidden" value="<%=Request("groupid")%>">
</td></tr>
<tr>
<td width="1%" class=td1>&nbsp;</td>
<td height="23" class=td1>�û���(�ȼ�)����</td>
<td height="23" class=td1 colspan=2><input size=35 name="title" type=text value="<%=Rs("UserTitle")%>"></td>
</tr>
<tr>
<td width="1%" class=td1><INPUT TYPE="checkbox" class="checkbox" NAME="CheckGroupPic"></td>
<td height="23" class=td1>�û���(�ȼ�)ͼƬ</td>
<td height="23"class=td1 colspan=2><input size=35 name="grouppic" type=text value="<%=rs("grouppic")%>"></td>
</tr>
<%
GroupPermission(rs("GroupSetting"))
%>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
<%
		Set Rs=Nothing
		Call Select_Group(Request("groupid"))
	End If
End Sub

Sub GroupOnline()
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" align=left>������ʾ</th>
</tr>
<tr align=left>
<td height="23" class="td1" style="LINE-HEIGHT: 140%">
<li>���Ը�ÿ���û���ֱ���������ͼ��ͼƬ����ͼƬ��ʾ���û������б����û���ǰ�棬<B>ͼƬ·��Ϊ���ģ����������·��</B>
<li>����Ϊ 0 ����ʾ����ҳ����ͼ��˵���У�������� 0 ������ʾ����ҳ������ͼ��˵����
</td>
</tr>
<tr align=left>
<td height="23" class="td2" style="LINE-HEIGHT: 140%">
<B>��ݲ���</B>�� <a href="?">�û��������ҳ</a> | <a href="?#1">ϵͳ��</a> | <a href="?#2">������</a> | <a href="?#3">��������</a> | <a href="?action=editgroup&groupid=4">�༭Ĭ��������</a>
</td>
</tr>
</table>
<BR>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" align=left colspan=3>�û�������ͼ������</th>
</tr>
<tr> 
<td width="20%" class=bodytitle><B>������</B></td>
<td height="23" width="10%" class=bodytitle><B>����</B></td>
<td width="*" class=bodytitle><B>��ͼ��</B></td>
</tr>
<FORM METHOD=POST ACTION="?action=saveonline">
<%
Dim rs
Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Order By ParentGID,UserGroupID")
Do While Not Rs.Eof
%>
<input type=hidden value="<%=rs("UserGroupID")%>" name="usertitleid">
<tr align=left>
<td height="23" class="td1">
<%=Rs("UserTitle")%>
</td>
<td height="23" class="td1">
<Input type=text size=5 name="Orders" value="<%=Rs("Orders")%>">
</td>
<td height="23" class="td1">
<Input type=text size=20 name="TitlePic" value="<%=Rs("TitlePic")%>">
<img src="../<%=Dvbbs.Forum_PicUrl%><%=rs("TitlePic")%>" border="0">
<%
If Rs("ParentGID")=0 Then Response.Write "�޸�ע���û����������� <a href=""?action=editgroup&groupid=4"">�༭Ĭ��������</a>"
%>
</td>
</tr>
<%
Rs.MoveNext
Loop
Rs.Close
Set Rs=Nothing
%>
<tr align=center>
<td colspan=3 height=25 class="td2">
<input type=submit class="button" name=submit value="�ύ����">
</td></tr>
</FORM>
</table>
<BR>
<%
End Sub

Sub iGroupSetting_UserName()
	Dvbbs.Name="GroupSetting_UserName"
	Dim i,Str,OutputStr,Outputvalue
	Dim Rs,SQL
	SQL = "Select UserGroupID,GroupSetting From [Dv_UserGroups] order by UserGroupID"
	Set Rs = Dvbbs.Execute(SQL)
	Do while not Rs.Eof
		Str = Str & Rs(0) &","& Split(Rs(1),",")(58)
		Str = Str & "|||"
	Rs.MoveNext
	Loop
	Rs.Close : Set Rs = Nothing
	Dvbbs.value = Left(str,Len(str)-3)
	Str = Split(Dvbbs.value,"|||")
	For i=0 to Ubound(Str)
		OutputStr = Split(Str(i),",")
		Outputvalue = Outputvalue & "GroupUserName["&OutputStr(0)&"]='"&Replace(Replace(Replace(Replace(OutputStr(1),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	Next
	Dvbbs.value = "var GroupUserName = new Array(); " & Outputvalue
End Sub

Sub SaveGroupOnline()
	Dim UserTitleID,Orders,TitlePic,i,rs
	For i=1 To Request.Form("usertitleid").Count
		UserTitleID=Dvbbs.CheckNumeric(Request.Form("usertitleid")(i))
		Orders=Dvbbs.CheckNumeric(Request.Form("Orders")(i))
		TitlePic=Dvbbs.CheckStr(Request.Form("TitlePic")(i))
		Dvbbs.Execute("Update Dv_UserGroups Set Orders="&Orders&",TitlePic='"&TitlePic&"' Where UserGroupID="&UserTitleID)
	Next
	Dv_suc("���������û�������ͼ�����ϳɹ���")
	Set Rs=Nothing
	Dvbbs.LoadGroupSetting():iGroupSetting_UserName()
End Sub

Sub iGroupSetting_UserName()
	Dvbbs.Name="GroupSetting_UserName"
	Dim i,Str,OutputStr,Outputvalue
	Dim Rs,SQL
	SQL = "Select UserGroupID,GroupSetting From [Dv_UserGroups] order by UserGroupID"
	Set Rs = Dvbbs.Execute(SQL)
	Do while not Rs.Eof
		Str = Str & Rs(0) &","& Split(Rs(1),",")(58)
		Str = Str & "|||"
	Rs.MoveNext
	Loop
	Rs.Close : Set Rs = Nothing
	Dvbbs.value = Left(str,Len(str)-3)
	Str = Split(Dvbbs.value,"|||")
	For i=0 to Ubound(Str)
		OutputStr = Split(Str(i),",")
		Outputvalue = Outputvalue & "GroupUserName["&OutputStr(0)&"]='"&Replace(Replace(Replace(Replace(OutputStr(1),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	Next
	Dvbbs.value = "var GroupUserName = new Array(); " & Outputvalue
End Sub
%>