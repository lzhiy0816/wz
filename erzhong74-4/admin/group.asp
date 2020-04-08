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
<th style="text-align:center;" align=left>&nbsp;操作提示</th>
</tr>
<tr align=left>
<td height="23" class="td1" style="LINE-HEIGHT: 140%">
<li>动网论坛用户组分为系统用户组、特殊用户组、注册用户组、多属性用户组四种类型
<li>系统用户组为内置固定用户组，不能添加，供论坛管理之用，不能随意更改，如删除则会引起论坛运行异常
<li>特殊用户组不随用户等级升降而变更，通常建立来分配给一些对论坛有特殊贡献或操作的人员
<li>多属性用户组不随用户等级升降而变更，该组用户<U>可设置享有多个不同用户组的权限</U>，通常建立来分配给一些对论坛有特殊贡献或操作的人员
<li>注册用户组即为传统的用户等级，每个组(等级)可设定不同的权限
<li>默认权限为添加新的用户组时使用其中一些定义好的权限设置，通常新添加用户组后都要再次定义其权限
</td>
</tr>
<tr align=left>
<td height="23" class="td2" style="LINE-HEIGHT: 140%">
<B>快捷操作</B>：<a href="#1">系统组</a> | <a href="#2">特殊组</a> | <a href="#3">多属性组</a>
 | <a href="#4">Vip用户组管理</a>
| <a href="?action=editgroup&groupid=4">编辑默认组资料</a> | <a href="?action=online">在线图例定制</a>
</td>
</tr>
</table>
<BR>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" colspan="6">注册用户组(等级)管理</th>
</tr>
<tr><td colspan=6 height=25 class="td1">
小提示：点击权限您可以分别设定每个注册用户组(等级)分别拥有不同的论坛权限
</td></tr>
<tr>
<td height="23" width="5%" class=bodytitle><B>组ID</B></td>
<td height="23" width="20%" class=bodytitle><B>用户组(等级)名称</B></td>
<td width="15%" class=bodytitle><B>最少发贴</B></td>
<td width="30%" class=bodytitle><B>组(等级)图片</B></td>
<td height="23" width="10%" class=bodytitle><B>用户数</B></td>
<td width="20%" class=bodytitle><B>操作</B></td>
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
<td class=td1><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">编辑</a> | <a href="user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">列出用户</a> | <a href="?action=delusergroup&id=<%=rs("UserGroupID")%>" onclick="{if(confirm('删除操作将会自动更新一部分用户的等级，并且不可恢复，确定吗?')){return true;}return false;}">删除</a></td>
</tr>
<%
Rs.MoveNext
Loop
Set Rs=Nothing
Set Trs=Nothing
%>
<input type=hidden value="0" name="usertitleid">
<tr>
<td class=td1 align=center><font color=blue>新</font></td>
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
<input type=submit class="button" name=submit value="提交更改">
</td></tr>
</FORM>
</table>
<BR>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" colspan="6">系统用户组管理<a name="1"></a></th>
</tr>
<tr><td colspan=6 height=25 class="td1">
小提示：点击权限您可以分别设定每个系统用户组分别拥有不同的论坛权限，系统组头衔和图标显示在前台用户信息中
</td></tr>
<tr>
<td height="23" width="5%" class=bodytitle><B>组ID</B></td>
<td height="23" width="20%" class=bodytitle><B>系统组头衔</B></td>
<td width="15%" class=bodytitle><B>系统中名称</B></td>
<td height="23" width="30%" class=bodytitle><B>系统组图标</B></td>
<td height="23" width="10%" class=bodytitle><B>用户数</B></td>
<td width="20%" class=bodytitle><B>操作</B></td>
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
<td class=td1><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">编辑</a> | <a href="user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">列出用户</a></td>
</tr>
<%
Rs.MoveNext
Loop
Set Rs=Nothing
Set Trs=Nothing
%>
<tr align=center>
<td colspan=6 height=25 class="td2">
<input type=submit class="button" name=submit value="提交更改">
</td></tr>
</FORM>
</table>

<BR>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" colspan="6">特殊用户组管理<a name="2"></a></th>
</tr>
<tr><td colspan=6 height=25 class="td1">
小提示：点击权限您可以分别设定每个特殊用户组分别拥有不同的论坛权限，通常建立来分配给论坛上比较特殊的用户群体，特殊组头衔和图标显示在前台用户信息中
</td></tr>
<tr>
<td width="5%" class=bodytitle><B>组ID</B></td>
<td height="23" width="15%" class=bodytitle><B>特殊组头衔</B></td>
<td width="15%" class=bodytitle><B>系统中名称</B></td>
<td width="30%" class=bodytitle><B>特殊组图片</B></td>
<td height="23" width="10%" class=bodytitle><B>用户数</B></td>
<td width="20%" class=bodytitle><B>操作</B></td>
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
<td class=td1><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">编辑</a> | <a href="user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">列出用户</a> | <a href="?action=delusergroup&id=<%=rs("UserGroupID")%>" onclick="{if(confirm('删除操作将会自动更新一部分用户的等级，并且不可恢复，确定吗?')){return true;}return false;}">删除</a></td>
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
<td class=td1 align=center><font color=blue>新</font></td>
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
<input type=submit class="button" name=submit value="提交更改">
</td></tr>
</FORM>
</table>
<!--
<BR>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" colspan="7">多属性用户组管理<a name="3"></a></th>
</tr>
<tr><td colspan=7 height=25 class="td1">
小提示：点击权限您可以分别设定每个多属性用户组的默认论坛权限，通常建立来分配给论坛上比较特殊的用户群体，多属性组头衔和图标显示在前台用户信息中，多属性用户组的用户可同时拥有多个用户组的权限。<BR><font color=blue>包含组ID请慎重填写，组ID的获取可参考上面的各个组列表，内容用竖线分隔(如：2|8)，如果发现不能更新，请仔细检查所填写组ID是否正确</font>
</td></tr>
<tr>
<td width="5%" class=bodytitle><B>组ID</B></td>
<td height="23" width="13%" class=bodytitle><B>多属性组头衔</B></td>
<td width="10%" class=bodytitle><B>系统中名称</B></td>
<td width="17%" class=bodytitle><B>包含组ID</B></td>
<td width="28%" class=bodytitle><B>多属性组图片</B></td>
<td height="23" width="8%" class=bodytitle><B>用户数</B></td>
<td width="25%" class=bodytitle><B>操作</B></td>
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
<td class=td1><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">编辑</a> | <a href="user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">列出用户</a> | <a href="?action=delusergroup&id=<%=rs("UserGroupID")%>" onclick="{if(confirm('删除操作将会自动更新一部分用户的等级，并且不可恢复，确定吗?')){return true;}return false;}">删除</a></td>
</tr>
<%
Rs.MoveNext
Loop
Set Rs=Nothing
Set Trs=Nothing
%>
<input type=hidden value="0" name="usertitleid">
<tr>
<td class=td1 align=center><font color=blue>新</font></td>
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
<input type=submit class="button" name=submit value="提交更改">
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
<th style="text-align:center;" colspan="6">Vip用户组管理<a name="4"></a></th>
</tr>
<tr><td colspan=6 height=25 class="td1">
小提示：VIP用户将有权限限期控制，当该用户的使用权限过期，系统将会自动将会员转到默认注册组。
</td></tr>
<tr>
<td width="5%" class=bodytitle><B>组ID</B></td>
<td height="23" width="15%" class=bodytitle><B>特殊组头衔</B></td>
<td width="15%" class=bodytitle><B>系统中名称</B></td>
<td width="30%" class=bodytitle><B>特殊组图片</B></td>
<td height="23" width="10%" class=bodytitle><B>用户数</B></td>
<td width="20%" class=bodytitle><B>操作</B></td>
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
<td class=td1><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">编辑</a> | <a href="user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">列出用户</a> | <a href="?action=delusergroup&id=<%=rs("UserGroupID")%>" onclick="{if(confirm('删除后VIP用户将失去相关的VIP权限，并且不可恢复，确定吗?')){return true;}return false;}">删除</a></td>
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
<td class=td1 align=center><font color=blue>新</font></td>
<td height="23" class=td1><input size=15 value="" name="usertitle" type=text></td>
<td class=td1><input size=15 value="Vip用户组" name="title" type=text ></td>
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
<input size=15 value="Vip用户组" name="title" type=hidden>
<input size=15 value="level0.gif" name="titlepic" type=hidden>
<%
End If
%>
<tr align=center>
<td colspan=6 height=25 class="td2">
<input type=submit class="button" name=submit value="提交更改">
</td></tr>
</FORM>
</table>
<br/>
<%
End Sub

'保存注册用户组(等级)批量更改信息
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
			'新加入用户组(等级)
			If Clng(UserTitleID) = 0 Then
				Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where UserGroupID=4")
				Dvbbs.Execute("Insert Into Dv_UserGroups (Title,UserTitle,GroupSetting,Orders,MinArticle,TitlePic,GroupPic,ParentGID) Values ('"&Rs("Title")&"','"&UserTitle&"','"&Rs("GroupSetting")&"',0,"&MinArticle&",'"&Rs("TitlePic")&"','"&TitlePic&"',3)")
			End If
		End If
	Next
	Dv_suc("批量更新用户组（等级）资料成功！")
	Set Rs=Nothing
	Dvbbs.LoadGroupSetting
End Sub

'保存系统、特殊、多属性用户组批量更改信息
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
		Errmsg = ErrMsg + "<BR><li>非法的用户组参数。"
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
					'多属性用户组，填写的UserGroupID不存在则不更新
					If InStr(GroupIDList,"|" & mIsSetting(k) & "|") = 0 Then
						FoundIsSetting = False
						Exit For
					End If
				Next
				If FoundIsSetting Then
					Dvbbs.Execute("Update Dv_UserGroups Set Title='"&Title&"',UserTitle='"&UserTitle&"',GroupPic='"&TitlePic&"',IsSetting='"&IsSetting&"' Where UserGroupID="&UserTitleID)
					'新加入用户组
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
				'新加入用户组
				If Clng(UserTitleID) = 0 Then
					Dim tGroupSetting	'修正下标越界，轻飘飘
					Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where UserGroupID=4")
					tGroupSetting=Rs("GroupSetting")
					tGroupSetting=Split(tGroupSetting,",")
					tGroupSetting(71)="0§0§0§0"
					tGroupSetting=Join(tGroupSetting,",")
					Dvbbs.Execute("Insert Into Dv_UserGroups (Title,UserTitle,GroupSetting,Orders,MinArticle,TitlePic,GroupPic,ParentGID) Values ('"&Title&"','"&UserTitle&"','"&Rs("GroupSetting")&"',0,0,'"&Rs("TitlePic")&"','"&TitlePic&"',"&ParentID&")")
				End If
			End If
		End If
	Next
	Dvbbs.LoadGroupSetting():iGroupSetting_UserName()
	Dv_suc("批量用户组资料成功！")
	Set Rs=Nothing
End Sub

'删除注册用户组(等级)信息
Sub DelUserGroup()
	Dim UserTitleID,tRs,rs
	UserTitleID = Request("id")
	If Not IsNumeric(UserTitleID) Or UserTitleID = "" Then
		Errmsg = ErrMsg + "<BR><li>请指定要删除的用户组(等级)。"
		Dvbbs_Error()
		Exit Sub
	End If
	UserTitleID = Clng(UserTitleID)
	'检测用户组是否存在以及取得临近用户组的信息
	'如果用户组为特殊、多属性组，则更新其用户信息为最低等级，用户登陆后会自动重新更新
	Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where (Not ParentGID=1) And UserGroupID = " & UserTitleID)
	If Rs.Eof And Rs.Bof Then
		Errmsg = ErrMsg + "<BR><li>指定要删除的用户组(等级)不存在。"
		Dvbbs_Error()
		Exit Sub
	ElseIf Not Rs("UserGroupID") = 8 And Rs("ParentGID") = 2 Then
		'删除特殊用户组（等级）之判断 2005-4-9 Dv.Yz
		Set tRs = Dvbbs.Execute("SELECT TOP 1 * FROM Dv_UserGroups WHERE ParentGID = 3 ORDER BY MinArticle Desc")
		If tRs.Eof And tRs.Bof Then
			Errmsg = ErrMsg + "<BR><li>注册用户组(等级)为空，不能删除，请先添加至少一个注册用户组(等级)。"
			Dvbbs_Error()
			Exit Sub
		Else
			Dvbbs.Execute("UPDATE Dv_User SET UserClass = '" & tRs("UserTitle") & "', TitlePic = '" & tRs("GroupPic") & "', UserGroupID = " & tRs("UserGroupID") & " WHERE UserGroupID = " & UserTitleID)
			Dvbbs.Execute("DELETE FROM Dv_UserGroups WHERE UserGroupID = " & UserTitleID)
		End If
	Else
		Set tRs=Dvbbs.Execute("Select Top 1 * From Dv_UserGroups Where ParentGID=3 And (Not UserGroupID="&UserTitleID&") And MinArticle<="&Rs("MinArticle")&" Order By MinArticle Desc")
		If tRs.Eof And tRs.Bof Then
			Errmsg = ErrMsg + "<BR><li>该用户组(等级)为最后一个注册用户组，不能删除。"
			Dvbbs_Error()
			Exit Sub
		Else
			Dvbbs.Execute("Update Dv_User Set UserClass='"&tRs("UserTitle")&"',TitlePic='"&tRs("GroupPic")&"',UserGroupID="&tRs("UserGroupID")&" Where UserGroupID="&UserTitleID)
			Dvbbs.Execute("Delete From Dv_UserGroups Where UserGroupID = " & UserTitleID)
		End If
		Set tRs=Nothing
	End If
Dvbbs.LoadGroupSetting():iGroupSetting_UserName()
	Dv_suc("用户组（等级）资料删除成功！")
	Set Rs=Nothing
End Sub

Sub EditGroup()
	If Not IsNumeric(Replace(Request("groupid"),",","")) Then
		Errmsg = ErrMsg + "<BR><li>请选择对应的用户组。"
		Dvbbs_Error()
		Exit Sub
	End If
	
	If Request("groupaction")="yes" Then
		Dim GroupSetting,SaveGroupid
		Dim UpdateStr,OldStr,NewStr,k,rs,sql
		If Request.Form("title")="" Then
			Errmsg = ErrMsg + "<BR><li>请输入用户组名称！"
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
		Dv_suc("用户组（等级）资料修改成功！")
		Dvbbs.LoadGroupSetting():iGroupSetting_UserName()
	Else
		Dim reGroupSetting
		Set Rs=Dvbbs.Execute("Select * From Dv_Usergroups Where UserGroupID="&Request("groupid"))
		If Rs.Eof And Rs.Bof Then
			Errmsg = ErrMsg + "<BR><li>未找到该用户组！"
			Dvbbs_Error()
			Exit Sub
		End If
		reGroupSetting=Split(Rs("GroupSetting"),",")
%>
<FORM METHOD=POST ACTION="?action=editgroup" name="TheForm">
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr>
<th style="text-align:center;" colspan="4">
论坛用户组权限设置
</th>
</tr>
<tr><td colspan=4 height=25 class="td1"><B>说明</B>：
<BR>①在这里您可以设置各个用户组（等级）在论坛中的默认权限；
<BR>②可以删除和编辑新添加的用户组；
<BR>③<font color="red">在更新多个用户组设置时，请选取最左边的复选表单，只有选取的设置项目才会更新；</font>
<BR>④不执行多用户组更新时，不需要选取左边的复选表单。
</td></tr>
<tr><td colspan=4 height=25 class="td1">
<b>设置功能</b>：
[<a href="#setting1">编辑用户组(等级)资料信息</a>] 
[<a href="#setting2">浏览相关选项</a>] 
[<a href="#setting3">发帖权限</a>] 
[<a href="#setting4">帖子/主题编辑权限</a>] 
[<a href="#setting5">上传权限设置</a>] 
[<a href="#setting6">管理权限</a>] 
[<a href="#setting7">短信权限</a>] 
[<a href="#setting8">其他权限</a>] 
[<a href="#setting9">重要权限设置</a>] 
</td></tr>
<tr><td colspan=4 height=25 class="td1">
<b>批量更新用户组设置</b>：<input type="button" class="button" value="选择用户组" onclick="getGroup('Select_Group');">
</td></tr>
<tr>
<th style="text-align:center;" colspan="4" align=left>
<a name="setting1"></a>
<INPUT TYPE="checkbox" class="checkbox" NAME="chkall" onclick="CheckAll(this.form);">[全选]
编辑用户组(等级)资料信息
</th>
</tr>
<tr><td colspan=4 height=25 class="bodytitle">
<B><a href="?">用户组(等级)管理</a> >> <%=SysGroupName(Rs("ParentGID"))%></B>
<%=Rs("UserTitle")%>
<input name="groupid" type="hidden" value="<%=Request("groupid")%>">
</td></tr>
<tr>
<td width="1%" class=td1>&nbsp;</td>
<td height="23" class=td1>用户组(等级)名称</td>
<td height="23" class=td1 colspan=2><input size=35 name="title" type=text value="<%=Rs("UserTitle")%>"></td>
</tr>
<tr>
<td width="1%" class=td1><INPUT TYPE="checkbox" class="checkbox" NAME="CheckGroupPic"></td>
<td height="23" class=td1>用户组(等级)图片</td>
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
<th style="text-align:center;" align=left>操作提示</th>
</tr>
<tr align=left>
<td height="23" class="td1" style="LINE-HEIGHT: 140%">
<li>可以给每个用户组分别定制其在线图例图片，该图片显示于用户在线列表的用户名前面，<B>图片路径为风格模板中所定义路径</B>
<li>排序为 0 则不显示于首页在线图例说明中，排序大于 0 则按序显示于首页的在线图例说明中
</td>
</tr>
<tr align=left>
<td height="23" class="td2" style="LINE-HEIGHT: 140%">
<B>快捷操作</B>： <a href="?">用户组管理首页</a> | <a href="?#1">系统组</a> | <a href="?#2">特殊组</a> | <a href="?#3">多属性组</a> | <a href="?action=editgroup&groupid=4">编辑默认组资料</a>
</td>
</tr>
</table>
<BR>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr> 
<th style="text-align:center;" align=left colspan=3>用户组在线图例管理</th>
</tr>
<tr> 
<td width="20%" class=bodytitle><B>组名称</B></td>
<td height="23" width="10%" class=bodytitle><B>排序</B></td>
<td width="*" class=bodytitle><B>组图例</B></td>
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
If Rs("ParentGID")=0 Then Response.Write "修改注册用户组名称请点击 <a href=""?action=editgroup&groupid=4"">编辑默认组资料</a>"
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
<input type=submit class="button" name=submit value="提交更改">
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
	Dv_suc("批量更新用户组在线图例资料成功！")
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