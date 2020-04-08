<!--#include file="../conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="../inc/md5.asp"-->
<!-- #include file="../inc/myadmin.asp" -->
<%
	Head()
	Dim	admin_flag
	admin_flag=",16,"
	if not Dvbbs.master	or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg +	"<BR><li>本页面为管理员专用，请<a href=../admin_login.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		dvbbs_error()
	else
		dim	body,username2,password2,oldpassword,oldusername,oldadduser,username1
'''''''''''''''
'取出用户组管理员的组名	2002-12-13
		dim	groupsname,titlepic
		set	rs=Dvbbs.Execute("select usertitle,grouppic	from [dv_UserGroups] where UserGroupID=1 ")
		groupsname=rs(0)
		titlepic=rs(1)
		set	rs=nothing
		Select Case Request("action")
			Case "updat" : update()
			Case "del" : Del()
			Case "pasword" : pasword()
			Case "newpass" : newpass()
			Case "add" : addadmin()
			Case "edit" : userinfo()
			Case "savenew" : savenew()
			Case Else
				userlist()
		End Select
		If ErrMsg<>"" Then Dvbbs_Error
		Footer()
	end	if

	sub	userlist()
%>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
				<tr>
				  <th height=22	colspan=5>管理员管理(点击用户名进行操作)</th>
				</tr>
				<tr	align=center>
				  <td width="30%" height=22	class="forumHeaderBackgroundAlternate"><B>用户名</B></td><td width="25%" class="forumHeaderBackgroundAlternate"><B>上次登录时间</B></td><td width="15%" class="forumHeaderBackgroundAlternate"><B>上次登陆IP</B></td><td width="15%" class="forumHeaderBackgroundAlternate"><B>操作</B></td>
				</tr>
<%
	set	rs=Dvbbs.Execute("select * from "&admintable&" order by LastLogin desc")
	do while not rs.eof
%>
				<tr>
				  <td class=forumrow><a	href="admin.asp?id=<%=rs("id")%>&action=pasword"><%=rs("username")%></a></td><td class=forumrow><%=rs("LastLogin")%></td><td	class=forumrow><%=rs("LastLoginIP")%></td><td class=forumrow><a	href="admin.asp?action=del&id=<%=rs("id")%>&name=<%=Rs("adduser")%>" onclick="{if(confirm('删除后该管理员将不可进入后台！\n\n确定删除吗?')){return true;}return false;}">删除</a>&nbsp;&nbsp;<a	href="admin.asp?id=<%=rs("id")%>&action=edit">编辑权限</a></td>
				</tr>
<%
	rs.movenext
	loop
	rs.close
	set	rs=nothing
%>
		   </table>
<%
	end	sub

Sub	Del()
	Dim	UserTitle
	Rem	更新撤销管理员后的等级名称 2004-4-29 Dvbbs.YangZheng
	Sql	= "SELECT Top 1	UserTitle From Dv_UserGroups Where MinArticle >	0 And ParentGID	= 4	Order By UserGroupID"
	Set	Rs = Dvbbs.Execute(Sql)
	If Rs.Eof And Rs.Bof Then
		UserTitle =	"新手上路"
	Else
		UserTitle =	Rs(0)
	End	If
	Dvbbs.Execute("DELETE FROM " & Admintable &	" WHERE	Id = " & Request("Id"))
	Dvbbs.Execute("UPDATE [Dv_User]	SET	Usergroupid	= 4, UserClass = '"	& UserTitle	& "' WHERE Username	= '" & Replace(Request("name"),"'","")	& "'")
	body="<li>管理员删除成功。"
	Dv_suc(body)
End	Sub

Sub	pasword()
	Dim AcceptIP,i,AddIP
	set	Rs=Dvbbs.Execute("select * from	"&admintable&" where id="&request("id"))
	oldpassword=rs("password")
	oldadduser=rs("adduser")
	AcceptIP = Rs("AcceptIP") &""
	AddIP = Dvbbs.UserTrueIP
	AddIP = Left(AddIP, InStrRev(AddIP, ".")-1)
	AddIP = Left(AddIP, InStrRev(AddIP, ".")-1)
	AddIP = AddIP &".*.*"
  %>
<form action="?action=newpass" method=post>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
			   <tr>
				  <th colspan=2	height=23>管理员资料管理－－密码修改
				  </th>
				</tr>
			   <tr >
			<td	width="26%"	align="right" class=forumrow>后台登录名称：</td>
			<td	width="74%"	class=forumrow>
			  <input type=hidden name="oldusername"	value="<%=rs("username")%>">
			  <input type=text name="username2"	value="<%=rs("username")%>">  (可与注册名不同)
			</td>
		  </tr>
		  <tr >
			<td	width="26%"	align="right" class=forumrow>后台登录密码：</td>
			<td	width="74%"	class=forumrow>
			  <input type="password" name="password2" value="<%=oldpassword%>">	 (可与注册密码不同,如要修改请直接输入)
			</td>
		  </tr>
		  <tr>
			<td	width="26%"	align="right" class=forumrow height=23>前台用户名称：</td>
			<td	width="74%"	class=forumrow><%=oldadduser%>
			</td>
		 </tr>
		<tr>
			<td	width="26%"	align="right" class=forumrow height=23>添加只允许登陆IP列表：
			</td>
			<td	width="74%"	class=forumrow>
			<textarea name="AddAcceptIP" cols="40" rows="8"><%
			If AcceptIP<>"" or not IsNull(AcceptIP) Then
				AcceptIP=Split(Trim(AcceptIP),"|")
				For i=0 To Ubound(AcceptIP)
					Response.Write AcceptIP(i)
					If i<Ubound(AcceptIP) Then Response.Write vbCrLf
				Next
			End If
			%></textarea><br><input type=button value="添加自已当前IP" onclick="AddAcceptIP.value+='\n<%=AddIP%>'"> <%=dvbbs.UserTrueIP%>
			<fieldset class="fieldset" style="margin:2px 2px 2px 2px">
			<legend><B>添加说明</B></legend>
			<ol>
			<LI><b>清空不填写即允许所有IP登陆后台。</b>
			<LI><b><font color=red>尽量采用IP段的方式，如：10.10.*.*。</font></b>
			<LI><b>注意：提交后在下次登陆将会生效，若IP填错将会无法登陆后台。</b>
			<LI>添加IP后，该管理员访问IP必需符合允许IP列表才能登陆后台。	<LI>您可以添加多个允许IP，每个IP用回车分隔，允许IP的书写方式如202.152.12.1就允许了202.152.12.1这个IP的登陆后台，如202.152.12.*就允许了以202.152.12开头的IP登陆后台。
			<LI>在添加多个IP的时候，请注意最后一个IP的后面不要加回车。
			</ol></fieldset>
			</td>
		  </tr>
		  <tr align="center">
			<td	colspan="2"	class=forumrow>
			  <input type=hidden name="adduser"	value="<%=oldadduser%>">
			  <input type=hidden name=id value="<%=request("id")%>">
			  <input type="submit" name="Submit" value="更 新">
			</td>
		  </tr>
		</table>
		</form>
<%		 Rs.close
		 Set Rs=nothing
End	Sub

Sub	newpass()
	dim	passnw,usernw,aduser
	Dim AcceptIP,Tempstr,i
	set	rs=Dvbbs.Execute("select * from	"&admintable&" where id="&request("id"))
	oldpassword=rs("password")
	if request("username2")="" then
		ErrMsg = "<li>请输入管理员名字。<a href=?>［ <font color=red>返回</font> ］</a>"
		exit sub
	else
		usernw=trim(request("username2"))
	end	if
	if request("password2")="" then
		ErrMsg = "<li>请输入您的密码。<a href=?>［ <font color=red>返回</font> ］</a>"
		exit sub
	elseif trim(request("password2"))=oldpassword then
		passnw=request("password2")
	else
		passnw=md5(request("password2"),16)
	end	if
	if request("adduser")="" then
		ErrMsg = "<li>请输入管理员名字。<a	href=?>［ <font	color=red>返回</font> ］</a>"
		exit sub
	else
		aduser=trim(request("adduser"))
	end	if
	Tempstr = Trim(Request.Form("AddAcceptIP"))
	If Tempstr<>"" Then
		Tempstr = Split(Tempstr,vbCrLf)
		For i = 0 To ubound(Tempstr)
			If Tempstr(i)<>"" and Tempstr(i)<>" " and Isnumeric(Replace(Replace(Tempstr(i),".",""),"*","")) and Instr(Tempstr(i),",")=0 Then
				If i=0 or AcceptIP="" Then
					AcceptIP = Tempstr(i)
				Else
					AcceptIP = AcceptIP & "|" & Tempstr(i)
				End If
			End If
		Next
	End If
	If Len(AcceptIP)>=255 Then
		ErrMsg = "<li>允许IP列表太多，超出了限制。<a href=?>［ <font color=red>返回</font> ］</a>"
		exit sub
	End If
	set	rs=server.createobject("adodb.recordset")
	sql="select	* from "&admintable&" where	username='"&trim(request("oldusername"))&"'"
	rs.open	sql,conn,1,3
	if not rs.eof and not rs.bof then
	Rs("username") = usernw
	Rs("adduser") = aduser
	Rs("password") = passnw
	Rs("AcceptIP") = AcceptIP
''''''''''''''
'更新用户的的级别
	'Dvbbs.Execute("update [dv_user]	set	usergroupid=1,userclass='"&groupsname&"',titlepic='"&titlepic&"' where username='"&trim(request("adduser"))&"'")	'
	body="<li>管理员资料更新成功，请记住更新信息。<br> 管理员："&request("username2")&"	<BR> 密	  码："&request("password2")&" <a href=?>［ <font	color=red>返回</font> ］</a>"
	Dv_suc(body)
	rs.update
	End	if
	rs.close
	set	rs=nothing
end	sub

sub	addadmin()
%>
<form action="?action=savenew" method=post>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr>
	<th colspan=2	height=23>管理员管理－－添加管理员</th>
</tr>
<tr>
	<td	width="26%"	align="right" class=forumrow>后台登录名称：</td>
	<td	width="74%"	class=forumrow>
	<input type=text name="username2" size=30>  (可与注册名不同)
	</td>
</tr>
<tr>
	<td	width="26%"	align="right" class=forumrow>后台登录密码：</td>
	<td	width="74%"	class=forumrow>
	<input type="password" name="password2" size=33>	(可与注册密码不同)
	</td>
</tr>
<tr>
	<td	width="26%"	align="right" class=forumrow height=23>前台用户名称：</td>
	<td	width="74%"	class=forumrow><input type=text	name="username1" size=30>  (本选项填写后不允许修改)
	</td>
</tr>
<tr>
	<td	width="26%"	align="right" class=forumrow height=23>在前台显示为管理员：</td>
	<td	width="74%"	class=forumrow>是<input name="isdisp" type=radio value="1" checked>&nbsp;否<input name="isdisp" type=radio value="0">  (如改为显示请到用户管理中编辑资料)
	</td>
</tr>
<tr align="center">
	<td	colspan="2"	class=forumrow>
	<input type="submit" name="Submit" value="添 加">
	</td>
</tr>
</table>
</form>
<%
end	sub

sub	savenew()
dim	adminuserid
	if request.form("username2")=""	then
		ErrMsg = "请输入后台登录用户名！"
		exit sub
	end	if
	if request.form("username1")=""	then
		ErrMsg = "请输入前台登录用户名！"
		exit sub
	end	if
	if request.form("password2")=""	then
		ErrMsg = "请输入后台登录密码！"
		exit sub
	end	if
	dim isdisp
	If request("isdisp")="" then
		isdisp=1
	else
		isdisp=cint(request("isdisp"))
	end if

	set	rs=Dvbbs.Execute("select userid	from [dv_user] where username='"&replace(request.form("username1"),"'","")&"'")
	if rs.eof and rs.bof then
		ErrMsg = "您输入的用户名不是一个有效的注册用户！"
		exit sub
	else
		adminuserid=rs(0)
	end	if

	set	rs=Dvbbs.Execute("select username from "&admintable&" where	username='"&replace(request.form("username2"),"'","")&"'")
	if not (rs.eof and rs.bof) then
		ErrMsg = "您输入的用户名已经在管理用户中存在！"
		exit sub
	end	if
	if isdisp=1 then
	Dvbbs.Execute("update [dv_user]	set	usergroupid=1 ,	userclass='"&groupsname&"',titlepic='"&titlepic&"' where userid="&adminuserid&" ")
	end if
	Dvbbs.Execute("insert into "&Admintable&" (username,[password],adduser)	values ('"&replace(request.form("username2"),"'","")&"','"&md5(replace(request.form("password2"),"'",""),16)&"','"&replace(request.form("username1"),"'","")&"')")
	body="用户ID:"&adminuserid&" 添加成功，请记住新管理员后台登录信息，如需修改请返回管理员管理！"
	Dv_suc(body)
end	sub

sub	userinfo()
dim	menu(8,10),trs,k
menu(0,0)="常规管理"
menu(0,1)="<a href=setting.asp target=main>基本设置</a>@@1"
menu(0,2)="<a href=ads.asp target=main>广告管理</a>@@2"
menu(0,3)="<a href=log.asp target=main>论坛日志</a>@@3"
menu(0,4)="<a href=help.asp target=main>帮助管理</a>@@4"
menu(0,5)="<a href=wealth.asp	target=main>积分设置</a>@@5"
menu(0,6)="<a href=message.asp target=main>短信管理</a>@@6"
menu(0,7)="<a href=announcements.asp?boardid=0&action=AddAnn target=_blank>公告管理</a>@@7"
menu(0,8)="<a href=menpai.asp	target=main>门派管理</a>@@8"
'menu(0,9)="<a href=""ForumPay.asp"" target=main>论坛交易信息管理</a>"
'menu(0,10)="<a href=""ForumNewsSetting.asp"" target=main>首页调用管理</a>"

menu(1,0)="论坛管理"
menu(1,1)="<a href=board.asp?action=add target=main>版面(分类)添加</a> | <a href=board.asp target=main>管理</a>@@9"
menu(1,2)="<a href=board.asp?action=permission target=main>分版面用户权限设置</a>@@10"
menu(1,3)="<a href=boardunite.asp	target=main>合并版面数据</a>@@11"
menu(1,4)="<a href=update.asp	target=main>重计论坛数据和修复</a>@@12"
menu(1,5)="<a href=link.asp?action=add target=main>友情论坛添加</a> |	<a href=link.asp target=main>管理</a>@@13"

menu(2,0)="用户管理"
menu(2,1)="<a href=user.asp target=main>用户资料(权限)管理</a>@@14"
menu(2,2)="<a href=group.asp?action=addgroup target=main>用户组添加</a> |	<a href=group.asp	target=main>管理</a>@@15"
menu(2,3)="<a href=admin.asp?action=add target=main>管理员添加</a> | <a href=admin.asp target=main>管理</a>@@16"
menu(2,4)="<a href=grade.asp?action=add target=main>用户等级添加</a> | <a	href=grade.asp target=main>管理</a>@@17"
menu(2,5)="<a href=update.asp?action=updateuser target=main>奖惩用户管理</a>@@18"
menu(2,6)="<a href=update.asp?action=updateuser target=main>重计用户各项数据</a>@@19"
menu(2,7)="<a href=update.asp?action=SenndEmail target=main>用户邮件通知</a>@@20"

menu(3,0)="外观设置"
menu(3,1)="<a href=template.asp target=main>风格界面模板总管理</a>@@20"
menu(3,2)="<a href=loadskin.asp target=main>模板导出</a> | <a	href=loadskin.asp?action=load	target=main>导入</a>@@21"

menu(4,0)="论坛帖子管理"
menu(4,1)="<a href=alldel.asp	target=main>批量删除</a> | <a href=alldel.asp?action=moveinfo	target=main>批量移动</a>@@22"
menu(4,2)="<a href=recycle.asp target=_blank>回收站管理</a>@@23"
menu(4,3)="<a href=postdata.asp?action=Nowused target=main>当前帖子数据表管理</a>@@24"
menu(4,4)="<a href=postdata.asp target=main>数据表间帖子转换</a>@@25"

menu(5,0)="替换/限制处理"
menu(5,1)="<a href=badword.asp?reaction=badword target=main>脏话过滤设置</a>@@26"
menu(5,2)="<a href=badword.asp?reaction=splitreg target=main>注册过滤字符</a>@@27"
menu(5,3)="<a href=lockip.asp?action=add target=main>IP来访限定添加</a> |	<a href=lockip.asp target=main>管理</a>@@28"
menu(5,4)="<a href=address.asp?action=add	target=main>论坛IP库添加</a> | <a href=address.asp target=main>管理</a>@@29"

menu(6,0)="数据处理(Access)"
menu(6,1)="<a href=data.asp?action=CompressData target=main>压缩数据库</a>@@30"
menu(6,2)="<a href=data.asp?action=BackupData	target=main>备份数据库</a>@@31"
menu(6,3)="<a href=data.asp?action=RestoreData target=main>恢复数据库</a>@@32"
menu(6,4)="<a href=data.asp?action=SpaceSize target=main>系统空间占用</a>@@33"

menu(7,0)="文件管理"
menu(7,1)="<a href=upUserface.asp	target=main>上传头像管理</a>@@34"
menu(7,2)="<a href=uploadlist.asp	target=main>上传文件管理</a>@@35"

menu(8,0)="菜单管理"
menu(8,1)="<a href=plus.asp target=main>论坛菜单管理</a>@@36"

dim	j,tmpmenu,menuname,menurl
set	rs=Dvbbs.Execute("select * from	"&admintable&" where id="&request("id"))
%>
<form action="admin.asp?action=updat"	method=post	name=adminflag>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr>
<th	height=25><b>管理员权限管理</b>(请选择相应的权限分配给管理员 <%=rs("username")%>)
</th>
</tr>
<tr>
<td	height=25 class="forumHeaderBackgroundAlternate"><b>>>全局权限</b></td></tr>
<tr><td	class=forumrow>
<%
Dim n
n = 0
for i=0 to ubound(menu,1)
Response.Write "<b>"&menu(i,0)&"</b><br>"
on error resume	next
for	j=1	to ubound(menu,2)
if isempty(menu(i,j)) then exit	for
tmpmenu=split(menu(i,j),"@@")
menuname=tmpmenu(0)
menurl=tmpmenu(1)
n = n+1
%>
<input type="checkbox" name="flag" <% if instr(","&session("flag")&",",",16,")=0 then response.write "disabled=true" %>	value="<%=n%>"	<% if instr(","&rs("flag")&",",","&n&",")>0 then response.write "checked" %>><%=menurl%>.<%=menuname%>&nbsp;&nbsp;
<%
next%><br><br>
<%next%>
<input type=hidden name=id value="<%=request("id")%>">
<input type="submit" name="Submit" value="更新"><input name=chkall type=checkbox value=on	onclick=CheckAll(this.form)>选择所有权限
</td>
</tr>
</table>
</form>
<%
rs.close
set	rs=nothing
end	sub

sub	update()
' 1, 2,	3, 4, 5, 6,	7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35
'Response.Write	request("flag")
'response.end
set	rs=server.createobject("adodb.recordset")
sql="select	* from "&admintable&" where	id="&request("id")
rs.open	sql,conn,1,3
if not rs.eof and not rs.bof then
rs("flag")=replace(Request("flag")," ","")
	body="<li>管理员更新成功，请记住更新信息。"
	Dv_suc(body)
rs.update
if rs("adduser")=Dvbbs.membername then session("flag")=replace(request("flag")," ","")
end	if
rs.close
set	rs=nothing
end	sub

%>