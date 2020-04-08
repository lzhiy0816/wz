<!--#include file="../conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="../inc/dv_clsother.asp"-->
<!--#include file="../inc/md5.asp"-->
<!--#include file="../inc/GroupPermission.asp"-->
<!--#include file="../dv_dpo/cls_dvapi.asp"-->
<%
Head()
Dim admin_flag,sqlstr,myrootid
FoundErr=False 
admin_flag=",16,"
CheckAdmin(admin_flag)

Dim tRs,UserInfo,UserTitle
UserMain(1)

Select Case Request("action")
Case "fix"
	Fixuser()
Case "userSearch"
	UserSearch()
Case "touser"
	ToUser()
Case "modify"
	UserModify()
Case "saveuserinfo"
	SaveUserInfo()
Case "UserPermission"
	UserPermission()
Case "UserBoardPermission"
	UserBoardPermission()
Case "saveuserpermission"
	SaveUserPermission()
Case "uniteuser"
	UniteUser()
Case Else
	UserIndex()
End Select

UserMain(0)
Footer()

'用户管理通用头部
Sub UserMain(Str)
	If Str = 1 Then
%>
<table cellpadding="2" cellspacing="1" border="0" width="100%" align=center>
<tr>
<th colspan=8 style="text-align:center;">用户管理</th>
</tr>
<tr>
<td width="20%" class=td2 align="center"><button Style="width:80;height:50;border: 1px outset;" class="button">注意事项</button></td>
<td width="80%" class=td2 colspan=7><li>①点删除按钮将删除所选定的用户，此操作是不可逆的；<li>②您可以批量移动用户到相应的组；<li>③点用户名进行相应的资料操作；<li>④点用户最后登陆IP可进行锁定IP操作；<li>⑤点用户Email将给该用户发送Email；<li>⑥点修复贴子将会修复该用户所发的贴子数据并更新其文章数，用于误删ID用户贴的修复。</td>
</tr>
<tr>
<td width=100% class=td2 colspan=8>
快速查看：<a href="user.asp">用户管理首页</a> | <a href="?action=userSearch&userSearch=1"><%If Request("userSearch")="1" Then%><font color=red><%End If%>所有用户<%If Request("userSearch")="1" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=2"><%If Request("userSearch")="2" Then%><font color=red><%End If%>发贴TOP100<%If Request("userSearch")="2" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=3"><%If Request("userSearch")="3" Then%><font color=red><%End If%>发贴END100<%If Request("userSearch")="3" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=4"><%If Request("userSearch")="4" Then%><font color=red><%End If%>24H内登录<%If Request("userSearch")="4" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=5"><%If Request("userSearch")="5" Then%><font color=red><%End If%>24H内注册<%If Request("userSearch")="5" Then%></font><%End If%></a><BR>
　　　　　<a href="?action=userSearch&userSearch=6"><%If Request("userSearch")="6" Then%><font color=red><%End If%>等待验证会员<%If Request("userSearch")="6" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=7"><%If Request("userSearch")="7" Then%><font color=red><%End If%>邮件验证<%If Request("userSearch")="7" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=8"><%If Request("userSearch")="8" Then%><font color=red><%End If%>管理　团队<%If Request("userSearch")="8" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=11"><%If Request("userSearch")="11" Then%><font color=red><%End If%>屏蔽　用户<%If Request("userSearch")="11" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=12"><%If Request("userSearch")="12" Then%><font color=red><%End If%>锁定 用户<%If Request("userSearch")="12" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=14"><%If Request("userSearch")="13" Then%><font color=red><%End If%>自定义权限用户<%If Request("userSearch")="13" Then%></font><%End If%></a>
 | <a href="?action=userSearch&userSearch=15"><%If Request("userSearch")="15" Then%><font color=red><%End If%>VIP用户<%If Request("userSearch")="15" Then%></font><%End If%></a>
</td>
</tr>
<tr>
<td width=100% class=td2 colspan=8>
功能选项：<a href="?action=uniteuser">合并用户</a> | <a href="update_user.asp">奖惩用户管理</a> <!--| <a href="boardmastergrade.asp">版主工作情况</a>-->
</td>
</tr>
<%
	Else
%>
</table>
<p></p>
<%
	End If
End Sub

'用户管理首页，搜索项
Sub UserIndex()
%>
<form action="?action=userSearch" method=post>
<tr>
<th colspan=7 style="text-align:center;">高级查询</th>
</tr>
<tr>
<td width="20%" class=td1>注意事项</td>
<td width="80%" class=td1 colspan=5>在记录很多的情况下搜索条件越多查询越慢，请尽量减少查询条件；最多显示记录数也不宜选择过大</td>
</tr>
<tr>
<td width="20%" class=td1>最多显示记录数</td>
<td width="80%" class=td1 colspan=5><input size=45 name="searchMax" type=text value=100></td>
</tr>
<tr>
<td width="20%" class=td1>用户名</td>
<td width="80%" class=td1 colspan=5><input size=45 name="username" type=text>&nbsp;<input type=checkbox class=checkbox name="usernamechk" value="yes" checked>用户名完整匹配</td>
</tr>
<tr>
<td width="20%" class=td1>用户组</td>
<td width="80%" class=td1 colspan=5>
<select size=1 name="usergroups">
<option value=0>任意</option>
<%
Dim rs
set rs=Dvbbs.Execute("select usergroupid,UserTitle,ParentGID from dv_usergroups Where Not ParentGID=0 order by ParentGID,usergroupid")
do while not rs.eof
response.write "<option value="&rs(0)&">"&SysGroupName(Rs(2)) & rs(1)&"</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
</tr>
<tr>
<td width="20%" class=td1>Email包含</td>
<td width="80%" class=td1 colspan=5><input size=45 name="userEmail" type=text></td>
</tr>
<tr>
<td width="20%" class=td1>用户IM包含</td>
<td width="80%" class=td1 colspan=5><input size=45 name="userim" type=text> 包括主页、OICQ、UC、ICQ、YAHOO、AIM、MSN</td>
</tr>
<tr>
<td width="20%" class=td1>登录IP包含</td>
<td width="80%" class=td1 colspan=5><input size=45 name="lastip" type=text></td>
</tr>
<tr>
<td width="20%" class=td1>头衔包含</td>
<td width="80%" class=td1 colspan=5><input size=45 name="usertitle" type=text></td>
</tr>
<tr>
<td width="20%" class=td1>签名包含</td>
<td width="80%" class=td1 colspan=5><input size=45 name="sign" type=text></td>
</tr>
<tr>
<td width="20%" class=td1>详细资料包含</td>
<td width="80%" class=td1 colspan=5><input size=45 name="userinfo" type=text></td>
</tr>
<!--shinzeal加入特殊搜索-->
<tr>
<th colspan=7 style="text-align:center;">特殊查询&nbsp;（注意： <多于> 或 <少于> 已默认包含 <等于>；条件留空则不使用此条件 ）</th>
</tr>
<tr>
<td class=td1 colspan=7>
<table ID="Table1" width="100%">
<tr>
<td width="50%" class=td2>登录次数:<input type=radio class=radio value=more name="loginR" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="loginR">&nbsp;少于&nbsp;&nbsp;<input size=5 name="loginT" type=text> 次</td>
<td width="50%" class=td2>消失天数:<input type=radio class=radio value=more name="vanishR" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="vanishR">&nbsp;少于&nbsp;&nbsp;<input size=5 name="vanishT" type=text> 天</td>
</tr>
<tr>
<td>注册天数:<input type=radio class=radio value=more name="regR" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="regR">&nbsp;少于&nbsp;&nbsp;<input size=5 name="regT" type=text> 天</td>
<td>发表帖数:<input type=radio class=radio value=more name="artcleR" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="artcleR">&nbsp;少于&nbsp;&nbsp;<input size=5 name="artcleT" type=text> 篇</td>
</tr>

<tr>
<td class=td2>用户金钱:<input type=radio class=radio value=more name="UWealth" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="UWealth">&nbsp;少于&nbsp;&nbsp;<input size=5 name="UWealth_value" type=text></td>
<td class=td2>用户经验:<input type=radio class=radio value=more name="UEP" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="UEP">&nbsp;少于&nbsp;&nbsp;<input size=5 name="UEP_value" type=text></td>
</tr>
<tr>
<td>用户魅力:<input type=radio class=radio value=more name="UCP" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="UCP">&nbsp;少于&nbsp;&nbsp;<input size=5 name="UCP_value" type=text></td>
<td>用户威望:<input type=radio class=radio value=more name="UPower" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="UPower">&nbsp;少于&nbsp;&nbsp;<input size=5 name="UPower_value" type=text></td>
</tr>
<tr>
<td class=td2>用户金币:<input type=radio class=radio value=more name="UMoney" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="UMoney">&nbsp;少于&nbsp;&nbsp;<input size=5 name="UMoney_value" type=text></td>
<td class=td2>用户点券:<input type=radio class=radio value=more name="UTicket" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="UTicket">&nbsp;少于&nbsp;&nbsp;<input size=5 name="UTicket_value" type=text></td>
</tr>
<tr>
<td class=td1><LI>以下条件请选取相应的VIP用户组进行查询</LI></td>
</tr>
<tr>
<td class=td2>Vip登记时间:<input type=radio class=radio value=more name="UVipStarTime" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="UVipStarTime">&nbsp;少于&nbsp;&nbsp;<input size=5 name="UVipStarTime_value" type=text></td>
<td class=td2>Vip截止时间:<input type=radio class=radio value=more name="UVipEndTime" checked>&nbsp;多于&nbsp;<input type=radio class=radio value=less name="UVipEndTime">&nbsp;少于&nbsp;&nbsp;<input size=5 name="UVipEndTime_value" type=text></td>
</tr>
</table>
</td></tr>

<!--特殊搜索结束-->
<tr>
<td width="100%" class=td1 align=center colspan=7><input name="submit" type=submit class=button value="   搜  索   "></td>
</tr>
<input type=hidden value="9" name="userSearch">
</form>
<%
End Sub

'用户搜索结果项
Sub UserSearch()
%>
<tr>
<th colspan=8 style="text-align:center;">搜索结果</th>
</tr>
<%
	dim currentpage,page_count,Pcount
	dim totalrec,endpage
	Dim rs,sql
	currentPage=request("page")
	if currentpage="" or not IsNumeric(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
		if err then
			currentpage=1
			err.clear
		end if
	end if
	Sql = " Userid, Username, Useremail, LastLogin, UserLastIP, UserPost, UserGroupID,Vip_StarTime,Vip_EndTime"
	Set Rs = Dvbbs.iCreateObject("ADODB.Recordset")
	Select Case Request("UserSearch")
	Case 1
		Sql = "SELECT " & Sql & " FROM [Dv_User] ORDER BY UserID DESC"
	Case 2
		Sql = "SELECT TOP 100 " & Sql & " FROM [Dv_User] ORDER BY UserPost DESC"
	case 3
		sql="select top 100 " & Sql & " from [dv_user]  order by UserPost"
	case 4
		If IsSqlDataBase=1 Then
		sql="select " & Sql & " from [dv_user]  where datediff(hour,LastLogin,"&SqlNowString&")<25 order by lastlogin desc"
		else
		sql="select " & Sql & " from [dv_user]  where datediff('h',LastLogin,"&SqlNowString&")<25 order by lastlogin desc"
		end if
	case 5
		If IsSqlDataBase=1 Then
		sql="select " & Sql & " from [dv_user]  where datediff(hour,JoinDate,"&SqlNowString&")<25 order by UserID desc"
		else
		sql="select " & Sql & " from [dv_user]  where datediff('h',JoinDate,"&SqlNowString&")<25 order by UserID desc"
		end if
	case 6
		sql="select " & Sql & " from [dv_user]  where usergroupid=5 order by UserID desc"
	case 7
		sql="select " & Sql & " from [dv_user]  where usergroupid=6 order by UserID desc"
	case 8
		sql="select " & Sql & " from [dv_user]  where usergroupid<4 order by usergroupid"
	case 10
		Sql = "select " & Sql & " from [dv_user]  where usergroupid="&request("usergroupid")&" order by UserID desc"
	case 11
		sql="select " & Sql & " from [dv_user]  where lockuser=2 order by userid desc"
	case 12
		sql="select " & Sql & " from [dv_user]  where lockuser=1 order by userid desc"
	case 13
		sql="select " & Sql & " from [dv_user]  where IsChallenge=1 order by userid desc"
	case 14
		Sql = "SELECT " & Sql & " FROM [Dv_User] WHERE UserID IN (SELECT Uc_UserID FROM Dv_UserAccess) ORDER BY Userid DESC"
	case 15
		Sql = "SELECT " & Sql & " FROM [dv_user]  WHERE UserGroupid IN (SELECT UserGroupID FROM Dv_UserGroups WHERE ParentGID=5) ORDER BY Vip_EndTime desc,UserID desc"
	case 9
		sqlstr=""
		if request("username")<>"" then
			if request("usernamechk")="yes" then
			sqlstr=" username='"&request("username")&"'"
			else
			sqlstr=" username like '%"&request("username")&"%'"
			end if
		end if
		if cint(request("usergroups"))>0 then
			if sqlstr="" then
			sqlstr=" usergroupid="&request("usergroups")&""
			else
			sqlstr=sqlstr & " and usergroupid="&CheckNumeric(request("usergroups"))
			end if
		end if
		'if request("userclass")<>"0" then
		'	if sqlstr="" then
		'	sqlstr=" userclass='"&request("userclass")&"'"
		'	else
		'	sqlstr=sqlstr & " and userclass='"&request("userclass")&"'"
		'	end if
		'end if

		'======shinzeal加入特殊搜索=======
		dim Tsqlstr
		if request("loginT")<>"" then
		   	if request("loginR")="more" then
			 Tsqlstr=" userlogins >= "&CheckNumeric(request("loginT"))
			else
			 Tsqlstr=" userlogins <= "&CheckNumeric(request("loginT"))
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if 
		end if

		if request("vanishT")<>"" then
		   	if request("vanishR")="more" then
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,lastlogin,"&SqlNowString&") >= "&CheckNumeric(request("vanishT"))
				Else
					Tsqlstr=" datediff('d',lastlogin,"&SqlNowString&") >= "&CheckNumeric(request("vanishT"))
				End If
			else
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,lastlogin,"&SqlNowString&") <= "&CheckNumeric(request("vanishT"))
				Else
					Tsqlstr=" datediff('d',lastlogin,"&SqlNowString&") <= "&CheckNumeric(request("vanishT"))
				End If
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if 
		end if

		if request("regT")<>"" then
		   	if request("regR")="more" then
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,JoinDate,"&SqlNowString&") >= "&CheckNumeric(request("regT"))
				Else
					Tsqlstr=" datediff('d',JoinDate,"&SqlNowString&") >= "&CheckNumeric(request("regT"))
				End If
			else
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,JoinDate,"&SqlNowString&") <= "&CheckNumeric(request("regT"))
				Else
					Tsqlstr=" datediff('d',JoinDate,"&SqlNowString&") <= "&CheckNumeric(request("regT"))
				End If
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if 
		end if

		if request("artcleT")<>"" then
		   	if request("artcleR")="more" then
			 Tsqlstr=" UserPost >= "&CheckNumeric(request("artcleT"))
			else
			 Tsqlstr=" UserPost <= "&CheckNumeric(request("artcleT"))
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if 
		end if

		if request("UWealth_value")<>"" then
			if request("UWealth")="more" then
				Tsqlstr=" userWealth >= "&CheckNumeric(Request("UWealth_value"))
			else
				Tsqlstr=" userWealth <= "&CheckNumeric(Request("UWealth_value"))
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if
		end if

		if request("UEP_value")<>"" then
			if request("UEP")="more" then
				Tsqlstr=" userEP >= "&CheckNumeric(Request("UEP_value"))
			else
				Tsqlstr=" userEP <= "&CheckNumeric(Request("UEP_value"))
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if
		end if

		if request("UCP_value")<>"" then
			if request("UCP")="more" then
				Tsqlstr=" userCP >= "&CheckNumeric(Request("UCP_value"))
			else
				Tsqlstr=" userCP <= "&CheckNumeric(Request("UCP_value"))
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if
		end if

		if request("UPower_value")<>"" then
			if request("UPower")="more" then
				Tsqlstr=" UserPower >= "&CheckNumeric(Request("UPower_value"))
			else
				Tsqlstr=" UserPower <= "&CheckNumeric(Request("UPower_value"))
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if
		end if

		if request("UMoney_value")<>"" then
			if request("UMoney")="more" then
				Tsqlstr=" UserMoney >= "&CheckNumeric(Request("UMoney_value"))
			else
				Tsqlstr=" UserMoney <= "&CheckNumeric(Request("UMoney_value"))
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if
		end if

		if request("UTicket_value")<>"" then
			if request("UTicket")="more" then
				Tsqlstr=" UserTicket >= "&CheckNumeric(Request("UTicket_value"))
			else
				Tsqlstr=" UserTicket <= "&CheckNumeric(Request("UTicket_value"))
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if
		end if

		if request("UVipStarTime_value")<>"" then
		   	if request("UVipStarTime")="more" then
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,Vip_StarTime,"&SqlNowString&") >= "&CheckNumeric(request("UVipStarTime_value"))
				Else
					Tsqlstr=" datediff('d',Vip_StarTime,"&SqlNowString&") >= "&CheckNumeric(request("UVipStarTime_value"))
				End If
			else
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,Vip_StarTime,"&SqlNowString&") <= "&CheckNumeric(request("UVipStarTime_value"))
				Else
					Tsqlstr=" datediff('d',Vip_StarTime,"&SqlNowString&") <= "&CheckNumeric(request("UVipStarTime_value"))
				End If
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if 
		end if
		if request("UVipEndTime_value")<>"" then
		   	if request("UVipEndTime")="more" then
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,Vip_EndTime,"&SqlNowString&") >= "&CheckNumeric(request("UVipEndTime_value"))
				Else
					Tsqlstr=" datediff('d',Vip_EndTime,"&SqlNowString&") >= "&CheckNumeric(request("UVipEndTime_value"))
				End If
			else
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,Vip_EndTime,"&SqlNowString&") <= "&CheckNumeric(request("UVipEndTime_value"))
				Else
					Tsqlstr=" datediff('d',Vip_EndTime,"&SqlNowString&") <= "&CheckNumeric(request("UVipEndTime_value"))
				End If
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & " and " & Tsqlstr
			end if 
		end if

		'======特殊搜索结束======
		if request("useremail")<>"" then
			if sqlstr="" then
			sqlstr=" useremail like '%"&request("useremail")&"%'"
			else
			sqlstr=sqlstr & " and useremail like '%"&request("useremail")&"%'"
			end if
		end if
		if request("userim")<>"" then
			if sqlstr="" then
			sqlstr=" UserIM like '%"&request("userim")&"%'"
			else
			sqlstr=sqlstr & " and UserIM like '%"&request("userim")&"%'"
			end if
		end if
		if request("lastip")<>"" then
			if sqlstr="" then
			sqlstr=" UserLastIP like '%"&request("lastip")&"%'"
			else
			sqlstr=sqlstr & " and UserLastIP like '%"&request("lastip")&"%'"
			end if
		end if
		if request("userinfo")<>"" then
			if sqlstr="" then
			sqlstr=" UserInfo like '%"&request("userinfo")&"%'"
			else
			sqlstr=sqlstr & " and UserInfo like '%"&request("userinfo")&"%'"
			end if
		end if
		'修正不能用头衔搜索 2005-4-9 Dv.Yz
		If Request("usertitle") <> "" Then
			If Sqlstr = "" Then
				Sqlstr = " UserTitle LIKE '%" & Request("usertitle") & "%'"
			Else
				Sqlstr = Sqlstr & " AND UserTitle LIKE '%" & Request("usertitle") & "%'"
			End If
		End If
		if request("sign")<>"" then
			if sqlstr="" then
			sqlstr=" usersign like '%"&request("sign")&"%'"
			else
			sqlstr=sqlstr & " and usersign like '%"&request("sign")&"%'"
			end if
		end if

		If Sqlstr = "" Then
			Response.Write "<tr><td colspan=8 class=td1>请指定搜索参数！</td></tr>"
			Response.End
		End If
		If Request("Searchmax") = "" Or Not Isnumeric(Request("Searchmax")) Then
			Sql = "SELECT TOP 1 "& Sql &" FROM [Dv_User] WHERE " & Sqlstr & " ORDER BY UserID DESC"
		Else
			Sql = "SELECT TOP " & Request("Searchmax") & Sql &" FROM [Dv_User] WHERE " & Sqlstr & " ORDER BY UserID DESC"
		End If
	case else
		Response.Write "<tr><td colspan=8 class=td1>错误的参数。</td></tr>"
		Response.End
	End Select
	'Response.Write sql
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=8 class=td1>没有找到相关记录。"
		If Request("userSearch")="15" Then
			Response.Write "（若未添加VIP用户组，请<a href=""group.asp""><font color=red>点击进入论坛用户组管理</font></a>进行添加。）"
		End If
		Response.Write "</td></tr>"
	else
%>
<FORM METHOD=POST ACTION="?action=touser">
<tr align=center height=23>
<td class=td2 width="10%"><B>用户名</B></td>
<td class=td2 width="15%"><B>Email</B></td>
<td class=td2 width="8%"><B>权限</B></td>
<td class=td2 width="8%"><B>数据修复</B></td>
<td class=td2 width="15%"><B>最后IP</B></td>
<td class=td2 width="15%"><B>最后登录</B></td>
<td class=td2 width="20%"><B>登记/终止日期</B></td>
<td class=td2><B>操作</B></td>
</tr>
<%
		rs.PageSize = Cint(Dvbbs.Forum_Setting(11))
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = Cint(Dvbbs.Forum_Setting(11)))
%>
<tr>
<td class=td1><a href="?action=modify&userid=<%=rs("userid")%>"><%=rs("username")%></a></td>
<td class=td1><a href="mailto:<%=rs("useremail")%>"><%=rs("useremail")%></a></td>
<td class=td1 align=center><a href="?action=UserPermission&userid=<%=rs("userid")%>&username=<%=rs("username")%>">编辑</a></td>
<td class=td1 align=center><a href="?action=fix&userid=<%=rs("userid")%>&username=<%=rs("username")%>">修复</a></td>
<td class=td1><a href="lockIP.asp?userip=<%=rs("UserLastIP")%>" title="点击锁定该用户IP"><%=rs("userlastip")%></a>&nbsp;</td>
<td class=td1><%if rs("lastlogin")<>"" and isdate(rs("lastlogin")) then%><%=rs("lastlogin")%><%end if%></td>
<td class=td1 align=center>
<%=rs("Vip_StarTime")%>/
<%=rs("Vip_EndTime")%>
</td>
<td class=td1 align=center><input type="checkbox" class=checkbox name="userid" value="<%=rs("userid")%>" <%if rs("userGroupid")=1 then response.write "disabled"%>></td>
</tr>
<%
		page_count = page_count + 1
		rs.movenext
		wend
		Pcount=rs.PageCount
%>
<tr><td colspan=8 class=td1 align=center>分页：
<%
Dim Searchstr,i
'修正头衔搜索用户的分页错误。
'修正最后登陆IP搜索用户的分页错误 2005.10.12 By Winder
Searchstr = "?userSearch=" & Request("userSearch") & "&username=" & Request("username") & "&useremail=" & Request("useremail") & "&userim=" & Request("userim") & "&lastip=" & Request("lastip") & "&usertitle=" & Request("usertitle") & "&sign=" & Request("sign") & "&userinfo=" & Request("userinfo") & "&action=" & Request("action") & "&loginR=" & Request("loginR") & "&loginT=" & Request("loginT") & "&vanishR=" & Request("vanishR") & "&vanishT=" & Request("vanishT") & "&regR=" & Request("regR") & "&regT=" & Request("regT") & "&artcleR=" & Request("artcleR") & "&artcleT=" & Request("artcleT") & "&UWealth=" & Request("UWealth") & "&UWealth_value=" & Request("UWealth_value") & "&UEP=" & Request("UEP") & "&UEP_value=" & Request("UEP_value") & "&UCP=" & Request("UCP") & "&UCP_value=" & Request("UCP_value") & "&UPower=" & Request("UPower") & "&UPower_value=" & Request("UPower_value") & "&UMoney=" & Request("UMoney") & "&UMoney_value=" & Request("UMoney_value") & "&UTicket=" & Request("UTicket") & "&UTicket_value=" & Request("UTicket_value") & "&searchmax=" & Request("searchmax") & "&UVipStarTime=" & Request("UVipStarTime") & "&UVipStarTime_value=" & Request("UVipStarTime_value") & "&UVipEndTime=" & Request("UVipEndTime") & "&UVipEndTime_value=" & Request("UVipEndTime_value")&"&usergroups="&Request("usergroups")&"&usergroupid="&Request("usergroupid")

	if currentpage > 4 then
		response.write "<a href="""&Searchstr&"&page=1"">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
		endpage=currentpage+3
	else
		endpage=Pcount
	end if
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color=red>["&i&"]</font>"
		else
        response.write " <a href="""&Searchstr&"&page="&i&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
		response.write "... <a href="""&Searchstr&"&page="&Pcount&""">["&Pcount&"]</a>"
	end if
%>
</td></tr>
<tr><td colspan=5 class=td1 align=center><B>请选择您需要进行的操作</B>：<input type="radio" class=radio name="useraction" value=1> 删除&nbsp;&nbsp;<input type="radio" class=radio name="useraction" value=3> 删除用户所有帖子&nbsp;&nbsp;<input type="radio" class=radio name="useraction" value=2 checked> 移动到用户组
<select size=1 name="selusergroup">
<%
set trs=Dvbbs.Execute("select usergroupid,UserTitle,ParentGID from dv_usergroups where not (usergroupid=1 or usergroupid=7) and (Not ParentGID=0) order by ParentGID,usergroupid")
do while not trs.eof
response.write "<option value="&trs(0)&">"&SysGroupName(tRs(2))&trs(1)&"</option>"
trs.movenext
loop
trs.close
set trs=nothing
%>
</select>
</td>
<td class=td1 colspan=8 align=center>全部选定<input type=checkbox class=checkbox value="on" name="chkall" onclick="CheckAll(this.form)">
</td>
</tr>
<tr><td colspan=8 class=td1 align=center>
<input type=submit class=button name=submit value="执行选定的操作"  onclick="{if(confirm('确定执行选择的操作吗?')){return true;}return false;}">
</td></tr>
</FORM>
<%
	end if
	rs.close
	set rs=nothing
End Sub

'操作用户，删除用户信息相关操作
Sub ToUser()
	Dim SQL,rs
	response.write "<tr><th colspan=8 style=""text-align:center;"">执行结果</th></tr>"
	if request("useraction")="" then
		response.write "<tr><td colspan=8 class=td1>请指定相关参数。</td></tr>"
		founderr=true
	end if
	if request("userid")="" then
		response.write "<tr><td colspan=8 class=td1>请选择相关用户。</td></tr>"
		founderr=true
	end if
	if not founderr then
		if request("useraction")=1 then
			Dim AllUserName
			AllUserName = ""
			'------------------shinzeal加入删除用户的短信-------------------------
			dim uid,i
			for i=1 to request("userid").count
				if request("userid").count=1 then
				uID=CLng( request("userid") )
				else
				uID=replace(request.form("userid")(i),"'","")
				end if
				set rs=Dvbbs.Execute("select username from [dv_User] where userid="&uid&"")
				if not (rs.eof and rs.bof) then
					AllUserName = AllUserName & Rs(0) & ","
					Dvbbs.Execute("update dv_message set delR=1 where incept='"&trim(rs(0))&"' and delR=0")
					Dvbbs.Execute("update dv_message set delS=1 where sender='"&trim(rs(0))&"' and delS=0 and issend=0")
					Dvbbs.Execute("update dv_message set delS=1 where sender='"&trim(rs(0))&"' and delS=0 and issend=1")
					Dvbbs.Execute("delete from dv_message where incept='"&rs(0)&"' and delR=1") 
					Dvbbs.Execute("update dv_message set delS=2 where sender='"&trim(rs(0))&"' and delS=1")
					Dvbbs.Execute("delete from dv_friend where F_username='"&rs(0)&"'") 
					Dvbbs.Execute("delete from dv_bookmark where username='"&rs(0)&"'") 
				end if 
				rs.close
			next
			If Right(AllUserName,1) = "," Then AllUserName = Left(AllUserName,Len(AllUserName)-1)
			'-------------------删除用户的短信------------------------
			'删除用户的帖子和精华
			Dvbbs.Execute("delete from dv_topic where PostUserID in ("&replace(request("userid"),"'","")&")")
			for i=0 to ubound(allposttable)
				Dvbbs.Execute("delete from "&allposttable(i)&" where PostUserID in ("&replace(request("userid"),"'","")&")")
			next
			Dvbbs.Execute("delete from dv_besttopic where PostUserID in ("&replace(request("userid"),"'","")&")")
			'删除用户上传表
			Dvbbs.Execute("delete from dv_upfile where F_UserID in ("&replace(request("userid"),"'","")&")")
			Dvbbs.Execute("delete from [dv_user] where userid in ("&replace(request("userid"),"'","")&")")
			Response.write "<tr><td colspan=8 class=td1>删除用户（ "& AllUserName &" ）操作成功。</td></tr>"

			'-----------------------------------------------------------------
			'系统整合
			'-----------------------------------------------------------------
			Dim DvApi_Obj,DvApi_SaveCookie,SysKey
			If DvApi_Enable Then
				'SysKey = Md5(DvApi_SysKey&AllUserName,16)
				Set DvApi_Obj = New DvApi
					DvApi_Obj.NodeValue "syskey",SysKey,0,False
					DvApi_Obj.NodeValue "action","delete",0,False
					DvApi_Obj.NodeValue "username",AllUserName,1,False
					Md5OLD = 1
					SysKey = Md5(DvApi_Obj.XmlNode("username")&DvApi_SysKey,16)
					Md5OLD = 0
					DvApi_Obj.NodeValue "syskey",SysKey,0,False
					DvApi_Obj.SendHttpData
					'If DvApi_Obj.Status = "1" Then
						'Response.redirect "showerr.asp?ErrCodes="& DvApi_Obj.Message &"&action=OtherErr"
					'End If
				Set DvApi_Obj = Nothing
			End If
			'-----------------------------------------------------------------
		elseif request("useraction")=2 then
			dim userclass,usertitlepic
			set rs=Dvbbs.Execute("select * from dv_usergroups where usergroupid="&request("selusergroup")&" order by minarticle")
			if not (rs.eof and rs.bof) then
				userclass=rs("usertitle")
				usertitlepic=rs("grouppic")
			end if
			Dvbbs.Execute("update [dv_user] set UserGroupID="&replace(request("selusergroup"),"'","")&",userclass='"&userclass&"',titlepic='"&usertitlepic&"' where userid in ("&replace(request("userid"),"'","")&")")
			response.write "<tr><td colspan=8 class=td1>操作成功。</td></tr>"
		elseif request("useraction")=3 then
			dim titlenum
			if request("userid")="" then
				response.write "<tr><td colspan=8 class=td1>请输入被删除帖子用户名。</td></tr>"
			end if
			titlenum=0
			for i=0 to ubound(allposttable)
			set rs=Dvbbs.Execute("Select Count(announceID) from "&allposttable(i)&" where postuserid in ("&replace(request("userid"),"'","")&")") 
   			titlenum=titlenum+rs(0)
			sql="update "&allposttable(i)&" set locktopic=boardid,boardid=444,isbest=0 where postuserid in ("&replace(request("userid"),"'","")&")"
			Dvbbs.Execute(sql)
			next
			Dvbbs.Execute("delete from dv_besttopic where postuserid in ("&replace(request("userid"),"'","")&")")
			set rs=Dvbbs.Execute("select topicid,posttable from dv_topic where postuserid in ("&replace(request("userid"),"'","")&")")
			do while not rs.eof
			Dvbbs.Execute("update "&rs(1)&" set locktopic=boardid,boardid=444,isbest=0 where rootid="&rs(0))
			rs.movenext
			loop
			set rs=nothing
			Dvbbs.Execute("update dv_topic set locktopic=boardid,boardid=444,isbest=0 where postuserid in ("&replace(request("userid"),"'","")&")")
			if isnull(titlenum) then titlenum=0
			sql="update [dv_user] set UserPost=UserPost-"&titlenum&",userWealth=userWealth-"&titlenum*Dvbbs.Forum_user(3)&",userEP=userEP-"&titlenum*Dvbbs.Forum_user(8)&",userCP=userCP-"&titlenum*Dvbbs.Forum_user(13)&" where userid in ("&replace(request("userid"),"'","")&")"
			Dvbbs.Execute(sql)
			response.write "<tr><td colspan=8 class=td1>删除成功，如果要完全删除帖子请到论坛回收站<BR>建议您到更新论坛数据中更新一下论坛数据，或者<a href=alldel.asp>返回</a></td></tr>"
		else
			response.write "<tr><td colspan=8 class=td1>错误的参数。</td></tr>"
		end if
	end if
End Sub

'修改用户资料表单
Sub UserModify()
dim realname,character,personal,country,province,city,shengxiao,blood,belief,occupation,marital, education,college,userphone,iaddress
Dim UserIM
Dim rs,sql
	response.write "<tr><th colspan=8 style=""text-align:center;"">用户资料操作</th></tr>"
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=8 class=td1>错误的用户参数。</td></tr>"
		founderr=true
	end if
	if not founderr then
		Set rs= Dvbbs.iCreateObject("ADODB.Recordset")
		sql="select * from [dv_user] where userid="&request("userid")
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
		response.write "<tr><td colspan=8 class=td1>没有找到相关用户。</td></tr>"
		founderr=true
		else
if rs("userinfo")<>"" then
	userinfo=split(Server.HtmlEncode(rs("userinfo")),"|||")
	if ubound(userinfo)=14 then
		realname=userinfo(0)
		character=userinfo(1)
		personal=userinfo(2)
		country=userinfo(3)
		province=userinfo(4)
		city=userinfo(5)
		shengxiao=userinfo(6)
		blood=userinfo(7)
		belief=userinfo(8)
		occupation=userinfo(9)
		marital=userinfo(10)
		education=userinfo(11)
		college=userinfo(12)
		userphone=userinfo(13)
		iaddress=userinfo(14)
	else
		realname=""
		character=""
		personal=""
		country=""
		province=""
		city=""
		shengxiao=""
		blood=""
		belief=""
		occupation=""
		marital=""
		education=""
		college=""
		userphone=""
		iaddress=""
	end if
else
	realname=""
	character=""
	personal=""
	country=""
	province=""
	city=""
	shengxiao=""
	blood=""
	belief=""
	occupation=""
	marital=""
	education=""
	college=""
	userphone=""
	iaddress=""
end if
UserIM = Split(Rs("UserIM"),"|||")
%>
<FORM METHOD=POST ACTION="?action=saveuserinfo">
<tr>
<td width=100% class=td1 valign=top colspan=8>对 <%=rs("username")%> 用户操作快捷选项：<BR><BR>
<a href="mailto:<%=rs("useremail")%>">发邮件</a> | <a href="../messanger.asp?action=new&touser=<%=rs("username")%>" target=_blank>发短信</a> | <a href="../dispuser.asp?id=<%=rs("userid")%>" target=_blank>预览用户资料</a> | <a href="../Query.asp?stype=1&nSearch=3&keyword=<%=rs("username")%>&SearchDate=30" target=_blank>用户新贴</a> | <a href="../Query.asp?stype=6&nSearch=0&pSearch=0&keyword=<%=rs("username")%>" target=_blank>用户精华</a> | <a href="../Query.asp?stype=4&nSearch=0&pSearch=0&keyword=<%=rs("username")%>" target=_blank>用户热贴</a> | <a href="../show.asp?username=<%=rs("username")%>" target=_blank>用户展区</a> | <a href="?action=UserPermission&userid=<%=rs("userid")%>&username=<%=rs("username")%>">编辑权限</a> | <a href="../TopicOther.asp?action=lookip&ip=<%=Rs("UserLastIP")%>&t=1" target=_blank>最后来源</a> | <a href="?action=touser&useraction=1&userid=<%=rs("userid")%>" onclick="{if(confirm('删除将不可恢复，并且将删除该用户在论坛的所有信息，确定删除吗?')){return true;}return false;}">删除用户</a>
</td>
</tr>
<tr><th colspan=6 style="text-align:center;">用户基本资料修改－－<%=rs("username")%></th></tr>
<tr><td class=td1 height=23 align=left colspan=6>
注意：新建管理员建议到管理员管理中进行，仅在此设置为管理员组的用户并无进入系统后台权限
</td></tr>
<tr>
<td width=20% class=td1>用户组</td>
<td width=80% class=td1 colspan=5>
<select size=1 name="usergroups">
<%
set trs=Dvbbs.Execute("select usergroupid,UserTitle,parentgid from dv_usergroups where Not ParentGID=0 order by ParentGID,usergroupid")
do while not trs.eof
response.write "<option value="&trs(0)
if rs("usergroupid")=trs(0) then response.write " selected "
response.write ">"&SysGroupName(tRs(2)) & trs(1)
'if trs(2)>0 then response.write "(自定义等级)"
response.write "</option>"
trs.movenext
loop
trs.close
set trs=nothing

%>
</select>
</td>
</tr>
<input name="userid" type=hidden value="<%=rs("userid")%>">
<tr>
<td width=20% class=td1>用户名</td>
<td width=80% class=td1 colspan=5><input size=45 name="username" type=text value="<%=Server.HtmlEncode(rs("username"))%>" disabled></td>
</tr>
<tr>
<td width=20% class=td1>密  码</td>
<td width=80% class=td1 colspan=5><input size=45 name="password" type=text>&nbsp;如果不修改请留空</td>
</tr>
<tr>
<td width=20% class=td1>密码问题</td>
<td width=80% class=td1 colspan=5><input size=45 name="quesion" type=text value="<%If Trim(rs("userquesion"))<>"" Then Response.Write Server.HtmlEncode(rs("userquesion"))%>"></td>
</tr>
<tr>
<td width=20% class=td1>密码答案</td>
<td width=80% class=td1 colspan=5><input size=45 name="answer" type=text>&nbsp;如果不修改请留空</td>
</tr><tr>
<td width=20% class=td1>用户性别</td>
<td width=80% class=td1 colspan=5>
女 <input type="radio" class=radio value="0" <%if rs("UserSex")=0 then%>checked<%end if%> name="sex">&nbsp;
男 <input type="radio" class=radio value="1" <%if rs("UserSex")=1 then%>checked<%end if%> name="sex">&nbsp;
</td>
</tr>
<tr>
<td width=20% class=td1>个人照片</td>
<td width=80% class=td1 colspan=5><input size=45 name="UserPhoto" type=text value="<%If Trim(rs("UserPhoto"))<>"" Then Response.Write Server.HtmlEncode(rs("UserPhoto"))%>"></td>
</tr>
<tr>
<td width=20% class=td1>Email</td>
<td width=80% class=td1 colspan=5><input size=45 name="userEmail" type=text value="<%If Trim(rs("useremail"))<>"" Then Response.Write Server.HtmlEncode(rs("useremail"))%>"></td>
</tr>
<tr>
<td width=20% class=td1>个人主页</td>
<td width=80% class=td1 colspan=5><input size=45 name="homepage" type=text value="<%=Server.HtmlEncode(UserIM(0))%>"></td>
</tr>
<tr>
<td width=20% class=td1>头像</td>
<td width=80% class=td1 colspan=5><input size=45 name="face" type=text value="<%If Trim(Rs("UserFace"))<>"" Then Response.Write Server.HtmlEncode(Split(rs("userface"),"|")(Ubound(Split(rs("userface"),"|"))))%>">&nbsp;宽度：<input size=3 name="width" type=text value="<%=rs("userwidth")%>">&nbsp;高度：<input size=3 name="height" type=text value="<%=rs("userheight")%>"></td>
</tr>
<tr>
<td width=20% class=td1>OICQ</td>
<td width=80% class=td1 colspan=5><input size=45 name="oicq" type=text value="<%=Server.HtmlEncode(UserIM(1))%>"></td>
</tr>
<tr>
<td width=20% class=td1>ICQ</td>
<td width=80% class=td1 colspan=5><input size=45 name="icq" type=text value="<%=Server.HtmlEncode(UserIM(2))%>"></td>
</tr>
<tr>
<td width=20% class=td1>MSN</td>
<td width=80% class=td1 colspan=5><input size=45 name="msn" type=text value="<%=Server.HtmlEncode(UserIM(3))%>"></td>
</tr>
<tr>
<td width=20% class=td1>AIM</td>
<td width=80% class=td1 colspan=5><input size=45 name="aim" type=text value="<%=Server.HtmlEncode(UserIM(4))%>"></td>
</tr>
<tr>
<td width=20% class=td1>YaHoo</td>
<td width=80% class=td1 colspan=5><input size=45 name="yahoo" type=text value="<%=Server.HtmlEncode(UserIM(5))%>"></td>
</tr>
<tr>
<td width=20% class=td1>UC</td>
<td width=80% class=td1 colspan=5><input size=45 name="uc" type=text value="<%=Server.HtmlEncode(UserIM(6))%>"></td>
</tr>
<tr>
<td width=20% class=td1>头衔</td>
<td width=80% class=td1 colspan=5><input size=45 name="usertitle" type=text value="<%If Trim(Rs("UserTitle"))<>"" Then Response.Write Server.HtmlEncode(rs("usertitle"))%>"></td>
</tr>
<tr><th colspan=6 style="text-align:center;">用户分值资料修改</th></tr>
<tr>
<td width=20% class=td1>发表文章</td>
<td width=80% class=td1 colspan=5><input size=45 name="article" type=text value="<%=rs("UserPost")%>"></td>
</tr>
<tr>
<td width=20% class=td1>被删文章</td>
<td width=80% class=td1 colspan=5><input size=45 name="Userdel" type=text value="<%=rs("userdel")%>"></td>
</tr>
<tr>
<td width=20% class=td1>精华文章</td>
<td width=80% class=td1 colspan=5><input size=45 name="userisbest" type=text value="<%=rs("userisbest")%>"></td>
</tr>
<tr>
<td width=20% class=td1>金币</td>
<td width=80% class=td1 colspan=5><input size=45 name="usermoney" type=text value="<%=rs("usermoney")%>"></td>
</tr>
<tr>
<td width=20% class=td1>点券</td>
<td width=80% class=td1 colspan=5><input size=45 name="UserTicket" type=text value="<%=rs("UserTicket")%>"></td>
</tr>
<tr>
<td width=20% class=td1>金钱</td>
<td width=80% class=td1 colspan=5><input size=45 name="userwealth" type=text value="<%=rs("userwealth")%>"></td>
</tr>
<tr>
<td width=20% class=td1>经验</td>
<td width=80% class=td1 colspan=5><input size=45 name="userep" type=text value="<%=rs("userep")%>"></td>
</tr>
<tr>
<td width=20% class=td1>魅力</td>
<td width=80% class=td1 colspan=5><input size=45 name="usercp" type=text value="<%=rs("usercp")%>"></td>
</tr>
<tr>
<td width=20% class=td1>威望</td>
<td width=80% class=td1 colspan=5><input size=45 name="userpower" type=text value="<%=rs("userpower")%>"></td>
</tr>
<tr><th colspan=6 style="text-align:center;">日期相关</th></tr>
<tr>
<td width=20% class=td1>生日</td>
<td width=80% class=td1 colspan=5><input size=45 name="birthday" type=text value="<%=rs("userbirthday")%>">&nbsp;格式：2001-2-2</td>
</tr>
<tr>
<td width=20% class=td1>注册时间</td>
<td width=80% class=td1 colspan=5><input size=45 name="adddate" type=text value="<%=rs("JoinDate")%>"></td>
</tr>
<tr>
<td width=20% class=td1>最后登录</td>
<td width=80% class=td1 colspan=5><input size=45 name="lastlogin" type=text value="<%=rs("lastlogin")%>"></td>
</tr>
<tr><th colspan=6 style="text-align:center;">用户详细资料</th></tr>
<tr>
<td width=20% class=td1>真实姓名</td>
<td width=80% class=td1 colspan=5><input size=45 name="realname" type=text value="<%=realname%>"></td>
</tr>
<tr>
<td width=20% class=td1>国　　家</td>
<td width=80% class=td1 colspan=5><input size=45 name="country" type=text value="<%=country%>"></td>
</tr>
<tr>
<td width=20% class=td1>联系电话</td>
<td width=80% class=td1 colspan=5><input size=45 name="userphone" type=text value="<%=userphone%>"></td>
</tr><tr>
<td width=20% class=td1>通信地址</td>
<td width=80% class=td1 colspan=5><input size=45 name="address" type=text value="<%=iaddress%>"></td>
</tr>
<tr>
<td width=20% class=td1>省　　份</td>
<td width=80% class=td1 colspan=5><input size=45 name="province" type=text value="<%=province%>"></td>
</tr>
<tr>
<td width=20% class=td1>城　　市</td>
<td width=80% class=td1 colspan=5><input size=45 name="city" type=text value="<%=city%>"></td>
</tr><tr>
<td width=20% class=td1>生　　肖</td>
<td width=80% class=td1 colspan=5>
<select size=1 name=shengxiao>
<option <%if shengxiao="" then%>selected<%end if%>></option>
<option value=鼠 <%if shengxiao="鼠" then%>selected<%end if%>>鼠</option>
<option value=牛 <%if shengxiao="牛" then%>selected<%end if%>>牛</option>
<option value=虎 <%if shengxiao="虎" then%>selected<%end if%>>虎</option>
<option value=兔 <%if shengxiao="兔" then%>selected<%end if%>>兔</option>
<option value=龙 <%if shengxiao="龙" then%>selected<%end if%>>龙</option>
<option value=蛇 <%if shengxiao="蛇" then%>selected<%end if%>>蛇</option>
<option value=马 <%if shengxiao="马" then%>selected<%end if%>>马</option>
<option value=羊 <%if shengxiao="羊" then%>selected<%end if%>>羊</option>
<option value=猴 <%if shengxiao="猴" then%>selected<%end if%>>猴</option>
<option value=鸡 <%if shengxiao="鸡" then%>selected<%end if%>>鸡</option>
<option value=狗 <%if shengxiao="狗" then%>selected<%end if%>>狗</option>
<option value=猪 <%if shengxiao="猪" then%>selected<%end if%>>猪</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=td1>血　　型</td>
<td width=80% class=td1 colspan=5>
<select size=1 name=blood>
<option <%if blood="" then%>selected<%end if%>></option>
<option value=A <%if blood="A" then%>selected<%end if%>>A</option>
<option value=B <%if blood="B" then%>selected<%end if%>>B</option>
<option value=AB <%if blood="AB" then%>selected<%end if%>>AB</option>
<option value=O <%if blood="O" then%>selected<%end if%>>O</option>
<option value=其他 <%if blood="其他" then%>selected<%end if%>>其他</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=td1>信　　仰</td>
<td width=80% class=td1 colspan=5>
<select size=1 name=belief>
<option <%if belief="" then%>selected<%end if%>></option>
<option value=佛教 <%if belief="佛教" then%>selected<%end if%>>佛教</option>
<option value=道教 <%if belief="道教" then%>selected<%end if%>>道教</option>
<option value=基督教 <%if belief="基督教" then%>selected<%end if%>>基督教</option>
<option value=天主教 <%if belief="天主教" then%>selected<%end if%>>天主教</option>
<option value=回教 <%if belief="回教" then%>selected<%end if%>>回教</option>
<option value=无神论者 <%if belief="无神论者" then%>selected<%end if%>>无神论者</option>
<option value=共产主义者 <%if belief="共产主义者" then%>selected<%end if%>>共产主义者</option>
<option value=其他 <%if belief="其他" then%>selected<%end if%>>其他</option>
</select>
</td>
</tr><tr>
<td width=20% class=td1>职　　业</td>
<td width=80% class=td1 colspan=5>
<select name=occupation>
<option <%if occupation="" then%>selected<%end if%>> </option>
<option value="财会/金融" <%if occupation="财会/金融" then%>selected<%end if%>>财会/金融</option>
<option value=工程师 <%if occupation="工程师" then%>selected<%end if%>>工程师</option>
<option value=顾问 <%if occupation="顾问" then%>selected<%end if%>>顾问</option>
<option value=计算机相关行业 <%if occupation="计算机相关行业" then%>selected<%end if%>>计算机相关行业</option>
<option value=家庭主妇 <%if occupation="家庭主妇" then%>selected<%end if%>>家庭主妇</option>
<option value="教育/培训" <%if occupation="教育/培训" then%>selected<%end if%>>教育/培训</option>
<option value="客户服务/支持" <%if occupation="客户服务/支持" then%>selected<%end if%>>客户服务/支持</option>
<option value="零售商/手工工人" <%if occupation="零售商/手工工人" then%>selected<%end if%>>零售商/手工工人</option>
<option value=退休 <%if occupation="退休" then%>selected<%end if%>>退休</option>
<option value=无职业 <%if occupation="无职业" then%>selected<%end if%>>无职业</option>
<option value="销售/市场/广告" <%if occupation="销售/市场/广告" then%>selected<%end if%>>销售/市场/广告</option>
<option value=学生 <%if occupation="学生" then%>selected<%end if%>>学生</option>
<option value=研究和开发 <%if occupation="研究和开发" then%>selected<%end if%>>研究和开发</option>
<option value="一般管理/监督" <%if occupation="一般管理/监督" then%>selected<%end if%>>一般管理/监督</option>
<option value="政府/军队" <%if occupation="政府/军队" then%>selected<%end if%>>政府/军队</option>
<option value="执行官/高级管理" <%if occupation="执行官/高级管理" then%>selected<%end if%>>执行官/高级管理</option>
<option value="制造/生产/操作" <%if occupation="制造/生产/操作" then%>selected<%end if%>>制造/生产/操作</option>
<option value=专业人员 <%if occupation="专业人员" then%>selected<%end if%>>专业人员</option>
<option value="自雇/业主" <%if occupation="自雇/业主" then%>selected<%end if%>>自雇/业主</option>
<option value=其他 <%if occupation="其他" then%>selected<%end if%>>其他</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=td1>婚姻状况</td>
<td width=80% class=td1 colspan=5>
<select size=1 name=marital>
<option <%if marital="" then%>selected<%end if%>></option>
<option value=未婚 <%if marital="未婚" then%>selected<%end if%>>未婚</option>
<option value=已婚 <%if marital="已婚" then%>selected<%end if%>>已婚</option>
<option value=离异 <%if marital="离异" then%>selected<%end if%>>离异</option>
<option value=丧偶 <%if marital="丧偶" then%>selected<%end if%>>丧偶</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=td1>最高学历</td>
<td width=80% class=td1 colspan=5>
<select size=1 name=education>
<option <%if education="" then%>selected<%end if%>></option>
<option value=小学 <%if education="小学" then%>selected<%end if%>>小学</option>
<option value=初中 <%if education="初中" then%>selected<%end if%>>初中</option>
<option value=高中 <%if education="高中" then%>selected<%end if%>>高中</option>
<option value=大学 <%if education="大学" then%>selected<%end if%>>大学</option>
<option value=硕士 <%if education="硕士" then%>selected<%end if%>>硕士</option>
<option value=博士 <%if education="博士" then%>selected<%end if%>>博士</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=td1>毕业院校</td>
<td width=80% class=td1 colspan=5><input size=45 name="college" type=text value="<%=college%>"></td>
</tr>
<tr>
<td width=20% class=td1>性　格</td>
<td width=80% class=td1 colspan=5>
<textarea name=character rows=4 cols=80><%=character%></textarea>
</td>
</tr><tr>
<td width=20% class=td1>个人简介</td>
<td width=80% class=td1 colspan=5>
<textarea name=personal rows=4 cols=80><%=personal%></textarea>
</td>
</tr><tr>
<td width=20% class=td1>用户签名</td>
<td width=80% class=td1 colspan=5>
<textarea name="sign" rows=4 cols=80><%If Trim(Rs("UserSign"))<>"" Then Response.Write Server.HtmlEncode(rs("usersign"))%></textarea>
</td>
</tr>
<tr><th colspan=6 style="text-align:center;">用户设置</th></tr>
<tr>
<td width=20% class=td1>用户状态</td>
<td width=80% class=td1 colspan=5>
正常 <input type="radio" class=radio value="0" <%if rs("lockuser")=0 then%>checked<%end if%> name="lockuser">&nbsp;
锁定 <input type="radio" class=radio value="1" <%if rs("lockuser")=1 then%>checked<%end if%> name="lockuser">&nbsp;
屏蔽 <input type="radio" class=radio value="2" <%if rs("lockuser")=2 then%>checked<%end if%> name="lockuser">
</td>
</tr>
<!--
<tr>
<td width=20% class=td1>阳光会员</td>
<td width=80% class=td1 colspan=5>
关闭 <input type="radio" class=radio value="0" <%if rs("IsChallenge")=0 then%>checked<%end if%> name="IsChallenge">&nbsp;
开启 <input type="radio" class=radio value="1" <%if rs("IsChallenge")=1 then%>checked<%end if%> name="IsChallenge">&nbsp;
&nbsp;&nbsp;
手机号&nbsp;<input type=text size=15 name="UserMobile" Value="<%=Rs("UserMobile")%>">
</td>
</tr>
-->
<tr>
<td width=20% class=td1>VIP用户登记时间</td>
<td width=80% class=td1 colspan=5>
<INPUT TYPE="text" NAME="Vip_StarTime" value="<%=Rs("Vip_StarTime")%>">
</td>
</tr>
<tr>
<td width=20% class=td1>VIP用户到期时间</td>
<td width=80% class=td1 colspan=5>
<INPUT TYPE="text" NAME="Vip_EndTime" value="<%=Rs("Vip_EndTime")%>">
</td>
</tr>
<tr>
<td width=100% class=td1 align=center colspan=6><input name="submit" type=submit class=button value="   更  新   "></td>
</tr>
</FORM>
<%
		end if
		rs.close
		set rs=nothing
	end if
End Sub

Sub SaveUserInfo()
	response.write "<tr><th colspan=8 style=""text-align:center;"">更新用户资料</th></tr>"
	userinfo=checkreal(request.Form("realname")) & "|||" & checkreal(request.Form("character")) & "|||" & checkreal(request.Form("personal")) & "|||" & checkreal(request.Form("country")) & "|||" & checkreal(request.Form("province")) & "|||" & checkreal(request.Form("city")) & "|||" & request.Form("shengxiao") & "|||" & request.Form("blood") & "|||" & request.Form("belief") & "|||" & request.Form("occupation") & "|||" & request.Form("marital") & "|||" & request.Form("education") & "|||" & checkreal(request.Form("college")) & "|||" & checkreal(request.Form("userphone")) & "|||" & checkreal(request.Form("address"))
	dim myuserim,rs,sql
	myuserim=checkreal(request.Form("homepage")) & "|||" & checkreal(request.Form("oicq")) & "|||" & checkreal(request.Form("icq")) & "|||" & checkreal(request.Form("msn")) & "|||" & checkreal(request.Form("aim")) & "|||" & checkreal(request.Form("yahoo")) & "|||" & request.Form("uc")
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=8 class=td1>错误的用户参数。</td></tr>"
		founderr=true
	end if
	'用户签名长度限制 2004-9-13 Dv.Yz
	If Dvbbs.StrLength(Request.Form("sign")) > 250 Then
		Response.Write "<tr><td colspan=8 class=td1>用户签名不能超过 250 个字符。</td></tr>"
		Founderr = True
	End If
	if not founderr then
	Dim iUserClass,iTitlePic
	Set Rs=Dvbbs.Execute("Select * From Dv_UserGroups Where UserGroupID = " & Request.Form("usergroups"))
	If Rs.Eof And Rs.Bof Then
		Response.Write "<tr><td colspan=8 class=td1>所选用户组信息并不存在。</td></tr>"
		Founderr = True
	Else
		iUserClass = Rs("UserTitle")
		iTitlePic = Rs("GroupPic")
	End If
	Dim UpUserName
	Set rs= Dvbbs.iCreateObject("ADODB.Recordset")
	sql="select * from [dv_user] where userid="&request("userid")
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=8 class=td1>没有找到相关用户。</td></tr>"
		founderr=true
	Else
		UpUserName = rs("username")
		Rs("UserPhoto")=Request.form("UserPhoto")
		'rs("username")=request.form("username")
		if request.form("password")<>"" then
			rs("userpassword")=md5(request.form("password"),16)
		end if
		rs("usergroupid")=request.form("usergroups")
		rs("userquesion")=request.form("quesion")
		if request.form("answer")<>"" then rs("useranswer")=md5(request.form("answer"),16)
		rs("userclass")=iUserClass
		rs("useremail")=request.form("useremail")
		Rs("UserSex")=request.form("sex")
		rs("userim")=myuserim
		rs("userface")=request.form("face")
		if isnumeric(request.form("width")) then rs("userwidth")=request.form("width")
		if isnumeric(request.form("height")) then rs("userheight")=request.form("height")
		rs("usertitle")=request.form("usertitle")
		rs("titlepic")=iTitlePic
		if isnumeric(request.form("article")) then rs("UserPost")=request.form("article")
		if isnumeric(request.form("userdel")) then rs("userdel")=request.form("userdel")
		if isnumeric(request.form("userisbest")) then rs("userisbest")=request.form("userisbest")
		if isnumeric(request.form("userpower")) then rs("userpower")=request.form("userpower")
		if isnumeric(request.form("userwealth")) then rs("userwealth")=request.form("userwealth")
		if isnumeric(request.form("usermoney")) then rs("usermoney")=request.form("usermoney")
		if isnumeric(request.form("UserTicket")) then rs("UserTicket")=request.form("UserTicket")
		if isnumeric(request.form("userep")) then rs("userep")=request.form("userep")
		if isnumeric(request.form("usercp")) then rs("usercp")=request.form("usercp")
		if isdate(request.form("birthday")) then rs("userbirthday")=request.form("birthday")
		if isdate(request.form("adddate")) then rs("JoinDate")=request.form("adddate")
		if isdate(request.form("lastlogin")) then rs("lastlogin")=request.form("lastlogin")
		if isdate(request.form("Vip_StarTime")) then rs("Vip_StarTime")=request.form("Vip_StarTime")
		if isdate(request.form("Vip_EndTime")) then rs("Vip_EndTime")=request.form("Vip_EndTime")
		if isnumeric(request.form("lockuser")) then rs("lockuser")=request.form("lockuser")
		rs("usersign")=request.form("sign")
		rs("userinfo")=userinfo
		'If request.form("IsChallenge")="0" Or Request.Form("UserMobile")="" Then
			Rs("IsChallenge")=0
			Rs("UserMobile")=""
		'Else
		'	Rs("IsChallenge")=1
		'	Rs("UserMobile")=Request.Form("UserMobile")
		'End If
		rs.update
	end if
	rs.close
	set rs=nothing
	end if
	if not founderr then
		'-----------------------------------------------------------------
		'系统整合
		'-----------------------------------------------------------------
		Dim DvApi_Obj,DvApi_SaveCookie,SysKey
		If DvApi_Enable Then
			Set DvApi_Obj = New DvApi
				DvApi_Obj.NodeValue "syskey",SysKey,0,False
				DvApi_Obj.NodeValue "action","update",0,False
				DvApi_Obj.NodeValue "username",UpUserName,1,False
				Md5OLD = 1
				SysKey = Md5(DvApi_Obj.XmlNode("username")&DvApi_SysKey,16)
				Md5OLD = 0
				DvApi_Obj.NodeValue "syskey",SysKey,0,False
				DvApi_Obj.NodeValue "password",Request.form("password"),1,False
				DvApi_Obj.NodeValue "answer",Request.Form("useranswer"),1,False
				DvApi_Obj.NodeValue "question",Request.Form("quesion"),1,False
				DvApi_Obj.NodeValue "email",Request.Form("useremail"),1,False
				DvApi_Obj.SendHttpData
				If DvApi_Obj.Status = "1" Then
					response.write "<tr><td colspan=8 class=td1>"&DvApi_Obj.Message&"</td></tr>"
				End If
			Set DvApi_Obj = Nothing
		End If
		'-----------------------------------------------------------------
	End If
	if founderr then
		response.write "<tr><td colspan=8 class=td1>更新失败。</td></tr>"
	else

		response.write "<tr><td colspan=8 class=td1>更新用户数据成功。</td></tr>"
	end if
End Sub

Sub UserPermission()
	Response.Write "<tr><th colspan=8 style=""text-align:center;"">编辑" & Request("Username") & "论坛权限（红色表示该用户在该版面有自定义权限）</th></tr>"
	If Not Isnumeric(Request("Userid")) Then
		Response.Write "<tr><td colspan=8 class=td1>错误的用户参数。</td></tr>"
		Founderr = True
	End If
	If Not Founderr Then
		Response.Write "<tr><td colspan=8 class=td1 height=25>①您可以设置该用户在不同论坛内的权限，红色表示为该用户组使用的是用户自定义属性<BR>②该权限不能继承，比如您设置了一个包含下级论坛的版面，那么只对您设置的版面生效而不对其下属论坛生效<BR>③如果您想设置生效，必须在设置页面<B>选择自定义设置</B>，选择了自定义设置后，这里设置的权限将<B>优先</B>于用户组设置和论坛权限设置，比如用户组默认或论坛权限设置该用户组不能管理帖子，而这里设置了该用户可管理帖子，那么该用户在这个版面就可以管理帖子</td></tr>"
		Set Trs = Dvbbs.Execute("SELECT Uc_UserId FROM Dv_UserAccess WHERE Uc_Boardid = 0 AND Uc_Userid = " & Request("Userid"))
		If Trs.Eof And Trs.Bof Then
		Response.Write "<tr><td colspan=8 class=td1 height=25><a href=?action=UserBoardPermission&boardid=0&userid=" & Request("Userid") & ">编辑该用户在全局的权限</a>（前台短信、前台用户信息、帖子和权限管理、进入后台权限等）</td></tr>"
		Else
		Response.Write "<tr><td colspan=8 class=td1 height=25><a href=?action=UserBoardPermission&boardid=0&userid=" & Request("Userid") & "><font color=red>编辑该用户在全局的权限</font></a>（前台短信、前台用户信息、帖子和权限管理、进入后台权限等）</td></tr>"
		End If
'----------------------boardinfo--------------------
		Response.Write "<tr><td colspan=8 class=td1><B>点击论坛名称进入编辑状态</B><BR>"
		Rem 改用数组代替循环查询 2004-5-6 Dvbbs.YangZheng
		Dim Bn,Sql,Rs,i
		Sql = "SELECT Depth, Child, Boardid, Parentid, Boardtype FROM Dv_Board ORDER BY Rootid, Orders"
		Set Rs = Dvbbs.Execute(Sql)
		If Not (Rs.Eof And Rs.Bof) Then
			Sql = Rs.GetRows(-1)
			Rs.Close:Set Rs = Nothing
			For Bn = 0 To Ubound(Sql,2)
				If Sql(0,Bn) > 0 Then
					For i = 1 To Sql(0,Bn)
						Response.Write "&nbsp;"
					Next
				End If
				If Sql(1,Bn) > 0 Then
					Response.Write "<img src=""../skins/default/plus.gif"">"
				Else
					Response.Write "<img src=""../skins/default/nofollow.gif"">"
				End If
%>
<a href="?action=UserBoardPermission&boardid=<%=Sql(2,Bn)%>&userid=<%=Request("Userid")%>">
<%
				Set Trs = Dvbbs.Execute("SELECT Uc_UserId FROM Dv_UserAccess WHERE Uc_Boardid = " & Sql(2,Bn) & " AND Uc_Userid = " & Request("Userid"))
				If Not (Trs.Eof And Trs.Bof) Then
					Response.Write "<font color=red>[自定义]"
				End If
				If Sql(3,Bn) = 0 Then Response.Write "<b>"
				Response.Write Sql(4,Bn)
				If Sql(3,Bn) = 0 Then Response.Write "</b>"
				If Sql(1,Bn) > 0 Then Response.Write "(" & Sql(1,Bn) & ")"
				Response.Write "</font></a><BR>"
			Next
		End If
		Response.Write "</td></tr>"
'-------------------end-------------------
	End If
End Sub

Sub UserBoardPermission()
	Dim rs
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=8 class=td1>错误的用户参数。</td></tr>"
		founderr=true
	end if
	if not isnumeric(request("boardid")) then
		response.write "<tr><td colspan=8 class=td1>错误的版面参数。</td></tr>"
		founderr=true
	end if
	if not founderr then
	set rs=Dvbbs.Execute("select u.UserGroupID,ug.title,u.username from [dv_user] u inner join dv_UserGroups UG on u.userGroupID=ug.userGroupID where u.userid="&request("userid"))
	Dvbbs.UserGroupID=rs(0)
	usertitle=rs(1)
	Dvbbs.membername=rs(2)
	dim boardtype
	set rs=Dvbbs.Execute("select boardtype from dv_board where boardid="&request("boardid"))
	if rs.eof and rs.bof then
	boardtype="论坛其他页面"
	else
	boardtype=rs(0)
	end if
	response.write "<tr><th colspan=8 style=""text-align:center;"">编辑 "&Dvbbs.membername&" 在 "&boardtype&" 权限</th></tr>"
	response.write "<tr><td colspan=8 height=25 class=td1>注意：该用户属于 <B>"&usertitle&"</B> 用户组中，如果您设置了他的自定义权限，则该用户权限将以自定义权限为主</td></tr>"
%>
<tr><td colspan=8 class=td1>
<%
Dim reGroupSetting
Dim FoundGroup,FoundUserPermission,FoundGroupPermission
FoundGroup=false
FoundUserPermission=false
FoundGroupPermission=false

set rs=Dvbbs.Execute("select * from dv_UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
if not (rs.eof and rs.bof) then
	reGroupSetting=rs("uc_Setting")
	FoundGroup=true
	FoundUserPermission=true
end if

if not foundgroup then
set rs=Dvbbs.Execute("select * from dv_BoardPermission where boardid="&request("boardid")&" and groupid="&DVbbs.UserGroupID)
if not(rs.eof and rs.bof) then
	reGroupSetting=rs("PSetting")
	FoundGroup=true
	FoundGroupPermission=true
end if
end if

if not foundgroup then
set rs=Dvbbs.Execute("select * from dv_usergroups where usergroupid="&DVbbs.UserGroupID)
if rs.eof and rs.bof then
	response.write "未找到该用户组！"
	response.end
else
	FoundGroup=true
	FoundGroupPermission=true
	reGroupSetting=rs("GroupSetting")
end if
end if
%>
<table width="100%" border="0" cellspacing="1" cellpadding="0" align=center>
<FORM METHOD=POST ACTION="?action=saveuserpermission">
<input type=hidden name="userid" value="<%=request("userid")%>">
<input type=hidden name="BoardID" value="<%=request("boardid")%>">
<input type=hidden name="username" value="<%=Dvbbs.membername%>">
<%If Dvbbs.BoardID <> 0 Then%>
<tr> 
<td width="100%" class=td1 colspan=2 height=25>
<font color=blue>保存目标</font>：<input type=radio class=radio name="savetype" value=0 checked>该版面&nbsp;<input type=radio class=radio name="savetype" value=1>所有版面&nbsp;<input type=radio class=radio name="savetype" value=2>相同分类下所有版面（不包括分类）&nbsp;<input type=radio class=radio name="savetype" value=3>相同分类下所有版面（包括分类）&nbsp;<input type=radio class=radio name="savetype" value=4>同分类同级别版面
</td>
</tr>
<tr> 
<td width="100%" class=td1 colspan=2 height=25>
<font color=blue>
这里指的分类仅指一级分类，而不是该版面的上级版面</font>，比如您目前设置的是一个五级版面，选择了相同分类下所有版面都更新，那么这里将更新包括该分类的一级、二级、三级、四级所有版面，如果您担心更新范围太大，可以选择更新同分类同级别版面。
</td>
</tr>
<%Else%>
<input type=hidden name="savetype" value=0>
<%End If%>
<tr> 
<td height="23" colspan="2" class=td1><input type=radio class=radio name="isdefault" value="1" <%if FoundGroupPermission then%>checked<%end if%>><B>使用用户组默认值</B> (注意: 这将删除任何之前所做的自定义设置)</td>
</tr>
<tr> 
<td height="23" colspan="2"  class=td1><input type=radio class=radio name="isdefault" value="0" <%if FoundUserPermission then%>checked<%end if%>><B>使用自定义设置</B> &nbsp;(<font color=blue>选择自定义才能使以下设置生效</font>)</td>
</tr>
<%
GroupPermission(reGroupSetting)
%>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
</td></tr>
<%
	end if
End Sub

Sub SaveUserPermission()
	Dim i
	response.write "<tr><th colspan=8 style=""text-align:center;"">编辑用户 "&request("username")&" 权限</th></tr>"
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=8 class=td1>错误的用户参数。</td></tr>"
		founderr=true
	end if
	if not isnumeric(request("boardid")) then
		response.write "<tr><td colspan=8 class=td1>错误的版面参数。</td></tr>"
		founderr=true
	end if
	if not founderr then
	dim myGroupSetting,rs
	Dim IsGroupSetting,MyIsGroupSetting,FoundSetting
	myGroupSetting=GetGroupPermission
	select case request("savetype")
	'当前版面
	case "0"
		if request("isdefault")=1 then
			Dvbbs.Execute("delete from dv_UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
			Set Rs=Dvbbs.Execute("Select Count(*) from dv_UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
			FoundSetting=Rs(0)
			If IsNull(FoundSetting) Or FoundSetting="" Then FoundSetting=0
			If Dvbbs.BoardID > 0 Then
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&request("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = ""
			Else
				IsGroupSetting = "," & Rs(0) & ","
				If FoundSetting=0 Then IsGroupSetting = Replace(IsGroupSetting,",0_"&request("userid"),"")
				IsGroupSetting=replace(IsGroupSetting,",,",",")
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&request("boardid"))
			End If
		else
			set rs=Dvbbs.Execute("select * from dv_UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
			if rs.eof and rs.bof then
				Dvbbs.Execute("insert into dv_UserAccess (uc_userid,uc_boardid,uc_setting) values ("&request("userid")&","&request("boardid")&",'"&myGroupSetting&"')")
			else
				Dvbbs.Execute("update dv_UserAccess set uc_setting='"&myGroupSetting&"' where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
			end if
			If Dvbbs.BoardID > 0 Then
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&request("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = 0
			Else
				IsGroupSetting = "," & Rs(0) & ","
				IsGroupSetting = Replace(IsGroupSetting,",0_"&request("userid"),"")
				IsGroupSetting=replace(IsGroupSetting,",,",",")
				IsGroupSetting = IsGroupSetting & "0_"&request("userid")&","
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&request("boardid"))
			Set Rs=Nothing
			End If
		end if
		If Dvbbs.BoardID > 0 Then Dvbbs.ReloadBoardCache request("boardid")
	'所有版面
	case "1"
		set trs=Dvbbs.Execute("select * from dv_board")
		do while not trs.eof
		if request("isdefault")=1 then
			Dvbbs.Execute("delete from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			Set Rs=Dvbbs.Execute("Select Count(*) from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			FoundSetting=Rs(0)
			If IsNull(FoundSetting) Or FoundSetting="" Then FoundSetting=0
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&trs("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = ""
			Else
				IsGroupSetting = "," & Rs(0) & ","
				If FoundSetting=0 Then IsGroupSetting = Replace(IsGroupSetting,",0_"&request("userid"),"")
				IsGroupSetting=replace(IsGroupSetting,",,",",")
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			FoundSetting=""
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&trs("boardid"))
		Else
			set rs=Dvbbs.Execute("select * from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			if rs.eof and rs.bof then
				Dvbbs.Execute("insert into dv_UserAccess (uc_userid,uc_boardid,uc_setting) values ("&request("userid")&","&trs("boardid")&",'"&myGroupSetting&"')")
			else
				Dvbbs.Execute("update dv_UserAccess set uc_setting='"&myGroupSetting&"' where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			end if
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&trs("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = 0
			Else
				IsGroupSetting = "," & Rs(0) & ","
				IsGroupSetting = Replace(IsGroupSetting,",0_"&request("userid"),"")
				IsGroupSetting=replace(IsGroupSetting,",,",",")
				IsGroupSetting = IsGroupSetting & "0_"&request("userid")&","
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&trs("boardid"))
		end if
		Dvbbs.ReloadBoardCache trs("boardid")
		trs.movenext
		loop
		trs.close
		set trs=nothing
		Set Rs=Nothing
	'相同分类下所有版面（不包括分类）
	case "2"
		set trs=Dvbbs.Execute("select rootid from dv_board where boardid="&request("boardid"))
		myrootid=trs(0)
		set trs=Dvbbs.Execute("select * from dv_board where (Not ParentID=0) and rootid="&myrootid)
		do while not trs.eof
		if request("isdefault")=1 then
			Dvbbs.Execute("delete from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			Set Rs=Dvbbs.Execute("Select Count(*) from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			FoundSetting=Rs(0)
			If IsNull(FoundSetting) Or FoundSetting="" Then FoundSetting=0
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&trs("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = ""
			Else
				IsGroupSetting = "," & Rs(0) & ","
				If FoundSetting=0 Then IsGroupSetting = Replace(IsGroupSetting,",0_"&request("userid"),"")
				IsGroupSetting=replace(IsGroupSetting,",,",",")
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			FoundSetting=""
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&trs("boardid"))
		else
			set rs=Dvbbs.Execute("select * from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			if rs.eof and rs.bof then
				Dvbbs.Execute("insert into dv_UserAccess (uc_userid,uc_boardid,uc_setting) values ("&request("userid")&","&trs("boardid")&",'"&myGroupSetting&"')")
			else
				Dvbbs.Execute("update dv_UserAccess set uc_setting='"&myGroupSetting&"' where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			end if
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&trs("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = 0
			Else
				IsGroupSetting = "," & Rs(0) & ","
				IsGroupSetting = Replace(IsGroupSetting,",0_"&request("userid"),"")
					IsGroupSetting=replace(IsGroupSetting,",,",",")
				IsGroupSetting = IsGroupSetting & "0_"&request("userid")&","
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&trs("boardid"))
		end if
		Dvbbs.ReloadBoardCache trs("boardid")
		trs.movenext
		loop
		trs.close
		set trs=nothing
		Set Rs=Nothing
	'相同分类下所有版面（包括分类）
	case "3"
		set trs=Dvbbs.Execute("select rootid from dv_board where boardid="&request("boardid"))
		myrootid=trs(0)
		set trs=Dvbbs.Execute("select * from dv_board where rootid="&myrootid)
		do while not trs.eof
		if request("isdefault")=1 then
			Dvbbs.Execute("delete from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			Set Rs=Dvbbs.Execute("Select Count(*) from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			FoundSetting=Rs(0)
			If IsNull(FoundSetting) Or FoundSetting="" Then FoundSetting=0
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&trs("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = ""
			Else
				IsGroupSetting = "," & Rs(0) & ","
				If FoundSetting=0 Then IsGroupSetting = Replace(IsGroupSetting,",0_"&request("userid"),"")
				IsGroupSetting=replace(IsGroupSetting,",,",",")
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			FoundSetting=""
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&trs("boardid"))
		else
			set rs=Dvbbs.Execute("select * from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			if rs.eof and rs.bof then
				Dvbbs.Execute("insert into dv_UserAccess (uc_userid,uc_boardid,uc_setting) values ("&request("userid")&","&trs("boardid")&",'"&myGroupSetting&"')")
			else
				Dvbbs.Execute("update dv_UserAccess set uc_setting='"&myGroupSetting&"' where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			end if
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&trs("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = 0
			Else
				IsGroupSetting = "," & Rs(0) & ","
				IsGroupSetting = Replace(IsGroupSetting,",0_"&request("userid"),"")
				IsGroupSetting=replace(IsGroupSetting,",,",",")
				IsGroupSetting = IsGroupSetting & "0_"&request("userid")&","
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&trs("boardid"))
		end if
		Dvbbs.ReloadBoardCache trs("boardid")
		trs.movenext
		loop
		trs.close
		set trs=nothing
		Set Rs=Nothing
	'同分类同级别版面
	case "4"
		dim myparentid,myparentstr
		set trs=Dvbbs.Execute("select rootid,ParentStr,ParentID from dv_board where boardid="&request("boardid"))
		myrootid=trs(0)
		myparentstr=trs(1)
		myparentid=trs(2)
		set trs=Dvbbs.Execute("select * from dv_board where rootid="&myrootid&" and ParentID="&myparentid&" and ParentStr='"&myparentstr&"'")
		do while not trs.eof
		if request("isdefault")=1 then
			Dvbbs.Execute("delete from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			Set Rs=Dvbbs.Execute("Select Count(*) from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			FoundSetting=Rs(0)
			If IsNull(FoundSetting) Or FoundSetting="" Then FoundSetting=0
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&trs("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = ""
			Else
				IsGroupSetting = "," & Rs(0) & ","
				If FoundSetting=0 Then IsGroupSetting = Replace(IsGroupSetting,",0_"&request("userid"),"")
				IsGroupSetting=replace(IsGroupSetting,",,",",")
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			FoundSetting=""
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&trs("boardid"))
		else
			set rs=Dvbbs.Execute("select * from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			if rs.eof and rs.bof then
				Dvbbs.Execute("insert into dv_UserAccess (uc_userid,uc_boardid,uc_setting) values ("&request("userid")&","&trs("boardid")&",'"&myGroupSetting&"')")
			else
				Dvbbs.Execute("update dv_UserAccess set uc_setting='"&myGroupSetting&"' where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			end if
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&trs("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = 0
			Else
				IsGroupSetting = "," & Rs(0) & ","
				IsGroupSetting = Replace(IsGroupSetting,",0_"&request("userid"),"")
				IsGroupSetting=replace(IsGroupSetting,",,",",")
				IsGroupSetting = IsGroupSetting & "0_"&request("userid")&","
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&trs("boardid"))
		end if
		Dvbbs.ReloadBoardCache trs("boardid")
		trs.movenext
		loop
		trs.close
		set trs=nothing
		Set Rs=Nothing
	end select
	if founderr then
		response.write "<tr><td colspan=8 class=td1>更新失败。</td></tr>"
	else
		response.write "<tr><td colspan=8 class=td1><li>设置用户权限成功。"
		If Request.Form("GroupSetting(70)") = "1" Then Response.Write "<li>您设置了该用户可进入论坛后台的权限，请到管理员管理中 <a href=""admin.asp?action=add"">添加</a> 该用户的后台账号和设置该帐户后台权限。"
		Response.Write "</td></tr>"
	end if
	End if
End Sub

Sub UniteUser()
	if request("auser")<>"" and request("buser")<>"" then
		dim auserid,buserid,rs,i
		dim c1,c2,c3,c4,c5,c6,c7,c8,c9
		set rs=dvbbs.execute("select userid,userpost,usertopic,userviews,userwealth,userep,usercp,userpower,userisbest,userdel,usergroupid from dv_user where username='"&replace(request("auser"),"'","''")&"'")
		if rs.eof and rs.bof then
			errmsg = errmsg + "<tr><td colspan=8 class=td1>没有找到被合并用户</td></tr>"
			founderr=true
		else
			auserid=rs(0)
			c1=rs(1)
			c2=rs(2)
			c3=rs(3)
			c4=rs(4)
			c5=rs(5)
			c6=rs(6)
			c7=rs(7)
			c8=rs(8)
			c9=rs(9)
			if rs(10) < 4 then
				errmsg = errmsg + "<tr><td colspan=8 class=td1>只允许对注册用户组进行合并用户操作</td></tr>"
				founderr=true
			end if
		end if
		set rs=dvbbs.execute("select userid from dv_user where username='"&replace(request("buser"),"'","''")&"'")
		if rs.eof and rs.bof then
			errmsg = errmsg + "<tr><td colspan=8 class=td1>没有找到合并的目标用户</td></tr>"
			founderr=true
		else
			buserid=rs(0)
		end if
		if auserid=buserid then
			errmsg = errmsg + "<tr><td colspan=8 class=td1>相同用户不能进行合并</td></tr>"
			founderr=true
		end if
		if founderr then
			Response.Write errmsg
		else
			'合并用户的资料
			dvbbs.execute("update dv_user set userpost=userpost+"&c1&",usertopic=usertopic+"&c2&",userviews=userviews+"&c3&",userwealth=userwealth+"&c4&",userep=userep+"&c5&",usercp=usercp+"&c6&",userpower=userpower+"&c7&",userisbest=userisbest+"&c8&",userdel=userdel+"&c9&" where userid="&buserid)
			'更新帖子数据
			for i=0 to ubound(allposttable)
				dvbbs.execute("update "&allposttable(i)&" set postuserid="&buserid&",username='"&replace(request("buser"),"'","''")&"' where postuserid="&auserid)
			next
			dvbbs.execute("update dv_topic set postuserid="&buserid&",postusername='"&replace(request("buser"),"'","''")&"' where postuserid="&auserid)
			'更新短信数据
			Dvbbs.Execute("update dv_message set incept='"&replace(request("buser"),"'","''")&"' where incept='"&replace(request("auser"),"'","''")&"'")
			Dvbbs.Execute("update dv_message set sender='"&replace(request("buser"),"'","''")&"' where sender='"&replace(request("auser"),"'","''")&"'")
			Dvbbs.Execute("update dv_friend set F_username='"&replace(request("buser"),"'","''")&"' where F_username='"&replace(request("auser"),"'","''")&"'") 
			Dvbbs.Execute("update dv_bookmark set username='"&replace(request("buser"),"'","''")&"' where username='"&replace(request("auser"),"'","''")&"'") 

			Dvbbs.Execute("update dv_besttopic set PostUserID="&buserid&",postusername='"&replace(request("buser"),"'","''")&"' where PostUserID="&auserid)
			'更新用户上传表
			Dvbbs.Execute("update dv_upfile set F_UserID="&buserid&",F_Username='"&replace(request("buser"),"'","''")&"' where F_UserID="&auserid)
			response.write "<tr><td colspan=8 class=td1>合并用户数据成功。</td></tr>"
		end if
	else
%>
<form action="?action=uniteuser" method=post>
<tr>
<th colspan=7 style="text-align:center;">合并用户</th>
</tr>
<tr>
<td width=20% class=td1>注意事项</td>
<td width=80% class=td1 colspan=5>被合并用户在论坛中的所有帖子（包括精华）、短信、上传、收藏等资料将合并到所指定的用户中</td>
</tr>
<tr>
<td width=20% class=td1>选项</td>
<td width=80% class=td1 colspan=5>把用户 <input size=25 name="auser" type=text> 资料合并到 <input size=25 name="buser" type=text> 用户 <input type=submit class=button name=submit value="提交"></td>
</tr>
</form>
<%
	end if
End Sub

Sub Fixuser()
	Dim Userid,i
	Userid = Request("Userid")
	If Not IsNumeric(Userid) Then
	Errmsg = ErrMsg + "<BR><li>参数错误!"
		Dvbbs_Error()
		Exit Sub
	End If
	Userid = CLng(Userid)
	Dim Rs, Username, UserArticle, UserIsBest
	UserArticle = 0
	Set Rs = Dvbbs.Execute("SELECT Username FROM [Dv_User] WHERE Userid = " & Userid & "")
	If Rs.Eof Or Rs.Bof Then
		Errmsg = ErrMsg + "<BR><li>找不到该用户，误删用户需要重新用原来的名字注册才可以修复数据!"
		Dvbbs_Error()
		Exit Sub
	Else
		Username = Rs(0)
		Rs.Close:Set Rs = Nothing
		'修复主题表
		Dvbbs.Execute ("Update Dv_Topic Set PostUserID = " & Userid & " WHERE PostUserName = '" & Username & "'")
		'修复所有数据表
		For i = 0 To Ubound(AllPostTable)
			Dvbbs.Execute ("Update " & AllPostTable(i) & " Set Postuserid = " & Userid & " WHERE UserName = '" & Username & "'")
			'计算用户发贴
			Set Rs = Dvbbs.Execute("SELECT COUNT(*) FROM " & AllPostTable(i) & " WHERE Postuserid = " & Userid & "")
			UserArticle = UserArticle + Rs(0)
			Rs.Close:Set Rs = Nothing
		Next
		'修复精华
		Dvbbs.Execute ("UPDATE Dv_BestTopic Set PostUserID = " & Userid & " WHERE PostUserName = '" & Username & "'")
		Set Rs = Dvbbs.Execute("SELECT COUNT(*) FROM Dv_BestTopic WHERE Postuserid = " & Userid &"")
		UserIsBest = Rs(0)
		Rs.Close:Set Rs = Nothing
		'修复上传文件列表
		Dvbbs.Execute ("UPDATE DV_Upfile SET F_UserID = " & Userid & " WHERE F_Username = '" & Username & "'")
		'更新发贴数
		Dvbbs.Execute ("UPDATE [Dv_User] SET UserPost = " & UserArticle & ", UserIsBest = " & UserIsBest & " WHERE Userid = " & Userid & "")
	End If
	Set Rs = Nothing
	Dv_Suc("用户<b>" & Username & "</b>数据修复成功！")
End Sub

Function CheckReal(v)
	Dim w
	If Not IsNull(v) Then
		w=Replace(v,"|||","§§§")
		CheckReal=w
	End If
End Function

Function CheckNumeric(Byval CHECK_ID)
	If CHECK_ID<>"" and IsNumeric(CHECK_ID) Then _
		CHECK_ID = Int(CHECK_ID) _
	Else _
		CHECK_ID = 0
	CheckNumeric = CHECK_ID
End Function
%>