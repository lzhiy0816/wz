<!-- #include file="setup.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
if adminpassword<>session("pass") then response.redirect "admin.asp?menu=login"

log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")


%>
<META http-equiv=Content-Type content=text/html;charset=gb2312>
<link href=images/skins/<%=Request.Cookies("skins")%>/bbs.css rel=stylesheet>
<br><center>
<%

select case Request("menu")

case "user"
user

case "user2"
user2

case "showall"
showall

case "showallok"
showallok

case "userdeltopic"
userdeltopic

case "userdel"
userdel

case "userok"
userok

case "fix"

conn.execute("update [user] set userpass='"&md5("123")&"'  where username='"&request("username")&"'")
error2("已经将 "&request("username")&" 的密码还原成 123 ")

case "sendmoney"


sql="select username from [user] where membercode="&Request("membercode")&""
rs.Open sql,Conn
do while not rs.eof
Count=Count+1

content=replace(server.htmlencode(Request("content")), "'", "&#39;")

conn.Execute("insert into message (author,incept,content) values ('"&Request.Cookies("username")&"','"&rs("username")&"','【系统广播】："&content&"')")

conn.execute("update [user] set newmessage=newmessage+1,savemoney=savemoney+"&Request("money")&" where username='"&rs("username")&"'")
rs.movenext
loop
rs.close
error2("发放完成！")

end select



sub showall





%>
用户资料：<b><font color=red><%=conn.execute("Select count(id)from [user]")(0)%></font></b> 条




<div align="center">


<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 colspan=2 align="center">用户管理</td>
  </tr>


<tr class=a4>
<td><form method="post" action="?menu=user2">
请输入会员的名称: <input size="13" name="username">
<input type="submit" value="确定">

</td></form>
<td>
<form method="post" action="?menu=showallok">
快速搜索 <select onchange="javascript:submit()" size="1" name="userSearch">
<option value="">请选择查询条件</option>


<option value="landtime">24小时内登录的用户</option>
<option value="regtime">24小时内注册的用户</option>
<option value=" and membercode=2">所有特邀嘉宾</option>
<option value=" and membercode=4">所有管理员</option>
<option value=" and membercode=5">所有社区区长</option>
<option value=" and experience<2">经验值少于2的用户</option>
<option value=" and degree<2">登录次数少于2的用户</option>
</select>



</td>
</tr>

  <tr height=25>
    <td class=a1 align=center colspan=2>高级查询</td>
  </tr>




  <tr height=25 class=a3>
    <td>最多显示记录数</td>
    <td><input size="45" value="50" name="searchMax"></td>
  </tr>


  <tr height=25 class=a3>
    <td>用户名包含</td>
    <td><input size="45" name="username"></td>
  </tr>


  <tr height=25 class=a3>
    <td>Email包含</td>
    <td><input size="45" name="usermail"></td>
  </tr>


  <tr height=25 class=a3>
    <td>主页包含</td>
    <td><input size="45" name="userhome"></td>
  </tr>


  <tr height=25 class=a3>
    <td>QQ包含</td>
    <td><input size="45" name="userqq"></td>
  </tr>


  <tr height=25 class=a3>
    <td>ICQ包含</td>
    <td><input size="45" name="icq"></td>
  </tr>


  <tr height=25 class=a3>
    <td>签名包含</td>
    <td><input size="45" name="sign"></td>
  </tr>


  <tr height=25 class=a3>
    <td>注册日期</td>
    <td><input size="45" name="regtime"></td>
  </tr>


  <tr height=25 class=a3>
    <td colspan="2" align="center">
	<input type="submit" value="   搜  索   " name="submit0"></td>
  </tr>

</form>
  </table>



</div>



<br>
<%
end sub







sub showallok


%>
<script>
function CheckAll(form){for (var i=0;i<form.elements.length;i++){var e = form.elements[i];if (e.name != 'chkall')e.checked = form.chkall.checked;}}
</script>


<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
<TR align=middle>
<TD class=a1>用户名</TD>
<TD class=a1 height="25">Email</TD>

<TD class=a1>短讯息</TD>

<TD class=a1 height="25">权限</TD>

<TD class=a1>注册时间</TD>
<TD class=a1>最后登录时间</TD>
<TD class=a1>删除</TD>
</TR>
<form method="post" action="?menu=userdel">
<%

item=Request("userSearch")

if item="landtime" then
item=" and landtime >"&SqlNowString&"-1 "
end if
if item="regtime" then
item=" and regtime >"&SqlNowString&"-1 "
end if


if Request("username")<>"" then
item=""&item&" and username like '%"&Request("username")&"%'"
end if

if Request("usermail")<>"" then
item=""&item&" and usermail like '%"&Request("usermail")&"%'"
end if

if Request("userhome")<>"" then
item=""&item&" and userhome like '%"&Request("userhome")&"%'"
end if

if Request("userqq")<>"" then
item=""&item&" and userqq like '%"&Request("userqq")&"%'"
end if

if Request("icq")<>"" then
item=""&item&" and icq like '%"&Request("icq")&"%'"
end if

if Request("sign")<>"" then
item=""&item&" and sign like '%"&Request("sign")&"%'"
end if

if Request("regtime")<>"" then
item=""&item&" and regtime like '%"&Request("regtime")&"%'"
end if





if item="" or Request("searchMax")="" then
error2("请输入您要搜索的条件")
end if 


item="where"&item&""
item=replace(item,"where and","where")


sql="select top "&Request("searchMax")&" * from [user] "&item&""
rs.Open sql,Conn
Do While Not RS.EOF 
%>
<TBODY><TR align=middle>
<TD class=a4><a target="_blank" href=Profile.asp?username=<%=rs("username")%>><%=rs("username")%></a>　</TD>
<TD class=a3><%=rs("usermail")%>　</TD>
<TD class=a4><a style=cursor:hand onclick="javascript:open('friend.asp?menu=post&incept=<%=rs("username")%>','','width=320,height=170')"><img border="0" src="images/message1.gif"></a></TD>
<TD class=a3>
<a href="admin_user.asp?menu=user2&username=<%=rs("username")%>">编辑</a></TD>
<TD class=a4><%=rs("regtime")%>　</TD>
<TD class=a3><%=rs("landtime")%>　</TD>
<TD class=a3><input type="checkbox" value="<%=rs("username")%>" name="username"></TD></TR>



<%
RS.MoveNext
loop
RS.Close
%>

	<TR align=middle class=a3>
<TD colspan="7" align="right">&nbsp;<input onclick="{if(confirm('您确定要删除您所选用户的全部资料?')){return true;}return false;}"  type="submit" value="   确 定   "> <input type=checkbox name=chkall onclick=CheckAll(this.form) value="ON">&nbsp;</TD></form>
	</TR>
</TABLE><center><br>
<a href="javascript:history.back()">返 回</a> 
<%
end sub







sub user


%>
用户资料：<b><font color=red><%=conn.execute("Select count(id)from [user]")(0)%></font></b> 条
<br>
<form method="post" action="?menu=sendmoney">
<table cellspacing="1" cellpadding="2" width="64%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>发放工资</td>
  </tr>
  <tr height=25>
    <td class=a3 align=middle colspan=2>给所有
<select size="1" name="membercode">
<option value="1">普通用户</option>
<option value="2">特邀嘉宾</option>
<option value="4" selected>管 理 员</option>
<option value="5">社区区长</option>
</select>&nbsp;发放本月工资 <input size="8" name="money" value="1000">
金币
<input type="submit" value="确定"> </td>
  </tr>
  <tr height=25 class=a3>
    <td align=middle width="13%">短<br>
		讯<br>
		内<br>
		容</td>
    <td align=middle width="85%"><textarea name="content" rows="5" cols="55">本月工资已经转账到您的银行存款中，请注意查收！</textarea></td>
  </tr>
   </table>

</form>
<%
end sub

sub user2

username=Request("username")
sql="select * from [user] where username='"&HTMLEncode(username)&"'"
rs.Open sql,Conn
if rs.eof then
error2(""&username&" 的用户资料不存在")
end if

%>
<form method="post" name=form action="?menu=userok">
<input type=hidden name=username value="<%=rs("username")%>">
<div align="center">
<center>
<table cellSpacing="1" cellPadding="3" border="0" width="389" class=a2>
<tr class=a1 id=TableTitleLink>
<td height="25" width="300" colspan="2" align="center">
社区信息息</td>
<td height="25" align="center" width="300" colspan="2"><font color="000000"><a target="_blank" href="Profile.asp?username=<%=rs("username")%>">
查看详细资料</a></font></td>
</tr>
<tr class=a3>
<td width="300" colspan="2" height="25">&nbsp;用户名称
&nbsp;<%=rs("username")%></td>
<td width="300" colspan="2" height="25">&nbsp;用户密码
 
&nbsp;<a onclick="{if(confirm('此操作将会把该用户的密码改成：123')){return true;}return false;}" href="?menu=fix&username=<%=rs("username")%>">还原密码</a>
</td>
</tr>
<tr class=a4>
<td width="300" colspan="2">&nbsp;用户头衔 <input size="10" name="honor" value="<%=rs("honor")%>"></td>
<td width="300" colspan="2">&nbsp;用户权限
<select size="1" name="membercode">
<option value=1 <%if rs("membercode")=1 then%>selected<%end if%>>普通会员</option>
<option value=2 <%if rs("membercode")=2 then%>selected<%end if%>>特别嘉宾</option>
<option value=4 <%if rs("membercode")=4 then%>selected<%end if%>>管 理 员</option>
<option value=5 <%if rs("membercode")=5 then%>selected<%end if%>>社区区长</option>
</select> </td>
</tr>
<tr class=a3>
<td width="300" colspan="2">&nbsp;发表文章 <input size="10" name="posttopic" value="<%=rs("posttopic")%>"></td>
<td width="300" colspan="2">&nbsp;回复文章 <input size="10" name="postrevert" value="<%=rs("postrevert")%>"></td>
</tr>
<tr class=a4>
<td width="300" colspan="2">&nbsp;社区金币 <input size="10" name="money" value="<%=rs("money")%>"></td>
<td width="300" colspan="2">&nbsp;经 验 值 <input size="10" name="experience" value="<%=rs("experience")%>"></td>
</tr>
<tr class=a3>
<td width="300" colspan="2">&nbsp;社区存款 <input size="10" name="savemoney" value="<%=rs("savemoney")%>"></td>
<td width="300" colspan="2">&nbsp;体 力 值 <input size="10" name="userlife" value="<%=rs("userlife")%>"></td>
</tr>
<tr class=a4>
<td width="600" colspan="4">&nbsp;用户签名 
<textarea name="sign" rows="3" cols="40"><%=rs("sign")%></textarea></td>
</tr>
<tr class=a1 id=TableTitleLink>
<td width="201" align="center" height="25">
<a onclick="{if(confirm('您确定要删除该用户所有发表过的帖子?')){return true;}return false;}" href="?menu=userdeltopic&username=<%=rs("username")%>">
删除该用户的所有帖子</a>


<td width="201" colspan="2" align="center" height="25">
<input type="submit" value=" 更 新 " name="Submit"></td>
<td width="201" align="center" height="25">

<a onclick="{if(confirm('您确定要删除该用户的所有资料?')){return true;}return false;}" href="?menu=userdel&username=<%=rs("username")%>">
删除该用户的所有资料</a></td>

</td>
</tr>
</table></center>
</div>
</form><A href="javascript:history.back()">返 回</A>
<Script>
document.form.sign.value = unybbcode(document.form.sign.value);
document.form.sign.focus();
function unybbcode(temp) {
temp = temp.replace(/<br>/ig,"\n");
return (temp);
}
</Script>
<%
end sub

sub userdeltopic
conn.execute("delete from [forum] where username='"&Request("username")&"'")
conn.execute("delete from [reforum] where username='"&Request("username")&"'")
%>
已经将 <%=Request("username")%> 所有发表过的文章全部删除<br><br><a href=javascript:history.back()>返 回</a>
<%
end sub

sub userdel

if Request("username")=Request.Cookies("username") then error2("操作错误！")

for each ho in Request("username")

conn.execute("delete from [user] where username='"&ho&"'")
conn.execute("delete from [reforum] where username='"&ho&"'")

next

%>
已经成功删除 <%=Request("username")%> 的所有资料<br><br><a href=javascript:history.back()>返 回</a>
<%
end sub


sub userok

if Request("membercode")="" then
error2("权限不能为空！")
end if

sql="select * from [user] where username='"&HTMLEncode(Request("username"))&"'"
rs.Open sql,Conn,1,3

rs("membercode")=Request("membercode")
rs("honor")=Request("honor")
rs("posttopic")=Request("posttopic")
rs("postrevert")=Request("postrevert")
rs("experience")=Request("experience")
rs("userlife")=Request("userlife")
rs("money")=Request("money")
rs("savemoney")=Request("savemoney")
sign=server.htmlencode(Request("sign"))
sign=replace(sign,vbCrlf,"<br>")
sign=replace(sign,"\","&#92;")
rs("sign")=sign
rs.update
rs.close
%> 更新成功</b></font></td>
</tr></table><br><br><a href=javascript:history.back()>返 回</a>
<%
end sub

htmlend

%>