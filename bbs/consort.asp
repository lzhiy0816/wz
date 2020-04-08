<!-- #include file="setup.asp" -->
<%
id=int(Request("id"))

content=replace(server.htmlencode(Request.Form("content")), "'", "&#39;")

top
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")


consort=Conn.Execute("Select consort From [user] where username='"&Request.Cookies("username")&"'")(0)







select case Request("menu")
case "add"

if content=empty then error("<li>表白内容不能为空")

aim=server.htmlencode(Trim(Request("aim")))

if aim=Request.Cookies("username") then error("<li>不能自己追求自己！")

If conn.Execute("Select id From [user] where username='"&aim&"'" ).eof Then
error("<li>系统不存在"&aim&"的资料")
end if

If Not conn.Execute("Select id From [consort] where username='"&Request.Cookies("username")&"' and aim='"&aim&"'" ).eof Then
error("<li>"&aim&"已经在您的追求列表当中")
end if


sql="insert into consort (username,aim,unburden) values ('"&Request.Cookies("username")&"','"&aim&"','"&content&"')"
conn.Execute(SQL)


sql="insert into message(author,incept,content) values ('"&Request.Cookies("username")&"','"&aim&"','"&content&"')"
conn.Execute(SQL)




conn.execute("update [user] set newmessage=newmessage+1 where username='"&aim&"'")

case "accept"

if consort<>empty then error("<li>您当前已有配偶")

aim=Conn.Execute("Select aim From [consort] where id="&id&"")(0)
if aim<>Request.Cookies("username") then error("<li>非法操作")

consort=Conn.Execute("Select username From [consort] where id="&id&"")(0)
if Conn.Execute("Select consort From [user] where username='"&consort&"'")(0)<>empty then error("<li>"&consort&"已经有配偶了")

conn.execute("update [user] set consort='"&aim&"' where username='"&consort&"'")
conn.execute("update [user] set consort='"&consort&"' where username='"&aim&"'")
conn.execute("delete from [consort] where id="&id&"")
succeed("<li>您已经接受了"&consort&"的追求<li><a href=consort.asp>返回社区配偶</a><meta http-equiv=refresh content=3;url=consort.asp>")



case "del"
conn.execute("delete from [consort] where id="&id&" and (username='"&Request.Cookies("username")&"' or aim='"&Request.Cookies("username")&"') ")

case "part"
conn.execute("update [user] set consort='' where username='"&consort&"'")
conn.execute("update [user] set consort='' where username='"&Request.Cookies("username")&"'")
succeed("<li>分手成功<li><a href=consort.asp>返回社区配偶</a><meta http-equiv=refresh content=3;url=consort.asp>")

end select

%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<TITLE>社区配偶</TITLE>

<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
<td vAlign="center" align="top">&nbsp;<font color=000000><img src="images/closedfold.gif">　<a href=main.asp><%=clubname%></a><br>
&nbsp;<img src="images/coner.gif"><img src="images/openfold.gif">　<a href="consort.asp">社区配偶</a></font></td>
</tr>
</table><br>
<center>
<SCRIPT>valigntop()</SCRIPT>
		<table cellspacing="1" cellpadding="6" width="97%" border="0" class=a2>
			<tr>
				<td width="100%" height="10" align="middle" class="a1" colspan="4">
				<font face="宋体">我的追求者</font></td>
			</tr>
			<tr class=a3>
				<td width="10%" height="5" align="middle">
				用户名</td>
				<td height="5" align="middle">
				爱的表白</td>
				<td width="20%" height="5" align="middle">
				<font face="宋体">追求</font>时间</td>
				<td width="15%" height="5" align="middle">
				操作</td>
			</tr>
			
<%
sql="select * from consort where aim='"&Request.Cookies("username")&"' order by id Desc"
rs.Open sql,Conn
Do While Not RS.EOF 
%>
			<tr class=a4>
				<td height="5" align="middle">
				<a href="Profile.asp?username=<%=rs("username")%>"><%=rs("username")%></a></td>
				<td height="5" align="middle">
				<%=rs("unburden")%></td>
				<td height="5" align="middle">
				<%=rs("addtime")%></td>
				<td height="5" align="middle">
				
				
<%if consort=rs("username") then%>
 <a href="?menu=part">分 手</a> 
<%else%>
 <a href="?menu=accept&id=<%=rs("id")%>">接受</a>  <a href="?menu=del&id=<%=rs("id")%>">拒绝</a>
<%end if%>
				
				</td>
			</tr>
<%
RS.MoveNext
loop
RS.Close
%>


			</table>
<SCRIPT>valignbottom()</SCRIPT>
<br><br><br>
<SCRIPT>valigntop()</SCRIPT>
		<table cellspacing="1" cellpadding="6" width="97%" border="0" class=a2>
			<tr>
				<td width="100%" height="10" align="middle" class="a1" colspan="4">
				<font face="宋体">我追求的人</font></td>
			</tr>
			<tr class=a3>
				<td width="10%" height="5" align="middle">
				用户名</td>
				<td height="5" align="middle">
				爱的表白</td>
				<td width="20%" height="5" align="middle">
				<font face="宋体">追求</font>时间</td>
				<td width="15%" height="5" align="middle">
				操作</td>
			</tr>
<%
sql="select * from consort where username='"&Request.Cookies("username")&"' order by id Desc"
rs.Open sql,Conn
Do While Not RS.EOF 
%>
			<tr class=a4>
				<td height="5" align="middle">
				<a href="Profile.asp?username=<%=rs("aim")%>"><%=rs("aim")%></a></td>
				<td height="5" align="middle">
				<%=rs("unburden")%></td>
				<td height="5" align="middle">
				<%=rs("addtime")%></td>
				<td height="5" align="middle">
<%if consort=rs("aim") then%>
<a href="?menu=part">分 手</a> 
<%else%>
 <a href="?menu=del&id=<%=rs("id")%>">取消追求</a>
<%end if%>

</td>
			</tr>
<%
RS.MoveNext
loop
RS.Close

on error resume next  '找不到配偶头像时候忽略错误
%>
			</table>
<SCRIPT>valignbottom()</SCRIPT>
<br><br><br>

<form action="consort.asp" method="post">
<input type="hidden" value="add" name="menu">
<SCRIPT>valigntop()</SCRIPT>
		<table cellspacing="1" cellpadding="6" width="97%" border="0" class=a2>
				<tr>
		
<%if consort<>empty then%>	
				<td width="19%" height="10" align="middle" class="a3" rowspan="4">
				<img src="<%=Conn.Execute("Select userface From [user] where username='"&consort&"'")(0)%>">
				
				
<br><br>
				当前配偶：<a href="profile.asp?username=<%=consort%>"><%=consort%></a><br>
				<br>
				<a href="?menu=part">与当前配偶分手</a></td>

<%end if%>			
			
				<td width="77%" height="2" align="middle" class="a1" colspan="2">
				<font face="宋体">添加我想追求的人</font></td>
			</tr>			
			<tr>
				<td width="12%" height="2" align="left" class="a4">
				对方用户名：</td>
				<td width="64%" height="2" align="left" class="a4">
				<input maxLength="30" size="15" name="aim"></td>
			</tr>
			<tr>
				<td width="12%" height="1" align="left" class="a4" valign="top">
				爱的表白：</td>
				<td width="64%" height="1" align="left" class="a4">
				<textarea name="content" rows="5" cols="60"></textarea></td>
			</tr>
			<tr>
				<td width="77%" height="1" align="center" class="a4" colspan="2">
				<input tabIndex="4" type="submit" value=" 确 定 ">
				<input onclick="{if(confirm('该项操作要清除全部的内容，您确定要清除吗?')){return true;}return false;}" type="reset" value=" 重 写 "></td>
			</tr>
			</table>
<SCRIPT>valignbottom()</SCRIPT>

			</form>
<%htmlend%>