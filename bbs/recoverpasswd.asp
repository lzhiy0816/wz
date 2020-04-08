<!-- #include file="setup.asp" -->
<!-- #include file="inc/MD5.asp" -->


<%
username=server.htmlencode(Trim(Request("username")))
userpass=Trim(Request("userpass"))
if Request("username")="" then
error("<li>用户名称没有填写")
end if

sql="select * from [user] where username='"&HTMLEncode(username)&"'"
rs.Open sql,Conn,1,3
if rs.eof then
error("<li>"&username&"的用户资料不存在")
end if
if ""&rs("birthday")&""="" then
error("<li>您注册的时候没有填写出生日期，所以无法通过此功能找回密码")
end if
if Request("birthday")<>rs("birthday") then
error("<li>出生日期填写错误")
end if

if Request("menu")=2 then

if ""&rs("answer")&""="" or ""&rs("question")&""="" then
error("<li>您注册的时候没有填写密码提示问题或者密码提示答案，所以无法通过此功能找回密码")
end if


if md5(Request("answer"))<>rs("answer") then
error("<li>答案错误")
end if
if Request("userpass")="" then
error("<li>请输入新的密码")
end if
if Request("userpass")<>Request("userpass2") then
error("<li>您2次输入的密码不同")
end if

rs("userpass")=md5(userpass)
rs.update
rs.close


message=message&"<li>更改密码成功<li><a href=main.asp>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=main.asp>")


end if

top
%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>




<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%">
<a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
<td vAlign="center" align="top">&nbsp;<font color=000000><img
src="images/closedfold.gif" width="14" height="14">　<a href=main.asp><%=clubname%></a><br>
&nbsp;<img
height=15
src="images/coner.gif"
width=15><img
src="images/openfold.gif">　更改密码</font></td>
</tr>
</table>


<br>











<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class=a2>
<form method="POST" action=RecoverPasswd.asp>
<input type=hidden name=username value=<%=Request("username")%>>
<input type=hidden name=birthyear value=<%=Request("birthyear")%>>
<input type=hidden name=birthmonth value=<%=Request("birthmonth")%>>
<input type=hidden name=birthday value=<%=Request("birthday")%>>
<input type=hidden name=menu value=2>
<tr>
<td width="100%" align="center" height="20" class=a1 colspan="2">更改密码</td>
</tr>
<tr class=a3>
<td width="50%" align="right" height="20">请回答问题：</td>
<td width="50%" height="20"><%=rs("question")%></td>
</tr>
<tr class=a4>
<td width="50%" align="right" height="20">答案：</td>
<td width="50%" height="20"><input size="15" value name="answer"></td>
</tr>
<tr class=a3>
<td width="50%" align="right" height="20">请输入新的密码：</td>
<td width="50%" height="20"><input type="text" size="15" name="userpass"></td>
</tr>
<tr class=a4>
<td width="50%" align="right" height="20">请再次输入密码：</td>
<td width="50%" height="20"><input type="text" size="15" name="userpass2"></td>
</tr>

<tr>
<td width="100%" align="center" height="20" colspan="2" bgcolor="FFFFFF"><input type="submit" value=" 确定 " name="Submit1">　<input type="reset" value=" 取消 " name="Submit"></td>
</tr>
</form>

</table>



<br>
<center>
<a href=javascript:history.back()>BACK </a><br>

<%
htmlend
%>