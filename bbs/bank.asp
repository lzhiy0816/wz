<!-- #include file="setup.asp" -->
<%
top
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")

sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn,1,3

select case Request("menu")
case "save"
save
case "draw"
draw
case "virement"
virement
end select

%>
<title>社区银行</title>
<center><br>
<table cellSpacing="0" cellPadding="0" width="622" border="0" style="border-collapse: collapse" bordercolor="111111"><tr>
<td align="center" width="590">
<table border="0" cellpadding="4" cellspacing="1" width="83%" class=a2>
<tr>
<td width="50%" rowspan="4" class=a4><img src="images/plus/bank.gif" width="249" height="90"></td>
<td width="50%" colspan="2" class=a1 align="center">您的银行账号</td>
</tr>
<tr class=a4>
<td width="24%">&nbsp;现金：</td>
<td width="26%"><b><font color="aa0000"><%=rs("money")%> </font>
金币</b></td>
</tr>
<tr class=a4>
<td width="24%">&nbsp;存款：</td>
<td width="26%"><b><font color="aa0000"><%=rs("savemoney")%></font>
金币</b></td>
</tr>
<tr class=a4>
<td width="24%">&nbsp;总共：</td>
<td width="26%"><b><font color="aa0000"><%=rs("savemoney")+rs("money")%></font>
金币</b></td>
</tr>
</table>
<p><br></p>


<table cellSpacing="1" cellPadding="3" border="0" width="377" height="47" class=a2><tr>
<td class=a1 height="25">&nbsp; <b>我要存款</b>&nbsp;</td>
<td class=a1 height="25" align="center">您有现金 <b><%=rs("money")%></b> <b>金币</b></td></tr><tr class=a4>
<td height="25" align="center">
<form action="bank.asp?menu=save" method="post">&nbsp; 我要存
<input size="10" value="1000" name="qmoney" MAXSIZE="32"><b> 金币</b>
</td>
<td height="25" align="center">
<input type="submit" value=" 存 了 " name="B2"></td></form></tr></table>
<br><br>



<table cellSpacing="1" cellPadding="3" border="0" width="377" height="47" class=a2><tr>
<td class=a1 height="25">&nbsp; <b>我要取款</b>&nbsp;
</td>
<td class=a1 height="25" align="center">您有存款 <b><%=rs("savemoney")%></b> <b>金币</b></td></tr><tr class=a4>
<td height="25" align="center">
<form action="bank.asp?menu=draw" method="post">&nbsp; 我要取
<input size="10" value="1000" name="qmoney" MAXSIZE="32"><b> 金币</b>
</td>
<td height="25" align="center">
<input type="submit" value=" 取 了 " name="B2"></td></form></tr></table>
<br><br>


<table cellSpacing="1" cellPadding="3" border="0" width="377" height="47" class=a2><tr>
<td class=a1 height="25">&nbsp; <b>我要转帐</b>&nbsp;
</td>
<td class=a1 height="25" align="right">最低转账金额为 <b>1000</b> <b>金币</b></td></tr><tr class=a4>
<td height="25" align="center" colspan="2">
<form action="bank.asp?menu=virement" method="post">&nbsp; 我要将
<input size="5" value="1000" name="qmoney" MAXSIZE="32"><b> 金币</b> 转到
<input size="10" name="dxname" MAXSIZE="32"> 的账户
<input type="submit" value=" 确 定 " name="B2"></td>
</form></tr></table>
<br><br>
<%
rs.close
htmlend

sub save
qmoney=int(request("qmoney"))
if qmoney > rs("money") then
message="<li>你的现金没有这么多吧！"
error(message)
end if
if qmoney<1 then
message="<li>存款不能为零！"
error(message)
end if
rs("savemoney")=rs("savemoney")+qmoney
rs("money")=rs("money")-qmoney
rs.update
rs.close
message="<li>存款成功<li><a href=bank.asp>返回银行</a><li><a href=main.asp>返回论坛首页</a>"
succeed(message&"<meta http-equiv=refresh content=3;url='bank.asp'")
end sub


sub draw
qmoney=int(request("qmoney"))
if qmoney>rs("savemoney") then
message="<li>您的存款不够！"
error(message)
end if
if qmoney<1 then
message="<li>取款不能为零！"
error(message)
end if
rs("savemoney")=rs("savemoney")-qmoney
rs("money")=rs("money")+qmoney
rs.update
rs.close
message="<li>取款成功<li><a href=bank.asp>返回银行</a><li><a href=main.asp>返回论坛首页</a>"
succeed(message&"<meta http-equiv=refresh content=3;url='bank.asp'")
end sub

sub virement
dxname=server.htmlencode(Trim(Request("dxname")))

if dxname=Request.Cookies("username") then error("<li>您输入的是自己的账号？")


qmoney=int(request("qmoney"))
if qmoney>rs("savemoney") then
error"<li>您的帐户余额不够！！"
end if
if qmoney<1000 then
error"<li>转帐不能低于1000！"
end if



If conn.Execute("Select id From [user] where username='"&dxname&"'" ).eof Then
error("<li>查无"&dxname&"的账号")
end if



rs("savemoney")=rs("savemoney")-qmoney
rs.update
rs.close

conn.execute("update [user] set savemoney=savemoney+"&qmoney&" where username='"&dxname&"'")


message="<li>转账成功<li><a href=bank.asp>返回银行</a><li><a href=main.asp>返回论坛首页</a>"
succeed(message&"<meta http-equiv=refresh content=3;url='bank.asp'")
end sub


%>