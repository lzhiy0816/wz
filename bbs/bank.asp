<!-- #include file="setup.asp" -->
<%
top
if Request.Cookies("username")=empty then error("<li>����δ<a href=login.asp>��¼</a>����")

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
<title>��������</title>
<center><br>
<table cellSpacing="0" cellPadding="0" width="622" border="0" style="border-collapse: collapse" bordercolor="111111"><tr>
<td align="center" width="590">
<table border="0" cellpadding="4" cellspacing="1" width="83%" class=a2>
<tr>
<td width="50%" rowspan="4" class=a4><img src="images/plus/bank.gif" width="249" height="90"></td>
<td width="50%" colspan="2" class=a1 align="center">���������˺�</td>
</tr>
<tr class=a4>
<td width="24%">&nbsp;�ֽ�</td>
<td width="26%"><b><font color="aa0000"><%=rs("money")%> </font>
���</b></td>
</tr>
<tr class=a4>
<td width="24%">&nbsp;��</td>
<td width="26%"><b><font color="aa0000"><%=rs("savemoney")%></font>
���</b></td>
</tr>
<tr class=a4>
<td width="24%">&nbsp;�ܹ���</td>
<td width="26%"><b><font color="aa0000"><%=rs("savemoney")+rs("money")%></font>
���</b></td>
</tr>
</table>
<p><br></p>


<table cellSpacing="1" cellPadding="3" border="0" width="377" height="47" class=a2><tr>
<td class=a1 height="25">&nbsp; <b>��Ҫ���</b>&nbsp;</td>
<td class=a1 height="25" align="center">�����ֽ� <b><%=rs("money")%></b> <b>���</b></td></tr><tr class=a4>
<td height="25" align="center">
<form action="bank.asp?menu=save" method="post">&nbsp; ��Ҫ��
<input size="10" value="1000" name="qmoney" MAXSIZE="32"><b> ���</b>
</td>
<td height="25" align="center">
<input type="submit" value=" �� �� " name="B2"></td></form></tr></table>
<br><br>



<table cellSpacing="1" cellPadding="3" border="0" width="377" height="47" class=a2><tr>
<td class=a1 height="25">&nbsp; <b>��Ҫȡ��</b>&nbsp;
</td>
<td class=a1 height="25" align="center">���д�� <b><%=rs("savemoney")%></b> <b>���</b></td></tr><tr class=a4>
<td height="25" align="center">
<form action="bank.asp?menu=draw" method="post">&nbsp; ��Ҫȡ
<input size="10" value="1000" name="qmoney" MAXSIZE="32"><b> ���</b>
</td>
<td height="25" align="center">
<input type="submit" value=" ȡ �� " name="B2"></td></form></tr></table>
<br><br>


<table cellSpacing="1" cellPadding="3" border="0" width="377" height="47" class=a2><tr>
<td class=a1 height="25">&nbsp; <b>��Ҫת��</b>&nbsp;
</td>
<td class=a1 height="25" align="right">���ת�˽��Ϊ <b>1000</b> <b>���</b></td></tr><tr class=a4>
<td height="25" align="center" colspan="2">
<form action="bank.asp?menu=virement" method="post">&nbsp; ��Ҫ��
<input size="5" value="1000" name="qmoney" MAXSIZE="32"><b> ���</b> ת��
<input size="10" name="dxname" MAXSIZE="32"> ���˻�
<input type="submit" value=" ȷ �� " name="B2"></td>
</form></tr></table>
<br><br>
<%
rs.close
htmlend

sub save
qmoney=int(request("qmoney"))
if qmoney > rs("money") then
message="<li>����ֽ�û����ô��ɣ�"
error(message)
end if
if qmoney<1 then
message="<li>����Ϊ�㣡"
error(message)
end if
rs("savemoney")=rs("savemoney")+qmoney
rs("money")=rs("money")-qmoney
rs.update
rs.close
message="<li>���ɹ�<li><a href=bank.asp>��������</a><li><a href=main.asp>������̳��ҳ</a>"
succeed(message&"<meta http-equiv=refresh content=3;url='bank.asp'")
end sub


sub draw
qmoney=int(request("qmoney"))
if qmoney>rs("savemoney") then
message="<li>���Ĵ�����"
error(message)
end if
if qmoney<1 then
message="<li>ȡ���Ϊ�㣡"
error(message)
end if
rs("savemoney")=rs("savemoney")-qmoney
rs("money")=rs("money")+qmoney
rs.update
rs.close
message="<li>ȡ��ɹ�<li><a href=bank.asp>��������</a><li><a href=main.asp>������̳��ҳ</a>"
succeed(message&"<meta http-equiv=refresh content=3;url='bank.asp'")
end sub

sub virement
dxname=server.htmlencode(Trim(Request("dxname")))

if dxname=Request.Cookies("username") then error("<li>����������Լ����˺ţ�")


qmoney=int(request("qmoney"))
if qmoney>rs("savemoney") then
error"<li>�����ʻ���������"
end if
if qmoney<1000 then
error"<li>ת�ʲ��ܵ���1000��"
end if



If conn.Execute("Select id From [user] where username='"&dxname&"'" ).eof Then
error("<li>����"&dxname&"���˺�")
end if



rs("savemoney")=rs("savemoney")-qmoney
rs.update
rs.close

conn.execute("update [user] set savemoney=savemoney+"&qmoney&" where username='"&dxname&"'")


message="<li>ת�˳ɹ�<li><a href=bank.asp>��������</a><li><a href=main.asp>������̳��ҳ</a>"
succeed(message&"<meta http-equiv=refresh content=3;url='bank.asp'")
end sub


%>