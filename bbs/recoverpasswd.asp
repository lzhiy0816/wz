<!-- #include file="setup.asp" -->
<!-- #include file="inc/MD5.asp" -->


<%
username=server.htmlencode(Trim(Request("username")))
userpass=Trim(Request("userpass"))
if Request("username")="" then
error("<li>�û�����û����д")
end if

sql="select * from [user] where username='"&HTMLEncode(username)&"'"
rs.Open sql,Conn,1,3
if rs.eof then
error("<li>"&username&"���û����ϲ�����")
end if
if ""&rs("birthday")&""="" then
error("<li>��ע���ʱ��û����д�������ڣ������޷�ͨ���˹����һ�����")
end if
if Request("birthday")<>rs("birthday") then
error("<li>����������д����")
end if

if Request("menu")=2 then

if ""&rs("answer")&""="" or ""&rs("question")&""="" then
error("<li>��ע���ʱ��û����д������ʾ�������������ʾ�𰸣������޷�ͨ���˹����һ�����")
end if


if md5(Request("answer"))<>rs("answer") then
error("<li>�𰸴���")
end if
if Request("userpass")="" then
error("<li>�������µ�����")
end if
if Request("userpass")<>Request("userpass2") then
error("<li>��2����������벻ͬ")
end if

rs("userpass")=md5(userpass)
rs.update
rs.close


message=message&"<li>��������ɹ�<li><a href=main.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=main.asp>")


end if

top
%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>




<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%">
<a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="����ҳ"></a></td>
<td vAlign="center" align="top">&nbsp;<font color=000000><img
src="images/closedfold.gif" width="14" height="14">��<a href=main.asp><%=clubname%></a><br>
&nbsp;<img
height=15
src="images/coner.gif"
width=15><img
src="images/openfold.gif">����������</font></td>
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
<td width="100%" align="center" height="20" class=a1 colspan="2">��������</td>
</tr>
<tr class=a3>
<td width="50%" align="right" height="20">��ش����⣺</td>
<td width="50%" height="20"><%=rs("question")%></td>
</tr>
<tr class=a4>
<td width="50%" align="right" height="20">�𰸣�</td>
<td width="50%" height="20"><input size="15" value name="answer"></td>
</tr>
<tr class=a3>
<td width="50%" align="right" height="20">�������µ����룺</td>
<td width="50%" height="20"><input type="text" size="15" name="userpass"></td>
</tr>
<tr class=a4>
<td width="50%" align="right" height="20">���ٴ��������룺</td>
<td width="50%" height="20"><input type="text" size="15" name="userpass2"></td>
</tr>

<tr>
<td width="100%" align="center" height="20" colspan="2" bgcolor="FFFFFF"><input type="submit" value=" ȷ�� " name="Submit1">��<input type="reset" value=" ȡ�� " name="Submit"></td>
</tr>
</form>

</table>



<br>
<center>
<a href=javascript:history.back()>BACK </a><br>

<%
htmlend
%>