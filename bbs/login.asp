<!-- #include file="setup.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
select case Request("menu")

case "add"


url=Request("url")
username=HTMLEncode(Request("username"))
userpass=md5(Request("userpass"))
if username=empty then error("<li>�û���û������")

If conn.Execute("Select id From [user] where username='"&username&"'" ).eof Then error("<li>���û�����δ<a href=register.asp?username="&username&">ע��</a>")

if userpass<>Conn.Execute("Select userpass From [user] where username='"&username&"'")(0) then  error("<li>��������������")

Response.Cookies("username")=username
Response.Cookies("userpass")=userpass
Response.Cookies("eremite")="0"

if Request("eremite")="1" then Response.Cookies("eremite")="1"

if Request("xuansave")=1 then
Response.Cookies("eremite").Expires=date+365
Response.Cookies("username").Expires=date+365
Response.Cookies("userpass").Expires=date+365
end if

if url<>empty and instr(url,"login.asp")=0 and instr(url,"ShowPost.asp")=0 and instr(url,"left.htm")=0 then
response.redirect url
else
response.write "<SCRIPT>top.location='Default.asp';</SCRIPT>"
end if

case "out"
conn.execute("delete from [online] where username='"&Request.Cookies("username")&"' and ip='"&remoteaddr&"'")
Response.Cookies("username")=""
Response.Cookies("userpass")=""
succtitle="�Ѿ��ɹ��˳�"
message=message&"<li><a href=main.asp>������ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=main.asp>")

end select
top
%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<title>��¼����</title>
<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="����ҳ"></a></td>
<td vAlign="center" align="top">&nbsp;<font color=000000><img src="images/closedfold.gif" width="14" height="14">��<a href=main.asp><%=clubname%></a><br>
&nbsp;<img height=15 src="images/coner.gif" width=15><img src="images/openfold.gif">����¼����</font></td>
</tr>
</table>
<br>

<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class=a2>

<form action="login.asp" method="post">
<input type="hidden" value="add" name="menu">
<input type="hidden" value="<%=Request.ServerVariables("HTTP_REFERER")%>" name="url">
<tr>
<td width="328" height="25" align="center" class=a1>
��¼����
</td></tr><tr>
<td height="19" width="328" valign="top" align="center" class=a3>
�û�����:
<input size="15" name="username" value="<%=Request.Cookies("username")%>">&nbsp; <a href="register.asp">û��ע��?</a><br>
�û�����: <input type="password" size="15" value name="userpass">&nbsp;
<a href="modification.asp">��������?</a><br>

<input type="checkbox" value="1" name="xuansave" id="xuansave"><label for="xuansave">��ס����</label>

<input type="checkbox" value="1" name="eremite" id="eremite"><label for="eremite">�����¼</label><br>
<input type="submit" value=" ��¼ " name="Submit1">��<input type="reset" value=" ȡ�� " name="Submit">
</td></tr> </form></table>

<br>
<center>
<a href=javascript:history.back()>BACK </a><br>


<%htmlend%>