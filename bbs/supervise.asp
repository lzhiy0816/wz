<!-- #include file="setup.asp" -->
<%
top
if Request.Cookies("username")=empty then error("<li>����δ<a href=login.asp>��¼</a>����")

forumid=int(Request("forumid"))


moderated=split(Conn.Execute("Select moderated From [bbsconfig] where id="&forumid&" ")(0),"|")

if membercode<4 and moderated(1)<>Request.Cookies("username") then
error("<li>����Ȩ�޲���<li>ֻ�������� <font color=red>"&moderated(1)&"</font> ������Ա������������ӵ�д�Ȩ��")
end if


select case Request("menu")
case "up"

if Request("bbsname")="" then error("<li>��������̳����")

if Len(Request("intro"))>255 then  error("<li>��̳��鲻�ܴ��� 255 ���ֽ�")

temp="|"&Request("moderated")&"|"
temp=replace(temp,"||","|")
master=split(temp,"|")
for i = 1 to ubound(master)-1
If conn.Execute("Select id From [user] where username='"&HTMLEncode(master(i))&"'" ).eof and master(i)<>"" Then error("<li>"&master(i)&"���û����ϻ�δע��")
next


sql="select * from bbsconfig where id="&forumid&""
rs.Open sql,Conn,1,3
rs("bbsname")=server.htmlencode(Request("bbsname"))
rs("moderated")=temp
rs("intro")=server.htmlencode(Request("intro"))
rs("icon")=server.htmlencode(Request("icon"))
rs("logo")=server.htmlencode(Request("logo"))
rs.update

rs.close

log("������̳��ID:"&forumid&"������Ϣ��")

message="<li>���³ɹ���<li><a href=ShowForum.asp?forumid="&forumid&">������̳</a><li><a href=main.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=ShowForum.asp?forumid="&forumid&">")


case "delusertopic"

username=HTMLEncode(request("username"))
if username="" then error("<li>��û�������ַ���")
conn.execute("update [forum] set toptopic=0,deltopic=1,lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where username='"&request("username")&"' and forumid="&forumid&" ")
error2("�Ѿ��� "&username&" ������ȫ��ɾ���ˣ�")

case "deltopic"


topic=server.htmlencode(request("topic"))
if topic="" then error("<li>��û�������ַ���")
conn.execute("update [forum] set toptopic=0,deltopic=1,lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where topic like '%"&topic&"%' and forumid="&forumid&" ")
error2("�Ѿ�������������� "&topic&" ������ȫ��ɾ���ˣ�")


end select




sql="select * from bbsconfig where id="&forumid&""
rs.Open sql,Conn

%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
 <title>������̳</title><table width=97% align="center" border="0"><tr><td vAlign="top" width="30%">
 <SCRIPT>if ("<%=rs("logo")%>"!=''){document.write("<img border=0 src=<%=rs("logo")%> onload='javascript:if(this.height>60)this.height=60;'")}else{document.write("<img border=0 src=images/lzylogo.gif>")}</SCRIPT>
</td><td vAlign="center" align="top">��<img src="images/closedfold.gif" border="0">��<a href=main.asp><%=clubname%></a><br>
��<img src="images/bar.gif" border="0"><img src="images/closedfold.gif" border="0">��<a href="ShowForum.asp?forumid=<%=forumid%>"><%=rs("bbsname")%></a><br>����&nbsp;<img src="images/bar.gif" border="0"><img src="images/openfold.gif">��������̳</td></tr></table>
<br><center>

<SCRIPT>valigntop()</SCRIPT>
<table width=97% cellspacing=1 cellpadding=4 border=0 class=a2>
<tr>
<td height="20" align="center" colspan="2" class=a1><b>������̳</b></td>
</tr>
<tr>
<td colspan="2" height="25" align="left" bgcolor="#FFFFFF"><b>&nbsp;<img src="images/1.gif">��������</b></td>
</tr>
<tr class=a3>
<td align="right" width="20%"><form name="form" method="post" action="supervise.asp?menu=delusertopic&forumid=<%=forumid%>">
����û� </td>
<td align="left" width="77%">
<input size="30" name="username"> ����������
<input type="submit" value=" ȷ ��" name="Submit0"></td></form>
</tr>
<tr class=a4>
<td align="right" width="20%"><form name="form" method="post" action="supervise.asp?menu=deltopic&forumid=<%=forumid%>">
ɾ����������� </td>
<td align="left" width="77%">
<input size="30" name="topic"> ����������
<input type="submit" value=" ȷ ��" name="Submit1"></td></form>
</tr>
<form name="form" method="post" action="supervise.asp">
<input type=hidden name=menu value="up">
<input type=hidden name=forumid value="<%=forumid%>">
<tr>
<td colspan="2" height="25" align="left" bgcolor="#FFFFFF"><b>&nbsp;<img src="images/1.gif">��̳��Ϣ</b></td>
</tr>
<tr>
<td class=a3 height="5" align="right" valign="middle" width="20%">��̳����<b>��</b></td>
<td class=a3 height="5" align="left" valign="middle" width="77%">
<input type="text" name="bbsname" size="30" maxlength="12" value="<%=rs("bbsname")%>">
</td>
</tr>
<tr>
<td class=a4 height="2" align="right" valign="middle" width="20%">��̳����<b>��</b></td>
<td class=a4 height="2" align="left" valign="middle" width="77%">
<input size="30" name="moderated" value="<%=rs("moderated")%>">
������������|�ָ����磺yuzi|ԣԣ
</td>
</tr>
<tr class=a3>
<td height="2" align="right" width="20%">��̳����<b>��</b></td>
<td height="2" align="left" valign="middle" width="77%">
<textarea name="intro" rows="4" cols="50"><%=rs("intro")%></textarea>&nbsp;
</td>
</tr>
<tr>
<td class=a4 height="1" align="right" valign="middle" width="20%">Сͼ��URL��</td>
<td class=a4 height="1" align="left" valign="middle" width="77%">
<input size="30" name="icon" value="<%=rs("icon")%>">����ʾ��������ҳ��̳�������</td>
</tr>
<tr>
<td class=a3 align="right" valign="bottom" width="20%">��ͼ��URL��</td>
<td class=a3 align="left" valign="bottom" width="77%">
<input size="30" name="logo" value="<%=rs("logo")%>">����ʾ����̳���Ͻ�</td>
</tr>
<tr>
<td align="right" width="98%" colspan="2" bgcolor="#FFFFFF">
<input type="submit" value=" �� �� &gt;&gt;�� һ �� " name="Submit"> </td>
</tr>
</table>
<SCRIPT>valignbottom()</SCRIPT>
</form>
<%
rs.close
htmlend
%>