<!-- #include file="setup.asp" -->
<%
menu=int(Request("menu"))
top
if Request.Cookies("username")=empty then error("<li>����δ<a href=login.asp>��¼</a>����")
%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>

<script>
if(top==self)document.location='Default.asp';
var key_word="<%=Request.Cookies("key_word")%>"
var cookieusername="<%=Request.Cookies("username")%>"
</script>


<title><%=clubname%></title>


<table width=97% align="center" border="0" id="table1">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="����ҳ"></a></td>
<td>��<img src="images/closedfold.gif" border="0">��<a href=main.asp><%=clubname%></a><br>
��<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">��<span id=bbsxpname></span></td>
</tr>
</table>
<br>



<table cellspacing="1" cellpadding="1" width="97%" align="center" border="0" class="a2">
  <TR id=TableTitleLink class=a1 height="25">
      <Td width="14%" align="center"><a href="?">��������</a></td>
      <TD width="14%" align="center"><a href="?menu=1">������������</a></td>
      <TD width="14%" align="center"><a href="?menu=2">������������</a></td>
      <TD width="14%" align="center"><a href="?menu=3">��������</a></td>
      <TD width="14%" align="center"><a href="?menu=4">ͶƱ����</a></td>
      <TD width="14%" align="center"><a href="?menu=5">�ҵ�����</a></td>
      </TR></TABLE>


<br>

<SCRIPT>valigntop()</SCRIPT>
<table cellspacing="1" cellpadding="0" width="97%" align="center" border="0" class="a2"><tr height="25" id=TableTitleLink><td width="3%" class="a1">��</td><td width="3%" class="a1">��</td><td align="middle" height="24" class="a1" width="45%">����</td><td align="middle" width="9%" height="24" class="a1">����</font></td><td align="middle" width="6%" height="24" class="a1">�ظ�</td><td align="middle" width="7%" height="24" class="a1">���</td><td width="27%" height="24" class="a1" align="center">������</td></tr>
<%
select case Request("menu")
case ""
sql="select top 20 * from forum where deltopic<>1 order by id Desc"
%><SCRIPT>bbsxpname.innerText="��������"</SCRIPT><%
case "1"
sql="select top 20 * from forum where deltopic<>1 and posttime>"&SqlNowString&"-7 order by Views Desc"
%><SCRIPT>bbsxpname.innerText="������������"</SCRIPT><%
case "2"
sql="select top 20 * from forum where deltopic<>1 and posttime>"&SqlNowString&"-7 order by replies Desc"
%><SCRIPT>bbsxpname.innerText="������������"</SCRIPT><%
case "3"
sql="select top 20 * from forum where goodtopic=1 and deltopic<>1 order by id Desc"
%><SCRIPT>bbsxpname.innerText="��������"</SCRIPT><%
case "4"
sql="select top 20 * from forum where polltopic<>'' and deltopic<>1 order by id Desc"
%><SCRIPT>bbsxpname.innerText="ͶƱ����"</SCRIPT><%
case "5"
sql="select top 20 * from forum where username='"&Request.Cookies("username")&"' and deltopic<>1 order by id Desc"
%><SCRIPT>bbsxpname.innerText="�ҵ�����"</SCRIPT><%
end select





rs.Open sql,Conn

Do While Not RS.EOF

newtopic=""
if rs("posttime")+1>now() then newtopic="<img src=images/new.gif>"

%><script>ShowForum("<%=rs("ID")%>","<%=rs("topic")%>","<%=newtopic%>","<%=rs("username")%>","<%=rs("Views")%>","<%=rs("icon")%>","<%=rs("toptopic")%>","<%=rs("locktopic")%>","<%=rs("pollresult")%>","<%=rs("goodtopic")%>","<%=rs("replies")%>","<%=rs("lastname")%>","<%=rs("lasttime")%>");</script>
<%
RS.MoveNext
loop
RS.Close
%></table>
<SCRIPT>valignbottom()</SCRIPT>
	<br><center><table cellspacing="4" cellpadding="0" width="80%" border="0"><tr><td nowrap width="200"><img alt="" src="images/f_new.gif" border="0"> ������ (�лظ�������)</td><td nowrap width="100"><img alt="" src="images/f_hot.gif" border="0"> �������� </td><td nowrap width="100"><img alt="" src="images/f_locked.gif" border="0"> �ر�����</td><td nowrap width="150"><img src="images/topicgood.gif"> ��������</td></tr><tr><td nowrap width="200"><img alt="" src="images/f_norm.gif" border="0"> ������ (û�лظ�������)</td><td nowrap width="100"><img alt="" src="images/f_poll.gif" border="0"> ͶƱ����</td><td nowrap width="100"><img alt="" src="images/f_top.gif" border="0"> �ö�����</td><td nowrap width="150"><img src="images/my.gif"> �Լ����������</td></tr></table></center></div></td></tr></table>
<iframe height=0 width=0 name=hiddenframe></iframe>


<%
htmlend
%>