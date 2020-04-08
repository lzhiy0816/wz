<!-- #include file="setup.asp" -->
<%
menu=int(Request("menu"))
top
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")
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
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
<td>　<img src="images/closedfold.gif" border="0">　<a href=main.asp><%=clubname%></a><br>
　<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">　<span id=bbsxpname></span></td>
</tr>
</table>
<br>



<table cellspacing="1" cellpadding="1" width="97%" align="center" border="0" class="a2">
  <TR id=TableTitleLink class=a1 height="25">
      <Td width="14%" align="center"><a href="?">社区新帖</a></td>
      <TD width="14%" align="center"><a href="?menu=1">本周人气帖子</a></td>
      <TD width="14%" align="center"><a href="?menu=2">本周热门帖子</a></td>
      <TD width="14%" align="center"><a href="?menu=3">精华帖子</a></td>
      <TD width="14%" align="center"><a href="?menu=4">投票帖子</a></td>
      <TD width="14%" align="center"><a href="?menu=5">我的帖子</a></td>
      </TR></TABLE>


<br>

<SCRIPT>valigntop()</SCRIPT>
<table cellspacing="1" cellpadding="0" width="97%" align="center" border="0" class="a2"><tr height="25" id=TableTitleLink><td width="3%" class="a1">　</td><td width="3%" class="a1">　</td><td align="middle" height="24" class="a1" width="45%">主题</td><td align="middle" width="9%" height="24" class="a1">作者</font></td><td align="middle" width="6%" height="24" class="a1">回复</td><td align="middle" width="7%" height="24" class="a1">点击</td><td width="27%" height="24" class="a1" align="center">最后更新</td></tr>
<%
select case Request("menu")
case ""
sql="select top 20 * from forum where deltopic<>1 order by id Desc"
%><SCRIPT>bbsxpname.innerText="社区新帖"</SCRIPT><%
case "1"
sql="select top 20 * from forum where deltopic<>1 and posttime>"&SqlNowString&"-7 order by Views Desc"
%><SCRIPT>bbsxpname.innerText="本周人气帖子"</SCRIPT><%
case "2"
sql="select top 20 * from forum where deltopic<>1 and posttime>"&SqlNowString&"-7 order by replies Desc"
%><SCRIPT>bbsxpname.innerText="本周热门帖子"</SCRIPT><%
case "3"
sql="select top 20 * from forum where goodtopic=1 and deltopic<>1 order by id Desc"
%><SCRIPT>bbsxpname.innerText="精华帖子"</SCRIPT><%
case "4"
sql="select top 20 * from forum where polltopic<>'' and deltopic<>1 order by id Desc"
%><SCRIPT>bbsxpname.innerText="投票帖子"</SCRIPT><%
case "5"
sql="select top 20 * from forum where username='"&Request.Cookies("username")&"' and deltopic<>1 order by id Desc"
%><SCRIPT>bbsxpname.innerText="我的帖子"</SCRIPT><%
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
	<br><center><table cellspacing="4" cellpadding="0" width="80%" border="0"><tr><td nowrap width="200"><img alt="" src="images/f_new.gif" border="0"> 打开主题 (有回复的主题)</td><td nowrap width="100"><img alt="" src="images/f_hot.gif" border="0"> 热门主题 </td><td nowrap width="100"><img alt="" src="images/f_locked.gif" border="0"> 关闭主题</td><td nowrap width="150"><img src="images/topicgood.gif"> 精华主题</td></tr><tr><td nowrap width="200"><img alt="" src="images/f_norm.gif" border="0"> 打开主题 (没有回复的主题)</td><td nowrap width="100"><img alt="" src="images/f_poll.gif" border="0"> 投票主题</td><td nowrap width="100"><img alt="" src="images/f_top.gif" border="0"> 置顶主题</td><td nowrap width="150"><img src="images/my.gif"> 自己发表的主题</td></tr></table></center></div></td></tr></table>
<iframe height=0 width=0 name=hiddenframe></iframe>


<%
htmlend
%>