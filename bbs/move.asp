<!-- #include file="setup.asp" -->
<%

id=int(Request("id"))

forumid=Conn.Execute("Select forumid From forum where id="&id)(0)

top
sql="select * from forum where ID="&ID&" and forumid="&forumid&""
rs.Open sql,Conn
%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>

<style>TABLE{BORDER-TOP:0px;BORDER-LEFT:0px;BORDER-BOTTOM:1px}TD{BORDER-RIGHT:0px;BORDER-TOP:0px}.title {FONT-SIZE: 9pt; COLOR: #FFFFFF;}
</style>
<title><%=rs("topic")%></title>



<center>

<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
<td vAlign="center" align="top">　<img src="images/closedfold.gif" border="0">　<a href=main.asp><%=clubname%></a><br>
　<img src="images/bar.gif" border="0"><img src="images/closedfold.gif" border="0">　<a href="ShowForum.asp?forumid=<%=forumid%>"><%=Conn.Execute("Select bbsname From bbsconfig where id="&forumid)(0)%></a><br>
　　&nbsp;<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">　<a onclick=min_yuzi() target=message href="ShowPost.asp?id=<%=id%>"><%=rs("topic")%></a></td>
</tr>
</table>
<br>

<form method="post" action=manage.asp>
<input type="hidden" value="move" name=menu>
<input type="hidden" value="<%=id%>" name="id">

<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class=a2>
<tr><td width="328" class=a1 height="25"><div align="center"><span>
移动帖子</span></div>
</td></tr><tr><td height="19" width="328" valign="top" class=a4>
<div align="center"><p><br><select name=moveid>




<option selected value="">将主题移动到...</option>

<%
rs.close


sql="select id,bbsname,classid from bbsconfig where hide=0 order by classid,id"
rs.Open sql,Conn
do while not rs.eof
Classid=Trim(Rs("classid"))
		if TClass <> Classid Then
		Response.write "<option style=BACKGROUND-COLOR:ECF5FF value=''>╋ "&Conn.Execute("Select classname From class where id="&Classid)(0)&"</option>"
		TClass = Classid
		end if
	Response.write "<option value="&rs("id")&">　├『"&rs("bbsname")&"』</option>"
	rs.movenext
loop
rs.close
%>
</select>

<br><br>
<input type=submit value=" 确 定 ">
</div></td></tr></table>
</form>

<a href=javascript:history.back()>BACK </a><br>






<%
htmlend
%>