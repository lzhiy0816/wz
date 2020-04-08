<!-- #include file="setup.asp" -->
<%
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")


if Request.Cookies("userpass") <> Conn.Execute("Select userpass From [user] where username='"&Request.Cookies("username")&"'")(0) then
error("<li>您的密码错误")
end if

membercode=Conn.Execute("Select membercode From [user] where username='"&Request.Cookies("username")&"'")(0)



if Request("menu")="ok" then

if membercode < 4 then
error("<li>您的权限不够，无法抓人入狱！")
end if

If conn.Execute("Select username From [user] where username='"&Request("username")&"'" ).eof Then error("<li>"&Request("username")&"的资料不存在")

if Request("username")="" then
error("<li>犯人的名称没有添写")
end if
if Request("causation")="" then
error("<li>您没有输入入狱原因")
end if



rs.Open "prison",Conn,1,3
rs.addnew
rs("username")=""&Trim(Request("username"))&""
rs("causation")=server.htmlencode(Request("causation"))
rs("constable")=""&Request.Cookies("username")&""
rs.update
rs.close


end if

if Request("menu")="release" then

if membercode < 4 then
error("<li>您的权限不够，无法释放犯人！")
end if

conn.execute("update [user] set membercode=1 where username='"&HTMLEncode(Request("username"))&"' and membercode=0")
conn.execute("delete from [prison] where username='"&HTMLEncode(Request("username"))&"'")

log("将 "&Request("username")&" 释放出监狱！")

error2("已经将 "&Request("username")&" 释放出监狱！")
end if


if Request("menu")="look" then
sql="select * from prison where username='"&HTMLEncode(Request("username"))&"'"
rs.Open sql,Conn


%>
<HTML><HEAD><meta http-equiv=Content-Type content=text/html; charset=gb2312><link href=images/skins/<%=Request.Cookies("skins")%>/bbs.css rel=stylesheet>
</HEAD>
<TITLE>探 监</TITLE>

<BODY background="images/plus/qiu.gif" topMargin=0>
<br><b><%=rs("username")%></b>
<SCRIPT>
tips = new Array;
tips[0] = "斜着眼睛瞟了一眼看守,嘟哝着:最近的点儿太背!";
tips[1] = "两眼汪汪的说:都是我不好!对不起大家了!";
tips[2] = "脸上露出诡异的笑容:嘿嘿!要不要进来看看!";
tips[3] = "感慨万分道:一失足,成千古恨!我一定重新做人!";
tips[4] = "毒毒地点点头:哼!二十天后,老子要重新做人!";
tips[5] = "望着布满电网和铁丝网的高墙,摇头叹息着!";
index = Math.floor(Math.random() * tips.length);
document.write("" + tips[index] + "");
  </SCRIPT><br><br>
入狱原因：<%=rs("causation")%><br><br>
入狱时间：<%=rs("cometime")%><br><br>
出狱时间：<%=rs("cometime")+prison%><br><br>

执行警官：<%=rs("constable")%>

<%
rs.close
responseend

end if

top

conn.execute("delete from [prison] where cometime<"&SqlNowString&"-"&prison&"")


%>
<TITLE>社区临时看守所</TITLE>

<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
<td vAlign="center" align="top">&nbsp;<font color=000000><img src="images/closedfold.gif">　<a href=main.asp><%=clubname%></a><br>
&nbsp;<img src="images/coner.gif"><img src="images/openfold.gif">　<a href="prison.asp">社区监狱</a></font></td>
</tr>
</table><br>

<SCRIPT>valigntop()</SCRIPT>
<CENTER><TABLE cellSpacing=1 cellPadding=0 border=0 class=a2 width=97%><TBODY>
<TR><TD valign=middle>
<TABLE cellSpacing=0 cellPadding=3 border=0 width=100%><TBODY><TR class=a1 height="25">
	<TD align=center><b>名字</b></TD>
<TD align=center><b>入狱时间</b></TD>
<TD align=center><b>执行警官</b></TD>
<TD align=center><b>动作</b></TD>
</TR>

<%
rs.Open "prison order by cometime Desc",Conn
Do While Not RS.EOF


response.write "<tr class=a4><TD align=center><a href=Profile.asp?username="&rs("username")&">"&rs("username")&"</a></TD><TD align=center>"&rs("cometime")&"</TD><TD align=center><a href=Profile.asp?username="&rs("constable")&">"&rs("constable")&"</a></TD><TD align=center><a href=?menu=release&username="&rs("username")&">释 放</a> | <a style=cursor:hand onClick=javascript:open('prison.asp?menu=look&username="&rs("username")&"','','resizable,scrollbars,width=220,height=160')>探 监</a></TD></tr>"


RS.MoveNext
loop
RS.Close

%>

<form METHOD=POST><input type=hidden name=menu value=ok><tr>
  <TD align=center colspan="4" class=a3>将
<input name="username" size="13"> 抓入监狱<br>
入狱原因：<input name="causation" size="20"> <input type="submit" value="确定"></TD>
  			</tr></form>
</TBODY></TABLE></TD></TR></TBODY></TABLE>
<SCRIPT>valignbottom()</SCRIPT>


<%
htmlend
%>