<!-- #include file="setup.asp" -->


<%
if Request.Cookies("username")=empty then error2("请登录后才能使用此功能！")

if Request.Cookies("userpass") <> Conn.Execute("Select userpass From [user] where username='"&Request.Cookies("username")&"'")(0) then error("您的密码错误")

username=HTMLEncode(Request("username"))

content=replace(server.htmlencode(Request.Form("content")), "'", "&#39;")




select case Request("menu")
case "add"
add
case "bad"
bad
case "del"
del
case "post"
post
case "look"
look
case "addpost"
addpost
case ""
index
end select



sub add
if username="" then error2("请输入您要添加的好友名字！")

if username=Request.Cookies("username") then error2("不能添加自己！")

If conn.Execute("Select id From [user] where username='"&username&"'" ).eof Then error2("数据库不存在此用户的资料！")



sql="select friend from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn,1,3
if instr(rs("friend"),"|"&username&"|")>0 then
error2("此好友已经添加！")
end if
rs("friend")=""&rs("friend")&""&username&"|"
rs.update
rs.close
error2("已经成功添加好友!")
end sub

sub bad
Response.Cookies("badlist")=""&Request.Cookies("badlist")&""&username&"|"
error2("已经成功将 "&username&" 加入黑名单!")
end sub


sub del
sql="select friend from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn,1,3
rs("friend")=replace(rs("friend"),"|"&username&"|","|")
rs.update
rs.close
index
end sub

sub look
page=Request("page")
if page<1 then
disabled="disabled=true"
page=0
end if
count=conn.execute("Select count(id)from message where incept='"&Request.Cookies("username")&"'")(0)

sql="select author,content from message where incept='"&Request.Cookies("username")&"' order by time Desc"
rs.Open sql,Conn
if Count-page<2 then
disabled2="disabled=true"
end if
if rs.eof then
error2("您没有短讯息！")
end if
RS.Move page

%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body><HTML><META http-equiv=Content-Type content="text/html; charset=gb2312">
<link href=images/skins/<%=Request.Cookies("skins")%>/bbs.css rel=stylesheet><TITLE>查看消息</TITLE>
<body topmargin=0 bgcolor="ECE9D8">
<style>
.bt {BORDER-RIGHT: 1px; BORDER-TOP: 1px; FONT-SIZE: 9pt; BORDER-LEFT: 1px; BORDER-BOTTOM: 1px;}
</style>
<TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0><TR><TD>
.&nbsp;昵称：<input readOnly type="text" value="<%=rs("author")%>" size="8"> Email：<input  readOnly type="text" value="<%=Conn.Execute("Select usermail From [user] where username='"&rs("author")&"'")(0)%>" size="10">
</TD><TD align=right><a target=_blank href=Profile.asp?username=<%=rs("author")%>><img border="0" src="<%=Conn.Execute("Select userface From [user] where username='"&rs("author")&"'")(0)%>" width="32" height="32" alt=用户详细资料>
</TD></TR><TR><TD VALIGN=top ALIGN=right colspan="2" bgcolor="F8F8F8">
    <textarea name="content" readOnly cols="39" rows="6"><%=rs("content")%></textarea>
</TD></TR></TABLE>
<TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0 height="30">
<tr ALIGN=center><TD><input type="button" value="回复讯息" onclick=javascript:open('friend.asp?menu=post&incept=<%=rs("author")%>','_top','width=320,height=170')>
</td><TD><input type="button" value="<<上一条" <%=disabled%> onclick=javascript:open('friend.asp?menu=look&page=<%=page-1%>','_top','')> </td><TD><input type="button" value="下一条>>" <%=disabled2%> onclick=javascript:open('friend.asp?menu=look&page=<%=page+1%>','_top','')>
</td>
</TR></TABLE>
</BODY></HTML>
<%
rs.close
conn.execute("update [user] set newmessage=0 where username='"&Request.Cookies("username")&"'")
end sub




sub post

if Request("incept")="" then
error2("对不起，您没有输入用户名称！")
end if


if Request("log")="1" then
log2="javascript:history.back()"
else
log2="javascript:open('friend.asp?menu=post&log=1&incept="&Request("incept")&"','_top','width=320,height=170');history.go(1)"
end if


sql="select username,userface,usermail from [user] where username='"&HTMLEncode(Request("incept"))&"'"
rs.Open sql,Conn

if rs.eof then
error2"系统不存在该用户的资料"
end if



%>

<HTML><META http-equiv=Content-Type content="text/html; charset=gb2312">
<link href=images/skins/<%=Request.Cookies("skins")%>/bbs.css rel=stylesheet>
<body topmargin=0 bgcolor="ECE9D8">
<style>
.bt {BORDER-RIGHT: 1px; BORDER-TOP: 1px; FONT-SIZE: 9pt; BORDER-LEFT: 1px; BORDER-BOTTOM: 1px;}
</style><TITLE>发送消息</TITLE>
<SCRIPT>
function check(theForm) {
if(theForm.content.value == "" ) {
alert("不能发空讯息！");
return false;
}
if (theForm.content.value.lengtd > 255){
alert("对不起，您的讯息不能超过 255 个字节！");
return false;
}
}
function presskey(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.form.submit();}}
</SCRIPT>
<TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0><TR><form name=form action="friend.asp" method="post">
<input type="hidden" name="menu" value="addpost">
<input type="hidden" name="incept" value="<%=rs("username")%>">
<TD>
&nbsp;昵称：<input readOnly type="text" value="<%=rs("username")%>" size="8"> Email：<input  readOnly type="text" value="<%=rs("usermail")%>" size="10">
</TD><TD align=right><a target=_blank href=Profile.asp?username=<%=rs("username")%>><img border="0" src="<%=rs("userface")%>" width="32" height="32" alt=用户详细资料>
</TD></TR><TR><TD VALIGN=top ALIGN=right colspan="2" bgcolor="F8F8F8">
    <textarea name="content" cols="39" rows="6" onkeydown=presskey()></textarea>
</TD></TR></TABLE><TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0 height="30">
<tr ALIGN=center><TD><input type="button" value="聊天记录" onclick=<%=log2%>>
</td><TD><input type="reset" value="取消发送" OnClick="window.close();"> </td><TD><input type="submit" value="发送讯息" onclick="return check(this.form)"></td>
</TR></form>
</TABLE>
<%
rs.close
if Request("log")<>"" then
%>
<body onload=resizeTo(330,300)>
<textarea name="content" readOnly cols="40" rows="6"><%
sql="select * from message where (author='"&Request.Cookies("username")&"' and incept='"&Request("incept")&"') or (author='"&Request("incept")&"' and incept='"&Request.Cookies("username")&"') order by time Desc"
rs.Open sql,Conn
do while not rs.eof
%>
(<%=rs("time")%>)   <%=rs("author")%>　
<%=rs("content")%>
<%
rs.movenext
loop
rs.close
%></textarea>

<%
end if

response.write "<body onload=resizeTo(330,206)>"
end sub

sub addpost

incept=server.htmlencode(Trim(Request("incept")))
if Request("incept")=Request.Cookies("username") then error2("不能给自己发送讯息！")
If conn.Execute("Select id From [user] where username='"&incept&"'" ).eof Then error("<li>系统不存在"&incept&"的资料")



sql="insert into message(author,incept,content) values ('"&Request.Cookies("username")&"','"&incept&"','"&content&"')"
conn.Execute(SQL)


conn.execute("update [user] set newmessage=newmessage+1 where username='"&incept&"'")
%>
发送成功！<script>close();</script>
<%
end sub


sub index

top


sql="select username from online where username<>''"
rs.Open sql,Conn
Do While Not RS.EOF
onlinelist=""&onlinelist&""&rs("username")&"|"
rs.movenext
loop
rs.close

%>

<title>控制面板</title>
<SCRIPT>
var onlinelist= "<%=onlinelist%>";

function add(){
var id=prompt("请输入您要添加的好友名称！","");
if(id){
document.location='friend.asp?menu=add&username='+id+'';
}
}
</SCRIPT>



<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
<td vAlign="center" align="top">　<img src="images/closedfold.gif" border="0">　<a href=main.asp><%=clubname%></a><br>
　<img src="images/bar.gif" border="0"><img src="images/openfold.gif">　控制面板</td>
</tr>
</table>
<br>

<table cellspacing="1" cellpadding="1" width="97%" align="center" border="0" class="a2">
  <TR id=TableTitleLink class=a1 height="25">
      <Td width="14%" align="center"><b><a href="usercp.asp">控制面板首页</a></b></td>
      <TD width="14%" align="center"><b><a href="EditProfile.asp">基本资料修改</a></b></td>
      <TD width="14%" align="center"><b><a href="EditProfile.asp?menu=contact">编辑论坛选项</a></b></td>
      <TD width="14%" align="center"><b><a href="EditProfile.asp?menu=pass">用户密码修改</a></b></td>
      <TD width="14%" align="center"><b><a href="message.asp">用户短信服务</a></b></td>
      <TD width="14%" align="center"><b><a href="friend.asp">编辑好友列表</a></b></td>
      <TD width="16%" align="center"><b><a href="favorites.asp">用户收藏管理</a></b></td></TR></TABLE>
<form method="POST">
<input type=hidden name="menu" value="del">
<SCRIPT>valigntop()</SCRIPT>
<table width="97%" cellSpacing=1 cellPadding=3 align=center border=0 class=a2>
  <tr>
    <td class=a1 align="center">
    昵称 </th>
    <td class=a1 align="center">
    邮件 </th>
    <td class=a1 align="center">
    主页 </th>
    <td class=a1 align="center">
    状态 </th>
    <td class=a1 align="center">
    发短信 </th>
    <td class=a1 align="center">
    操作 </th>
  </tr>
<%

on error resume next '找不到好友资料时候忽略错误

sql="select friend,userface from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn
master=split(rs("friend"),"|")
for i = 1 to ubound(master)-1

usermail=Conn.Execute("Select usermail From [user] where username='"&master(i)&"'")(0)
userhome=Conn.Execute("Select userhome From [user] where username='"&master(i)&"'")(0)

%>
  <tr>
    <td vAlign=center align=middle bgcolor="FFFFFF"><a href=Profile.asp?username=<%=master(i)%>><%=master(i)%></a>　</td>
    <td align=middle bgcolor="FFFFFF"><a href=mailto:<%=usermail%>><%=usermail%></a>　</td>
    <td bgcolor="FFFFFF"><a href=<%=userhome%> target=_blank><%=userhome%></a>　</td>
    <td align=middle bgcolor="FFFFFF">
    
<Script>
if(onlinelist.indexOf('<%=master(i)%>|') == -1 ){
document.write("<img border=0 src=images/offline1.gif alt='离线'><br>");
}else{
document.write("<img border=0 src=images/online1.gif alt='在线'><br>");
}
</Script>
    
    </td>
    <td align=middle bgcolor="FFFFFF"><a style=cursor:hand onclick="javascript:open('friend.asp?menu=post&incept=<%=master(i)%>','','width=320,height=170')">发送</a></td>
    <td align=middle bgcolor="FFFFFF"><INPUT type=radio value=<%=master(i)%> name=username></td>
  </tr>
  
  
  

<%
next
rs.close
%>


  
  

  
  <tr>
    <td vAlign="center" align="right" colSpan="6" bgcolor="FFFFFF">
    <input onclick="javascript:add();" type="button" value="添加好友" name="action">&nbsp;<input onclick="{if(confirm('确定删除选定的好友吗?')){return true;}return false;}" type="submit" value="删除"></td>
  </tr></form>
</table>
<SCRIPT>valignbottom()</SCRIPT>
<%

htmlend
end sub

responseend
%>