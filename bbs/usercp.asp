<!-- #include file="setup.asp" -->

<%
top
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")

%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>








<title>控制面板</title>
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

<br>
<TABLE cellSpacing=0 cellPadding=0 width=97% align=center>

  <TR>
    <TD vAlign=top width="28%">
      <TABLE style="WIDTH: 95%" height="100%" cellSpacing=1 cellPadding=3 border=0 class=a2>
        <TBODY>
        <TR>
          <TD height=25 class=a1 align="center" colspan="2">
          <b>用户头像</TH></b></TR>
        <TR align=middle>
          <TD colspan="2" bgcolor="FFFFFF"><img src="<%=userface%>" onload='javascript:if(this.width>200)this.width=200'></TD></TR>
        <TR>
          <TD height=25 class=a1 align="center" colspan="2">
          <b>基本资料</TH></b></TR>


       <%
sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn
response.write "<TR><TD bgcolor=FFFFFF>用户等级： <Script>document.write(level("&rs("experience")&","&rs("membercode")&",'','')+levelname)</Script><br>用户门派： "&rs("faction")&"<br>用户配偶： "&rs("consort")&"<br>用户体力： "&rs("userlife")&"<br>发表原贴： "&rs("posttopic")&"<br>发表回贴： "&rs("postrevert")&"<br>精华帖数： "&rs("goodtopic")&"<br>被删原贴： "&rs("deltopic")&"<br>社区金币： "&rs("money")&"<br>用户经验： "&rs("experience")&"<br>登录次数： "&rs("degree")&"<br>注册时间： "&rs("regtime")&"<br>上次登录： "&rs("landtime")&"</TD></TR>"

rs.close
%>   


</TABLE></TD>
    <TD vAlign=top>
      <TABLE style="WIDTH: 100%" height=29 cellSpacing=1 cellPadding=3 align=center border=0 class=a2>

        <TR>
          <TD align=left height=25 class=a1>
          <b>-=&gt; 论坛短讯息</b></TD></TR>
        <TR>
          <TD align="center" bgcolor="FFFFFF"><a href="message.asp"><FONT 
            color=ff0000>收件箱</FONT></a><font color="FF0000">&nbsp; </font>共有 <B>[<%=conn.execute("Select count(id)from message where incept='"&Request.Cookies("username")&"'")(0)%>]</B> 条讯息，<a href="message.asp?send=1"><FONT 
            color=ff0000>已发送讯息</FONT></a><B> [<%=conn.execute("Select count(id)from message where author='"&Request.Cookies("username")&"'")(0)%>]</B> 条。<BR></TD></TR></TBODY></TABLE><BR>
      <TABLE  style="WIDTH: 100%" cellSpacing=1 cellPadding=3 align=center border=0 class=a2>
        <TBODY>
        <TR>
          <TD align=left colSpan=5 height=25 class=a1>
          <b>-=&gt; 最新收到的短讯</TH></b></TR>
        <TR>
          <TD vAlign=center align=middle 
          width=30 bgcolor="FFFFFF">　</TD>
          <TD vAlign=center align=middle 
          width=100 bgcolor="FFFFFF"><B>发件人</B></TD>
          <TD vAlign=center align=middle 
          width=300 bgcolor="FFFFFF"><B>内容</B></TD>
          <TD align=middle 
          width=160 bgcolor="FFFFFF"><B>日期</B></TD>
          <TD vAlign=center align=middle 
          width=50 bgcolor="FFFFFF"><B>大小</B></TD></TR>
          
<%

sql="select top 5 author,content,time from message where incept='"&Request.Cookies("username")&"' order by time Desc"
rs.Open sql,Conn
Do While Not RS.EOF 
size=Len(""&rs("content")&"")
if size>20 then
content=left(""&rs("content")&"",16)&"..."
else
content=rs("content")
end if
%>
   
        <TR>
          <TD vAlign=center align=middle bgcolor="FFFFFF">
          <IMG 
            src="images/f_norm.gif"></TD>
          <TD vAlign=center align=middle bgcolor="FFFFFF"><A href="Profile.asp?username=<%=rs("author")%>" ><%=rs("author")%></A>　</TD>
          <TD bgcolor="FFFFFF"><A href=message.asp><%=content%></A>　</TD>
          <TD align="center" bgcolor="FFFFFF"><%=rs("time")%></FONT>　</TD>
          <TD align="center" bgcolor="FFFFFF"><%=size%>　</TD></TR>
          
<%          
RS.MoveNext
loop
RS.Close      
%>          
          
          
          
          </TD></TR></TABLE><BR>
      <TABLE style="WIDTH: 100%" cellSpacing=1 cellPadding=3 align=center border=0 class=a2>

        <TR>
          <TD align=left height=25 class=a1>
          <b>-=&gt; 
  最近发表的文章</TH></b></TR>
    <%
sql="select top 5 * from forum where deltopic<>1 and username='"&Request.Cookies("username")&"' order by lasttime Desc  "
rs.Open sql,Conn

Do While Not RS.EOF 

%>

  <tr>
    <td align="left" bgcolor="FFFFFF">&nbsp;□　<a onclick=min_yuzi() target=message href=ShowPost.asp?id=<%=rs("id")%>><%=rs("topic")%></a> -- <%=rs("lasttime")%></td>
  </tr>


<%


RS.MoveNext
loop
RS.Close

  %>
</table></TD></TR></TBODY></TABLE>
  
<%
htmlend
%>