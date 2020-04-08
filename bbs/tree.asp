<!-- #include file="setup.asp" -->
<%

sql="select id,bbsname,classid from bbsconfig where hide=0 order by classid,id"
rs.Open sql,Conn
do while not rs.eof
Classid=Trim(Rs("classid"))
if TClass <> Classid Then
		
classname=Conn.Execute("Select classname From class where id="&Classid)(0)
if Len(""&classname&"")>10 then
classname=left(""&classname&"",8)&"..."
end if

alltree=""&alltree&"</DIV><img src=images/bullet.gif><A href=javascript:expands('yuzi"&rs("id")&"')><font color=215DC6>"&classname&"</font></A><BR><DIV class=child id=yuzi"&rs("id")&"Child style='display:none'>"
TClass = Classid
end if

bbsname=rs("bbsname")
if Len(""&bbsname&"")>9 then
bbsname=left(""&bbsname&"",7)&"..."
end if

alltree=""&alltree&"<img src=images/bar.gif> <A href=ShowForum.asp?forumid="&rs("id")&" target=yuzi_frame onClick=message()>"&bbsname&"</A><br>"

rs.movenext
loop
rs.close
%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<SCRIPT>
parent.followTd.innerHTML="<%=alltree%></DIV>";
</SCRIPT>

<%responseend%>