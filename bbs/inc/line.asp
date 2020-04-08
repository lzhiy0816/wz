<%
sql="select * from online where ip='"&remoteaddr&"' and username='"&Request.Cookies("username")&"' or sessionid='"&session.sessionid&"' "
rs.Open sql,Conn,1,3
if rs.eof then
rs.addnew
rs("act")=""&act&""
rs("acturl")=""&acturl&""
rs("forumid")=""&forumid&""
rs("bbsname")=""&bbsname&""
if Request.Cookies("eremite")<>empty then rs("eremite")=Request.Cookies("eremite")
rs("userface")=userface
rs("ip")=remoteaddr
rs("username")=Request.Cookies("username")
rs("sessionid")=session.sessionid
else
rs("act")=""&act&""
rs("acturl")=""&acturl&""
rs("forumid")=forumid
rs("bbsname")=""&bbsname&""
rs("lasttime")=now
if Request.Cookies("eremite")<>empty then rs("eremite")=Request.Cookies("eremite")
rs("userface")=userface
rs("ip")=remoteaddr
rs("username")=Request.Cookies("username")
rs("sessionid")=session.sessionid
end if
rs.update
rs.close
conn.execute("delete from [online] where DateDiff("&SqlChar&"s"&SqlChar&",lasttime,"&SqlNowString&")>"&OnlineTime&" ")
onlinemany=conn.execute("Select count(sessionid)from online")(0)
%>