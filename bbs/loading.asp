<!-- #include file="setup.asp" -->
<%
id=int(Request("id"))


if id="0" then

if Request("forumid")="0" then
sql="select * from online where username<>'' and eremite<>1"
else
sql="select * from online where forumid="&int(Request("forumid"))&" and username<>'' and eremite<>1"
end if

rs.Open sql,Conn

do while not rs.eof

if NO_count < 6 then
NO_count=NO_count+1
else
NO_count=1
end if


allline=""&allline&"<td width=16% style=word-break:break-all><img src="&rs("userface")&" width=16 height=16> <a href=Profile.asp?username="&rs("username")&">"&rs("username")&"</a></td>"

if NO_count = 6 then
allline=""&allline&"</tr><tr>"
end if

rs.movenext
loop
rs.close

allline="<TABLE cellSpacing=0 cellPadding=3 width=100% border=0><TR>"&allline&"</TR></TABLE>"
%>
<META http-equiv=Content-Type content=text/html;charset=gb2312>
<SCRIPT>
parent.followImg0.loaded="yes";
parent.followTd0.innerHTML="<table width=100% cellpadding=5 style=TABLE-LAYOUT:fixed><tr><td width=100% style=WORD-WRAP:break-word><%=allline%></td></tr></table>";
</SCRIPT>
<%

else

%>
<META http-equiv=Content-Type content=text/html;charset=gb2312>
<SCRIPT>
parent.followImg<%=id%>.loaded="yes";
parent.followTd<%=id%>.innerHTML="<table width=100% cellpadding=10 style=TABLE-LAYOUT:fixed><tr><td width=100% style=WORD-WRAP:break-word><%=Conn.Execute("Select content From forum where id="&id)(0)%></td></tr></table>";
</SCRIPT>
<%
end if
responseend




%>