<!-- #include file="setup.asp" -->
<%

if Request.Cookies("username")=empty then
error2("���¼���ٽ��в�����")
end if

if Request.Cookies("userpass") <> Conn.Execute("Select userpass From [user] where username='"&Request.Cookies("username")&"'")(0) then
error("<li>�����������")
end if

id=int(Request("id"))
forumid=Conn.Execute("Select forumid From forum where id="&id)(0)

membercode=Conn.Execute("Select membercode From [user] where username='"&Request.Cookies("username")&"'")(0)
if membercode > 3 then
pass=1
elseif instr(Conn.Execute("Select moderated From [bbsconfig] where id="&forumid&" ")(0),"|"&Request.Cookies("username")&"|")>0 then
pass=1
end if
if pass<>1 then
error("<li>����Ȩ�޲���")
end if

username=Conn.Execute("Select username From [forum] where id="&id&" and forumid="&forumid&"")(0)
select case Request("menu")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "top"
if membercode > 3 then
conn.execute("update [forum] set toptopic=2,lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where id="&id&" and forumid="&forumid&"")
succtitle="���ö�����ɹ�"
else
error("<li>����Ȩ�޲���")
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "untop"
if membercode > 3 then
conn.execute("update [forum] set toptopic=0,lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where id="&id&" and forumid="&forumid&"")
succtitle="ȡ�����ö�����ɹ�"
else
error("<li>����Ȩ�޲���")
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "movenew"
conn.execute("update [forum] set lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where id="&id&" and forumid="&forumid&"")
succtitle="��ǰ����ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "move"
if Request("moveid")="" then:error("<li>��û��ѡ��Ҫ�������ƶ��ĸ���̳"):end if
conn.execute("update [forum] set forumid="&Request("moveid")&",toptopic=0,goodtopic=0,locktopic=0,lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where id="&id&" and forumid="&forumid&"")
succtitle="�ƶ�����ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "deltopic"
if isnumeric(""&Request("retopicid")&"") then
conn.execute("delete from [reforum] where topicid="&id&" and id="&Request("retopicid")&"")
conn.execute("update [forum] set replies=replies-1,lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where id="&id&" and forumid="&forumid&"")
conn.execute("update [bbsconfig] set tolrestore=tolrestore-1 where id="&forumid&"")
succtitle="ɾ�������ɹ�"
else
conn.execute("update [user] set deltopic=deltopic+1 where username='"&username&"'")
conn.execute("update [forum] set toptopic=0,deltopic=1,lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where id="&id&" and forumid="&forumid&" and deltopic=0")
conn.execute("update [bbsconfig] set toltopic=toltopic-1,tolrestore=tolrestore-1 where id="&forumid&"")
succtitle="ɾ������ɹ�"
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "goodtopic"
if Conn.Execute("Select goodtopic From [forum] where id="&id&" ")(0)=1 then:error("<li>�������Ѿ����뾫�����ˣ������ظ����"):end if
conn.execute("update [forum] set goodtopic=1,lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where id="&id&" and forumid="&forumid&"")
conn.execute("update [user] set goodtopic=goodtopic+1 where username='"&username&"'")
succtitle="���뾫�����ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "delgoodtopic"
if Conn.Execute("Select goodtopic From [forum] where id="&id&" ")(0)=0 then:error("<li>�������Ѿ��Ƴ���������"):end if
conn.execute("update [forum] set goodtopic=0,lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where id="&id&" and forumid="&forumid&"")
conn.execute("update [user] set goodtopic=goodtopic-1 where username='"&username&"'")
succtitle="�Ƴ��������ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "toptopic"
conn.execute("update [forum] set toptopic=1,lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where id="&id&" and forumid="&forumid&"")
succtitle="�ö�����ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "deltoptopic"
conn.execute("update [forum] set toptopic=0,lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where id="&id&" and forumid="&forumid&"")
succtitle="ȡ���ö�����ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "locktopic"
conn.execute("update [forum] set locktopic=1,lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where id="&id&" and forumid="&forumid&"")
succtitle="�ر�����ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "dellocktopic"
conn.execute("update [forum] set locktopic=0,lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where id="&id&" and forumid="&forumid&"")
succtitle="��������ɹ�"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
case "lookip"
if isnumeric(""&Request("retopicid")&"") then
error2("IP��ַ��"&Conn.Execute("Select postip From [reforum] where id="&Request("retopicid")&" ")(0)&"")
else
error2("IP��ַ��"&Conn.Execute("Select postip From [forum] where id="&id&" ")(0)&"")
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

end select
if succtitle="" then
error("<li>��Ч����")
end if
message="<li>"&succtitle&"<li><a href=ShowForum.asp?forumid="&forumid&">������̳</a><li><a href=main.asp>����������ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=ShowForum.asp?forumid="&forumid&">")
%>